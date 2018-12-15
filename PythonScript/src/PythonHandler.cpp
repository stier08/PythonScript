#include "stdafx.h"
#include "PythonHandler.h"

#include "Scintilla.h"
#include "ScintillaWrapper.h"
#include "ScintillaPython.h"
#include "NotepadPlusWrapper.h"
#include "NotepadPython.h"
#include "PythonConsole.h"
#include "MenuManager.h"
#include "WcharMbcsConverter.h"
#include "GILManager.h"
#include "ConfigFile.h"

std::string extractStringFromPyStr(PyObject* strObj)
{
	std::string ret;
#  if PY_VERSION_HEX >= 0x03000000
	PyObject* bytes = PyUnicode_AsUTF8String(strObj);
	ret = PyBytes_AsString(bytes);
	if (PyErr_Occurred()) return "";
	Py_DECREF(bytes);
#else
	ret = PyString_AsString(strObj);
#endif
	return ret;
}

bool pyStrCheck(PyObject* strObj)
{
#  if PY_VERSION_HEX >= 0x03000000
	return PyUnicode_Check(strObj);
#else
	return PyString_Check(strObj);
#endif

}

std::string getPythonErrorString()
{
	// Extra paranoia...
	if (!PyErr_Occurred())
	{
		return "No Python error";
	}

	PyObject *type, *value, *traceback;
	PyErr_Fetch(&type, &value, &traceback);
	PyErr_Clear();

	std::string message = "Python error: ";
	if (type)
	{
		type = PyObject_Str(type);

		message += extractStringFromPyStr(type);
	}

	if (value)
	{
		value = PyObject_Str(value);
		message += ": ";
		message += extractStringFromPyStr(value);
	}

	Py_XDECREF(type);
	Py_XDECREF(value);
	Py_XDECREF(traceback);

	return message;
}

/*
Modeled after a function from Mark Hammond.

Obtains a string from a Python traceback.  This is the exact same string as
"traceback.print_exception" would return.

Result is a string which must be free'd using PyMem_Free()
*/
#define TRACEBACK_FETCH_ERROR(what) {errMsg = what; goto done;}

std::string PyTracebackToString(void)
{
	std::string errMsg; /* holds a local error message */
	std::string result; /* a valid, allocated result. */

	PyObject *modStringIO = NULL;
	PyObject *modTB = NULL;
	PyObject *obStringIO = NULL;
	PyObject *obResult = NULL;

	PyObject *type, *value, *traceback;

	PyErr_Fetch(&type, &value, &traceback);
	PyErr_NormalizeException(&type, &value, &traceback);

#  if PY_VERSION_HEX >= 0x03000000
	modStringIO = PyImport_ImportModule("io");
#else
	modStringIO = PyImport_ImportModule("cStringIO");
#endif

	if (modStringIO == NULL)
		TRACEBACK_FETCH_ERROR("cant import cStringIO\n");

	obStringIO = PyObject_CallMethod(modStringIO, "StringIO", NULL);

	/* Construct a cStringIO object */
	if (obStringIO == NULL)
		TRACEBACK_FETCH_ERROR("cStringIO.StringIO() failed\n");

	modTB = PyImport_ImportModule("traceback");
	if (modTB == NULL)
		TRACEBACK_FETCH_ERROR("cant import traceback\n");

	obResult = PyObject_CallMethod(modTB, "print_exception",
		"OOOOO",
		type, value ? value : Py_None,
		traceback ? traceback : Py_None,
		Py_None,
		obStringIO);

	if (obResult == NULL)
		TRACEBACK_FETCH_ERROR("traceback.print_exception() failed\n");
	Py_DECREF(obResult);

	obResult = PyObject_CallMethod(obStringIO, "getvalue", NULL);
	if (obResult == NULL)
		TRACEBACK_FETCH_ERROR("getvalue() failed.\n");

	/* And it should be a string all ready to go - duplicate it. */
	if (!pyStrCheck(obResult))
		TRACEBACK_FETCH_ERROR("getvalue() did not return a string\n");

	result = extractStringFromPyStr(obResult);
done:

	/* All finished - first see if we encountered an error */
	if (result.empty() && errMsg.size()) {
		result = errMsg;
	}

	Py_XDECREF(modStringIO);
	Py_XDECREF(modTB);
	Py_XDECREF(obStringIO);
	Py_XDECREF(obResult);
	Py_XDECREF(value);
	Py_XDECREF(traceback);
	Py_XDECREF(type);

	return result;
}

namespace NppPythonScript
{

PythonHandler::PythonHandler(TCHAR *pluginsDir, TCHAR *configDir, HINSTANCE hInst, HWND nppHandle, HWND scintilla1Handle, HWND scintilla2Handle, boost::shared_ptr<PythonConsole> pythonConsole)
	: PyProducerConsumer<RunScriptArgs>(),
	  m_nppHandle(nppHandle),
      m_scintilla1Handle(scintilla1Handle),
	  m_scintilla2Handle(scintilla2Handle),
	  m_hInst(hInst),
	  m_machineBaseDir(pluginsDir),
	  m_userBaseDir(configDir),
	  mp_console(pythonConsole),
	  m_currentView(0),
	  mp_mainThreadState(NULL),
	  m_consumerStarted(false)
{
	m_machineBaseDir.append(_T("\\PythonScript\\"));
	m_userBaseDir.append(_T("\\PythonScript\\"));

	mp_notepad = createNotepadPlusWrapper();
	mp_scintilla = createScintillaWrapper();
	mp_scintilla1.reset(new ScintillaWrapper(scintilla1Handle, m_nppHandle));
	mp_scintilla2.reset(new ScintillaWrapper(scintilla2Handle, m_nppHandle));
}

PythonHandler::~PythonHandler(void)
{
	try
	{
		if (Py_IsInitialized())
		{
			if (consumerBusy())
			{
				stopScript();
			}

			// We need to swap back to the main thread
			GILLock lock;  // It's actually pointless, as we don't need it anymore,
			               // but we'll grab it anyway, just in case we need to wait for something to finish

			// Can't call finalize with boost::python.
			// Py_Finalize();

		}


		// To please Lint, let's NULL these handles
		m_hInst = NULL;
		m_nppHandle = NULL;
		mp_mainThreadState = NULL;
	}
	catch (...)
	{
		// I don't know what to do with that, but a destructor should never throw, so...
	}
}


boost::shared_ptr<ScintillaWrapper> PythonHandler::createScintillaWrapper()
{
	m_currentView = mp_notepad->getCurrentView();
	return boost::shared_ptr<ScintillaWrapper>(new ScintillaWrapper(m_currentView ? m_scintilla2Handle : m_scintilla1Handle, m_nppHandle));
}

boost::shared_ptr<NotepadPlusWrapper> PythonHandler::createNotepadPlusWrapper()
{
	return boost::shared_ptr<NotepadPlusWrapper>(new NotepadPlusWrapper(m_hInst, m_nppHandle));
}

void PythonHandler::initPython()
{
	if (Py_IsInitialized())
		return;

	DEBUG_TRACE("PythonHandler::initPython");
	preinitScintillaModule();

	// Don't import site - if Python 2.7 doesn't find it as part of Py_Initialize,
	// it does an exit(1) - AGH!
	Py_NoSiteFlag = 1;

	Py_Initialize();
    // Initialise threading and create & acquire Global Interpreter Lock
	PyEval_InitThreads();


	std::shared_ptr<char> machineBaseDir = WcharMbcsConverter::tchar2char(m_machineBaseDir.c_str());
	std::shared_ptr<char> configDir = WcharMbcsConverter::tchar2char(m_userBaseDir.c_str());

	bool machineIsUnicode = containsExtendedChars(machineBaseDir.get());
	bool userIsUnicode    = containsExtendedChars(configDir.get());


	std::string smachineDir(machineBaseDir.get());
	std::string suserDir(configDir.get());


	// Init paths
	char initBuffer[1024];
    char pathCommands[500];

    // If the user wants to use their installed python version, append the paths.
    // If not (and they want to use the bundled python install), the default, then prepend the paths
    if (ConfigFile::getInstance()->getSetting(_T("PREFERINSTALLEDPYTHON")) == _T("1")) {
        strcpy_s<500>(pathCommands, "import sys\n"
            "sys.path.append(r'%slib'%s)\n"
            "sys.path.append(r'%slib'%s)\n"
            "sys.path.append(r'%sscripts'%s)\n"
            "sys.path.append(r'%sscripts'%s)\n"
			"sys.path.append(r'%slib\\lib-tk'%s)\n" );
	} else {
        strcpy_s<500>(pathCommands, "import sys\n"
            "sys.path.insert(0,r'%slib'%s)\n"
            "sys.path.insert(1,r'%slib'%s)\n"
            "sys.path.insert(2,r'%sscripts'%s)\n"
            "sys.path.insert(3,r'%sscripts'%s)\n"
            "sys.path.insert(4,r'%slib\\lib-tk'%s)\n"
			);
	}

	_snprintf_s(initBuffer, 1024, 1024,
        pathCommands,
		smachineDir.c_str(),
		machineIsUnicode ? ".decode('utf8')" : "",

		suserDir.c_str(),
		userIsUnicode ? ".decode('utf8')" : "",

		smachineDir.c_str(),
		machineIsUnicode ? ".decode('utf8')" : "",

		suserDir.c_str(),
		userIsUnicode ? ".decode('utf8')" : "",

		smachineDir.c_str(),
		machineIsUnicode ? ".decode('utf8')" : ""
		);

	PyRun_SimpleString(initBuffer);

    initSysArgv();


	// Init Notepad++/Scintilla modules
	initModules();

    mp_mainThreadState = PyEval_SaveThread();

}

void PythonHandler::initSysArgv()
{
    LPWSTR commandLine = ::GetCommandLineW();
    int argc;
    LPWSTR* argv = ::CommandLineToArgvW(commandLine, &argc);


    boost::python::list argvList;
    for(int currentArg = 0; currentArg != argc; ++currentArg)
	{
        std::shared_ptr<char> argInUtf8 = WcharMbcsConverter::wchar2char(argv[currentArg]);
        PyObject* unicodeArg = PyUnicode_FromString(argInUtf8.get());

		argvList.append(boost::python::handle<>(unicodeArg));
    }

    boost::python::object sysModule(boost::python::handle<>(boost::python::borrowed(PyImport_AddModule("sys"))));
    sysModule.attr("argv") = argvList;


}

void PythonHandler::initModules()
{
	importScintilla(mp_scintilla, mp_scintilla1, mp_scintilla2);
	importNotepad(mp_notepad);
	importConsole(mp_console);
}


bool PythonHandler::containsExtendedChars(char *s)
{
	bool retVal = false;
	for(int pos = 0; s[pos]; ++pos)
	{
		if (s[pos] & 0x80)
		{
			retVal = true;
			break;
		}
	}
	return retVal;
}

void PythonHandler::runStartupScripts()
{

	// Machine scripts (N++\Plugins\PythonScript\scripts dir)
	tstring startupPath(m_machineBaseDir);
	startupPath.append(_T("scripts\\startup.py"));
	if (::PathFileExists(startupPath.c_str()))
	{

		runScript(startupPath, true);
	}

	// User scripts ($CONFIGDIR$\PythonScript\scripts dir)
	startupPath = m_userBaseDir;
	startupPath.append(_T("scripts\\startup.py"));
	if (::PathFileExists(startupPath.c_str()))
	{
		runScript(startupPath, true);
	}

}

bool PythonHandler::runScript(const tstring& scriptFile,
							  bool synchronous /* = false */,
							  bool allowQueuing /* = false */,
							  HANDLE completedEvent /* = NULL */,
							  bool isStatement /* = false */)
{
	return runScript(scriptFile.c_str(), synchronous, allowQueuing, completedEvent, isStatement);
}

bool PythonHandler::runScript(const TCHAR *filename,
							  bool synchronous /* = false */,
							  bool allowQueuing /* = false */,
							  HANDLE completedEvent /* = NULL */,
							  bool isStatement /* = false */)
{
	bool retVal;

	if (!allowQueuing && consumerBusy())
	{
		retVal = false;
	}
	else
	{
		std::shared_ptr<RunScriptArgs> args(
			new RunScriptArgs(
				filename,
				mp_mainThreadState,
				synchronous,
				completedEvent,
				isStatement));

		if (!synchronous)
		{
			retVal = produce(args);
			if (!m_consumerStarted)
			{
				startConsumer();
			}
		}
		else
		{
			runScriptWorker(args);
			retVal = true;
		}
	}
	return retVal;
}

void PythonHandler::consume(std::shared_ptr<RunScriptArgs> args)
{
	runScriptWorker(args);
}

void PythonHandler::runScriptWorker(const std::shared_ptr<RunScriptArgs>& args)
{

    GILLock gilLock;
	DEBUG_TRACE("PythonHandler::runScriptWorker");
	if (args->m_isStatement)
	{
		DEBUG_TRACE("PythonHandler::runScriptWorker - 1");
		if (PyRun_SimpleString(WcharMbcsConverter::tchar2char(args->m_filename.c_str()).get()) == -1)
		{
			if (ConfigFile::getInstance()->getSetting(_T("ADDEXTRALINETOOUTPUT")) == _T("1"))
			{
				mp_console->writeText(boost::python::str("\n"));
			}
			
			if (ConfigFile::getInstance()->getSetting(_T("OPENCONSOLEONERROR")) == _T("1"))
			{
				mp_console->pythonShowDialog();
			}
		}
	}
	else
	{
		DEBUG_TRACE("PythonHandler::runScriptWorker - 2");
		std::shared_ptr<char> filenameUFT8 = WcharMbcsConverter::tchar2char(args->m_filename.c_str());

		if (containsExtendedChars(filenameUFT8.get()))
		{
			// First obtain the size needed by passing NULL and 0.
			const long initLength = GetShortPathName(args->m_filename.c_str(), NULL, 0);
			if (initLength > 0)
			{
				// Dynamically allocate the correct size
				// (terminating null char was included in length)
				tstring buffer(initLength, 0);

				// Now simply call again using same long path.

				long length = GetShortPathName(args->m_filename.c_str(), const_cast<LPWSTR>(buffer.c_str()), initLength);
				if (length > 0)
				{
					filenameUFT8 = WcharMbcsConverter::tchar2char(buffer.c_str());
				}
			}
		}

		// We assume PyFile_FromString won't modify the file name passed in param
		// (that would be quite troubling) and that the missing 'const' is simply an oversight
		// from the Python API developers.
		// We also assume the second parameter, "r" won't be modified by the function call.
		//lint -e{1776}  Converting a string literal to char * is not const safe (arg. no. 2)
		PyObject* pyio = PyImport_ImportModule("io");
		DEBUG_TRACE(filenameUFT8.get());
		if (pyio)
		{
			DEBUG_TRACE("PythonHandler::runScriptWorker - pyio - OK");
			PyObject* pyname = PyUnicode_FromString("open");
			if (pyname)
			{
				DEBUG_TRACE("PythonHandler::runScriptWorker - pyname - OK");
				PyObject* pyioopen = PyObject_GetAttr(pyio, pyname);
				if (pyioopen)
				{
					DEBUG_TRACE("PythonHandler::runScriptWorker - pyioopen - OK");
					PyObject* pyfname = PyUnicode_FromString(filenameUFT8.get());
					if (pyfname)
					{
						DEBUG_TRACE("PythonHandler::runScriptWorker - pyfname - OK");
					}
					PyObject* pyArg = PyTuple_New(0);
					if (pyArg)
					{
						DEBUG_TRACE("PythonHandler::runScriptWorker - pyArg - OK");
					}
					PyObject* pyFile = PyObject_Call(pyioopen, pyfname, pyArg);

					FILE* cfile = fopen(filenameUFT8.get(), "r");
					if (cfile)
					{
						DEBUG_TRACE("PythonHandler::runScriptWorker - cfile - OK");
					}
					if (cfile)
					{
						DEBUG_TRACE("PythonHandler::runScriptWorker - pyFile - OK");

						int pyret= PyRun_SimpleFile(cfile, filenameUFT8.get());
						DEBUG_TRACE("PythonHandler::runScriptWorker - PyRun_SimpleFile - Done");
						if (pyret == -1)
						{
							DEBUG_TRACE("PythonHandler::runScriptWorker - pyret - -1");
							if (ConfigFile::getInstance()->getSetting(_T("ADDEXTRALINETOOUTPUT")) == _T("1"))
							{
								mp_console->writeText(boost::python::str("\n"));
							}

							if (ConfigFile::getInstance()->getSetting(_T("OPENCONSOLEONERROR")) == _T("1"))
							{
								mp_console->pythonShowDialog();
							}
						}
						if (pyFile)
						{
							Py_DECREF(pyFile);
						}
						Py_DECREF(pyArg);
						}
					Py_DECREF(pyioopen);
				}
					Py_DECREF(pyname);
			}
				Py_DECREF(pyio);
		}
		else
		{
			DEBUG_TRACE("PythonHandler::runScriptWorker - pyio - ERROR");
		}
	}



	try
	{
		PyObject *py_main, *py_dict;
		py_main = PyImport_AddModule("__main__");

		if (py_main)
		{
			DEBUG_TRACE("PythonHandler::runScriptWorker - py_main - OK");
		}
		else
		{
			DEBUG_TRACE("PythonHandler::runScriptWorker - py_main - ERROR");
			std::string s = PyTracebackToString();
			DEBUG_TRACE(s.c_str());

		}

		py_dict = PyModule_GetDict(py_main);

		if (py_dict)
		{
			DEBUG_TRACE("PythonHandler::runScriptWorker - py_dict - OK");
		}
		else
		{
			DEBUG_TRACE("PythonHandler::runScriptWorker - py_dict - ERROR");
			std::string s = PyTracebackToString();
			DEBUG_TRACE(s.c_str());

		}

		PyObject*  obj1=PyRun_String("open(r'd:\temp\b.txt','w').write('x')",
			Py_file_input,
			py_dict,
			py_dict
		);

		if (obj1)
		{
			DEBUG_TRACE("PythonHandler::runScriptWorker - obj1 - OK");
		}
		else
		{
			DEBUG_TRACE("PythonHandler::runScriptWorker - obj1 - ERROR");
			std::string s = PyTracebackToString();
			DEBUG_TRACE(s.c_str());

		}

		PyObject*  obj2 = PyRun_String("\n",
			Py_file_input,
			py_dict,
			py_dict
		);

		if (obj2)
		{
			DEBUG_TRACE("PythonHandler::runScriptWorker - obj2 - OK");
		}
		else
		{
			DEBUG_TRACE("PythonHandler::runScriptWorker - obj2 - ERROR");
			std::string s = PyTracebackToString();
			DEBUG_TRACE(s.c_str());

		}

		DEBUG_TRACE("PythonHandler::runScriptWorker - PyRun_String - 4x");
		PyRun_String("open(r'd:\temp\b.txt','w').write('x')",

			Py_file_input,
			PyDict_New(),
			PyDict_New()
		);
		DEBUG_TRACE("PythonHandler::runScriptWorker - PyRun_String - 5x");
	}
	catch (...)
	{
		std::string s = PyTracebackToString();
		DEBUG_TRACE(s.c_str());
	}



	if (NULL != args->m_completedEvent)
	{
		SetEvent(args->m_completedEvent);
	}
	DEBUG_TRACE("PythonHandler::runScriptWorker - END");
}

void PythonHandler::notify(SCNotification *notifyCode)
{
	if (notifyCode->nmhdr.hwndFrom == m_scintilla1Handle || notifyCode->nmhdr.hwndFrom == m_scintilla2Handle)
	{
		mp_scintilla->notify(notifyCode);
	}
	else if (notifyCode->nmhdr.hwndFrom != mp_console->getScintillaHwnd()) // ignore console notifications
	{
		// Change the active scintilla handle for the "buffer" variable if the active buffer has changed
		if (notifyCode->nmhdr.code == NPPN_BUFFERACTIVATED)
		{
			int newView = mp_notepad->getCurrentView();
			if (newView != m_currentView)
			{
				m_currentView = newView;
				mp_scintilla->setHandle(newView ? m_scintilla2Handle : m_scintilla1Handle);
			}
		}

		mp_notepad->notify(notifyCode);
	}
}

void PythonHandler::queueComplete()
{
	MenuManager::getInstance()->stopScriptEnabled(false);
}


void PythonHandler::stopScript()
{
	DWORD threadID;
	CreateThread(NULL, 0, reinterpret_cast<LPTHREAD_START_ROUTINE>(stopScriptWorker), this, 0, &threadID);
}


void PythonHandler::stopScriptWorker(PythonHandler *handler)
{
    GILLock gilLock;

	PyThreadState_SetAsyncExc((long)handler->getExecutingThreadID(), PyExc_KeyboardInterrupt);

}


}