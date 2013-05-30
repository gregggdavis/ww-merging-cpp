
#include "stdafx.h"

#include "XML2Word.h"

#include "Ini.h"
extern CIni cIni;


BOOL Xml2Word(CString cstrPathMerged,
              CString cstrPathTemplates,
              CString cstrFileName,
              CString cstrCourseName,
              CString cstrAssignmentName,
              CString cstrXmlData)
{
  HRESULT hresult;
  CLSID clsid;
  BOOL bReturnValue = FALSE;

  CoInitialize(NULL);	//initialize COM library
  hresult=CLSIDFromProgID(OLESTR("XML2Word.clMain"), &clsid);    //retrieve CLSID of component
  if(FAILED(hresult)) {
    system("regsvr32 XML2Word.dll");
    hresult=CLSIDFromProgID(OLESTR("XML2Word.clMain"), &clsid);    //retrieve CLSID of component
    if(FAILED(hresult)) {
      AfxMessageBox("CLSIDFromProgID Failed");
      CoUninitialize();  //Unintialize the COM library
      return bReturnValue;
    }
  }

  _clMain *t; 
  hresult=CoCreateInstance(clsid,NULL,CLSCTX_INPROC_SERVER,__uuidof(_clMain),(LPVOID *) &t);
  if(FAILED(hresult)) {
    system("regsvr32 XML2Word.dll");
    hresult=CoCreateInstance(clsid,NULL,CLSCTX_INPROC_SERVER,__uuidof(_clMain),(LPVOID *) &t);
    if(FAILED(hresult)) {
      AfxMessageBox("CoCreateInstance Failed");
      CoUninitialize();  //Unintialize the COM library
      return bReturnValue;
    }
  }

  CString cstrPrefix   = cstrFileName;
  CString cstrPath     = cstrPathMerged + "\\";
  CString cstrTemplate = cstrPathTemplates + "\\" + cstrCourseName + "\\" + cstrCourseName + "_" + cstrAssignmentName + ".DOT";
  //CString strXml      = cstrPathMerged + "\\" + cstrFileName + ".xml";

  if (cIni.GetBool("Config", "Debugging", false)) {
    AfxMessageBox(cstrPrefix);
    AfxMessageBox(cstrPath);
    AfxMessageBox(cstrTemplate);
    AfxMessageBox(cstrXmlData);
  }

  BSTR bstrPrefix    = NULL;
  BSTR bstrPath      = NULL;
  BSTR bstrTemplate  = NULL;
  BSTR bstrXmlString = NULL;

  bstrPrefix    = cstrPrefix.AllocSysString();
  bstrPath      = cstrPath.AllocSysString();
  bstrTemplate  = cstrTemplate.AllocSysString();
  bstrXmlString = cstrXmlData.AllocSysString();

  t->PutDocumentPrefix(&bstrPrefix);
  t->PutPath(&bstrPath);
  t->PutWordTemplate(&bstrTemplate);
  t->PutXMLString(&bstrXmlString);

  bReturnValue = t->Execute();
  t->Release();

  SysFreeString(bstrPrefix);
  SysFreeString(bstrPath);
  SysFreeString(bstrTemplate);
  SysFreeString(bstrXmlString);

  bReturnValue = TRUE;

  CoUninitialize();  //Unintialize the COM library

  return bReturnValue;
}



bool CloseWord(bool bGracefully)
{
  OSVERSIONINFO  osver;      
  HINSTANCE      hInstLib;
  LPDWORD        lpdwPIDs;      
  DWORD          dwSize, dwSize2, dwIndex;
  HMODULE        hMod;      
  HANDLE         hProcess;
  char           szFileName[ MAX_PATH ];

  bool bReturnOk = true;

  // PSAPI Function Pointers.
  BOOL  (WINAPI *lpfEnumProcesses)( DWORD *, DWORD cb, DWORD * );
  BOOL  (WINAPI *lpfEnumProcessModules)( HANDLE, HMODULE *, DWORD, LPDWORD );
  DWORD (WINAPI *lpfGetModuleFileNameEx)( HANDLE, HMODULE, PTSTR, DWORD );      

  // Check to see if we're running under Windows95 or Windows NT.
  osver.dwOSVersionInfoSize = sizeof( osver ) ;
  if (!GetVersionEx(&osver)) {
    AfxMessageBox("running under Windows95 or Windows NT");
    return bReturnOk;
  }
  // If Windows NT:
  if (osver.dwPlatformId != VER_PLATFORM_WIN32_NT) {
    AfxMessageBox("Windows NT");
    return bReturnOk;
  }
  // Load library and get the procedures explicitly. We do
  // this so that we don't have to worry about modules using
  // this code failing to load under Windows 95, because
  // it can't resolve references to the PSAPI.DLL.
  hInstLib = LoadLibraryA( "PSAPI.DLL" ) ;         
  if (hInstLib == NULL) {
    AfxMessageBox("Can't Loadlibrary PSAPI.DLL");
    return bReturnOk;
  }

  // Get procedure addresses.
  lpfEnumProcesses = (BOOL(WINAPI *)(DWORD *,DWORD,DWORD*))
                     GetProcAddress(hInstLib, "EnumProcesses");

  lpfEnumProcessModules = (BOOL(WINAPI *)(HANDLE, HMODULE *,DWORD, LPDWORD)) 
                          GetProcAddress(hInstLib, "EnumProcessModules");

  lpfGetModuleFileNameEx = (DWORD (WINAPI *)(HANDLE, HMODULE, LPTSTR, DWORD )) 
                           GetProcAddress(hInstLib, "GetModuleFileNameExA");
  
  if ((lpfEnumProcesses == NULL)
  ||  (lpfEnumProcessModules == NULL)
  ||  (lpfGetModuleFileNameEx == NULL)) {
    FreeLibrary(hInstLib);
    AfxMessageBox("Can't Enum Processes");
    return bReturnOk;
  }

  // Call the PSAPI function EnumProcesses to get all of the
  // ProcID's currently in the system.
  // NOTE: In the documentation, the third parameter of
  // EnumProcesses is named cbNeeded, which implies that you
  // can call the function once to find out how much space to
  // allocate for a buffer and again to fill the buffer.
  // This is not the case. The cbNeeded parameter returns
  // the number of PIDs returned, so if your buffer size is
  // zero cbNeeded returns zero.
  // NOTE: The "HeapAlloc" loop here ensures that we
  // actually allocate a buffer large enough for all the
  // PIDs in the system.
  dwSize2 = 256 * sizeof(DWORD);
  lpdwPIDs = NULL ;

  do {
    if (lpdwPIDs) {
      HeapFree(GetProcessHeap(), 0, lpdwPIDs);
      dwSize2 *= 2 ;
    }
    lpdwPIDs = (LPDWORD)HeapAlloc(GetProcessHeap(), 0, dwSize2);
    if (lpdwPIDs == NULL) {
      FreeLibrary(hInstLib);
      AfxMessageBox("PIDs are NULL");
      return bReturnOk;
    }
    if (!lpfEnumProcesses(lpdwPIDs, dwSize2, &dwSize)) {
      HeapFree(GetProcessHeap(), 0, lpdwPIDs);
      FreeLibrary(hInstLib);
      AfxMessageBox("No processes remaining");
      return bReturnOk;
    }
  } while(dwSize == dwSize2);

  // How many ProcID's did we get?
  dwSize /= sizeof(DWORD);

  // Loop through each ProcID.
  for (dwIndex = 0; dwIndex < dwSize;  dwIndex++) {
    szFileName[0] = 0;
    // Open the process (if we can... security does not permit every process in the system).
    hProcess = OpenProcess(PROCESS_QUERY_INFORMATION | PROCESS_VM_READ, FALSE, lpdwPIDs[dwIndex]);
    if (hProcess != NULL) {
      // Here we call EnumProcessModules to get only the
      // first module in the process this is important,
      // because this will be the .EXE module for which we
      // will retrieve the full path name in a second.
      if (lpfEnumProcessModules(hProcess, &hMod, sizeof(hMod), &dwSize2)) {
        // Get Full pathname:
        if (!lpfGetModuleFileNameEx(hProcess, hMod, szFileName, sizeof(szFileName))) {
          szFileName[0] = 0;
        }
      }
      CloseHandle(hProcess);
    }

    CString strFilename = szFileName;
	int i = 0;
    for (i = strFilename.GetLength() - 1;  i >= 0;  i--) {
      if (strFilename.GetAt(i) == '\\') {
        i++;
        break;
      }
    }
    strFilename = strFilename.Mid(i, strFilename.GetLength() - i);

    CString strMsg;
    strMsg.Format("%ld  == %s", lpdwPIDs[dwIndex], strFilename);

    // Modified to find winword.exe process and close it
    if (strFilename.CompareNoCase("WINWORD.EXE") == 0) {
      if (cIni.GetBool("Config", "Debugging", false)) {
        AfxMessageBox(strMsg);
      }

      if (lpdwPIDs[dwIndex] != 0) {
        HANDLE hProcess = OpenProcess(PROCESS_TERMINATE
                                      | PROCESS_QUERY_INFORMATION
                                      | PROCESS_VM_READ,
                                      FALSE, lpdwPIDs[dwIndex]);

        if (bGracefully) {

          DWORD dwCode;
          int iFeedback = IDRETRY;
          while (GetExitCodeProcess(hProcess, &dwCode) && (dwCode == STILL_ACTIVE) && (iFeedback == IDRETRY)) {

            iFeedback = AfxMessageBox("Merging has detected that there are one or more instances of Microsoft Word open on your computer.  Please close them before clicking 'Retry' and proceeding.", MB_RETRYCANCEL);
          }
          bReturnOk = (iFeedback != IDCANCEL);

        } else {

          if (hProcess != NULL) {   
            SafeTerminateProcess(hProcess, 0, 60000);  // 1 minute
            CloseHandle(hProcess);
          }

        }
      }

    }

  }
  HeapFree(GetProcessHeap(), 0, lpdwPIDs);

  // Free the library.
  FreeLibrary( hInstLib);

  return bReturnOk;
}


// Terminate a running process safely
BOOL SafeTerminateProcess(
	HANDLE hProcess, 
	UINT   uExitCode,
	DWORD  dwTimeout
	)
{
    DWORD dwTID, dwCode, dwErr = 0;
    HANDLE hProcessDup = INVALID_HANDLE_VALUE;
    HANDLE hRT = NULL;
    HINSTANCE hKernel = GetModuleHandle("Kernel32");
    BOOL bSuccess = FALSE;

    BOOL bDup = DuplicateHandle(GetCurrentProcess(), 
                                hProcess, 
                                GetCurrentProcess(), 
                                &hProcessDup, 
                                PROCESS_ALL_ACCESS, 
                                FALSE, 
                                0);

    // Detect the special case where the process is 
    // already dead...
    if ( GetExitCodeProcess((bDup) ? hProcessDup : hProcess, &dwCode) && 
         (dwCode == STILL_ACTIVE) ) 
    {
        FARPROC pfnExitProc;
           
        pfnExitProc = GetProcAddress(hKernel, "ExitProcess");

        hRT = CreateRemoteThread((bDup) ? hProcessDup : hProcess, 
                                 NULL, 
                                 0, 
                                 (LPTHREAD_START_ROUTINE)pfnExitProc,
                                 (PVOID)uExitCode, 0, &dwTID);

        if ( hRT == NULL )
            dwErr = GetLastError();
    }
    else
    {
        dwErr = ERROR_PROCESS_ABORTED;
    }


    if ( hRT )
    {
        // Must wait process to terminate to 
        // guarantee that it has exited...
        DWORD dwResult = WaitForSingleObject(
			(bDup) ? hProcessDup : hProcess,  dwTimeout);

		bSuccess = dwResult == WAIT_OBJECT_0; 

		if (!bSuccess)
			dwErr = ERROR_PROCESS_ABORTED;


        CloseHandle(hRT);
    }

    if ( bDup )
        CloseHandle(hProcessDup);

    if ( !bSuccess )
        SetLastError(dwErr);

    return bSuccess;
}
