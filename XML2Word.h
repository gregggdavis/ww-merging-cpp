
BOOL Xml2Word(CString cstrPathMerged,
              CString cstrPathTemplates,
              CString cstrFileName,
              CString cstrCourseName,
              CString cstrAssignmentName,
              CString cstrXmlData);

bool CloseWord(bool bGracefully);

BOOL SafeTerminateProcess(HANDLE hProcess, 
	                        UINT   uExitCode,
	                        DWORD  dwTimeout
	                        );
