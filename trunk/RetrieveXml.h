
bool GetDirectoryListing(CString cstrHttpPath, CString *pcstrDirectoryListing);

bool GetNextXmlFileName(CString *pcstrDirectoryListing, CString *pcstrXmlFileName);

void GetXmlDataFile(CString cstrUrlXmlData,
                    CString cstrXmlFileName, 
                    CString cstrPathMerged,
                    CString *pcstrXMlData);

void CallProcessedXmlCgi(CString cstrHttpPathName, CString cstrXmlFileName);
