#include "stdafx.h"

#include "RetrieveXml.h"
#include "Request.h"

#include "Ini.h"
extern CIni cIni;


bool GetDirectoryListing(CString cstrHttpPath, CString *pcstrDirectoryListing)
{
  cstrHttpPath = cstrHttpPath + "/";

  // Retrieve XML data files
  CString  cstrHeaderSend    = "";
  CString  cstrHeaderReceive = "";
  CString  cstrMessage       = "";

  Request  myRequest;
  myRequest.SendRequest(false,
                        (LPCTSTR)cstrHttpPath,
                        cstrHeaderSend,
                        cstrHeaderReceive,
                        cstrMessage);
  if (cIni.GetBool("Config", "Debugging", false)) {
    AfxMessageBox(cstrMessage);
  }
  *pcstrDirectoryListing = cstrMessage;
  return(cstrMessage.GetLength() > 0);
}



bool GetNextXmlFileName(CString *pcstrDirectoryListing, CString *pcstrXmlFileName)
{
  // Get "<A...HREF=\""
  int iHrefTag = (*pcstrDirectoryListing).Find("HREF=");
  if (iHrefTag < 0) {
	iHrefTag = (*pcstrDirectoryListing).Find("href=");
	if (iHrefTag < 0) {
		return 0;
	}
  }
  CString cstrHrefTag = (*pcstrDirectoryListing).Mid(iHrefTag+6);

  // Get "\">" (end quote)
  int iHrefTagEnd = cstrHrefTag.Find(">");
  CString cstrHrefTagDataAndListing = cstrHrefTag.Mid(iHrefTagEnd);

  // Save the rest of the Dir listing
  *pcstrDirectoryListing = cstrHrefTagDataAndListing;

  // Get data in HREF=\"data\"
  *pcstrXmlFileName = cstrHrefTag.Mid(0, iHrefTagEnd-1);

  return((*pcstrXmlFileName).GetLength() > 0);
}



void GetXmlDataFile(CString cstrUrlXmlData,
                    CString cstrXmlFileName, 
                    CString cstrPathMerged,
                    CString *pcstrXMlData)
{
  cstrUrlXmlData = cstrUrlXmlData + "/" + cstrXmlFileName;

  // Retrieve XML data files
  CString  cstrHeaderSend    = "";
  CString  cstrHeaderReceive = "";
  CString  cstrMessage       = "";

  Request  myRequest;
  myRequest.SendRequest(false,
                        (LPCTSTR)cstrUrlXmlData,
                        cstrHeaderSend,
                        cstrHeaderReceive,
                        cstrMessage);

  *pcstrXMlData = cstrMessage;
}




void CallProcessedXmlCgi(CString cstrHttpPathName, CString cstrXmlFileName)
{
  CString  cstrHeaderSend    = "";
  CString  cstrHeaderReceive = "";
  CString  cstrMessage       = "";

  cstrHttpPathName = cstrHttpPathName + "?filename=" + cstrXmlFileName;

  if (cIni.GetBool("Config", "Debugging", false)) {
    AfxMessageBox(cstrHttpPathName);
  }

  Request  myRequest;
  myRequest.SendRequest(false,
                        (LPCTSTR)cstrHttpPathName,
                        cstrHeaderSend,
                        cstrHeaderReceive,
                        cstrMessage);
}



