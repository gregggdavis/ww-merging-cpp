// MergingDlg.cpp : implementation file
//

#include "stdafx.h"
#include "Merging.h"
#include "MergingDlg.h"
#include "XML2Word.h"
#include "RetrieveXml.h"
#include "PathDialog.h"
#include "Ini.h"
//#include ".\mergingdlg.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif


CIni cIni(".\\Merging.ini");


// CAboutDlg dialog used for App About

class CAboutDlg : public CDialog
{
public:
	CAboutDlg();

// Dialog Data
	enum { IDD = IDD_ABOUTBOX };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support

// Implementation
protected:
	DECLARE_MESSAGE_MAP()
};

CAboutDlg::CAboutDlg() : CDialog(CAboutDlg::IDD)
{
}

void CAboutDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
}

BEGIN_MESSAGE_MAP(CAboutDlg, CDialog)
END_MESSAGE_MAP()


// CMergingDlg dialog



CMergingDlg::CMergingDlg(CWnd* pParent /*=NULL*/)
	: CDialog(CMergingDlg::IDD, pParent)
  , cstrUrlXmlData(_T(""))
  , cstrPathTemplates(_T(""))
  , cstrPathMerged(_T(""))
{
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
}

void CMergingDlg::DoDataExchange(CDataExchange* pDX)
{
  CDialog::DoDataExchange(pDX);
  DDX_Text(pDX, IDC_EDIT2, cstrUrlXmlData);
  DDX_Text(pDX, IDC_EDIT1, cstrPathTemplates);
  DDX_Text(pDX, IDC_EDIT3, cstrPathMerged);
}

BEGIN_MESSAGE_MAP(CMergingDlg, CDialog)
	ON_WM_SYSCOMMAND()
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	//}}AFX_MSG_MAP
  ON_BN_CLICKED(IDCANCEL, OnBnClickedCancel)
  ON_BN_CLICKED(IDOK, OnBnClickedOk)
  ON_BN_CLICKED(IDC_BUTTON1, OnBnClickedButton1)
  ON_BN_CLICKED(IDC_BUTTON3, OnBnClickedButton3)
  ON_EN_KILLFOCUS(IDC_EDIT2, OnEnKillfocusEdit2)
  ON_EN_KILLFOCUS(IDC_EDIT1, OnEnKillfocusEdit1)
  ON_EN_KILLFOCUS(IDC_EDIT3, OnEnKillfocusEdit3)
END_MESSAGE_MAP()


// CMergingDlg message handlers

BOOL CMergingDlg::OnInitDialog()
{
	CDialog::OnInitDialog();

	// Add "About..." menu item to system menu.

	// IDM_ABOUTBOX must be in the system command range.
	ASSERT((IDM_ABOUTBOX & 0xFFF0) == IDM_ABOUTBOX);
	ASSERT(IDM_ABOUTBOX < 0xF000);

	CMenu* pSysMenu = GetSystemMenu(FALSE);
	if (pSysMenu != NULL)
	{
		CString strAboutMenu;
		strAboutMenu.LoadString(IDS_ABOUTBOX);
		if (!strAboutMenu.IsEmpty())
		{
			pSysMenu->AppendMenu(MF_SEPARATOR);
			pSysMenu->AppendMenu(MF_STRING, IDM_ABOUTBOX, strAboutMenu);
		}
	}

	// Set the icon for this dialog.  The framework does this automatically
	//  when the application's main window is not a dialog
	SetIcon(m_hIcon, TRUE);			// Set big icon
	SetIcon(m_hIcon, FALSE);		// Set small icon

	// TODO: Add extra initialization here


  // Get INI file initial values
  cstrUrlXmlData    = cIni.GetString("Locations", "UrlXmlData",    "");
  cstrPathTemplates = cIni.GetString("Locations", "PathTemplates", "");
  cstrPathMerged    = cIni.GetString("Locations", "PathMerged",    "");
  UpdateData(FALSE);

	return TRUE;  // return TRUE  unless you set the focus to a control
}

void CMergingDlg::OnSysCommand(UINT nID, LPARAM lParam)
{
	if ((nID & 0xFFF0) == IDM_ABOUTBOX)
	{
		CAboutDlg dlgAbout;
		dlgAbout.DoModal();
	}
	else
	{
		CDialog::OnSysCommand(nID, lParam);
	}
}

// If you add a minimize button to your dialog, you will need the code below
//  to draw the icon.  For MFC applications using the document/view model,
//  this is automatically done for you by the framework.

void CMergingDlg::OnPaint() 
{
	if (IsIconic())
	{
		CPaintDC dc(this); // device context for painting

		SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

		// Center icon in client rectangle
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// Draw the icon
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialog::OnPaint();
	}
}

// The system calls this function to obtain the cursor to display while the user drags
//  the minimized window.
HCURSOR CMergingDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}

void CMergingDlg::OnBnClickedCancel()
{
  // TODO: Add your control notification handler code here
  OnCancel();
}

void CMergingDlg::OnBnClickedOk()
{
  UpdateData(TRUE);

  // Verify input paths
  if ((CPathDialog::MakeSurePathExists(cstrPathTemplates) != 0)
  ||  (CPathDialog::MakeSurePathExists(cstrPathMerged)    != 0)) {
    return;
  }


  // Get directory listing of xml data files
  CString cstrDirectoryListing;
  GetDirectoryListing(cstrUrlXmlData, &cstrDirectoryListing);


  // Get each Xml Data Files
  CString cstrXmlFileName;
  CString cstrXmlData;

  int iNumLinksBeforeFiles = cIni.GetInt("Config", "DirListLinksBeforeFiles", 5);
  for (int i = 0;  i < iNumLinksBeforeFiles;  i++) {
    GetNextXmlFileName(&cstrDirectoryListing, &cstrXmlFileName);
  }
  while (GetNextXmlFileName(&cstrDirectoryListing, &cstrXmlFileName)) {

/*
cstrXmlFileName = "gdwm_Gregg_Davis_GC1_GC01.xml";

CFile cFile;
cFile.Open("G:\\Templates\\Merged\\gdwm_Gregg_Davis_GC1_GC01.xml", CFile::modeRead);
char szSampleText[2000];
UINT uiBytesRead = cFile.Read(szSampleText,2000);
szSampleText[uiBytesRead] = '\0';
cFile.Close();
cstrXmlData = szSampleText;
*/


    GetXmlDataFile(cstrUrlXmlData, cstrXmlFileName, cstrPathMerged, &cstrXmlData);

    if (cIni.GetBool("Config", "Debugging", false)) {
      AfxMessageBox(cstrXmlData);
    }

    // Remove DOCTYPE line from data
    cstrXmlData.Replace("<!DOCTYPE merge SYSTEM 'merge.dtd' >", "");
    // Replace &'s with 'and'
    cstrXmlData.Replace("&", "and");

    //Get raw FileName without XML extension
    int iFileName = cstrXmlFileName.Find(".xml");
    CString cstrFileName = cstrXmlFileName.Mid(0, iFileName);
    // Replace %20's with spaces
    cstrFileName.Replace("%20", " ");

    // Get CourseName (between second and third '_')
    int iStudentName = cstrFileName.Find("_") + 1;
    CString cstrStudentName = cstrFileName.Mid(iStudentName, cstrFileName.GetLength());
    int iCourseName = cstrStudentName.Find("_") + 1;
    CString cstrCourseNameAll = cstrStudentName.Mid(iCourseName, cstrStudentName.GetLength());
    int iAssignmentName = cstrCourseNameAll.Find("_") + 1;
    CString cstrCourseName = cstrCourseNameAll.Mid(0, iAssignmentName-1);


    // Get AssignmentName (after third '_')
    CString cstrAssignmentName = cstrCourseNameAll.Mid(iAssignmentName, cstrCourseNameAll.GetLength());


    // Perform Merge
    if (Xml2Word(cstrPathMerged,
                 cstrPathTemplates,
                 cstrFileName,
                 cstrCourseName,
                 cstrAssignmentName,
                 cstrXmlData)) {

      CFile::Rename(cstrPathMerged + "\\" + cstrFileName + "1.DOCX",
                    cstrPathMerged + "\\" + cstrFileName + ".DOCX");

      if (cIni.GetBool("Config", "ShowSuccessful", false)) {
        AfxMessageBox("Successfully Merged " + cstrFileName);
      }

      // Kill the WINWORD.EXE process
      CloseWord(false);

      // Re-Replace spaces with %20's
      cstrFileName.Replace(" ", "%20");

      // Call MoveSuccessfullAssignment.cgi on server
      if (cIni.GetBool("Config", "CallProcessed", true)) {
        CallProcessedXmlCgi(cIni.GetString("Config", "UrlProcessed", "http://www.wwcampus.com/cgi-bin/formhandler/ProcessedXml.pl"), cstrFileName + ".xml");
      }

    } else {
      AfxMessageBox("Failed to Merge " + cstrFileName);
    }
  }
}

//
// TEMPLATES PATH
//
void CMergingDlg::OnBnClickedButton1()
{
  UpdateData(TRUE);

  CString strInitialPath = cstrPathTemplates;
  CString strYourCaption(_T("Select Templates Directory"));
  CString strYourTitle(_T("Templates Directory"));

  CPathDialog dlg(strYourCaption, strYourTitle, strInitialPath);

  if (dlg.DoModal() == IDOK) {
    cstrPathTemplates = dlg.GetPathName();
    UpdateData(FALSE);
    cIni.WriteString("Locations", "PathTemplates", cstrPathTemplates);
  }
}

//
// MERGED ASSIGNMENTS PATH
//
void CMergingDlg::OnBnClickedButton3()
{
  UpdateData(TRUE);

  CString strInitialPath = cstrPathMerged;
  CString strYourCaption(_T("Select Merged Assignments Directory"));
  CString strYourTitle(_T("Merged Assignments Directory"));

  CPathDialog dlg(strYourCaption, strYourTitle, strInitialPath);

  if (dlg.DoModal() == IDOK) {
    cstrPathMerged = dlg.GetPathName();
    UpdateData(FALSE);
    cIni.WriteString("Locations", "PathMerged", cstrPathMerged);
  }
}



void CMergingDlg::OnEnKillfocusEdit2()
{
  UpdateData(TRUE);
  cIni.WriteString("Locations", "UrlXmlData", cstrUrlXmlData);
}

void CMergingDlg::OnEnKillfocusEdit1()
{
  UpdateData(TRUE);
  cIni.WriteString("Locations", "PathTemplates", cstrPathTemplates);
}

void CMergingDlg::OnEnKillfocusEdit3()
{
  UpdateData(TRUE);
  cIni.WriteString("Locations", "PathMerged", cstrPathMerged);
}


INT_PTR CMergingDlg::DoModal()
{
  // TODO: Add your specialized code here and/or call the base class

  return CDialog::DoModal();
}
