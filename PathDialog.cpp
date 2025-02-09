//////////////////////////////////////////////////////////////////////////
//PathDialog.h file
//
//Written by Nguyen Tan Hung <tanhung@yahoo.com>
//////////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "PathDialog.h"
#include <io.h>

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

#define IDC_FOLDERTREE		0x3741
#define IDC_TITLE			0x3742
#define IDC_STATUSTEXT		0x3743

#define IDC_NEW_EDIT_PATH	0x3744

// Class CDlgWnd
BEGIN_MESSAGE_MAP(CPathDialogSub, CWnd)
        ON_BN_CLICKED(IDOK, OnOK)
		ON_EN_CHANGE(IDC_NEW_EDIT_PATH, OnChangeEditPath)
END_MESSAGE_MAP()

void CPathDialogSub::OnOK()
{
	::GetWindowText(::GetDlgItem(m_hWnd, IDC_NEW_EDIT_PATH),
		m_pPathDialog->m_szPathName, MAX_PATH);

	if(CPathDialog::MakeSurePathExists(m_pPathDialog->m_szPathName)==0)
	{
		m_pPathDialog->m_bGetSuccess=TRUE;
		::EndDialog(m_pPathDialog->m_hWnd, IDOK);
	}
	else
	{
		::SetFocus(::GetDlgItem(m_hWnd, IDC_NEW_EDIT_PATH));
	}
}

void CPathDialogSub::OnChangeEditPath()
{
	::GetWindowText(::GetDlgItem(m_hWnd, IDC_NEW_EDIT_PATH),
		m_pPathDialog->m_szPathName, MAX_PATH);
	BOOL bEnableOKButton = (_tcslen(m_pPathDialog->m_szPathName)>0);
	SendMessage(BFFM_ENABLEOK, 0, bEnableOKButton);
}
/////////////////////////////////////////////////////////////////////////////
// CPathDialog dialog


CPathDialog::CPathDialog(LPCTSTR lpszCaption, 
						 LPCTSTR lpszTitle, 
						 LPCTSTR lpszInitialPath, 
						 CWnd* pParent)
{
	m_hWnd=NULL;
	m_PathDialogSub.m_pPathDialog= this;
    m_bParentDisabled = FALSE;

    // Get the true parent of the dialog
    m_pParentWnd = CWnd::GetSafeOwner(pParent);

	m_lpszCaption = lpszCaption;
	m_lpszInitialPath = lpszInitialPath;

	memset(&m_bi, 0, sizeof(BROWSEINFO) );
	m_bi.hwndOwner = (m_pParentWnd==NULL)?NULL:m_pParentWnd->GetSafeHwnd();
	m_bi.pszDisplayName = 0;
	m_bi.pidlRoot = 0;
	m_bi.ulFlags = BIF_RETURNONLYFSDIRS | BIF_STATUSTEXT;
	m_bi.lpfn = BrowseCallbackProc;
	m_bi.lpszTitle = lpszTitle;
}


/////////////////////////////////////////////////////////////////////////////
// CPathDialog message handlers

CString CPathDialog::GetPathName()
{
	return CString(m_szPathName);
}

int CALLBACK CPathDialog::BrowseCallbackProc(HWND hwnd,UINT uMsg,LPARAM lParam, LPARAM pData) 
{
	CPathDialog* pDlg = (CPathDialog*)pData;

	switch(uMsg) 
	{
	case BFFM_INITIALIZED: 
		{
			RECT rc;
			HWND hEdit;
			HFONT hFont;

			pDlg->m_hWnd = hwnd;

			if(pDlg->m_lpszCaption!=NULL)
			{
				::SetWindowText(hwnd, pDlg->m_lpszCaption);
			}

			VERIFY(pDlg->m_PathDialogSub.SubclassWindow(hwnd));
			::ShowWindow(::GetDlgItem(hwnd, IDC_STATUSTEXT), SW_HIDE);
			::GetWindowRect(::GetDlgItem(hwnd, IDC_FOLDERTREE), &rc);
			rc.bottom = rc.top - 4;
			rc.top = rc.bottom - 23;
			::ScreenToClient(hwnd, (LPPOINT)&rc);
			::ScreenToClient(hwnd, ((LPPOINT)&rc)+1);
			hEdit = ::CreateWindowEx(WS_EX_CLIENTEDGE, _T("EDIT"), _T(""),
				WS_CHILD|WS_TABSTOP|WS_VISIBLE|ES_AUTOHSCROLL,
				rc.left, rc.top, 
				rc.right-rc.left, rc.bottom-rc.top, 
				hwnd, NULL, NULL, NULL);
			::SetWindowLong(hEdit, GWL_ID, IDC_NEW_EDIT_PATH);
			::ShowWindow(hEdit, SW_SHOW);

			hFont = (HFONT)::SendMessage(hwnd, WM_GETFONT, 0, 0);
			::SendMessage(hEdit, WM_SETFONT, (WPARAM)hFont, MAKELPARAM(TRUE, 0));

			LPCTSTR lpszPath = pDlg->m_lpszInitialPath;
			TCHAR szTemp[MAX_PATH];
			if(lpszPath==NULL)
			{
				::GetCurrentDirectory(MAX_PATH, szTemp );
				lpszPath = szTemp;
			}
			// WParam is TRUE since you are passing a path.
			// It would be FALSE if you were passing a pidl.
			::SendMessage(hwnd,BFFM_SETSELECTION,TRUE,
				(LPARAM)lpszPath);
			break;
		}
	case BFFM_SELCHANGED:
		{
      BOOL failed = FALSE;
      char szSelection[MAX_PATH];
      char szLwrSelection[MAX_PATH];

      szLwrSelection[0] = NULL; 

      // Win2000: nethood is the name of folder with 
      // shortcuts to other network resourses. The 
      // folders are displayed, however they are
      // only links, and returned string must not be 
      // treated as a folder.
      if(::SHGetPathFromIDList(
         (LPITEMIDLIST)lParam, 
         szSelection))
      {
         strcpy(szLwrSelection, szSelection);

         //Make lower-case for comparison
         strlwr(szLwrSelection);
      }

      if (szLwrSelection[0] != NULL &&
         (strstr(szLwrSelection, "nethood\\") == NULL))
      {
         // Valid folder
         ::SendMessage(hwnd, BFFM_ENABLEOK, 0, TRUE);
      }
      else
      {
         //Invalid folder
         szSelection[0] = '\0';
         ::SendMessage(hwnd, BFFM_ENABLEOK, 0, FALSE);
      }
      ::SendMessage(
         hwnd,
         BFFM_SETSTATUSTEXT,
         0,
         (LPARAM)szSelection);

      ::SetWindowText(
         ::GetDlgItem(hwnd, IDC_NEW_EDIT_PATH), 
         szSelection);
      break;
		}
	default:
		break;
	}
	return 0;
}



int CPathDialog::DoModal() 
{

	/////////////////////////////////////////////////////////
	TCHAR szPathTemp[MAX_PATH];
    m_bi.lpfn = BrowseCallbackProc;  // address of callback function
    m_bi.lParam = (LPARAM)this;      // pass address of object to callback function
	m_bi.pszDisplayName = szPathTemp;

	LPITEMIDLIST pidl;
	LPMALLOC pMalloc;

	int iResult=-1;
	if(SUCCEEDED(SHGetMalloc(&pMalloc)))
	{
//		if((m_pParentWnd!=NULL) && m_pParentWnd->IsWindowEnabled())
//		{
//		  m_pParentWnd->EnableWindow(FALSE);
//		  m_bParentDisabled = TRUE;
//		}
		m_bGetSuccess = FALSE;
		pidl = SHBrowseForFolder(&m_bi);
		if (pidl!=NULL) 
		{
			//not need do this because OnOK function did
			//bSucceeded = SHGetPathFromIDList(pidl, m_szPathName);
			// In C++: 
			pMalloc->Free(pidl);
			//In C:
			//pMalloc->lpVtbl->Free(pMalloc,pidl);
			//pMalloc->lpVtbl->Release(pMalloc);
		}
		if(m_bGetSuccess)
		{
			iResult = IDOK;
		}
		pMalloc->Release();
	}

    if(m_bParentDisabled && (m_pParentWnd!=NULL))
	{
		m_pParentWnd->EnableWindow(TRUE);
	}
    m_bParentDisabled=FALSE;

	return iResult;
}

BOOL CPathDialog::IsFileNameValid(LPCTSTR lpFileName)
{
	if(lpFileName==NULL)
	{
		return FALSE;
	}

	int nLen = (int)_tcslen(lpFileName);
	if(nLen<=0)
	{
		return FALSE;
	}

	//check first char
	switch(lpFileName[0])
	{
	case _T('.'):
	case _T(' '):
	case _T('\t'):
		return FALSE;
	}

	//check last char
	switch(lpFileName[nLen-1])
	{
	case _T('.'):
	case _T(' '):
	case _T('\t'):
		return FALSE;
	}

	//check all
	int i=0;
	while(lpFileName[i]!=0)
	{
		switch(lpFileName[i])
		{
		case _T('\\'):
		case _T('/'):
		case _T(':'):
		case _T('*'):
		case _T('?'):
		case _T('\"'):
		case _T('<'):
		case _T('>'):
		case _T('|'):
			return FALSE;
		}
		i++;
	}


	return TRUE;
}

const TCHAR c_FolderDoesNotExist[] = _T(
		"The folder:\n\n"
		"%s\n\n"
		"does not exist. Do you want the folder to be created?");
const TCHAR c_szErrInvalidPath[] = _T(
		"The folder:"
		"\n\n"
		"%s\n\n"
		"is invalid. Please reenter.");
const TCHAR c_szErrCreatePath[] = _T(
		"The folder:"
		"\n\n"
		"%s"
		"\n\ncan not be created. Please double check.");

//return -1: user break;
//return 0: no error
//return 1: lpPath is invalid
//return 2: can not create lpPath
int CPathDialog::MakeSurePathExists(LPCTSTR lpPath)
{
	CString strMsg;
	int iRet;
	try
	{

#if 0
    //validate path
    iRet=Touch(lpPath, TRUE);
    if(iRet!=0)
    {
       throw iRet;
    }
#endif


		if(_taccess(lpPath, 0)==0)
		{
			return (int)0;
		}

		strMsg.Format(c_FolderDoesNotExist, lpPath);
		if(AfxMessageBox(strMsg, MB_YESNO|MB_ICONQUESTION) != IDYES)
		{
			return (int)-1;
		}

		//create path
		iRet=Touch(lpPath, FALSE);
		if(iRet!=0)
		{
			throw iRet; 
		}

		return 0;
	}
	catch(int nErrCode)
	{
		switch(nErrCode)
		{
		case 1:
			strMsg.Format(c_szErrInvalidPath, lpPath);
			break;
		case 2:
		default:
			strMsg.Format(c_szErrCreatePath, lpPath);
			break;
		}

		AfxMessageBox(strMsg, MB_OK|MB_ICONEXCLAMATION);

	}

	return iRet;
}

//return 0: no error
//return 1: lpPath is invalid
//return 2: lpPath can not be created(bValidate==FALSE)
int CPathDialog::Touch(LPCTSTR lpPath, BOOL bValidate)
{
	if(lpPath==NULL)
	{
		return 1;
	}

	TCHAR szPath[MAX_PATH];
	_tcscpy(szPath, lpPath);
	int nLen = (int)_tcslen(szPath);

	//path must be "x:\..."
	if( ( nLen<3 ) || 
		( ( szPath[0]<_T('A') || _T('Z')<szPath[0] ) && 
		  ( szPath[0]<_T('a') || _T('z')<szPath[0] ) ||
		( szPath[1]!=_T(':') )|| 
		( szPath[2]!=_T('\\') )
		)
	  )
	{
		return 1;
	}

	int i;
	if(nLen==3)
	{
		if(!bValidate)
		{
			if(_access(szPath, 0)!=0)
			{
				return 2;
			}
		}
		return 0;
	}

	i = 3;
	BOOL bLastOne=TRUE;
	LPTSTR lpCurrentName;
	while(szPath[i]!=0)
	{
		lpCurrentName = &szPath[i];
		while( (szPath[i]!=0) && (szPath[i]!=_T('\\')) )
		{
			i++;
		}

		bLastOne =(szPath[i]==0);
		szPath[i] = 0;

		if( !IsFileNameValid(lpCurrentName) )
		{
			return 1;
		}

		if(!bValidate)
		{
			CreateDirectory(szPath, NULL);
			if(_taccess(szPath, 0)!=0)
			{
				return 2;
			}
		}

		if(bLastOne)
		{
			break; //it's done
		}
		else
		{
			szPath[i] = _T('\\');
		}

		i++;
	}

	return (bLastOne?0:1);
}

//return 0: ok
//return 1: error
int CPathDialog::ConcatPath(LPTSTR lpRoot, LPCTSTR lpMorePath)
{
	if(lpRoot==NULL)
	{
		return 1;
	}

	int nLen = (int)_tcslen(lpRoot);

	if(nLen<3)
	{
		return 1;
	}

	if(lpMorePath==NULL)
	{
		return 0;
	}

	if(nLen==3)
	{
		_tcscat(lpRoot, lpMorePath);
		return 0;
	}

	_tcscat(lpRoot, _T("\\"));
	_tcscat(lpRoot, lpMorePath);

	return 0;
}
