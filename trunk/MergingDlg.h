// MergingDlg.h : header file
//

#pragma once


// CMergingDlg dialog
class CMergingDlg : public CDialog
{
// Construction
public:
	CMergingDlg(CWnd* pParent = NULL);	// standard constructor

// Dialog Data
	enum { IDD = IDD_MERGING_DIALOG };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV support


// Implementation
protected:
	HICON m_hIcon;

	// Generated message map functions
	virtual BOOL OnInitDialog();
	afx_msg void OnSysCommand(UINT nID, LPARAM lParam);
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	DECLARE_MESSAGE_MAP()
public:
  afx_msg void OnBnClickedCancel();
  afx_msg void OnBnClickedOk();
  afx_msg void OnBnClickedButton1();
  afx_msg void OnBnClickedButton3();
  CString cstrUrlXmlData;
  CString cstrPathTemplates;
  CString cstrPathMerged;
  afx_msg void OnEnKillfocusEdit2();
  afx_msg void OnEnKillfocusEdit1();
  afx_msg void OnEnKillfocusEdit3();
  virtual INT_PTR DoModal();
};
