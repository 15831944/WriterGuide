
// WriterGuideDlg.h : header file
//

#pragma once
#include "afxcmn.h"
#include "odbcinst.h"
#include "afxdb.h"


// CWriterGuideDlg dialog
class CWriterGuideDlg : public CDialogEx
{
// Construction
public:
	CWriterGuideDlg(CWnd* pParent = NULL);	// standard constructor

// Dialog Data
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_WRITERGUIDE_DIALOG };
#endif

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV support

private:
	void ResetListControl();
	void ShowListControl();
	void CreateListControl();
	void UpdateListControl();
	void DeleteListControl();
	void GetDbConnect(CDatabase * database); 



// Implementation
protected:
	HICON m_hIcon;

	// Generated message map functions
	virtual BOOL OnInitDialog();
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	DECLARE_MESSAGE_MAP()
public:
	afx_msg void OnBnClickedButtonRead();
	CListCtrl m_ListControl;
	afx_msg void OnLvnItemchangedListData(NMHDR *pNMHDR, LRESULT *pResult);
	afx_msg void OnBnClickedButtonRead4();
	afx_msg void OnBnClickedButtonCreate();
	afx_msg void OnBnClickedButtonUpdate();
	afx_msg void OnBnClickedButtonDelete();


};
