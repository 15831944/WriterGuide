
// WriterGuideDlg.cpp : implementation file
//

#include "stdafx.h"
#include "WriterGuide.h"
#include "WriterGuideDlg.h"
#include "afxdialogex.h"
#include "odbcinst.h"
#include "afxdb.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif


// CWriterGuideDlg dialog



CWriterGuideDlg::CWriterGuideDlg(CWnd* pParent /*=NULL*/)
	: CDialogEx(IDD_WRITERGUIDE_DIALOG, pParent)
{
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
}

void CWriterGuideDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_LIST_DATA, m_ListControl);
}

BEGIN_MESSAGE_MAP(CWriterGuideDlg, CDialogEx)
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	ON_BN_CLICKED(IDC_BUTTON_READ, &CWriterGuideDlg::OnBnClickedButtonRead)
	ON_NOTIFY(LVN_ITEMCHANGED, IDC_LIST_DATA, &CWriterGuideDlg::OnLvnItemchangedListData)
	ON_BN_CLICKED(IDC_BUTTON_CREATE, &CWriterGuideDlg::OnBnClickedButtonCreate)
	ON_BN_CLICKED(IDC_BUTTON_UPDATE, &CWriterGuideDlg::OnBnClickedButtonUpdate)
	ON_BN_CLICKED(IDC_BUTTON_DELETE, &CWriterGuideDlg::OnBnClickedButtonDelete)
END_MESSAGE_MAP()


// CWriterGuideDlg message handlers

BOOL CWriterGuideDlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// Set the icon for this dialog.  The framework does this automatically
	//  when the application's main window is not a dialog
	SetIcon(m_hIcon, TRUE);			// Set big icon
	SetIcon(m_hIcon, FALSE);		// Set small icon

	// TODO: Add extra initialization here

	return TRUE;  // return TRUE  unless you set the focus to a control
}

// If you add a minimize button to your dialog, you will need the code below
//  to draw the icon.  For MFC applications using the document/view model,
//  this is automatically done for you by the framework.

void CWriterGuideDlg::OnPaint()
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
		CDialogEx::OnPaint();
	}
}

// The system calls this function to obtain the cursor to display while the user drags
//  the minimized window.
HCURSOR CWriterGuideDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}


void CWriterGuideDlg::OnBnClickedButtonRead()
{
	ShowListControl();
}

// Reset List control
void CWriterGuideDlg::ResetListControl() {
	m_ListControl.DeleteAllItems();
	int iNbrOfColumns = 0;
	CHeaderCtrl* pHeader = (CHeaderCtrl*)m_ListControl.GetDlgItem(0);
	if (pHeader) {
		iNbrOfColumns = pHeader->GetItemCount();
	}
	for (int i = iNbrOfColumns; i >= 0; i--) {
		m_ListControl.DeleteColumn(i);
	}
}

// shows data inside of list control
void CWriterGuideDlg::ShowListControl() {
	CDatabase database, *databaseRef = &database;
	CString SqlString;
	CString strEnglishName, strLocalName, strPastEnglishName, strPastLocalName;
	// You must change above path if it's different
	int iRec = 0;

	TRY{
		// Open the database
	GetDbConnect(databaseRef);

	// Allocate the recordset
	CRecordset recset(&database);

	// Build the SQL statement
	SqlString = "SELECT * FROM [City names$]";

	// Execute the query

	recset.Open(CRecordset::forwardOnly,SqlString,CRecordset::readOnly);
	// Reset List control if there is any data
	ResetListControl();
	// populate Grids
	ListView_SetExtendedListViewStyle(m_ListControl,LVS_EX_GRIDLINES);

	// Column width and heading
	m_ListControl.InsertColumn(0,L"English Name",LVCFMT_LEFT,-1,0);
	m_ListControl.InsertColumn(1,L"Local Name",LVCFMT_LEFT,-1,1);
	m_ListControl.InsertColumn(2, L"Past English Name or Local Name", LVCFMT_LEFT, -1, 1);
	m_ListControl.InsertColumn(3, L"Country", LVCFMT_LEFT, -1, 1);
	m_ListControl.SetColumnWidth(0, 150);
	m_ListControl.SetColumnWidth(1, 150);
	m_ListControl.SetColumnWidth(2, 150);
	m_ListControl.SetColumnWidth(3, 150);

	// Loop through each record
	while (!recset.IsEOF()) {
		// Copy each column into a variable
		int colNum = 0;
		recset.GetFieldValue(L"English name",strEnglishName);
		recset.GetFieldValue(L"Local name",strLocalName);
		recset.GetFieldValue(L"Past English or local names", strPastEnglishName);
		recset.GetFieldValue(L"Country", strPastLocalName);

		// Insert values into the list control
		iRec = m_ListControl.InsertItem(0, strEnglishName, 0);
		m_ListControl.SetItemText(0, 1, strLocalName);
		m_ListControl.SetItemText(0, 2, strPastEnglishName);
		m_ListControl.SetItemText(0, 3, strPastLocalName);

		// goto next record
		recset.MoveNext();
	}
	// Close the database
	database.Close();
	}CATCH(CDBException, e) {
		// If a database exception occured, show error msg
		AfxMessageBox(L"Database error: " + e->m_strError);
	}
	END_CATCH;
}


void CWriterGuideDlg::OnLvnItemchangedListData(NMHDR *pNMHDR, LRESULT *pResult)
{
	LPNMLISTVIEW pNMLV = reinterpret_cast<LPNMLISTVIEW>(pNMHDR);
	// TODO: Add your control notification handler code here
	*pResult = 0;
}


void CWriterGuideDlg::OnBnClickedButtonCreate()
{
	CreateListControl();
	ShowListControl();
}


void CWriterGuideDlg::OnBnClickedButtonUpdate()
{
	UpdateListControl();
	ShowListControl();
}


void CWriterGuideDlg::OnBnClickedButtonDelete()
{
	DeleteListControl();
	ShowListControl();
}

void CWriterGuideDlg::GetDbConnect(CDatabase * database) {
	CString sDsn;
	CString sDriver = L"MICROSOFT EXCEl DRIVER (*.xls)";
	CString sFile = L"C:\\Temp\\City_and_country_names.xls";

	// Build ODBC connection string
	sDsn.Format(L"ODBC;DRIVER={%s};DSN='';DBQ=%s", sDriver, sFile);
	database->Open(NULL, false, false, sDsn);

}

// Create List control
void CWriterGuideDlg::CreateListControl() {

		CDatabase database, * databaseRef = &database;
		CString SqlString;
	
		TRY{
			// Open the database
			GetDbConnect(databaseRef);
	
			SqlString = L"INSERT INTO [City names$] ([English name]) VALUES ('test12')";
			database.ExecuteSQL(SqlString);
			// Close the database
			database.Close();
		}CATCH(CDBException, e) {
			// If a database exception occured, show error msg
			AfxMessageBox(L"Database error: " + e->m_strError);
		}
		END_CATCH;
}

// Update List control
void CWriterGuideDlg::UpdateListControl() {

	CDatabase database, *databaseRef = &database;
	CString SqlString;

	TRY{
		// Open the database
		GetDbConnect(databaseRef);

		SqlString = L"UPDATE [City names$] SET [English name] = '113' WHERE [English name] = '112'";
		database.ExecuteSQL(SqlString);
		// Close the database
		database.Close();
	}CATCH(CDBException, e) {
		// If a database exception occured, show error msg
		AfxMessageBox(L"Database error: " + e->m_strError);
	}
	END_CATCH;
}

// Delete List control
void CWriterGuideDlg::DeleteListControl() {

	CDatabase database, *databaseRef = &database;
	CString SqlString;

	TRY{
		// Open the database
		GetDbConnect(databaseRef);

		SqlString = "DELETE FROM [City names$] WHERE [English name] = '112'";
		database.ExecuteSQL(SqlString);
		// Close the database
		database.Close();
	}CATCH(CDBException, e) {
		// If a database exception occured, show error msg
		AfxMessageBox(L"Database error: " + e->m_strError);
	}
	END_CATCH;
}
