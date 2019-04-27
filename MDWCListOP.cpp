// MDWCListOP.cpp : 定义 DLL 的初始化例程。
//

#include "stdafx.h"
#include "MDWCListOP.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif

//
//TODO:  如果此 DLL 相对于 MFC DLL 是动态链接的，
//		则从此 DLL 导出的任何调入
//		MFC 的函数必须将 AFX_MANAGE_STATE 宏添加到
//		该函数的最前面。
//
//		例如: 
//
//		extern "C" BOOL PASCAL EXPORT ExportedFunction()
//		{
//			AFX_MANAGE_STATE(AfxGetStaticModuleState());
//			// 此处为普通函数体
//		}
//
//		此宏先于任何 MFC 调用
//		出现在每个函数中十分重要。  这意味着
//		它必须作为函数中的第一个语句
//		出现，甚至先于所有对象变量声明，
//		这是因为它们的构造函数可能生成 MFC
//		DLL 调用。
//
//		有关其他详细信息，
//		请参阅 MFC 技术说明 33 和 58。
//

// CMDWCListOPApp

BEGIN_MESSAGE_MAP(CMDWCListOPApp, CWinApp)
END_MESSAGE_MAP()


// CMDWCListOPApp 构造

CMDWCListOPApp::CMDWCListOPApp()
{
	// TODO:  在此处添加构造代码，
	// 将所有重要的初始化放置在 InitInstance 中

}


// 唯一的一个 CMDWCListOPApp 对象

CMDWCListOPApp theApp;


// CMDWCListOPApp 初始化

/*
函数名称：InitInstance()
创建日期：20180425 LuckyRen
功能描述：初始化
*/
BOOL CMDWCListOPApp::InitInstance()
{
	CWinApp::InitInstance();
	HRESULT hr = CoInitialize(NULL);
	if (FAILED(hr))
	{
		AfxMessageBox(_T("初始化COM库失败！"));
		return FALSE;
	}
	CWinApp::InitInstance();
	Accessconnect = FALSE;
	GetModuleFileName(NULL, szPath, MAX_PATH);//双击打开文件时定位文件路径
	CString path;
	path = szPath;
	int iIndex = path.ReverseFind('\\');
	szPath[iIndex] = '\0';
	Str_szpath = szPath;
	return TRUE;
}
/*
函数名称：ConnectAccess()
创建日期：20180425 LuckyRen
功能描述：连接数据库
参数描述：
*/
int  CMDWCListOPApp::ConnectAccess()
{
	HRESULT hr;
	try
	{
		hr = m_pConnection.CreateInstance(__uuidof(Connection));

		if (SUCCEEDED(hr))
		{
			
			CString szConnect;
			szConnect = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=";
			szConnect += szPath;
			szConnect = szConnect + _T("\\MDWListOP.tbs;Persist Security Info=FALSE;Jet OLEDB:Database Password=#datacofahu#");
			m_pConnection->Open((LPCTSTR)szConnect, _T(""), _T(""), adModeUnknown);
			m_pRecordset.CreateInstance(__uuidof(Recordset));

			m_pADOBomSerialNO.CreateInstance(__uuidof(Recordset));
			pRecord.CreateInstance(__uuidof(Recordset));
			pSet.CreateInstance(__uuidof(Recordset));
			m_pRecordset->CursorLocation = adUseClient;//设置此属性后,GetRecordCout()返回值才大于0
			pSet->CursorLocation = adUseClient;//设置此属性后,GetRecordCout()返回值才大于0
		}
	}
	catch (_com_error &e)
	{
		AfxMessageBox(e.Description());
		return 0;
	}
	return 1;
}
/*
函数名称：
创建日期：
功能描述：
*/
BOOL CMDWCListOPApp::mFB_ConnectDB(_ConnectionPtr Conn, CString PVS_DBFullName)
{
	try{
		if (Conn->State)
			Conn->Close();

		Conn->ConnectionTimeout = 20;
		HRESULT hr = Conn->Open("Provider=SQLOLEDB.1;Persist Security Info=False;User ID=mdmcd;Password=mdmcd20090722;	\
																			Initial Catalog=" + _bstr_t(PVS_DBFullName) + ";Data Source=" + "192.168.16.20", "", "", adOpenUnspecified);
		if (FAILED(hr))
		{
			MessageBox(0, _T("打开数据库出错！"), _T("数据库初始化"), MB_OK);
			return FALSE;
		}
	}
	catch (_com_error e)
	{
		_bstr_t LVS_Source = e.Source();
		_bstr_t LVS_Cause = e.Description();

		return FALSE;
	}
	return TRUE;
}
/*
函数名称：ExcuteSql()
创建日期：20180425 LuckyRen
功能描述：连接打开数据库
*/
void CMDWCListOPApp::ExcuteSql(_RecordsetPtr &rs, _variant_t sql)
{
	try
	{
		if (!m_pConnection)
		{
			if (!ConnectAccess())
			{
				AfxMessageBox(TEXT("连接数据库失败！"));
				return;
			}
		}
		if (rs->GetState() == adStateOpen)
			rs->Close();
		rs->Open(sql, _variant_t((IDispatch*)m_pConnection, true), adOpenStatic, adLockOptimistic, adCmdText);
	}
	catch (_com_error &e)
	{
		AfxMessageBox(e.Description());
	}
}
/*
函数名称：ADOBOMExecute()
创建日期：20181128 LuckyRen
功能描述：连接打开数据库
*/
BOOL CMDWCListOPApp::ADOBOMExecute(_RecordsetPtr &ADOSet, _variant_t strSQL)
{
	if (ADOSet->State == adStateOpen)
		ADOSet->Close();
	try
	{
		ADOSet->CursorLocation = adUseClient;//设置此属性后,GetRecordCout()返回值才大于

		ADOSet->Open(strSQL, ADOBomConn.GetInterfacePtr(), \
			adOpenStatic, adLockOptimistic, adCmdUnknown);
		return TRUE;
	}
	catch (_com_error &e)
	{
		CString err;
		err.Format(_T("ADO Error:%s"), (char*)e.Description());
		AfxMessageBox(err);
		return FALSE;
	}
}
/*******************************************************************************************/
/*
函数名称：ExitInstance()
创建日期：20180425 LuckyRen
功能描述：
*/
int CMDWCListOPApp::ExitInstance()
{
	if (m_pRecordset)
	{
		if (m_pRecordset->State)
			m_pRecordset->Close();
	}
	if (pSet)
	{
		if (pSet->State)
			pSet->Close();
	}
	CoUninitialize(); //结束com
	return CWinApp::ExitInstance();
}


