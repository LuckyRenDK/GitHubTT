// MDWCListOP.cpp : ���� DLL �ĳ�ʼ�����̡�
//

#include "stdafx.h"
#include "MDWCListOP.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif

//
//TODO:  ����� DLL ����� MFC DLL �Ƕ�̬���ӵģ�
//		��Ӵ� DLL �������κε���
//		MFC �ĺ������뽫 AFX_MANAGE_STATE ����ӵ�
//		�ú�������ǰ�档
//
//		����: 
//
//		extern "C" BOOL PASCAL EXPORT ExportedFunction()
//		{
//			AFX_MANAGE_STATE(AfxGetStaticModuleState());
//			// �˴�Ϊ��ͨ������
//		}
//
//		�˺������κ� MFC ����
//		������ÿ��������ʮ����Ҫ��  ����ζ��
//		��������Ϊ�����еĵ�һ�����
//		���֣������������ж������������
//		������Ϊ���ǵĹ��캯���������� MFC
//		DLL ���á�
//
//		�й�������ϸ��Ϣ��
//		����� MFC ����˵�� 33 �� 58��
//

// CMDWCListOPApp

BEGIN_MESSAGE_MAP(CMDWCListOPApp, CWinApp)
END_MESSAGE_MAP()


// CMDWCListOPApp ����

CMDWCListOPApp::CMDWCListOPApp()
{
	// TODO:  �ڴ˴���ӹ�����룬
	// ��������Ҫ�ĳ�ʼ�������� InitInstance ��

}


// Ψһ��һ�� CMDWCListOPApp ����

CMDWCListOPApp theApp;


// CMDWCListOPApp ��ʼ��

/*
�������ƣ�InitInstance()
�������ڣ�20180425 LuckyRen
������������ʼ��
*/
BOOL CMDWCListOPApp::InitInstance()
{
	CWinApp::InitInstance();
	HRESULT hr = CoInitialize(NULL);
	if (FAILED(hr))
	{
		AfxMessageBox(_T("��ʼ��COM��ʧ�ܣ�"));
		return FALSE;
	}
	CWinApp::InitInstance();
	Accessconnect = FALSE;
	GetModuleFileName(NULL, szPath, MAX_PATH);//˫�����ļ�ʱ��λ�ļ�·��
	CString path;
	path = szPath;
	int iIndex = path.ReverseFind('\\');
	szPath[iIndex] = '\0';
	Str_szpath = szPath;
	return TRUE;
}
/*
�������ƣ�ConnectAccess()
�������ڣ�20180425 LuckyRen
�����������������ݿ�
����������
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
			m_pRecordset->CursorLocation = adUseClient;//���ô����Ժ�,GetRecordCout()����ֵ�Ŵ���0
			pSet->CursorLocation = adUseClient;//���ô����Ժ�,GetRecordCout()����ֵ�Ŵ���0
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
�������ƣ�
�������ڣ�
����������
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
			MessageBox(0, _T("�����ݿ����"), _T("���ݿ��ʼ��"), MB_OK);
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
�������ƣ�ExcuteSql()
�������ڣ�20180425 LuckyRen
�������������Ӵ����ݿ�
*/
void CMDWCListOPApp::ExcuteSql(_RecordsetPtr &rs, _variant_t sql)
{
	try
	{
		if (!m_pConnection)
		{
			if (!ConnectAccess())
			{
				AfxMessageBox(TEXT("�������ݿ�ʧ�ܣ�"));
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
�������ƣ�ADOBOMExecute()
�������ڣ�20181128 LuckyRen
�������������Ӵ����ݿ�
*/
BOOL CMDWCListOPApp::ADOBOMExecute(_RecordsetPtr &ADOSet, _variant_t strSQL)
{
	if (ADOSet->State == adStateOpen)
		ADOSet->Close();
	try
	{
		ADOSet->CursorLocation = adUseClient;//���ô����Ժ�,GetRecordCout()����ֵ�Ŵ���

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
�������ƣ�ExitInstance()
�������ڣ�20180425 LuckyRen
����������
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
	CoUninitialize(); //����com
	return CWinApp::ExitInstance();
}


