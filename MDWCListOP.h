// MDWCListOP.h : MDWCListOP DLL ����ͷ�ļ�
//

#pragma once

#ifndef __AFXWIN_H__
	#error "�ڰ������ļ�֮ǰ������stdafx.h�������� PCH �ļ�"
#endif

#include "resource.h"		// ������


// CMDWCListOPApp
// �йش���ʵ�ֵ���Ϣ������� MDWCListOP.cpp
//

class CMDWCListOPApp : public CWinApp
{
public:
	CMDWCListOPApp();
	_ConnectionPtr m_pConnection;//�������ݿ⣻
	int  ConnectAccess();
	TCHAR szPath[MAX_PATH];
	void ExcuteSql(_RecordsetPtr &rs, _variant_t sql);
	_RecordsetPtr  m_pRecordset, pRecord, pSet;
	CString Str_szpath;
	BOOL Accessconnect;
	BOOL EOFflag;
	CString mVS_EOFPath;
	CString ConnectDirection;
	CString WaterDirection;
	CString UnitIDflag;
	//20181128 LuckRen
	_ConnectionPtr	ADOBomConn;
	_RecordsetPtr	m_pADOBomSerialNO;
	BOOL ADOBOMExecute(_RecordsetPtr &ADOSet, _variant_t strSQL);
	BOOL mFB_ConnectDB(_ConnectionPtr Conn, CString PVS_DBFullName);

// ��д
public:
	virtual BOOL InitInstance();

	DECLARE_MESSAGE_MAP()
	virtual int ExitInstance();
};
extern  CMDWCListOPApp theApp;