// MDWCListOP.h : MDWCListOP DLL 的主头文件
//

#pragma once

#ifndef __AFXWIN_H__
	#error "在包含此文件之前包含“stdafx.h”以生成 PCH 文件"
#endif

#include "resource.h"		// 主符号


// CMDWCListOPApp
// 有关此类实现的信息，请参阅 MDWCListOP.cpp
//

class CMDWCListOPApp : public CWinApp
{
public:
	CMDWCListOPApp();
	_ConnectionPtr m_pConnection;//连接数据库；
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

// 重写
public:
	virtual BOOL InitInstance();

	DECLARE_MESSAGE_MAP()
	virtual int ExitInstance();
};
extern  CMDWCListOPApp theApp;