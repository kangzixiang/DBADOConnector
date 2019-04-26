#pragma once
/////////////////////////////////////////////////////////////////
//	SfmDBADOConn.h : interface of the CSfmDBADOConn class
//
//	Description:
//
//	Date:		26 April 2019
//	Revision:	1.00
//
//	By:			
//
//
/////////////////////////////////////////////////////////////////
/********************************************
History
13 Apr, 2019
-- integrate loop data base code
*********************************************/
#pragma once
#ifndef SfmDBADOConn_H
#define SfmDBADOConn_H
// If you encounter any compilation problem caused by the following line,
// please add "c:\Program Files\common files\system\ado" to your VC++ IDE
// menu "Tools"->"Options"->"Projects"->"VC++ Directories"->"Include Files"

#ifdef _WIN64
#import "c:\\Program Files\\Common Files\\System\\ado\\msado15.dll" no_namespace \
rename("EOF","adoEOF") rename("DataTypeEnum","adoDataTypeEnum") \
rename("FieldAttributeEnum", "adoFielAttributeEnum") rename("EditModeEnum", "adoEditModeEnum") \
rename("LockTypeEnum", "adoLockTypeEnum") rename("RecordStatusEnum", "adoRecordStatusEnum") \
rename("ParameterDirectionEnum", "adoParameterDirectionEnum")
#else
#import "c:\\Program Files (x86)\\Common Files\\System\\ado\\msado15.dll" no_namespace \
rename("EOF","adoEOF") rename("DataTypeEnum","adoDataTypeEnum") \
rename("FieldAttributeEnum", "adoFielAttributeEnum") rename("EditModeEnum", "adoEditModeEnum") \
rename("LockTypeEnum", "adoLockTypeEnum") rename("RecordStatusEnum", "adoRecordStatusEnum") \
rename("ParameterDirectionEnum", "adoParameterDirectionEnum")
#endif


#ifdef SPC_DATA_EXPORTS
#define SPC_DATA_EXPORTS _declspec(dllexport)
#else
#define SPC_DATA_EXPORTS _declspec(dllimport)
#endif

#include<list>
#include<vector>
using namespace std;

class SPC_DATA_EXPORTS CSfmDBADOConn
{
public:
	CSfmDBADOConn(void);
	virtual~CSfmDBADOConn(void);
	CSfmDBADOConn* CloneObject(void) const;

protected:
	_ConnectionPtr m_pConnection;
	//_RecordsetPtr  m_pRecordset;
	CString        m_szServer;
	CString        m_szDatabase;
	CString        m_szUser;
	CString        m_szPassword;
	CMutex		   m_mutx;

public:
	BOOL OnInitADOConn(const CString& szServer, const CString& szDataBase,
		const CString& szUser = _T("User"), const CString& szPassword = _T(""));
	BOOL OnInitADODB(const CString& szServer, const CString& szDataBase,
		const CString& szUser = _T("User"), const CString& szPassword = _T(""));
	BOOL DeleteDatabase(CString szDataBase, CString szRestorePath);
	bool GetRecordSet(_RecordsetPtr& pRecordset, const _bstr_t& bstrSQL);
	BOOL ExecuteSQL(const _bstr_t& bstrSQL);
	BOOL GetOneRecord(_RecordsetPtr& recordset, char* fieldname, CString& value);
	void CloseConnect();
	static void GetDBFieldItem(_RecordsetPtr& record, CString szfieldName, CString& strFieldValue);
};
class SPParameter
{
private:
	CString m_szName;
	adoDataTypeEnum m_eType;
	adoParameterDirectionEnum m_eDirection;
	int m_nSize;
	CString m_szValue;
public:

	SPParameter(CString szName, adoDataTypeEnum eType, adoParameterDirectionEnum eDirection, int nSize, CString szValue)
	{
		m_szName = szName;
		m_eType = eType;
		m_eDirection = eDirection;
		m_nSize = nSize;
		m_szValue = szValue;
	}

	~SPParameter(void) {}

	CString GetName(void) { return m_szName; }

	adoDataTypeEnum GetType(void) { return m_eType; }

	adoParameterDirectionEnum GetDirection(void) { return m_eDirection; }

	int GetSize(void) { return m_nSize; }

	CString GetValue(void) { return m_szValue; }
};

class  SPC_DATA_EXPORTS DBManage
{
public:
	DBManage(void);
	virtual ~DBManage(void);

protected:
	_ConnectionPtr m_pConnection;
	_CommandPtr  m_pCommand;
	_RecordsetPtr  m_pRecordset;
	_ParameterPtr m_pParam;
	list<SPParameter> ParamList;
	CString m_szServer;
	CString m_szDatabase;
	CString m_szUser;
	CString m_szPassword;

public:

	void ClearParamList();

	int AddParamList(SPParameter para);

	BOOL OpenConnection(const CString& szServer, const CString& szDataBase, const CString& szUser, const CString& szPassword);

	BOOL DBManage::CreateDB(const CString& szServer, const CString& szDataBase, const CString& szUser, const CString& szPassword);

	void CloseConnection();

	void ExecuteSP(const CString& szSPName, const CString& szColumnName, CStringArray& arrResult);

	_RecordsetPtr ExecuteSP(const CString& szSPName);

	BOOL GetRecordsByField(vector<vector<CString>>& table);

	BOOL ExecuteSPNoneQuery(const CString& szSPName);

};

#endif