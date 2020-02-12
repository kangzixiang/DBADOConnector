/////////////////////////////////////////////////////////////////
//	SfmDBADOConn.cpp : implementation file
//
//	Description:
//
//	Date:		26 April 2019
//	Revision:	1.00
//
//	By:		Kang	
//
/////////////////////////////////////////////////////////////////
#include "stdafx.h"

#include "SfmDBADOConn.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif 

CSfmDBADOConn::CSfmDBADOConn(void)
{
	m_szServer = _T("");
	m_szDatabase = _T("");
	m_szUser = _T("");
	m_szPassword = _T("");

	m_pConnection = NULL;
	//m_pRecordset = NULL;
}

CSfmDBADOConn::~CSfmDBADOConn(void)
{
	this->CloseConnect();
}

CSfmDBADOConn* CSfmDBADOConn::CloneObject(void) const
{
	CSfmDBADOConn* pObject = new CSfmDBADOConn(*this);
	return pObject;
}

BOOL CSfmDBADOConn::OnInitADOConn(const CString& szServer, const CString& szDataBase,
	const CString& szUser, const CString& szPassword)
{
	::CoInitialize(NULL);

	try
	{
		m_szServer = szServer;
		m_szDatabase = szDataBase;
		m_szUser = szUser;
		m_szPassword = szPassword;

		HRESULT hr = m_pConnection.CreateInstance("ADODB.Connection");

		if (hr != S_OK || m_pConnection == NULL)
		{
			return FALSE;
		}

		CString szConn;
		szConn = "Driver={SQL Server};Server=" + szServer + ";Database=" + szDataBase + ";Uid=" + szUser + ";Pwd=" + szPassword + ";";

		m_pConnection->Open(_bstr_t(szConn), _bstr_t(szUser), _bstr_t(szPassword), adModeUnknown);
		//m_pConnection->CursorLocation = CursorLocationEnum::adUseClient;
		return TRUE;
	}
	catch (_com_error e)
	{
		AfxMessageBox(e.Description());
		m_pConnection = NULL;
		return FALSE;
	}

	return TRUE;
}

BOOL CSfmDBADOConn::OnInitADODB(const CString& szServer, const CString& szDataBase,
	const CString& szUser, const CString& szPassword)
{
	::CoInitialize(NULL);

	try
	{
		m_szServer = szServer;
		m_szDatabase = szDataBase;
		m_szUser = szUser;
		m_szPassword = szPassword;
		CString szSystemDB = "master";
		CString szConn = "";

		HRESULT hr = m_pConnection.CreateInstance("ADODB.Connection");
		szConn = "Driver={SQL Server};Server=" + szServer + ";Database=" + szSystemDB + ";Uid=" + szUser + ";Pwd=" + szPassword + ";";

		if (hr != S_OK || m_pConnection == NULL)
		{
			return FALSE;
		}

		m_pConnection->Open(_bstr_t(szConn), _bstr_t(szUser), _bstr_t(szPassword), adModeUnknown);

		CString	szSql = "IF  NOT EXISTS (SELECT name FROM sys.databases WHERE name = N";
		//szSql = szSql + _T("'") + szDataBase + _T("')");
		//szSql = szSql + _T(" ");
		//szSql = szSql + "CREATE DATABASE ";
		//szSql = szSql + szDataBase;
		szSql.Append(_T("'"));
		szSql.Append(szDataBase);
		szSql.Append(_T("')"));
		szSql.Append(_T(" "));
		szSql.Append("CREATE DATABASE ");
		szSql.Append(szDataBase);

		_bstr_t bstrtmp = (_bstr_t)szSql;
		m_pConnection->Execute(bstrtmp, NULL, adCmdText);
		SysFreeString(bstrtmp);
		m_pConnection->Close();

		szConn = "Driver={SQL Server};Server=" + szServer + ";Database=" + szDataBase + ";Uid=" + szUser + ";Pwd=" + szPassword + ";";
		m_pConnection->Open(_bstr_t(szConn), _bstr_t(szUser), _bstr_t(szPassword), adModeUnknown);
		//m_pConnection->CursorLocation = CursorLocationEnum::adUseClient;

		return TRUE;
	}
	catch (_com_error e)
	{
		m_pConnection = NULL;
		return FALSE;
	}
}

BOOL CSfmDBADOConn::DeleteDatabase(CString szDataBase, CString szRestorePath) //3.2.4.5
{
	if (m_pConnection != NULL)
	{
		CloseConnect();
	}
	try
	{
		CString szSystemDB = "master";
		CString szConn = "";

		HRESULT hr = m_pConnection.CreateInstance("ADODB.Connection");
		szConn = "Driver={SQL Server};Server=" + m_szServer + ";Database=" + szSystemDB + ";Uid=" + m_szUser + ";Pwd=" + m_szPassword + ";";

		if (hr != S_OK || m_pConnection == NULL)
		{
			return FALSE;
		}

		m_pConnection->Open(_bstr_t(szConn), _bstr_t(m_szUser), _bstr_t(m_szPassword), adModeUnknown);

		CString	szSql = "IF  EXISTS (SELECT name FROM sys.databases WHERE name = N";
		//szSql = szSql + _T("'") + szDataBase + _T("')");
		//szSql = szSql + _T(" ");
		//szSql = szSql + "DROP DATABASE ";
		//szSql = szSql + szDataBase;
		szSql.Append(_T("'"));
		szSql.Append(szDataBase);
		szSql.Append(_T("')"));
		szSql.Append(_T(" "));
		szSql.Append("DROP DATABASE ");
		szSql.Append(szDataBase);

		m_pConnection->Execute((_bstr_t)szSql, NULL, adCmdText);

		CFileFind filefind;
		CString szFile = "C:\\Program Files\\Microsoft SQL Server\\MSSQL.1\\MSSQL\\Data\\" + m_szDatabase + ".mdf";
		if (filefind.FindFile(szFile))
		{
			DeleteFile(szFile);
		}

		szFile = "C:\\Program Files\\Microsoft SQL Server\\MSSQL.1\\MSSQL\\Data\\" + m_szDatabase + "_log.LDF";
		if (filefind.FindFile(szFile))
		{
			DeleteFile(szFile);
		}

		szSql = "RESTORE DATABASE [";
		//szSql = szSql + szDataBase;
		//szSql = szSql + "] FROM DISK = '";
		//szSql = szSql + szRestorePath + szDataBase + ".bak'";
		szSql.Append(szDataBase);
		szSql.Append("] FROM DISK = '");
		szSql.Append(szRestorePath);
		szSql.Append(szDataBase);
		szSql.Append(".bak'");

		m_pConnection->Execute((_bstr_t)szSql, NULL, adCmdText);
		m_pConnection->Close();
		m_pConnection = NULL;

		return TRUE;
	}
	catch (_com_error e)
	{
		m_pConnection = NULL;
		return FALSE;
	}
}

bool CSfmDBADOConn::GetRecordSet(_RecordsetPtr& pRecordset, const _bstr_t& bstrSQL)
{
	m_mutx.Lock();
	TRACE1("%d S_CSfmDBADOConn::GetRecordSet\n", GetTickCount()); //TIMTTEST

	if (m_pConnection == NULL)
	{
		return false;
	}

	try
	{
		if (m_pConnection->GetState() == adStateClosed)
		{
			OnInitADOConn(m_szServer, m_szDatabase, m_szUser, m_szPassword);
		}

		HRESULT hr = pRecordset.CreateInstance(__uuidof(Recordset));
		if (hr == S_OK && pRecordset != NULL)
		{
			pRecordset->Open(bstrSQL, m_pConnection.GetInterfacePtr(), adOpenForwardOnly, /*adOpenDynamic,*/
				adLockOptimistic, adCmdText);
		}
	}
	catch (_com_error e)
	{
		pRecordset = NULL;
		//if(pRecordset!=NULL)
		//{
		//	if(pRecordset->GetState()==adStateOpen)
		//		pRecordset->Close();
		//	pRecordset = NULL;
		//}
	}

	TRACE1("%d E_CSfmDBADOConn::GetRecordSet\n", GetTickCount()); //TIMTTEST
	m_mutx.Unlock();
	return true;
}

BOOL CSfmDBADOConn::ExecuteSQL(const _bstr_t& bstrSQL)
{
	m_mutx.Lock();
	if (m_pConnection == NULL)
	{
		m_mutx.Unlock();
		return FALSE;
	}

	try
	{
		if (m_pConnection->GetState() == adStateClosed)
		{
			OnInitADOConn(m_szServer, m_szDatabase, m_szUser, m_szPassword);
		}

		m_pConnection->Execute(bstrSQL, NULL, adCmdText);

		m_mutx.Unlock();
		return TRUE;
	}
	catch (_com_error e)
	{
		m_mutx.Unlock();
		return FALSE;
	}
}

BOOL CSfmDBADOConn::GetOneRecord(_RecordsetPtr& recordset, char* fieldname, CString& value)
{
	try
	{
		_bstr_t		record;
		record = recordset->GetCollect(fieldname);

		value.Format("%s", (char*)record);
		SysFreeString(record);
		return TRUE;
	}
	catch (_com_error e)
	{
		value = _T("");
		return FALSE;
	}
}

void CSfmDBADOConn::CloseConnect()
{
	if (m_pConnection != NULL && m_pConnection->GetState() != adStateClosed)
	{
		m_pConnection->Close();
		m_pConnection = NULL;
	}

	::CoUninitialize();
}

void CSfmDBADOConn::GetDBFieldItem(_RecordsetPtr& record, CString szfieldName, CString& strFieldValue)
{
	COleVariant olevar;
	_bstr_t fieldName = szfieldName;
	olevar = record->Fields->GetItem(fieldName)->Value;

	if (olevar.vt != VT_EMPTY && olevar.vt == VT_BSTR)
	{
		BSTR fieldValue = olevar.Detach().bstrVal;
		strFieldValue = fieldValue;
		::SysFreeString(fieldValue);
	}
	else if (olevar.vt != VT_EMPTY && olevar.vt == VT_DATE)
	{
		COleDateTime oledate = olevar.Detach();
		strFieldValue = oledate.Format(_T("%Y-%m-%d %H:%M:%S"));
	}
	else if ((olevar.vt == VT_NULL || olevar.vt == VT_EMPTY))
	{
		strFieldValue = "";
	}
	::SysFreeString(fieldName);
}

//////////////////////////////////////////////////////////////////////////////////
//////////////////////////////////////////////////////////////////////////////////

DBManage::DBManage(void)
{
	::CoInitialize(NULL);
	m_szServer = _T("");
	m_szDatabase = _T("");
	m_szUser = _T("");
	m_szPassword = _T("");
	m_pConnection = NULL;
	m_pCommand = NULL;
	m_pRecordset = NULL;
}

DBManage::~DBManage(void)
{
	//CloseConnection();
}

int DBManage::AddParamList(SPParameter para)
{
	ParamList.push_back(para);
	return ParamList.size();
}

void DBManage::ClearParamList()
{
	ParamList.clear();
}

BOOL DBManage::OpenConnection(const CString& szServer, const CString& szDataBase, const CString& szUser, const CString& szPassword)
{
	try
	{
		m_szServer = szServer;
		m_szDatabase = szDataBase;
		m_szUser = szUser;
		m_szPassword = szPassword;
		HRESULT hr = m_pConnection.CreateInstance("ADODB.Connection");
		if (hr != S_OK || m_pConnection == NULL)
		{
			return FALSE;
		}
		CString szConn;
		szConn = _T("Driver={SQL Server};Server=") + szServer + _T(";Database=") + szDataBase + _T(";Uid=") + szUser + _T(";Pwd=") + szPassword + _T(";");
		m_pConnection->Open(_bstr_t(szConn), _bstr_t(szUser), _bstr_t(szPassword), adModeUnknown);
		return TRUE;
	}
	catch (_com_error e)
	{
		AfxMessageBox(e.Description());
		if (m_pConnection != NULL)
		{
			//m_pConnection->Close();
			m_pConnection = NULL;
		}
		return FALSE;
	}
}

void DBManage::CloseConnection()
{
	m_szServer = _T("");
	m_szDatabase = _T("");
	m_szUser = _T("");
	m_szPassword = _T("");
	if (m_pRecordset != NULL)
	{
		m_pRecordset->Close();
		m_pRecordset = NULL;
	}
	if (m_pConnection != NULL)
	{
		m_pConnection->Close();
		m_pConnection = NULL;
	}
	::CoUninitialize();
}

void DBManage::ExecuteSP(const CString& szSPName, const CString& szColumnName, CStringArray& arrResult)
{
	if (m_pConnection == NULL)
	{
		return;
	}
	try
	{
		if (m_pConnection->GetState() == adStateClosed)
		{
			OpenConnection(m_szServer, m_szDatabase, m_szUser, m_szPassword);
		}
		HRESULT hr = m_pCommand.CreateInstance(__uuidof(Command));
		m_pRecordset.CreateInstance(__uuidof(Recordset));

		if (hr != S_OK || m_pCommand == NULL)
		{
			return;
		}
		m_pCommand->ActiveConnection = m_pConnection;
		m_pCommand->CommandText = _bstr_t(szSPName);
		while (ParamList.size() > 0)
		{
			SPParameter spPara = ParamList.front();
			ParamList.pop_front();
			HRESULT hr = m_pParam.CreateInstance("ADODB.Parameter");
			if (hr != S_OK)
			{
				return;
			}
			m_pParam = m_pCommand->CreateParameter((_bstr_t)spPara.GetName(), spPara.GetType(), spPara.GetDirection(), spPara.GetSize(), (_variant_t)spPara.GetValue());
			if (m_pParam == NULL)
			{
				return;
			}
			m_pCommand->Parameters->Append(m_pParam);
			m_pParam->Release();
			m_pParam = NULL;
		}
		m_pRecordset = m_pCommand->Execute(NULL, NULL, adCmdStoredProc); //adAsyncExecute
		while (!m_pRecordset->adoEOF)
		{
			_variant_t value = m_pRecordset->GetCollect((_variant_t)szColumnName);
			if (value.vt != VT_NULL)
			{
				arrResult.Add(value);
			}
			m_pRecordset->MoveNext();
		}
		m_pRecordset->Close();
		m_pRecordset = NULL;
		//m_pCommand->Release();
		m_pCommand = NULL;
		return;
	}
	catch (_com_error e)
	{
		AfxMessageBox(e.Description());
		if (m_pParam != NULL)
		{
			m_pParam = NULL;
		}
		if (m_pRecordset != NULL)
		{
			m_pRecordset = NULL;
		}
		if (m_pCommand != NULL)
		{
			m_pCommand = NULL;
		}
		return;
	}
}

_RecordsetPtr DBManage::ExecuteSP(const CString& szSPName)
{
	if (m_pConnection == NULL)
	{
		m_pRecordset = NULL;
		return m_pRecordset;
	}
	if (m_pRecordset != NULL)
	{
		m_pRecordset->Close();
		m_pRecordset = NULL;
	}
	try
	{
		if (m_pConnection->GetState() == adStateClosed)
		{
			OpenConnection(m_szServer, m_szDatabase, m_szUser, m_szPassword);
		}
		HRESULT hr = m_pCommand.CreateInstance(__uuidof(Command));
		m_pRecordset.CreateInstance(__uuidof(Recordset));
		if (hr != S_OK || m_pCommand == NULL)
		{
			return NULL;
		}
		m_pCommand->ActiveConnection = m_pConnection;
		m_pCommand->CommandText = _bstr_t(szSPName);
		while (ParamList.size() > 0)
		{
			SPParameter spPara = ParamList.front();
			ParamList.pop_front();
			HRESULT hr = m_pParam.CreateInstance("ADODB.Parameter");
			if (hr != S_OK)
			{
				return NULL;
			}
			m_pParam = m_pCommand->CreateParameter((_bstr_t)spPara.GetName(), spPara.GetType(), spPara.GetDirection(), spPara.GetSize(), (_variant_t)spPara.GetValue());
			if (m_pParam == NULL)
			{
				return NULL;
			}
			m_pCommand->Parameters->Append(m_pParam);
			m_pParam->Release();
			m_pParam = NULL;
		}
		m_pRecordset = m_pCommand->Execute(NULL, NULL, adCmdStoredProc);
		m_pCommand = NULL;
	}
	catch (_com_error e)
	{
		AfxMessageBox(e.Description());
		if (m_pParam != NULL)
		{
			m_pParam = NULL;
		}
		if (m_pRecordset != NULL)
		{
			m_pRecordset = NULL;
		}
		if (m_pCommand != NULL)
		{
			m_pCommand = NULL;
		}
	}
	return m_pRecordset;
}

BOOL DBManage::GetRecordsByField(vector<vector<CString>>& table)
{
	try
	{
		if (m_pRecordset == NULL)
		{
			return FALSE;
		}
		Fields* fields = NULL;
		m_pRecordset->get_Fields(&fields);
		long ColCount;
		if (fields == NULL)
		{
			return false;
		}
		fields->get_Count(&ColCount);
		while (!m_pRecordset->adoEOF)
		{
			vector<CString> line;
			for (long i = 0; i < ColCount; i++)
			{
				Field* field = NULL;
				field = fields->GetItem((_variant_t)(i));
				if (field != NULL)
				{
					_variant_t value = m_pRecordset->GetCollect(field->GetName());
					if (value.vt != VT_NULL)
					{
						line.push_back(value);
					}
				}
			}
			table.push_back(line);
			m_pRecordset->MoveNext();
		}
		if (fields != NULL)
		{
			fields->Release();
		}
		return TRUE;
	}
	catch (_com_error e)
	{
		return FALSE;
	}
}

BOOL DBManage::ExecuteSPNoneQuery(const CString& szSPName)
{
	if (m_pConnection == NULL)
	{
		return FALSE;
	}
	try
	{
		if (m_pConnection->GetState() == adStateClosed)
		{
			OpenConnection(m_szServer, m_szDatabase, m_szUser, m_szPassword);
		}
		HRESULT hr = m_pCommand.CreateInstance(__uuidof(Command));
		if (hr != S_OK)
		{
			return NULL;
		}
		m_pCommand->ActiveConnection = m_pConnection;
		m_pCommand->CommandText = _bstr_t(szSPName);
		while (ParamList.size() > 0)
		{
			SPParameter spPara = ParamList.front();
			ParamList.pop_front();
			HRESULT hr = m_pParam.CreateInstance("ADODB.Parameter");
			if (hr != S_OK)
			{
				return FALSE;
			}
			m_pParam = m_pCommand->CreateParameter((_bstr_t)spPara.GetName(), spPara.GetType(), spPara.GetDirection(), spPara.GetSize(), (_variant_t)spPara.GetValue());
			if (m_pParam == NULL)
			{
				return FALSE;
			}
			m_pCommand->Parameters->Append(m_pParam);
			m_pParam->Release();
			m_pParam = NULL;
		}
		m_pCommand->Execute(NULL, NULL, adCmdStoredProc);
		m_pCommand.Release();
		m_pCommand = NULL;
		return TRUE;
	}
	catch (_com_error e)
	{
		AfxMessageBox(e.Description());
		if (m_pParam != NULL)
		{
			m_pParam = NULL;
		}
		if (m_pCommand != NULL)
		{
			m_pCommand = NULL;
		}
		return FALSE;
	}
}

