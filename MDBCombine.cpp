#include "MDBCombine.h"




MDBCombine::MDBCombine()
	:m_pConnection(nullptr), m_pTableSchema(nullptr), m_pFileSchema(nullptr), m_FilePathName(""),
	m_ConnectionString("")
{
	m_ClassCount++;
	OutputDebugString(L"\n>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>�����ڱ��޲������죡  " + _bstr_t(m_ClassCount) + _bstr_t(L"��  \n"));
}


MDBCombine::MDBCombine(_bstr_t FileName)
	:m_FilePathName(FileName), m_pConnection(nullptr), m_pTableSchema(nullptr), m_pFileSchema(nullptr),
	m_ConnectionString("")
{
	m_ClassCount++;
	OutputDebugString(L"\n>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>�����ں����ļ����Ĺ��죡  " + _bstr_t(m_ClassCount) + _bstr_t(L"��  \n"));
}

MDBCombine::~MDBCombine()
{
	try {
		m_ClassCount--;
		OutputDebugString(L"\n>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>������������  " + _bstr_t(m_ClassCount) + _bstr_t(L"��  \n"));
		if (m_pConnection)
			if (m_pConnection->State == adStateOpen)
				m_pConnection->Close();
		if (m_pTableSchema)
			if (m_pTableSchema->State == adStateOpen)
				m_pTableSchema->Close();
		if (m_pFileSchema)
			if (m_pFileSchema->State == adStateOpen)
				m_pFileSchema->Close();
		if (m_pTableConstraint)
			if (m_pTableConstraint->State == adStateOpen)
				m_pTableConstraint->Close();
		//delete this;
	}
	catch (std::exception &e)
	{
		const char * aaa = e.what();
	}
}

_bstr_t MDBCombine::getFileName()
{
	return m_FilePathName;
}

_ConnectionPtr MDBCombine::get_ConnectPtr()
{
	if (m_pConnection) 
		return m_pConnection;
	else
	{
		if (Connect())
		{
				return m_pConnection;
		}
	}
	return nullptr;
}

_RecordsetPtr MDBCombine::get_FileSchema()
{
	if (m_pFileSchema) return m_pFileSchema;
	if (!m_pConnection)
	{
		if (Connect())
		{
			if (GetFileSchema())
			{
				return m_pFileSchema;
			}
		}

	}
	else
	{
		if (GetFileSchema())
		{
			return m_pFileSchema;
		}
	}
	return nullptr;
}

_RecordsetPtr MDBCombine::get_TableSchema()
{
	if (m_pTableSchema) return m_pTableSchema;
	if (!m_pConnection) 
	{
		if (Connect())
		{
			if (GetTableSchema())
			{
				return m_pTableSchema;
			}
		}
		
	}
	else {
		if (GetTableSchema())
		{
			return m_pTableSchema;
		}
	}
	return nullptr;
}

_RecordsetPtr MDBCombine::get_TableConstraint()
{
	return m_pTableConstraint;
}

_bstr_t MDBCombine::get_ConnectionString()
{
	if (m_ConnectionString == _bstr_t(L""))
		Connect();
	return m_ConnectionString;
}

void MDBCombine::SetConnectInfo(_bstr_t FilePathName)
{
	m_FilePathName = FilePathName;
}

bool MDBCombine::Connect()
{
	//if (nullptr != m_pConnection) return;//����ע�������Ӻ�����ֶ��ͷţ�����ÿ��ִ�ж���ˢ������
	if (m_FilePathName == (_bstr_t)T2W(L""))
	{
		MessageBox(NULL, L"���ӻ�δ��ʼ����", L"��ʾ", NULL);
		return false;
	}
		
	if (FAILED(::CoInitialize(NULL)))
	{
		MessageBox(NULL, L"��ʼ��COM����ʧ�ܣ�", L"��ʾ", NULL);
		return false;
	}
	_bstr_t Provider("Provider = Microsoft.ACE.OLEDB.12.0; Data Source = ");
	_bstr_t Semicolon(";");
	m_ConnectionString = Provider + m_FilePathName + Semicolon;
	HRESULT hr = S_OK;
	try
	{
		hr = m_pConnection.CreateInstance(__uuidof(Connection));
	}
	catch (_com_error &err)
	{
		MessageBox(NULL,err.Description(), L"\n��������", MB_OK);
		return false;
	}
	if (hr != S_OK)
	{
		MessageBox(NULL, L"\n��������ʵ��ʱ��������\n (1) ϵͳ�ϰ�װ�Ų�ͬ��ADO�汾��\n(2) ����װ�Ĳ���64λ��office!\n(3) ���ļ������û�а�װ64λ��MDB����!\n(4) �����������µ�ַ����\nhttps://www.microsoft.com/zh-cn/download/details.aspx?id=13255", L"��������", MB_OK);
	}
	try {
		m_pConnection->Open(m_ConnectionString, "admin", "", adConnectUnspecified);
	}
	catch (_com_error &e) {
		MessageBox(NULL, (_bstr_t)T2W(L"���ļ��Ĺ����з�������\n")+e.Description(),L"\n���ӳ��� �� ԭ���п���: <1> ϵͳ�ϰ�װ�Ų�ͬ��ADO�汾��\n<2> ����װ�Ĳ���64λ��office!\n<3> ���ļ������û�а�װ64λ��MDB���棡\n<4> �����������µ�ַ����\nhttps://www.microsoft.com/zh-cn/download/details.aspx?id=13255",MB_OK);
		return false;
	}
	return true;
}

bool MDBCombine::DisConnect()
{
	try {
		if (m_pConnection)
		{
			if (m_pConnection->State == adStateOpen)
				m_pConnection->Close();
		}
		m_pConnection = nullptr;
		if (m_pTableSchema)
		{
			if (m_pTableSchema->State == adStateOpen)
				m_pTableSchema->Close();
		}
		m_pTableSchema = nullptr;
		if (m_pFileSchema)
		{
			if (m_pFileSchema->State == adStateOpen)
				m_pFileSchema->Close();
		}
		m_pFileSchema = nullptr;
		if (m_pTableConstraint)
		{
			if (m_pTableConstraint->State == adStateOpen)
				m_pTableConstraint->Close();
		}
		m_pTableConstraint = nullptr;
	}
	catch (std::exception &err)
	{
		MessageBox(NULL,T2W(LPTSTR(err.what())), L"��������", MB_OK);
		return false;
	}
	return true;
}

bool MDBCombine::GetTableSchema()
{
	if (!m_pConnection)
	{
		if (!Connect())
		{
			return false;
		}
	}
	try {
		m_pTableSchema = m_pConnection->OpenSchema(adSchemaColumns);//��������Ϣ
		//m_pTableSchema = m_pConnection->OpenSchema(adSchemaTableConstraints);
	}
	catch (_com_error &err)
	{
		MessageBox(NULL, (_bstr_t)T2W(L"��ȡ���ṹʱ��������\n") + err.Description(), L"��������", MB_OK);
		return false;
	}
	return true;
}

bool MDBCombine::GetFileSchema()
{
	if (!m_pConnection)
	{
		if (!Connect())
		{
			return false;
		}
	}
	try {
		m_pFileSchema = m_pConnection->OpenSchema(adSchemaTables);
	}
	catch (_com_error &err)
	{
		MessageBox(NULL, (_bstr_t)T2W(L"��ȡ�ļ��ṹʱ��������\n") + err.Description(), L"��������", MB_OK);
		return false;
	}
	return true;
}

bool MDBCombine::GetTableConstraint()
{
	if (!m_pConnection)
	{
		if (!Connect())
		{
			return false;
		}
	}
	try {
		m_pTableConstraint=m_pConnection->OpenSchema(adSchemaConstraintColumnUsage);
	}
	catch (_com_error &err)
	{
		MessageBox(NULL, (_bstr_t)T2W(L"��ȡ�ļ�Լ��ʱ��������\n") + err.Description(), L"��������", MB_OK);
		return false;
	}
	return true;
}

//ת��Variant�ڶ���
_bstr_t MDBCombine::ConvertVariantV2(_variant_t var)
{
	CString str;
	switch (var.vt)
	{
	case	VT_EMPTY:
	{
		break;
	}
	case	VT_NULL:
	{
		str = "NULL";
		//str = W2T(_bstr_t(var));
		break;
	}
	case	VT_I2:
	{
		str.Format(_T("%d"), (int)var.iVal);
		str.Format(_T("%d"), (int)var.iVal);
		break;
	}
	case 	VT_I4:
	{
		str.Format(_T("%d"), var.lVal);
		break;
	}
	case 	VT_R4:
	{
		str.Format(_T("%10.6f"), (double)var.fltVal);
		break;
	}
	case 	VT_R8:
	{
		str.Format(_T("%10.6f"), var.dblVal);
		break;
	}
	case 	VT_CY:
	{
		str = COleCurrency(var).Format();
		break;
	}
	case 	VT_DATE:
	{
		str = COleDateTime(var).Format();
		str = "'" + str + "'";
		//str = W2T(_bstr_t(var));
		break;
	}
	case 	VT_BSTR:
	{
		str = var.bstrVal;
		str = "'" + str + "'";
		break;
	}
	case 	VT_DISPATCH:
	{
		break;
	}
	case 	VT_ERROR:
	{
		break;
	}
	case 	VT_BOOL:
	{
		str = (var.boolVal == 0) ? "FALSE" : "TRUE";
		break;
	}
	case 	VT_VARIANT:
	{
		break;
	}
	case 	VT_UNKNOWN:
	{
		break;
	}
	case 	VT_DECIMAL:
	{
		//str.Format(_T("%15.6d"),var.dblVal);
		str = W2T(_bstr_t(var));
		break;
	}
	case 	VT_I1:
	{
		break;
	}
	case 	VT_UI1:
	{
		break;
	}
	case 	VT_UI2:
	{
		break;
	}
	case 	VT_UI4:
	{
		break;
	}
	case 	VT_I8:
	{
		break;
	}
	case 	VT_UI8:
	{
		break;
	}
	case 	VT_INT:
	{
		break;
	}
	case 	VT_UINT:
	{
		break;
	}
	case 	VT_VOID:
	{
		break;
	}
	case 	VT_HRESULT:
	{
		break;
	}
	case 	VT_PTR:
	{
		break;
	}
	case 	VT_SAFEARRAY:
	{
		break;
	}
	case 	VT_CARRAY:
	{
		break;
	}
	case 	VT_USERDEFINED:
	{
		break;
	}
	case 	VT_LPSTR:
	{
		break;
	}
	case 	VT_LPWSTR:
	{
		break;
	}
	case 	VT_RECORD:
	{
		break;
	}
	case 	VT_INT_PTR:
	{
		break;
	}
	case 	VT_UINT_PTR:
	{
		break;
	}
	case 	VT_FILETIME:
	{
		break;
	}
	case 	VT_BLOB:
	{
		break;
	}
	case 	VT_STREAM:
	{
		break;
	}
	case 	VT_STORAGE:
	{
		break;
	}
	case 	VT_STREAMED_OBJECT:
	{
		break;
	}
	case 	VT_STORED_OBJECT:
	{
		break;
	}
	case 	VT_BLOB_OBJECT:
	{
		break;
	}
	case 	VT_CF:
	{
		break;
	}
	case 	VT_CLSID:
	{
		break;
	}
	case 	VT_VERSIONED_STREAM:
	{
		break;
	}
	case 	VT_BSTR_BLOB:
	{
		break;
	}
	case 	VT_VECTOR:
	{
		break;
	}
	case 	VT_ARRAY:
	{
		break;
	}
	case 	VT_BYREF:
	{
		break;
	}
	case 	VT_RESERVED:
	{
		break;
	}
	case 	VT_ILLEGAL:
	{
		break;
	}
	//case 	VT_ILLEGALMASKED = 0xfff,
	//case 	VT_TYPEMASK = 0xfff
	default:
	{

	}
	}
	return str.AllocSysString();
}
