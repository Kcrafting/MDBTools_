#include "MDBClass.h"
#import "C:\\Program Files (x86)\\Common Files\\system\\ado\\msado15.dll" no_namespace rename("EOF", "EndOfFile")  
#import "C:\\Program Files (x86)\\Common Files\\system\\ado\\msadox.dll" 


MDBClass::MDBClass()
{
}


MDBClass::~MDBClass()
{
}

_bstr_t MDBClass::ConvertVariant(_variant_t var)
{
	CString str;

	//���´�����ʾ���ת��ΪC��׼�ַ�����  
	switch (var.vt)
	{
	case VT_BSTR:
	{
		str = var.bstrVal;
		str = "'" + str + "'";
		break;
	}
	case VT_I2: //var is short int type   
	{
		str.Format(_T("%d"), (int)var.iVal);
		break;
	}
	case VT_I4: //var is long int type  
	{
		str.Format(_T("%d"), var.lVal);
		break;
	}
	case VT_R4: //var is float type  
	{
		str.Format(_T("%10.6f"), (double)var.fltVal);
		break;
	}
	case VT_R8: //var is double type  
	{
		str.Format(_T("%10.6f"), var.dblVal);
		break;
	}
	case VT_CY: //var is CY type  
	{
		str = COleCurrency(var).Format();
		break;
	}
	case VT_DATE: //var is DATE type  
	{
		str = COleDateTime(var).Format();
		break;
	}
	case VT_BOOL: //var is VARIANT_BOOL  
	{
		str = (var.boolVal == 0) ? "FALSE" : "TRUE";
		break;
	}
	//case VARENUM::
	default:
	{
		str.Format(_T("Unk type %d\n"), var.vt);
		TRACE("Unknown type %d\n", var.vt);
		break;
	}
	}
	return str.AllocSysString();
}
//ת��Variant�ڶ���
_bstr_t MDBClass::ConvertVariantV2(_variant_t var)
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
//����1����ȡһ���ļ������Ʒ��أ�Ϊ�ϲ�Ŀ���ļ�
_bstr_t MDBClass::GetSingleFileName()
{
	_bstr_t Result="";
	OPENFILENAME ofn;
	TCHAR szOpenFileNames[80 * MAX_PATH];
	TCHAR szPath[MAX_PATH];
	TCHAR szFileName[80 * MAX_PATH];
	TCHAR* p;
	int nLen = 0;
	ZeroMemory(&ofn, sizeof(ofn));
	ofn.Flags = OFN_PATHMUSTEXIST | OFN_FILEMUSTEXIST | OFN_EXPLORER;
	ofn.lStructSize = sizeof(ofn);
	ofn.lpstrFile = szOpenFileNames;
	ofn.nMaxFile = sizeof(szOpenFileNames);
	ofn.lpstrFile[0] = '\0';
	ofn.lpstrFilter = TEXT("MDB���ݿ��ļ�\0*.MDB\0\0");
	ofn.lpstrTitle = LPCWSTR(L"ѡ��Ŀ���ļ�");
	if (GetOpenFileName(&ofn))
	{
		lstrcpyn(szPath, szOpenFileNames, ofn.nFileOffset);
		szPath[ofn.nFileOffset] = '/0';
		nLen = lstrlen(szPath);
		return (T2W(szOpenFileNames));
		//return Result;
	}
	return Result;
}
//����2����ȡ����ļ������Ʒ��أ����Ŀ���ļ��������У�����
vector<_bstr_t> MDBClass::GetMultiFileName()
{
	vector<_bstr_t> Result;
	OPENFILENAME ofn;
	TCHAR szOpenFileNames[80 * MAX_PATH];
	TCHAR szPath[MAX_PATH];
	TCHAR szFileName[80 * MAX_PATH];
	TCHAR* p;
	int nLen = 0;
	ZeroMemory(&ofn, sizeof(ofn));
	ofn.Flags = OFN_EXPLORER | OFN_ALLOWMULTISELECT;
	ofn.lStructSize = sizeof(ofn);
	ofn.lpstrFile = szOpenFileNames;
	ofn.nMaxFile = sizeof(szOpenFileNames);
	ofn.lpstrFile[0] = '\0';
	ofn.lpstrFilter = TEXT("MDB���ݿ��ļ�\0*.MDB\0\0");
	ofn.lpstrTitle = LPCWSTR(L"ѡ����Ҫ���кϲ����ļ��ɶ�ѡ");
	if (GetOpenFileName(&ofn))
	{
		lstrcpyn(szPath, szOpenFileNames, ofn.nFileOffset);
		szPath[ofn.nFileOffset] = '/0';
		nLen = lstrlen(szPath);
		if (szPath[nLen - 1] != '//')   //���ѡ�˶���ļ�,��������'//'
		{
			lstrcat(szPath, TEXT("\\"));
		}
		p = szOpenFileNames + ofn.nFileOffset; //��ָ���Ƶ���һ���ļ�
		ZeroMemory(szFileName, sizeof(szFileName));
		while (*p)
		{
			lstrcat(szFileName, szPath);  //���ļ�������·��  
			lstrcat(szFileName, p);    //�����ļ���
			Result.push_back(T2W(szFileName));
			//lstrcat(szFileName, TEXT("\n\/\/")); //����   
			p += lstrlen(p) + 1;     //������һ���ļ�
			ZeroMemory(szFileName, sizeof(szFileName));
		}
		return Result;
	}
	return Result;
}
//����3����ȡMDB��ṹ���أ�OUT�����ƣ����ͣ�IN���ļ�·������
bool MDBClass::GetTableList(_bstr_t FilePath, vector<_bstr_t>& TABLE_NAME, vector<_bstr_t>& TABLE_TYPE)
{
	if (FAILED(::CoInitialize(NULL)))
		return false;
	//��ʼ��COM
	_ConnectionPtr pConnection = NULL;
	_RecordsetPtr pTableSchema = NULL;
	_bstr_t Pro("Provider = Microsoft.ACE.OLEDB.12.0; Data Source = ");
	_bstr_t FileNameEnd(";");
	_bstr_t strCnn = Pro + FilePath + FileNameEnd;
	try {
		HRESULT hr = pConnection.CreateInstance(__uuidof(Connection));
		if (hr != S_OK)
		{
			AfxMessageBox((_bstr_t)T2W(_T("��������ʵ���������ADO�汾���ܺͱ���汾��һ�£�")));
			return false;
		}
	}
	catch (_com_error &e)
	{
		AfxMessageBox((_bstr_t)T2W(_T("��������ʵ���������ADO�汾���ܺͱ���汾��һ�£�")) + e.Description());
		return false;
	}
	try {
		pConnection->Open(strCnn, "admin", "", -1);
	}
	catch (_com_error &e) {
		AfxMessageBox((_bstr_t)T2W(_T("�����ӳ�����ļ�����Ͽ���û�а�װX64�汾�����氲װ\nhttps://www.microsoft.com/zh-cn/download/details.aspx?id=13255\n")) + e.Description());
		return false;
	}
	try {
		pTableSchema = pConnection->OpenSchema(adSchemaTables);
		//AfxMessageBox((_bstr_t)T2W(_T("����ǰ��AccessDatabaseEngine�뵱ǰҪ��Ĳ�����\n�뵽����ĵ�ַ����X64�汾�����氲װ\nhttps://www.microsoft.com/zh-cn/download/details.aspx?id=13255")));
	}
	catch (_com_error &e)
	{
		cout << GetLastError();
		AfxMessageBox((_bstr_t)T2W(_T("��ѯ�ļ��ṹ����")) + e.Description());
		CString MSG = W2T(e.Description());
		if (MSG.Find(L"��ȷ��װ", 0))
		{
			AfxMessageBox((_bstr_t)T2W(_T("����ǰ��AccessDatabaseEngine�뵱ǰҪ��Ĳ�����\n�뵽����ĵ�ַ����X64�汾�����氲װ\nhttps://www.microsoft.com/zh-cn/download/details.aspx?id=13255")));
		}
		return false;
	}
	while (!(pTableSchema->EndOfFile))
	{
		if ((CString(W2T((_bstr_t)(pTableSchema->Fields->GetItem("TABLE_NAME")->Value)))).Left(2) != "MS")
		{
			/*
			QMessageBox msgBox;
			msgBox.setText(QString::number(pTableSchema->Fields->Count, 10));
			msgBox.exec();
			int i = 0;
			
			while (i < pTableSchema->Fields->Count)
			{ 
				OutputDebugString(LPCWSTR(pTableSchema->Fields->Item[long(i)]->Name)+ (_bstr_t)T2W(_T("\n")));

				_variant_t dd = pTableSchema->Fields->Item[long(i)]->Name;
				i++;
			}
			*/
			//msgBox.setText(pTableSchema->Fields->Item[0]->Name);
			//msgBox.exec();
			/*******************************************************
			TABLE_CATALOG
			TABLE_SCHEMA
			TABLE_NAME
			TABLE_TYPE
			TABLE_GUID
			DESCRIPTION
			TABLE_PROPID
			DATE_CREATED
			DATE_MODIFIED
			********************************************************/
			try {
				_variant_t ddd = pTableSchema->Fields->GetItem("TABLE_NAME")->Value;
			}
			catch (std::exception &e) {
				AfxMessageBox(L"��ȡ���г���");
				MessageBox(NULL,L"",LPCWSTR(e.what()),NULL);
				//return false;
			}
		}
		pTableSchema->MoveNext();
	}
	if (pTableSchema)
		if (pTableSchema->State == adStateOpen)
			pTableSchema->Close();
	if (pConnection)
		if (pConnection->State == adStateOpen)
			pConnection->Close();
	CoUninitialize();
	return true;
}
//����4����ȡ����ֶ����ͷ���
bool MDBClass::GetColumnAndType(_bstr_t FilePath, _bstr_t TableName, vector<_bstr_t>& COLUMN_NAME, vector<DataTypeEnum>& COLUMN_TYPE)
{
	if (FAILED(::CoInitialize(NULL)))
		return false;
	_ConnectionPtr pConnection = NULL;
	_RecordsetPtr pColumnSchema = NULL;
	_bstr_t Pro("Provider = Microsoft.ACE.OLEDB.12.0; Data Source = ");
	_bstr_t FileNameEnd(";");
	_bstr_t strCnn = Pro + FilePath + FileNameEnd;
	try {
		IfCreateSuccess(pConnection.CreateInstance(__uuidof(Connection)));
		pConnection->Open(strCnn, "admin", "", -1);
		pColumnSchema = pConnection->OpenSchema(adSchemaColumns);
	}
	catch (_com_error &e)
	{
		AfxMessageBox((_bstr_t)T2W(_T("�������ļ������ӳ���")) + e.Description());
		return false;
	}
	_variant_t ExecResult;
	_RecordsetPtr ReturnColumn = NULL;
	while (!(pColumnSchema->EndOfFile))
	{
		if ((_bstr_t)(pColumnSchema->Fields->GetItem("TABLE_NAME")->Value) == TableName)
		{
			_variant_t Column_Name = pColumnSchema->Fields->GetItem("COLUMN_NAME")->Value;
			COLUMN_NAME.push_back(Column_Name);
			//��ȡ������������
			_bstr_t SQLTxt = (_bstr_t)W2T(_T("Select ")) + (_bstr_t)Column_Name + (_bstr_t)W2T(_T(" From ")) + TableName;
			//_RecordsetPtr ReturnColumn=NULL;

			try {
				//ReturnColumn = pConnection->Execute(SQLTxt, &ExecResult, adCmdText);
				ReturnColumn = pConnection->Execute(SQLTxt, &ExecResult, adCmdText);
			}
			catch (_com_error &e) {
				AfxMessageBox((_bstr_t)T2W(_T("ִ�д���")) + e.Description());
			}
			COLUMN_TYPE.push_back(ReturnColumn->Fields->GetItem(Column_Name)->Type);
		}
		pColumnSchema->MoveNext();
	}

	if (ReturnColumn)
		if (ReturnColumn->State == adStateOpen)
			ReturnColumn->Close();
	if (pColumnSchema)
		if (pColumnSchema->State == adStateOpen)
			pColumnSchema->Close();
	if (pConnection)
		if (pConnection->State == adStateOpen)
			pConnection->Close();
	CoUninitialize();
	return true;
}
//����5���ļ��ϲ���Ŀ���ļ���·���б���ȥ���Ѻϲ��ļ�����¼������Ϣ
vector<_bstr_t> MDBClass::CombineMDB(_bstr_t TarFileName, vector<_bstr_t>OrgFiles)
{
	vector<_bstr_t> a;
	return a;
}
//����6: ���������Ϣ
void MDBClass::ShowMessage()
{
	/*
	for (size_t start = 0; start < ErrorMessage.size(); start++)
	{
	cout << ErrorMessage[(long)start];
	}*/
}
//����7��ԭ�ļ��ϲ���Ŀ���ļ�
bool MDBClass::CombineFileToFile(_bstr_t TarFile, _bstr_t OrgFile)
{
	cout << "��ʼ�ϲ���\n";
	vector<_bstr_t> TarTableName;
	vector<_bstr_t> TarTableType;
	vector<_bstr_t> OrgTableName;
	vector<_bstr_t> OrgTableType;
	if (GetTableList(TarFile, TarTableName, TarTableType))
	{
		if (GetTableList(OrgFile, OrgTableName, OrgTableType))
		{
			for (size_t i = 0; i < TarTableType.size(); i++)
			{
				cout << "�����" << i + 1 << "/" << TarTableType.size() << "\n";
				if (i == (size_t)12)
				{
					cout << "stophere";
				}
				if (TarTableType[i] == (_bstr_t)T2W(_T("TABLE")))
				{
					for (size_t j = 0; j < OrgTableType.size(); j++)
					{
						if (OrgTableType[j] == (_bstr_t)T2W(_T("TABLE")))
						{
							if (TarTableName[i] == OrgTableName[j])//same table
							{
								if (TarTableName[i] == (_bstr_t)T2W(_T("FBF")))
								{
									cout << "";
								}
								vector<_bstr_t> TarColumnName;
								vector<DataTypeEnum> TarColumnType;
								vector<_bstr_t> OrgColumnName;
								vector<DataTypeEnum> OrgColumnType;
								if (GetColumnAndType(TarFile, TarTableName[i], TarColumnName, TarColumnType))
								{
									if (GetColumnAndType(TarFile, OrgTableName[j], OrgColumnName, OrgColumnType))
									{
										vector<int> TempOut;
										vector<int> TempIn;
										for (size_t ii = 0; ii < TarColumnName.size(); ii++)
										{
											for (size_t jj = 0; jj < OrgColumnName.size(); jj++)
											{

												if (TarColumnName[ii] == OrgColumnName[jj] && TarColumnType[ii] == OrgColumnType[jj])
												{
													TempOut.push_back(jj);
													TempIn.push_back(ii);
												}
											}

										}
										if (TempOut.size() != size_t(0))
										{
											_variant_t TempResult;
											if (FAILED(::CoInitialize(NULL)))
												return false;
											_RecordsetPtr OrgData = NULL;
											_ConnectionPtr OrgCon = NULL;
											_bstr_t Provider("Provider = Microsoft.ACE.OLEDB.12.0; Data Source = ");
											_bstr_t ProviderEnd(";");
											_bstr_t OrgConStr = Provider + OrgFile + ProviderEnd;
											IfCreateSuccess(OrgCon.CreateInstance(__uuidof(Connection)));
											OrgCon->Open(OrgConStr, "admin", "", adConnectUnspecified);
											_bstr_t SQLTxt10 = (_bstr_t)W2T(_T("Select * ")) + (_bstr_t)W2T(_T(" from ")) + OrgTableName[j];
											OrgData = OrgCon->Execute(SQLTxt10, &TempResult, adCmdText);

											_RecordsetPtr TarData = NULL;
											_ConnectionPtr TarCon = NULL;
											_bstr_t TarConStr = Provider + TarFile + ProviderEnd;
											IfCreateSuccess(TarCon.CreateInstance(__uuidof(Connection)));
											TarCon->Open(TarConStr, "admin", "", adConnectUnspecified);
											_bstr_t Tables;
											for (size_t iii = 0; iii < TempOut.size(); iii++)
											{
												if ((iii + 1) == TempOut.size())
												{
													Tables += OrgColumnName[TempOut[iii]];
												}
												else
												{
													Tables += OrgColumnName[TempOut[iii]] + (_bstr_t)W2T(_T(","));
												}
											}
											while (!(OrgData->EndOfFile))
											{
												_bstr_t Values;
												//CString Values;
												try {
													for (size_t iii = 0; iii < TempOut.size(); iii++)
													{
														if ((iii + 1) == TempOut.size())
														{
															//if (ConvertVariantV2(OrgData->Fields->GetItem(OrgColumnName[TempOut[iii]])->Value).GetBSTR() == NULL)
															if (OrgData->Fields->GetItem(OrgColumnName[TempOut[iii]])->Value.vt == VT_NULL || ConvertVariantV2(OrgData->Fields->GetItem(OrgColumnName[TempOut[iii]])->Value) == (_bstr_t)W2T(_T("")))
																Values += (_bstr_t)W2T(_T("FUCK"));
															else
																Values += ConvertVariantV2(OrgData->Fields->GetItem(OrgColumnName[TempOut[iii]])->Value);
														}
														else
														{
															try {
																//if (ConvertVariantV2(OrgData->Fields->GetItem(OrgColumnName[TempOut[iii]])->Value).GetBSTR() == NULL)
																if (OrgData->Fields->GetItem(OrgColumnName[TempOut[iii]])->Value.vt == VT_NULL || ConvertVariantV2(OrgData->Fields->GetItem(OrgColumnName[TempOut[iii]])->Value) == (_bstr_t)W2T(_T("")))
																	Values += (_bstr_t)W2T(_T("'FUCK',"));
																else
																{
																	if (ConvertVariantV2(OrgData->Fields->GetItem(OrgColumnName[TempOut[iii]])->Value) == (_bstr_t)W2T(_T(" ")))
																		Values += (_bstr_t)W2T(_T("'FUCK',"));
																	Values += ConvertVariantV2(OrgData->Fields->GetItem(OrgColumnName[TempOut[iii]])->Value) + (_bstr_t)W2T(_T(","));
																}
																//Values += ConvertVariantV2(OrgData->Fields->GetItem(OrgColumnName[TempOut[iii]])->Value) + (_bstr_t)W2T(_T(","));
																//type_info aaae=((_bstr_t)OrgData->Fields->GetItem(OrgColumnName[TempOut[iii]])->Value);	
															}
															catch (_com_error &EE) {
																//long err = GetLastError();
																AfxMessageBox((_bstr_t)W2T(_T("��ת����")) + EE.Description());
															}
														}
													}
												}
												catch (_com_error &e) {
													//ErrorMessage.push_back("����ת��������" + Values);
													cout << "����ת��������" + Tables + Values + "\n" << e.Description() << "\n";
													//AfxMessageBox("����ת��������"+e.Description());
												}
												_bstr_t SQLTxt11 = (_bstr_t)W2T(_T("INSERT INTO ")) + TarTableName[i] + (_bstr_t)W2T(_T(" ( ")) + Tables + (_bstr_t)W2T(_T(") VALUES( ")) + Values + (_bstr_t)W2T(_T(")"));
												//AfxMessageBox(SQLTxt11);
												try {
													TarCon->Execute(SQLTxt11, &TempResult, adCmdText);
												}
												catch (_com_error &e) {
													//ErrorMessage.push_back("���ݲ��������" + SQLTxt11 + e.Description());

													cout << "���ݲ��������" + SQLTxt11 + e.Description() + "\n";
													//AfxMessageBox(SQLTxt11);
													AfxMessageBox("���ݲ��������" + e.Description());
												}
												OrgData->MoveNext();
											}
											if (OrgData)
												if (OrgData->State == adStateOpen)
													OrgData->Close();
											if (OrgCon)
												if (OrgCon->State == adStateOpen)
													OrgCon->Close();
											if (TarData)
												if (TarData->State == adStateOpen)
													TarData->Close();
											if (TarCon)
												if (TarCon->State == adStateOpen)
													TarCon->Close();
										}

										TempIn.clear();
										TempOut.clear();
									}
								}
							}
						}
					}
				}
				else break;
			}
		}
	}
	return true;
}
bool MDBClass::CombineFileToFileEx(_bstr_t TarFile, _bstr_t OrgFile, vector<_variant_t> Err)
{
	cout << "��ʼ�ϲ���\n";
	vector<_bstr_t> TarTableName;
	vector<_bstr_t> TarTableType;
	vector<_bstr_t> OrgTableName;
	vector<_bstr_t> OrgTableType;
	if (GetTableList(TarFile, TarTableName, TarTableType))
	{
		if (GetTableList(OrgFile, OrgTableName, OrgTableType))
		{
			for (size_t i = 0; i < TarTableType.size(); i++)
			{
				//cout << "�����" << i + 1 << "/" << TarTableType.size() << "\n";
				if (i == (size_t)12)
				{
					//cout << "stophere";
				}
				if (TarTableType[i] == (_bstr_t)T2W(_T("TABLE")))
				{
					for (size_t j = 0; j < OrgTableType.size(); j++)
					{
						if (OrgTableType[j] == (_bstr_t)T2W(_T("TABLE")))
						{
							if (TarTableName[i] == OrgTableName[j])//same table
							{
								if (TarTableName[i] == (_bstr_t)T2W(_T("FBF")))
								{
									//cout << "";
								}
								vector<_bstr_t> TarColumnName;
								vector<DataTypeEnum> TarColumnType;
								vector<_bstr_t> OrgColumnName;
								vector<DataTypeEnum> OrgColumnType;
								if (GetColumnAndType(TarFile, TarTableName[i], TarColumnName, TarColumnType))
								{
									if (GetColumnAndType(TarFile, OrgTableName[j], OrgColumnName, OrgColumnType))
									{
										vector<int> TempOut;
										vector<int> TempIn;
										for (size_t ii = 0; ii < TarColumnName.size(); ii++)
										{
											for (size_t jj = 0; jj < OrgColumnName.size(); jj++)
											{

												if (TarColumnName[ii] == OrgColumnName[jj] && TarColumnType[ii] == OrgColumnType[jj])
												{
													TempOut.push_back(jj);
													TempIn.push_back(ii);
												}
											}

										}
										if (TempOut.size() != size_t(0))
										{
											_variant_t TempResult;
											if (FAILED(::CoInitialize(NULL)))
												return false;
											_RecordsetPtr OrgData = NULL;
											_ConnectionPtr OrgCon = NULL;
											_bstr_t Provider("Provider = Microsoft.ACE.OLEDB.12.0; Data Source = ");
											_bstr_t ProviderEnd(";");
											_bstr_t OrgConStr = Provider + OrgFile + ProviderEnd;
											IfCreateSuccess(OrgCon.CreateInstance(__uuidof(Connection)));
											OrgCon->Open(OrgConStr, "admin", "", adConnectUnspecified);
											_bstr_t SQLTxt10 = (_bstr_t)W2T(_T("Select * ")) + (_bstr_t)W2T(_T(" from ")) + OrgTableName[j];
											OrgData = OrgCon->Execute(SQLTxt10, &TempResult, adCmdText);

											_RecordsetPtr TarData = NULL;
											_ConnectionPtr TarCon = NULL;
											_bstr_t TarConStr = Provider + TarFile + ProviderEnd;
											IfCreateSuccess(TarCon.CreateInstance(__uuidof(Connection)));
											TarCon->Open(TarConStr, "admin", "", adConnectUnspecified);
											_bstr_t Tables;
											for (size_t iii = 0; iii < TempOut.size(); iii++)
											{
												if ((iii + 1) == TempOut.size())
												{
													Tables += OrgColumnName[TempOut[iii]];
												}
												else
												{
													Tables += OrgColumnName[TempOut[iii]] + (_bstr_t)W2T(_T(","));
												}
											}
											while (!(OrgData->EndOfFile))
											{
												_bstr_t Values;
												//CString Values;
												try {
													for (size_t iii = 0; iii < TempOut.size(); iii++)
													{
														if ((iii + 1) == TempOut.size())
														{
															if (ConvertVariantV2(OrgData->Fields->GetItem(OrgColumnName[TempOut[iii]])->Value).GetBSTR() == NULL)
																Values += BSTR("FUCK");
															else
																Values += ConvertVariantV2(OrgData->Fields->GetItem(OrgColumnName[TempOut[iii]])->Value);
														}
														else
														{
															try {
																if (ConvertVariantV2(OrgData->Fields->GetItem(OrgColumnName[TempOut[iii]])->Value).GetBSTR() == NULL)
																	Values += BSTR("FUCK,");
																else
																	Values += ConvertVariantV2(OrgData->Fields->GetItem(OrgColumnName[TempOut[iii]])->Value) + (_bstr_t)W2T(_T(","));
																//type_info aaae=((_bstr_t)OrgData->Fields->GetItem(OrgColumnName[TempOut[iii]])->Value);	
															}
															catch (_com_error &EE) {
																//long err = GetLastError();
																AfxMessageBox((_bstr_t)W2T(_T("�����ݽ��е�һ������ת��ʱ�������󣡣�")) + EE.Description());
															}
														}
													}
												}
												catch (_com_error &e) {
													//ErrorMessage.push_back("����ת��������" + Values);
													//cout << "����ת��������" + Tables + Values + "\n" << e.Description() << "\n";
													//AfxMessageBox("����ת��������"+e.Description());
												}
												_bstr_t SQLTxt11 = (_bstr_t)W2T(_T("INSERT INTO ")) + TarTableName[i] + (_bstr_t)W2T(_T(" ( ")) + Tables + (_bstr_t)W2T(_T(") VALUES( ")) + Values + (_bstr_t)W2T(_T(")"));
												//AfxMessageBox(SQLTxt11);
												try {
													TarCon->Execute(SQLTxt11, &TempResult, adCmdText);
												}
												catch (_com_error &e) {
													//ErrorMessage.push_back("���ݲ��������" + SQLTxt11 + e.Description());

													cout << "���ݲ��������" + SQLTxt11 + e.Description() + "\n";
													//AfxMessageBox(SQLTxt11);
													AfxMessageBox("���ݲ��������" + e.Description());
												}
												OrgData->MoveNext();
											}
											if (OrgData)
												if (OrgData->State == adStateOpen)
													OrgData->Close();
											if (OrgCon)
												if (OrgCon->State == adStateOpen)
													OrgCon->Close();
											if (TarData)
												if (TarData->State == adStateOpen)
													TarData->Close();
											if (TarCon)
												if (TarCon->State == adStateOpen)
													TarCon->Close();
										}

										TempIn.clear();
										TempOut.clear();
									}
								}
							}
						}
					}
				}
				else break;
			}
		}
	}
	return true;
}
//��ѡ����ļ���ʾ����--��������������ʵ��
void MDBClass::ShowSourceFilesWeSelected(_bstr_t fileName)
{
	
}

_ConnectionPtr MDBClass::SetConnectionToFile(_bstr_t FilePath)
{
	if (FAILED(::CoInitialize(NULL)))
		return nullptr;//ʵ����COM�����ʧ�ܷ��ؿ�ָ��
	_ConnectionPtr pConnection = NULL;//���ڷ��ص�����ָ��
	_bstr_t Pro("Provider = Microsoft.ACE.OLEDB.12.0; Data Source = ");
	_bstr_t FileNameEnd(";");
	_bstr_t strCnn = Pro + FilePath + FileNameEnd;
	try {
		HRESULT hr = pConnection.CreateInstance(__uuidof(Connection));
		if (hr != S_OK)
		{
			AfxMessageBox((_bstr_t)T2W(_T("û�в�ѯ����Ч�Ľӿڣ�")));
			return nullptr;
		}
	}
	catch (_com_error &e)
	{
		AfxMessageBox((_bstr_t)T2W(_T("ʵ�������ӳ�����ļ������ȱ����ӦADO�����/n����ADO�汾�ͱ���汾��һ�£�/n")) + e.Description());
		return nullptr;
	}
	return pConnection;
}

void MDBClass::GetMDBSchema(_ConnectionPtr pConnection, vector<_bstr_t>& TABLE_NAME, vector<_bstr_t>& TABLE_TYPE)
{
	_RecordsetPtr pTableSchema = NULL;
	try {
		pTableSchema = pConnection->OpenSchema(adSchemaTables);
		//AfxMessageBox((_bstr_t)T2W(_T("����ǰ��AccessDatabaseEngine�뵱ǰҪ��Ĳ�����\n�뵽����ĵ�ַ����X64�汾�����氲װ\nhttps://www.microsoft.com/zh-cn/download/details.aspx?id=13255")));
	}
	catch (_com_error &e)
	{
		AfxMessageBox((_bstr_t)T2W(_T("��ѯ�ļ��ṹ����")) + e.Description());
		CString MSG = W2T(e.Description());
		if (MSG.Find(L"��ȷ��װ", 0))
		{
			AfxMessageBox((_bstr_t)T2W(_T("����ǰ��AccessDatabaseEngine�뵱ǰҪ��Ĳ�����\n�뵽����ĵ�ַ����X64�汾�����氲װ\nhttps://www.microsoft.com/zh-cn/download/details.aspx?id=13255")));
		}
		return;
	}
	while (!(pTableSchema->EndOfFile))
	{
		if ((CString(W2T((_bstr_t)(pTableSchema->Fields->GetItem("TABLE_NAME")->Value)))).Left(2) != "MS")
		{
			try {
				TABLE_NAME.push_back(pTableSchema->Fields->GetItem("TABLE_NAME")->Value);
				TABLE_TYPE.push_back(pTableSchema->Fields->GetItem("TABLE_TYPE")->Value);
			}
			catch (...) {
				AfxMessageBox(L"��ȡ���г���");
				//return false;
			}
		}
		pTableSchema->MoveNext();
	}
	if (pTableSchema)
		if (pTableSchema->State == adStateOpen)
			pTableSchema->Close();
	//if (pConnection)
		//if (pConnection->State == adStateOpen)
			//pConnection->Close();
	//���������������ͷ�
	CoUninitialize();
}
