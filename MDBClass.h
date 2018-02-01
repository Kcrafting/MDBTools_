#pragma once
#define _AFXDLL 
//#include <afxwin.h>
#include <afxwin.h>
#include <afxdisp.h>//COleCurrency class
#include <vector>//vector
#include <comutil.h>//_bstr_t
#include <typeinfo>
#include <iostream>
//#include <afxdisp.h>//COleCurrency class
#include <iomanip>//er jin zhi 
#include <comdef.h>
#include <qmessagebox.h>
#include <Windows.h>
//�����
#import "C:\\Program Files (x86)\\Common Files\\system\\ado\\msado15.dll" no_namespace rename("EOF", "EndOfFile")  
#import "C:\\Program Files (x86)\\Common Files\\system\\ado\\msadox.dll" 

#define IfCreateSuccess(x) if FAILED(x) _com_issue_error(x);//COMʵ�������Ƿ�ɹ�
using std::vector;
using std::cout;
class MDBClass
{
public:
public:
	MDBClass();
	~MDBClass();
	_bstr_t GetSingleFileName();
	vector<_bstr_t> GetMultiFileName();
	bool GetTableList(_bstr_t FilePath, vector<_bstr_t>& TABLE_NAME, vector<_bstr_t>& TABLE_TYPE);
	bool GetColumnAndType(_bstr_t FilePath, _bstr_t TableName, vector<_bstr_t>& COLUMN_NAME, vector<DataTypeEnum>& COLUMN_TYPE);
	vector<_bstr_t> CombineMDB(_bstr_t TarFileName, vector<_bstr_t>OrgFiles);
	bool CombineFileToFile(_bstr_t TarFile, _bstr_t OrgFile);
	_bstr_t ConvertVariant(_variant_t var);
	_bstr_t ConvertVariantV2(_variant_t var);
	void ShowMessage();
	bool CombineFileToFileEx(_bstr_t TarFile, _bstr_t OrgFile, vector<_variant_t> Err);
	void ShowSourceFilesWeSelected(_bstr_t fileName);//��ʾ����ѡ�����Ҫ���кϲ����ļ�
	_ConnectionPtr SetConnectionToFile(_bstr_t FilePath);//�������ӵ�ʵ��ָ��
	void GetMDBSchema(_ConnectionPtr pConnection,vector<_bstr_t>& TABLE_NAME, vector<_bstr_t>& TABLE_TYPE);//����ָ��ָ��ı�ṹ
};

