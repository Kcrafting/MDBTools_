#pragma once
/*
#define _AFXDLL
#include <afxwin.h>
#include <comdef.h>
#include <tchar.h>
#include <atlconv.h>
#include <exception>
*/

#define _AFXDLL 
//#include <afxwin.h>
#include <afxdisp.h>//COleCurrency class
#include <vector>//vector
#include <comutil.h>//_bstr_t
#include <typeinfo>
#include <iostream>
#include <iomanip>//er jin zhi 
#include <comdef.h>
//ע�Ȿ��ֻ�ܱ�new��������Ϊ���ڵ�����������ʹ����delete����˲������ⲿdelete

#import "C:\\Program Files (x86)\\Common Files\\system\\ado\\msado15.dll" no_namespace rename("EOF", "EndOfFile")  
#import "C:\\Program Files (x86)\\Common Files\\system\\ado\\msadox.dll" 

class MDBCombine
{
private:
	_bstr_t m_FilePathName;//�ļ���
	_ConnectionPtr m_pConnection;//����ָ��
	_RecordsetPtr m_pFileSchema;//�ļ���¼��
	_RecordsetPtr m_pTableSchema;//����¼��
	_RecordsetPtr m_pTableConstraint;//��ȡ����Լ��
	_bstr_t m_ConnectionString;//�����ַ���
	static int m_ClassCount;
public:
	MDBCombine(_bstr_t FilePathName);
	MDBCombine();
	~MDBCombine();
	_bstr_t getFileName();//��ȡ���ļ���
	_ConnectionPtr get_ConnectPtr();//��ȡ���ļ����ݿ������
	_RecordsetPtr get_FileSchema();//��ȡ�ļ��ṹ|����|
	_RecordsetPtr get_TableSchema();//��ȡ���ṹ
	_RecordsetPtr get_TableConstraint();//��ȡ����Լ��
	_bstr_t get_ConnectionString();//
	void SetConnectInfo(_bstr_t FilePathName);//�������ӵ��ļ���
	bool Connect();//���ӵ��ļ�
	bool DisConnect();//�Ͽ��ļ�����
	bool GetTableSchema();//��ȡ�ļ��ṹ
	bool GetFileSchema();//��ȡ���ṹ
	bool GetTableConstraint();//��ȡ���Լ��
	_bstr_t ConvertVariantV2(_variant_t var);//variantת������
};

