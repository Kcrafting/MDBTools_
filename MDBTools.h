#pragma once
#define _AFXDLL 
#include <afxwin.h>
#include <afxdisp.h>//COleCurrency class
#include <QtWidgets/QMainWindow>
#include "ui_MDBTools.h"
#include <comdef.h>
#include <vector>
#include "MDBCombine.h"
using std::vector;
class MDBTools : public QMainWindow
{
	Q_OBJECT

public:
	MDBTools(QWidget *parent = Q_NULLPTR);
	bool CombineFile(_RecordsetPtr Org, MDBCombine* MDBobj);
	void SetGlobleProGressValue(int Value);
private:
	QTableWidget * m_TarSchema;
	vector<QProgressBar*> m_ControlsDocker;//����������
	vector<QTableWidget*> m_TabDocker;//Tab������
	int m_fileCount;//�ļ�����
private:
	void ShowSourceFilesWeSelected(_bstr_t fileName,_bstr_t Percent);//��ʾ����ѡ�����Ҫ���кϲ����ļ�
	Ui::MDBToolsClass ui;
public slots://����Ϊ�Զ���Ĳۺ���
void slots_OnTargetFileSelect();//����ѡ��Ŀ���ļ��¼�
void slots_OnSourceFileSelect();//����ѡ��Դ�ļ�
void slots_OnStartCombineClick();//�����ϲ���ť
void slots_OnHelpClick();//��������
};
