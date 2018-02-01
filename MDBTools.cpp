#include <qpushbutton.h>
#include <qtablewidget.h>
#include <qtoolbar.h>
#include <qmessagebox.h>
#include <algorithm> //for_each����
#include <iterator>//������
#include <functional>//mem_fun_ref ����
#include <qprogressbar.h>
#include "MDBTools.h"
#include <memory>
#include "MDBClass.h"
#define ProGressMax 10000
#define ProGressMin 0

using std::shared_ptr;
struct MapWord
{
	_bstr_t ColumnName;
	_bstr_t DefaultValue;
	int LengthRequire;//�жϳ���Ҫ��
};
MDBClass *MDB;
MDBCombine *MDBCM;//Ŀ���ļ�
_bstr_t g_TarFileName;
vector<bstr_t> g_SourceFileNames;
using std::for_each;
long EveryProGressMax = 0;//ÿ�������������ֵ
long ProGressIndex = -1;//ÿ���������±�
long GetMDBFileMaxRow(_RecordsetPtr pRs, _ConnectionPtr pCon, _RecordsetPtr pConRECORD);
void RecordSetDisplayInTableWidget(_RecordsetPtr pRSP, QTableWidget* QWD);//����¼����ʾ��tablewidget
bool GetRecordMatchTar(_RecordsetPtr pRecotd, _bstr_t TableName, _bstr_t ColumnName, _bstr_t DataType);//���������д���
int GetTableRowCount(_ConnectionPtr pRecord, _bstr_t ColumnName, _bstr_t TableName);//��ȡĳ����������/����ı����ֻ����һ������/������ֶ��в��ܰ���������
_bstr_t GetDefaultValue(int Code);//�ֶ�Ϊ�ǿ�ʱ��Ĭ���ֶ�
bool IfSrouceTableInTarget(_RecordsetPtr Source, _RecordsetPtr Target);//�ж�Դ�ļ��еı���Ŀ�����Ƿ���ͬ����
MDBTools::MDBTools(QWidget *parent)
	: QMainWindow(parent),m_fileCount(0), m_TarSchema(nullptr)
{
	ui.setupUi(this);
	g_TarFileName = (_bstr_t)T2W(L"");
	if (FAILED(::CoInitialize(NULL)))
	{
		QMessageBox msgBox;
		msgBox.setText("��ʼ��COM���ʧ�ܣ�");
		msgBox.exec();
		
		ui.label->setText(QString::fromStdWString(L"��ʼ��COM���ʧ�ܣ�"));
		return;
	}
	
	//ui.setWindowIcon(QIcon(":/UploadClient/UploadClient.ico"));
	//this->setFixedHeight(this->height());
	//�����ļ��б��߶�
	//ui.tabWidget->addTab(this->ui.tabWidget, "thires");
	//ui.tabWidget->setCurrentIndex(0);//������ʾ��ҳǩ
	//ui.tableWidget->setFixedHeight(this->height() / 5);//�����ؼ��߶�
	ui.tableWidget->setShowGrid(true);//��ʾ����
	ui.tableWidget->setColumnWidth(0, this->width() / 4);//���ó�ʼ�ĵ�Ԫ����
	ui.tableWidget->setColumnWidth(1, this->width() / 4);
	ui.tableWidget->setColumnWidth(2, this->width() / 4);
	ui.tableWidget->setColumnWidth(3, this->width() / 4);
	//ui.listWidget->setCellWidget(row, column, new QLineEdit);
	/*
	ui.tableWidget->setRowCount(2);//���ñ����õ��������
	QTableWidgetItem *TItem;//����һ��������
	TItem = new QTableWidgetItem(QString::fromStdWString(L"�ϲ�ȫ����ɣ�"));//���캯��
	ui.tableWidget->setItem(0,0, TItem);//������ʾ����
	*/
	ui.mainToolBar->setToolButtonStyle(Qt::ToolButtonTextUnderIcon);//���ö���������ʾģʽ
	//ui.mainToolBar->toolButtonStyle(Qt::ToolButtonTextBesideIcon);
	QToolBar* pToolBar = ui.mainToolBar;//��ȡ������ָ��
	QAction* ActButton = new QAction(QIcon(QPixmap(":/MDBTools/100.png")), QString::fromStdWString(L"Ŀ���ļ�"));//����һ����ť
	pToolBar->addAction(ActButton);//���뵽������
	QObject::connect(ActButton, SIGNAL(triggered(bool)), this, SLOT(slots_OnTargetFileSelect()));//�󶨲�
	ActButton = new QAction(QIcon(QPixmap(":/MDBTools/101.png")), QString::fromStdWString(L"��Դ�ļ�"));
	pToolBar->addAction(ActButton);
	QObject::connect(ActButton, SIGNAL(triggered(bool)), this, SLOT(slots_OnSourceFileSelect()));
	ActButton = new QAction(QIcon(QPixmap(":/MDBTools/102.png")), QString::fromStdWString(L"��ʼ�ϲ�"));
	pToolBar->addAction(ActButton);
	QObject::connect(ActButton, SIGNAL(triggered(bool)), this, SLOT(slots_OnStartCombineClick()));
	ActButton = new QAction(QIcon(QPixmap(":/MDBTools/103.png")), QString::fromStdWString(L"ʹ��˵��"));
	pToolBar->addAction(ActButton);
	QObject::connect(ActButton, SIGNAL(triggered(bool)), this, SLOT(slots_OnHelpClick()));
	QCoreApplication::processEvents();//������VBA DoEvents
	MDB = new MDBClass();
	MDBCM = new MDBCombine();
	ui.progressBar->setRange(ProGressMin, ProGressMax);
	ui.progressBar->setValue(ProGressMin);
	ui.statusBar->setToolTip(QString::fromStdWString(L"Ŀ���ļ�"));
	ui.statusBar->showMessage(QString::fromStdWString(L"������"));
}

void MDBTools::slots_OnTargetFileSelect()
{
	if (nullptr != m_TarSchema && ui.tabWidget_2->count() >= 2) ui.tabWidget_2->removeTab(1);
	ui.label->setText(QString::fromStdWString(L"Ŀ���ļ�   "));
	g_TarFileName = MDB->GetSingleFileName();//�Ȼ�ȡ�ļ�����
	if (g_TarFileName == (_bstr_t)T2W(L"")) return;
	std::wstring ws = (wchar_t *)g_TarFileName;//��Ҫ������ʾ��ҪתΪqstring
	ui.label->setText(QString::fromStdWString(L"Ŀ���ļ�   ")+ QString::fromStdWString(ws));//������ʾ
	MDBCM->SetConnectInfo(g_TarFileName);
	if (!MDBCM->Connect()) return;
	//if (!MDBCM->GetFileSchema()) return;
	if (!MDBCM->GetTableSchema()) return;
	//MDBCM->GetTableConstraint();	
	m_TarSchema = new QTableWidget();//����һ���µ�TabWidgetʵ��
	ui.tabWidget_2->insertTab(1, m_TarSchema, QString::fromStdWString(L"Ŀ���ļ�����"));//��Tab���½�Tab,��������ؼ�
	//_RecordsetPtr TT1 = MDBCM->get_TableSchema();
	RecordSetDisplayInTableWidget(MDBCM->get_TableSchema(), m_TarSchema);
	ui.tabWidget_2->setCurrentIndex(1);
}
void RecordSetDisplayInTableWidget(_RecordsetPtr pRSP,QTableWidget* QWD)
{
	QStringList headers;//һ�����ڼ�¼��ͷ���ı�����
	QWD->setColumnCount(pRSP->Fields->Count);//���������
	int i = 0;
	while (i < pRSP->Fields->Count)
	{
		headers << QString(pRSP->Fields->Item[long(i)]->Name);
		OutputDebugString(LPCWSTR(pRSP->Fields->Item[long(i)]->Name) + (_bstr_t)T2W(_T("\n")));
		_variant_t dd = pRSP->Fields->Item[long(i)]->Name;
		i++;
	}
	QWD->setHorizontalHeaderLabels(headers);//д���ͷ
	int j = -1;
	while (!(pRSP->EndOfFile))//ѭ��¼������
	{
		j++;
		QWD->setRowCount(j + 1);//���������
		for (int i = 0; i < pRSP->Fields->Count; i++)
		{
			QTableWidgetItem *TItem;
			_bstr_t temp = MDBCM->ConvertVariantV2(pRSP->Fields->GetItem(long(i))->Value);
			TItem = new QTableWidgetItem(QString::fromStdWString((std::wstring)MDBCM->ConvertVariantV2(pRSP->Fields->GetItem(long(i))->Value)));
			//QWD->setRowCount(j);
			//QWD->setColumnCount(i);
			QWD->setItem(j, i, TItem);
		}
		pRSP->MoveNext();
	}
}

_bstr_t ReplaceDYH(_bstr_t Words)//�滻���ݿ��д��е����ŵ����� �滻������
{
	CString Temp_4 = CString((LPCSTR)Words);//��ʼ��һ��CString
	Temp_4.Replace(L"'", L"");//�滻��������
	return _bstr_t(Temp_4);
}
_bstr_t AddDYH(_bstr_t Words)
{
	CString Temp_4 = CString((LPCSTR)Words);//��ʼ��һ��CString
	return _bstr_t(L"'" + Temp_4 + L"'");
}
long GetMDBFileMaxRow(_RecordsetPtr pRs,_ConnectionPtr pCon, _RecordsetPtr pTar)//��ȡ���ļ����������//��������Ķ�
{									//Դ�ļ�
	long RowCount = 0;//���м���
	long FirstRowCount = 0;//����ȡֵ
	//_RecordsetPtr pConRcrd;//
	//pConRcrd = pCon->Open();
	try {
		_bstr_t tablename = ReplaceDYH(MDBCM->ConvertVariantV2(pRs->Fields->GetItem("TABLE_NAME")->Value));//��ʼ����//tablename������һ��������� �Ա�͵�ǰ��Ա�
			while (!pRs->EndOfFile)//������¼
			{
				
				if ((CString(W2T(ReplaceDYH(MDBCM->ConvertVariantV2(pRs->Fields->GetItem("TABLE_NAME")->Value))))).Left(2) != "MS")//�����MS��ͷ����ϵͳ���ϵͳ��񲻽��кϲ�
				{
					//if (IfSrouceTableInTarget(pRs, pTar))
					{
						//��һ�����У���ȡ��һ���������
						OutputDebugString(_bstr_t(">>>>>>>>>>")  + ReplaceDYH(MDBCM->ConvertVariantV2(pRs->Fields->GetItem("TABLE_NAME")->Value)) + _bstr_t("<<<<<<<<<<\n"));
						if (FirstRowCount == 0) FirstRowCount = GetTableRowCount(pCon, ReplaceDYH(MDBCM->ConvertVariantV2(pRs->Fields->GetItem("COLUMN_NAME")->Value)), ReplaceDYH(MDBCM->ConvertVariantV2(pRs->Fields->GetItem("TABLE_NAME")->Value)));
						_RecordsetPtr pRRs;
						_bstr_t Temp10 = MDBCM->ConvertVariantV2(pRs->Fields->GetItem("TABLE_NAME")->Value);//��ȡ����ǰ�ı���
						if (ReplaceDYH(tablename) != ReplaceDYH(Temp10))//���������±�������������������ᵼ�µ�һ������������� ���е�һ����Ҫ����������
						{
							//_variant_t re;
							//_bstr_t Temp = bstr_t(L"select count(") + ReplaceDYH(MDBCM->ConvertVariantV2(pRs->Fields->GetItem("COLUMN_NAME")->Value)) +
								//_bstr_t(L") as rowcount from ") + ReplaceDYH(MDBCM->ConvertVariantV2(pRs->Fields->GetItem("TABLE_NAME")->Value));
							RowCount += GetTableRowCount(pCon, ReplaceDYH(MDBCM->ConvertVariantV2(pRs->Fields->GetItem("COLUMN_NAME")->Value)), ReplaceDYH(MDBCM->ConvertVariantV2(pRs->Fields->GetItem("TABLE_NAME")->Value)));
							OutputDebugString(_bstr_t("���м���     :") + ReplaceDYH(MDBCM->ConvertVariantV2(pRs->Fields->GetItem("TABLE_NAME")->Value)) + _bstr_t("  :   ") +
								_bstr_t(GetTableRowCount(pCon, ReplaceDYH(MDBCM->ConvertVariantV2(pRs->Fields->GetItem("COLUMN_NAME")->Value)), ReplaceDYH(MDBCM->ConvertVariantV2(pRs->Fields->GetItem("TABLE_NAME")->Value)))) + _bstr_t("\n"));
							OutputDebugString(_bstr_t(CString(W2T(tablename))) + _bstr_t(L"   ���ҽ��   ") + _bstr_t((CString(W2T(tablename))).Left(2) != "MS") + _bstr_t(L"\n"));
						}
					}
					//pRRs->open
				}
				tablename = ReplaceDYH(MDBCM->ConvertVariantV2(pRs->Fields->GetItem("TABLE_NAME")->Value));
				pRs->MoveNext();

			}
	}
	catch (_com_error &e)
	{
		MessageBox(NULL,e.Description(),L"",NULL);
	}
	OutputDebugString(_bstr_t(L"���м�¼����   ") + _bstr_t(RowCount + FirstRowCount) + _bstr_t(L"�У�"));
	return RowCount + FirstRowCount;
}
bool IfSrouceTableInTarget(_RecordsetPtr Source, _RecordsetPtr Target)//�ж�Դ�ļ��еı���Ŀ�����Ƿ���ͬ����
{
	Source->MoveFirst();
	Target->MoveFirst();
	while (!Target->EndOfFile)
	{
		while (!Source->EndOfFile)
		{
			OutputDebugString(ReplaceDYH(MDBCM->ConvertVariantV2(Source->Fields->GetItem("TABLE_NAME")->Value)) + _bstr_t(L"    �Ա�    ") + ReplaceDYH(MDBCM->ConvertVariantV2(Target->Fields->GetItem("TABLE_NAME")->Value)) + _bstr_t(L"    \n"));
			if (ReplaceDYH(MDBCM->ConvertVariantV2(Source->Fields->GetItem("TABLE_NAME")->Value)) == ReplaceDYH(MDBCM->ConvertVariantV2(Target->Fields->GetItem("TABLE_NAME")->Value)))
			{
				OutputDebugString(ReplaceDYH(MDBCM->ConvertVariantV2(Source->Fields->GetItem("TABLE_NAME")->Value)) + _bstr_t(L"    �Ա�    ") + ReplaceDYH(MDBCM->ConvertVariantV2(Target->Fields->GetItem("TABLE_NAME")->Value)) + _bstr_t(L"    \n"));
				return true;
			}
			Source->MoveNext();
		}
		Target->MoveNext();
	}
	return false;
}
int GetTableRowCount(_ConnectionPtr pRecord,_bstr_t ColumnName,_bstr_t TableName=L"")//��ȡĳ����������/����ı����ֻ����һ������/������ֶ��в��ܰ���������
{
	_RecordsetPtr pRRs;
	_variant_t re;
	if (TableName == _bstr_t(L""))
	{
		_bstr_t Temp1 = bstr_t(L"select * from ") + ColumnName + bstr_t(L" ;");
		try {
			pRRs = pRecord->Execute(Temp1, &re, adCmdText);
		}
		catch (...)
		{
			return 0;
		}
		return 1;
	}
	int RowCount = 0;

	_bstr_t Temp = bstr_t(L"select count(") + ColumnName +
		_bstr_t(L") as rowcount from ") + TableName;
	try {
		pRRs = pRecord->Execute(Temp, &re, adCmdText);
	}
	catch (...)
	{
		return 0;
	}
	if (!(pRRs->EndOfFile && pRRs->BOF))
	{
		pRRs->MoveFirst();
		RowCount += _ttoi((CString(W2T(MDBCM->ConvertVariantV2(pRRs->Fields->GetItem("rowcount")->Value)))));
	}
	return RowCount;
}
bool MDBTools::CombineFile(_RecordsetPtr Tar, MDBCombine* MDBobj)//�ϲ��ļ�������
{
	//�˴�λ�ϲ�����Ҫ����
	//���� 
	long TableWidgetRowCount = 1;
	long CurrentRow = 0;
	long MaxRow = 0;
	//MaxRow=MDBobj->get_TableSchema()->PageCount;
	MaxRow = GetMDBFileMaxRow(MDBobj->get_TableSchema(), MDBobj->get_ConnectPtr(),Tar);
	(*m_ControlsDocker.at(ProGressIndex)).setRange(0, MaxRow);
	/***************************************************
	1.�жϱ�����ͬ��������ͬ������������ͬ��Լ����ͬ
	****************************************************/
	//MDBobj->get_TableSchema()->Open("SELECT * FROM CBDKXX"/*MDBobj->get_ConnectionString()*/,
		//MDBCM->get_ConnectPtr().GetInterfacePtr(), adOpenStatic, adLockOptimistic, adCmdTable);
	//long MaxRow = MDBobj->get_TableSchema()->GetRecordCount();
	_RecordsetPtr Sou = MDBobj->get_TableSchema();
	_bstr_t LastTableName = T2W(L"");//��ʱ�ÿ�
	_bstr_t ColumnName = L"";
	try {
		Tar->MoveFirst();
		//��ʼ��ʼ��
		LastTableName = MDBCM->ConvertVariantV2(Tar->Fields->GetItem("TABLE_NAME")->Value);
		//Tar->Fields->GetItem("TABLE_NAME")->Value;
		//Tar->Open(MDBCM->get_ConnectionString(), _variant_t(MDBCM->get_ConnectPtr()), adOpenStatic, adLockReadOnly, adCmdText);
		//bool aa = Tar->GetState() == adStateOpen;
			//Tar->MoveFirst();
		HRESULT Hr = S_OK;
		_bstr_t NeceColumnValue = L"";//�ǿ���
		_bstr_t UnNeceColumnValue = L"";//�ɿմ�����
		_bstr_t INSERTTEXT = L"";//
		vector<_bstr_t> ColumnNamesOfTable_Nece;//������
		vector<_bstr_t> ColumnNamesOfTable_NeceNotExist;//���뵫��ԭ���в����ڵ���
		vector<_bstr_t> ColumnNamesOfTable_UnNece;//�Ǳ�����
		vector<_bstr_t> ColumnNamesOfTable;//����
		vector<MapWord> ColumnDataMap;//����
		//��ʼ��������
		while (!Tar->EndOfFile)//��ʼ��ȡ
		{
			//if ((CString(W2T((_bstr_t)(Tar->Fields->GetItem("TABLE_NAME")->Value)))).Left(2) != "MS")//��MS��ͷ�Ĵ��ʱϵͳ��񣬺���
			
				HRESULT HRS = S_OK;
				_bstr_t SQL = L"";
				//_bstr_t ColumnName = L"";
				//if (Sou->Supports(adSeek) == VARIANT_FALSE) return false;//��ѯ�Ƿ�֧��seek
				if (MDBCM->ConvertVariantV2(Tar->Fields->GetItem("TABLE_NAME")->Value) == LastTableName)//���ж��Ƿ���Ҫ�ύ
					//��������仯��ʱ���ύ��һ�����ĸĶ�
				{
					if (MDBCM->ConvertVariantV2(Tar->Fields->GetItem("IS_NULLABLE")->Value) == _bstr_t(L"FALSE"))
					{
						//�����ʱ�Ǳ������ô�ڲ����ʱ����в��ң��Ҳ����򲻲������
						//_bstr_t SQLText = _bstr_t(L"TABLE_NAME = '") + MDBCM->ConvertVariantV2(Tar->Fields->GetItem("TABLE_NAME")->Value) +bstr_t(L"' and ");
						//Hr=Tar->Seek(_variant_t("dkxx,fbf"), adSeekFirstEQ);��֧��
						//HRS = Tar->Find(_bstr_t(L"TABLE_NAME like '%")+ MDBCM->ConvertVariantV2(Tar->Fields->GetItem("TABLE_NAME")->Value)+_bstr_t(L"%'"),1,adSearchForward);
						_bstr_t Temp_0 = MDBCM->ConvertVariantV2(Tar->Fields->GetItem("COLUMN_NAME")->Value);//��ȡ������
						if (GetRecordMatchTar(Sou, MDBCM->ConvertVariantV2(Tar->Fields->GetItem("TABLE_NAME")->Value),
							MDBCM->ConvertVariantV2(Tar->Fields->GetItem("COLUMN_NAME")->Value),
							MDBCM->ConvertVariantV2(Tar->Fields->GetItem("DATA_TYPE")->Value)))//�жϸ�������Դ���ݿ����Ƿ���ڣ����������
						{
							MapWord ThisColumnReqiuire;
							ThisColumnReqiuire.ColumnName = ReplaceDYH( MDBCM->ConvertVariantV2(Tar->Fields->GetItem("COLUMN_NAME")->Value));
							ThisColumnReqiuire.DefaultValue = GetDefaultValue(atoi((const char*)MDBCM->ConvertVariantV2(Tar->Fields->GetItem("DATA_TYPE")->Value)));
							ThisColumnReqiuire.LengthRequire = atoi((const char*)MDBCM->ConvertVariantV2(Tar->Fields->GetItem("CHARACTER_MAXIMUM_LENGTH")->Value));
							OutputDebugString(_bstr_t(L"�ǿ��ֶ�:")+ ThisColumnReqiuire.ColumnName + _bstr_t(L"   Ĭ��ֵΪ:")+ ThisColumnReqiuire.DefaultValue  + _bstr_t(L"   Ҫ�󳤶�:") + _bstr_t(ThisColumnReqiuire.LengthRequire) + _bstr_t(L"\n"));
							m_TabDocker.at(ProGressIndex)->setRowCount(TableWidgetRowCount++);
							m_TabDocker.at(ProGressIndex)->setItem(TableWidgetRowCount - 2, 0, new QTableWidgetItem(QString::fromStdWString((wchar_t *)(_bstr_t(L"�ǿ��ֶ�:") + ThisColumnReqiuire.ColumnName + _bstr_t(L"   Ĭ��ֵΪ:") + ThisColumnReqiuire.DefaultValue + _bstr_t(L"   Ҫ�󳤶�:") + _bstr_t(ThisColumnReqiuire.LengthRequire) + _bstr_t(L"\n")))));
							ColumnDataMap.push_back(ThisColumnReqiuire);
							_bstr_t Temp_3 = MDBCM->ConvertVariantV2(Tar->Fields->GetItem("COLUMN_NAME")->Value);//��ȡ������
							CString Temp_4 = CString((LPCSTR)Temp_3);//��ʼ��һ��CString
							Temp_4.Replace(L"'", L"");//�滻��������
							//_bstr_t Temp_2 = _bstr_t(CString((LPCSTR)MDBCM->ConvertVariantV2(Tar->Fields->GetItem("COLUMN_NAME")->Value)).Replace(L"'", L""));
							if (NeceColumnValue == _bstr_t(L""))//�����ֶ�
								NeceColumnValue += _bstr_t(Temp_4);//�ж��ۼ�
							else
								NeceColumnValue += L"," + _bstr_t(Temp_4);//��Ҫ�Ӷ��ŵ��ۼ�

							ColumnNamesOfTable_Nece.push_back(MDBCM->ConvertVariantV2(Tar->Fields->GetItem("COLUMN_NAME")->Value));//����������vector
						}
						else//�жϸ�������Դ���ݿ����Ƿ���ڣ������������
						{
							CString Temp_5 = CString((LPCSTR)MDBCM->ConvertVariantV2(Tar->Fields->GetItem("COLUMN_NAME")->Value));//ȥ������
							Temp_5.Replace(L"'", L"");//ȥ������
							if (INSERTTEXT == _bstr_t(L""))//�ж��ۼӷ�ʽ
								INSERTTEXT = GetDefaultValue(atoi((const char*)MDBCM->ConvertVariantV2(Tar->Fields->GetItem("DATA_TYPE")->Value))) + L" as " +
								_bstr_t(Temp_5);//Ϊ�ո���
							else//�ж��ۼӷ�ʽ
								INSERTTEXT += _bstr_t(L",") + GetDefaultValue(atoi((const char*)MDBCM->ConvertVariantV2(Tar->Fields->GetItem("DATA_TYPE")->Value))) + L" as " +
								_bstr_t(Temp_5);//��ֵ�ۼ�
							ColumnNamesOfTable_NeceNotExist.push_back(MDBCM->ConvertVariantV2(Tar->Fields->GetItem("COLUMN_NAME")->Value));//����������vector
						}
						//else ���ԭ���в����ڸ��У������
					}
					else//����п���Ϊ��
					{
						//_bstr_t Temp_1 = MDBCM->ConvertVariantV2(Tar->Fields->GetItem("COLUMN_NAME")->Value);//������
						if (GetRecordMatchTar(Sou, MDBCM->ConvertVariantV2(Tar->Fields->GetItem("TABLE_NAME")->Value),
							MDBCM->ConvertVariantV2(Tar->Fields->GetItem("COLUMN_NAME")->Value),
							MDBCM->ConvertVariantV2(Tar->Fields->GetItem("DATA_TYPE")->Value)))//�жϸ����Ƿ���ԭ���д���
						{
							//��λ����ʱ�����Ҹ��У��Ҳ�����ʹ��Ĭ��ֵ�������
							_bstr_t Temp_3 = MDBCM->ConvertVariantV2(Tar->Fields->GetItem("COLUMN_NAME")->Value);//��ȡ������
							CString Temp_4 = CString((LPCSTR)Temp_3);//��ʼ��һ��CString
							Temp_4.Replace(L"'", L"");//�滻��������
							if (UnNeceColumnValue == _bstr_t(L""))//�ɿ��ֶ�
								UnNeceColumnValue += _bstr_t(Temp_4);//�ۼӷ�ʽ
							else
								UnNeceColumnValue += L"," + _bstr_t(Temp_4);//�ۼӷ�ʽ
							ColumnNamesOfTable_UnNece.push_back(MDBCM->ConvertVariantV2(Tar->Fields->GetItem("COLUMN_NAME")->Value));//����vector
						}
						//else//ԭ���в����ڸ��� ��Ϊ�ɿ� ���Է���
						//{
							//_bstr_t x;
							//int i = atoi((const char*)x);
							//�Ǳ����ֶβ鲻��������
						//}
					}
				}
				else//��ͬʱ������ύ����
				{
					vector<_bstr_t>::iterator iter;
					//MDB->GetTableList();
					for (iter = ColumnNamesOfTable_Nece.begin(); iter != ColumnNamesOfTable_Nece.end(); iter++)
					{
						CString T1 = CString((LPCSTR)*iter);
						T1.Replace(L"'", L"");
						ColumnNamesOfTable.push_back(_bstr_t(T1));
						OutputDebugString(_bstr_t(L"��������ֶ�:")+ _bstr_t(T1)+_bstr_t(L"\n"));
						m_TabDocker.at(ProGressIndex)->setRowCount(TableWidgetRowCount++);
						m_TabDocker.at(ProGressIndex)->setItem(TableWidgetRowCount - 2, 0, new QTableWidgetItem(QString::fromStdWString((wchar_t *)(_bstr_t(L"��������ֶ�:") + _bstr_t(T1) + _bstr_t(L"\n")))));
					}
					for (iter = ColumnNamesOfTable_UnNece.begin(); iter != ColumnNamesOfTable_UnNece.end(); iter++)
					{
						CString T2 = CString((LPCSTR)*iter);
						T2.Replace(L"'", L"");
						ColumnNamesOfTable.push_back(_bstr_t(T2));
					}
					for (iter = ColumnNamesOfTable_NeceNotExist.begin(); iter != ColumnNamesOfTable_NeceNotExist.end(); iter++)
					{
						CString T3 = CString((LPCSTR)*iter);
						T3.Replace(L"'", L"");
						ColumnNamesOfTable.push_back(_bstr_t(T3));
					}
					ColumnNamesOfTable_Nece.clear();
					ColumnNamesOfTable_UnNece.clear();
					ColumnNamesOfTable_NeceNotExist.clear();
					_bstr_t InsertText;
					for (iter = ColumnNamesOfTable.begin(); iter != ColumnNamesOfTable.end(); iter++)
					{
						if (InsertText == _bstr_t(L""))
						{
							//CString T1= CString((LPCSTR)*iter);
							//T1.Replace(L"'", L"");
							InsertText = *iter;//_bstr_t(T1);
						}
						else
						{
							//CString T2 = CString((LPCSTR)*iter);
							//T2.Replace(L"'", L"");
							InsertText += *iter; //_bstr_t(L",") + _bstr_t(T2);
						}
					}
					CString TableName = CString((LPCSTR)LastTableName);
					TableName.Replace(L"'", L"");
					LastTableName = _bstr_t(TableName);//ȥ�������еĵ�����
					_bstr_t Temp_0;//= _bstr_t(L"select ");//��ѯ��� ���������Ĵ������ ������
					if (NeceColumnValue != _bstr_t(L""))
						Temp_0 += NeceColumnValue;
					if (UnNeceColumnValue != _bstr_t(L""))
						Temp_0 += _bstr_t(L",") + UnNeceColumnValue;
					if (INSERTTEXT != _bstr_t(L""))
						Temp_0 += _bstr_t(L",") + INSERTTEXT;
					Temp_0 = _bstr_t(L"select ") + Temp_0 + _bstr_t(L" from ") + _bstr_t(TableName);
					//OutputDebugString(Temp_0+L"\n");//���ѡ���sql���
					//OutputDebugString((_bstr_t)T2W(_T("\n")) + UnNeceColumnValue + NeceColumnValue);
					_variant_t ExeResultStatus;
					_RecordsetPtr Result = MDBobj->get_ConnectPtr()->Execute(Temp_0, &ExeResultStatus, adCmdText);//�����е���Ϣ����Recordset�ٱ�������
					m_TabDocker.at(ProGressIndex)->setRowCount(TableWidgetRowCount++);
					m_TabDocker.at(ProGressIndex)->setItem(TableWidgetRowCount - 2, 0, new QTableWidgetItem(QString::fromStdWString((wchar_t *)Temp_0)));
					//long NowTheTableWidgetRowCount = m_TabDocker.at(ProGressIndex)->rowCount();
					while (!Result->EndOfFile)
					{
						_bstr_t InsertValue;
						_bstr_t InsertVolumns;
						//RecordSetDisplayInTableWidget(Result,ui.tableWidget);
						for (long i = 0; i < Result->Fields->Count; i++)
						{
							if (InsertValue == _bstr_t(L""))
							{

								for (long j = 0; j < Result->Fields->Count; j++)
								{
									//OutputDebugString(L"\n" + Result->Fields->Item[j]->Name);
								}
								//_bstr_t Temp = ColumnNamesOfTable.at(i);
								_bstr_t Value_Temp= MDBCM->ConvertVariantV2(Result->Fields->GetItem(_variant_t(ColumnNamesOfTable.at(i)))->Value);
								//**********************************
								vector<MapWord>::iterator iter;//���ǿ��ֶα��벻Ϊ��
								//MDB->GetTableList();
								for (iter = ColumnDataMap.begin(); iter != ColumnDataMap.end(); iter++)
								{
									if (iter->ColumnName == ColumnNamesOfTable.at(i))//�����ֶ�
									{
										if (MDBCM->ConvertVariantV2(Result->Fields->GetItem(_variant_t(ColumnNamesOfTable.at(i)))->Value) == _bstr_t(L"")
											|| MDBCM->ConvertVariantV2(Result->Fields->GetItem(_variant_t(ColumnNamesOfTable.at(i)))->Value) == _bstr_t(L"NULL")
											|| MDBCM->ConvertVariantV2(Result->Fields->GetItem(_variant_t(ColumnNamesOfTable.at(i)))->Value) == _bstr_t(L"''"))
										{
											Value_Temp = iter->DefaultValue;
											int TLen = CString(W2T(ReplaceDYH(iter->DefaultValue))).GetLength();
											if (!(TLen <= iter->LengthRequire || iter->LengthRequire == 0))//�������Ҫ�󳤶ȣ�Ҫ�󳤶ȱ��벻Ϊ0
											{
												Value_Temp = (CString(W2T(ReplaceDYH(iter->DefaultValue)))).Left(iter->LengthRequire);
												Value_Temp = AddDYH(Value_Temp);
											}
											break;
										}
									}
								}
								//**********************************
								InsertValue = Value_Temp;
								InsertVolumns = ColumnNamesOfTable.at(i);
							}
							else
							{
								_bstr_t Value_Temp = MDBCM->ConvertVariantV2(Result->Fields->GetItem(_variant_t(ColumnNamesOfTable.at(i)))->Value);
								//**********************************
								vector<MapWord>::iterator iter;//���ǿ��ֶα��벻Ϊ��
															   //MDB->GetTableList();
								for (iter = ColumnDataMap.begin(); iter != ColumnDataMap.end(); iter++)
								{
									if (iter->ColumnName == ColumnNamesOfTable.at(i))//�����ֶ�
									{
										if (LastTableName==_bstr_t(L"CBF_JTCY"))//��if������
										{
											_bstr_t Temp = MDBCM->ConvertVariantV2(Result->Fields->GetItem(_variant_t(ColumnNamesOfTable.at(i)))->Value);
											//OutputDebugString(Temp+L"\n");
											
										}
										if (MDBCM->ConvertVariantV2(Result->Fields->GetItem(_variant_t(ColumnNamesOfTable.at(i)))->Value) == _bstr_t(L"")
											|| MDBCM->ConvertVariantV2(Result->Fields->GetItem(_variant_t(ColumnNamesOfTable.at(i)))->Value) == _bstr_t(L"NULL")
											|| MDBCM->ConvertVariantV2(Result->Fields->GetItem(_variant_t(ColumnNamesOfTable.at(i)))->Value) == _bstr_t(L"''"))
										{
											Value_Temp = iter->DefaultValue;
											int TLen = CString(W2T(ReplaceDYH(iter->DefaultValue))).GetLength();
											if (!(TLen <= iter->LengthRequire || iter->LengthRequire == 0))//
											{
												Value_Temp = (CString(W2T(ReplaceDYH(iter->DefaultValue)))).Left(iter->LengthRequire);
												Value_Temp = AddDYH(Value_Temp);
											}
											break;
										}
									}
								}
								//**********************************
								InsertValue += L"," + Value_Temp;
								//InsertValue += L"," + MDBCM->ConvertVariantV2(Result->Fields->GetItem(_variant_t(ColumnNamesOfTable.at(i)))->Value);
								InsertVolumns += L"," + ColumnNamesOfTable.at(i);
							}
						}
						_variant_t ExeResult;
						if ((CString(W2T(LastTableName))).Left(2) != "MS")//����ϵͳ��
						{
							_bstr_t Temp = L"insert into " + LastTableName + L"(" + InsertVolumns + L")Values(" + InsertValue + ")";
							OutputDebugString(Temp + L"\n");//����������
							CurrentRow++;
							MDBCM->get_ConnectPtr()->Execute(Temp, &ExeResult, adCmdText);
							(*m_ControlsDocker.at(ProGressIndex)).setValue(CurrentRow);//���ý�����
							QTableWidgetItem *TItem;//����һ��������
							TItem = new QTableWidgetItem(_bstr_t(CurrentRow) + _bstr_t(L"/") + QString::fromStdWString((wchar_t *)( _bstr_t((*m_ControlsDocker.at(ProGressIndex)).maximum())  + _bstr_t(L"  �����/����Ŀ")  )));//���캯��
							//long ProGressValue = ui.progressBar->value() + EveryProGressMax * (float(CurrentRow / float((*m_ControlsDocker.at(ProGressIndex)).maximum())));
							long ProGressValue = EveryProGressMax*(float(CurrentRow) /(*m_ControlsDocker.at(ProGressIndex)).maximum());
							//m_TabDocker.at(CurrentRow)->setColumnCount()
							//m_TabDocker.at(ProGressIndex)->setRowCount(TableWidgetRowCount++);
							//m_TabDocker.at(ProGressIndex)->setItem(TableWidgetRowCount, 1, new QTableWidgetItem(_bstr_t(L"��������ֶ�:") + _bstr_t(T1) + _bstr_t(L"\n")));
							m_TabDocker.at(ProGressIndex)->setRowCount(TableWidgetRowCount++);
							QTableWidgetItem *TabTItem = new QTableWidgetItem(QString::fromStdWString((wchar_t *)Temp));//����һ�������� 
							m_TabDocker.at(ProGressIndex)->setItem(TableWidgetRowCount - 2,0, TabTItem);
							ui.progressBar->setValue(ProGressIndex*EveryProGressMax + ProGressValue);
							ui.tableWidget->setItem(ProGressIndex, 2, TItem);//��ʾ����
							QApplication::processEvents();
							//delete TItem;
							//delete TabTItem;
						}
						Result->MoveNext();
						//MDBCM->get_ConnectPtr()->Execute();
					}
					ColumnNamesOfTable.clear();
					NeceColumnValue = L"";
					UnNeceColumnValue = L"";
					INSERTTEXT = L"";
					//MessageBox(NULL, UnNeceColumnValue + NeceColumnValue,L"",MB_OK);
				}
				bool IfNewTable = MDBCM->ConvertVariantV2(Tar->Fields->GetItem("TABLE_NAME")->Value) == LastTableName;//�������ж����ϱ�ͬ�򱾴�û��ִ�б��ĺϲ����̣����Բ������α����ƶ���
				LastTableName = MDBCM->ConvertVariantV2(Tar->Fields->GetItem("TABLE_NAME")->Value);
				if (IfNewTable)
					Tar->MoveNext();
				//else
					//Tar->MovePrevious();
			
		}
		//�ϲ����
		
	}
	catch (_com_error &e)
	{
		
		_bstr_t bb = e.Description();
		//MessageBox(NULL, L"��������\n"+e.Description() + L"\n" + _bstr_t((CString(W2T((_bstr_t)(Tar->Fields->GetItem("TABLE_NAME")->Value)))).Left(2)), L"������ʾ��", NULL);
		//return false;
		m_TabDocker.at(ProGressIndex)->setRowCount(TableWidgetRowCount++);
		m_TabDocker.at(ProGressIndex)->setItem(TableWidgetRowCount - 2, 0, new QTableWidgetItem(QString::fromStdWString((wchar_t *)(L"��������" + e.Description() + _bstr_t((CString(W2T((_bstr_t)(Tar->Fields->GetItem("TABLE_NAME")->Value)))).Left(2)) + L"������ʾ��"))));
	
	}
	return true;
}
bool GetRecordMatchTar(_RecordsetPtr pRecord, _bstr_t TableName, _bstr_t ColumnName, _bstr_t DataType)//��ȡ����Դ�ļ���¼��ͬ����
{
	pRecord->MoveFirst();
	while (!pRecord->EndOfFile)
	{
		_bstr_t TABLE_NAME = MDBCM->ConvertVariantV2(pRecord->Fields->GetItem("TABLE_NAME")->Value);
		_bstr_t COLUMN_NAME = MDBCM->ConvertVariantV2(pRecord->Fields->GetItem("COLUMN_NAME")->Value);
		_bstr_t DATA_TYPE = MDBCM->ConvertVariantV2(pRecord->Fields->GetItem("DATA_TYPE")->Value);
		if (MDBCM->ConvertVariantV2(pRecord->Fields->GetItem("TABLE_NAME")->Value) == TableName &&
			MDBCM->ConvertVariantV2(pRecord->Fields->GetItem("COLUMN_NAME")->Value) == ColumnName &&
			MDBCM->ConvertVariantV2(pRecord->Fields->GetItem("DATA_TYPE")->Value) == DataType)
		{
			return true;
		}
		pRecord->MoveNext();
	}
	return false;
}
_bstr_t GetDefaultValue(int Code)//��ȡ��ĳ�����͵�Ĭ��ֵ
{
	if (Code == 0x2000) return _bstr_t(L"'AdArray'");
	if (Code == 20) return _bstr_t(L"0");//_bstr_t(L"adBigInt");
	if (Code == 128) return _bstr_t(L"'adBinary'");
	if (Code == 11) return _bstr_t(L"'FALSE'");//_bstr_t(L"adBoolean");
	if (Code == 8) return _bstr_t(L"'adBSTR'");
	if (Code == 136) return _bstr_t(L"'adChapter'");
	if (Code == 129) return _bstr_t(L"'adChar'");
	if (Code == 6) return _bstr_t(L"'adCurrency'");
	if (Code == 7) return _bstr_t(L"'2000/01/01'");//_bstr_t(L"adDate");
	if (Code == 133) return _bstr_t(L"'2000/01/01'");//_bstr_t(L"adDBDate");
	if (Code == 134) return _bstr_t(L"'12:12:12'");//_bstr_t(L"adDBTime");
	if (Code == 135) return _bstr_t(L"'2000/01/01 12:12:12'");//_bstr_t(L"adDBTimeStamp");
	if (Code == 14) return _bstr_t(L"0.0");//_bstr_t(L"adDecimal");
	if (Code == 5) return _bstr_t(L"0.0");//_bstr_t(L"adDouble");
	if (Code == 0) return _bstr_t(L"'adEmpty'");
	if (Code == 10) return _bstr_t(L"'adError'");
	if (Code == 64) return _bstr_t(L"'2000/01/01'");//_bstr_t(L"adFileTime");
	if (Code == 72) return _bstr_t(L"'adGUID'");
	if (Code == 9) return _bstr_t(L"'adIDispatch'");
	if (Code == 3) return  _bstr_t(L"0");//_bstr_t(L"adInteger");
	if (Code == 13) return _bstr_t(L"'adIUnknown'");
	if (Code == 205) return _bstr_t(L"'adLongVarBinary'");
	if (Code == 201) return _bstr_t(L"'adLongVarChar'");
	if (Code == 203) return _bstr_t(L"'adLongVarWChar'");
	if (Code == 131) return _bstr_t(L"0");//_bstr_t(L"adNumeric");
	if (Code == 138) return _bstr_t(L"'adPropVariant'");
	if (Code == 4) return _bstr_t(L"0.0");//_bstr_t(L"adSingle");
	if (Code == 2) return _bstr_t(L"0");//_bstr_t(L"adSmallInt");
	if (Code == 16) return _bstr_t(L"0");//_bstr_t(L"adTinyInt");
	if (Code == 21) return _bstr_t(L"0");//_bstr_t(L"adUnsignedBigInt");
	if (Code == 19) return _bstr_t(L"0");//_bstr_t(L"adUnsignedInt");
	if (Code == 18) return _bstr_t(L"0");//_bstr_t(L"adUnsignedSmallInt");
	if (Code == 17) return _bstr_t(L"0");//_bstr_t(L"adUnsignedTinyInt");
	if (Code == 132) return _bstr_t(L"'adUserDefined'");
	if (Code == 204) return _bstr_t(L"'adVarBinary'");
	if (Code == 200) return _bstr_t(L"'adVarChar'");
	if (Code == 12) return _bstr_t(L"'adVariant'");
	if (Code == 139) return _bstr_t(L"0"); //_bstr_t(L"adVarNumeric");
	if (Code == 202) return _bstr_t(L"'adVarWChar'");
	if (Code == 130) return _bstr_t(L"'adWChar'");
	return _bstr_t(L"'NULL'");

}

void MDBTools::SetGlobleProGressValue(int Value)//�����������ϵĽ�����
{
	if (Value >= ProGressMin && Value <= ProGressMax)
		ui.progressBar->setValue(Value);
}

void MDBTools::ShowSourceFilesWeSelected(_bstr_t fileName,_bstr_t Percent)//��ʾѡ���Դ�ļ�������
{
	ui.statusBar->showMessage(QString::fromStdWString(L"���������ļ���Ϣ��   ") + Percent, 2000);
	MDBCombine TempClass(fileName);
	ui.tableWidget->setRowCount(ui.tableWidget->rowCount() + 1);//���������/���������һ��/�ǵ�ǰ��
	QTableWidgetItem *TItem;//����һ��������
	TItem = new QTableWidgetItem(QString::fromStdWString((wchar_t *)fileName));//���캯��
	/*****************����**************
	QMessageBox msgBox;
	msgBox.setText(QString::number(ui.tableWidget->rowCount(),10));
	msgBox.exec();
	*************************************/
	ui.tableWidget->setItem((ui.tableWidget->rowCount()-1), 0, TItem);//������ʾ����
	QProgressBar *EveryProgress = new QProgressBar();//�½�һ��������ʵ��
	EveryProgress->setTextVisible(false);//���ý��������ֲ��ɼ�
	ui.tableWidget->setCellWidget((ui.tableWidget->rowCount() - 1),1, EveryProgress);//���õ�����������tablewidget
	m_ControlsDocker.push_back(EveryProgress);//����������������
	//��̬���Tab�ķ���
	QTableWidget *EveryNewTab = new QTableWidget();//����һ���µ�Tabʵ��
	ui.tabWidget->insertTab(0, EveryNewTab, QString::number(m_fileCount,10));//��tab���½�tab,��������ؼ�
	m_TabDocker.push_back(EveryNewTab);
	QApplication::processEvents();
	if (!TempClass.Connect())
	{
		//delete TempClass;
		return;
	}
		
	if (!TempClass.GetTableSchema())
	{
		//delete TempClass;
		return;
	}
	//RecordSetDisplayInTableWidget(TempClass.get_TableSchema(), EveryNewTab);
	ui.tabWidget_2->setCurrentIndex(0);
	//delete TempClass;
}

void MDBTools::slots_OnSourceFileSelect()//ѡ��Դ�ļ�
{
	m_ControlsDocker.clear();
	if (m_ControlsDocker.size() != 0)
	{
		vector<QProgressBar *>::iterator iter;
		for (iter = m_ControlsDocker.begin(); iter != m_ControlsDocker.end(); iter++)
		{
			delete *iter;
		}

	}
	if (m_TabDocker.size() != 0)
	{
		vector<QTableWidget *>::iterator iter;
		for (iter = m_TabDocker.begin(); iter != m_TabDocker.end(); iter++)
		{
			delete *iter;
		}
	}
	m_fileCount = 0;
	ui.tableWidget->clearContents();
	ui.tabWidget->clear();
	ui.tableWidget->setRowCount(0);
	g_SourceFileNames = MDB->GetMultiFileName();//��ȡѡ�е��ļ�����
	if (0 == g_SourceFileNames.size()) return;
	//for_each(g_SourceFileNames.begin(), g_SourceFileNames.end(), std::mem_fun_ref(&MDBTools::ShowSourceFilesWeSelected));//������ʾ
	vector<_bstr_t>::iterator iter;
	//MDB->GetTableList();
	for (iter = g_SourceFileNames.begin(); iter != g_SourceFileNames.end(); iter++)
	{
		if (*iter != g_TarFileName)
		{
			m_fileCount++;
			ShowSourceFilesWeSelected(*iter, _bstr_t(m_fileCount) + _bstr_t("/") + _bstr_t(g_SourceFileNames.size()));
		}
	}
	ui.statusBar->showMessage(QString::fromStdWString(L"�ļ���Ϣ�Ѿ�ȫ��������ɣ�����"));
}

void MDBTools::slots_OnStartCombineClick()
{
	if (!MDBCM->get_ConnectPtr())
	{
		MessageBox(NULL,L"��û��ѡ��Ŀ���ļ���",L"������ʾ",MB_OK);
		return;
	}
  	if (g_SourceFileNames.size() == (size_t)0)
	{
		MessageBox(NULL, L"��û��ѡ��Ҫ�ϲ���Դ�ļ���", L"������ʾ", MB_OK);
		return;
	}
	vector<_bstr_t>::iterator iter;
	EveryProGressMax = ProGressMax / g_SourceFileNames.size();
	for (iter = g_SourceFileNames.begin(); iter != g_SourceFileNames.end(); iter++)
	{
		if (*iter != g_TarFileName)//Ŀ���Դ�ļ�������ͬһ�ļ�
		{
			ProGressIndex++;
			MDBCombine  pCombine(*iter);
			pCombine.Connect();
			pCombine.GetTableSchema();
			m_TabDocker.at(ProGressIndex)->setColumnCount(1);
			m_TabDocker.at(ProGressIndex)->setHorizontalHeaderLabels(QStringList()<<QString::fromStdWString(L"�ϲ���Ϣ"));
			m_TabDocker.at(ProGressIndex)->setColumnWidth(0,this->width());
			//m_TabDocker.at(ProGressIndex)->setRowCount(1000);
			if (!CombineFile(MDBCM->get_TableSchema(), &pCombine))
				ui.tableWidget->setItem(ProGressIndex,2,new QTableWidgetItem(QString::fromStdWString(L"�ϲ�ʧ�ܣ�")));
				//MessageBox(NULL, *iter + L"\n�ϲ�ʧ�ܣ�", L"��ʾ", NULL);
			//delete pCombine;
		}
	}
	delete MDBCM;
	MessageBox(NULL, L"�ϲ���ɣ�", L"��ʾ", MB_OK);
	//delete MDBCM;
}

void MDBTools::slots_OnHelpClick()
{
	QMessageBox msgBox;
	QString msg =
		QString::fromStdWString(L"��ע�⣬������������MDB�ļ��ϲ�����������Ҫע�������\n1.���кϲ����ļ��ı��У��������ͣ�") +
		QString::fromStdWString(L"Լ����ȫ������Ŀ���ļ�Ϊ��\n2.��������Ҫ��װx64��MDB��������\n3.ͬʱ��Ҫ��װ64λ��Office��������ʹ��") +
		QString::fromStdWString(L"\n4.��ԭ���е�������Ŀ�����е�Լ������ʱ������ԭ���ݽ��������Խ��кϲ�����ضϣ���ӳ�\n5.��ԭ���еķǿ�����ԭ����û��ʱ��������Ĭ�����ݣ�")+
		QString::fromStdWString(L"\n6.��ͬ���������Ͷ�Ӧ��Ĭ��ֵ��ͬ������������±�");
	msgBox.setText(msg);
	msgBox.exec();
}
