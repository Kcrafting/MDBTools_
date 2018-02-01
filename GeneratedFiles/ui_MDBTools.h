/********************************************************************************
** Form generated from reading UI file 'MDBTools.ui'
**
** Created by: Qt User Interface Compiler version 5.9.0
**
** WARNING! All changes made in this file will be lost when recompiling UI file!
********************************************************************************/

#ifndef UI_MDBTOOLS_H
#define UI_MDBTOOLS_H

#include <QtCore/QVariant>
#include <QtWidgets/QAction>
#include <QtWidgets/QApplication>
#include <QtWidgets/QButtonGroup>
#include <QtWidgets/QHBoxLayout>
#include <QtWidgets/QHeaderView>
#include <QtWidgets/QLabel>
#include <QtWidgets/QMainWindow>
#include <QtWidgets/QProgressBar>
#include <QtWidgets/QStatusBar>
#include <QtWidgets/QTabWidget>
#include <QtWidgets/QTableWidget>
#include <QtWidgets/QToolBar>
#include <QtWidgets/QVBoxLayout>
#include <QtWidgets/QWidget>

QT_BEGIN_NAMESPACE

class Ui_MDBToolsClass
{
public:
    QWidget *centralWidget;
    QVBoxLayout *verticalLayout;
    QLabel *label;
    QTabWidget *tabWidget_2;
    QWidget *tab;
    QHBoxLayout *horizontalLayout;
    QTableWidget *tableWidget;
    QTabWidget *tabWidget;
    QProgressBar *progressBar;
    QToolBar *mainToolBar;
    QStatusBar *statusBar;

    void setupUi(QMainWindow *MDBToolsClass)
    {
        if (MDBToolsClass->objectName().isEmpty())
            MDBToolsClass->setObjectName(QStringLiteral("MDBToolsClass"));
        MDBToolsClass->resize(948, 725);
        centralWidget = new QWidget(MDBToolsClass);
        centralWidget->setObjectName(QStringLiteral("centralWidget"));
        verticalLayout = new QVBoxLayout(centralWidget);
        verticalLayout->setSpacing(6);
        verticalLayout->setContentsMargins(11, 11, 11, 11);
        verticalLayout->setObjectName(QStringLiteral("verticalLayout"));
        label = new QLabel(centralWidget);
        label->setObjectName(QStringLiteral("label"));
        QFont font;
        font.setPointSize(14);
        label->setFont(font);

        verticalLayout->addWidget(label);

        tabWidget_2 = new QTabWidget(centralWidget);
        tabWidget_2->setObjectName(QStringLiteral("tabWidget_2"));
        tab = new QWidget();
        tab->setObjectName(QStringLiteral("tab"));
        horizontalLayout = new QHBoxLayout(tab);
        horizontalLayout->setSpacing(6);
        horizontalLayout->setContentsMargins(11, 11, 11, 11);
        horizontalLayout->setObjectName(QStringLiteral("horizontalLayout"));
        tableWidget = new QTableWidget(tab);
        if (tableWidget->columnCount() < 4)
            tableWidget->setColumnCount(4);
        QTableWidgetItem *__qtablewidgetitem = new QTableWidgetItem();
        tableWidget->setHorizontalHeaderItem(0, __qtablewidgetitem);
        QTableWidgetItem *__qtablewidgetitem1 = new QTableWidgetItem();
        tableWidget->setHorizontalHeaderItem(1, __qtablewidgetitem1);
        QTableWidgetItem *__qtablewidgetitem2 = new QTableWidgetItem();
        tableWidget->setHorizontalHeaderItem(2, __qtablewidgetitem2);
        QTableWidgetItem *__qtablewidgetitem3 = new QTableWidgetItem();
        tableWidget->setHorizontalHeaderItem(3, __qtablewidgetitem3);
        tableWidget->setObjectName(QStringLiteral("tableWidget"));
        tableWidget->setStyleSheet(QStringLiteral(""));

        horizontalLayout->addWidget(tableWidget);

        tabWidget_2->addTab(tab, QString());

        verticalLayout->addWidget(tabWidget_2);

        tabWidget = new QTabWidget(centralWidget);
        tabWidget->setObjectName(QStringLiteral("tabWidget"));

        verticalLayout->addWidget(tabWidget);

        progressBar = new QProgressBar(centralWidget);
        progressBar->setObjectName(QStringLiteral("progressBar"));
        progressBar->setValue(24);
        progressBar->setTextVisible(false);

        verticalLayout->addWidget(progressBar);

        MDBToolsClass->setCentralWidget(centralWidget);
        tabWidget->raise();
        progressBar->raise();
        label->raise();
        tabWidget_2->raise();
        mainToolBar = new QToolBar(MDBToolsClass);
        mainToolBar->setObjectName(QStringLiteral("mainToolBar"));
        MDBToolsClass->addToolBar(Qt::TopToolBarArea, mainToolBar);
        statusBar = new QStatusBar(MDBToolsClass);
        statusBar->setObjectName(QStringLiteral("statusBar"));
        MDBToolsClass->setStatusBar(statusBar);

        mainToolBar->addSeparator();

        retranslateUi(MDBToolsClass);

        tabWidget_2->setCurrentIndex(0);
        tabWidget->setCurrentIndex(-1);


        QMetaObject::connectSlotsByName(MDBToolsClass);
    } // setupUi

    void retranslateUi(QMainWindow *MDBToolsClass)
    {
        MDBToolsClass->setWindowTitle(QApplication::translate("MDBToolsClass", "MDBTools", Q_NULLPTR));
        label->setText(QApplication::translate("MDBToolsClass", "\347\233\256\346\240\207\346\226\207\344\273\266", Q_NULLPTR));
        QTableWidgetItem *___qtablewidgetitem = tableWidget->horizontalHeaderItem(0);
        ___qtablewidgetitem->setText(QApplication::translate("MDBToolsClass", "\346\226\207\344\273\266\345\220\215\347\247\260", Q_NULLPTR));
        QTableWidgetItem *___qtablewidgetitem1 = tableWidget->horizontalHeaderItem(1);
        ___qtablewidgetitem1->setText(QApplication::translate("MDBToolsClass", "\345\220\210\345\271\266\350\277\233\345\272\246", Q_NULLPTR));
        QTableWidgetItem *___qtablewidgetitem2 = tableWidget->horizontalHeaderItem(2);
        ___qtablewidgetitem2->setText(QApplication::translate("MDBToolsClass", "\345\220\210\345\271\266\347\273\223\346\236\234", Q_NULLPTR));
        QTableWidgetItem *___qtablewidgetitem3 = tableWidget->horizontalHeaderItem(3);
        ___qtablewidgetitem3->setText(QApplication::translate("MDBToolsClass", "\346\230\257\345\220\246\345\217\221\347\224\237\351\224\231\350\257\257", Q_NULLPTR));
        tabWidget_2->setTabText(tabWidget_2->indexOf(tab), QApplication::translate("MDBToolsClass", "\346\272\220\346\226\207\344\273\266\345\210\227\350\241\250", Q_NULLPTR));
    } // retranslateUi

};

namespace Ui {
    class MDBToolsClass: public Ui_MDBToolsClass {};
} // namespace Ui

QT_END_NAMESPACE

#endif // UI_MDBTOOLS_H
