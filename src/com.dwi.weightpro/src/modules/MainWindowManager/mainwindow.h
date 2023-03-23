#pragma once
#include "allotmentmanager.h"
#include "materialcodemanager.h"
#include "partycodemanager.h"
#include "sodatamanager.h"
#include "src/modules/RfidManager/rfidmanager.h"
#include "wbcodemanager.h"
#ifndef MAINWINDOW_H
#define MAINWINDOW_H

#include <QMainWindow>
#include <QDialog>
#include <QSqlDatabase>
#include <QNetworkAccessManager>
#include <QDebug>
#include "psqlmanager.h"
#include "login.h"
#include <QLibrary>
#include "weighmachinethread.h"

#include <QTest>

#include "DebugManager.h"

namespace Ui {
    class MainWindow;
}

#include "sqlite3.h"

/***
 *  TODOS
 * - Main Window should know who the user is
 * - Let the PSQL manager know who the user is
 *
 ***/

class QApplication;
class DebugManager;

class MainWindow : public QMainWindow
{
    Q_OBJECT

    public:
        explicit MainWindow(QApplication *parent = 0,QString appname = "test");
        ~MainWindow();
        void whoami();
        PsqlManager db;
        int loginDialog();
        QString appname;
        QString username;
        void pSetup();


    public slots:
        bool logincheck(QString username, QString password, QString mode);

    private slots:
        void saveIpSettings();
        void on_pushButtonGetClicked();
        void on_pushButtonWbsaveClicked();
        void on_pushButtonGetweightClicked();
        void on_pushButtonPlayClicked();
        void on_serialChangeWeightgm();
        void onpushButtonLogClicked();
        void processFrame(QImage test);
        void onpushButtonGetInventoryClicked();
        void onpushButtonClearInventoryClicked();
        void onpushButtonRfidUpdateClicked();
        void onpushButtonRfidNewClicked();
        void onpushButtonRfidReissueClicked();
        void onActionLogoutTriggered();

        void onpushVehicleActionsClicked();
        void onpushSoManagerClicked();
        void onpushAllotmentManager();
        void onpushWbCodeManager();
        void onpushMaterialCodeManager();
        void onpushPartyCodeManager();

    signals:
        void rfidupdatecalled();
        void rfidnewcalled();
        void rfidreissuecalled();
        void vehicleactionscalled();
        void allotmentmanagercalled();
        void wbcodemanagercalled();
        void materialcodemanagercalled();
        void partycodemanagercalled();

    private:
        QString password;
        QString mode;
        QString accessLevel;
        DebugManager *debugger;
        WeighMachineThread *wthread;
        QNetworkAccessManager *manager;
        QNetworkRequest request;
        Ui::MainWindow *ui;
        Login *login;
        sqlite3* db1;
        RfidManager *rfidcontroller;
        SoDataManager *sodatacontroller;
        AllotmentManager *allotmanagercontroller;
        WbCodeManager *wbcodemanager;
        MaterialCodeManager *materialcodemanager;
        PartyCodeManager *partycodemanager;



        //
        void setupDefaults();
        void setupWeighMachine();
        void setupDatabase();
        void setupSignals();
        void setupIpSettingsDefaults();
        void setAdminIpSettingsRow(QString name, QString ip, int port, QString uname, QString password);
        void setupNetworkManager();
        void setupWeighBridgeSettingsDefaults();
//        void closeEvent (QCloseEvent *event);
        void closeEvent();

        bool MainWindow::superUserLoginCheck(QString username, QString password);


};



#endif // MAINWINDOW_H
