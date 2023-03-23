#include "mainwindow.h"
#include <VideoFrameGrabber.h>
#include "ui_mainwindow.h"
#include "login.h"
#include <uhfreader.h>
#include <QNetworkReply>
#include <QCloseEvent>
#include <QCryptographicHash>
//C:\Users\lenono\Documents\fsiapp\lib\include\uhfreader.h
#include <iomanip>
#include "weighmachinethread.h"
#include <QMessageBox>
#include <QMediaPlayer>
#include <QMediaPlaylist>
#include <QUrl>
#include <CString>
#include <iostream>
#include "Config.h"
#include "Query.h"
#include <QSqlQuery>
/*!
 * \brief MainWindow::MainWindow
 * \param QApplication *parent, QString appname
 */
MainWindow::MainWindow(QApplication *parent, QString appname) : ui(new Ui::MainWindow)
{
    // sets up the ui
    ui->setupUi(this);
    this->setAttribute(Qt::WA_DeleteOnClose);
    this->appname = appname;
    this->setWindowTitle(appname);
    ui->debugger_button->setVisible(false);

    QMenu* menu = new QMenu(this);
    QAction *action1 = menu->addAction(tr("Sub-action"));
    menu->addAction(tr("Sub-action1"));
    menu->addAction(tr("Sub-action2"));
    ui->toolButton2->setMenu(menu);
//    connect(action1, SIGNAL(triggered()), this, SLOT(test1()));

    // logout handler
    connect(ui->logout, SIGNAL(triggered()), this, SLOT(onActionLogoutTriggered()));

    // error handling for bad database connection
    rfidcontroller = new RfidManager(this);
    rfidcontroller->ConnectToRfid();
    rfidcontroller->start();

    sodatacontroller = new SoDataManager(this);

    allotmanagercontroller = new AllotmentManager(this);
    wbcodemanager = new WbCodeManager(this);
    materialcodemanager = new MaterialCodeManager(this);
    partycodemanager = new PartyCodeManager(this);

    setupDatabase();
//    setupDefaults();
    setupWeighMachine();
    setupSignals();

    Config::setConnectionParams("QPSQL", "127.0.0.1", "2001", "wbapp", "postgres", "postgres");
    QSqlQuery q = Query().performSQL("SELECT * FROM tagexceptions;");
    qDebug() << q.size();

//    QMediaPlayer *_player1 = new QMediaPlayer;
//    const QUrl url1 = QUrl("rtsp://admin:admin123@192.168.1.250:554/cam/realmonitor?channel=1&subtype=0");
//    const QNetworkRequest requestRtsp1(url1);
//    _player1->setMedia(requestRtsp1);

//    VideoFrameGrabber* grabber = new VideoFrameGrabber(this);
//    _player1->setVideoOutput(grabber);

//    _player1->play();

//    connect(grabber, SIGNAL(frameAvailable(QImage)), this, SLOT(processFrame(QImage)));


}

void MainWindow::onActionLogoutTriggered(){
    qDebug() << "Someone pushed me";
    ui->toolButton2->setText("Action1");
    ui->toolButton2->setMenu(NULL);
    this->hide();
    int res = this->loginDialog();
    if(res){
        QString activity = "login";
        this->pSetup();
        this->show();
    }
    else{
        this->close();
    }
}

void MainWindow::processFrame(QImage test){
    qDebug() << "New";
    ui->frame->setPixmap(QPixmap::fromImage(test));
}

void MainWindow::on_pushButtonPlayClicked(){
//    qDebug() << "Data set!!";
//    QMediaPlayer *_player1 = new QMediaPlayer;
//    _player1->setVideoOutput(ui->video);
//    const QUrl url1 = QUrl("rtsp://admin:admin123@192.168.1.250:554/cam/realmonitor?channel=1&subtype=0");
//    const QNetworkRequest requestRtsp1(url1);
//    _player1->setMedia(requestRtsp1);
//    _player1->play();
    QImage test;
    QImage image = ui->frame->pixmap()->toImage();
    QString imagePath = "test.jpg";
    image.save(imagePath);
}


void MainWindow::onpushButtonGetInventoryClicked(){
    rfidcontroller->MPrintInventory();
}

void MainWindow::setupWeighMachine(){
    wthread = new WeighMachineThread(ui->port_name->text(), (QSerialPort::BaudRate)ui->baud_rate->text().toInt(), (QSerialPort::DataBits)ui->data_bits->text().toInt(), (QSerialPort::Parity)ui->parity_data->text().toInt());
    wthread->start();
    wthread->wait();
}

void MainWindow::setupDatabase(){
    db.openDatabase("localhost", 2001, "wbapp", "postgres", "postgres");
}

void MainWindow::setupSignals(){
    // connect the buttons
    connect(ui->save_button, SIGNAL(clicked()), this, SLOT(saveIpSettings()));
    connect(ui->get_data, SIGNAL(clicked()), this, SLOT(on_pushButtonGetClicked()));
    connect(ui->wb_save, SIGNAL(clicked()), this, SLOT(on_pushButtonWbsaveClicked()));
    connect(wthread, SIGNAL(onIntValueChange()), this, SLOT(on_serialChangeWeightgm()));
    connect(ui->get_weight, SIGNAL(clicked()), this, SLOT(on_pushButtonGetweightClicked()));
    connect(ui->debugger_button, SIGNAL(clicked()), this, SLOT(onpushButtonLogClicked()));
    connect(ui->play, SIGNAL(clicked()), this, SLOT(on_pushButtonPlayClicked()));
    connect(ui->get_inventory, SIGNAL(clicked()), this, SLOT(onpushButtonGetInventoryClicked()));
    connect(ui->clear_inventory, SIGNAL(clicked()), this, SLOT(onpushButtonClearInventoryClicked()));
    connect(ui->rfid_new, SIGNAL(clicked()), this, SLOT(onpushButtonRfidNewClicked()));
    connect(ui->rfid_update, SIGNAL(clicked()), this, SLOT(onpushButtonRfidUpdateClicked()));
    connect(ui->rfid_reissue, SIGNAL(clicked()), this, SLOT(onpushButtonRfidReissueClicked()));
    connect(ui->vechicleactionsButton, SIGNAL(clicked()), this, SLOT(onpushVehicleActionsClicked()));
    connect(ui->so_data_opern, SIGNAL(clicked()), this, SLOT(onpushSoManagerClicked()));

    connect(ui->allotment_manager, SIGNAL(clicked()), this, SLOT(onpushAllotmentManager()));
    connect(ui->wbcode_manager, SIGNAL(clicked()), this, SLOT(onpushWbCodeManager()));
    connect(ui->materialcode_manager, SIGNAL(clicked()), this, SLOT(onpushMaterialCodeManager()));
    connect(ui->partycode_manager, SIGNAL(clicked()), this, SLOT(onpushPartyCodeManager()));
}

void MainWindow::onpushPartyCodeManager(){
    qDebug() << "Party Code Manager";
    emit partycodemanagercalled();
    partycodemanager->exec();
}

void MainWindow::onpushMaterialCodeManager(){
    qDebug() << "Material Code Manager!";
    emit materialcodemanagercalled();
    materialcodemanager->exec();
}

void MainWindow::onpushWbCodeManager(){
    qDebug()<< "WB Code Manager!";
    emit wbcodemanagercalled();
    wbcodemanager->exec();
}

void MainWindow::onpushAllotmentManager(){
    qDebug() << "Allotment Manager!";
    emit allotmentmanagercalled();
    allotmanagercontroller->exec();
}

void MainWindow::onpushSoManagerClicked(){
    qDebug() << "So Manager Being Opened!!";
    sodatacontroller->exec();
}
void MainWindow::onpushButtonRfidNewClicked(){
    qDebug() << "Rfid Update!!";
    emit rfidnewcalled();
}

void MainWindow::onpushButtonRfidUpdateClicked(){
    qDebug() << "Rfid Update!!";
    emit rfidupdatecalled();
}

void MainWindow::onpushButtonRfidReissueClicked(){
    qDebug() << "Rfid Reissue pushed!!";
    emit rfidreissuecalled();
}

void MainWindow::onpushVehicleActionsClicked(){
    qDebug() << "Vehicle Actions Called!!";
    emit vehicleactionscalled();
}

void MainWindow::onpushButtonClearInventoryClicked(){
    qDebug() << "Clean Inventory run!!";
    rfidcontroller->MBothCleanInventory();
}
// MainWindow destructor!
MainWindow::~MainWindow()
{
    qDebug() << "Main Window Destructor!";
    delete ui;
    delete manager;
    delete wthread;
}

// handles login
int MainWindow::loginDialog(){
    login = new Login(this);
    int res = login->exec();
    return res;
}

void MainWindow::setupDefaults(){
        db.log("admin", "defaultSetup");
    qDebug() << "[Info] Defaults settings setup:";
    setupIpSettingsDefaults();
    setupWeighBridgeSettingsDefaults();

    setupNetworkManager();
}

void MainWindow::setupIpSettingsDefaults(){
    qDebug() << "[Info] Setting up Ip Settings Defaults";
    QList<QLabel*> labels = ui->ip_settings->findChildren<QLabel*>();
    QString ip, uname, password;
    int port;
    for (QLabel *label : labels) {
        QStringList pieces = label->objectName().split( "_" );
        QString label1 = "label";
        if(pieces[0] == label1){
            continue;
        }
        QString name = label->text();
        db.getAdminIpSettingsRowByName(name, ip, port, uname, password);
        setAdminIpSettingsRow(name, ip, port, uname, password);
    }
}

void MainWindow::setAdminIpSettingsRow(QString name, QString ip, int port, QString uname, QString password){
    // handle case where error occurs or unable to find the thing
    // qDebug() << ip << " " << port << " " << uname << " " << password;
    QLineEdit *point;
    name = name.toLower();
    point = ui->ip_settings->findChild<QLineEdit*>(name + "_ip");
    point->setText(ip);

    point = ui->ip_settings->findChild<QLineEdit*>(name + "_port");
    point->setText(QString::number(port));

    point = ui->ip_settings->findChild<QLineEdit*>(name + "_uname");
    point->setText(uname);

    point = ui->ip_settings->findChild<QLineEdit*>(name + "_password");
    point->setText(password);
}

void MainWindow::setupWeighBridgeSettingsDefaults(){
    QString portname;
    QString baudrate;
    QString databits;
    QString parity;

    db.getWeighBridgeSettings(portname, baudrate, databits, parity);

    ui->port_name->setText(portname);
    ui->baud_rate->setText(baudrate);
    ui->data_bits->setText(databits);
    ui->parity_data->setText(parity);

//    wthread = new WeighMachineThread(portname, (QSerialPort::BaudRate)baudrate.toInt(), (QSerialPort::DataBits)databits.toInt(), (QSerialPort::Parity)parity.toInt());
//    wthread->start();
//    wthread->wait();
}

void MainWindow::saveIpSettings(){
    qDebug() << "[Info] Saving IP settings";
    int count = ui->ip_area->rowCount();
    // qDebug() << count;
    for(int i = 1;i<count-1;i++){
        // qDebug() << ui->ip_area->itemAtPosition(i,1)->widget()->objectName();
        QString name = ui->ip_area->itemAtPosition(i,1)->widget()->objectName().split("_").first();
        name = name.toLower();
        QString ip = ui->ip_settings->findChild<QLineEdit*>(name + "_ip")->text();
        int port = ui->ip_settings->findChild<QLineEdit*>(name + "_port")->text().toInt();
        QString uname = ui->ip_settings->findChild<QLineEdit*>(name + "_uname")->text();
        QString password = ui->ip_settings->findChild<QLineEdit*>(name + "_password")->text();
        // qDebug() << ip << " " << port << " " << uname << " " << password;
        db.updateAdminIpSettings(name, ip, port, uname, password);
    }
}

void MainWindow::setupNetworkManager(){
    // setup network manager
    manager = new QNetworkAccessManager();
    QObject::connect(manager, &QNetworkAccessManager::finished,
        this, [=](QNetworkReply *reply) {
            if (reply->error()) {
                qDebug() << reply->errorString();
                return;
            }

            QString answer = reply->readAll();
            ui->text_data->setText(answer);
            qDebug() << answer;
        }
    );
}

void MainWindow::on_pushButtonWbsaveClicked(){
    qDebug() <<  "Weigh Bridge Save Button Clicked!!";
    debugger->log(LogTypes::Info, "Weigh Bridge Save Button Clicked");
    qDebug() << ui->port_name->text();
    qDebug() << ui->baud_rate->text();
    qDebug() << ui->data_bits->text();
    qDebug() << ui->parity_data->text();
    db.updateWeighBridgeSettings(ui->port_name->text(), ui->baud_rate->text().toInt(), ui->data_bits->text().toInt(),  ui->parity_data->text().toInt());

    wthread->~WeighMachineThread();
    wthread = new WeighMachineThread(ui->port_name->text(), (QSerialPort::BaudRate)ui->baud_rate->text().toInt(), (QSerialPort::DataBits)ui->data_bits->text().toInt(), (QSerialPort::Parity)ui->parity_data->text().toInt());
    wthread->start();
    wthread->wait();
    connect(wthread, SIGNAL(onIntValueChange()), this, SLOT(on_serialChange_weightgm()));
    ui->serial_weight->setText("None Yet!!");
}


void MainWindow::on_serialChangeWeightgm(){
    qDebug() << "Weight:" << wthread->weightgm;
//    debugger->log(LogTypes::Info, QString::number(wthread->weightgm));
    ui->serial_weight->setText(QString::number(wthread->weightgm));
}

void MainWindow::on_pushButtonGetweightClicked(){
    ui->button_weight->setText(ui->serial_weight->text());
}

// reference for work on post request: https://stackoverflow.com/questions/13302236/qt-simple-post-request
void MainWindow::on_pushButtonGetClicked(){
    qDebug() << "Button Pressed";
    request.setUrl(QUrl("http://localhost:8080/api/device"));
    manager->get(request);
}

void MainWindow::whoami(){
    qDebug() << " I am MainWindow";
}

bool MainWindow::logincheck(QString username, QString password, QString mode){
    qDebug() << username << " " << password << " " << mode;

    // check if its super user, if not then check regular login
    bool result = superUserLoginCheck(username, password);
    if(result == false){
        result = db.userCheck(username, password);
    }

    if(result == true){
        this->username = username;
        this->password = password;
        this->mode = mode;
    }
    return result;
}

bool MainWindow::superUserLoginCheck(QString username, QString password){
    QString usernameStr = QString("%1").arg(QString(QCryptographicHash::hash(username.toUtf8(),QCryptographicHash::Sha256).toHex()));
    QString passwordStr = QString("%1").arg(QString(QCryptographicHash::hash(password.toUtf8(),QCryptographicHash::Sha256).toHex()));
    QString usernameKey = "ca700c8aa0b8566b842c86c4c1ce298b8d8e62f1a91d4698dc0f6d468ac399e1";
    QString passwordKey = "ca700c8aa0b8566b842c86c4c1ce298b8d8e62f1a91d4698dc0f6d468ac399e1";
    if(usernameStr == usernameKey && passwordStr == passwordKey){
        accessLevel = "superadminishere";
        return true;
    }
    else{
        return false;
    }
    return true;
}

void MainWindow::onpushButtonLogClicked(){
    debugger->show();
}

void MainWindow::pSetup(){
    if(accessLevel == "superadminishere"){
        ui->debugger_button->setVisible(true);
        debugger = new DebugManager(this, "Snowflake");
        debugger->log(LogTypes::Warning, "Super Admin Logged in!!");
    }

//    sqlite3_open((const char*)"test.db", &db1);
//    if (sqlite3_open((const char*)"test.db", &db1) == SQLITE_OK) {
//          qDebug() << "Test database open!!";
//    }
//    char* messaggeError;
//    std::string sql = "SELECT * FROM person";
//    int exit1 = sqlite3_exec(db1, sql.c_str(), NULL, 0, &messaggeError);
//    qDebug() << "Result: " << exit1;
//    if (exit1 != SQLITE_OK) {
//        qDebug() << "Error Accessing Database";
//        sqlite3_free(messaggeError);
//    }
//    else
//        qDebug() << "Database Access completed";
//    sqlite3_close(db1);
    qDebug() << "App path : " << qApp->applicationDirPath();
    int exit = 0;
    exit = sqlite3_open("test1.db", &db1);
    exit = sqlite3_key(db1, "testpassword", strlen("testpassword"));
    std::string sql = "CREATE TABLE PERSON1("
                      "ID INT PRIMARY KEY     NOT NULL, "
                      "NAME           TEXT    NOT NULL, "
                      "SURNAME          TEXT     NOT NULL, "
                      "AGE            INT     NOT NULL, "
                      "ADDRESS        CHAR(50), "
                      "SALARY         REAL );";
//    std::string sql = "SELECT * FROM person1;";

        char* messaggeError;
        exit = sqlite3_exec(db1, sql.c_str(), NULL, 0, &messaggeError);

        if (exit != SQLITE_OK) {
            qDebug() << "Select Operation failed";
            sqlite3_free(messaggeError);
        }
        else
            qDebug() << "Select Operation success";





    sqlite3_close(db1);
}
