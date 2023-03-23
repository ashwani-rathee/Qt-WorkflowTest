#include <QString>
#include "DebugManager.h"
#include "qsqlerror.h"
#include "ui_DebugManager.h"
#include "qlistwidget.h"
#include "qscrollbar.h"
#include <QDebug>
#include <QDateTime>
#include "mainwindow.h"
#include <QMessageBox>

DebugManager::DebugManager(MainWindow *parent, QString name) : ui(new Ui::DebugManager){
    ui->setupUi(this);

    this->parent = parent;
    map.insert(LogTypes::Info, Qt::blue);
    map.insert(LogTypes::Warning, Qt::red);
    map.insert(LogTypes::Alert, Qt::yellow);

    mapStr.insert(LogTypes::Info, "INFO");
    mapStr.insert(LogTypes::Warning, "WARNING");
    mapStr.insert(LogTypes::Alert, "ALERT");
    connect(ui->logger_save , SIGNAL(clicked()), this, SLOT(onpushButtonSaveClicked()));
    connect(ui->test_connection_button, SIGNAL(clicked()), this, SLOT(onpushButtonTestConnectionClicked()));
}

DebugManager::~DebugManager(){

}


void DebugManager::log(LogTypes a, QString logstring){
    ui->debug_console->addItem(QDateTime::currentDateTime().toString() + " [" + mapStr[a] + "] " + QString::number(ui->debug_console->count()));
    qDebug() << map[a] << QDateTime::currentDateTime().toString();
    ui->debug_console->item(ui->debug_console->count()-1)->setForeground(map[a]);
    if(ui->debug_console->count() > MaxLines){
        QListWidgetItem *item = ui->debug_console->item(0);
        delete item;
    }
    ui->debug_console->verticalScrollBar()->setValue(ui->debug_console->verticalScrollBar()->maximum());
}

void DebugManager::onpushButtonSaveClicked(){
    MaxLines = ui->maxLinesLineEdit->text().toInt();
}

void DebugManager::onpushButtonTestConnectionClicked(){
    QString server = ui->serverNameLineEdit->text();
    int port = ui->portLineEdit->text().toInt();
    QString database = ui->databaseLineEdit->text();
    QString username = ui->usernameLineEdit->text();
    QString password = ui->passwordLineEdit->text();

    QSqlDatabase db = QSqlDatabase::addDatabase("QPSQL");
    db.setHostName(server);
    db.setPort(port);
    db.setDatabaseName(database);
    db.setUserName(username);
    db.setPassword(password);

    if (db.open()) {
        // if connection succesful!
        QMessageBox::information(0, QObject::tr("Connection Test"), "Database Connection Successful!!");

    }
    else{
        // in case any thing else happens
        QMessageBox::warning(0, QObject::tr("Database Error"), db.lastError().text());
    }
    db.close();
    db.removeDatabase("QPSQL");
}
