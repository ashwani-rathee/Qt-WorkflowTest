#include "rfidnew.h"
#include "qstatusbar.h"
#include "ui_rfidnew.h"
#include "rfidmanager.h"
#include <QDebug>
#include "mainwindow.h"

RfidNew::RfidNew(RfidManager *parent) : ui(new Ui::RfidNew)
{
    this->main = parent;
    ui->setupUi(this);
    bar = new QStatusBar(this);
    ui->statusbar->addWidget(bar);
    ui->weight_type_dispatch->setChecked(true);

    ui->tagTripsLineEdit_2->setValidator(new QIntValidator(1, 9999, this) );


    connect(ui->read_inventory, SIGNAL(clicked()), this, SLOT(onpushButtonReadTagPushed()));
    connect(ui->clear_form, SIGNAL(clicked()), this, SLOT(ClearLineEdits()));
    connect(ui->save_form, SIGNAL(clicked()), this, SLOT(SaveDataForm()));
    connect(ui->weight_type, SIGNAL(buttonClicked(int)), this, SLOT(onButtonClickedChangeWeighmentPage(int)));
}

void RfidNew::onButtonClickedChangeWeighmentPage(int i){
    qDebug() << "Pushed and idx: " << i;
    qDebug() << ui->weight_type->checkedId();
    ui->weightment_forms->setCurrentIndex(-i-2);
}

void RfidNew::onpushButtonReadTagPushed(){
    this->main->read_data();
    //
    if(this->main->tags.length() > 0){
        ui->epc_number->setText(this->main->tags.at(0).epcdata);
        if(this->main->tags.length() > 1){
            bar->showMessage("More than 1 RFID tag has been detected!");
        }
        else{
            bar->showMessage("1 RFID tag has been detected!");
        }
    }
    else{
        ui->epc_number->setText("");
        bar->showMessage("No RFID tag found, click Read Tag again!");
    }
}


RfidNew::~RfidNew()
{
    delete ui;
}

void RfidNew::ClearLineEdits(){
    foreach(QLineEdit* le, ui->rfidnew_main->findChildren<QLineEdit*>()) {
        le->clear();
    }
}

void RfidNew::SaveDataForm(){

    QString TAGNO = ui->epc_number->text();
//    QString TAGNO = ui->tagTripsLineEdit_2->text();
    if(TAGNO == ""){
        return;
    }
    qDebug() << "Saving Data form";
    QString V_NO = ui->vehicleNoLineEdit_2->text();
    QDate ISSUE = QDate::currentDate();
    QDate EXPIRY = ui->ExpiryDateEdit_2->date();
    QString TC_CODE = "2000-01-01";
//    QDate TC_CODE = ui->ExpiryDateEdit_2->date();
    QString OWNER = ui->truckOwnerLineEdit_2->text();
    QString OWNER_ADDRESS = ui->ownerAddressLineEdit_2->text();
    QString OWNER_EMAIL = ui->ownerEmailLineEdit_2->text();
    QString OWNER_PHONE = ui->ownerPhoneLineEdit_2->text();
    QString DRIVER = ui->truckDriverLineEdit_2->text();
    QString DRIVER_ADDRESS = ui->driverAddressLineEdit_2->text();
    QString DRIVER_EMAIL = ui->driverEmailLineEdit->text();
    QString DRIVER_PHONE =  ui->driverPhoneLineEdit_2->text();
    QString PHOTO =  "asdas";
    QString TM_CODE = "test";
    qint64 RLW = ui->rLWInKgLineEdit_2->text().toInt();
    qint64 GVW = 1000;
    qint64 MAX_TARE_WT = 1000;
    qint64 MIN_TARE_WT = 1000;

    this->main->main->db.addNewRfid(TAGNO, V_NO, ISSUE, EXPIRY, TC_CODE, OWNER, OWNER_ADDRESS, OWNER_EMAIL, OWNER_PHONE, DRIVER, DRIVER_ADDRESS, DRIVER_EMAIL, DRIVER_PHONE, PHOTO, TM_CODE, RLW, GVW, MAX_TARE_WT, MIN_TARE_WT);
}
