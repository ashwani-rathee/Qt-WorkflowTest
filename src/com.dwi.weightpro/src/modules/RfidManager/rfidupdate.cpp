#include "rfidupdate.h"
#include "ui_rfidupdate.h"
#include "rfidmanager.h"
#include "mainwindow.h"
RfidUpdate::RfidUpdate(RfidManager *parent): ui(new Ui::RfidUpdate){
    this->main = parent;
    ui->setupUi(this);
//    this->setWindowTitle(parent->appname);
    connect(ui->read_inventory, SIGNAL(clicked()), this, SLOT(onpushButtonReadTagPushed()));
    connect(ui->clear_form, SIGNAL(clicked()), this, SLOT(ClearLineEdits()));
    connect(ui->save_form, SIGNAL(clicked()), this, SLOT(SaveDataForm()));
}


void RfidUpdate::onpushButtonReadTagPushed(){
    this->main->read_data();
    //
    if(this->main->tags.length() > 0){
        ui->new_tagid->setText(this->main->tags.at(0).epcdata);
        ui->newTagNoLineEdit->setText(this->main->tags.at(0).epcdata);
    }
    else{
        ui->new_tagid->setText("");
        ui->newTagNoLineEdit->setText("");
    }

    QString TAGNO ="";
    QString V_NO = ui->vehicle_number->text();
    QDate ISSUE = QDate::currentDate();
    QDate EXPIRY = QDate::currentDate();
    QString TC_CODE = "";

    QString OWNER = "";
    QString OWNER_ADDRESS = "";
    QString OWNER_EMAIL = "";
    QString OWNER_PHONE = "";
    QString VALID = "0";
    QString DRIVER = "";
    QString DRIVER_ADDRESS = "";
    QString DRIVER_EMAIL = "";
    QString DRIVER_PHONE =  "";
    QString PHOTO =  "";
    QString TM_CODE = "";
    qint64 RLW = 1;
    qint64 GVW = 1;
    qint64 MAX_TARE_WT = 1;
    qint64 MIN_TARE_WT = 1;
    QString DO_NO = "";
    QString COLL_CODE = "";
    QString WB_CODE = "";
    int TAGTRIPS = 1;
    int TRIPS_DONE = 1;
    QString V_TYPE = "";
    QString TAG_TYPE = "";
    QString TSNO = "";
    QString UNIT = "";
    QString MODE = "";
    QString WMODE = "";
    QString DEST = "";
    QString TRIPTIME = "";
    this->main->main->db.getTagDataByVehicleNo(TAGNO, V_NO, ISSUE, EXPIRY, TC_CODE, OWNER, OWNER_ADDRESS, OWNER_EMAIL, OWNER_PHONE, DRIVER, DRIVER_ADDRESS, DRIVER_EMAIL, DRIVER_PHONE, PHOTO, TM_CODE, RLW, GVW, MAX_TARE_WT, MIN_TARE_WT, DO_NO, COLL_CODE, WB_CODE, VALID, TAGTRIPS, TRIPS_DONE, V_TYPE, TAG_TYPE, TSNO, UNIT, MODE, WMODE, DEST,TRIPTIME);
    ui->oldTagNoLineEdit->setText(TAGNO);
    ui->tagTripsLineEdit_2->setText(QString::number(TAGTRIPS));
    ui->ExpiryDateEdit_2->setDate(EXPIRY);
    ui->vehicleTypeLineEdit_2->setText(V_TYPE);
    ui->rLWInKgLineEdit_2->setText(QString::number(RLW));
    ui->truckOwnerLineEdit_2->setText(OWNER);
    ui->ownerAddressLineEdit_2->setText(OWNER_ADDRESS);
    ui->ownerEmailLineEdit_2->setText(OWNER_EMAIL);
    ui->ownerPhoneLineEdit_2->setText(OWNER_PHONE);

    ui->truckDriverLineEdit_2->setText(DRIVER);
    ui->driverAddressLineEdit_2->setText(DRIVER_ADDRESS);
    ui->driverPhoneLineEdit_2->setText(DRIVER_EMAIL);
    ui->wBCodeLineEdit_2->setText(WB_CODE);
}

RfidUpdate::~RfidUpdate(){

}

void RfidUpdate::ClearLineEdits(){
    qDebug() << "Cleaning Form!!";
    foreach(QLineEdit* le, ui->RfidChange->findChildren<QLineEdit*>()) {
        le->clear();
    }
}

void RfidUpdate::SaveDataForm(){
    QString comment = ui->commentLineEdit->toPlainText();
    QString newtag = ui->newTagNoLineEdit->text();
    QString v_no = ui->vehicle_number->text();
    QString oldtagno = ui->oldTagNoLineEdit->text();
    QString wb = ui->wBCodeLineEdit_2->text();
    // update in tags
    this->main->main->db.updateRfidTag(newtag, v_no);
    QDateTime current = QDateTime::currentDateTime();
    oldtagno = "asd";

    this->main->main->db.logExceptions(current.toString(), newtag, v_no, oldtagno, "TAG_CHANGE", "TAG CHANGE", this->main->main->username, "area code not known", wb, "expected", "recorded", comment, "1");
//    ClearLineEdits();
}
