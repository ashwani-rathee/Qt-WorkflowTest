#include "vehicleactions.h"
#include "src/modules/RfidManager/rfidreissue.h"
#include "ui_vehicleactions.h"
#include "rfidmanager.h"
#include "mainwindow.h"
#include "rfidnew.h"
#include "ui_rfidnew.h"
#include "rfidupdate.h"
#include "ui_rfidupdate.h"
#include "ui_rfidreissue.h"

VehicleActions::VehicleActions(RfidManager *ancestor) :
    ui(new Ui::VehicleActions)
{
    this->parent = ancestor;
    ui->setupUi(this);
    ui->issuenewRfid->setVisible(false);
    ui->changeRfid->setVisible(false);
    ui->resissueRfid->setVisible(false);
    connect(ui->get_actions, SIGNAL(clicked()), this, SLOT(onpushGetActionsClicked()));
    connect(ui->issuenewRfid, SIGNAL(clicked()), this, SLOT(onpushIssueNewClicked()));
    connect(ui->changeRfid, SIGNAL(clicked()), this, SLOT(onpushChangeRfidClicked()));
    connect(ui->resissueRfid, SIGNAL(clicked()), this, SLOT(onpushReissueRfidClicked()));

}

void VehicleActions::onpushIssueNewClicked(){
    qDebug() << "Issue New Clicked()!!";
    parent->newform->ClearLineEdits();
    parent->newform->ui->vehicleNoLineEdit_2->setText(ui->vehicleNumberLineEdit->text());
    parent->newform->exec();
}

void VehicleActions::onpushChangeRfidClicked(){
    qDebug() <<  "Change Rfid Clicked()";
    parent->updateform->ClearLineEdits();
    parent->updateform->ui->vehicle_number->setText(ui->vehicleNumberLineEdit->text());
    parent->updateform->exec();
}

void VehicleActions::onpushReissueRfidClicked(){
    qDebug()<<" Reissue Rfid Clicked()";
    parent->reissueform->ClearLineEdits();
    parent->reissueform->ui->vehicle_number->setText(ui->vehicleNumberLineEdit->text());
    parent->reissueform->ui->driverAddressLineEdit->setText(row->DRIVER_ADDRESS);
    parent->reissueform->ui->driverEmailLineEdit->setText(row->DRIVER_EMAIL);
    parent->reissueform->ui->driverNameLineEdit->setText(row->DRIVER);
    parent->reissueform->ui->driverPhoneLineEdit->setText(row->DRIVER_PHONE);
    parent->reissueform->ui->ownerNameLineEdit->setText(row->OWNER);
    parent->reissueform->ui->ownerAddressLineEdit->setText(row->OWNER_ADDRESS);
    parent->reissueform->ui->ownerEmailLineEdit->setText(row->OWNER_EMAIL);
    parent->reissueform->ui->ownerPhoneLineEdit->setText(row->OWNER_PHONE);
    parent->reissueform->ui->soNoLineEdit->setText(row->DO_NO);
    parent->reissueform->ui->tagTripsLineEdit->setText((row->TAGTRIPS));
    parent->reissueform->exec();
}

VehicleActions::~VehicleActions()
{
    delete ui;
}

void VehicleActions::onpushGetActionsClicked(){
    resetButtons();
    qDebug() << "Get button clicked!! WITH" << ui->vehicleNumberLineEdit->text();
    int a1 = 0;
    a1 = parent->main->db.getValidityStatus(a1, ui->vehicleNumberLineEdit->text());
    QString TAGNO ="";
    QString V_NO = ui->vehicleNumberLineEdit->text();
    QDate ISSUE = QDate::currentDate() ;
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
    qint64 RLW = 0;
    qint64 GVW = 0;
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
    this->parent->main->db.getTagDataByVehicleNo(TAGNO, V_NO, ISSUE, EXPIRY, TC_CODE, OWNER, OWNER_ADDRESS, OWNER_EMAIL, OWNER_PHONE, DRIVER, DRIVER_ADDRESS, DRIVER_EMAIL, DRIVER_PHONE, PHOTO, TM_CODE, RLW, GVW, MAX_TARE_WT, MIN_TARE_WT, DO_NO, COLL_CODE, WB_CODE, VALID, TAGTRIPS, TRIPS_DONE, V_TYPE, TAG_TYPE, TSNO, UNIT, MODE, WMODE, DEST,TRIPTIME);
    row = new TagRow(TAGNO, V_NO, ISSUE, EXPIRY, TC_CODE, OWNER, OWNER_ADDRESS, OWNER_EMAIL, OWNER_PHONE, DRIVER, DRIVER_ADDRESS, DRIVER_EMAIL, DRIVER_PHONE, PHOTO, TM_CODE, RLW, GVW, MAX_TARE_WT, MIN_TARE_WT, DO_NO, COLL_CODE, WB_CODE, VALID, TAGTRIPS, TRIPS_DONE, V_TYPE, TAG_TYPE, TSNO, UNIT, MODE, WMODE, DEST,TRIPTIME);
    a1 = row->VALID.toInt();
    if(a1 == -1){
        ui->issuenewRfid->setVisible(true);
    } else if(a1 == 0){
        ui->changeRfid->setVisible(true);
        ui->resissueRfid->setVisible(true);
    }
    else if(a1 > 0){
        ui->changeRfid->setVisible(true);
    }
}


void VehicleActions::reset(){
    ui->vehicleNumberLineEdit->setText("");
    ui->issuenewRfid->setVisible(false);
    ui->changeRfid->setVisible(false);
    ui->resissueRfid->setVisible(false);
}

void VehicleActions::resetButtons(){
    ui->issuenewRfid->setVisible(false);
    ui->changeRfid->setVisible(false);
    ui->resissueRfid->setVisible(false);
}
