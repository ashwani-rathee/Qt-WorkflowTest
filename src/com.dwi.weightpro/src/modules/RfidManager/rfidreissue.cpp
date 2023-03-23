#include "rfidreissue.h"
#include "ui_rfidreissue.h"
#include <QDebug>
#include "mainwindow.h"
#include "vehicleactions.h"

RfidReissue::RfidReissue(RfidManager *ancestor) : ui(new Ui::RfidReissue)
{
    this->parent = ancestor;
    ui->setupUi(this);
    connect(ui->clear_form, SIGNAL(clicked()), this, SLOT(ClearLineEdits()));
    connect(ui->save_form, SIGNAL(clicked()), this, SLOT(SaveDataForm()));
}

RfidReissue::~RfidReissue()
{
    delete ui;
}


void RfidReissue::ClearLineEdits(){
    qDebug() << "Cleaning Form!!";
    foreach(QLineEdit* le, ui->RfidReissueBox->findChildren<QLineEdit*>()) {
        le->clear();
    }
}

void RfidReissue::SaveDataForm(){
//    QString comment = ui->commentLineEdit->toPlainText();
//    QString newtag = ui->newTagNoLineEdit->text();
//    QString v_no = ui->vehicle_number->text();
//    QString oldtagno = ui->oldTagNoLineEdit->text();
//    QString wb = ui->wBCodeLineEdit_2->text();
//    // update in tags
//    this->main->main->db.updateRfidTag(newtag, v_no);
//    QDateTime current = QDateTime::currentDateTime();
//    oldtagno = "asd";
    TagRow *row1 = this->parent->vehmanager->row;
//    this->parent->main->db.updateResissueCaseForm(row1->OWNER, row1->OWNER_ADDRESS, row1->OWNER_EMAIL, row1->OWNER_PHONE, row1->DRIVER, row1->DRIVER_ADDRESS, row1->DRIVER_EMAIL, row1->DRIVER_EMAIL, row1->DO_NO, row1->TAGTRIPS);
//    this->main->main->db.logExceptions(current.toString(), newtag, v_no, oldtagno, "TAG_CHANGE", "TAG CHANGE", this->main->main->username, "area code not known", wb, "expected", "recorded", comment, "1");
    ClearLineEdits();
}
