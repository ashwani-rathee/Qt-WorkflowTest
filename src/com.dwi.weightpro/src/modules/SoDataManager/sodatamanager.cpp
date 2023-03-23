#include "sodatamanager.h"
#include "ui_sodatamanager.h"
#include "mainwindow.h"

SoDataManager::SoDataManager(MainWindow *ancestor) :
    ui(new Ui::SoDataManager)
{
    this->parent = ancestor;
    ui->setupUi(this);
    connect(ui->get_so_data, SIGNAL(clicked()), this, SLOT(onpushSoDataGetClicked()));
    connect(ui->save_so_data, SIGNAL(clicked()), this, SLOT(onpushSoDataPushClicked()));
    connect(ui->update_so_data, SIGNAL(clicked()), this, SLOT(onpushSoDataUpdateClicked()));
}

void SoDataManager::onpushSoDataGetClicked(){
    qDebug() << "Fetching data";
    QString auction_id ="";
    QString customer_code = "";
    QString customer_name = "";
    QString so_no = ui->so_no_find->text();
    QString so_date = "";
    QString so_grade = "";
    QString so_coal_size = "";
    QString so_qty = "";
    QString valid_start_date = "";
    QString valid_end_date = "";
    QString location_id = "";
    QString location_desc = "";
    QString bal_qty = "";
    QString state_code = "";
    parent->db.getSoData(auction_id, customer_code, customer_name, so_no, so_date, so_grade, so_coal_size, so_qty, valid_start_date, valid_end_date, location_id, location_desc, bal_qty, state_code);
    ui->auctionIdLineEdit->setText(auction_id);
    ui->customerCodeLineEdit->setText(customer_code);
    ui->customerNameLineEdit->setText(customer_name);
    ui->soNoLineEdit->setText(so_no);
    ui->soDateLineEdit->setText(so_date);
    ui->soGradeLineEdit->setText(so_grade);
    ui->soCoalSizeLineEdit->setText(so_coal_size);
    ui->soQuantityLineEdit->setText(so_qty);
    ui->validStartDateLineEdit->setText(valid_start_date);
    ui->validEndDateLineEdit->setText(valid_end_date);
    ui->locationIdLineEdit->setText(location_id);
    ui->locationDescLineEdit->setText(location_desc);
    ui->balQtyLineEdit->setText(bal_qty);
    ui->stateCodeLineEdit->setText(state_code);
}
void SoDataManager::onpushSoDataPushClicked(){
    qDebug() << "Save data";

    QString auction_id =ui->auctionIdLineEdit_2->text();
    QString customer_code = ui->customerCodeLineEdit_2->text();
    QString customer_name = ui->customerNameLineEdit_2->text();
    QString so_no = ui->soNoLineEdit_2->text();
    QString so_date = ui->soDateLineEdit_2->text();
    QString so_grade = ui->soGradeLineEdit_2->text();
    QString so_coal_size = ui->soCoalSizeLineEdit_2->text();
    QString so_qty = ui->soQuantityLineEdit_2->text();
    QString valid_start_date = ui->validStartDateLineEdit_2->text();
    QString valid_end_date =  ui->validEndDateLineEdit_2->text();
    QString location_id = ui->locationIdLineEdit_2->text();
    QString location_desc = ui->locationDescLineEdit_2->text();
    QString bal_qty = ui->balQtyLineEdit_2->text();
    QString state_code = ui->stateCodeLineEdit_2->text();
    parent->db.pushSoData(auction_id, customer_code, customer_name, so_no, so_date, so_grade, so_coal_size, so_qty, valid_start_date, valid_end_date, location_id, location_desc, bal_qty, state_code);

}

void SoDataManager::onpushSoDataUpdateClicked(){
    qDebug() << "Update SO date";

    QString auction_id =ui->auctionIdLineEdit->text();
    QString customer_code = ui->customerCodeLineEdit->text();
    QString customer_name = ui->customerNameLineEdit->text();
    QString so_no = ui->soNoLineEdit->text();
    QString so_date = ui->soDateLineEdit->text();
    QString so_grade = ui->soGradeLineEdit->text();
    QString so_coal_size = ui->soCoalSizeLineEdit->text();
    QString so_qty = ui->soQuantityLineEdit->text();
    QString valid_start_date = ui->validStartDateLineEdit->text();
    QString valid_end_date =  ui->validEndDateLineEdit->text();
    QString location_id = ui->locationIdLineEdit->text();
    QString location_desc = ui->locationDescLineEdit->text();
    QString bal_qty = ui->balQtyLineEdit->text();
    QString state_code = ui->stateCodeLineEdit->text();
    parent->db.updateSoData(auction_id, customer_code, customer_name, so_no, so_date, so_grade, so_coal_size, so_qty, valid_start_date, valid_end_date, location_id, location_desc, bal_qty, state_code);

}

SoDataManager::~SoDataManager()
{
    delete ui;
}
