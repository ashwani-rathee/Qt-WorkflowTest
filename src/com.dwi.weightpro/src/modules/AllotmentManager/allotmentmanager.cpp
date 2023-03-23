#include "allotmentmanager.h"
#include "Selector.h"
#include "Updater.h"
#include "Inserter.h"
#include "Where.h"
#include "qdatetime.h"
#include "ui_allotmentmanager.h"
#include "Query.h"
#include <QDebug>

AllotmentManager::AllotmentManager(MainWindow *parent) : ui(new Ui::AllotmentManager)
{
    this->parent = parent;
    ui->setupUi(this);
    connect(ui->get_button, SIGNAL(clicked()), this, SLOT(onpushGet()));
    connect(ui->update_button, SIGNAL(clicked()), this, SLOT(onpushUpdate()));
    connect(ui->insert_button, SIGNAL(clicked()), this, SLOT(onpushInsert()));
}

void AllotmentManager::onpushGet(){

    QString test = ui->doNoLineEdit->text();
    QString test1 = ui->doNoLineEdit->text();
    auto res = Query("allotment")
            .select()
            .where(OP::EQ("do_no", test))
            .perform();

    auto dict = res.first().toMap();
    // qDebug() << res.count();

    this->ui->doNoLineEdit->setText(dict["do_no"].toString());
    this->ui->wDateLineEdit->setText(dict["w_date"].toString());
    this->ui->allotmentLineEdit->setText(dict["allotment"].toString());
}

void AllotmentManager::onpushUpdate(){
    QString do_no = ui->doNoLineEdit->text();
    QString w_date = ui->wDateLineEdit->text();
    QString allotment = ui->allotmentLineEdit->text();

    bool ok = Query("allotment")
            .update({{"do_no", do_no}, {"w_date", w_date}, {"allotment", allotment}})
            .where(OP::LE("do_no", do_no))
            .perform();
    if(ok){
        qDebug()<<"Pass!";
    }
    else{
        qDebug()<<"Fail!";
    }
}

void AllotmentManager::onpushInsert(){
    qDebug() << "Insert!";
    QString do_no = ui->doNoLineEdit_2->text();
    QString w_date = ui->wDateLineEdit_2->text();
    QString allotment = ui->allotmentLineEdit_2->text();
    // 2021-08-08T00:00:00.000
    QDateTime dateTime = QDateTime::fromString(w_date, "yyyy-MM-ddThh:mm:ss.zzzZ");

    auto ids = Query("allotment")
            .insert({"do_no", "w_date", "allotment"})
            .values({do_no, QDateTime::currentDateTime(), allotment})
            .perform();
}

AllotmentManager::~AllotmentManager()
{
    delete ui;
}
