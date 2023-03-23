#include "wbcodemanager.h"
#include "Inserter.h"
#include "Selector.h"
#include "Updater.h"
#include "qdebug.h"
#include "ui_wbcodemanager.h"
#include "Query.h"

WbCodeManager::WbCodeManager(MainWindow *parent) : ui(new Ui::WbCodeManager)
{
    this->parent = parent;
    ui->setupUi(this);
    connect(ui->get_button, SIGNAL(clicked()), this, SLOT(onpushGet()));
    connect(ui->update_button, SIGNAL(clicked()), this, SLOT(onpushUpdate()));
    connect(ui->insert_button, SIGNAL(clicked()), this, SLOT(onpushInsert()));
}

WbCodeManager::~WbCodeManager()
{
    delete ui;
}

void WbCodeManager::onpushGet(){
    QString wbcode = ui->wbCodeLineEdit->text();
    auto res = Query("wbcode")
            .select()
            .where(OP::EQ("wbcode", wbcode))
            .perform();
    auto dict = res.first().toMap();
    this->ui->wbNameLineEdit->setText(dict["wbname"].toString());
}

void WbCodeManager::onpushInsert(){
    QString wbcode = ui->wbCodeLineEdit_2->text();
    QString wbname = ui->wbNameLineEdit_2->text();
    auto res = Query("wbcode")
            .insert({"wbcode", "wbname"})
            .values({wbcode, wbname})
            .perform();
}

void WbCodeManager::onpushUpdate(){
    QString wbcode = ui->wbCodeLineEdit->text();
    QString wbname = ui->wbNameLineEdit->text();
    bool ok = Query("wbcode")
            .update({{"wbcode", wbcode}, {"wbname", wbname}})
            .where(OP::LE("wbcode", wbcode))
            .perform();
    if(ok){
        qDebug()<<"Pass!";
    }
    else{
        qDebug()<<"Fail!";
    }
}
