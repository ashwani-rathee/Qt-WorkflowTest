#include "materialcodemanager.h"
#include "Inserter.h"
#include "Updater.h"
#include "ui_materialcodemanager.h"
#include "Query.h"
#include "Selector.h"
#include <QDebug>

MaterialCodeManager::MaterialCodeManager(MainWindow *parent) : ui(new Ui::MaterialCodeManager)
{
   this->parent = parent;
    ui->setupUi(this);

    connect(ui->get_button, SIGNAL(clicked()), this, SLOT(onpushGet()));
    connect(ui->update_button, SIGNAL(clicked()), this, SLOT(onpushUpdate()));
    connect(ui->insert_button, SIGNAL(clicked()), this, SLOT(onpushInsert()));

}

MaterialCodeManager::~MaterialCodeManager()
{
    delete ui;
}

void MaterialCodeManager::onpushGet(){
    QString mcode = ui->mCodeLineEdit->text();
    auto res = Query("mater")
            .select()
            .where(OP::EQ("m_code", mcode))
            .perform();
    auto dict = res.first().toMap();
    this->ui->mNameLineEdit->setText(dict["m_name"].toString());
}

void MaterialCodeManager::onpushInsert(){
    QString m_code = ui->mCodeLineEdit_2->text();
    QString m_name = ui->mNameLineEdit_2->text();
    qDebug() << m_code << " " << m_name;
    auto res = Query("mater")
            .insert({"m_code", "m_name"})
            .values({m_code, m_name})
            .perform();
}

void MaterialCodeManager::onpushUpdate(){
    QString m_code = ui->mCodeLineEdit->text();
    QString m_name = ui->mNameLineEdit->text();
    bool ok = Query("mater")
            .update({{"m_code", m_code}, {"m_name", m_name}})
            .where(OP::LE("m_code", m_code))
            .perform();
    if(ok){
        qDebug()<<"Pass!";
    }
    else{
        qDebug()<<"Fail!";
    }
}
