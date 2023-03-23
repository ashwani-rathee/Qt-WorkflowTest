#include "partycodemanager.h"
#include "ui_partycodemanager.h"

PartyCodeManager::PartyCodeManager(MainWindow *parent) : ui(new Ui::PartyCodeManager)
{
    this->parent = parent;
    ui->setupUi(this);

    connect(ui->get_button, SIGNAL(clicked()), this, SLOT(onpushGet()));
    connect(ui->update_button, SIGNAL(clicked()), this, SLOT(onpushUpdate()));
    connect(ui->insert_button, SIGNAL(clicked()), this, SLOT(onpushInsert()));

}

PartyCodeManager::~PartyCodeManager()
{
    delete ui;
}

void PartyCodeManager::onpushGet(){
//    QString
}

void PartyCodeManager::onpushInsert(){

}

void PartyCodeManager::onpushUpdate(){

}
