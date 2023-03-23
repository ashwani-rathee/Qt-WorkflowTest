#include "rfidmanager.h"
#include "src/modules/RfidManager/rfidreissue.h"
#include "uhfreader.h"
#include "mainwindow.h"
#include <QDebug>
#include <QDateTime>

RfidManager::RfidManager(MainWindow *parent)
{
    this->main = parent;
    updateform = new RfidUpdate(this);
    newform = new RfidNew(this);
    reissueform = new RfidReissue(this);
    vehmanager = new VehicleActions(this);
    connect(main, SIGNAL(rfidupdatecalled()), this, SLOT(onpushbuttonRfidUpdateClicked()));
    connect(main, SIGNAL(rfidnewcalled()), this, SLOT(onpushbuttonRfidNewClicked()));
    connect(main, SIGNAL(rfidreissuecalled()), this, SLOT(onpushbuttonRfidReissueClicked()));
    connect(main, SIGNAL(vehicleactionscalled()), this, SLOT(onpushbuttonVehicleActionsClicked()));

}

RfidManager::~RfidManager()
{
    qDebug() << "Deleted!!";
}

void RfidManager::onpushbuttonRfidUpdateClicked(){
    qDebug() << "Button Clicked()!!";
    updateform->exec();
}

void RfidManager::onpushbuttonRfidNewClicked(){
    qDebug() << "Button Clicked()!!";
    newform->exec();
}

void RfidManager::onpushbuttonRfidReissueClicked(){
    qDebug() << "Reissue Button Clicked()!!";
    newform->exec();
}

void RfidManager::onpushbuttonVehicleActionsClicked(){
    qDebug() << "Vehicle Actions clicked()!!";
    vehmanager->exec();
    vehmanager->reset();
}

int RfidManager::ConnectToRfid(){
    char localip[15] = "10.0.0.8";
    int localport = 27001;
    char readerip[15] = "10.0.0.10";
    int readerport = 27001;
    int a = Net_Connect(localip, localport, readerip, readerport);
    qDebug() << "RfidManager Created and Net_Connect result:" << a;
    return a;
}

int RfidManager::MRFIDCleanInventory(){
    qDebug() << "Hardware Inventory Clean!!";
    CleanInventory();
    return 0;
}

int RfidManager::MRFIDControllerCleanInventory(){
    qDebug() << "Software Inventory Clean!!";
    tags.clear();
    return 0;
}


int RfidManager::MBothCleanInventory(){
    qDebug() << "Hardware and Software Clean Both!!!!";
    CleanInventory();
    tags.clear();
    return 0;
}



int RfidManager::MAddTag(int counter, int freqant, int rssi, QString pc, int epclen, QString epcdata){
    struct Tag a(counter, freqant, rssi, pc, epclen, epcdata);
    tags.push_back(a);
    return 0;
}

int RfidManager::MAddTagExistAware(int counter, int freqant, int rssi, QString pc, int epclen, QString epcdata){
    auto it = std::find_if(tags.begin(), tags.end(), [&epcdata](Tag s) { return epcdata == s.epcdata; } );
    if (tags.end() != it){
        it->counter = counter;
        it->freqant = freqant;
        it->rssi = rssi;
        it->pc =  pc;
        it->epclen = epclen;
    }
    else{
        struct Tag a(counter, freqant, rssi, pc, epclen, epcdata);
        tags.push_back(a);
    }
    return 0;
}

QString getStringFromUnsignedChar(unsigned char *str)
{

    QString s;
    QString result = "";
    int rev = 300;

    // Print String in Reverse order....
    for ( int i = 0; i<rev; i++)
        {
           s = QString("%1").arg(str[i],0,16);

           if(s == "0"){
              s="00";
             }
         result.append(s);

         }
   return result;
}

QString getStr(unsigned char *str, int start, int end)
{

    QString s;
    QString result = "";

    // Print String in Reverse order....
    for ( int i = start; i<end; i++)
        {
           s = QString("%1").arg(str[i],0,16);

           if(s == "0"){
              s="00";
             }
         result.append(s);

         }
   return result;
}


int RfidManager::MGetInventory(){
    unsigned char repeat = 1;
    unsigned char outData[3000] = {0}; // initialize to all zeroes
    unsigned char tagNum[20] = {0}; // initialize to all zeroes
    QString time_format = "yyyy-MM-dd  HH:mm:ss";
    QDateTime a = QDateTime::currentDateTime();
    QString as = a.toString(time_format);
    qDebug() << as;

    int c = Inventory(repeat, outData, tagNum);

    QString str;
    int TagNum = tagNum[0];
    qDebug() << "TagNum = " << TagNum;
    if(TagNum >0){
       int a = MRFIDCleanInventory();
       qDebug() << "Clean Inventory:" << a;
       for(int j=0;j<=5;j++){
           MAddTagExistAware(outData[30*j+1], outData[30*j+2], outData[30*j+3], getStr(outData, 30*j+4,30*j+5), outData[30*j+6], getStr(outData, 30*j+7,30*j+7+outData[30*j+6]) );
       }
    }
    return 0;
}

int RfidManager::MProcessInventory(){
    return 0;
}

void RfidManager::MPrintInventory(){
    for(auto it: tags){
        qDebug() << "Tag: " << it.epcdata << ", Counter:" << it.counter << ", EPC Len:" << it.epclen;
    }
}


void RfidManager::run(){
//    unsigned char repeat = 1;
//    unsigned char outData[3000] = {0}; // initialize to all zeroes
//    unsigned char tagNum[20] = {0}; // initialize to all zeroes
//    QString time_format = "yyyy-MM-dd  HH:mm:ss";
//    QString as;
//    while(true){
//        as = QDateTime::currentDateTime().toString(time_format);
////        qDebug() << as;
//        Inventory(repeat, outData, tagNum);
//        int TagNum = tagNum[0];
////        qDebug() << "TagNum = " << TagNum;
//        if(TagNum >0){
//           for(int j=0;j<=5;j++){
//               // qDebug() << "Information on data packet: Counter:"<< outData[30*j+1] << ", FreqAnt:" <<outData[30*j+2] << " , RSSI:" << outData[30*j+3] << " , PC: " << outData[30*j+4] << outData[30*j+5] << " , EPC Len: " << outData[30*j+6] << getStr(outData, 30*j+7,30*j+7+outData[30*j+6]);
//               QString a = "";
//               if(getStr(outData, 30*j+7,30*j+7+outData[30*j+6]) != a){
//                    MAddTagExistAware(outData[30*j+1], outData[30*j+2], outData[30*j+3], getStr(outData, 30*j+4,30*j+5), outData[30*j+6], getStr(outData, 30*j+7,30*j+7+outData[30*j+6]) );
//               }
//           }
//        }
////        MPrintInventory();
//        int sleep_msec = 100;
//        QThread::msleep(sleep_msec );
//    }
}

void RfidManager::read_data(){
    unsigned char repeat = 1;
    unsigned char outData[3000] = {0}; // initialize to all zeroes
    unsigned char tagNum[20] = {0}; // initialize to all zeroes
    QString time_format = "yyyy-MM-dd  HH:mm:ss";
    QString as;
    MBothCleanInventory();
    as = QDateTime::currentDateTime().toString(time_format);
    qDebug() << as;
    Inventory(repeat, outData, tagNum);
    int TagNum = tagNum[0];
    qDebug() << "TagNum = " << TagNum;
    if(TagNum >0){
       for(int j=0;j<=5;j++){
           // qDebug() << "Information on data packet: Counter:"<< outData[30*j+1] << ", FreqAnt:" <<outData[30*j+2] << " , RSSI:" << outData[30*j+3] << " , PC: " << outData[30*j+4] << outData[30*j+5] << " , EPC Len: " << outData[30*j+6] << getStr(outData, 30*j+7,30*j+7+outData[30*j+6]);
           QString a = "";
           if(getStr(outData, 30*j+7,30*j+7+outData[30*j+6]) != a){
                MAddTagExistAware(outData[30*j+1], outData[30*j+2], outData[30*j+3], getStr(outData, 30*j+4,30*j+5), outData[30*j+6], getStr(outData, 30*j+7,30*j+7+outData[30*j+6]) );
           }
       }
    }
    MPrintInventory();
}
