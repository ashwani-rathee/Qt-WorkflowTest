#ifndef RFIDMANAGER_H
#define RFIDMANAGER_H

#include "qobject.h"
#include "qthread.h"
#include "rfidupdate.h"
#include "rfidnew.h"
#include "vehicleactions.h"
#include "src/modules/RfidManager/vehicleactions.h"
#include <QString>
#include <QVector>

struct Tag {
  int counter;
  int freqant;
  int rssi;
  QString pc;
  int epclen;
  QString epcdata;

  Tag(int counter, int freqant, int rssi, QString pc, int epclen, QString epcdata){
    this->counter = counter;
    this->freqant = freqant;
    this->rssi = rssi;
    this->pc = pc;
    this->epclen = epclen;
    this->epcdata = epcdata;
  }
};
class MainWindow;
class RfidManager: public QThread
{
    Q_OBJECT
    friend class RfidReissue;
    friend class RfidNew;
    friend class RfidUpdate;
    friend class VehicleActions;
public:
    RfidManager(MainWindow *parent);
    ~RfidManager();
    int ConnectToRfid();
    int MRFIDCleanInventory();
    int MRemoveTags();
    int MAddTag(int counter, int freqant, int rssi, QString pc, int epclen, QString epcdata);
    int MAddTagExistAware(int counter, int freqant, int rssi, QString pc, int epclen, QString epcdata);
    int MGetInventory();
    int MProcessInventory();
    void MPrintInventory();
    int MRFIDControllerCleanInventory();
    int MBothCleanInventory();
    void run();
    void read_data();
    QVector<struct Tag> tags;

private slots:
    void onpushbuttonRfidUpdateClicked();
    void onpushbuttonRfidNewClicked();
    void onpushbuttonRfidReissueClicked();
    void onpushbuttonVehicleActionsClicked();


private:
    MainWindow *main;
    RfidUpdate *updateform;
    RfidNew *newform;
    RfidReissue *reissueform;
    VehicleActions *vehmanager;

};

#endif // RFIDMANAGER_H
