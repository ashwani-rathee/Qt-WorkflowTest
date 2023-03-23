#ifndef VEHICLEACTIONS_H
#define VEHICLEACTIONS_H

#include "qdatetime.h"
#include <QDialog>

struct TagRow{
    QString TAGNO;
    QString V_NO;
    QDate ISSUE;
    QDate EXPIRY;
    QString TC_CODE;
    QString OWNER;
    QString OWNER_ADDRESS;
    QString OWNER_EMAIL;
    QString OWNER_PHONE;

    QString DRIVER;
    QString DRIVER_ADDRESS;
    QString DRIVER_EMAIL;
    QString DRIVER_PHONE;

    QString PHOTO;
    QString TM_CODE;
    qint64 RLW;
    qint64 GVW;
    qint64 MAX_TARE_WT;
    qint64 MIN_TARE_WT;

    QString DO_NO;
    QString COLL_CODE;
    QString WB_CODE;
    QString VALID;

    QString TAGTRIPS;
    QString TRIPS_DONE;
    QString V_TYPE;
    QString TAG_TYPE;
    QString TSNO;
    QString UNIT;
    QString MODE;
    QString WMODE;
    QString DEST;
    QString TRIPTIME;

    TagRow(QString TAGNO, QString V_NO, QDate ISSUE, QDate EXPIRY,  QString TC_CODE, QString OWNER, QString OWNER_ADDRESS, QString OWNER_EMAIL, QString OWNER_PHONE, QString DRIVER, QString DRIVER_ADDRESS, QString DRIVER_EMAIL, QString DRIVER_PHONE, QString PHOTO, QString TM_CODE, qint64 RLW, qint64 GVW, qint64 MAX_TARE_WT, qint64 MIN_TARE_WT, QString DO_NO, QString COLL_CODE, QString WB_CODE, QString VALID, int TAGTRIPS, int TRIPS_DONE, QString V_TYPE, QString TAG_TYPE, QString TSNO, QString UNIT, QString MODE, QString WMODE, QString DEST, QString TRIPTIME){
        this->TAGNO = TAGNO;
        this->V_NO = V_NO;
        this->ISSUE = ISSUE;
        this->EXPIRY = EXPIRY;
        this->TC_CODE = TC_CODE;
        this->OWNER = OWNER;
        this->OWNER_ADDRESS = OWNER_ADDRESS;
        this->OWNER_EMAIL = OWNER_EMAIL;
        this->OWNER_PHONE = OWNER_PHONE;

        this->DRIVER = DRIVER;
        this->DRIVER_ADDRESS = DRIVER_ADDRESS;
        this->DRIVER_EMAIL = DRIVER_EMAIL;
        this->DRIVER_PHONE = DRIVER_PHONE;

        this->PHOTO = PHOTO;
        this->TM_CODE = TM_CODE;
        this->RLW = RLW;
        this->GVW = GVW;
        this->MAX_TARE_WT = MAX_TARE_WT;
        this->MIN_TARE_WT = MIN_TARE_WT;

        this->DO_NO = DO_NO;
        this->COLL_CODE = COLL_CODE;
        this->WB_CODE = WB_CODE;
        this->VALID = VALID;

        this->TAGTRIPS = TAGTRIPS;
        this->TRIPS_DONE = TRIPS_DONE;
        this->V_TYPE = V_TYPE;
        this->TAG_TYPE = TAG_TYPE;
        this->TSNO = TSNO;
        this->UNIT = UNIT;
        this->MODE = MODE;
        this->WMODE = WMODE;
        this->DEST = DEST;
        this->TRIPTIME = TRIPTIME;
    }
};

namespace Ui {
class VehicleActions;
}

class RfidManager;
class VehicleActions : public QDialog
{
    Q_OBJECT
    friend class RfidReissue;
public:
    explicit VehicleActions(RfidManager *parent = nullptr);
    ~VehicleActions();
    void reset();
    void resetButtons();

private slots:
    void onpushGetActionsClicked();
    void onpushIssueNewClicked();
    void onpushChangeRfidClicked();
    void onpushReissueRfidClicked();

private:
    TagRow *row;
    RfidManager *parent;
    Ui::VehicleActions *ui;
};

#endif // VEHICLEACTIONS_H
