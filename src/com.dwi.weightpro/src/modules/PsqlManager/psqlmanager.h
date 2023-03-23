#ifndef PSQLMANAGER_H
#define PSQLMANAGER_H
#include <QSqlDatabase>
#include <QDebug>
#include <QDate>

class PsqlManager{

public:
    PsqlManager();
    ~PsqlManager();

    bool isOpen() const;
    bool userCheck(QString username, QString password);
    void log(QString username, QString activity);
    void openDatabase(QString hostname = "localhost", int port = 2001, QString databaseName = "wbapp", QString userName = "postgres", QString password = "postgres");
    void closeDatabase();

    // to update settings for ip_settings by admin
    void updateAdminIpSettings(QString name, QString ip, int port, QString uname, QString password);
    void insertAdminIpSettings(QString name, QString ip, int port, QString uname, QString password);
    void getAdminIpSettingsRowByName(QString name, QString &ip, int &port,  QString &uname, QString &password);


    // to update and get settings of weigh bridge
    void updateWeighBridgeSettings(QString portname, int baudrate, int databits, int parity);
    void getWeighBridgeSettings(QString &portname, QString &baudrate, QString &databits, QString &parity);

    void addNewRfid(QString TAGNO="", QString V_NO="", QDate ISSUE=QDate::currentDate(), QDate EXPIRY=QDate::currentDate(),  QString TC_CODE="", QString OWNER="", QString OWNER_ADDRESS="", QString OWNER_EMAIL="", QString OWNER_PHONE="", QString DRIVER="", QString DRIVER_ADDRESS="", QString DRIVER_EMAIL="", QString DRIVER_PHONE="", QString PHOTO="", QString TM_CODE="", qint64 RLW=100, qint64 GVW=100, qint64 MAX_TARE_WT=100, qint64 MIN_TARE_WT=100, QString DO_NO="1", QString COLL_CODE="collc", QString WB_CODE="wbc", QString VALID="0", int TAGTRIPS=1, int TRIPS_DONE=1, QString V_TYPE="1", QString TAG_TYPE="tagtype", QString TSNO="tsno", QString UNIT="unit", QString MODE="mode", QString WMODE="wmode", QString DEST="dest", QString TRIPTIME="triptime");
    void getTagDataByVehicleNo(QString &TAGNO, QString V_NO, QDate &ISSUE, QDate &EXPIRY,  QString &TC_CODE, QString &OWNER, QString &OWNER_ADDRESS, QString &OWNER_EMAIL, QString &OWNER_PHONE, QString &DRIVER, QString &DRIVER_ADDRESS, QString &DRIVER_EMAIL, QString &DRIVER_PHONE, QString &PHOTO, QString &TM_CODE, qint64 &RLW, qint64 &GVW, qint64 &MAX_TARE_WT, qint64 &MIN_TARE_WT, QString &DO_NO, QString &COLL_CODE, QString &WB_CODE, QString &VALID, int &TAGTRIPS, int &TRIPS_DONE, QString &V_TYPE, QString &TAG_TYPE, QString &TSNO, QString &UNIT, QString &MODE, QString &WMODE, QString &DEST, QString &TRIPTIME);
    int getValidityStatus(int &a, QString vehicleno);
    void updateRfidTag(QString newtag, QString v_no);

    void getSoData(QString &auction_id, QString &customer_code, QString &customer_name, QString &so_no, QString &so_date, QString &so_grade, QString &so_coal_size, QString &so_qty, QString &valid_start_date, QString &valid_end_date, QString &location_id, QString &location_desc, QString &bal_qty, QString &state_code);
    void pushSoData(QString auction_id, QString customer_code, QString customer_name, QString so_no, QString so_date, QString so_grade, QString so_coal_size, QString so_qty, QString valid_start_date, QString valid_end_date, QString location_id, QString location_desc, QString bal_qty, QString state_code);
    void updateSoData(QString auction_id, QString customer_code, QString customer_name, QString so_no, QString so_date, QString so_grade, QString so_coal_size, QString so_qty, QString valid_start_date, QString valid_end_date, QString location_id, QString location_desc, QString bal_qty, QString state_code);


    void logExceptions(QString edatetime,QString  tagno,QString  v_no,QString  oldtagno,QString  optype,QString  description,QString  o_name,QString  area,QString  wb,QString  expected,QString  recorded,QString  comments,QString  sent);
private:
    QSqlDatabase database;
};

#endif // PSQLMANAGER_H
