#include "psqlmanager.h"
#include <QSqlQuery>
#include <QSqlError>
#include <QSqlRecord>
#include <QDebug>
#include <QDate>

// used to create the user table
// CREATE TABLE users ( ID INT PRIMARY KEY     NOT NULL, USERNAME           CHAR(50)    NOT NULL, PASSWORD           CHAR(50)    NOT NULL );


// used to create the user_events table
//  CREATE TABLE user_events ( ID SERIAL INT PRIMARY KEY     NOT NULL, USERNAME CHAR(50) NOT NULL, SERVERTIME TIMESTAMP DEFAULT NOW(), EVENT CHAR(50) );

// used to change ip admin settings
// CREATE TABLE admin_ip_settings ( ID SERIAL PRIMARY KEY NOT NULL, NAME CHAR(50) NOT NULL, IP CHAR(50) NOT NULL, PORT INT NOT NULL, UNAME CHAR(50) NOT NULL, PASSWORD  CHAR(50) NOT NULL);


//wbapp=# \COPY users FROM 'C:\Users\lenono\Documents\data\users.csv' DELIMITER ',' CSV HEADER;
//COPY 1
//wbapp=# \COPY user_events FROM 'C:\Users\lenono\Documents\data\user_events.csv' DELIMITER ',' CSV HEADER;
//COPY 0
//wbapp=# \COPY admin_ip_settings FROM 'C:\Users\lenono\Documents\data\admin_ip_settings.csv' DELIMITER ',' CSV HEADER;
//COPY 8

//wbapp=# \COPY users TO 'C:\Users\lenono\Documents\data\test.csv' WITH (FORMAT CSV, HEADER);
//COPY 1
//wbapp=# \COPY user_events TO 'C:\Users\lenono\Documents\data\user_events.csv' WITH (FORMAT CSV, HEADER);
//COPY 0
//wbapp=# \COPY admin_ip_settings TO 'C:\Users\lenono\Documents\data\admin_ip_settings.csv' WITH (FORMAT CSV, HEADER);
//COPY 8

//
//CREATE TABLE tags(
//	SNO SERIAL PRIMARY KEY NOT NULL,
//	TAGNO varchar(30) NOT NULL,
//	V_NO varchar(15) NULL,
//	ISSUE date NULL,
//	EXPIRY date NULL,
//	TC_CODE varchar(20) NULL,
//	OWNER varchar(50) NULL,
//	OWNER_ADDRESS varchar(100) NULL,
//	OWNER_EMAIL varchar(50) NULL,
//	OWNER_PHONE varchar(20) NULL,
//	DRIVER varchar(50) NULL,
//	DRIVER_ADDRESS varchar(100) NULL,
//	DRIVER_EMAIL varchar(50) NULL,
//	DRIVER_PHONE varchar(20) NULL,
//	PHOTO varchar(200) NULL,
//	TM_CODE varchar(8) NULL,
//	RLW bigint NULL,
//	GVW bigint NULL,
//	MAX_TARE_WT bigint NULL,
//	MIN_TARE_WT bigint NULL,
//	DO_NO varchar(50) NULL,
//	COLL_CODE varchar(20) NULL,
//	WB_CODE varchar(20) NULL,
//	VALID varchar(5) NULL,
//	TAGTRIPS int NULL,
//	TRIPS_DONE int NULL,
//	V_TYPE varchar(15) NULL,
//	TAG_TYPE varchar(15) NULL,
//	TSNO varchar(50) NULL,
//	UNIT varchar(15) NULL,
//	MODE varchar(15) NULL,
//	WMODE varchar(15) NULL,
//	DEST varchar(15) NULL,
//	TRIPTIME varchar(15) NULL
//);
// constructor of the class
PsqlManager::PsqlManager(){
    this->database = QSqlDatabase::addDatabase("QPSQL");
}

// destructor
PsqlManager::~PsqlManager(){
    if (this->database.isOpen())
    {
        this->database.close();
    }
}

// check if the connection is open or not?
bool PsqlManager::isOpen() const{
    return this->database.isOpen();

}

//void PsqlManager::createDatabase(){

//}
//void PsqlManager::checkDatabaseExistence(){

//}
// connects to the database
void PsqlManager::openDatabase(QString hostname, int port, QString databaseName, QString userName, QString password){
    this->database.setHostName(hostname);
    this->database.setPort(port);
    this->database.setDatabaseName(databaseName);
    this->database.setUserName(userName);
    this->database.setPassword(password);
    if (!this->database.open())
    {
        qDebug() << "Error: connection with database fail";
    }
    else
    {
        qDebug() << "Database: connection ok";
    }
}

void PsqlManager::closeDatabase(){
    qDebug() << "Attempting to close database!!";
    this->database.close();
    if (this->database.isOpen())
    {
        qDebug() << "Error: close failed";
    }
    else
    {
        qDebug() << "Database: closed successfully";
    }
    return;
}

bool PsqlManager::userCheck(QString username, QString password){
    QSqlQuery query;
    qDebug() << "select * from users where USERNAME ='"+username+"' and PASSWORD = '"+password+"'";
    query.prepare("select * from users where USERNAME ='"+username+"' and PASSWORD = '"+password+"'");
    query.exec();

    while(query.next()){
            return true;
    }

    return false;
}

void PsqlManager::log(QString username, QString activity)
{
    QSqlQuery query(this->database);
    qDebug() << "Logging: " << activity;
    query.prepare("INSERT INTO user_events (username, event) VALUES (:user, :deets)");
    query.bindValue(":user", username);
    query.bindValue(":deets", activity);
    if (query.exec()) {
        qDebug() << "Query success1";
    } else {
        qDebug() << query.lastError().text();
        return;
    }
}

void PsqlManager::updateAdminIpSettings(QString name, QString ip, int port, QString uname, QString password){
    // check if query was succesful or not
    QSqlQuery query(this->database);
    query.prepare("UPDATE admin_ip_settings SET ip = :ip, port = :port, uname = :uname, password = :password WHERE name= :name;");
    query.bindValue(":ip", ip);
    query.bindValue(":port", port);
    query.bindValue(":uname", uname);
    query.bindValue(":password", password);
    query.bindValue(":name", name.toUpper());
    query.exec();
}

void PsqlManager::insertAdminIpSettings(QString name, QString ip, int port, QString uname, QString password){
    QSqlQuery query(this->database);
    //  INSERT INTO "admin_ip_settings" (name, ip, port, uname, password) VALUES ('nasdsd', 'Asd', 2, 'asduname', 'asdpassword');
    query.prepare("INSERT INTO admin_ip_settings (name, ip, port, uname, password) VALUES (:name, :ip, :port, :uname, :password);");
    query.bindValue(":name", name.toUpper());
    query.bindValue(":ip", ip);
    query.bindValue(":port", port);
    query.bindValue(":uname", uname);
    query.bindValue(":password", password);
    query.exec();
}

void PsqlManager::getAdminIpSettingsRowByName(QString name,  QString &ip, int &port, QString &uname, QString &password){
    // handle case where this query returns nothing
    QSqlQuery query(this->database);
    // SELECT ip, port, uname, password FROM admin_ip_settings WHERE name = 'RFID1';
    query.prepare("SELECT ip, port, uname, password FROM admin_ip_settings WHERE name = :name;");
    query.bindValue(":name", name);
    query.exec();
    //    ip = "ads";
    //    port = 10;
    //    uname = "sad";
    //    password = "asd";
    query.next();
    ip = query.value(0).toString().trimmed();
    port = query.value(1).toInt();
    uname = query.value(2).toString().trimmed();
    password = query.value(3).toString().trimmed();
}

// weigh bridge related
void PsqlManager::updateWeighBridgeSettings(QString portname, int baudrate, int databits, int parity){
    // qDebug() << "Saving Weigh Bridge Settings";
    QSqlQuery query(this->database);
    // INSERT INTO admin_serial_connections VALUES (1, 'Base', 'COM5', 2400, 8, 0);
    // UPDATE admin_serial_connections SET portname = 'COM5', baudrate = 2400, databits = 8, parity = 0 WHERE name= 'Base';
    query.prepare("UPDATE admin_serial_connections SET portname = :portname, baudrate = :baudrate, databits = :databits, parity = :parity WHERE name= 'Base';");
    query.bindValue(":portname", portname);
    query.bindValue(":baudrate", baudrate);
    query.bindValue(":databits", databits);
    query.bindValue(":parity", parity);
    query.exec();
}


void PsqlManager::getWeighBridgeSettings(QString &portname, QString &baudrate, QString &databits, QString &parity){
    // qDebug() << "Fetching Weigh Bridge Settings";
    QSqlQuery query(this->database);
    //SELECT portname, baudrate, databits, parity FROM admin_serial_connections WHERE name = 'Base';
    query.prepare("SELECT portname, baudrate, databits, parity FROM admin_serial_connections WHERE name = 'Base';");
    query.exec();
    query.next();
    portname = query.value(0).toString().trimmed();
    baudrate = query.value(1).toString().trimmed();
    databits = query.value(2).toString().trimmed();
    parity = query.value(3).toString().trimmed();
}

void PsqlManager::addNewRfid(QString TAGNO, QString V_NO, QDate ISSUE, QDate EXPIRY,  QString TC_CODE, QString OWNER, QString OWNER_ADDRESS, QString OWNER_EMAIL, QString OWNER_PHONE, QString DRIVER, QString DRIVER_ADDRESS, QString DRIVER_EMAIL, QString DRIVER_PHONE, QString PHOTO, QString TM_CODE, qint64 RLW, qint64 GVW, qint64 MAX_TARE_WT, qint64 MIN_TARE_WT, QString DO_NO, QString COLL_CODE, QString WB_CODE, QString VALID, int TAGTRIPS, int TRIPS_DONE, QString V_TYPE, QString TAG_TYPE, QString TSNO, QString UNIT, QString MODE, QString WMODE, QString DEST, QString TRIPTIME){
    QSqlQuery query(this->database);
    query.prepare("INSERT INTO tags (tagno, v_no, issue, expiry, tc_code, owner, owner_address, owner_email, owner_phone, driver, driver_address, driver_email, driver_phone, photo, tm_code, rlw, gvw, max_tare_wt, min_tare_wt, do_no, coll_code, wb_code, valid, tagtrips, trips_done, v_type, tag_type, tsno, unit, mode, wmode, dest, triptime) VALUES (:TAGNO, :V_NO, :ISSUE, :EXPIRY, :TC_CODE, :OWNER, :OWNER_ADDRESS, :OWNER_EMAIL, :OWNER_PHONE, :DRIVER, :DRIVER_ADDRESS, :DRIVER_EMAIL, :DRIVER_PHONE, :PHOTO, :TM_CODE, :RLW, :GVW, :MAX_TARE_WT, :MIN_TARE_WT, :DO_NO, :COLL_CODE, :WB_CODE, :VALID, :TAGTRIPS, :TRIPS_DONE, :V_TYPE, :TAG_TYPE, :TSNO, :UNIT, :MODE, :WMODE, :DEST, :TRIPTIME)");
    query.bindValue(":TAGNO", TAGNO);
    query.bindValue(":V_NO", V_NO);
    query.bindValue(":ISSUE", ISSUE);
    query.bindValue(":EXPIRY", EXPIRY);
    query.bindValue(":TC_CODE", TC_CODE);

    query.bindValue(":OWNER", OWNER);
    query.bindValue(":OWNER_ADDRESS", OWNER_ADDRESS);
    query.bindValue(":OWNER_EMAIL", OWNER_EMAIL);
    query.bindValue(":OWNER_PHONE", OWNER_PHONE);

    query.bindValue(":DRIVER", DRIVER);
    query.bindValue(":DRIVER_ADDRESS", DRIVER_ADDRESS);
    query.bindValue(":DRIVER_EMAIL", DRIVER_EMAIL);
    query.bindValue(":DRIVER_PHONE", DRIVER_PHONE);

    query.bindValue(":PHOTO", PHOTO);
    query.bindValue(":TM_CODE", TM_CODE);
    query.bindValue(":RLW", RLW);
    query.bindValue(":GVW", GVW);
    query.bindValue(":MAX_TARE_WT", MAX_TARE_WT);
    query.bindValue(":MIN_TARE_WT", MIN_TARE_WT);
    query.bindValue(":DO_NO", DO_NO);
    query.bindValue(":COLL_CODE", COLL_CODE);
    query.bindValue(":WB_CODE", WB_CODE);
    query.bindValue(":VALID", VALID);
    query.bindValue(":TAGTRIPS", TAGTRIPS);
    query.bindValue(":TRIPS_DONE", TRIPS_DONE);
    query.bindValue(":V_TYPE", V_TYPE);
    query.bindValue(":TAG_TYPE", TAG_TYPE);
    query.bindValue(":TSNO", TSNO);
    query.bindValue(":UNIT", UNIT);
    query.bindValue(":MODE", MODE);
    query.bindValue(":WMODE", WMODE);
    query.bindValue(":DEST", DEST);
    query.bindValue(":TRIPTIME", TRIPTIME);

    log( "test", "NEW RFID Tag Issued!");
    if (query.exec()) {
        qDebug() << "Query success";
    } else {
        qDebug() << query.lastError().text();
        return;
    }
}

 void PsqlManager::getTagDataByVehicleNo(QString &TAGNO, QString V_NO, QDate &ISSUE, QDate &EXPIRY,  QString &TC_CODE, QString &OWNER, QString &OWNER_ADDRESS, QString &OWNER_EMAIL, QString &OWNER_PHONE, QString &DRIVER, QString &DRIVER_ADDRESS, QString &DRIVER_EMAIL, QString &DRIVER_PHONE, QString &PHOTO, QString &TM_CODE, qint64 &RLW, qint64 &GVW, qint64 &MAX_TARE_WT, qint64 &MIN_TARE_WT, QString &DO_NO, QString &COLL_CODE, QString &WB_CODE, QString &VALID, int &TAGTRIPS, int &TRIPS_DONE, QString &V_TYPE, QString &TAG_TYPE, QString &TSNO, QString &UNIT, QString &MODE, QString &WMODE, QString &DEST, QString &TRIPTIME){
    QSqlQuery query(this->database);
    qDebug() << V_NO;
//    query.prepare("SELECT * FROM tags WHERE v_no=':V_N0';");
    query.prepare("select * from tags where v_no ='"+V_NO+"'");
//    query.prepare("SELECT * FROM tags;");
    query.bindValue(":V_NO", V_NO);

//    qDebug() << query.;
    log( "test", "TagData fetched by Vehicle No.");
    if (query.exec()) {
        qDebug() << "Query Successs: TagData fetched by Vehicle No.";
        qDebug() << query.lastError().text();
    } else {
        qDebug() << query.lastError().text();
        return;
    }
    query.first();
    qDebug() << query.value(0).toString();
    TAGNO = query.value(1).toString();
    V_NO = query.value(2).toString();
    ISSUE = query.value(3).toDate();
    EXPIRY = query.value(4).toDate();
    TC_CODE = query.value(5).toString();

    OWNER = query.value(6).toString();
    OWNER_ADDRESS =  query.value(7).toString();
    OWNER_EMAIL = query.value(8).toString();
    OWNER_PHONE = query.value(9).toString();

    DRIVER = query.value(10).toString();
    DRIVER_ADDRESS = query.value(11).toString();
    DRIVER_EMAIL = query.value(12).toString();
    DRIVER_PHONE = query.value(13).toString();

    PHOTO = query.value(14).toString();
    TM_CODE = query.value(15).toString();
    RLW = query.value(16).toInt();
    GVW = query.value(17).toInt();
    MAX_TARE_WT = query.value(18).toInt();
    MIN_TARE_WT = query.value(19).toInt();

    DO_NO = query.value(20).toString();
    COLL_CODE = query.value(21).toString();
    WB_CODE = query.value(22).toString();
    VALID = query.value(23).toString();

    TAGTRIPS = query.value(24).toInt();
    TRIPS_DONE = query.value(25).toInt();
    V_TYPE = query.value(26).toString();
    TAG_TYPE = query.value(27).toString();
    TSNO = query.value(28).toString();
    UNIT = query.value(29).toString();
    MODE = query.value(30).toString();
    WMODE = query.value(31).toString();;
    DEST = query.value(32).toString();
    TRIPTIME = query.value(33).toString();
}

int PsqlManager::getValidityStatus(int &a, QString vehicleno){

    QSqlQuery query(this->database);
    qDebug() << "Veh:" << vehicleno;
    query.prepare("select valid from tags where v_no ='"+vehicleno+"'");
    query.exec();

    while(query.next()){
        int validity = query.value(0).toInt();
        qDebug() << "Validity Status:"<< validity;
        if(validity > 0 ){
            return 1;
        }
        else{
            return 0;
        }
    }

    return -1;
}

void PsqlManager::getSoData(QString &auction_id, QString &customer_code, QString &customer_name, QString &so_no, QString &so_date, QString &so_grade, QString &so_coal_size, QString &so_qty, QString &valid_start_date, QString &valid_end_date, QString &location_id, QString &location_desc, QString &bal_qty, QString &state_code){
    QSqlQuery query(this->database);
    query.prepare("SELECT * FROM sodata WHERE so_no=:so_no;");
    query.bindValue(":so_no", so_no);
    if (query.exec()) {
        qDebug() << "Query success of SO Data Fetch";
    } else {
        qDebug() << query.lastError().text();
        return;
    }
    query.next();
    auction_id = query.value(0).toString();
    customer_code = query.value(1).toString();
    customer_name = query.value(2).toString();
    so_no = query.value(3).toString();
    so_date = query.value(4).toString();
    so_grade = query.value(5).toString();
    so_coal_size = query.value(6).toString();
    so_qty = query.value(7).toString();
    valid_start_date = query.value(8).toString();
    valid_end_date = query.value(9).toString();
    location_id = query.value(10).toString();
    location_desc = query.value(11).toString();
    bal_qty = query.value(12).toString();
    state_code = query.value(13).toString();
    log( "test", "SO Data Fetched");

}

void PsqlManager::pushSoData(QString auction_id, QString customer_code, QString customer_name, QString so_no, QString so_date, QString so_grade, QString so_coal_size, QString so_qty, QString valid_start_date, QString valid_end_date, QString location_id, QString location_desc, QString bal_qty, QString state_code){
    QSqlQuery query(this->database);
    query.prepare("INSERT INTO sodata VALUES(:auction_id, :customer_code, :customer_name, :so_no, :so_date, :so_grade, :so_coal_size, :so_qty, :valid_start_date, :valid_end_date, :location_id, :location_desc, :bal_qty, :state_code);");
    query.bindValue(":auction_id", auction_id);
    query.bindValue(":customer_code", customer_code);
    query.bindValue(":customer_name", customer_name);
    query.bindValue(":so_no", so_no);
    query.bindValue(":so_date", so_date);
    query.bindValue(":so_grade", so_grade);
    query.bindValue(":so_coal_size", so_coal_size);
    query.bindValue(":so_qty", so_qty);
    query.bindValue(":valid_start_date", valid_start_date);
    query.bindValue(":valid_end_date", valid_end_date);
    query.bindValue(":location_id", location_id);
    query.bindValue(":location_desc", location_desc);
    query.bindValue(":bal_qty", bal_qty);
    query.bindValue(":state_code", state_code);
    log( "test", "SO Data Pushed");
    if (query.exec()) {
        qDebug() << "Query success of SO Data Push";
    } else {
        qDebug() << query.lastError().text();
        return;
    }
}

void PsqlManager::updateSoData(QString auction_id, QString customer_code, QString customer_name, QString so_no, QString so_date, QString so_grade, QString so_coal_size, QString so_qty, QString valid_start_date, QString valid_end_date, QString location_id, QString location_desc, QString bal_qty, QString state_code){
    QSqlQuery query(this->database);
    // query.prepare("INSERT INTO sodata VALUES(:auction_id, :customer_code, :customer_name, :so_no, :so_date, :so_grade, :so_coal_size, :so_qty, :valid_start_date, :valid_end_date, :location_id, :location_desc, :bal_qty, :state_code);");
    query.prepare("UPDATE sodata SET (auction_id, customer_code, customer_name, so_no, so_date, so_grade, so_coal_size, so_qty, valid_start_date, valid_end_date, location_id, location_desc, bal_qty, state_code) = (:auction_id, :customer_code, :customer_name, :so_no, :so_date, :so_grade, :so_coal_size, :so_qty, :valid_start_date, :valid_end_date, :location_id, :location_desc, :bal_qty, :state_code) WHERE so_no=:so_no;");
    query.bindValue(":auction_id", auction_id);
    query.bindValue(":customer_code", customer_code);
    query.bindValue(":customer_name", customer_name);
    query.bindValue(":so_no", so_no);
    query.bindValue(":so_date", so_date);
    query.bindValue(":so_grade", so_grade);
    query.bindValue(":so_coal_size", so_coal_size);
    query.bindValue(":so_qty", so_qty);
    query.bindValue(":valid_start_date", valid_start_date);
    query.bindValue(":valid_end_date", valid_end_date);
    query.bindValue(":location_id", location_id);
    query.bindValue(":location_desc", location_desc);
    query.bindValue(":bal_qty", bal_qty);
    query.bindValue(":state_code", state_code);
    log( "test", "SO Data Pushed");
    if (query.exec()) {
        qDebug() << "Query success of SO Data Update";
    } else {
        qDebug() << query.lastError().text();
        return;
    }
}


// exception table
//DROP TABLE tagexceptions /**WEAK**/;
//-- SQLINES LICENSE FOR EVALUATION USE ONLY
//CREATE TABLE tagexceptions (
//sno SERIAL PRIMARY KEY NOT NULL,
//edatetime TEXT,
//tagno TEXT,
//v_no TEXT,
//oldtagno TEXT,
//optype TEXT,
//description TEXT,
//o_name TEXT,
//area TEXT,
//wb TEXT,
//expected TEXT,
//recorded TEXT,
//comments TEXT,
//sent TEXT DEFAULT 0
//);
void PsqlManager::updateRfidTag(QString newtag, QString v_no){
    QSqlQuery query(this->database);
    query.prepare("UPDATE tags SET tagno=:tagno WHERE v_no=:v_no;");
    query.bindValue(":tagno", newtag);
    query.bindValue(":v_no", v_no);
    if (query.exec()) {
        qDebug() << "Query success: Rfid Tag Updated";
    } else {
        qDebug() << query.lastError().text();
        return;
    }
//    query.next();
}


//INSERT INTO tagexceptions (edatetime, tagno, v_no, oldtagno, optype, description, o_name, area, wb, expected, recorded, comments, sent) VALUES ('2021-08-08 08:08:08', 'TEST1', 'TRUCK1', 'TEST1OLD', 'OPTYPE', 'DESC', 'ASH', 'AREA','WB','EXPECTED', 'RECORDED', 'TEST COMMENT', '1') ;

void PsqlManager::logExceptions(QString edatetime,QString  tagno,QString  v_no,QString  oldtagno,QString  optype,QString  description,QString  o_name,QString  area,QString  wb,QString  expected,QString  recorded,QString  comments,QString  sent){
    QSqlQuery query(this->database);
    query.prepare("INSERT INTO tagexceptions (edatetime, tagno, v_no, oldtagno, optype, description, o_name, area, wb, expected, recorded, comments, sent) VALUES (:edatetime, :tagno, :v_no, :oldtagno, :optype, :description, :o_name, :area, :wb, :expected, :recorded, :comments, :sent) ;");
    query.bindValue(":edatetime", edatetime);
    query.bindValue(":tagno", tagno);
    query.bindValue(":v_no", v_no);
    query.bindValue(":oldtagno", oldtagno);
    query.bindValue(":optype",optype);
    query.bindValue(":description",description);
    query.bindValue(":o_name",o_name);
    query.bindValue(":area",area);
    query.bindValue(":wb",wb);
    query.bindValue(":expected",expected);
    query.bindValue(":recorded",recorded);
    query.bindValue(":comments",comments);
    query.bindValue(":sent",sent);
    if (query.exec()) {
        qDebug() << "Query success: Log Exception";
    } else {
        qDebug() << query.lastError().text();
        return;
    }
}
