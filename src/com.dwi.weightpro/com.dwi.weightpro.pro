QT += widgets serialport sql network quick core gui
QT += multimedia multimediawidgets
QT += testlib

greaterThan(QT_MAJOR_VERSION, 4): QT += widgets

CONFIG += c++17
CONFIG += console
CONFIG += testcase

TARGET = weightpro

#Release:DESTDIR = C:\Users\lenono\Documents\WeighBridgeAppCode\build\com.dwi.weightpro
#Debug:DESTDIR = C:\Users\lenono\Documents\WeighBridgeAppCode\build\com.dwi.weightpro

#CONFIG(debug, debug|release) {
#    DESTDIR = build\debug
#} else {
#    DESTDIR = build\release
#}

SOURCES += src/main.cpp \
    src/modules/MainWindowManager/mainwindow.cpp \
    src/modules/PsqlManager/psqlmanager.cpp \
    src/modules/PsqlManager/sqlbuilder/Config.cpp \
    src/modules/PsqlManager/sqlbuilder/Deleter.cpp \
    src/modules/PsqlManager/sqlbuilder/Inserter.cpp \
    src/modules/PsqlManager/sqlbuilder/Query.cpp \
    src/modules/PsqlManager/sqlbuilder/Selector.cpp \
    src/modules/PsqlManager/sqlbuilder/Updater.cpp \
    src/modules/PsqlManager/sqlbuilder/Where.cpp \
    src/modules/RfidManager/rfidmanager.cpp \
    src/modules/RfidManager/rfidreissue.cpp \
    src/modules/RfidManager/rfidupdate.cpp \
    src/modules/RfidManager/rfidnew.cpp \
    src/modules/RfidManager/vehicleactions.cpp \
    src/modules/DebugManager/DebugManager.cpp \
    src/modules/MaterialCodeManager/materialcodemanager.cpp \
    src/modules/PartyCodeManager/partycodemanager.cpp \
    src/modules/SoDataManager/sodatamanager.cpp \
    src/modules/TestManager/TestSuite.cpp \
    src/modules/TestManager/testqstring.cpp \
    src/modules/WbCodeManager/wbcodemanager.cpp \
    src/modules/LoginManager/login.cpp \
    src/modules/WeighMachineManager/weighmachinethread.cpp \
    src/modules/CctvManager/videoframegrabber.cpp \
    src/modules/AllotmentManager/allotmentmanager.cpp

HEADERS += src/modules/LoginManager/login.h \
    src/modules/MainWindowManager/mainwindow.h \
    src/modules/MaterialCodeManager/materialcodemanager.h \
    src/modules/PartyCodeManager/partycodemanager.h \
    src/modules/SoDataManager/sodatamanager.h \
    src/modules/TestManager/TestSuite.h \
    src/modules/TestManager/testqstring.h \
    src/modules/WbCodeManager/wbcodemanager.h \
    src/modules/PsqlManager/psqlmanager.h \
    src/modules/DebugManager/DebugManager.h \
    src/modules/PsqlManager/sqlbuilder/Config.h \
    src/modules/PsqlManager/sqlbuilder/Deleter.h \
    src/modules/PsqlManager/sqlbuilder/Inserter.h \
    src/modules/PsqlManager/sqlbuilder/Query.h \
    src/modules/PsqlManager/sqlbuilder/Selector.h \
    src/modules/PsqlManager/sqlbuilder/Updater.h \
    src/modules/PsqlManager/sqlbuilder/Where.h \
    src/modules/RfidManager/rfidmanager.h \
    src/modules/RfidManager/rfidreissue.h \
    src/modules/RfidManager/rfidupdate.h \
    src/modules/RfidManager/rfidnew.h \
    src/modules/RfidManager/vehicleactions.h \
    src/modules/WeighMachineManager/weighmachinethread.h \
    src/modules/CctvManager/videoframegrabber.h \
    src/modules/AllotmentManager/allotmentmanager.h

FORMS += src/modules/LoginManager/login.ui \
    src/modules/MainWindowManager/mainwindow.ui \
    src/modules/DebugManager/DebugManager.ui \
    src/modules/MaterialCodeManager/materialcodemanager.ui \
    src/modules/PartyCodeManager/partycodemanager.ui \
    src/modules/SoDataManager/sodatamanager.ui \
    src/modules/WbCodeManager/wbcodemanager.ui \
    src/modules/RfidManager/rfidreissue.ui \
    src/modules/RfidManager/rfidupdate.ui \
    src/modules/RfidManager/rfidnew.ui \
    src/modules/RfidManager/vehicleactions.ui \
    src/modules/AllotmentManager/allotmentmanager.ui


INCLUDEPATH += $$PWD/src/modules/PsqlManager/
INCLUDEPATH += $$PWD/src/modules/PsqlManager/sqlbuilder
INCLUDEPATH += $$PWD/src/modules/WeighMachine/
INCLUDEPATH += $$PWD/src/modules/DebugManager/
INCLUDEPATH += $$PWD/src/modules/LoginManager/
INCLUDEPATH += $$PWD/src/modules/MainWindowManager/
INCLUDEPATH += $$PWD/src/modules/CctvManager/
INCLUDEPATH += $$PWD/src/modules/WeighMachineManager/
INCLUDEPATH += $$PWD/src/modules/AllotmentManager/
INCLUDEPATH += $$PWD/src/modules/WbCodeManager/
INCLUDEPATH += $$PWD/src/modules/PartyCodeManager/
INCLUDEPATH += $$PWD/src/modules/SoDataManager/
INCLUDEPATH += $$PWD/src/modules/MaterialCodeManager/

INCLUDEPATH += $$PWD/src/modules/TestManager/


INCLUDEPATH += $$PWD/libs/uhfreader/include
LIBS += "libs/uhfreader/lib/uhfreader.lib"

INCLUDEPATH += $$PWD/libs/sqlcipher/include
LIBS += "libs/sqlcipher/lib/sqlite3.lib"

# Default rules for deployment.
qnx: target.path = /tmp/$${TARGET}/bin
else: unix:!android: target.path = /opt/$${TARGET}/bin
!isEmpty(target.path): INSTALLS += target

RESOURCES += \
    src/resources/resource.qrc
