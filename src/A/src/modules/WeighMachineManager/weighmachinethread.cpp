#include "weighmachinethread.h"
#include <QMessageBox>
// hellothread/hellothread.cpp
#include <QtSerialPort\QSerialPort>
#include <QRegularExpression>
#include <QTextCodec>

WeighMachineThread::WeighMachineThread(QString portname, QSerialPort::BaudRate baudrate, QSerialPort::DataBits databits, QSerialPort::Parity parity){
    this->portname = portname;
    this->baudrate = baudrate;
    this->databits = databits;
    this->parity = parity;

    serial.setPortName(portname);
     if(!serial.setBaudRate(baudrate))
         qDebug() << serial.errorString();
     if(!serial.setDataBits(databits))
         qDebug() << serial.errorString();
     if(!serial.setParity(parity))
         qDebug() << serial.errorString();
     if(!serial.setFlowControl(QSerialPort::HardwareControl))
         qDebug() << serial.errorString();
     if(!serial.setStopBits(QSerialPort::OneStop))
         qDebug() << serial.errorString();
     if(!serial.open(QIODevice::ReadOnly))
         qDebug() << serial.errorString();


     QObject::connect(&serial, &QSerialPort::readyRead, this, [&]
     {
         if(serial.canReadLine()){
            QString weight = "";
            QByteArray data = serial.readLine().simplified();
            for(int i=0; i<data.length();i++){
                if(data[i] >= '0' && data[i] <='9'){
                    weight = weight + data[i];
                }
            }
            weightgm = weight.toInt();
            emit onIntValueChange();
        }
     });
}


WeighMachineThread::~WeighMachineThread(){
    qDebug() << "Weigh Machine Thread Destructor!!" ;

}

void WeighMachineThread::run()
{
     // qDebug() << "New Machine Thread:" << thread()->currentThreadId();
     // setup serial port manager

}

void WeighMachineThread::setProperties(QString portname, QSerialPort::BaudRate baudrate, QSerialPort::DataBits databits, QSerialPort::Parity parity){

}
