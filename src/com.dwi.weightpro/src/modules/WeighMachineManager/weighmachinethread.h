#ifndef WEIGHMACHINETHREAD_H
#define WEIGHMACHINETHREAD_H

#include <QThread>
#include <qlogging.h>
#include <QtSerialPort\QSerialPort>
#include <QDebug>
class WeighMachineThread: public QThread{

    Q_OBJECT

public:
    WeighMachineThread(QString portname = "COM5", QSerialPort::BaudRate baudrate = QSerialPort::Baud2400, QSerialPort::DataBits databits = QSerialPort::Data8, QSerialPort::Parity parity = QSerialPort::NoParity);
    int weightgm = 0;
    void setProperties(QString portname, QSerialPort::BaudRate baudrate, QSerialPort::DataBits databits, QSerialPort::Parity parity);
    ~WeighMachineThread();

signals:
    void onIntValueChange();

private:
    void run();
    QSerialPort serial;

    QString portname;
    QSerialPort::BaudRate baudrate;
    QSerialPort::DataBits databits;
    QSerialPort::Parity parity;

};

#endif // WEIGHMACHINETHREAD_H
