#pragma once
#include "qlineedit.h"
#ifndef LOGIN_H
#define LOGIN_H

#include <QDialog>
#include <QSqlDatabase>
#include <QDebug>

namespace Ui{
    class Login;
}

class MainWindow;
class Login : public QDialog
{
    Q_OBJECT

    public:
        explicit Login(MainWindow *parent = 0);
        ~Login();
        void GetData(QString &username, QString &password, QString &mode);

    private slots:
        void on_pushButtonLoginClicked();
        void closeEvent(QCloseEvent *event);
        void onPressed();
        void onReleased();

    private:
        MainWindow *main;
        Ui::Login *ui;
        QToolButton *button;
};

#endif // LOGIN_H
