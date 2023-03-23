#ifndef SODATAMANAGER_H
#define SODATAMANAGER_H

//#include "mainwindow.h"
#include <QDialog>

namespace Ui {
class SoDataManager;
}

class MainWindow;
class SoDataManager : public QDialog
{
    Q_OBJECT

public:
    explicit SoDataManager(MainWindow *ancestor = nullptr);
    ~SoDataManager();

private slots:
    void onpushSoDataGetClicked();
    void onpushSoDataPushClicked();
    void onpushSoDataUpdateClicked();

private:
    MainWindow *parent;
    Ui::SoDataManager *ui;
};

#endif // SODATAMANAGER_H
