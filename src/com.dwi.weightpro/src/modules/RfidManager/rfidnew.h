#ifndef RFIDNEW_H
#define RFIDNEW_H

#include "qstatusbar.h"
#include <QDialog>

namespace Ui {
    class RfidNew;
}

class RfidManager;
class RfidNew : public QDialog
{
    Q_OBJECT
    friend class VehicleActions;
public:
    explicit RfidNew(RfidManager *parent = nullptr);
    ~RfidNew();

private:
    RfidManager *main;
    Ui::RfidNew *ui;

    QStatusBar *bar;

private slots:
    void onpushButtonReadTagPushed();
    void onButtonClickedChangeWeighmentPage(int i);
    void ClearLineEdits();
    void SaveDataForm();
};

#endif // RFIDNEW_H
