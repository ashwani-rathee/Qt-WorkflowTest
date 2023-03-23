#ifndef RFIDUPDATE_H
#define RFIDUPDATE_H
#include "qdialog.h"

namespace Ui {
    class RfidUpdate;
}


class RfidManager;
class RfidUpdate: public QDialog
{
    Q_OBJECT
    friend class VehicleActions;
    public:
        explicit RfidUpdate(RfidManager *parent = 0);
        ~RfidUpdate();

    private:
        RfidManager *main;
        Ui::RfidUpdate *ui;


    private slots:
        void onpushButtonReadTagPushed();
        void ClearLineEdits();
        void SaveDataForm();

};

#endif // RFIDUPDATE_H
