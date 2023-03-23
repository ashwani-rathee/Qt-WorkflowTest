#ifndef RFIDREISSUE_H
#define RFIDREISSUE_H

#include <QDialog>

namespace Ui {
class RfidReissue;
}

class RfidManager;
class RfidReissue : public QDialog
{
    Q_OBJECT
    friend class VehicleActions;
public:
    explicit RfidReissue(RfidManager *ancestor = nullptr);
    ~RfidReissue();

private:
    RfidManager *parent;
    Ui::RfidReissue *ui;

private slots:
    void ClearLineEdits();
    void SaveDataForm();
};

#endif // RFIDREISSUE_H
