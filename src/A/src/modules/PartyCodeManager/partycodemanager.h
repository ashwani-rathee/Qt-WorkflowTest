#ifndef PARTYCODEMANAGER_H
#define PARTYCODEMANAGER_H

#include <QDialog>

namespace Ui {
class PartyCodeManager;
}

class MainWindow;
class PartyCodeManager : public QDialog
{
    Q_OBJECT

public:
    explicit PartyCodeManager(MainWindow *parent = nullptr);
    ~PartyCodeManager();

private slots:
    void onpushGet();
    void onpushUpdate();
    void onpushInsert();

private:
    MainWindow *parent;
    Ui::PartyCodeManager *ui;
};

#endif // PARTYCODEMANAGER_H
