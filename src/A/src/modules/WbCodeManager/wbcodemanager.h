#ifndef WBCODEMANAGER_H
#define WBCODEMANAGER_H

#include <QDialog>

namespace Ui {
class WbCodeManager;
}

class MainWindow;
class WbCodeManager : public QDialog
{
    Q_OBJECT

public:
    explicit WbCodeManager(MainWindow *parent = nullptr);
    ~WbCodeManager();

private slots:
    void onpushGet();
    void onpushUpdate();
    void onpushInsert();

private:
    MainWindow *parent;
    Ui::WbCodeManager *ui;
};

#endif // WBCODEMANAGER_H
