#ifndef ALLOTMENTMANAGER_H
#define ALLOTMENTMANAGER_H

#include <QDialog>

namespace Ui {
class AllotmentManager;
}

class MainWindow;
class AllotmentManager : public QDialog
{
    Q_OBJECT

public:
    explicit AllotmentManager(MainWindow *parent = nullptr);
    ~AllotmentManager();

private slots:
    void onpushGet();
    void onpushUpdate();
    void onpushInsert();

private:
    MainWindow *parent;
    Ui::AllotmentManager *ui;
};

#endif // ALLOTMENTMANAGER_H
