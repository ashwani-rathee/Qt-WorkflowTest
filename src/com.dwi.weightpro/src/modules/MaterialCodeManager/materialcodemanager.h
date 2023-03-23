#ifndef MATERIALCODEMANAGER_H
#define MATERIALCODEMANAGER_H

#include <QDialog>

namespace Ui {
class MaterialCodeManager;
}

class MainWindow;
class MaterialCodeManager : public QDialog
{
    Q_OBJECT

public:
    explicit MaterialCodeManager(MainWindow *parent = nullptr);
    ~MaterialCodeManager();

private slots:
    void onpushGet();
    void onpushUpdate();
    void onpushInsert();

private:
    MainWindow *parent;
    Ui::MaterialCodeManager *ui;
};

#endif // MATERIALCODEMANAGER_H
