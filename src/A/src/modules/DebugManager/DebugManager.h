#pragma once
#ifndef DEBUGMANAGER_H
#define DEBUGMANAGER_H

#include <QMap>
#include <QDialog>
#include <QMetaEnum>
#include <QSqlDatabase>

enum class LogTypes
  {
     Info,
     Warning,
     Alert,
  };

namespace Ui{
    class DebugManager;
}

class MainWindow;
class DebugManager : public QDialog {
    Q_OBJECT;
    Q_ENUM(LogTypes);
    public:
        explicit DebugManager(MainWindow *parent = 0, QString name = "SnowFlake");
        ~DebugManager();
        void log(LogTypes a, QString data);
        QMap<LogTypes, Qt::GlobalColor> map;
        QMap<LogTypes, QString> mapStr;
        QString name;
    private slots:
        void onpushButtonSaveClicked();
        void onpushButtonTestConnectionClicked();

    private:
        QSqlDatabase db;
        MainWindow *parent;
        Ui::DebugManager *ui;
        int MaxLines = 20;
};

#endif // DEBUGMANAGER_H
