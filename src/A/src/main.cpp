// Includes
#include <QApplication>
#include <mainwindow.h>

#include "sqlite3.h"
#include "uhfreader.h"

#include "TestSuite.h"

// environment variables will mess you up!

// Driver Code
int main(int argc, char *argv[])
{

    // setup lambda
    int status = 0;
    auto runTest = [&status, argc, argv](QObject* obj) {
        status |= QTest::qExec(obj, argc, argv);
    };

    // run suite
    auto &suite = TestSuite::suite();
    for (auto it = suite.begin(); it != suite.end(); ++it) {
        runTest(*it);
    }

    // Hi!!
    // QT application declaration
    QApplication app(argc, argv);

    // app name
    QString appname = "Test Company";

    // create main window
    MainWindow *main = new MainWindow(&app, appname);

    int res = main->loginDialog();
    if(res){
        QString activity = "login";
        main->pSetup();
        main->show();
        // return result of app
        return app.exec();
    }
    else{
        main->close();
        return 0;
    }

    return app.exec();
}
