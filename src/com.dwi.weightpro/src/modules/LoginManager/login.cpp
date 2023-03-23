#include "login.h"
#include "ui_login.h"
#include <QDialog>
#include <QMessageBox>
#include "mainwindow.h"
#include <QCloseEvent>
#include <QAction>
#include <QToolButton>
Login::Login(MainWindow *parent): ui(new Ui::Login){
    ui->setupUi(this);
    this->setWindowFlags(windowFlags().setFlag(Qt::WindowContextHelpButtonHint, false));
    this->main = parent;
    this->setWindowTitle(main->appname);
    this->setAttribute(Qt::WA_DeleteOnClose);

    ui->dispatch->setChecked(true);
    QPixmap foo( ":/base/weight.png" );
    bool found = !foo.isNull(); //
    qDebug() << "Found: "<< found;

    QAction *action = ui->passwordLineEdit->addAction(QIcon(":/base/eyeOff.png"), QLineEdit::TrailingPosition);
    button = qobject_cast<QToolButton *>(action->associatedWidgets().last());
    button->setCursor(QCursor(Qt::PointingHandCursor));
    connect(button, &QToolButton::pressed, this, &Login::onPressed);
    connect(button, &QToolButton::released, this, &Login::onReleased);

    connect(ui->login_button, SIGNAL(clicked()), this, SLOT(on_pushButtonLoginClicked()));
}

Login::~Login(){
    qDebug() << "Login Destructor!";
    delete ui;
}


void Login::onPressed(){
    QToolButton *button = qobject_cast<QToolButton *>(sender());
    button->setIcon(QIcon(":/base/eyeOn.png"));
    ui->passwordLineEdit->setEchoMode(QLineEdit::Normal);
}

void Login::onReleased(){
    QToolButton *button = qobject_cast<QToolButton *>(sender());
    button->setIcon(QIcon(":/base/eyeOff.png"));
    ui->passwordLineEdit->setEchoMode(QLineEdit::Password);
}
void Login::GetData(QString &username, QString &password, QString &mode){
    username = ui->usernameLineEdit->text();
    password = ui->passwordLineEdit->text();
    mode = ui->buttonbase->checkedButton()->text();
}

void Login::on_pushButtonLoginClicked(){
    QString username;
    QString password;
    QString mode;
    this->GetData(username, password, mode);
    // check the validity of username and password in the main window
    bool check = main->logincheck(username, password, mode);
    if(check){
       QMessageBox::information(this, main->appname, "Valid Login Attempt");
       QString activity = "login";
       main->db.log(username, activity);
       this->accept();
    } else {
        QMessageBox::warning(this,main->appname, "Invalid Login Attempt");
    }
    return;
}


void Login::closeEvent(QCloseEvent *event){
    QMessageBox::StandardButton resBtn = QMessageBox::warning(this, main->appname,
                                                                tr("Do you want to exit?\n"),
                                                                QMessageBox::Cancel | QMessageBox::No | QMessageBox::Yes,
                                                                QMessageBox::No);
    if (resBtn != QMessageBox::Yes) {
        event->ignore();
    } else {
        event->accept();
    }
}

