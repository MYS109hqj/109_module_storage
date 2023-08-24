#include "mainwindow.h"
#include "ui_mainwindow.h"

MainWindow::MainWindow(QWidget *parent)
    : QMainWindow(parent)
    , ui(new Ui::MainWindow)
{

    ui->setupUi(this);



    int buttonWidth = 64;
    int buttonHeight = 64;
    QString startImagePath = "D:\\temp1\\images\\start.jpg";
    QString pauseImagePath = "D:\\temp1\\images\\pause.jpg";
    QString resetImagePath = "D:\\temp1\\images\\reset.jpg";

    setButtonIconAndSize(ui->startButton, startImagePath, buttonWidth, buttonHeight);
    setButtonIconAndSize(ui->pauseButton, pauseImagePath, buttonWidth, buttonHeight);
    setButtonIconAndSize(ui->resetButton, resetImagePath, buttonWidth, buttonHeight);

    ui->timeLabel->setText("00:00:00");
    QFont f("仿宋",28);
    ui->timeLabel->setFont(f);
    ui->timeLabel->setAlignment(Qt::AlignCenter);

    time.setHMS(0,0,0,0);
    timer = new QTimer(this);
    connect(timer,SIGNAL(timeout()),this,SLOT(update()));

    //另一个一样的计时器，不符合设计目标，未实现
    setButtonIconAndSize(ui->startButton_2, startImagePath, buttonWidth, buttonHeight);
    setButtonIconAndSize(ui->pauseButton_2, pauseImagePath, buttonWidth, buttonHeight);
    setButtonIconAndSize(ui->resetButton_2, resetImagePath, buttonWidth, buttonHeight);

    ui->timeLabel_2->setText("00:00:00");
    ui->timeLabel_2->setFont(f);
    ui->timeLabel_2->setAlignment(Qt::AlignCenter);

}

MainWindow::~MainWindow()
{


    delete ui;
}

void MainWindow::update()
{
    static quint32 time_out = 0;
    time_out++;
    time = time.addSecs(1);
    ui->timeLabel->setText(time.toString("hh:mm:ss"));
}

void MainWindow::on_startButton_clicked()
{
    timer->start(1000);
}

void MainWindow::setButtonIconAndSize(QPushButton* button,const QString& imagePath,int width,int height)
{
    QIcon con(imagePath);
    if(con.isNull())
    {
        qDebug() << "创建对象失败，错误地址："<<imagePath;
    }
    button->setFixedSize(width,height);
    button->setIcon(con);
    button->setIconSize(QSize(width,height));
}

void MainWindow::on_pauseButton_clicked()
{
    timer->stop();
}


void MainWindow::on_resetButton_clicked()
{
    timer->stop();
    time.setHMS(0,0,0,0);
    ui->timeLabel->setText(time.toString("hh:mm:ss"));
}


void MainWindow::on_startButton_2_clicked()
{

}

void MainWindow::on_pauseButton_2_clicked()
{

}

void MainWindow::on_resetButton_2_clicked()
{

}
