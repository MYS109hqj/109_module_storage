#ifndef MAINWINDOW_H
#define MAINWINDOW_H

#include <QMainWindow>
#include <QDebug>
#include<QPushButton>
#include<QTimer>
#include<QTime>

QT_BEGIN_NAMESPACE
namespace Ui { class MainWindow; }
QT_END_NAMESPACE

class MainWindow : public QMainWindow
{
    Q_OBJECT

public:
    MainWindow(QWidget *parent = nullptr);
    ~MainWindow();
    QTimer* timer;
    QTime time;

private slots:
    void on_startButton_clicked();

    void on_pauseButton_clicked();

    void on_resetButton_clicked();

    void update();

    void on_startButton_2_clicked();
    void on_pauseButton_2_clicked();
    void on_resetButton_2_clicked();

private:
    Ui::MainWindow *ui;
    void setButtonIconAndSize(QPushButton* button, const QString& imagePath, int width, int height);
};
#endif // MAINWINDOW_H
