#ifndef MAINWINDOW_H
#define MAINWINDOW_H

#include <QMainWindow>
#include <ActiveQt/QAxObject>
#include <QDebug>
#include <QDir>
#include <QRegularExpression>
#include <QRegularExpressionMatch>
#include <QString>

QT_BEGIN_NAMESPACE
namespace Ui { class MainWindow; }
QT_END_NAMESPACE

class MainWindow : public QMainWindow
{
    Q_OBJECT

public:
    MainWindow(QWidget *parent = nullptr);
    ~MainWindow();
    void OpenExcel();
    void AddNewExcel();
    void SaveAndClose();
    int GetRowsCount();
    int GetActualRowsCount();
    QString GetCell(int row,int column);
    QString GetCell(QString number);
    void SetCell(int row,int column,QString value);
    void SetCell(QString number,QString value);
    void SetCellColor(int row,int column, QColor color);
    void SetRangeColor(int startRow, int startColumn, int endRow, int endColumn, QColor color);
    QString columnIndexToLetter(int column);
    void SetRangeColor(const QString &startColumnStr, int startRow, const QString &endColumnStr, int endRow, QColor color);
    int columnLetterToIndex(const QString &columnLetter);
    void SetCellTextCenter(int row,int col);
    void SetRangeValues(const QString &startCell, const QString &endCell, const QString &value);
    void SetRangeTextCenter(const QString& startCell,const QString& endCell);

private:
    Ui::MainWindow *ui;
    QAxObject *m_pExcel;
    QAxObject *m_pWorkBooks;
    QAxObject *m_pWorkBook;
    QAxObject *m_pWorkSheets;
    QAxObject *m_pWorkSheet;
};
#endif // MAINWINDOW_H
