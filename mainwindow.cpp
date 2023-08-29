#include "mainwindow.h"
#include "ui_mainwindow.h"

MainWindow::MainWindow(QWidget *parent)
    : QMainWindow(parent)
    , ui(new Ui::MainWindow)
{
    ui->setupUi(this);
}

MainWindow::~MainWindow()
{
    delete ui;
}

void MainWindow::SetCellTextCenter(int row,int col)
{
    QAxObject *range = m_pWorkSheet->querySubObject("Cells(int,int)",row,col);
    range->setProperty("HorizontalAlignment",-4108);
}

void MainWindow::SetRangeTextCenter(const QString& startCell, const QString& endCell)
{
    QRegularExpression re("([A-Z]+)(\\d+)");
    QRegularExpressionMatch startMatch = re.match(startCell);
    QRegularExpressionMatch endMatch = re.match(endCell);

    if (!startMatch.hasMatch() || !endMatch.hasMatch()) {
        qDebug() << "Invalid cell coordinates";
        return;
    }

    QString startColumnStr = startMatch.captured(1);
    QString endColumnStr = endMatch.captured(1);

    int startRow = startMatch.captured(2).toInt();
    int endRow = endMatch.captured(2).toInt();

    qDebug() << "Rows:" << startRow << endRow;
    qDebug() << "Columns:" << startColumnStr << endColumnStr;

    int startCol = 0;
    int endCol = 0;

    for (QChar ch : startColumnStr) {
        startCol = startCol * 26 + (ch.unicode() - 'A' + 1);
    }

    for (QChar ch : endColumnStr) {
        endCol = endCol * 26 + (ch.unicode() - 'A' + 1);
    }

    qDebug() << "Column indices:" << startCol << endCol;

    for (int row = startRow; row <= endRow; ++row) {
        for (int col = startCol; col <= endCol; ++col) {
            QAxObject *range = m_pWorkSheet->querySubObject("Cells(int,int)", row, col);
            range->setProperty("HorizontalAlignment", -4108);
            qDebug() << "Successfully set Cell" << row << col;
        }
    }
}

int MainWindow::GetActualRowsCount()
{
    int iRows = GetRowsCount();

    // 从最后一行开始逐行向上搜索非空行
    for (int row = 2; row <= iRows; row++)
    {
        QAxObject *pRange = m_pWorkSheet->querySubObject("Cells(int,int)", row, 2);
        QVariant value = pRange->property("Value");
        if (value.isNull() || value.toString().isEmpty())
        {
            return row;
        }
    }

    return iRows;
}

void MainWindow::OpenExcel()
{
    QString strExcelPath = "D:\\temp1\\temper run 2\\神庙逃亡2.xlsx";

    m_pWorkBooks = m_pExcel->querySubObject("WorkBooks");

    m_pWorkBook = m_pWorkBooks->querySubObject("Open(const QString&)",strExcelPath);

    if(m_pWorkBook)
    {
        qDebug()<<"Open Excel Success!";
    }

    m_pWorkSheets = m_pWorkBook->querySubObject("Sheets");

    m_pWorkSheet = m_pWorkSheets->querySubObject("Item(int)",2);
}

void MainWindow::AddNewExcel()
{
    m_pWorkBooks = m_pExcel->querySubObject("WorkBooks");

    m_pWorkBooks->dynamicCall("Add");
    m_pWorkBook = m_pExcel->querySubObject("ActiveWorkBook");
    m_pWorkSheets = m_pWorkBook->querySubObject("Sheets");
    m_pWorkSheet = m_pWorkSheets->querySubObject("Item(int)",1);
}

void MainWindow::SaveAndClose()
{
    QString strSavePath = "D:\\temp1\\temper run 2\\神庙逃亡2";
    m_pWorkBook->dynamicCall("SaveAs(const QString&)",QDir::toNativeSeparators(strSavePath));
    m_pWorkBook->dynamicCall("Close()");
    m_pExcel->dynamicCall("Quit()");
    delete m_pExcel;
    m_pExcel = NULL;
}

int MainWindow::GetRowsCount()
{
    int iRows = 0;
    QAxObject *pRows = m_pWorkSheet->querySubObject("Rows");
    iRows = pRows->property("Count").toInt();
    return iRows;
}

QString MainWindow::GetCell(int row,int column)
{
    QAxObject *range = m_pWorkSheet->querySubObject("Cells(int,int)", row, column);
    QString temp =  range->property("Value2").toString();
    qDebug()<<"test GetCell:"<<temp;
    delete range;
    return temp;
}

QString MainWindow::GetCell(QString number)
{
    QRegularExpression re("([A-Z]+)(\\d+)");
    QRegularExpressionMatch startMatch = re.match(number);


    if (!startMatch.hasMatch() ) {
        qDebug() << "Invalid cell coordinates";
        return NULL;
    }

    QString startColumnStr = startMatch.captured(1);
    int startRow = startMatch.captured(2).toInt();


    qDebug() << "Rows:" << startRow;
    qDebug() << "Columns:" << startColumnStr;

    int startCol = 0;

    for (QChar ch : startColumnStr) {
        startCol = startCol * 26 + (ch.unicode() - 'A' + 1);
    }

    qDebug() << "Column indices:" << startCol;


    QAxObject *range = m_pWorkSheet->querySubObject("Cells(int,int)",startRow,startCol);
    QString temp =  range->property("Value2").toString();

    qDebug()<<"test GetCell:"<<temp;
    delete range;
    return temp;
}

void MainWindow::SetCell(int row, int column, QString value)
{

    QAxObject *range = m_pWorkSheet->querySubObject("Cells(int,int)", row, column);
    range->setProperty("Value2", value);
    delete range;

}


void MainWindow::SetCell(QString number,QString value)
{
    QRegularExpression re("([A-Z]+)(\\d+)");
    QRegularExpressionMatch startMatch = re.match(number);


    if (!startMatch.hasMatch() ) {
        qDebug() << "Invalid cell coordinates";
        return;
    }

    QString startColumnStr = startMatch.captured(1);
    int startRow = startMatch.captured(2).toInt();


    qDebug() << "Rows:" << startRow;
    qDebug() << "Columns:" << startColumnStr;

    int startCol = 0;

    for (QChar ch : startColumnStr) {
        startCol = startCol * 26 + (ch.unicode() - 'A' + 1);
    }

    qDebug() << "Column indices:" << startCol;


    QAxObject *range = m_pWorkSheet->querySubObject("Cells(int,int)", startRow, startCol);
    qDebug() << "Successfully set Cell" << startRow << startCol;
    range->dynamicCall("SetValue(const QVariant&)", QVariant(value));

}

void MainWindow::SetRangeValues(const QString &startCell, const QString &endCell, const QString &value)
{
    QAxObject *range = m_pWorkSheet->querySubObject("Range(const QVariant&, const QVariant&)", QVariant(startCell), QVariant(endCell));
    if (range)
    {
        range->dynamicCall("SetValue(const QVariant&)", QVariant(value));
        qDebug() << "Range values set successfully";
    }
    else
    {
        qDebug() << "Failed to get range object";
    }
}


void MainWindow::SetCellColor(int row,int column,QColor color)
{
    QAxObject *pCell = m_pWorkSheet->querySubObject("Cells(int,int)",row,column);
    QAxObject *pInterior = pCell->querySubObject("Interior");
    pInterior->setProperty("Color",color);
}

void MainWindow::SetRangeColor(int startRow, int startColumn, int endRow, int endColumn, QColor color)
{
    QString startColumnStr = columnIndexToLetter(startColumn);
    QString endColumnStr = columnIndexToLetter(endColumn);

    QString rangeAddress = QString("%1%2:%3%4").arg(startColumnStr).arg(startRow)
                               .arg(endColumnStr).arg(endRow);

    QAxObject *pRange = m_pWorkSheet->querySubObject("Range(const QString&)", rangeAddress);

    QAxObject *pInterior = pRange->querySubObject("Interior");
    pInterior->setProperty("Color", color);
}

QString MainWindow::columnIndexToLetter(int column)
{
    QString columnStr;
    while (column > 0)
    {
        int remainder = (column - 1) % 26;
        columnStr.prepend(QChar('A' + remainder));
        column = (column - 1) / 26;
    }
    return columnStr;
}


void MainWindow::SetRangeColor(const QString &startColumnStr, int startRow, const QString &endColumnStr, int endRow, QColor color)
{

    QString rangeAddress = QString("%1%2:%3%4").arg(startColumnStr).arg(startRow)
                               .arg(endColumnStr).arg(endRow);

    QAxObject *pRange = m_pWorkSheet->querySubObject("Range(const QString&)", rangeAddress);

    QAxObject *pInterior = pRange->querySubObject("Interior");
    pInterior->setProperty("Color", color);
}

int MainWindow::columnLetterToIndex(const QString &columnLetter)
{
    int index = 0;
    for (QChar ch : columnLetter)
    {
        index = index * 26 + (ch.unicode() - 'A' + 1);
    }
    return index;
}

