#pragma once

#include <ActiveQt/QAxObject>
#include <QtCore/QString>

class ExcelApplication
{
	//Disable copy of object
	ExcelApplication(const ExcelApplication& ref);
    ExcelApplication& operator=(const ExcelApplication& ref);

public:
    explicit ExcelApplication(bool closeExcelOnExit = true);
    ~ExcelApplication();

	void Open(const QString& fileName);
	void New(const QString& fileName);
	void Save();
	void SaveAs(const QString& fileName);

	void AddWorkSheet(const QString& sheetName);

	QVariant GetCellValue(int row, int col);
	void SetCellValue(int row, int col, const QString& value);
    
	int GetRowStart();
	int GetColumnStart();
	int GetRowCount();
	int GetColumnCount();


private:
    QAxObject* m_excelApplication;
    QAxObject* m_workBooks;
    QAxObject* m_workBook;
    QAxObject* m_workSheets;
    QAxObject* m_workSheet;
	QAxObject* m_usedRange;
	QAxObject* m_rows;
	QAxObject* m_cols;
    bool m_closeExcelOnExit;
};