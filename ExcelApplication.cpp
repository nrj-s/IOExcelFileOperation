#include "ExcelApplication.h"
#include <QtCore/QFile>
#include <stdexcept>

using namespace std;

ExcelApplication::ExcelApplication(bool closeExcelOnExit)
	: m_excelApplication(new QAxObject("Excel.Application", 0)), m_workBooks(NULL)
	, m_workBook(NULL), m_workSheets(NULL), m_workSheet(NULL), m_usedRange(NULL)
	, m_rows(NULL), m_cols(NULL), m_closeExcelOnExit(closeExcelOnExit)
{
    if (m_excelApplication == NULL)
        throw invalid_argument("Failed to initialize interop with Excel (probably Excel is not installed)");

    m_excelApplication->dynamicCall( "SetVisible(bool)", false ); // hide excel
    m_excelApplication->setProperty( "DisplayAlerts", 0); // disable alerts
}

ExcelApplication::~ExcelApplication()
{
    if (m_excelApplication != NULL)
    {
        if (!m_closeExcelOnExit)
        {
            m_excelApplication->setProperty("DisplayAlerts", 1);
            m_excelApplication->dynamicCall("SetVisible(bool)", true );
        }

        if (m_workBook != NULL && m_closeExcelOnExit)
        {
            m_workBook->dynamicCall("Close (Boolean)", true);
            m_excelApplication->dynamicCall("Quit (void)");
        }
    }
	
	delete m_cols;
	delete m_rows;
	delete m_usedRange;
    delete m_workSheet;
    delete m_workSheets;
    delete m_workBook;
    delete m_workBooks;
    delete m_excelApplication;
}

void ExcelApplication::Open(const QString& fileName)
{
	m_workBooks = m_excelApplication->querySubObject("Workbooks");
	m_workBook = m_workBooks->querySubObject("Open(const QString&)", fileName);
	m_workSheets = m_workBook->querySubObject("Worksheets");
	m_workSheet = m_workBook->querySubObject("Worksheets(int)", 1);
	m_usedRange = m_workSheet->querySubObject("UsedRange");
	m_rows = m_usedRange->querySubObject("Rows");
	m_cols = m_usedRange->querySubObject("Columns");
}

void ExcelApplication::New(const QString& fileName)
{
    /*m_closeExcelOnExit = closeExcelOnExit;
    m_excelApplication = NULL;
    m_sheet = NULL;
    m_sheets = NULL;
    m_workbook = NULL;
    m_workbooks = NULL;
    m_excelApplication = NULL;

    m_excelApplication = new QAxObject( "Excel.Application", 0 );//{00024500-0000-0000-C000-000000000046}

    if (m_excelApplication == NULL)
        throw invalid_argument("Failed to initialize interop with Excel (probably Excel is not installed)");

    m_excelApplication->dynamicCall( "SetVisible(bool)", false ); // hide excel
    m_excelApplication->setProperty( "DisplayAlerts", 0); // disable alerts

    m_workbooks = m_excelApplication->querySubObject( "Workbooks" );
    m_workbook = m_workbooks->querySubObject( "Add" );
    m_sheets = m_workbook->querySubObject( "Worksheets" );
    m_sheet = m_sheets->querySubObject( "Add" );*/
}

void ExcelApplication::Save()
{
    m_workBook->dynamicCall("Save");
}

void ExcelApplication::SaveAs(const QString& fileName)
{
    m_workBook->dynamicCall("SaveAs (const QString&)", fileName);
}

void ExcelApplication::AddWorkSheet(const QString& sheetName)
{
	// set of sheets
	QAxObject* sheets = m_workBook->querySubObject( "Worksheets" );

	// Sheets number
	int intCount = sheets->property("Count").toInt();

	// Capture last sheet and add new sheet
	QAxObject* lastSheet = sheets->querySubObject("Item(int)", intCount);
	sheets->dynamicCall("Add(QVariant)", lastSheet->asVariant());

	// Capture the new sheet and move to after last sheet
	QAxObject* newSheet = sheets->querySubObject("Item(int)", intCount);
	newSheet->setProperty("Name", sheetName);
	lastSheet->dynamicCall("Move(QVariant)", newSheet->asVariant());
}

QVariant ExcelApplication::GetCellValue(int row, int col)
{
	QVariant value;
    QAxObject *cell = m_workSheet->querySubObject("Cells(int,int)", row, col);
    value = cell->dynamicCall("Value()");
    delete cell;

	return value;
}

void ExcelApplication::SetCellValue(int row, int col, const QString& value)
{
    QAxObject *cell = m_workSheet->querySubObject("Cells(int,int)", row, col);
    cell->setProperty("Value",value);
    delete cell;
}

int ExcelApplication::GetRowStart()
{
	return m_usedRange->property("Row").toInt();
}

int ExcelApplication::GetColumnStart()
{
	return m_usedRange->property("Column").toInt();
}

int ExcelApplication::GetRowCount()
{
	return m_rows->property("Count").toInt();
}

int ExcelApplication::GetColumnCount()
{
	return m_cols->property("Count").toInt();
}