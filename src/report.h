#ifndef REPORT_H_
#define REPORT_H_


#include "xlsxdocument.h"
#include "xlsxcellrange.h"
#include "xlsxworksheet.h"

class Report
{
public:
	Report();
	~Report();

	void WriteHorizontalAlignCell(QXlsx::Document &xlsx, const QString &cell, const QString &text, QXlsx::Format::HorizontalAlignment align);
	void WriteVerticalAlignCell(QXlsx::Document &xlsx, const QString &range, const QString &text, QXlsx::Format::VerticalAlignment align);
	void WriteSolidFillCell(QXlsx::Document &xlsx, const QString &cell,const QColor &color);
	void WriteBorderStyleCell(QXlsx::Document &xlsx, const QString &cell, const QString &text, QXlsx::Format::BorderStyle bs);
	void WritePatternFillCell(QXlsx::Document &xlsx, const QString &cell, QXlsx::Format::FillPattern pattern, const QColor &color);
	void WriteBorderAndFontColorCell(QXlsx::Document &xlsx, const QString &cell, const QString &text, const QColor &color);
	void WriteFontNameCell(QXlsx::Document &xlsx, const QString &cell, const QString &text);
	void WriteFontSizeCell(QXlsx::Document &xlsx, const QString &cell, int size);
	void WriteInternalNumFormatsCell(QXlsx::Document &xlsx, int row, double value, int numFmt);
	void WriteCustomNumFormatCell(QXlsx::Document &xlse, int row, double value, const QString &numFmt);
};





#endif

