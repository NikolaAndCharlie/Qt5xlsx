#include "report.h"


using namespace QXlsx;


Report::Report()
{
	
}

Report::~Report()
{
	
}

void Report::WriteHorizontalAlignCell(QXlsx::Document& xlsx, const QString& cell, const QString& text, QXlsx::Format::HorizontalAlignment align)
{
	Format format;
	format.setHorizontalAlignment(align);
	format.setBorderStyle(Format::BorderThin);
	xlsx.write(cell, text, format);
}


void Report::WriteVerticalAlignCell(QXlsx::Document& xlsx, const QString& range, const QString& text, QXlsx::Format::VerticalAlignment align)
{
	Format format;
	format.setVerticalAlignment(align);
	CellRange r(range);
	xlsx.write(r.firstRow(), r.firstColumn(), text);
	xlsx.mergeCells(r, format);
}



void Report::WriteBorderStyleCell(QXlsx::Document& xlsx, const QString& cell, const QString& text, QXlsx::Format::BorderStyle bs)
{
	Format format;
	format.setBorderStyle(bs);
	xlsx.write(cell, text, format);
}


void Report::WriteSolidFillCell(QXlsx::Document& xlsx, const QString& cell,const QColor &color)
{
	Format format;
	format.setPatternBackgroundColor(color);
	xlsx.write(cell, QVariant(), format);
}

void Report::WritePatternFillCell(QXlsx::Document& xlsx, const QString& cell, QXlsx::Format::FillPattern pattern, const QColor& color)
{
	Format format;
	format.setPatternForegroundColor(color);
	format.setFillPattern(pattern);
	xlsx.write(cell, QVariant(), format);
}



void Report::WriteBorderAndFontColorCell(QXlsx::Document& xlsx, const QString& cell, const QString& text, const QColor& color)
{
	Format format;
	format.setBorderStyle(Format::BorderThin);
	format.setBorderColor(color);
	xlsx.write(cell, text, format);
}

void Report::WriteFontNameCell(QXlsx::Document& xlsx, const QString& cell, const QString& text)
{
	Format format;
	format.setFontName(text);
	format.setFontSize(16);
	xlsx.write(cell, text, format);
}

void Report::WriteFontSizeCell(QXlsx::Document& xlsx, const QString& cell, int size)
{
	Format format;
	format.setFontSize(size);
	xlsx.write(cell, "Qt Xlsx", format);
}

void Report::WriteInternalNumFormatsCell(QXlsx::Document& xlsx, int row, double value, int numFmt)
{
	Format format;
	format.setNumberFormatIndex(numFmt);
	xlsx.write(row, 1, value);
	xlsx.write(row, 2, QString("Builtin NumFmt %1").arg(numFmt));
	xlsx.write(row, 3, value, format);
}

void Report::WriteCustomNumFormatCell(QXlsx::Document& xlse, int row, double value, const QString& numFmt)
{
	Format format;
	format.setNumberFormat(numFmt);
	xlse.write(row, 1, value);
	xlse.write(row, 2, numFmt);
	xlse.write(row, 3, value, format);
}






