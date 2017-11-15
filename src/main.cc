#include<QtCore>
#include "xlsxdocument.h"
#include "xlsxformat.h"
#include "xlsxcellrange.h"
#include "xlsxworksheet.h"
#include "report.h"

using namespace QXlsx;

int main()
{
	Document xlsx;

	xlsx.addSheet("Aligns & Borders");
	xlsx.setColumnWidth(2, 20);
	xlsx.setColumnWidth(8, 12);
	xlsx.currentWorksheet()->setGridLinesVisible(false);

	Report *report = new Report;
	report->WriteHorizontalAlignCell(xlsx, "B3", "AlignLeft", Format::AlignLeft);
	report->WriteHorizontalAlignCell(xlsx, "B5", "AlignHCenter", Format::AlignHCenter);
	report->WriteHorizontalAlignCell(xlsx, "D3::D7", "AlignTop", Format::AlignRight);
	report->WriteVerticalAlignCell(xlsx, "D3:D7", "AlignTop", Format::AlignTop);
	report->WriteVerticalAlignCell(xlsx, "F3:F7", "AlignVCenter", Format::AlignVCenter);
	report->WriteVerticalAlignCell(xlsx, "H3:H7", "AlignBottom", Format::AlignBottom);

	
	report->WriteBorderStyleCell(xlsx,"B3","BorderMedium",Format::BorderMedium);
	report->WriteBorderStyleCell(xlsx, "B15", "BorderDashed", Format::BorderDashed);
	report->WriteBorderStyleCell(xlsx, "B17", "BorderDotted", Format::BorderDotted);
	report->WriteBorderStyleCell(xlsx, "B19", "BorderThick", Format::BorderThick);
	report->WriteBorderStyleCell(xlsx, "B21", "BorderDouble", Format::BorderDouble);
	report->WriteBorderStyleCell(xlsx, "B23", "BorderDashDot", Format::BorderDashDot);


	report->WriteSolidFillCell(xlsx, "D13", Qt::red);
	report->WriteSolidFillCell(xlsx, "D15", Qt::blue);
	report->WriteSolidFillCell(xlsx, "D17", Qt::yellow);
	report->WriteSolidFillCell(xlsx, "D19", Qt::magenta);
	report->WriteSolidFillCell(xlsx, "D21", Qt::green);
	report->WriteSolidFillCell(xlsx, "D23", Qt::gray);

	report->WritePatternFillCell(xlsx, "F13",Format::PatternMediumGray,Qt::red);
	report->WritePatternFillCell(xlsx, "F15", Format::PatternDarkHorizontal, Qt::blue);
	report->WritePatternFillCell(xlsx, "F17", Format::PatternDarkVertical, Qt::yellow);
	report->WritePatternFillCell(xlsx, "F19", Format::PatternDarkDown, Qt::magenta);
	report->WritePatternFillCell(xlsx, "F21", Format::PatternLightVertical, Qt::green);
	report->WritePatternFillCell(xlsx, "F23", Format::PatternLightTrellis, Qt::gray);

	report->WriteBorderAndFontColorCell(xlsx, "H13", "Qt::red", Qt::red);
	report->WriteBorderAndFontColorCell(xlsx, "H15", "Qt::blue", Qt::blue);
	report->WriteBorderAndFontColorCell(xlsx, "H17", "Qt::yellow", Qt::yellow);
	report->WriteBorderAndFontColorCell(xlsx, "H19", "Qt::magenta", Qt::magenta);
	report->WriteBorderAndFontColorCell(xlsx, "H21", "Qt::green", Qt::green);
	report->WriteBorderAndFontColorCell(xlsx, "H23", "Qt::gray", Qt::gray);

	//create another sheet
	xlsx.addSheet("Fonts");
	xlsx.write("B3", "Normal");
	Format font_bold;
	font_bold.setFontBold(true);
	xlsx.write("B4", "Blod", font_bold);
	Format font_italic;
	font_italic.setFontItalic(true);
	xlsx.write("B5", "Italic", font_italic);
	Format font_underline;
	font_underline.setFontUnderline(Format::FontUnderlineSingle);
	xlsx.write("B6", "underline", font_underline);
	Format font_strikeout;
	font_strikeout.setFontStrikeOut(true);
	xlsx.write("B7", "StrikeOut", font_strikeout);
	
	report->WriteFontNameCell(xlsx, "D3", "Arial");
	report->WriteFontNameCell(xlsx, "D4", "Arial Black");
	report->WriteFontNameCell(xlsx, "D5", "Comic Sans MS");
	report->WriteFontNameCell(xlsx, "D6", "Courier New");
	report->WriteFontNameCell(xlsx, "D7", "Impact");
	report->WriteFontNameCell(xlsx, "D8", "Times New Roman");
	report->WriteFontNameCell(xlsx, "D9", "Verdana");

	report->WriteFontSizeCell(xlsx, "G3", 10);
	report->WriteFontSizeCell(xlsx, "G4", 12);
	report->WriteFontSizeCell(xlsx, "G5", 14);
	report->WriteFontSizeCell(xlsx, "G6", 16);
	report->WriteFontSizeCell(xlsx, "G7", 18);
	report->WriteFontSizeCell(xlsx, "G8", 20);
	report->WriteFontSizeCell(xlsx, "G9", 25);

	Format font_vertical;
	font_vertical.setRotation(255);
	font_vertical.setFontSize(16);
	xlsx.write("J3", "vertical", font_vertical);
	xlsx.mergeCells("J3:J9");



	xlsx.addSheet("Formulas");
	xlsx.setColumnWidth(1, 2, 40);
	Format rAlign;
	rAlign.setHorizontalAlignment(Format::AlignRight);
	Format lAlign;
	lAlign.setHorizontalAlignment(Format::AlignLeft);
	xlsx.write("B3", 40, lAlign);
	xlsx.write("B4", 30, lAlign);
	xlsx.write("B5", 50, lAlign);
	xlsx.write("A7", "SUM(B3:B5)=", rAlign);
	xlsx.write("B7", "=SUM(B3:B5)", lAlign);
	xlsx.write("A8", "AVERAGE(B3:B5)=", rAlign);
	xlsx.write("B8", "=AVERAGE(B3:B5)", lAlign);
	xlsx.write("A9", "MAX(B3:B5)=", rAlign);
	xlsx.write("B9", "=MAX(B3:B5)", lAlign);
	xlsx.write("A10", "MIN(B3:B5)=", rAlign);
	xlsx.write("B10", "=MIN(B3:B5)", lAlign);
	xlsx.write("A11", "COUNT(B3:B5)=", rAlign);
	xlsx.write("B11", "=COUNT(B3:B5)", lAlign);

	xlsx.write("A13", "IF(B7>100,\"large\",\"small\")=", rAlign);
	xlsx.write("B13", "=IF(B7>100,\"large\",\"small\")", lAlign);

	xlsx.write("A15", "SQRT(25)=", rAlign);
	xlsx.write("B15", "=SQRT(25)", lAlign);
	xlsx.write("A16", "RAND()=", rAlign);
	xlsx.write("B16", "=RAND()", lAlign);
	xlsx.write("A17", "2*PI()=", rAlign);
	xlsx.write("B17", "=2*PI()", lAlign);

	xlsx.write("A19", "UPPER(\"qtxlsx\")=", rAlign);
	xlsx.write("B19", "=UPPER(\"qtxlsx\")", lAlign);
	xlsx.write("A20", "LEFT(\"ubuntu\",3)=", rAlign);
	xlsx.write("B20", "=LEFT(\"ubuntu\",3)", lAlign);
	xlsx.write("A21", "LEN(\"Hello Qt!\")=", rAlign);
	xlsx.write("B21", "=LEN(\"Hello Qt!\")", lAlign);

	Format dateFormat;
	dateFormat.setHorizontalAlignment(Format::AlignLeft);
	dateFormat.setNumberFormat("yyyy-mm-dd");
	xlsx.write("A23", "DATE(2013,8,13)=", rAlign);
	xlsx.write("B23", "=DATE(2013,8,13)", dateFormat);
	xlsx.write("A24", "DAY(B23)=", rAlign);
	xlsx.write("B24", "=DAY(B23)", lAlign);
	xlsx.write("A25", "MONTH(B23)=", rAlign);
	xlsx.write("B25", "=MONTH(B23)", lAlign);
	xlsx.write("A26", "YEAR(B23)=", rAlign);
	xlsx.write("B26", "=YEAR(B23)", lAlign);
	xlsx.write("A27", "DAYS360(B23,TODAY())=", rAlign);
	xlsx.write("B27", "=DAYS360(B23,TODAY())", lAlign);

	xlsx.write("A29", "B3+100*(2-COS(0)))=", rAlign);
	xlsx.write("B29", "=B3+100*(2-COS(0))", lAlign);
	xlsx.write("A30", "ISNUMBER(B29)=", rAlign);
	xlsx.write("B30", "=ISNUMBER(B29)", lAlign);
	xlsx.write("A31", "AND(1,0)=", rAlign);
	xlsx.write("B31", "=AND(1,0)", lAlign);

	xlsx.write("A33", "HYPERLINK(\"http://qt-project.org\")=", rAlign);
	xlsx.write("B33", "=HYPERLINK(\"http://qt-project.org\")", lAlign);

	xlsx.addSheet("NumFormats");
	xlsx.setColumnWidth(2, 40);
	report->WriteInternalNumFormatsCell(xlsx, 4, 2.5681, 2);
	report->WriteInternalNumFormatsCell(xlsx, 5, 2500000, 3);
	report->WriteInternalNumFormatsCell(xlsx, 6, -500, 5);
	report->WriteInternalNumFormatsCell(xlsx, 7, -0.25, 9);
	report->WriteInternalNumFormatsCell(xlsx, 8, 890, 11);
	report->WriteInternalNumFormatsCell(xlsx, 9, 0.75, 12);
	report->WriteInternalNumFormatsCell(xlsx, 10, 41499, 14);
    report->WriteInternalNumFormatsCell(xlsx, 11, 41499, 17);

	report->WriteCustomNumFormatCell(xlsx, 13, 20.5627, "#.###");
	report->WriteCustomNumFormatCell(xlsx, 14, 4.8, "#.00");
	report->WriteCustomNumFormatCell(xlsx, 15, 1.23, "0.00 \"RMB\"");
	report->WriteCustomNumFormatCell(xlsx, 16, 60, "[Red][<=100];[Green][>100]");

	xlsx.addSheet("Merage");
	Format centerAlign;
	centerAlign.setHorizontalAlignment(Format::AlignHCenter);
	centerAlign.setVerticalAlignment(Format::AlignVCenter);
	xlsx.write("B4", "Hello Qt");
	xlsx.mergeCells("B4:F6", centerAlign);
	xlsx.write("B8", 1);
	xlsx.mergeCells("B8:C21", centerAlign);
	xlsx.write("E8", 2);
	xlsx.mergeCells("E8:F21", centerAlign);


	xlsx.addSheet("Grouping");
	qsrand(QDateTime::currentMSecsSinceEpoch());
	for(int row = 2;row<31;row++)
	{
		for(int col = 1;col<10;col++)
		{
			xlsx.write(row, col, qrand() % 100);
		}
	}

	xlsx.groupRows(4, 7);
	xlsx.groupRows(11, 26, false);
	xlsx.groupRows(15, 17, false);
	xlsx.groupRows(20, 22, false);
	xlsx.setColumnWidth(1, 10, 10.0);
	xlsx.groupColumns(1, 2);
	xlsx.groupColumns(5, 8, false);

	xlsx.saveAs("BOOK2.xlsx");

	Document xlsx2("BOOK1.xlsx");
	xlsx2.saveAs("Books.xlsx");
	return 0;

}