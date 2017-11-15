#include <QtCore>
#include "xlsxdocument.h"
#include "xlsxconditionalformatting.h"

using namespace QXlsx;

int main()
{
	Document xlsx;
	Format format;
	format.setFontBold(true);

	xlsx.write("B1", "(-inf,40)", format);
	xlsx.write("C1", "[30,70]", format);
	xlsx.write("D1", "startsWith 2", format);
	xlsx.write("E1", "dataBaR", format);
	xlsx.write("F1", "colorScale", format);

	for(int row = 3;row<22;++row)
	{
		for(int col = 2;col<22;++col)
		{
			xlsx.write(row, col, qrand() % 100);
		}
	}

	ConditionalFormatting cf1;
	Format fmt1;
	fmt1.setFontColor(Qt::green);
	fmt1.setBorderStyle(Format::BorderDashed);
	cf1.addHighlightCellsRule(ConditionalFormatting::Highlight_LessThan, "40", fmt1);
	cf1.addRange("B3:B21");
	xlsx.addConditionalFormatting(cf1);


	ConditionalFormatting cf2;
	Format fmt2;
	fmt2.setBorderStyle(Format::BorderDotted);
	fmt2.setBorderColor(Qt::blue);
	cf2.addHighlightCellsRule(ConditionalFormatting::Highlight_Between, "30", "70", fmt2);
	cf2.addRange("C3:C21");
	xlsx.addConditionalFormatting(cf2);

	ConditionalFormatting cf3;
	Format fmt3;
	fmt3.setFontStrikeOut(true);
	fmt3.setFontBold(true);
	cf3.addHighlightCellsRule(ConditionalFormatting::Highlight_BeginsWith, "2", fmt3);
	cf3.addRange("D3:D21");
	xlsx.addConditionalFormatting(cf3);

	ConditionalFormatting cf4;
	cf4.addDataBarRule(Qt::blue);
	cf4.addRange("E3:E21");
	xlsx.addConditionalFormatting(cf4);


	ConditionalFormatting cf5;
	cf5.add2ColorScaleRule(Qt::blue, Qt::red);
	cf5.addRange("F3:K21");
	xlsx.addConditionalFormatting(cf5);

	xlsx.saveAs("BOOK1.xlsx");
//	Document xlsx2("BOOK1.xlsx");

//	xlsx2.saveAs("BOOK1.xlsx");

	return 0;
	
}