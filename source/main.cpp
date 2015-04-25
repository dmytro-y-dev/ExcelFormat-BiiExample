#include <ExcelFormat/ExcelFormat.h>

using namespace ExcelFormat;

static void example1(const char* path)
{
	BasicExcel xls;

	 // create sheet 1 and get the associated BasicExcelWorksheet pointer
	xls.New(1);
	BasicExcelWorksheet* sheet = xls.GetWorksheet(0);

	XLSFormatManager fmt_mgr(xls);


	 // Create a table containing an header row in bold and four rows below.

	ExcelFont font_bold;
	font_bold._weight = FW_BOLD; // 700

	CellFormat fmt_bold(fmt_mgr);
	fmt_bold.set_font(font_bold);

	int col, row = 0;

	for(col=0; col<10; ++col) {
		BasicExcelCell* cell = sheet->Cell(row, col);

		cell->Set("TITLE");
		cell->SetFormat(fmt_bold);
	}

	while(++row < 4) {
		for(int col=0; col<10; ++col)
			sheet->Cell(row, col)->Set("text");
	}


	++row;

	ExcelFont font_red_bold;
	font_red_bold._weight = FW_BOLD;
	font_red_bold._color_index = EGA_RED;

	CellFormat fmt_red_bold(fmt_mgr, font_red_bold);
	fmt_red_bold.set_color1(COLOR1_PAT_SOLID);			// solid background
	fmt_red_bold.set_color2(MAKE_COLOR2(EGA_BLUE,0));	// blue background

	CellFormat fmt_green(fmt_mgr, ExcelFont().set_color_index(EGA_GREEN));

	for(col=0; col<10; ++col) {
		BasicExcelCell* cell = sheet->Cell(row, col);

		cell->Set("xxx");
		cell->SetFormat(fmt_red_bold);

		cell = sheet->Cell(row, ++col);
		cell->Set("yyy");
		cell->SetFormat(fmt_green);
	}


	xls.SaveAs(path);
}

int main(int argc, char **argv)
{
    example1("example1.xls");
    
    return 0;
}