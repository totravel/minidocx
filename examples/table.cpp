
#include "minidocx.hpp"
#include <iostream>

using namespace docx;

int main()
{
  Document doc("./table.docx");

  auto tbl = doc.AppendTable(4, 5);

  tbl.GetCell(0, 0).FirstParagraph().AppendRun("AAA");
  tbl.GetCell(0, 1).FirstParagraph().AppendRun("BBB");
  tbl.GetCell(0, 2).FirstParagraph().AppendRun("CCC");
  tbl.GetCell(0, 3).FirstParagraph().AppendRun("DDD");

  tbl.GetCell(1, 0).FirstParagraph().AppendRun("EEE");
  tbl.GetCell(1, 1).FirstParagraph().AppendRun("FFF");

  tbl.SetAlignment(Table::Alignment::Centered);

  tbl.SetTopBorders(Table::BorderStyle::Single, 1, "FF0000");
  tbl.SetBottomBorders(Table::BorderStyle::Dotted, 2, "00FF00");
  tbl.SetLeftBorders(Table::BorderStyle::Dashed, 3, "0000FF");
  tbl.SetRightBorders(Table::BorderStyle::DotDash, 1, "FFFF00");
  tbl.SetInsideHBorders(Table::BorderStyle::Double, 1, "FF00FF");
  tbl.SetInsideVBorders(Table::BorderStyle::Wave, 1, "00FFFF");

  tbl.SetCellMarginLeft(CM2Twip(0.19));
  tbl.SetCellMarginRight(CM2Twip(0.19));
  
  // std::cout << doc;
  
  doc.Save();
  return 0;
}
