
#include "minidocx.hpp"
#include <iostream>

using namespace docx;

int main()
{
  Document doc("./table.docx");

  auto t1 = doc.AppendTable(3, 4);

  t1.GetCell(0, 0).FirstParagraph().AppendRun("AAA");
  t1.GetCell(0, 1).FirstParagraph().AppendRun("BBB");
  t1.GetCell(0, 2).FirstParagraph().AppendRun("CCC");
  t1.GetCell(0, 3).FirstParagraph().AppendRun("DDD");

  t1.GetCell(1, 0).FirstParagraph().AppendRun("EEE");
  t1.GetCell(1, 1).FirstParagraph().AppendRun("FFF");
  t1.GetCell(1, 2).FirstParagraph().AppendRun("GGG");

  t1.SetAlignment(Table::Alignment::Centered);

  t1.SetTopBorders(Table::BorderStyle::Single, 1, "FF0000");
  t1.SetBottomBorders(Table::BorderStyle::Dotted, 2, "00FF00");
  t1.SetLeftBorders(Table::BorderStyle::Dashed, 3, "0000FF");
  t1.SetRightBorders(Table::BorderStyle::DotDash, 1, "FFFF00");
  t1.SetInsideHBorders(Table::BorderStyle::Double, 1, "FF00FF");
  t1.SetInsideVBorders(Table::BorderStyle::None);

  t1.SetCellMarginLeft(CM2Twip(0.19));
  t1.SetCellMarginRight(CM2Twip(0.19));

  doc.Save();
  std::cout << doc;
  return 0;
}
