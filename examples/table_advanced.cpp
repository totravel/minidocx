
#include "minidocx.hpp"
#include <iostream>

using namespace docx;

int main()
{
  Document doc("./table_advanced.docx");

  doc.AppendParagraph("Table 1");
  auto t1 = doc.AppendTable(4, 5);

  auto c11 = t1.GetCell(1, 1);
  auto c12 = t1.GetCell(1, 2);
  if (t1.MergeCells(c11, c12)) {
    std::cout << "c11 c12 merged\n";
  }

  auto c04 = t1.GetCell(0, 4);
  auto c14 = t1.GetCell(1, 4);
  if (t1.MergeCells(c04, c14)) {
    std::cout << "c04 c14 merged\n";
  }

  auto c11_12 = t1.GetCell(1, 2);
  auto c13 = t1.GetCell(1, 3);
  if (t1.MergeCells(c12, c13)) {
    std::cout << "c11_12 c13 merged\n";
  }

  auto c10 = t1.GetCell(1, 0);
  auto c11_12_13 = t1.GetCell(1, 1);
  if (t1.MergeCells(c10, c13)) {
    std::cout << "c10 c11_12_13 merged\n";
  }

  auto c24 = t1.GetCell(2, 4);
  auto c34 = t1.GetCell(3, 4);
  if (t1.MergeCells(c24, c34)) {
    std::cout << "c24 c34 merged\n";
  }

  if (t1.MergeCells(c24, c04)) {
    std::cout << "c24 c04 merged\n";
  }

  t1.GetCell(0, 0).FirstParagraph().AppendRun("AAA");
  t1.GetCell(0, 1).FirstParagraph().AppendRun("BBB");
  t1.GetCell(0, 2).FirstParagraph().AppendRun("CCC");
  t1.GetCell(0, 3).FirstParagraph().AppendRun("DDD");
  t1.GetCell(1, 0).FirstParagraph().AppendRun("FFF");

  c04.SetVerticalAlignment(TableCell::Alignment::Center);
  auto c04p1 = c04.FirstParagraph();
  c04p1.SetAlignment(Paragraph::Alignment::Centered);
  c04p1.AppendRun("EEE");

  doc.AppendParagraph("Table 2");
  auto t2 = doc.AppendTable(4, 4);

  t2.MergeCells(t2.GetCell(0, 1), t2.GetCell(1, 1));
  t2.MergeCells(t2.GetCell(0, 1), t2.GetCell(2, 1));
  t2.MergeCells(t2.GetCell(0, 2), t2.GetCell(1, 2));
  t2.MergeCells(t2.GetCell(0, 2), t2.GetCell(2, 2));
  t2.MergeCells(t2.GetCell(0, 1), t2.GetCell(2, 2));

  t2.MergeCells(t2.GetCell(0, 0), t2.GetCell(1, 0));
  t2.MergeCells(t2.GetCell(1, 0), t2.GetCell(2, 0));

  t2.MergeCells(t2.GetCell(0, 0), t2.GetCell(0, 1));

  // std::cout << doc;

  doc.Save();
  return 0;
}
