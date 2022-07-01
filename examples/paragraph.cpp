
#include <iostream>
#include "minidocx.hpp"

int main()
{
  docx::Document doc("./e.docx");

  doc.AppendParagraph("This is the 2nd paragraph.");
  doc.AppendParagraph("This is the 3rd paragraph.");
  doc.AppendParagraph("This is the 4th paragraph.");
  doc.AppendParagraph("This is the 5th paragraph.");
  doc.PrependParagraph("This is the 1st paragraph.");

  auto p1 = doc.FirstParagraph();
  auto p2 = p1.Next();
  auto p3 = p2.Next();
  auto p5 = doc.LastParagraph();
  auto p4 = p5.Prev();

  std::cout << p1.GetText() << std::endl;
  std::cout << p2.GetText() << std::endl;
  std::cout << p3.GetText() << std::endl;
  std::cout << p4.GetText() << std::endl;
  std::cout << p5.GetText() << std::endl;

  p4.InsertBefore().AppendRun("New paragraph before the 4th paragraph.");
  p4.InsertAfter().AppendRun("New paragraph after the 4th paragraph.");

  doc.Save();
  return 0;
}
