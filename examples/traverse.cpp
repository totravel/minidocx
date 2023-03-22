
#include <iostream>
#include "minidocx.hpp"

int main()
{
  docx::Document doc("./traverse.docx");

  // Section 1
  doc.AppendParagraph("This is the 1st paragraph.");
  doc.AppendParagraph("This is the 2nd paragraph.");
  doc.AppendParagraph("This is the 3rd paragraph.");

  auto p = doc.AppendSectionBreak();
  p.AppendRun("This is the 1st run. ");
  p.AppendRun("This is the 2nd run. ");

  auto r = p.AppendRun();
  r.AppendText("This is the 1st text. ");
  r.AppendText("This is the 2nd text. ");
  r.AppendText("This is the 3rd text.");

  // Section 2
  doc.AppendParagraph("This is the 5th paragraph.");
  doc.AppendParagraph("This is the 6th paragraph.");
  doc.AppendSectionBreak().AppendRun("This is the 7th paragraph.");

  // Section 3
  doc.AppendParagraph("This is the 8th paragraph.");
  doc.AppendParagraph("This is the 9th paragraph.");


  for (auto p = doc.FirstParagraph(); p; p = p.Next()) {
    std::cout << "paragraph: \n";
    for (auto r = p.FirstRun(); r; r = r.Next()) {
      std::cout << "run: " << r.GetText() << std::endl;
    }
  }

  auto s2 = doc.FirstSection().Next();
  auto last = s2.LastParagraph();
  for (auto p = s2.FirstParagraph(); p; p = p.Next()) {
    p.SetAlignment(docx::Paragraph::Alignment::Centered);
    if (p == last) break;
  }

  doc.Save();
  return 0;
}
