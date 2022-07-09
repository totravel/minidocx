
#include <iostream>
#include "minidocx.hpp"

int main()
{
  docx::Document doc("./f.docx");

  auto p1 = doc.AppendParagraph("This is the 1st paragraph.");
  auto p2 = doc.AppendSectionBreak();
  p2.AppendRun("This is the 2nd paragraph.");

  auto p3 = doc.AppendParagraph("This is the 3rd paragraph.");
  auto p4 = doc.AppendParagraph("This is the 4th paragraph.");

  auto p5 = doc.AppendParagraph("This is the 5th paragraph.");
  auto p6 = doc.AppendParagraph("This is the 6th paragraph.");
  p6.InsertSectionBreak();

  auto p7 = doc.AppendParagraph("This is the 7th paragraph.");
  auto p8 = doc.AppendParagraph("This is the 8th paragraph.");

  auto s1 = p1.GetSection();
  auto s2 = p4.InsertSectionBreak();
  auto s3 = s2.Next();
  auto s4 = doc.LastSection();

  std::cout << s1.LastParagraph().GetText() << std::endl;
  std::cout << s2.LastParagraph().GetText() << std::endl;
  std::cout << s3.LastParagraph().GetText() << std::endl;
  std::cout << s4.LastParagraph().GetText() << std::endl;

  auto firstSection = doc.FirstSection();
  if (s1 == firstSection) {
    std::cout << "They're the same Section\n";
  }

  for (auto s = doc.LastSection(); s; s = s.Prev()) {
    std::cout << s.LastParagraph().GetText() << std::endl;
  }

  s1.SetPageSize(docx::MM2Twip(297), docx::MM2Twip(420)); // A3
  s3.SetPageOrient(docx::Section::Orientation::Landscape);

  int w, h;
  s2.GetPageSize(w, h);
  std::cout << "Page Size: " 
            << docx::Twip2MM(w) << "mm" 
            << " * " 
            << docx::Twip2MM(h) << "mm\n";

  auto orient = s2.GetPageOrient() == docx::Section::Orientation::Landscape
              ? "Landscape"
              : "Portrait";
  std::cout << "Orientation: " << orient << std::endl;

  s4.SetPageMargin(docx::CM2Twip(2.54),  docx::CM2Twip(2.54),
                   docx::CM2Twip(3.175), docx::CM2Twip(3.175));
  s4.SetPageMargin(docx::CM2Twip(1.5),   docx::CM2Twip(1.75));

  std::cout << doc.GetFormatedBody();

  s2.Merge();
  std::cout << s2.LastParagraph().GetText() << std::endl;

  p2.SetAlignment(docx::Paragraph::Alignment::Left);
  p4.SetAlignment(docx::Paragraph::Alignment::Left);
  p6.SetAlignment(docx::Paragraph::Alignment::Left);

  doc.Save();
  return 0;
}
