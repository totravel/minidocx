
#include <iostream>
#include "minidocx.hpp"

int main()
{
  docx::Document doc("./section.docx");

  auto p1 = doc.AppendParagraph("This is the 1st paragraph of the 1st section (A3).");
  auto p2 = doc.AppendSectionBreak();
  p2.AppendRun("This is the 2nd paragraph of the 1st section (A3).");

  auto p3 = doc.AppendParagraph("This is the 3rd paragraph of the 2nd section.");
  auto p4 = doc.AppendParagraph("This is the 4th paragraph of the 2nd section.");

  auto p5 = doc.AppendParagraph("This is the 5th paragraph of the 3rd section (Landscape).");
  auto p6 = doc.AppendParagraph("This is the 6th paragraph of the 3rd section (Landscape).");
  p6.InsertSectionBreak();

  auto p7 = doc.AppendParagraph("This is the 7th paragraph of the 4th section.");
  auto p8 = doc.AppendParagraph("This is the 8th paragraph of the 4th section.");

  auto s1 = p1.GetSection();
  auto s2 = p4.InsertSectionBreak();
  auto s3 = s2.Next();
  auto s4 = doc.LastSection();

  // std::cout << doc;

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

  s1.SetPageSize(docx::Inch2Twip(docx::A3_W), docx::Inch2Twip(docx::A3_H)); // A3
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

  // std::cout << doc;

  //s2.Merge();
  std::cout << s2.LastParagraph().GetText() << std::endl;

  p2.SetAlignment(docx::Paragraph::Alignment::Left);
  p4.SetAlignment(docx::Paragraph::Alignment::Centered);
  p6.SetAlignment(docx::Paragraph::Alignment::Right);

  doc.Save();
  return 0;
}
