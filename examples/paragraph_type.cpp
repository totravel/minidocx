
#include <iostream>
#include "minidocx.hpp"

int main()
{
  docx::Document doc("./g.docx");

  auto p1 = doc.AppendParagraph("This is the 1st paragraph.");
  auto p2 = doc.AppendParagraph("This is the 2nd paragraph.");
  auto p3 = doc.AppendParagraph("This is the 3rd paragraph.");
  auto p4 = doc.AppendParagraph("This is the 4rd paragraph.");
  auto p5 = doc.AppendParagraph("This is the 5rd paragraph.");
  auto p6 = doc.AppendParagraph("This is the 6rd paragraph.");

  p2.AsPageBreak();
  p4.InsertSectionBreak();
  
  std::cout << doc.GetFormatedBody();

  for (auto p = doc.FirstParagraph(); p; p = p.Next()) {
    std::cout << "Type: ";
    switch (p.GetType()) {
      case docx::Paragraph::Type::Text:
        std::cout << "Text\n";
        break;
      case docx::Paragraph::Type::PageBreak:
        std::cout << "PageBreak\n";
        break;
      case docx::Paragraph::Type::SectionBreak:
        std::cout << "SectionBreak\n";
        break;
    }
  }

  doc.Save();
  return 0;
}
