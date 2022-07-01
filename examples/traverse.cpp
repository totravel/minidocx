
#include <iostream>
#include "minidocx.hpp"

int main()
{
  docx::Document doc("./b.docx");

  doc.AppendParagraph("This is the 1st paragraph.");
  doc.AppendParagraph("This is the 2nd paragraph.");
  doc.AppendParagraph("This is the 3rd paragraph.");
  
  auto p = doc.AppendParagraph();
  p.AppendRun("This is the 1st run. ");
  p.AppendRun("This is the 2nd run. ");
  auto r = p.AppendRun();
  r.AppendText("This is the 3rd run. ");
  r.AppendText("This is the 3rd run. ");
  r.AppendText("This is the 3rd run.");

  for (auto p = doc.FirstParagraph(); p; p = p.Next()) {
    std::cout << "paragraph: \n";
    for (auto r = p.FirstRun(); r; r = r.Next()) {
      std::cout << "run: " << r.GetText() << std::endl;
    }
  }

  doc.Save();
  return 0;
}
