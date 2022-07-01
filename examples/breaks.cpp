
#include <iostream>
#include "minidocx.hpp"

int main()
{
  docx::Document doc("./c.docx");

  auto p = doc.AppendParagraph();
  p.SetAlignment(docx::Paragraph::Alignment::Left);

  auto r = p.AppendRun();
  r.AppendText("This is");
  r.AppendLineBreak();
  r.AppendText("a simple sentence.");

  auto pageBreak = doc.AppendPageBreak();

  doc.AppendParagraph("see you next page.");

  std::cout << doc.GetFormatedBody();
  doc.Save();
  return 0;
}
