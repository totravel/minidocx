
#include <iostream>
#include "minidocx.hpp"

int main()
{
  docx::Document doc;

  auto p = doc.AppendParagraph();
  p.SetAlignment(docx::Paragraph::Alignment::Left);

  auto r = p.AppendRun();
  r.AppendText("Sample text here...");
  r.AppendLineBreak();
  r.AppendText("And another line of sample text here...");
  p.AppendPageBreak();

  doc.AppendParagraph("Page 2");

  doc.Save("breaks.docx");
  return 0;
}
