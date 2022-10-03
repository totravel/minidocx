
#include <iostream>
#include "minidocx.hpp"

int main()
{
  docx::Document doc("./breaks.docx");

  auto p = doc.AppendParagraph();
  p.SetAlignment(docx::Paragraph::Alignment::Left);

  auto r = p.AppendRun();
  r.AppendText("This is");
  r.AppendLineBreak();
  r.AppendText("a simple sentence.");
  p.AppendPageBreak();

  doc.AppendParagraph("see you next page.");

  doc.Save();
  return 0;
}
