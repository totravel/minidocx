
#include "minidocx.hpp"

using namespace docx;

int main()
{
  Document doc;

  auto p1 = doc.AppendParagraph("Hello, World!", 12, "Times New Roman");
  auto p2 = doc.AppendParagraph(u8"你好，世界！", 14, u8"宋体");
  auto p3 = doc.AppendParagraph(u8"Hello, 世界！", 16, "Times New Roman", u8"宋体");

  auto p4 = doc.AppendParagraph();
  p4.SetAlignment(Paragraph::Alignment::Centered);

  auto p4r1 = p4.AppendRun("Sample text here...", 12, "Arial");
  p4r1.AppendLineBreak();
  p4r1.SetCharacterSpacing(Pt2Twip(2));

  auto p4r2 = p4.AppendRun("And another line of sample text here...");
  p4r2.SetFontSize(14);
  p4r2.SetFont("Times New Roman");
  p4r2.SetFontColor("FF0000");
  p4r2.SetFontStyle(Run::Bold | Run::Italic);

  //doc.SetReadOnly();
  doc.Save("basic.docx");
  return 0;
}
