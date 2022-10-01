
#include "minidocx.hpp"

using namespace docx;

int main()
{
  Document doc("./a.docx");

  auto p1 = doc.AppendParagraph("Hello, World!", 12, "Times New Roman");
  auto p2 = doc.AppendParagraph(u8"你好，世界！", 14, u8"宋体");
  auto p3 = doc.AppendParagraph(u8"你好，World!", 16, "Times New Roman", u8"宋体");

  auto p4 = doc.AppendParagraph();
  p4.SetAlignment(Paragraph::Alignment::Centered);

  auto p4r1 = p4.AppendRun("This is a simple sentence. ", 12, "Arial");
  p4r1.SetCharacterSpacing(Pt2Twip(2));

  auto p4r2 = p4.AppendRun(u8"这是一个简单的句子。");
  p4r2.SetFontSize(14);
  p4r2.SetFont(u8"黑体");
  p4r2.SetFontStyle(Run::Bold | Run::Italic);

  doc.Save();
  return 0;
}
