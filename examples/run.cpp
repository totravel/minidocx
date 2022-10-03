
#include <iostream>
#include <string>
#include "minidocx.hpp"

int main()
{
  docx::Document doc("./run.docx");

  auto p = doc.AppendParagraph();
  auto r = p.AppendRun(u8"你好，World!", 16, "Times New Roman", "Microsoft YaHei UI");

  r.SetCharacterSpacing(docx::Pt2Twip(2));

  r.SetFontStyle(docx::Run::Bold | docx::Run::Underline);
  r.SetFontStyle(docx::Run::Bold | docx::Run::Italic);
  auto fontStyle = r.GetFontStyle();
  r.SetFontStyle(fontStyle | docx::Run::Strikethrough);

  auto fontSize = r.GetFontSize();
  std::cout << "Font Size: " << fontSize << std::endl;

  auto characterSpacing = docx::Twip2Pt(r.GetCharacterSpacing());
  std::cout << "Character Spacing: " << characterSpacing << std::endl;

  std::string fontAscii, fontEastAsia;
  r.GetFont(fontAscii, fontEastAsia);
  std::cout << "Font Ascii: "     << fontAscii << std::endl;
  std::cout << "Font East Asia: " << fontEastAsia << std::endl;

  doc.Save();
  return 0;
}
