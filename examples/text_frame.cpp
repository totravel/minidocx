
#include "minidocx.hpp"

using namespace docx;

int main()
{
  Document doc("./text_frame.docx");

  auto p1 = doc.AppendParagraph("Hello, World!");

  auto frame = doc.AppendTextFrame(3500, 3500);
  frame.AppendRun("TEST");
  frame.SetPositionX(TextFrame::Position::Left, TextFrame::Anchor::Page);
  frame.SetPositionY(TextFrame::Position::Top, TextFrame::Anchor::Margin);
  // frame.SetPositionX(CM2Twip(1), TextFrame::Anchor::Margin);
  // frame.SetPositionY(CM2Twip(1), TextFrame::Anchor::Margin);
  frame.SetBorders();
  frame.SetTextWrapping(TextFrame::Wrapping::Around);

  doc.Save();
  return 0;
}
