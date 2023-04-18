
#include "minidocx.hpp"
#include <iostream>

using namespace docx;

int main()
{
  Document doc("./text_frame.docx");

  doc.AppendParagraph("Hello, World!");

  auto frame = doc.AppendTextFrame(CM2Twip(4), CM2Twip(5));
  frame.AppendRun("This is the text frame paragraph.");
  frame.SetPositionX(TextFrame::Position::Left, TextFrame::Anchor::Page);
  frame.SetPositionY(TextFrame::Position::Top, TextFrame::Anchor::Margin);
  // frame.SetPositionX(CM2Twip(1), TextFrame::Anchor::Margin);
  // frame.SetPositionY(CM2Twip(1), TextFrame::Anchor::Margin);
  frame.SetTextWrapping(TextFrame::Wrapping::Around);

  doc.Save();
  return 0;
}
