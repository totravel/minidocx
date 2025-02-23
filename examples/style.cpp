
#include "minidocx/minidocx.hpp"
#include <iostream>


int main()
{
  using namespace md;
  try
  {
    Document doc;
    SectionPointer sect = doc.addSection();

    ParagraphStyle paraStyle;
    paraStyle.name_ = "My Heading 1";
    paraStyle.align_ = Alignment::Centered;
    paraStyle.outlineLevel_ = ParagraphProperties::OutlineLevel::Level1;
    paraStyle.fontSize_ = 32;
    paraStyle.color_ = "FF0000";
    doc.addParagraphStyle(paraStyle);

    paraStyle.name_ = "My Heading 2";
    paraStyle.outlineLevel_ = ParagraphProperties::OutlineLevel::Level2;
    paraStyle.fontSize_ = 28;
    paraStyle.color_ = "0000FF";
    doc.addParagraphStyle(paraStyle);

    ParagraphPointer para = sect->addParagraph();
    para->addRichText("Quick Start");
    para->prop_.style_ = "My Heading 1";

    ParagraphPointer para2 = sect->addParagraph();
    para2->addRichText("Download and Installation");
    para2->prop_.style_ = "My Heading 2";

    doc.saveAs("out/style.docx");
  }
  catch (const Exception& ex)
  {
    std::cerr << ex.what() << std::endl;
  }
  return 0;
}
