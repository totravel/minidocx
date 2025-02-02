
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
    paraStyle.pPr_.align_ = Alignment::Centered;
    paraStyle.pPr_.outlineLevel_ = ParagraphProperties::OutlineLevel::Level1;
    paraStyle.rPr_.fontSize_ = 32;
    paraStyle.rPr_.color_ = "FF0000";
    doc.addParagraphStyle(paraStyle);

    paraStyle.name_ = "My Heading 2";
    paraStyle.pPr_.outlineLevel_ = ParagraphProperties::OutlineLevel::Level2;
    paraStyle.rPr_.fontSize_ = 28;
    paraStyle.rPr_.color_ = "0000FF";
    doc.addParagraphStyle(paraStyle);

    ParagraphPointer para = sect->addParagraph();
    para->addRichText("Quick Start");
    para->properties().style_ = "My Heading 1";

    ParagraphPointer para2 = sect->addParagraph();
    para2->addRichText("Download and Installation");
    para2->properties().style_ = "My Heading 2";

    doc.saveAs("style.docx");
  }
  catch (const exception& ex)
  {
    std::cerr << ex.what() << std::endl;
  }
  return 0;
}
