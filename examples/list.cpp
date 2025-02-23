
#include "minidocx/minidocx.hpp"
#include <iostream>


int main()
{
  using namespace md;
  try
  {
    Document doc;
    SectionPointer sect = doc.addSection();
    const NumberingId numId1 = doc.addBulletedListDefinition();
    const NumberingId numId2 = doc.addNumberedListDefinition();

    sect->addParagraph()->addRichText("Example 1:");

    ParagraphPointer para1 = sect->addParagraph();
    para1->addRichText("Level one");
    para1->numId_ = numId1;

    ParagraphPointer para2 = sect->addParagraph();
    para2->addRichText("Level two");
    para2->numId_ = numId1;
    para2->level_ = NumberingLevel::Level2;

    ParagraphPointer para3 = sect->addParagraph();
    para3->addRichText("Level three");
    para3->numId_ = numId1;
    para3->level_ = NumberingLevel::Level3;


    sect->addParagraph()->addRichText("Example 2:");

    ParagraphPointer para4 = sect->addParagraph();
    para4->addRichText("Level one");
    para4->numId_ = numId2;

    ParagraphPointer para5 = sect->addParagraph();
    para5->addRichText("Level two");
    para5->numId_ = numId2;
    para5->level_ = NumberingLevel::Level2;

    ParagraphPointer para6 = sect->addParagraph();
    para6->addRichText("Level three");
    para6->numId_ = numId2;
    para6->level_ = NumberingLevel::Level3;

    doc.saveAs("out/list.docx");
  }
  catch (const Exception& ex)
  {
    std::cerr << ex.what() << std::endl;
  }
  return 0;
}
