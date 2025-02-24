
#include "minidocx/minidocx.hpp"
#include <iostream>

int main()
{
  using namespace md;
  try
  {
    Document doc;
    SectionPointer sect = doc.addSection();
    sect->prop_.landscape_ = true;

    ParagraphPointer para = sect->addParagraph();
    para->prop_.align_ = Alignment::Centered;

    RichTextPointer rich = para->addRichText("Happy Chinese New Year!");
    rich->prop_.fontSize_ = 32;
    rich->prop_.color_ = "FF0000";

    sect->addParagraph()->addRichText(
      "Spring Festival, known as the Chinese New Year, "
      "is the most important festival celebrated by the Chinese people. "
      "UNESCO inscribed Spring Festival on the Representative List of "
      "the Intangible Cultural Heritage of Humanity in 2024.");

    PicturePointer pict = sect->addParagraph({ .align_ = Alignment::Centered })
      ->addPicture(doc.addImage("assets/samples/17533.jpg"));
    pict->prop_.extent_.setSize(4643, 6199, 300, 20);

    doc.prop_.title_ = "Chinese New Year";
    doc.prop_.author_ = "Quinn";
    doc.prop_.lastModifiedBy_ = "John";
    doc.saveAs("out/example.docx");
  }
  catch (const Exception& ex)
  {
    std::cerr << ex.what() << std::endl;
  }

  return 0;
}
