
#include "minidocx/minidocx.hpp"
#include <iostream>

int main()
{
  using namespace md;
  try
  {
    Document doc;
    SectionPointer sect = doc.addSection();

    ParagraphPointer para = sect->addParagraph();
    para->properties().align_ = Alignment::Centered;

    RichTextPointer rich = para->addRichText("Happy Chinese New Year!");
    rich->properties().fontSize_ = 32;
    rich->properties().color_ = "FF0000";

    sect->addParagraph()->addRichText(
      "Spring Festival, known as the Chinese New Year, "
      "is the most important festival celebrated by the Chinese people. "
      "UNESCO inscribed Spring Festival on the Representative List of "
      "the Intangible Cultural Heritage of Humanity in 2024.");

    PicturePointer pict = sect->addParagraph()->addPicture(doc.addImage("samples/17533.jpg"));
    pict->properties().extent_.setSize(4643, 6199, 300, 20);

    doc.properties().title_ = "Chinese New Year";
    doc.properties().author_ = "Quinn";
    doc.properties().lastModifiedBy_ = "John";
    doc.saveAs("example.docx");
  }
  catch (const exception& ex)
  {
    std::cerr << ex.what() << std::endl;
  }

  return 0;
}
