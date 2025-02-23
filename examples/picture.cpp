
#include "minidocx/minidocx.hpp"
#include <iostream>

#define TEXT1 u8"\
In China, the spring festival marks the beginning of the new year. It falls on \
the first day of the first month of the Chinese calendar and involves a variety \
of social practices to usher in the new year, pray for good fortune, celebrate \
family reunions and promote community harmony. This process of celebration is \
known as ‘guonian’ (crossing the year)."

#define TEXT2 u8"\
In the days preceding the festival, people clean their homes, stock provisions \
and prepare food. On New Year’s Eve, families dine together and stay up late to \
welcome the new year. During the festival, people wear new clothes, make \
offerings to heaven, earth and ancestors, and extend greetings to elders, \
relatives, friends and neighbours. Public festivities are held by communities, \
cultural institutions, social groups and art troupes."

#define TEXT3 "\
The spring festival promotes family values, social cohesion and peace while \
providing a sense of identity and continuity for the Chinese people."


int main()
{
  using namespace md;
  try
  {
    Document doc;
    SectionPointer sect = doc.addSection();

    sect->addParagraph()->addRichText(TEXT1);

    PicturePointer pict1 = sect->addParagraph()->addPicture(doc.addImage("assets/samples/17528.jpg"));
    pict1->prop_.extent_.setSize(4725, 3173, 300, 20);

    sect->addParagraph()->addRichText(TEXT2);

    PicturePointer pict2 = sect->addParagraph()->addPicture(doc.addImage("assets/samples/17529.jpg"));
    pict2->prop_.extent_.setSize(4838, 3323, 300, 20);

    sect->addParagraph()->addRichText(TEXT3);

    doc.saveAs("out/picture.docx");
  }
  catch (const Exception& ex)
  {
    std::cerr << ex.what() << std::endl;
  }
  return 0;
}
