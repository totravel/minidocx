
#include "minidocx/minidocx.hpp"
#include <iostream>


int main()
{
  using namespace md;
  try
  {
    Document doc;
    SectionPointer sect = doc.addSection();

    sect->addParagraph()->addRichText("Example:");

    TablePointer tbl = sect->addTable(5, 7);
    tbl->prop_.width_.type_ = TableProperties::WidthType::Percent;

    tbl->merge(1, 1, 2, 3);
    tbl->merge(2, 4, 2, 2);
    tbl->merge(3, 2, 2, 2);
    tbl->merge(0, 6, 5, 1);
    tbl->dumpStructure();

    tbl->cellAt(0, 0)->addParagraph()->addRichText("AAA");
    tbl->cellAt(0, 1)->addParagraph()->addRichText("BBB");
    tbl->cellAt(0, 2)->addParagraph()->addRichText("CCC");
    tbl->cellAt(0, 3)->addParagraph()->addRichText("DDD");
    tbl->cellAt(0, 4)->addParagraph()->addRichText("EEE");
    tbl->cellAt(0, 5)->addParagraph()->addRichText("FFF");
    tbl->cellAt(0, 6)->addParagraph()->addRichText("GGG");

    doc.saveAs("out/table.docx");
  }
  catch (const Exception& ex)
  {
    std::cerr << ex.what() << std::endl;
  }
  return 0;
}
