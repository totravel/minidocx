
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

    tbl->cellAt(0, 0)->addParagraph()->addRichText("AAA");

    tbl->merge(1, 1, 2, 3);
    tbl->cellAt(1, 1)->addParagraph()->addRichText("BBB");

    tbl->merge(2, 4, 2, 2);
    tbl->cellAt(2, 4)->addParagraph()->addRichText("CCC");

    tbl->merge(3, 2, 2, 2);
    tbl->cellAt(3, 2)->addParagraph()->addRichText("DDD");

    tbl->merge(0, 6, 5, 1);
    tbl->cellAt(0, 6)->addParagraph()->addRichText("EEE");

    tbl->dumpStructure();
    
    doc.saveAs("out/table.docx");
  }
  catch (const Exception& ex)
  {
    std::cerr << ex.what() << std::endl;
  }
  return 0;
}
