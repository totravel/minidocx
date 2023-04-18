
#include "minidocx.hpp"

using namespace docx;

int main()
{
  docx::Document doc("./page_num.docx");

  auto p1 = doc.AppendParagraph("This is the 3rd paragraph in the 1st section.");
  auto p2 = doc.AppendParagraph("This is the 4th paragraph in the 1st section.");
  p2.InsertSectionBreak();

  auto p3 = doc.AppendParagraph("This is the 5th paragraph in the 2nd section.");
  auto p4 = doc.AppendParagraph("This is the 6th paragraph in the 2nd section.");

  auto s1 = p1.GetSection();
  auto s2 = p3.GetSection();

  s1.SetPageNumber(docx::Section::PageNumberFormat::NumberInDash);
  s2.SetPageNumber(docx::Section::PageNumberFormat::UpperLetter, 1);

  std::cout << doc << std::endl;
  doc.Save();
  return 0;
}
