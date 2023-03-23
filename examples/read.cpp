
#include "minidocx.hpp"

using namespace docx;

int main()
{
  Document doc("");
  doc.Open("test.docx");

  // std::cout << doc;
  
  return 0;
}
