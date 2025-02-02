
#pragma once

#include "wordprocessing/properties/paragraph.hpp"
#include "wordprocessing/properties/richtext.hpp"


namespace NAMESPACE
{
  struct StyleDefinition
  {
    std::string name_;
    std::string basedOn_;
  };

  struct RichTextStyle : StyleDefinition
  {
    RichTextProperties rPr_;
  };

  struct ParagraphStyle : RichTextStyle
  {
    std::string next_;
    ParagraphProperties pPr_;
  };

}
