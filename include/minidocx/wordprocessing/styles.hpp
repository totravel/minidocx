/**
 * Copyright (C) 2022-2025, Xie Zequn <totravel@foxmail.com>. All rights reserved.
 * Distributed under the MIT License (http://opensource.org/licenses/MIT)
 * Report bugs and download new versions at https://github.com/totravel/minidocx
 */

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
