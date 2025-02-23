/**
 * Copyright (C) 2022-2025, Xie Zequn <totravel@foxmail.com>. All rights reserved.
 * Distributed under the MIT License (http://opensource.org/licenses/MIT)
 * Report bugs and download new versions at https://github.com/totravel/minidocx
 */

#pragma once

#include "word/main/properties/paragraph.hpp"
#include "word/main/properties/richtext.hpp"


namespace MINIDOCX_NAMESPACE
{
  struct StyleDefinition
  {
    std::string name_;
    std::string basedOn_;
  };

  struct CharacterStyle : StyleDefinition, RichTextProperties
  {
  };

  struct ParagraphStyle : CharacterStyle, ParagraphProperties
  {
    std::string next_;
  };

}
