/**
 * Copyright (C) 2022-2025, Xie Zequn <totravel@foxmail.com>. All rights reserved.
 * Distributed under the MIT License (http://opensource.org/licenses/MIT)
 * Report bugs and download new versions at https://github.com/totravel/minidocx
 */

#pragma once

#include "word/main/base.hpp"
#include "word/main/properties/richtext.hpp"

#include <string>
#include <utility>


namespace MINIDOCX_NAMESPACE
{
  class MINIDOCX_API RichText : public Run
  {
  public:
    RichText(const char* text)
      : Run(RunType::RichText), text_{ text } {}

    RichText(const char8_t* text)
      : RichText(reinterpret_cast<const char*>(text)) {}

    RichText(std::string text)
      : Run(RunType::RichText), text_{ std::move(text) } {}
	  
    RichText(std::u8string text)
      : RichText(std::string(reinterpret_cast<const char*>(text.c_str()), text.size())) {}

    ~RichText() override = default;
    
    RichTextProperties prop_;

  private:
    std::string text_;

  public:
    inline std::string& text() { return text_; }
    inline const std::string& text() const { return text_; }

    inline void setText(const char* text) { text_ = text; }
    inline void setText(const char8_t* text) { setText(reinterpret_cast<const char*>(text)); }
    inline void setText(std::string text) { text_ = std::move(text); }
    inline void setText(std::u8string text) { setText(std::string(reinterpret_cast<const char*>(text.c_str()), text.size())); }

  public:
    void clear() override
    {
      text_.clear();
      text_.shrink_to_fit();
    }
  };
}
