/**
 * Copyright (C) 2022-2025, Xie Zequn <totravel@foxmail.com>. All rights reserved.
 * Distributed under the MIT License (http://opensource.org/licenses/MIT)
 * Report bugs and download new versions at https://github.com/totravel/minidocx
 */

#pragma once

#include "wordprocessing/base.hpp"
#include "wordprocessing/properties/richtext.hpp"

#include <string>
#include <utility>


namespace NAMESPACE
{
  class RichText : public Run, public Configurable<RichTextProperties>
  {
  public:
    RichText(const char* text)
      : Run(RunType::RichText), text_{ text } {}

    RichText(const char8_t* text)
      : RichText(reinterpret_cast<const char*>(text)) {}

    RichText(std::string text)
      : Run(RunType::RichText), text_{ std::move(text) } {}

    ~RichText() override = default;

  private:
    std::string text_;

  public:
    inline std::string& text() { return text_; }
    inline const std::string& text() const { return text_; }

    inline void setText(const char* text) { text_ = text; }
    inline void setText(const char8_t* text) { setText(reinterpret_cast<const char*>(text)); }
    inline void setText(std::string text) { text_ = std::move(text); }

  public:
    void clear() override
    {
      text_.clear();
      text_.shrink_to_fit();
    }
  };
}
