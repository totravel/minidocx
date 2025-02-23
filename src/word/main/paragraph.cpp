/**
 * Copyright (C) 2022-2025, Xie Zequn <totravel@foxmail.com>. All rights reserved.
 * Distributed under the MIT License (http://opensource.org/licenses/MIT)
 * Report bugs and download new versions at https://github.com/totravel/minidocx
 */

#include "word/main/paragraph.hpp"
#include "word/main/richtext.hpp"
#include "word/main/picture.hpp"
#include "utils/exceptions.hpp"


namespace MINIDOCX_NAMESPACE
{
  RichTextPointer Paragraph::addRichText(const char* text)
  {
    auto run{ std::make_shared<RichText>(text) };
    runs_.push_back(run);
    return run;
  }

  RichTextPointer Paragraph::addRichText(std::string text)
  {
    auto run{ std::make_shared<RichText>(std::move(text)) };
    runs_.push_back(run);
    return run;
  }

  PicturePointer Paragraph::addPicture(const RelationshipId id)
  {
    auto run{ std::make_shared<Picture>(id) };
    runs_.push_back(run);
    return run;
  }

  void Paragraph::deleteRun(const RunPointer& run)
  {
    run->destroy();
    runs_.remove(run);
  }

  void Paragraph::clear()
  {
    for (auto& run : runs_)
      run->destroy();
    runs_.clear();
  }
}
