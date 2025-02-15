/**
 * Copyright (C) 2022-2025, Xie Zequn <totravel@foxmail.com>. All rights reserved.
 * Distributed under the MIT License (http://opensource.org/licenses/MIT)
 * Report bugs and download new versions at https://github.com/totravel/minidocx
 */

#pragma once

#include "wordprocessing/base.hpp"
#include "wordprocessing/properties/paragraph.hpp"
#include "wordprocessing/numbering.hpp"
#include "packaging/relationship.hpp"
#include "utils/file.hpp"

#include <memory>
#include <list>


namespace NAMESPACE
{
  class RichText;
  class Picture;
  using RunPointer = std::shared_ptr<Run>;
  using RichTextPointer = std::shared_ptr<RichText>;
  using PicturePointer = std::shared_ptr<Picture>;


  class Paragraph : public Block, public Configurable<ParagraphProperties>
  {
  public:
    Paragraph() : Block(BlockType::Paragraph) {};
    ~Paragraph() override = default;

    NumberingId numId_ = 0;
    NumberingLevel level_ = NumberingLevel::Level1;

  private:
    std::list<RunPointer> runs_;

  public:
    inline std::list<RunPointer> runs() const { return runs_; }

    RichTextPointer addRichText(const char* text);
    inline RichTextPointer addRichText(const char8_t* text) { return addRichText(reinterpret_cast<const char*>(text)); }
    RichTextPointer addRichText(std::string text);

    PicturePointer addPicture(const RelationshipId id);

    void deleteRun(const RunPointer& run);

  public:
    void clear() override;
  };
}
