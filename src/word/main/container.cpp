/**
 * Copyright (C) 2022-2025, Xie Zequn <totravel@foxmail.com>. All rights reserved.
 * Distributed under the MIT License (http://opensource.org/licenses/MIT)
 * Report bugs and download new versions at https://github.com/totravel/minidocx
 */

#include "word/main/container.hpp"
#include "word/main/paragraph.hpp"
#include "word/main/table.hpp"


namespace MINIDOCX_NAMESPACE
{
  ParagraphPointer Container::addParagraph()
  {
    auto block{ std::make_shared<Paragraph>() };
    blocks_.push_back(block);
    return block;
  }

  ParagraphPointer Container::addParagraph(ParagraphProperties prop)
  {
    auto block{ addParagraph() };
    block->prop_ = std::move(prop);
    return block;
  }

  TablePointer Container::addTable(const size_t rows, const size_t cols)
  {
    auto block{ std::make_shared<Table>(rows, cols) };
    blocks_.push_back(block);
    return block;
  }

  void Container::deleteBlock(const BlockPointer& block)
  {
    block->destroy();
    blocks_.remove(block);
  }

  void Container::clear()
  {
    for (auto& block : blocks_)
      block->destroy();
    blocks_.clear();
  }
}
