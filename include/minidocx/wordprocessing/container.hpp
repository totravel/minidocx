/**
 * Copyright (C) 2022-2025, Xie Zequn <totravel@foxmail.com>. All rights reserved.
 * Distributed under the MIT License (http://opensource.org/licenses/MIT)
 * Report bugs and download new versions at https://github.com/totravel/minidocx
 */

#pragma once

#include "wordprocessing/base.hpp"

#include <memory>
#include <list>


namespace NAMESPACE
{
  class Paragraph;
  class Table;
  using BlockPointer = std::shared_ptr<Block>;
  using ParagraphPointer = std::shared_ptr<Paragraph>;
  using TablePointer = std::shared_ptr<Table>;


  class Container : public Destroyable
  {
    friend class Document;

  private:
    std::list<BlockPointer> blocks_;

  public:
    inline std::list<BlockPointer> blocks() const { return blocks_; }

    ParagraphPointer addParagraph();
    TablePointer addTable(const size_t rows, const size_t cols);
    
    void deleteBlock(const BlockPointer& block);

  public:
    void clear() override;
  };

}
