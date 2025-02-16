/**
 * Copyright (C) 2022-2025, Xie Zequn <totravel@foxmail.com>. All rights reserved.
 * Distributed under the MIT License (http://opensource.org/licenses/MIT)
 * Report bugs and download new versions at https://github.com/totravel/minidocx
 */

#pragma once

#include "wordprocessing/base.hpp"
#include "wordprocessing/properties/table.hpp"
#include "utils/geometry.hpp"

#include <vector>


namespace NAMESPACE
{

  class Cell;
  using CellPointer = std::shared_ptr<Cell>;


  class Table : public Block, public Configurable<TableProperties>
  {
  public:
    Table(const size_t rows, const size_t cols);
    ~Table() override = default;

  private:
    Rect rect_;

  public:
    inline const Rect& rect() const { return rect_; };

  private:
    std::vector<std::vector<size_t>> grid_;
    std::vector<CellPointer> cells_;
    std::vector<Rect> merged_;

  public:
    // Access specified cell without bounds checking.
    inline CellPointer cell(const size_t row, const size_t col) const
    {
      return cells_[grid_[row][col]];
    }

    // Access specified cell with bounds checking.
    CellPointer cellAt(const size_t row, const size_t col) const;

    CellPointer merge(const Rect rect);

    // Merge cells within specified range into a single cell.
    inline CellPointer merge(const size_t row, const size_t col, const size_t rows, const size_t cols)
    {
      return merge({ col, row, cols, rows });
    }

    // Split specified merged cell.
    void split(const size_t row, const size_t col);

    // Dump table's structure to console.
    void dumpStructure() const;

  public:
    //std::string toHtml() const;

  public:
    void clear() override;
  };
}
