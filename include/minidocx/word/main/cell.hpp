/**
 * Copyright (C) 2022-2025, Xie Zequn <totravel@foxmail.com>. All rights reserved.
 * Distributed under the MIT License (http://opensource.org/licenses/MIT)
 * Report bugs and download new versions at https://github.com/totravel/minidocx
 */

#pragma once

#include "word/main/container.hpp"
#include "utils/geometry.hpp"

#include <list>


namespace MINIDOCX_NAMESPACE
{
  class MINIDOCX_API Cell : public Container
  {
    friend class Table;

  public:
    Cell(const size_t row, const size_t col) : rect_{ col, row } {};

  private:
    Rect rect_;

    inline void span(const size_t rows, const size_t cols)
    {
      rect_.setGrid(rows, cols);
    }

  public:
    inline const Rect& rect() const { return rect_; };
  };
}
