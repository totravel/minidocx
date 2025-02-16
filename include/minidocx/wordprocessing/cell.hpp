/**
 * Copyright (C) 2022-2025, Xie Zequn <totravel@foxmail.com>. All rights reserved.
 * Distributed under the MIT License (http://opensource.org/licenses/MIT)
 * Report bugs and download new versions at https://github.com/totravel/minidocx
 */

#pragma once

#include "wordprocessing/container.hpp"
#include "utils/geometry.hpp"

#include <list>


namespace NAMESPACE
{
  class Cell : public Container
  {
    friend class Table;

  public:
    Cell(const size_t row, const size_t col) : rect_{ col, row } {};

  private:
    Rect rect_;

  public:
    inline const Rect& rect() const { return rect_; };
  };
}
