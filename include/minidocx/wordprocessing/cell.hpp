
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
    Cell(const Rect& rect) : rect_{ rect } {};
    Cell(const size_t row, const size_t col) : rect_{ col, row } {};

  private:
    Rect rect_;

  public:
    inline const Rect& rect() const { return rect_; };
  };
}
