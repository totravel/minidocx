/**
 * Copyright (C) 2022-2025, Xie Zequn <totravel@foxmail.com>. All rights reserved.
 * Distributed under the MIT License (http://opensource.org/licenses/MIT)
 * Report bugs and download new versions at https://github.com/totravel/minidocx
 */

#pragma once

#include "config.hpp"
#include <algorithm>


namespace MINIDOCX_NAMESPACE
{
  struct Point
  {
    size_t x_ = 0;
    size_t y_ = 0;

    Point() = default;
    Point(const size_t x, const size_t y) : x_{ x }, y_{ y } {}

    inline bool equal(const Point& pt) const
    {
      return x_ == pt.x_ && y_ == pt.y_;
    }

    inline bool operator==(const Point& rhs) const
    {
      return equal(rhs);
    }
  };


  class Rect
  {
  private:
    Point topLeft_;
    Point bottomRight_;

  public:
    Rect(const Point topLeft, const Point bottomRight)
      : topLeft_{ topLeft }, bottomRight_{ bottomRight }
    {
    }

    Rect(const size_t x, const size_t y, const size_t width = 1, const size_t height = 1)
      : topLeft_{ x, y }, bottomRight_{ topLeft_.x_ + width, topLeft_.y_ + height }
    {
    }


    inline size_t top() const { return topLeft_.y_; }
    inline size_t bottom() const { return bottomRight_.y_; }
    inline size_t left() const { return topLeft_.x_; }
    inline size_t right() const { return bottomRight_.x_; }

    inline void setTop(const size_t top) { topLeft_.y_ = top; }
    inline void setBottom(const size_t bottom) { bottomRight_.y_ = bottom; }
    inline void setLeft(const size_t left) { topLeft_.x_ = left; }
    inline void setRight(const size_t right) { bottomRight_.x_ = right; }


    inline Point topLeft() const { return topLeft_; }
    inline Point bottomRight() const { return bottomRight_; }

    inline void setTopLeft(const size_t top, const size_t left) { setTop(top); setLeft(left); }
    inline void setBottomRight(const size_t bottom, const size_t right) { setBottom(bottom); setRight(right); }


    inline size_t x() const { return left(); }
    inline size_t y() const { return top(); }

    inline void setX(const size_t x) { setLeft(x); }
    inline void setY(const size_t y) { setTop(y); }


    inline size_t row() const { return top(); }
    inline size_t col() const { return left(); }

    inline size_t endRow() const { return bottom(); }
    inline size_t endCol() const { return right(); }

    inline size_t rrow() const { return endRow() - 1; }
    inline size_t rcol() const { return endCol() - 1; }


    inline size_t width() const { return bottomRight_.x_ - topLeft_.x_; }
    inline size_t height() const { return bottomRight_.y_ - topLeft_.y_; }

    inline size_t rows() const { return height(); }
    inline size_t cols() const { return width(); }

    inline void setWidth(const size_t width) { bottomRight_.x_ = topLeft_.x_ + width; }
    inline void setHeight(const size_t height) { bottomRight_.y_ = topLeft_.y_ + height; }
    inline void setSize(const size_t width, const size_t height) { setWidth(width); setHeight(height); }

    inline void setRows(const size_t rows) { setHeight(rows); }
    inline void setCols(const size_t cols) { setWidth(cols); }
    inline void setGrid(const size_t rows, const size_t cols) { setRows(rows); setCols(cols); }


    inline size_t area() const { return width() * height(); }


    inline bool valid() const
    {
      return bottomRight_.x_ > topLeft_.x_ && bottomRight_.y_ > topLeft_.y_;
    }

    inline operator bool() const
    {
      return valid();
    }


    inline bool equal(const Rect& other) const
    {
      return topLeft_ == other.topLeft_ && bottomRight_ == other.bottomRight_;
    }

    inline bool operator==(const Rect& rhs) const
    {
      return equal(rhs);
    }


    inline Rect intersect(const Rect& other) const
    {
      return Rect(
        {
          std::max(topLeft_.x_, other.topLeft_.x_),
          std::max(topLeft_.y_, other.topLeft_.y_)
        },
        {
          std::min(bottomRight_.x_, other.bottomRight_.x_),
          std::min(bottomRight_.y_, other.bottomRight_.y_)
        });
    }

    inline Rect operator&(const Rect& rhs) const
    {
      return intersect(rhs);
    }


    inline bool contains(const Rect& other) const
    {
      return intersect(other).equal(other);
    }

    inline bool contains(const size_t row, const size_t col) const
    {
      return row >= topLeft_.y_ && row < bottomRight_.y_ && col >= topLeft_.x_ && col < bottomRight_.x_;
    }


    inline Rect bound(const Rect& other) const
    {
      return Rect(
        {
          std::min(topLeft_.x_, other.topLeft_.x_),
          std::min(topLeft_.y_, other.topLeft_.y_)
        },
        {
          std::max(bottomRight_.x_, other.bottomRight_.x_),
          std::max(bottomRight_.y_, other.bottomRight_.y_)
        });
    }

    inline Rect operator|(const Rect& rhs) const
    {
      return bound(rhs);
    }
  };
}
