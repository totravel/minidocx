
#pragma once

#include <algorithm>


namespace NAMESPACE
{
  struct Point
  {
    size_t x_;
    size_t y_;

    Point() : x_{ 0 }, y_{ 0 } {}

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
    Rect(const size_t x1, const size_t y1, const size_t x2, const size_t y2)
      : topLeft_{ x1, y1 }, bottomRight_{ x2, y2 }
    {}

    Rect(const size_t x1, const size_t y1)
      : topLeft_{ x1, y1 }, bottomRight_{ x1 + 1, y1 + 1 }
    {}

    Rect(const Point& pt) : Rect(pt.x_, pt.y_) {}

    inline size_t top() const { return topLeft_.y_; }
    inline size_t bottom() const { return bottomRight_.y_; }

    inline size_t left() const { return topLeft_.x_; }
    inline size_t right() const { return bottomRight_.x_; }

    inline void setTop(const size_t top) { topLeft_.y_ = top; }
    inline void setBottom(const size_t bottom) { bottomRight_.y_ = bottom; }

    inline void setLeft(const size_t left) { topLeft_.x_ = left; }
    inline void setRight(const size_t right) { bottomRight_.x_ = right; }

    inline void setTopLeft(const size_t top, const size_t left) { setTop(top); setLeft(left); }
    inline void setBottomRight(const size_t bottom, const size_t right) { setBottom(bottom); setRight(right); }

    inline size_t x() const { return left(); }
    inline size_t y() const { return top(); }

    inline void setX(const size_t x) { setLeft(x); }
    inline void setY(const size_t y) { setTop(y); }

    inline size_t row() const { return top(); }
    inline size_t col() const { return left(); }

    inline size_t rrow() const { return bottom() - 1; }
    inline size_t rcol() const { return right() - 1; }

    inline size_t endRow() const { return bottom(); }
    inline size_t endCol() const { return right(); }

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

    inline bool equal(const Rect& rect) const
    {
      return topLeft_ == rect.topLeft_ && bottomRight_ == rect.bottomRight_;
    }

    inline bool operator==(const Rect& rhs) const
    {
      return equal(rhs);
    }

    inline Rect intersect(const Rect& rect) const
    {
      return Rect(
        std::max(topLeft_.x_, rect.topLeft_.x_),
        std::max(topLeft_.y_, rect.topLeft_.y_),
        std::min(bottomRight_.x_, rect.bottomRight_.x_),
        std::min(bottomRight_.y_, rect.bottomRight_.y_));
    }

    inline Rect operator&(const Rect& rhs) const
    {
      return intersect(rhs);
    }

    inline bool contains(const Rect& rect) const
    {
      return intersect(rect).equal(rect);
    }

    inline bool contains(const size_t row, const size_t col) const
    {
      return row >= topLeft_.y_ && row < bottomRight_.y_ && col >= topLeft_.x_ && col < bottomRight_.x_;
    }

    inline Rect bound(const Rect& rect) const
    {
      return Rect(
        std::min(topLeft_.x_, rect.topLeft_.x_),
        std::min(topLeft_.y_, rect.topLeft_.y_),
        std::max(bottomRight_.x_, rect.bottomRight_.x_),
        std::max(bottomRight_.y_, rect.bottomRight_.y_));
    }

    inline Rect operator|(const Rect& rhs) const
    {
      return bound(rhs);
    }
  };
}
