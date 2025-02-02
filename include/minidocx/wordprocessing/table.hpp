
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
    friend class Document;

  public:
    Table(const size_t rows, const size_t cols);
    ~Table() override = default;

  private:
    Rect rect_;

  public:
    inline const Rect& rect() const { return rect_; };

  private:
    std::vector<std::vector<size_t>> indices_;
    std::vector<CellPointer> cells_;
    std::vector<Rect> merged_;

  public:
    inline CellPointer cell(const size_t row, const size_t col) const
    {
      return cells_[indices_[row][col]];
    }

    CellPointer cellAt(const size_t row, const size_t col) const;

  public:
    CellPointer merge(const Rect rect);

    inline CellPointer merge(const size_t row, const size_t col, const size_t rows, const size_t cols)
    {
      return merge({ col, row, col + cols, row + rows });
    }

    void split(const size_t row, const size_t col);

    void dumpStructure() const;

  public:
    //std::string toHtml() const;

  public:
    void clear() override;
  };
}
