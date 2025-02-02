
#include "wordprocessing/table.hpp"
#include "wordprocessing/cell.hpp"
#include "utils/exceptions.hpp"

#include <iostream>


namespace NAMESPACE
{
  Table::Table(const size_t rows, const size_t cols)
    : Block(BlockType::Table), rect_{ 0, 0, cols, rows }, indices_(rows, std::vector<size_t>(cols))
  {
    size_t k = 0;
    cells_.reserve(rows * cols);
    for (size_t i = 0; i < rows; i++) {
      for (size_t j = 0; j < cols; j++) {
        indices_[i][j] = k++;
        cells_.push_back(std::make_shared<Cell>(i, j));
      }
    }
  }

  CellPointer Table::cellAt(const size_t row, const size_t col) const
  {
    if (!rect_.contains(row, col))
      throw invalid_parameter();
    return cell(row, col);
  }

  CellPointer Table::merge(const Rect expected)
  {
    if (!expected || expected.area() == 1 || !rect_.contains(expected))
      throw invalid_parameter();

    Rect rect = expected;
    auto mergedIter = merged_.begin();
    while (mergedIter != merged_.end()) {
      const Rect& merged = *mergedIter;
      const Rect tmp = rect & merged;
      if (tmp && tmp != merged) {
        rect = rect | merged;
        mergedIter = merged_.begin();
      }
      else {
        mergedIter++;
      }
    }
    std::erase_if(merged_, [&rect](Rect& r) { return rect & r; });
    merged_.push_back(rect);

    const size_t k = indices_[rect.row()][rect.col()];
    for (size_t i = rect.row(); i < rect.endRow(); i++) {
      for (size_t j = rect.col(); j < rect.endCol(); j++) {
        indices_[i][j] = k;
      }
    }
    cells_[k]->rect_.setBottomRight(rect.bottom(), rect.right());
    return cells_[k];
  }

  void Table::split(const size_t row, const size_t col)
  {
    Rect pos(col, row);
    if (!rect_.contains(pos))
      return;

    for (auto mergedIter = merged_.begin(); mergedIter != merged_.end(); mergedIter++) {
      const Rect& merged = *mergedIter;
      if (merged & pos) {

        size_t k = indices_[merged.row()][merged.col()];
        for (size_t i = merged.row(); i < merged.endRow(); i++) {
          size_t l = k;
          for (size_t j = merged.col(); j < merged.endCol(); j++) {
            indices_[i][j] = l;
            cells_[l]->rect_.setSize(1, 1);
            l++;
          }
          k += rect().cols();
        }
        merged_.erase(mergedIter);
        break;
      }
    }
  }

  void Table::dumpStructure() const
  {
    std::clog << "-----> Table: " << rect_.rows() << 'x' << rect_.cols() << '\n';

    const size_t rows = rect_.rows();
    const size_t cols = rect_.cols();

    std::vector<std::vector<int>> grid;
    grid.resize(rows, std::vector<int>(cols, 3));

    for (const auto& rect : merged_) {
      size_t i = rect.row();
      while (i < rect.rrow()) {
        size_t j = rect.col();
        while (j < rect.rcol()) {
          grid[i][j] = 0;
          j++;
        }
        grid[i][j] = 1;
        i++;
      }
      size_t j = rect.col();
      while (j < rect.rcol()) {
        grid[i][j] = 2;
        j++;
      }
      grid[i][j] = 3;
    }

    std::clog << '+';
    for (size_t i = 0; i < cols; i++)
      std::clog << "---+";
    std::clog << '\n';

    for (size_t i = 0; i < rows; i++) {
      std::clog << '|';
      for (size_t j = 0; j < cols; j++) {
        if (grid[i][j] == 1 || grid[i][j] == 3)
          std::clog << "   |";
        else
          std::clog << "    ";
      }
      std::clog << '\n';

      std::clog << '+';
      for (size_t j = 0; j < cols; j++) {
        if (grid[i][j] == 2 || grid[i][j] == 3)
          std::clog << "---+";
        else if (grid[i][j] == 1)
          std::clog << "   +";
        else
          std::clog << "    ";
      }
      std::clog << '\n';
    }
  }

  void Table::clear()
  {
    for (auto& cell : cells_)
      cell->destroy();
    cells_.clear();
    cells_.shrink_to_fit();

    indices_.clear();
    indices_.shrink_to_fit();

    merged_.clear();
    merged_.shrink_to_fit();
  }
}