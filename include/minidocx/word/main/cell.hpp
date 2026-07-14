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
    struct CellProperties
    {
        struct Shade {
            // Required pattern name like clear (background), solid (foreground), horzStripe,
			// vertStripe, reverseDiagStripe, diagStripe, horzCross, diagCross, thinHorzStripe, 
            // thinVertStripe, thinReverseDiagStripe, thinDiagStripe, thinHorzCross, thinDiagCross
            std::string val_;
            // Foreground shade color
            // If omitted, the value is assumed to be "auto"
            // This color can either be presented as a hex value (in RRGGBB format), 
            // or auto to automatically choose an appropriate color based on the background.
            // 用十六进制数（RRGGBB 格式）表示具体颜色，或 auto 表示根据背景自动调整。
            std::optional<std::string> color_;
            // Background shade color
            // If omitted, the value is assumed to be "auto"
            // This color can either be presented as a hex value (in RRGGBB format), 
            // or auto to automatically choose an appropriate color based on the background.
            // 用十六进制数（RRGGBB 格式）表示具体颜色，或 auto 表示根据背景自动调整。
            std::optional<std::string> fill_;
        };
        std::optional<Shade> shade_;

        std::optional<bool> wrap_;

        //preferred cell width
        std::optional<TableProperties::Width> width_;
    };

  class MINIDOCX_API Cell : public Container
  {
    friend class Table;

  public:
   Cell(const size_t row, const size_t col) : rect_{ col, row } {};

   CellProperties prop_;

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
