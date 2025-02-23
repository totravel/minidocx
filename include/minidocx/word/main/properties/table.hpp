/**
 * Copyright (C) 2022-2025, Xie Zequn <totravel@foxmail.com>. All rights reserved.
 * Distributed under the MIT License (http://opensource.org/licenses/MIT)
 * Report bugs and download new versions at https://github.com/totravel/minidocx
 */

#pragma once

#include "word/main/properties/base.hpp"


namespace MINIDOCX_NAMESPACE
{
  struct TableProperties
  {
    // Table Layout 表格布局
    enum class Layout {
      // Uses the preferred widths on the table items to generate
      // the sizing of the table, but then uses the contents of
      // each cell to determine the final column widths.
      // 根据内容自动调整表格。
      Auto,
      // Uses the preferred widths on the table items to generate
      // the final sizing of the table. The width of the table is
      // not changed regardless of the contents of the cells. 
      // 固定列宽。
      Fixed
    } layout_ = Layout::Auto;


    // Specifies the units of the width_ property. 百分比或绝对值。
    enum class WidthType {
      Auto, // Specifies that width is determined by the overall table layout algorithm.
      Percent, // 50 = 1%
      Absolute // 1440 = 1 Inch = 72 points
    };

    // Table Width 首选表格宽度
    struct Width {
      WidthType type_ = WidthType::Auto;
      size_t value_ = 5000;
    } width_;


    // Table Alignment 表格对齐方式
    // This property will not affect the alignment of the table
    // if the table spans the entire width of the page. Also it
    // does not affect the justification of text within the cells
    // of the table.
    // 不影响宽度达到最大的表格和单元格中的内容。
    enum class Alignment {
      Left,   // Align To Leading Edge 左对齐
      Right,  // Align to Trailing Edge 右对齐
      Center, // Align Center 居中对齐
    } align_ = Alignment::Left;

    // Table Borders 表格边框
    TableBorders borders_;
  };

}