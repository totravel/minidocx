/**
 * Copyright (C) 2022-2025, Xie Zequn <totravel@foxmail.com>. All rights reserved.
 * Distributed under the MIT License (http://opensource.org/licenses/MIT)
 * Report bugs and download new versions at https://github.com/totravel/minidocx
 */

#pragma once

#include <optional>
#include <string>


namespace MINIDOCX_NAMESPACE
{
  enum class Alignment {
    Left,       // 左对齐
    Centered,   // 居中对齐
    Right,      // 右对齐
    Justified,  // 两端对齐
    Distributed // 分散对齐
  };

  enum class BorderStyle {
    Single,    // 单实线
    Double,    // 双实线
    Triple,    // 三实线
    Dotted,    // 点虚线
    Dashed,    // 短划线
    DotDash,   // 点划线
    Wave,      // 波浪线
    DoubleWave // 双波浪线
  };

  struct BorderProperties
  {
    BorderStyle style_ = BorderStyle::Single;
    size_t width_ = 4; // 8 = 1 pt
    std::string color_ = "auto"; // "auto" or "RRGGBB"
  };

  struct OutsideBorders
  {
    BorderProperties top_, bottom_, left_, right_;
  };

  struct InsideBorders
  {
    BorderProperties insideHorizontal_, insideVertical_;
  };

  struct TableBorders : OutsideBorders, InsideBorders {};
}
