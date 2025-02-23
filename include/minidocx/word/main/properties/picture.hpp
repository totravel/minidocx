/**
 * Copyright (C) 2022-2025, Xie Zequn <totravel@foxmail.com>. All rights reserved.
 * Distributed under the MIT License (http://opensource.org/licenses/MIT)
 * Report bugs and download new versions at https://github.com/totravel/minidocx
 */

#pragma once

#include "word/main/properties/base.hpp"


namespace MINIDOCX_NAMESPACE
{
  struct PictureProperties
  {
    // Drawing Object Size 绘制对象尺寸
    // Specifies the size of the extents rectangle in EMUs.
    // 以 EMUs 为单位指定范围矩形的尺寸。
    struct Extent {
      size_t width_ = 4 * 914400; // 4 inches
      size_t height_ = 4 * 914400; // 4 inches

      inline void setSize(const size_t cols, const size_t rows, const size_t dpi, const size_t scale = 100)
      {
        const size_t factor = 914400 / dpi * scale / 100;
        width_ = factor * cols;
        height_ = factor * rows;
      }
    } extent_;


    // 裁剪
    struct Cropping {
      double top_ = 0;
      double bottom_ = 0;
      double left_ = 0;
      double right_ = 0;
    };

    std::optional<Cropping> cropping_;


    // 伸缩
    struct Stretching {
      double top_ = 0;
      double bottom_ = 0;
      double left_ = 0;
      double right_ = 0;
    };

    std::optional<Stretching> stretching_;

  };

}