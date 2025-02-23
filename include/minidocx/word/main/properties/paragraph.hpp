/**
 * Copyright (C) 2022-2025, Xie Zequn <totravel@foxmail.com>. All rights reserved.
 * Distributed under the MIT License (http://opensource.org/licenses/MIT)
 * Report bugs and download new versions at https://github.com/totravel/minidocx
 */

#pragma once

#include "word/main/properties/base.hpp"


namespace MINIDOCX_NAMESPACE
{
  struct ParagraphProperties
  {
    std::string style_;


    // Paragraph Alignment 对齐方式
    std::optional<Alignment> align_;


    // Outline Level 大纲级别
    // Specifies the outline level associated with the paragraph.
    // It is used to build the table of contents and does not affect
    // the appearance of the text.
    // 指示段落的大纲级别，用于创建目录，对外观没有影响。
    enum class OutlineLevel {
      Level1, Level2, Level3, Level4, Level5, Level6, Level7, Level8, Level9, BodyText,
    } outlineLevel_ = OutlineLevel::BodyText;


    enum class SpecialIndentationType {
      None,

      // Specifies the additional indentation which shall be applied
      // to the first line of the parent paragraph.
      // This additional indentation is specified relative to the paragraph indentation
      // which is specified for all other lines in the parent paragraph.
      // 首行缩进
      FirstLine,

      // Specifies the indentation which shall be removed from the first line
      // of the parent paragraph, by moving the indentation on the first line
      // back towards the beginning of the direction of text flow.
      // This indentation is specified relative to the paragraph indentation
      // which is specified for all other lines in the parent paragraph.
      // 首行悬挂
      Hanging
    };


    // Paragraph Indentation 缩进
    struct Indentation {
      // Left/Right
      // 左侧/右侧
      struct {
        // Specifies the indentation in hundreths of a character (100 = 1 character)
        // or absolute units (1440 = 1 Inch = 72 points).
        // 以一个字符的百分之一为单位指示缩进（100 = 1 字符）
        // 或者以英寸、厘米、磅等单位指示缩进（1440 = 1 英寸 = 72 磅）。
        bool chars_ = false;
        size_t value_ = 0;
      } left_, right_;

      // Special
      // 特殊
      struct {
        SpecialIndentationType type_ = SpecialIndentationType::None;
        bool chars_ = true;
        size_t value_ = 0;
      } special_;
    };

    std::optional<Indentation> indent_;


    enum class SpacingType {
      // Specifies that the spacing should be determined automatically
      // by the consumer/wordprocessor and the value is ignored.
      // 由字处理软件自动决定是否添加间距。
      Auto,

      // Specifies the spacing in hundreths of a line (100 = 1 line).
      // 以一行的百分之一为单位指示间距（100 = 1 行）。
      Lines,

      // Specifies the spacing in absolute units (1440 = 1 Inch = 72 points).
      // 以英寸、厘米、磅等单位指示间距（1440 = 1 英寸 = 72 磅）。
      Absolute
    };


    enum class LineSpacingType {
      // Single / 1.5 lines / Double / Multiple (240 = 1 line)
      // 单倍行距 / 1.5 倍行距 / 2 倍行距 / 多倍行距（240 = 1 行）
      Lines,

      // At Least (240 = 12 pt)
      // 最小值（240 = 12 磅）
      AtLeast,

      // Exactly (240 = 12 pt)
      // 固定值（240 = 12 磅）
      Exactly
    };


    // Spacing 间距
    struct Spacing {
      // Spacing between paragraphs
      // 段前/段后
      struct {
        SpacingType type_ = SpacingType::Auto;
        size_t value_ = 100;
      } before_, after_;

      // Spacing between lines of a paragaph
      // 行距
      struct {
        LineSpacingType type_ = LineSpacingType::Lines;
        size_t value_ = 240;
      } lineSpacing_;
    };

    std::optional<Spacing> spacing_;


    // Borders 边框
    using ParagraphBorders = OutsideBorders;
    std::optional<ParagraphBorders> borders_;


    // Keep with next 与下段同页
    // Specifies that the paragraph (or at least part of it)
    // should be rendered on the same page as the next paragraph
    // when possible.
    bool keepNext_ = false;

    // Keep lines together 段中不分页
    // Specifies that all lines of the paragraph are to be kept
    // on a single page when possible.
    bool keepLines_ = false;

    // Page break before 段前分页
    // Specifies that the contents of this paragraph are rendered
    // on the start of a new page.
    bool pageBreakBefore_ = false;

    // Page break after 段后分页
    // Specifies that the contents of the next paragraph are rendered
    // on the start of a new page.
    bool pageBreakAfter_ = false;
  };

}