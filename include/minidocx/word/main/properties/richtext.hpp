/**
 * Copyright (C) 2022-2025, Xie Zequn <totravel@foxmail.com>. All rights reserved.
 * Distributed under the MIT License (http://opensource.org/licenses/MIT)
 * Report bugs and download new versions at https://github.com/totravel/minidocx
 */

#pragma once

#include "word/main/properties/base.hpp"


namespace MINIDOCX_NAMESPACE
{
  struct RichTextProperties
  {
    std::string style_;


    enum class FontTypeHint {
      Default, // No font hint
      EastAsia,
      ComplexScript
    };


    // Font 字体
    struct Font {

      // Font 西文字体
      std::string ascii_ = "Courier New";
      // ASCII (i.e., the first 128 Unicode code points)

      // Asian text font 中文字体
      std::string eastAsia_ = "Simsun";
      // East Asian (i.e., CJK)
      // 东亚（即，中日韩）

      // High ANSI
      std::string hAnsi_ = "Times New Roman";

      // Complex Script (e.g., Arabic)
      // 复杂文种（如，阿拉伯语）
      std::string cs_ = "Times New Roman";

      // Font Type Hint 字体类型提示
      // Specifies the font type which shall be used to format any ambiguous characters
      // that can be mapped into multiple categories of the four mentioned above such as ellipsis.
      // 为组别不唯一的字符，如省略号，选定一种字体。
      FontTypeHint hint_ = FontTypeHint::Default;
    };
    
    std::optional<Font> font_;


    // Font style 字形
    struct FontStyle {
      bool bold_ = false;   // 加粗
      bool italic_ = false; // 倾斜
    } fontStyle_;


    // Font size 字号
    // Specifies the font size in half points (2 = 1 pt).
    // 以二分之一磅为单位指示字号（2 = 1 磅）。
    size_t fontSize_ = 0;


    // Font color 字体颜色
    // This color can either be presented as a hex value (in RRGGBB format), 
    // or auto to automatically choose an appropriate color based on the background.
    // 用十六进制数（RRGGBB 格式）表示具体颜色，或 auto 表示根据背景自动调整。
    std::string color_;


    enum class UnderlineStyle {
      None,
      Words,           // 字下加线

      Single,          // 单实线
      Double,          // 双实线
      Thick,           // 粗单实线

      Dotted,          // 点虚线
      DottedHeavy,     // 粗点虚线

      Dash,            // 短划线
      DashedHeavy,     // 粗短划线

      DashLong,        // 长划线
      DashLongHeavy,   // 粗长划线

      DotDash,         // 点划线
      DashDotHeavy,    // 粗点划线

      DotDotDash,      // 点点划线
      DashDotDotHeavy, // 粗点点划线

      Wave,            // 波浪线
      WavyDouble,      // 双波浪线
      WavyHeavy,       // 粗波浪线
    };


    // Underline 下划线
    struct Underline {
      UnderlineStyle style_ = UnderlineStyle::None;
      std::string color_ = "auto";
    } underline_;


    enum class StrikeStyle {
      None,
      Single, // 删除线
      Double, // 双删除线
    };

    enum class VertAlign {
      None,
      Superscript, // 上标
      Subscript,   // 下标
    };


    // Effects 效果
    struct Effects {

      // Strikethrough 删除线
      StrikeStyle strike_ = StrikeStyle::None;

      // Subscript/Superscript 上标/下标
      VertAlign vertAlign_ = VertAlign::None;
    } effects_;


    // Highlighting 高亮
    enum class Highlight {
      None,
      Black,       // 黑
      White,       // 白
      Red,         // 红
      Green,       // 绿
      Blue,        // 蓝
      Yellow,      // 黄
      Cyan,        // 青
      Magenta,     // 粉
      DarkRed,     // 深红
      DarkGreen,   // 深绿
      DarkBlue,    // 深蓝
      DarkYellow,  // 深黄
      DarkCyan,    // 深青
      DarkMagenta, // 深粉
      DarkGray,    // 深灰
      LightGray,   // 浅灰
    } highlight_ = Highlight::None;


    // Scale 缩放
    // Specifies the amount to stretch or compress each character.
    // Note the minimum value is 1% and the maximum is 600%.
    // 拉伸或压缩字符。范围 1-600。
    size_t scale_ = 100;


    enum class SpacingType {
      Normal,    // 标准
      Expanded,  // 加宽
      Condensed, // 紧缩
    };


    // Spacing 间距
    // Specifies the amount of character pitch to be added or removed after each character.
    // 增加或减少字符间距。
    struct Spacing {

      SpacingType type_ = SpacingType::Normal;

      // 20 = 1 pt, 72 pt = 1 inch
      // 20 = 1 磅，72 磅 = 1 英寸
      size_t by_ = 20;

    } spacing_;


    enum class PositionType {
      Normal,  // 标准
      Raised,  // 上升
      Lowered, // 下降
    };


    // Position 位置
    // Specifies the amount by which text shall be raised or lowered.
    // 提升或降低文本位置。
    struct Position {

      PositionType type_ = PositionType::Normal;

      // 2 = 1 pt, 72 pt = 1 inch
      // 2 = 1 磅，72 磅 = 1 英寸
      size_t by_ = 2;

    } position_;


    // Text Border 文本边框
    std::optional<BorderProperties> border_;


    // True if whitespace needs to be preserved. 是否保留空格。
    bool whitespace_ = false;
  };

}