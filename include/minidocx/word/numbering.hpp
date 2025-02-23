/**
 * Copyright (C) 2022-2025, Xie Zequn <totravel@foxmail.com>. All rights reserved.
 * Distributed under the MIT License (http://opensource.org/licenses/MIT)
 * Report bugs and download new versions at https://github.com/totravel/minidocx
 */

#pragma once

#include "word/main/properties/paragraph.hpp"
#include "word/main/properties/richtext.hpp"

#include <string>
#include <array>
#include <map>


namespace MINIDOCX_NAMESPACE
{
  using NumberingId = size_t;

  enum class NumberingLevel: unsigned int
  {
    Level1, Level2, Level3, Level4, Level5, Level6, Level7, Level8, Level9
  };

  enum class NumberStyle
  {
    Bullet,      // Bullets 项目符号
    Decimal,     // 1, 2, 3, etc.
    UpperRoman,  // I, II, III, IV, etc.
    LowerRoman,  // i, ii, iii, iv, etc.
    UpperLetter, // A, B, C, etc.
    LowerLetter, // a, b, c, etc.
    OrdinalText, // First, Second, Third, etc.
    CardinalText // One, Two, Three, etc.
  };

  struct LevelDefinition : ParagraphProperties, RichTextProperties
  {
    size_t numStart_ = 1;
    NumberStyle numStyle_ = NumberStyle::Decimal;
    std::string numFmt_ = "%1.";
    Alignment numAlign_ = Alignment::Left;
  };

  enum class NumberingType
  {
    SingLevel, MultiLevel, HybridMultiLevel
  };

  const char * const CHAR_SOLID_CIRCLE = reinterpret_cast<const char *>(u8"\uF0B7");
  const char * const CHAR_SOLID_SQUARE = reinterpret_cast<const char *>(u8"\uF0A7");

  struct AbstractNumberingDefinition
  {
    NumberingType type_ = NumberingType::HybridMultiLevel;
    std::array<LevelDefinition, 9> levels_;

    // Default numbered list
    static AbstractNumberingDefinition makeNumberedList()
    {
      AbstractNumberingDefinition numDef;

      ParagraphProperties::Indentation indent;
      indent.special_.type_ = ParagraphProperties::SpecialIndentationType::Hanging;
      indent.special_.chars_ = false;
      indent.special_.value_ = 360;

      indent.left_.value_ = 720;
      numDef.levels_[0].indent_ = indent;
      numDef.levels_[0].numFmt_ = "%1.";
      numDef.levels_[0].numStyle_ = NumberStyle::Decimal;

      indent.left_.value_ = 1440;
      numDef.levels_[1].indent_ = indent;
      numDef.levels_[1].numFmt_ = "%2.";
      numDef.levels_[1].numStyle_ = NumberStyle::LowerLetter;

      indent.left_.value_ = 2160;
      numDef.levels_[2].indent_ = indent;
      numDef.levels_[2].numFmt_ = "%3.";
      numDef.levels_[2].numStyle_ = NumberStyle::LowerRoman;
      numDef.levels_[2].numAlign_ = Alignment::Right;


      indent.left_.value_ = 2880;
      numDef.levels_[3].indent_ = indent;
      numDef.levels_[3].numFmt_ = "%4.";
      numDef.levels_[3].numStyle_ = NumberStyle::Decimal;

      indent.left_.value_ = 3600;
      numDef.levels_[4].indent_ = indent;
      numDef.levels_[4].numFmt_ = "%5.";
      numDef.levels_[4].numStyle_ = NumberStyle::LowerLetter;

      indent.left_.value_ = 4320;
      numDef.levels_[5].indent_ = indent;
      numDef.levels_[5].numFmt_ = "%6.";
      numDef.levels_[5].numStyle_ = NumberStyle::LowerRoman;
      numDef.levels_[5].numAlign_ = Alignment::Right;


      indent.left_.value_ = 5040;
      numDef.levels_[6].indent_ = indent;
      numDef.levels_[6].numFmt_ = "%7.";
      numDef.levels_[6].numStyle_ = NumberStyle::Decimal;

      indent.left_.value_ = 5760;
      numDef.levels_[7].indent_ = indent;
      numDef.levels_[7].numFmt_ = "%8.";
      numDef.levels_[7].numStyle_ = NumberStyle::LowerLetter;

      indent.left_.value_ = 6480;
      numDef.levels_[8].indent_ = indent;
      numDef.levels_[8].numFmt_ = "%9.";
      numDef.levels_[8].numStyle_ = NumberStyle::LowerRoman;
      numDef.levels_[8].numAlign_ = Alignment::Right;

      return numDef;
    };

    // Default bulleted list
    static AbstractNumberingDefinition makeBulletedList()
    {
      AbstractNumberingDefinition numDef;

      ParagraphProperties::Indentation indent;
      indent.special_.type_ = ParagraphProperties::SpecialIndentationType::Hanging;
      indent.special_.chars_ = false;
      indent.special_.value_ = 360;

      RichTextProperties::Font font;
      font.eastAsia_ = "";
      font.cs_ = "";

      indent.left_.value_ = 720;
      font.ascii_ = "Symbol";
      font.hAnsi_ = "Symbol";
      numDef.levels_[0].indent_ = indent;
      numDef.levels_[0].font_ = font;
      numDef.levels_[0].numFmt_ = CHAR_SOLID_CIRCLE;
      numDef.levels_[0].numStyle_ = NumberStyle::Bullet;

      indent.left_.value_ = 1440;
      font.ascii_ = "Courier New";
      font.hAnsi_ = "Courier New";
      numDef.levels_[1].indent_ = indent;
      numDef.levels_[1].font_ = font;
      numDef.levels_[1].numFmt_ = "o";
      numDef.levels_[1].numStyle_ = NumberStyle::Bullet;

      indent.left_.value_ = 2160;
      font.ascii_ = "Wingdings";
      font.hAnsi_ = "Wingdings";
      numDef.levels_[2].indent_ = indent;
      numDef.levels_[2].font_ = font;
      numDef.levels_[2].numFmt_ = CHAR_SOLID_SQUARE;
      numDef.levels_[2].numStyle_ = NumberStyle::Bullet;


      indent.left_.value_ = 2880;
      font.ascii_ = "Symbol";
      font.hAnsi_ = "Symbol";
      numDef.levels_[3].indent_ = indent;
      numDef.levels_[3].font_ = font;
      numDef.levels_[3].numFmt_ = CHAR_SOLID_CIRCLE;
      numDef.levels_[3].numStyle_ = NumberStyle::Bullet;

      indent.left_.value_ = 3600;
      font.ascii_ = "Courier New";
      font.hAnsi_ = "Courier New";
      numDef.levels_[4].indent_ = indent;
      numDef.levels_[4].font_ = font;
      numDef.levels_[4].numFmt_ = "o";
      numDef.levels_[4].numStyle_ = NumberStyle::Bullet;

      indent.left_.value_ = 4320;
      font.ascii_ = "Wingdings";
      font.hAnsi_ = "Wingdings";
      numDef.levels_[5].indent_ = indent;
      numDef.levels_[5].font_ = font;
      numDef.levels_[5].numFmt_ = CHAR_SOLID_SQUARE;
      numDef.levels_[5].numStyle_ = NumberStyle::Bullet;


      indent.left_.value_ = 5040;
      font.ascii_ = "Symbol";
      font.hAnsi_ = "Symbol";
      numDef.levels_[6].indent_ = indent;
      numDef.levels_[6].font_ = font;
      numDef.levels_[6].numFmt_ = CHAR_SOLID_CIRCLE;
      numDef.levels_[6].numStyle_ = NumberStyle::Bullet;

      indent.left_.value_ = 5760;
      font.ascii_ = "Courier New";
      font.hAnsi_ = "Courier New";
      numDef.levels_[7].indent_ = indent;
      numDef.levels_[7].font_ = font;
      numDef.levels_[7].numFmt_ = "o";
      numDef.levels_[7].numStyle_ = NumberStyle::Bullet;

      indent.left_.value_ = 6480;
      font.ascii_ = "Wingdings";
      font.hAnsi_ = "Wingdings";
      numDef.levels_[8].indent_ = indent;
      numDef.levels_[8].font_ = font;
      numDef.levels_[8].numFmt_ = CHAR_SOLID_SQUARE;
      numDef.levels_[8].numStyle_ = NumberStyle::Bullet;

      return numDef;
    }
  };

  struct NumberingDefinition
  {
    NumberingDefinition() = default;
    NumberingDefinition(const NumberingId id) : id_{ id } {}

    NumberingId id_ = 0; // Abstract Numbering Definition ID
    std::map<NumberingLevel, LevelDefinition> levelOverrides_;
  };

}
