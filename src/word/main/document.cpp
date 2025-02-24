/**
 * Copyright (C) 2022-2025, Xie Zequn <totravel@foxmail.com>. All rights reserved.
 * Distributed under the MIT License (http://opensource.org/licenses/MIT)
 * Report bugs and download new versions at https://github.com/totravel/minidocx
 */

#include "word/main/document.hpp"
#include "word/main/section.hpp"
#include "word/main/paragraph.hpp"
#include "word/main/richtext.hpp"
#include "word/main/picture.hpp"
#include "word/main/table.hpp"
#include "word/main/cell.hpp"
#include "utils/file.hpp"
#include "utils/string.hpp"
#include "utils/exceptions.hpp"

#include "pugixml.hpp"


namespace MINIDOCX_NAMESPACE
{
  Document::Document()
  {
    init();
  }

  void Document::saveAs(const std::string& filename)
  {
    Zip::open(filename, OpenMode::Create);
    flush();
    Zip::close();
  }

  void Document::load(const std::string& filename)
  {
    Zip::open(filename, OpenMode::ReadOnly);
    clear();
    load();
    Zip::close();
  }


  void Document::init()
  {
    Package::init();
    initContentTypes();
    initRelationships();
  }

  void Document::flush()
  {
    writeOfficeDocument();
    writeStyles();
    writeNumDefinitions();
    Package::flush();
  }

  void Document::clear()
  {
    clearSections();
    Package::clear();
  }

  void Document::load()
  {
    Package::load();
    //readRelationshipsFor(mainPart_);
    //readOfficeDocument();
  }


  void Document::initContentTypes()
  {
    registerOverrideContentType(mainPart_, toPartMediaType(PartType::OfficeDocument));
  }

  void Document::initRelationships()
  {
    addPkgRelationship(PartType::OfficeDocument, mainPart_, Relationship::TargetMode::Internal);
    initRelationshipsFor(mainPart_);
  }


  static void writeParagraphAlignment(pugi::xml_node w_pPr, const Alignment& align)
  {
    switch (align) {
    case Alignment::Left:
      w_pPr.append_child("w:jc").append_attribute("w:val") = "start";
      break;

    case Alignment::Right:
      w_pPr.append_child("w:jc").append_attribute("w:val") = "end";
      break;

    case Alignment::Centered:
      w_pPr.append_child("w:jc").append_attribute("w:val") = "center";
      break;

    case Alignment::Justified:
      w_pPr.append_child("w:jc").append_attribute("w:val") = "both";
      break;

    case Alignment::Distributed:
      w_pPr.append_child("w:jc").append_attribute("w:val") = "distribute";
    }
  }

  static void writeParagraphIndentation(pugi::xml_node w_pPr, const ParagraphProperties::Indentation& indent)
  {
    pugi::xml_node w_ind = w_pPr.append_child("w:ind");

    w_ind.append_attribute(
      indent.left_.chars_ ? "w:leftChars" : "w:left") = indent.left_.value_;

    w_ind.append_attribute(
      indent.right_.chars_ ? "w:rightChars" : "w:right") = indent.right_.value_;

    switch (indent.special_.type_)
    {
    case ParagraphProperties::SpecialIndentationType::FirstLine:
      w_ind.append_attribute(
        indent.special_.chars_ ? "w:firstLineChars" : "w:firstLine") = indent.special_.value_;
      break;

    case ParagraphProperties::SpecialIndentationType::Hanging:
      w_ind.append_attribute(
        indent.special_.chars_ ? "w:hangingChars" : "w:hanging") = indent.special_.value_;
      break;
    }
  }

  static void writeParagraphSpacing(pugi::xml_node w_pPr, const ParagraphProperties::Spacing& spacing)
  {
    pugi::xml_node w_spacing = w_pPr.append_child("w:spacing");

    switch (spacing.before_.type_)
    {
    case ParagraphProperties::SpacingType::Auto:
      w_spacing.append_attribute("w:beforeAutospacing") = "1";
      break;

    case ParagraphProperties::SpacingType::Lines:
      w_spacing.append_attribute("w:beforeAutospacing") = "0";
      w_spacing.append_attribute("w:beforeLines") = spacing.before_.value_;
      break;

    case ParagraphProperties::SpacingType::Absolute:
      w_spacing.append_attribute("w:beforeAutospacing") = "0";
      w_spacing.append_attribute("w:before") = spacing.before_.value_;
      break;
    }

    switch (spacing.after_.type_)
    {
    case ParagraphProperties::SpacingType::Auto:
      w_spacing.append_attribute("w:afterAutospacing") = "1";
      break;

    case ParagraphProperties::SpacingType::Lines:
      w_spacing.append_attribute("w:afterAutospacing") = "0";
      w_spacing.append_attribute("w:afterLines") = spacing.after_.value_;
      break;

    case ParagraphProperties::SpacingType::Absolute:
      w_spacing.append_attribute("w:afterAutospacing") = "0";
      w_spacing.append_attribute("w:after") = spacing.after_.value_;
      break;
    }

    switch (spacing.lineSpacing_.type_)
    {
    case ParagraphProperties::LineSpacingType::Lines:
      w_spacing.append_attribute("w:lineRule") = "auto";
      w_spacing.append_attribute("w:line") = spacing.lineSpacing_.value_;
      break;

    case ParagraphProperties::LineSpacingType::AtLeast:
      w_spacing.append_attribute("w:lineRule") = "atLeast";
      w_spacing.append_attribute("w:line") = spacing.lineSpacing_.value_;
      break;

    case ParagraphProperties::LineSpacingType::Exactly:
      w_spacing.append_attribute("w:lineRule") = "exact";
      w_spacing.append_attribute("w:line") = spacing.lineSpacing_.value_;
      break;
    }
  }

  static void writeBorderProperties(pugi::xml_node w_border, const BorderProperties& prop)
  {
    switch (prop.style_) {
    case BorderStyle::Single:
      w_border.append_attribute("w:val") = "single";
      break;

    case BorderStyle::Double:
      w_border.append_attribute("w:val") = "double";
      break;

    case BorderStyle::Triple:
      w_border.append_attribute("w:val") = "triple";
      break;

    case BorderStyle::Dotted:
      w_border.append_attribute("w:val") = "dotted";
      break;

    case BorderStyle::Dashed:
      w_border.append_attribute("w:val") = "dashed";
      break;

    case BorderStyle::DotDash:
      w_border.append_attribute("w:val") = "dotDash";
      break;

    case BorderStyle::Wave:
      w_border.append_attribute("w:val") = "wave";
      break;

    case BorderStyle::DoubleWave:
      w_border.append_attribute("w:val") = "doubleWave";
      break;
    }

    w_border.append_attribute("w:sz") = prop.width_;
    w_border.append_attribute("w:color") = prop.color_.c_str();
  }

  static void writeParagraphBorders(pugi::xml_node w_pPr, const ParagraphProperties::ParagraphBorders& borders)
  {
    pugi::xml_node w_pBdr = w_pPr.append_child("w:pBdr");

    pugi::xml_node w_top = w_pBdr.append_child("w:top");
    pugi::xml_node w_bottom = w_pBdr.append_child("w:bottom");
    pugi::xml_node w_between = w_pBdr.append_child("w:between");
    pugi::xml_node w_left = w_pBdr.append_child("w:left");
    pugi::xml_node w_right = w_pBdr.append_child("w:right");

    writeBorderProperties(w_top, borders.top_);
    writeBorderProperties(w_bottom, borders.bottom_);
    writeBorderProperties(w_between, borders.bottom_);
    writeBorderProperties(w_left, borders.left_);
    writeBorderProperties(w_right, borders.right_);
  }

  static void writeParagraphProperties(pugi::xml_node w_pPr, const ParagraphProperties& prop)
  {
    if (prop.style_.size() > 0)
      w_pPr.append_child("w:pStyle").append_attribute("w:val") = removeSpaces(prop.style_).c_str();

    if (prop.align_.has_value())
      writeParagraphAlignment(w_pPr, prop.align_.value());

    if (prop.outlineLevel_ != ParagraphProperties::OutlineLevel::BodyText)
      w_pPr.append_child("w:outlineLvl").append_attribute("w:val") = static_cast<unsigned int>(prop.outlineLevel_);

    if (prop.indent_.has_value())
      writeParagraphIndentation(w_pPr, prop.indent_.value());

    if (prop.spacing_.has_value())
      writeParagraphSpacing(w_pPr, prop.spacing_.value());

    if (prop.borders_.has_value())
      writeParagraphBorders(w_pPr, prop.borders_.value());

    if (prop.keepNext_)
      w_pPr.append_child("w:keepNext");

    if (prop.keepLines_)
      w_pPr.append_child("w:keeplines");

    if (prop.pageBreakBefore_)
      w_pPr.append_child("w:pageBreakBefore");
  }

  static void writeRichTextProperties(pugi::xml_node w_rPr, const RichTextProperties& prop)
  {
    if (prop.style_.size() > 0)
      w_rPr.append_child("w:rStyle").append_attribute("w:val") = removeSpaces(prop.style_).c_str();

    if (prop.font_.has_value()) {
      pugi::xml_node w_rFonts = w_rPr.append_child("w:rFonts");
      const RichTextProperties::Font& font = prop.font_.value();

      if (font.ascii_.size() > 0)
        w_rFonts.append_attribute("w:ascii") = font.ascii_.c_str();

      if (font.eastAsia_.size() > 0)
        w_rFonts.append_attribute("w:eastAsia") = font.eastAsia_.c_str();

      if (font.hAnsi_.size() > 0)
        w_rFonts.append_attribute("w:hAnsi") = font.hAnsi_.c_str();

      if (font.cs_.size() > 0)
        w_rFonts.append_attribute("w:cs") = font.cs_.c_str();

      switch (font.hint_)
      {
      case RichTextProperties::FontTypeHint::EastAsia:
        w_rFonts.append_attribute("w:hint") = "eastAsia";
        break;

      case RichTextProperties::FontTypeHint::ComplexScript:
        w_rFonts.append_attribute("w:hint") = "cs";
        break;
      }
    }

    if (prop.fontStyle_.bold_) {
      w_rPr.append_child("w:b");
      w_rPr.append_child("w:bCs");
    }

    if (prop.fontStyle_.italic_) {
      w_rPr.append_child("w:i");
      w_rPr.append_child("w:iCs");
    }

    if (prop.fontSize_ > 0) {
      w_rPr.append_child("w:sz").append_attribute("w:val") = prop.fontSize_;
      w_rPr.append_child("w:szCs").append_attribute("w:val") = prop.fontSize_;
    }

    if (prop.color_.size() > 0) {
      w_rPr.append_child("w:color").
        append_attribute("w:val") = prop.color_.c_str();
    }

    if (prop.underline_.style_ != RichTextProperties::UnderlineStyle::None) {
      pugi::xml_node w_u = w_rPr.append_child("w:u");

      switch (prop.underline_.style_)
      {
      case RichTextProperties::UnderlineStyle::Words:
        w_u.append_attribute("w:val") = "words";
        break;

      case RichTextProperties::UnderlineStyle::Single:
        w_u.append_attribute("w:val") = "single";
        break;

      case RichTextProperties::UnderlineStyle::Double:
        w_u.append_attribute("w:val") = "double";
        break;

      case RichTextProperties::UnderlineStyle::Thick:
        w_u.append_attribute("w:val") = "thick";
        break;

      case RichTextProperties::UnderlineStyle::Dotted:
        w_u.append_attribute("w:val") = "dotted";
        break;

      case RichTextProperties::UnderlineStyle::DottedHeavy:
        w_u.append_attribute("w:val") = "dottedHeavy";
        break;

      case RichTextProperties::UnderlineStyle::Dash:
        w_u.append_attribute("w:val") = "dash";
        break;

      case RichTextProperties::UnderlineStyle::DashedHeavy:
        w_u.append_attribute("w:val") = "dashedHeavy";
        break;

      case RichTextProperties::UnderlineStyle::DashLong:
        w_u.append_attribute("w:val") = "dashLong";
        break;

      case RichTextProperties::UnderlineStyle::DashLongHeavy:
        w_u.append_attribute("w:val") = "dashLongHeavy";
        break;

      case RichTextProperties::UnderlineStyle::DotDash:
        w_u.append_attribute("w:val") = "dotDash";
        break;

      case RichTextProperties::UnderlineStyle::DashDotHeavy:
        w_u.append_attribute("w:val") = "dashDotHeavy";
        break;

      case RichTextProperties::UnderlineStyle::DotDotDash:
        w_u.append_attribute("w:val") = "dotDotDash";
        break;

      case RichTextProperties::UnderlineStyle::DashDotDotHeavy:
        w_u.append_attribute("w:val") = "dashDotDotHeavy";
        break;

      case RichTextProperties::UnderlineStyle::Wave:
        w_u.append_attribute("w:val") = "wave";
        break;

      case RichTextProperties::UnderlineStyle::WavyDouble:
        w_u.append_attribute("w:val") = "wavyDouble";
        break;

      case RichTextProperties::UnderlineStyle::WavyHeavy:
        w_u.append_attribute("w:val") = "wavyHeavy";
        break;
      }

      w_u.append_attribute("w:color") = prop.underline_.color_.c_str();
    }

    switch (prop.effects_.strike_)
    {
    case RichTextProperties::StrikeStyle::Single:
      w_rPr.append_child("w:strike");
      break;
    case RichTextProperties::StrikeStyle::Double:
      w_rPr.append_child("w:dstrike");
      break;
    }

    switch (prop.effects_.vertAlign_)
    {
    case RichTextProperties::VertAlign::Superscript:
      w_rPr.append_child("w:vertAlign").append_attribute("w:val") = "superscript";
      break;
    case RichTextProperties::VertAlign::Subscript:
      w_rPr.append_child("w:vertAlign").append_attribute("w:val") = "subscript";
      break;
    }

    switch (prop.highlight_)
    {
    case RichTextProperties::Highlight::Black:
      w_rPr.append_child("w:highlight").append_attribute("w:val") = "black";
      break;
    case RichTextProperties::Highlight::White:
      w_rPr.append_child("w:highlight").append_attribute("w:val") = "white";
      break;
    case RichTextProperties::Highlight::Red:
      w_rPr.append_child("w:highlight").append_attribute("w:val") = "red";
      break;
    case RichTextProperties::Highlight::Green:
      w_rPr.append_child("w:highlight").append_attribute("w:val") = "green";
      break;
    case RichTextProperties::Highlight::Blue:
      w_rPr.append_child("w:highlight").append_attribute("w:val") = "blue";
      break;
    case RichTextProperties::Highlight::Yellow:
      w_rPr.append_child("w:highlight").append_attribute("w:val") = "yellow";
      break;
    case RichTextProperties::Highlight::Cyan:
      w_rPr.append_child("w:highlight").append_attribute("w:val") = "cyan";
      break;
    case RichTextProperties::Highlight::Magenta:
      w_rPr.append_child("w:highlight").append_attribute("w:val") = "magenta";
      break;
    case RichTextProperties::Highlight::DarkRed:
      w_rPr.append_child("w:highlight").append_attribute("w:val") = "darkRed";
      break;
    case RichTextProperties::Highlight::DarkGreen:
      w_rPr.append_child("w:highlight").append_attribute("w:val") = "darkGreen";
      break;
    case RichTextProperties::Highlight::DarkBlue:
      w_rPr.append_child("w:highlight").append_attribute("w:val") = "darkBlue";
      break;
    case RichTextProperties::Highlight::DarkYellow:
      w_rPr.append_child("w:highlight").append_attribute("w:val") = "darkYellow";
      break;
    case RichTextProperties::Highlight::DarkCyan:
      w_rPr.append_child("w:highlight").append_attribute("w:val") = "darkCyan";
      break;
    case RichTextProperties::Highlight::DarkMagenta:
      w_rPr.append_child("w:highlight").append_attribute("w:val") = "darkMagenta";
      break;
    case RichTextProperties::Highlight::DarkGray:
      w_rPr.append_child("w:highlight").append_attribute("w:val") = "darkGray";
      break;
    case RichTextProperties::Highlight::LightGray:
      w_rPr.append_child("w:highlight").append_attribute("w:val") = "lightGray";
      break;
    }

    if (prop.scale_ != 100)
      w_rPr.append_child("w:w").append_attribute("w:val") = prop.scale_;

    switch (prop.spacing_.type_)
    {
    case RichTextProperties::SpacingType::Normal:
      break;
    case RichTextProperties::SpacingType::Expanded:
      w_rPr.append_child("w:spacing").append_attribute("w:val") = prop.spacing_.by_;
      break;
    case RichTextProperties::SpacingType::Condensed:
      w_rPr.append_child("w:spacing").append_attribute("w:val") = -static_cast<long long>(prop.spacing_.by_);
      break;
    }

    switch (prop.position_.type_)
    {
    case RichTextProperties::PositionType::Normal:
      break;
    case RichTextProperties::PositionType::Raised:
      w_rPr.append_child("w:position").append_attribute("w:val") = prop.position_.by_;
      break;
    case RichTextProperties::PositionType::Lowered:
      w_rPr.append_child("w:position").append_attribute("w:val") = -static_cast<long long>(prop.position_.by_);
      break;
    }

    if (prop.border_.has_value()) {
      pugi::xml_node w_bdr = w_rPr.append_child("w:bdr");
      writeBorderProperties(w_bdr, prop.border_.value());
    }
  }

  static int hasChar(const char ch, const char* list, const size_t len)
  {
    for (int i = 0; i < len; i++)
      if (list[i] == ch)
        return i;
    return -1;
  }

  static void writeText(pugi::xml_node w_r, const char* str, const size_t len, const bool whitespace)
  {
    if (len > 0) {
      const char chars[] = "\n\t\r";
      const char* tags[] = { "w:br", "w:tab", nullptr };
      const size_t num = sizeof(chars) - 1;

      const char* start = str;
      const char* const end = start + len;
      const char* p = start;

      int index = -1;
      while (p < end && -1 == (index = hasChar(*p, chars, num))) {
        p++;
      }
      while (p < end) {
        const int count = p - start;
        if (count > 0) {
          pugi::xml_node w_t = w_r.append_child("w:t");
          if (whitespace)
            w_t.append_attribute("xml:space") = "preserve";
          w_t.text().set(start, count);
        }

        auto tag = tags[index];
        if (tag != nullptr)
          w_r.append_child(tag);

        start = ++p;
        if (start >= end)
          break;

        while (p < end && -1 == (index = hasChar(*p, chars, num))) {
          p++;
        }
      }
      if (start < end) {
        w_r.append_child("w:t").text().set(start);
      }
    }
  }

  static void writeRichText(pugi::xml_node w_p, const RichText& rich)
  {
    pugi::xml_node w_r = w_p.append_child("w:r");
    pugi::xml_node w_rPr = w_r.append_child("w:rPr");

    writeRichTextProperties(w_rPr, rich.prop_);
    writeText(w_r, rich.text().c_str(), rich.text().size(), rich.prop_.whitespace_);
  }

  static size_t pictCount = 0;

  static void writePicture(pugi::xml_node w_p, const Picture& pict)
  {
    const unsigned int index = ++pictCount;
    const std::string name("Picture " + std::to_string(index));

    const PictureProperties& prop = pict.prop_;
    const auto width = prop.extent_.width_;
    const auto height = prop.extent_.height_;

    pugi::xml_node wp_inline = w_p.append_child("w:r")
      .append_child("w:drawing").append_child("wp:inline");
    wp_inline.append_attribute("distT") = "0";
    wp_inline.append_attribute("distB") = "0";
    wp_inline.append_attribute("distL") = "0";
    wp_inline.append_attribute("distR") = "0";

    pugi::xml_node wp_extent = wp_inline.append_child("wp:extent");
    wp_extent.append_attribute("cx") = width;
    wp_extent.append_attribute("cy") = height;

    pugi::xml_node wp_effectExtent = wp_inline.append_child("wp:effectExtent");
    wp_effectExtent.append_attribute("l") = "0";
    wp_effectExtent.append_attribute("t") = "0";
    wp_effectExtent.append_attribute("r") = "0";
    wp_effectExtent.append_attribute("b") = "0";

    pugi::xml_node wp_docPr = wp_inline.append_child("wp:docPr");
    wp_docPr.append_attribute("id") = index;
    wp_docPr.append_attribute("name") = name.c_str();

    pugi::xml_node a_graphicFrameLocks = wp_inline.append_child("wp:cNvGraphicFramePr")
      .append_child("a:graphicFrameLocks");
    a_graphicFrameLocks.append_attribute("xmlns:a")
      .set_value("http://schemas.openxmlformats.org/drawingml/2006/main");
    a_graphicFrameLocks.append_attribute("noChangeAspect") = "1";

    pugi::xml_node a_graphic = wp_inline.append_child("a:graphic");
    pugi::xml_node a_graphicData = a_graphic.append_child("a:graphicData");
    pugi::xml_node pic_pic = a_graphicData.append_child("pic:pic");

    a_graphic.append_attribute("xmlns:a")
      .set_value("http://schemas.openxmlformats.org/drawingml/2006/main");
    a_graphicData.append_attribute("uri")
      .set_value("http://schemas.openxmlformats.org/drawingml/2006/picture");
    pic_pic.append_attribute("xmlns:pic")
      .set_value("http://schemas.openxmlformats.org/drawingml/2006/picture");

    pugi::xml_node pic_nvPicPr = pic_pic.append_child("pic:nvPicPr");
    pugi::xml_node pic_cNvPr = pic_nvPicPr.append_child("pic:cNvPr");
    pic_cNvPr.append_attribute("id") = index;
    pic_cNvPr.append_attribute("name") = name.c_str();
    pugi::xml_node a_picLocks = pic_nvPicPr.append_child("pic:cNvPicPr")
      .append_child("a:picLocks");
    a_picLocks.append_attribute("noChangeAspect") = "1";
    a_picLocks.append_attribute("noChangeArrowheads") = "1";

    pugi::xml_node pic_blipFill = pic_pic.append_child("pic:blipFill");
    pugi::xml_node a_blip = pic_blipFill.append_child("a:blip");
    a_blip.append_attribute("r:embed") = Relationship::stringifyId(pict.id_).c_str();
    pic_blipFill.append_child("a:srcRect"); // Cropping
    pic_blipFill.append_child("a:stretch").append_child("a:fillRect"); // Stretching 

    pugi::xml_node pic_spPr = pic_pic.append_child("pic:spPr");
    pic_spPr.append_attribute("bwMode") = "auto";

    pugi::xml_node a_xfrm = pic_spPr.append_child("a:xfrm");
    pugi::xml_node a_off = a_xfrm.append_child("a:off");
    a_off.append_attribute("x") = "0";
    a_off.append_attribute("y") = "0";
    pugi::xml_node a_ext = a_xfrm.append_child("a:ext");
    a_ext.append_attribute("cx") = width;
    a_ext.append_attribute("cy") = height;

    pugi::xml_node a_prstGeom = pic_spPr.append_child("a:prstGeom");
    a_prstGeom.append_attribute("prst") = "rect";
    a_prstGeom.append_child("a:avLst");

    pic_spPr.append_child("a:noFill");
    pic_spPr.append_child("a:ln").append_child("a:noFill");
  }

  static void writeRun(pugi::xml_node w_p, Run& run)
  {
    switch (run.type())
    {
    case RunType::RichText:
      writeRichText(w_p, dynamic_cast<RichText&>(run));
      break;

    case RunType::Picture:
      writePicture(w_p, dynamic_cast<Picture&>(run));
      break;
    }
  }

  static pugi::xml_node writeParagraph(pugi::xml_node w_body, Paragraph& para)
  {
    pugi::xml_node w_p = w_body.append_child("w:p");
    pugi::xml_node w_pPr = w_p.append_child("w:pPr");

    if (para.numId_ > 0) {
      pugi::xml_node w_numPr = w_pPr.append_child("w:numPr");
      w_numPr.append_child("w:numId").append_attribute("w:val") = para.numId_;
      w_numPr.append_child("w:ilvl").append_attribute("w:val") = static_cast<unsigned int>(para.level_);
    }

    writeParagraphProperties(w_pPr, para.prop_);

    for (auto& run : para.runs())
      writeRun(w_p, *run);

    return w_p;
  }

  static void writeTableBorders(pugi::xml_node w_tblPr, const TableBorders& borders)
  {
    pugi::xml_node w_tblBorders = w_tblPr.append_child("w:tblBorders");

    pugi::xml_node w_top = w_tblBorders.append_child("w:top");
    pugi::xml_node w_bottom = w_tblBorders.append_child("w:bottom");
    pugi::xml_node w_start = w_tblBorders.append_child("w:start");
    pugi::xml_node w_end = w_tblBorders.append_child("w:end");
    pugi::xml_node w_insideH = w_tblBorders.append_child("w:insideH");
    pugi::xml_node w_insideV = w_tblBorders.append_child("w:insideV");

    writeBorderProperties(w_top, borders.top_);
    writeBorderProperties(w_bottom, borders.bottom_);
    writeBorderProperties(w_start, borders.left_);
    writeBorderProperties(w_end, borders.right_);
    writeBorderProperties(w_insideH, borders.insideHorizontal_);
    writeBorderProperties(w_insideV, borders.insideVertical_);
  }

  static void writeTableProperties(pugi::xml_node w_tblPr, const TableProperties& prop)
  {
    if (prop.layout_ == TableProperties::Layout::Fixed)
      w_tblPr.append_child("w:tblLayout").append_attribute("w:type") = "fixed";

    pugi::xml_node w_tblW = w_tblPr.append_child("w:tblW");
    switch (prop.width_.type_)
    {
    case TableProperties::WidthType::Auto:
      w_tblW.append_attribute("w:type") = "auto";
      break;

    case TableProperties::WidthType::Percent:
      w_tblW.append_attribute("w:type") = "pct";
      w_tblW.append_attribute("w:w") = prop.width_.value_;
      break;

    case TableProperties::WidthType::Absolute:
      w_tblW.append_attribute("w:type") = "dxa";
      w_tblW.append_attribute("w:w") = prop.width_.value_;
      break;
    }

    pugi::xml_node w_jc = w_tblPr.append_child("w:jc");
    switch (prop.align_) {
    case TableProperties::Alignment::Left:
      w_jc.append_attribute("w:val") = "start";
      break;

    case TableProperties::Alignment::Right:
      w_jc.append_attribute("w:val") = "end";
      break;

    case TableProperties::Alignment::Center:
      w_jc.append_attribute("w:val") = "center";
      break;
    }

    writeTableBorders(w_tblPr, prop.borders_);
  }

  static pugi::xml_node writeBlock(pugi::xml_node w_body, Block& block);

  static pugi::xml_node writeTable(pugi::xml_node w_body, const Table& tbl)
  {
    pugi::xml_node w_tbl = w_body.append_child("w:tbl");
    pugi::xml_node w_tblPr = w_tbl.append_child("w:tblPr");

    writeTableProperties(w_tblPr, tbl.prop_);

    const Rect& tblRect = tbl.rect();

    pugi::xml_node w_tblGrid = w_tbl.append_child("w:tblGrid");
    for (size_t j = 0; j < tblRect.cols(); j++)
      w_tblGrid.append_child("w:gridCol");

    const size_t colWidth = 5000 / tblRect.cols();

    for (size_t i = 0; i < tblRect.rows(); i++) {
      pugi::xml_node w_tr = w_tbl.append_child("w:tr");

      size_t j = 0;
      while (j < tblRect.cols()) {
        const Cell& cell = *tbl.cellAtUnsafe(i, j);
        const Rect& cellRect = cell.rect();

        pugi::xml_node w_tc = w_tr.append_child("w:tc");
        pugi::xml_node w_tcPr = w_tc.append_child("w:tcPr");

        if (cellRect.cols() > 1)
          w_tcPr.append_child("w:gridSpan").append_attribute("w:val") = cellRect.cols();

        if (i == cellRect.row()) {
          if (cellRect.rows() > 1)
            w_tcPr.append_child("w:vMerge").append_attribute("w:val") = "restart";

          pugi::xml_node w_tcW = w_tcPr.append_child("w:tcW");
          w_tcW.append_attribute("w:type") = "pct";
          w_tcW.append_attribute("w:w") = colWidth * cellRect.cols();

          if (cell.blocks().size() == 0) {
            w_tc.append_child("w:p");
          }
          else {
            for (auto& block : cell.blocks())
              writeBlock(w_tc, *block);
          }
        }
        else {
          if (i == cellRect.rrow())
            w_tcPr.append_child("w:vMerge");
          else
            w_tcPr.append_child("w:vMerge").append_attribute("w:val") = "continue";
          w_tc.append_child("w:p");
        }

        j += cellRect.cols();
      }
    }

    return w_tbl;
  }

  static pugi::xml_node writeBlock(pugi::xml_node w_body, Block& block)
  {
    switch (block.type())
    {
    case BlockType::Paragraph:
      return writeParagraph(w_body, dynamic_cast<Paragraph&>(block));

    case BlockType::Table:
      return writeTable(w_body, dynamic_cast<Table&>(block));

    default:
      return pugi::xml_node();
    }
  }

  static void writeSectionProperties(pugi::xml_node w_pPr, const SectionProperties& prop)
  {
    pugi::xml_node w_sectPr = w_pPr.append_child("w:sectPr");

    pugi::xml_node w_type = w_sectPr.append_child("w:type");
    w_type.append_attribute("w:val") = "nextPage";

    pugi::xml_node w_pgSz = w_sectPr.append_child("w:pgSz");
    if (prop.landscape_) {
      w_pgSz.append_attribute("w:w") = prop.size_.height_;
      w_pgSz.append_attribute("w:h") = prop.size_.width_;
      w_pgSz.append_attribute("w:orient") = "landscape";
    }
    else {
      w_pgSz.append_attribute("w:w") = prop.size_.width_;
      w_pgSz.append_attribute("w:h") = prop.size_.height_;
    }

    pugi::xml_node w_pgMar = w_sectPr.append_child("w:pgMar");
    w_pgMar.append_attribute("w:top") = prop.margins_.top_;
    w_pgMar.append_attribute("w:bottom") = prop.margins_.bottom_;
    w_pgMar.append_attribute("w:left") = prop.margins_.left_;
    w_pgMar.append_attribute("w:right") = prop.margins_.right_;
    w_pgMar.append_attribute("w:header") = prop.margins_.header_;
    w_pgMar.append_attribute("w:footer") = prop.margins_.footer_;
    w_pgMar.append_attribute("w:gutter") = prop.margins_.gutter_;
  }

  void Document::writeOfficeDocument()
  {
    pugi::xml_document doc;
    pugi::xml_node root = doc.append_child("w:document");
    root.append_attribute("xmlns:w")
      .set_value("http://schemas.openxmlformats.org/wordprocessingml/2006/main");
    root.append_attribute("xmlns:r")
      .set_value("http://schemas.openxmlformats.org/officeDocument/2006/relationships");
    root.append_attribute("xmlns:wp")
      .set_value("http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");

    pugi::xml_node body = root.append_child("w:body");

    if (sections_.size() == 0)
      addSection();

    for (auto iter = sections_.begin(); iter != std::prev(sections_.end()); iter++) {
      auto& ptr = *iter;

      if (ptr->blocks().size() == 0)
        ptr->addParagraph();

      pugi::xml_node w_p;
      for (auto& block : ptr->blocks())
        w_p = writeBlock(body, *block);

      if (ptr->blocks().back()->type() == BlockType::Table) {}
      w_p = body.append_child("w:p");

      pugi::xml_node w_pPr = w_p.child("w:pPr");
      if (!w_pPr)
        w_pPr = w_p.append_child("w:pPr");

      writeSectionProperties(w_pPr, ptr->prop_);
    }

    auto& ptr = sections_.back();
    if (ptr->blocks().size() == 0)
      ptr->addParagraph();

    for (auto& block : ptr->blocks())
      writeBlock(body, *block);

    writeSectionProperties(body, ptr->prop_);

    writePart(mainPart_, doc);
    pictCount = 0;
  }


  SectionPointer Document::addSection()
  {
    auto section{ std::make_shared<Section>() };
    sections_.push_back(section);
    return section;
  }

  void Document::deleteSection(const SectionPointer& section)
  {
    section->destroy();
    sections_.remove(section);
  }

  void Document::clearSections()
  {
    for (auto& section : sections_)
      section->destroy();
    sections_.clear();
  }


  RelationshipId Document::addImage(Buffer buf, const FileType type)
  {
    registerDefaultContentType(getFileExtension(type) + 1, getFileMediaType(type));

    const size_t count = ++imgCount_;
    const PartName name(IMG_PREFIX + std::to_string(count) + getFileExtension(type));

    addBufferedPart(name, std::move(buf));

    return addRelationshipFor(mainPart_, PartType::Image, name, Relationship::TargetMode::Internal);
  }

  RelationshipId Document::addImage(const FileName& filename)
  {
    const FileType type = getFileType(filename);
    if (type == FileType::Unknown)
      throw invalid_parameter();
    
    Buffer buf;
    readFile(buf, filename);
    return addImage(std::move(buf), type);
  }


  void Document::writeStyles()
  {
    if (paragraphStyles_.size() + characterStyles_.size() == 0)
      return;
    registerOverrideContentType(stylePart_, toPartMediaType(PartType::Style));
    addRelationshipFor(mainPart_, PartType::Style, stylePart_, Relationship::TargetMode::Internal);

    pugi::xml_document doc;
    pugi::xml_node root = doc.append_child("w:styles");
    root.append_attribute("xmlns:w")
      .set_value(toPartRootNamespace(PartType::Style));

    for (const auto& ref : paragraphStyles_) {
      pugi::xml_node w_style = root.append_child("w:style");
      w_style.append_attribute("w:type") = "paragraph";
      w_style.append_attribute("w:styleId") = ref.first.c_str();
      w_style.append_child("w:name").append_attribute("w:val") = ref.second.name_.c_str();
      w_style.append_child("w:qFormat"); // showed in the styles gallery

      if (ref.second.next_ != "")
        w_style.append_child("w:next").append_attribute("w:val") = removeSpaces(ref.second.next_).c_str();

      if (ref.second.basedOn_ != "")
        w_style.append_child("w:basedOn").append_attribute("w:val") = removeSpaces(ref.second.basedOn_).c_str();

      writeParagraphProperties(w_style.append_child("w:pPr"), ref.second);
      writeRichTextProperties(w_style.append_child("w:rPr"), ref.second);
    }

    for (const auto& ref : characterStyles_) {
      pugi::xml_node w_style = root.append_child("w:style");
      w_style.append_attribute("w:type") = "character";
      w_style.append_attribute("w:styleId") = ref.first.c_str();
      w_style.append_child("w:name").append_attribute("w:val") = ref.second.name_.c_str();
      w_style.append_child("w:qFormat");

      if (ref.second.basedOn_ != "")
        w_style.append_child("w:basedOn").append_attribute("w:val") = removeSpaces(ref.second.basedOn_).c_str();

      writeRichTextProperties(w_style.append_child("w:rPr"), ref.second);
    }

    writePart(stylePart_, doc);
  }

  void Document::addParagraphStyle(const ParagraphStyle& style)
  {
    paragraphStyles_[removeSpaces(style.name_)] = style;
  }

  void Document::addCharacterStyle(const CharacterStyle& style)
  {
    characterStyles_[removeSpaces(style.name_)] = style;
  }


  void Document::writeNumDefinitions()
  {
    if (abstractNumDefinitions_.size() == 0)
      return;
    registerOverrideContentType(numPart_, toPartMediaType(PartType::Numbering));
    addRelationshipFor(mainPart_, PartType::Numbering, numPart_, Relationship::TargetMode::Internal);

    pugi::xml_document doc;
    pugi::xml_node root = doc.append_child("w:numbering");
    root.append_attribute("xmlns:w")
      .set_value(toPartRootNamespace(PartType::Numbering));

    for (const auto& numDef : abstractNumDefinitions_) {
      pugi::xml_node w_abstractNum = root.append_child("w:abstractNum");
      w_abstractNum.append_attribute("w:abstractNumId") = numDef.first;

      switch (numDef.second.type_)
      {
      case NumberingType::SingLevel:
        w_abstractNum.append_child("w:multiLevelType").append_attribute("w:val") = "singlelevel";
        break;

      case NumberingType::MultiLevel:
        w_abstractNum.append_child("w:multiLevelType").append_attribute("w:val") = "multilevel";
        break;

      case NumberingType::HybridMultiLevel:
        w_abstractNum.append_child("w:multiLevelType").append_attribute("w:val") = "hybridMultilevel";
        break;

      default:
        w_abstractNum.append_child("w:multiLevelType").append_attribute("w:val") = "hybridMultilevel";
        break;
      }

      size_t count = 0;
      for (const auto& lvl : numDef.second.levels_) {
        pugi::xml_node w_lvl = w_abstractNum.append_child("w:lvl");
        w_lvl.append_attribute("w:ilvl") = count++;

        w_lvl.append_child("w:start").append_attribute("w:val") = lvl.numStart_;

        switch (lvl.numStyle_)
        {
        case NumberStyle::Decimal:
          w_lvl.append_child("w:numFmt").append_attribute("w:val") = "decimal";
          break;

        case NumberStyle::LowerLetter:
          w_lvl.append_child("w:numFmt").append_attribute("w:val") = "lowerLetter";
          break;

        case NumberStyle::LowerRoman:
          w_lvl.append_child("w:numFmt").append_attribute("w:val") = "lowerRoman";
          break;

        default:
          w_lvl.append_child("w:numFmt").append_attribute("w:val") = "bullet";
          break;
        }
        
        w_lvl.append_child("w:lvlText").append_attribute("w:val") = lvl.numFmt_.c_str();

        switch (lvl.numAlign_) {
        case Alignment::Left:
          w_lvl.append_child("w:lvlJc").append_attribute("w:val") = "left";
          break;

        case Alignment::Right:
          w_lvl.append_child("w:lvlJc").append_attribute("w:val") = "right";
          break;

        case Alignment::Centered:
          w_lvl.append_child("w:lvlJc").append_attribute("w:val") = "center";
          break;

        case Alignment::Justified:
          w_lvl.append_child("w:lvlJc").append_attribute("w:val") = "both";
          break;

        case Alignment::Distributed:
          w_lvl.append_child("w:lvlJc").append_attribute("w:val") = "distribute";
        }
        
        writeParagraphProperties(w_lvl.append_child("w:pPr"), lvl);
        writeRichTextProperties(w_lvl.append_child("w:rPr"), lvl);
      }
    }

    for (const auto& numDef : numDefinitions_) {
      pugi::xml_node w_num = root.append_child("w:num");
      w_num.append_attribute("w:numId") = numDef.first;
      w_num.append_child("w:abstractNumId").append_attribute("w:val") = numDef.second.id_;
      // TODO: writes numbering level overrides
    }

    writePart(numPart_, doc);
  }
}
