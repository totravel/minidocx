
#include <cstring>
#include <cmath>
#include "minidocx.hpp"
#include "zip.h"

// template of parts (a .xml file) of a package (a .docx/.zip file) 
// used to create a new package
#define _RELS R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>)"
#define DOCUMENT_XML R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex" xmlns:cx1="http://schemas.microsoft.com/office/drawing/2015/9/8/chartex" xmlns:cx2="http://schemas.microsoft.com/office/drawing/2015/10/21/chartex" xmlns:cx3="http://schemas.microsoft.com/office/drawing/2016/5/9/chartex" xmlns:cx4="http://schemas.microsoft.com/office/drawing/2016/5/10/chartex" xmlns:cx5="http://schemas.microsoft.com/office/drawing/2016/5/11/chartex" xmlns:cx6="http://schemas.microsoft.com/office/drawing/2016/5/12/chartex" xmlns:cx7="http://schemas.microsoft.com/office/drawing/2016/5/13/chartex" xmlns:cx8="http://schemas.microsoft.com/office/drawing/2016/5/14/chartex" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:aink="http://schemas.microsoft.com/office/drawing/2016/ink" xmlns:am3d="http://schemas.microsoft.com/office/drawing/2017/model3d" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:oel="http://schemas.microsoft.com/office/2019/extlst" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" xmlns:w16cex="http://schemas.microsoft.com/office/word/2018/wordml/cex" xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid" xmlns:w16="http://schemas.microsoft.com/office/word/2018/wordml" xmlns:w16sdtdh="http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash" xmlns:w16se="http://schemas.microsoft.com/office/word/2015/wordml/symex" xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk" xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" mc:Ignorable="w14 w15 w16se w16cid w16 w16cex w16sdtdh wp14"><w:body><w:sectPr><w:pgSz w:w="11906" w:h="16838" /><w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800" w:header="851" w:footer="992" w:gutter="0" /><w:cols w:space="425" /><w:docGrid w:type="lines" w:linePitch="312" /></w:sectPr></w:body></w:document>)"
#define CONTENT_TYPES_XML R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>)"

namespace docx
{
  void GetCharLen(std::string text, int &ascii, int &eastAsia)
  {
    ascii = eastAsia = 0;
    for (int i = 0; i < text.length(); i++) {
      if (text[i] < 0) {
        i += 2;
        eastAsia++;
        continue;
      }
      ascii++;
    }
  }

  int GetCharLen(std::string s)
  {
    int c = 0;
    for (int i = 0; i < s.length(); i++) {
      if (s[i] < 0) i += 2;
      c++;
    }
    return c;
  }

  int Pt2Twip(const double pt)
  {
    return pt * 20;
  }

  double Twip2Pt(const int twip)
  {
    return twip / 20.0;
  }

  double Inch2Pt(const double inch)
  {
    return inch * 72;
  }

  double Pt2Inch(const double pt)
  {
    return pt / 72;
  }

  double MM2Inch(const int mm)
  {
    return mm / 25.4;
  }

  int Inch2MM(const double inch)
  {
    return inch * 25.4;
  }

  double CM2Inch(const double cm)
  {
    return cm / 2.54;
  }

  double Inch2CM(const double inch)
  {
    return inch * 2.54;
  }

  int Inch2Twip(const double inch)
  {
    return inch * 1440;
  }

  double Twip2Inch(const int twip)
  {
    return twip / 1440.0;
  }

  int MM2Twip(const int mm)
  {
    return Inch2Twip(MM2Inch(mm));
  }

  int Twip2MM(const int twip)
  {
    return Inch2MM(Twip2Inch(twip));
  }

  int CM2Twip(const double cm)
  {
    return Inch2Twip(CM2Inch(cm));
  }

  double Twip2CM(const int twip)
  {
    return Inch2CM(Twip2Inch(twip));
  }

  struct xml_string_writer: pugi::xml_writer
  {
    std::string result;

    virtual void write(const void *data, size_t size)
    {
      result.append(static_cast<const char *>(data), size);
    }
  };

  pugi::xml_node GetLastChild(pugi::xml_node node, const char *name)
  {
    pugi::xml_node child = node.last_child();
    while (!child.empty() && strcmp(name, child.name()) != 0) {
      child = child.previous_sibling(name);
    }
    return child;
  }

  // class Document
  Document::Document(const std::string path): path_(path)
  {
    doc_.load_buffer(DOCUMENT_XML, strlen(DOCUMENT_XML));
    body_ = doc_.child("w:document").child("w:body");
  }

  void Document::Save()
  {
    // <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    pugi::xml_node decl = doc_.prepend_child(pugi::node_declaration);
    decl.append_attribute("version")    = "1.0";
    decl.append_attribute("encoding")   = "UTF-8";
    decl.append_attribute("standalone") = "yes";

    xml_string_writer writer;
    doc_.save(writer, "", pugi::format_raw);
    const char *buf = writer.result.c_str();

    struct zip_t *zip = zip_open(path_.c_str(), ZIP_DEFAULT_COMPRESSION_LEVEL, 'w');

    zip_entry_open(zip, "_rels/.rels");
    zip_entry_write(zip, _RELS, strlen(_RELS));
    zip_entry_close(zip);
    
    zip_entry_open(zip, "word/document.xml");
    zip_entry_write(zip, buf, strlen(buf));
    zip_entry_close(zip);
    
    zip_entry_open(zip, "[Content_Types].xml");
    zip_entry_write(zip, CONTENT_TYPES_XML, strlen(CONTENT_TYPES_XML));
    zip_entry_close(zip);

    zip_close(zip);
  }

  std::string Document::GetFormatedBody() {
    xml_string_writer writer;
    body_.print(writer, " ");
    return writer.result;
  }

  Paragraph Document::FirstParagraph()
  {
    auto p = body_.child("w:p");
    auto pPr = p.child("w:pPr");
    return Paragraph(body_, p, pPr);
  }

  Paragraph Document::LastParagraph()
  {
    auto p = GetLastChild(body_, "w:p");
    auto pPr = p.child("w:pPr");
    return Paragraph(body_, p, pPr);
  }

  Section Document::FirstSection()
  {
    return FirstParagraph().GetSection();
  }
  Section Document::LastSection()
  {
    return LastParagraph().GetSection();
  }

  Paragraph Document::AppendParagraph()
  {
    auto p = body_.append_child("w:p");
    auto pPr = p.append_child("w:pPr");
    return Paragraph(body_, p, pPr);
  }

  Paragraph Document::AppendParagraph(const std::string text)
  {
    auto p = AppendParagraph();
    p.AppendRun(text);
    return p;
  }

  Paragraph Document::AppendParagraph(const std::string text, 
                                      const double fontSize)
  {
    auto p = AppendParagraph();
    p.AppendRun(text, fontSize);
    return p;
  }

  Paragraph Document::AppendParagraph(const std::string text, 
                                      const double fontSize, 
                                      const std::string fontAscii, 
                                      const std::string fontEastAsia)
  {
    auto p = AppendParagraph();
    p.AppendRun(text, fontSize, fontAscii, fontEastAsia);
    return p;
  }

  Paragraph Document::PrependParagraph()
  {
    auto p = body_.prepend_child("w:p");
    auto pPr = p.append_child("w:pPr");
    return Paragraph(body_, p, pPr);
  }

  Paragraph Document::PrependParagraph(const std::string text)
  {
    auto p = PrependParagraph();
    p.AppendRun(text);
    return p;
  }

  Paragraph Document::PrependParagraph(const std::string text, 
                                       const double fontSize)
  {
    auto p = PrependParagraph();
    p.AppendRun(text, fontSize);
    return p;
  }

  Paragraph Document::PrependParagraph(const std::string text, 
                                       const double fontSize, 
                                       const std::string fontAscii, 
                                       const std::string fontEastAsia)
  {
    auto p = PrependParagraph();
    p.AppendRun(text, fontSize, fontAscii, fontEastAsia);
    return p;
  }

  Paragraph Document::AppendPageBreak()
  {
    auto p = AppendParagraph();
    p.AppendPageBreak();
    return p;
  }

  Paragraph Document::AppendSectionBreak()
  {
    auto p = AppendParagraph();
    p.InsertSectionBreak();
    return p;
  }


  // class Paragraph
  Run Paragraph::FirstRun()
  {
    auto r = p_.child("w:r");
    auto rPr = r.child("w:rPr");
    return Run(p_, r, rPr);
  }

  Run Paragraph::AppendRun()
  {
    auto r = p_.append_child("w:r");
    auto rPr = r.append_child("w:rPr");
    return Run(p_, r, rPr);
  }

  Run Paragraph::AppendRun(const std::string text)
  {
    auto r = AppendRun();
    if (!text.empty()) {
      r.AppendText(text);
    }
    return r;
  }

  Run Paragraph::AppendRun(const std::string text, 
                           const double fontSize)
  {
    auto r = AppendRun(text);
    if (fontSize != 0) {
      r.SetFontSize(fontSize);
    }
    return r;
  }

  Run Paragraph::AppendRun(const std::string text, 
                           const double fontSize, 
                           const std::string fontAscii, 
                           const std::string fontEastAsia)
  {
    auto r = AppendRun(text, fontSize);
    if (!fontAscii.empty()) {
      r.SetFont(fontAscii, fontEastAsia);
    }
    return r;
  }

  Run Paragraph::AppendPageBreak()
  {
    auto r = p_.append_child("w:r");
    auto br = r.append_child("w:br");
    br.append_attribute("w:type") = "page";
    return Run(p_, r, br);
  }

  void Paragraph::SetAlignment(Alignment alignment)
  {
    const char *val;
    switch (alignment) {
      case Alignment::Left:
        val = "start";
        break;
      case Alignment::Right:
        val = "end";
        break;
      case Alignment::Centered:
        val = "center";
        break;
      case Alignment::Justified:
        val = "both";
        break;
      case Alignment::Distributed:
        val = "distribute";
        break;
    }

    auto jc = pPr_.child("w:jc");
    if (!jc) {
      jc = pPr_.append_child("w:jc");
    }
    auto jcVal = jc.attribute("w:val");
    if (!jcVal) {
      jcVal = jc.append_attribute("w:val");
    }
    jcVal.set_value(val);
  }

  void Paragraph::SetLineSpacingSingle()
  {
    auto spacing = pPr_.child("w:spacing");
    if (!spacing) return;
    auto spacingLineRule = spacing.attribute("w:lineRule");
    if (spacingLineRule) {
      spacing.remove_attribute(spacingLineRule);
    }
    auto spacingLine = spacing.attribute("w:line");
    if (spacingLine) {
      spacing.remove_attribute(spacingLine);
    }
  }

  void Paragraph::SetLineSpacingLines(double at)
  {
    // A normal single-spaced paragaph has a w:line value of 240, or 12 points.
    // 
    // If the value of lineRule is auto, then the value of line 
    // is interpreted as 240th of a line, e.g. 360 = 1.5 lines.
    SetLineSpacing(at * 240, "auto");
  }

  void Paragraph::SetLineSpacingAtLeast(int at)
  {
    // If the value of the lineRule attribute is atLeast or exactly, 
    // then the value of the line attribute is interpreted as 240th of a point.
    // (Not really. Actually, values are in twentieths of a point, e.g. 240 = 12 pt.)
    SetLineSpacing(at, "atLeast");
  }

  void Paragraph::SetLineSpacingExactly(int at)
  {
    SetLineSpacing(at, "exact");
  }

  void Paragraph::SetLineSpacing(int at, const char *lineRule)
  {
    auto spacing = pPr_.child("w:spacing");
    if (!spacing) {
      spacing = pPr_.append_child("w:spacing");
    }

    auto spacingLineRule = spacing.attribute("w:lineRule");
    if (!spacingLineRule) {
      spacingLineRule = spacing.append_attribute("w:lineRule");
    }

    auto spacingLine = spacing.attribute("w:line");
    if (!spacingLine) {
      spacingLine = spacing.append_attribute("w:line");
    }

    spacingLineRule.set_value(lineRule);
    spacingLine.set_value(at);
  }

  void Paragraph::SetBeforeSpacingAuto()
  {
    SetSpacingAuto("w:beforeAutospacing");
  }

  void Paragraph::SetAfterSpacingAuto()
  {
    SetSpacingAuto("w:afterAutospacing");
  }

  void Paragraph::SetSpacingAuto(const char *attrNameAuto)
  {
    auto spacing = pPr_.child("w:spacing");
    if (!spacing) {
      spacing = pPr_.append_child("w:spacing");
    }
    auto spacingAuto = spacing.attribute(attrNameAuto);
    if (!spacingAuto) {
      spacingAuto = spacing.append_attribute(attrNameAuto);
    }
    // Any value for before or beforeLines is ignored.
    spacingAuto.set_value("true");
  }

  void Paragraph::SetBeforeSpacingLines(double beforeSpacing)
  {
    // To specify units in hundreths of a line, 
    // use attributes 'afterLines'/'beforeLines'.
    SetSpacing(beforeSpacing * 100, "w:beforeAutospacing", "w:beforeLines");
  }

  void Paragraph::SetAfterSpacingLines(double afterSpacing)
  {
    SetSpacing(afterSpacing * 100, "w:afterAutospacing", "w:afterLines");
  }

  void Paragraph::SetBeforeSpacing(int beforeSpacing)
  {
    SetSpacing(beforeSpacing, "w:beforeAutospacing", "w:before");
  }

  void Paragraph::SetAfterSpacing(int afterSpacing)
  {
    SetSpacing(afterSpacing, "w:afterAutospacing", "w:after");
  }

  void Paragraph::SetSpacing(int twip, const char *attrNameAuto, const char *attrName)
  {
    auto elemSpacing = pPr_.child("w:spacing");
    if (!elemSpacing) {
      elemSpacing = pPr_.append_child("w:spacing");
    }

    auto attrSpacingAuto = elemSpacing.attribute(attrNameAuto);
    if (attrSpacingAuto) {
      elemSpacing.remove_attribute(attrSpacingAuto);
    }

    auto attrSpacing = elemSpacing.attribute(attrName);
    if (!attrSpacing) {
      attrSpacing = elemSpacing.append_attribute(attrName);
    }
    attrSpacing.set_value(twip);
  }

  void Paragraph::SetLeftIndentChars(double leftIndent)
  {
    // To specify units in hundreths of a character, 
    // use attributes leftChars/endChars, rightChars/endChars, etc. 
    SetIndent(leftIndent * 100, "w:leftChars");
  }

  void Paragraph::SetRightIndentChars(double rightIndent)
  {
    SetIndent(rightIndent * 100, "w:rightChars");
  }

  void Paragraph::SetLeftIndent(int leftIndent)
  {
    SetIndent(leftIndent, "w:left");
  }

  void Paragraph::SetRightIndent(int rightIndent)
  {
    SetIndent(rightIndent, "w:right");
  }

  void Paragraph::SetFirstLineChars(double indent)
  {
    SetIndent(indent * 100, "w:firstLineChars");
  }

  void Paragraph::SetHangingChars(double indent)
  {
    SetIndent(indent * 100, "w:hangingChars");
  }

  void Paragraph::SetFirstLine(int indent)
  {
    SetIndent(indent, "w:firstLine");
  }

  void Paragraph::SetHanging(int indent)
  {
    SetIndent(indent, "w:hanging");
    SetLeftIndent(indent);
  }

  void Paragraph::SetIndent(int indent, const char *attrName)
  {
    auto elemIndent = pPr_.child("w:ind");
    if (!elemIndent) {
      elemIndent = pPr_.append_child("w:ind");
    }

    auto attrIndent = elemIndent.attribute(attrName);
    if (!attrIndent) {
      attrIndent = elemIndent.append_attribute(attrName);
    }
    attrIndent.set_value(indent);
  }

  void Paragraph::SetFontSize(const double fontSize)
  {
    for (auto r = FirstRun(); r; r = r.Next()) {
      r.SetFontSize(fontSize);
    }
  }

  void Paragraph::SetFont(const std::string fontAscii, 
                          const std::string fontEastAsia)
  {
    for (auto r = FirstRun(); r; r = r.Next()) {
      r.SetFont(fontAscii, fontEastAsia);
    }
  }

  void Paragraph::SetFontStyle(Run::FontStyle fontStyle)
  {
    for (auto r = FirstRun(); r; r = r.Next()) {
      r.SetFontStyle(fontStyle);
    }
  }

  void Paragraph::SetCharacterSpacing(const int characterSpacing)
  {
    for (auto r = FirstRun(); r; r = r.Next()) {
      r.SetCharacterSpacing(characterSpacing);
    }
  }

  std::string Paragraph::GetText()
  {
    std::string text;
    for (auto r = FirstRun(); r; r = r.Next()) {
      text += r.GetText();
    }
    return text;
  }

  Paragraph Paragraph::InsertBefore()
  {
    auto p = body_.insert_child_before("w:p", p_);
    auto pPr = p.append_child("w:pPr");
    return Paragraph(body_, p, pPr);
  }

  Paragraph Paragraph::InsertAfter()
  {
    auto p = body_.insert_child_after("w:p", p_);
    auto pPr = p.append_child("w:pPr");
    return Paragraph(body_, p, pPr);
  }

  void Paragraph::Remove()
  {
    body_.remove_child(p_);
  }

  Paragraph Paragraph::Next()
  {
    auto p = p_.next_sibling("w:p");
    auto pPr = p.child("w:pPr");
    return Paragraph(body_, p, pPr);
  }

  Paragraph Paragraph::Prev()
  {
    auto p = p_.previous_sibling("w:p");
    auto pPr = p.child("w:pPr");
    return Paragraph(body_, p, pPr);
  }

  Paragraph::operator bool()
  {
    return p_;
  }

  Section Paragraph::GetSection()
  {
    return Section(body_, p_, pPr_);
  }

  Section Paragraph::InsertSectionBreak()
  {
    auto s = Section(body_, p_, pPr_);
    // this paragraph will be the last paragraph of the new section
    s.Split();
    return s;
  }

  Section Paragraph::RemoveSectionBreak()
  {
    auto s = Section(body_, p_, pPr_);
    if (s.IsSplit()) s.Merge();
    return s;
  }

  bool Paragraph::HasSectionBreak()
  {
    return GetSection().IsSplit();
  }


  // class Section
  Section::Section(pugi::xml_node body, 
                   pugi::xml_node p, 
                   pugi::xml_node pPr): body_(body), 
                                        p_(p), 
                                        pPr_(pPr)
  {
    GetSectPr();
  }

  void Section::GetSectPr()
  {
    pugi::xml_node pNext = p_, p, pPr, sectPr;
    do {
      p = pNext;
      pPr = p.child("w:pPr");
      sectPr = pPr.child("w:sectPr");
      pNext = p.next_sibling();
    } while (sectPr.empty() && !pNext.empty());

    pLast_         = p;
    pPrLast_       = pPr;
    sectPr_        = sectPr;

    if (sectPr_.empty()) sectPr_ = body_.child("w:sectPr");
  }

  void Section::Split()
  {
    if (IsSplit()) return;
    pLast_ = p_;
    pPrLast_ = pPr_;
    sectPr_ = pPr_.append_copy(sectPr_);
  }

  bool Section::IsSplit()
  {
    return pPr_.child("w:sectPr");
  }

  void Section::Merge()
  {
    if (pPr_.child("w:sectPr").empty()) return;
    pPrLast_.remove_child(sectPr_);
    GetSectPr();
  }

  void Section::SetPageSize(const int w, const int h)
  {
    auto pgSz = sectPr_.child("w:pgSz");
    if (!pgSz) {
      pgSz = sectPr_.append_child("w:pgSz");
    }
    auto pgSzW = pgSz.attribute("w:w");
    if (!pgSzW) {
      pgSzW = pgSz.append_attribute("w:w");
    }
    auto pgSzH = pgSz.attribute("w:h");
    if (!pgSzH) {
      pgSzH = pgSz.append_attribute("w:h");
    }
    pgSzW.set_value(w);
    pgSzH.set_value(h);
  }

  void Section::GetPageSize(int &w, int &h)
  {
    w = h = 0;
    auto pgSz = sectPr_.child("w:pgSz");
    if (!pgSz) return;
    auto pgSzW = pgSz.attribute("w:w");
    if (!pgSzW) return;
    auto pgSzH = pgSz.attribute("w:h");
    if (!pgSzH) return;
    w = pgSzW.as_int();
    h = pgSzH.as_int();
  }

  void Section::SetPageOrient(Orientation orient)
  {
    auto pgSz = sectPr_.child("w:pgSz");
    if (!pgSz) {
      pgSz = sectPr_.append_child("w:pgSz");
    }
    auto pgSzH = pgSz.attribute("w:h");
    if (!pgSzH) {
      pgSzH = pgSz.append_attribute("w:h");
    }
    auto pgSzW = pgSz.attribute("w:w");
    if (!pgSzW) {
      pgSzW = pgSz.append_attribute("w:w");
    }
    auto pgSzOrient = pgSz.attribute("w:orient");
    if (!pgSzOrient) {
      pgSzOrient = pgSz.append_attribute("w:orient");
    }
    int hVal = pgSzH.as_int();
    int wVal = pgSzW.as_int();
    switch (orient) {
      case Orientation::Landscape:
        if (hVal < wVal) return;
        pgSzOrient.set_value("landscape");
        pgSzH.set_value(wVal);
        pgSzW.set_value(hVal);
        break;
      case Orientation::Portrait:
        if (hVal > wVal) return;
        pgSzOrient.set_value("portrait");
        pgSzH.set_value(wVal);
        pgSzW.set_value(hVal);
        break;
    }
  }

  Section::Orientation Section::GetPageOrient()
  {
    Orientation orient = Orientation::Portrait;
    auto pgSz = sectPr_.child("w:pgSz");
    if (!pgSz) return orient;
    auto pgSzOrient = pgSz.attribute("w:orient");
    if (!pgSzOrient) return orient;
    if (std::string(pgSzOrient.value()).compare("landscape") == 0) {
      orient = Orientation::Landscape;
    }
    return orient;
  }

  void Section::SetPageMargin(const int top, const int bottom, 
                              const int left, const int right)
  {
    auto pgMar = sectPr_.child("w:pgMar");
    if (!pgMar) {
      pgMar = sectPr_.append_child("w:pgMar");
    }
    auto pgMarTop = pgMar.attribute("w:top");
    if (!pgMarTop) {
      pgMarTop = pgMar.append_attribute("w:top");
    }
    auto pgMarBottom = pgMar.attribute("w:bottom");
    if (!pgMarBottom) {
      pgMarBottom = pgMar.append_attribute("w:bottom");
    }
    auto pgMarLeft = pgMar.attribute("w:left");
    if (!pgMarLeft) {
      pgMarLeft = pgMar.append_attribute("w:left");
    }
    auto pgMarRight = pgMar.attribute("w:right");
    if (!pgMarRight) {
      pgMarRight = pgMar.append_attribute("w:right");
    }
    pgMarTop.set_value(top);
    pgMarBottom.set_value(bottom);
    pgMarLeft.set_value(left);
    pgMarRight.set_value(right);
  }

  void Section::GetPageMargin(int &top, int &bottom,
                              int &left, int &right)
  {
    top = bottom = left = right = 0;
    auto pgMar = sectPr_.child("w:pgMar");
    if (!pgMar) return;
    auto pgMarTop = pgMar.attribute("w:top");
    if (!pgMarTop) return;
    auto pgMarBottom = pgMar.attribute("w:bottom");
    if (!pgMarBottom) return;
    auto pgMarLeft = pgMar.attribute("w:left");
    if (!pgMarLeft) return;
    auto pgMarRight = pgMar.attribute("w:right");
    if (!pgMarRight) return;
    top    = pgMarTop.as_int();
    bottom = pgMarBottom.as_int();
    left   = pgMarLeft.as_int();
    right  = pgMarRight.as_int();
  }

  void Section::SetPageMargin(const int header, const int footer)
  {
    auto pgMar = sectPr_.child("w:pgMar");
    if (!pgMar) {
      pgMar = sectPr_.append_child("w:pgMar");
    }
    auto pgMarHeader = pgMar.attribute("w:header");
    if (!pgMarHeader) {
      pgMarHeader = pgMar.append_attribute("w:header");
    }
    auto pgMarFooter = pgMar.attribute("w:footer");
    if (!pgMarFooter) {
      pgMarFooter = pgMar.append_attribute("w:footer");
    }
    pgMarHeader.set_value(header);
    pgMarFooter.set_value(footer);
  }

  void Section::GetPageMargin(int &header, int &footer)
  {
    header = footer = 0;
    auto pgMar = sectPr_.child("w:pgMar");
    if (!pgMar) return;
    auto pgMarHeader = pgMar.attribute("w:header");
    if (!pgMarHeader) return;
    auto pgMarFooter = pgMar.attribute("w:footer");
    if (!pgMarFooter) return;
    header = pgMarHeader.as_int();
    footer = pgMarFooter.as_int();
  }

  Paragraph Section::FirstParagraph()
  {
    return Paragraph(body_, pLast_, pPrLast_);
  }

  Paragraph Section::LastParagraph()
  {
    return Paragraph(body_, pLast_, pPrLast_);
  }

  Section Section::Next()
  {
    auto p = pLast_.next_sibling();
    if (p.empty()) return Section();
    return Section(body_, p, p.child("w:pPr"));
  }

  Section Section::Prev()
  {
    pugi::xml_node pPrev, p, pPr, sectPr;

    pPrev = p_.previous_sibling();
    if (pPrev.empty()) return Section();

    do {
      p = pPrev;
      pPr = p.child("w:pPr");
      sectPr = pPr.child("w:sectPr");
      pPrev = p.previous_sibling();
    } while (sectPr.empty() && !pPrev.empty());

    return Section(body_, p, pPr, sectPr);
  }

  Section::operator bool()
  {
    return sectPr_;
  }


  // class Run
  void Run::AppendText(const std::string text)
  {
    auto t = r_.append_child("w:t");
    if (isspace(text.front()) || isspace(text.back())) {
      t.append_attribute("xml:space") = "preserve";
    }
    t.text().set(text.c_str());
  }

  std::string Run::GetText()
  {
    std::string text;
    for (auto t = r_.child("w:t"); t; t = t.next_sibling("w:t")) {
      text += t.text().get();
    }
    return text;
  }

  void Run::ClearText()
  {
    r_.remove_children();
  }

  void Run::AppendLineBreak()
  {
    r_.append_child("w:br");
  }

  void Run::SetFontSize(const double fontSize)
  {
    auto sz = rPr_.child("w:sz");
    if (!sz) {
      sz = rPr_.append_child("w:sz");
    }
    auto szVal = sz.attribute("w:val");
    if (!szVal) {
      szVal = sz.append_attribute("w:val");
    }
    // font size in half-points (1/144 of an inch)
    szVal.set_value(fontSize * 2);
  }

  double Run::GetFontSize()
  {
    auto sz = rPr_.child("w:sz");
    if (!sz) return 0;
    auto szVal = sz.attribute("w:val");
    if (!szVal) return 0;
    return szVal.as_int() / 2.0;
  }

  void Run::SetFont(const std::string fontAscii, 
                    const std::string fontEastAsia)
  {
    auto rFonts = rPr_.child("w:rFonts");
    if (!rFonts) {
      rFonts = rPr_.append_child("w:rFonts");
    }
    auto rFontsAscii = rFonts.attribute("w:ascii");
    if (!rFontsAscii) {
      rFontsAscii = rFonts.append_attribute("w:ascii");
    }
    auto rFontsEastAsia = rFonts.attribute("w:eastAsia");
    if (!rFontsEastAsia) {
      rFontsEastAsia = rFonts.append_attribute("w:eastAsia");
    }
    rFontsAscii.set_value(fontAscii.c_str());
    rFontsEastAsia.set_value(fontEastAsia.empty() 
                           ? fontAscii.c_str()
                           : fontEastAsia.c_str());
  }

  void Run::GetFont(std::string &fontAscii, 
                    std::string &fontEastAsia)
  {
    auto rFonts = rPr_.child("w:rFonts");
    if (!rFonts) return;

    auto rFontsAscii = rFonts.attribute("w:ascii");
    if (rFontsAscii) fontAscii = rFontsAscii.value();

    auto rFontsEastAsia = rFonts.attribute("w:eastAsia");
    if (rFontsEastAsia) fontEastAsia = rFontsEastAsia.value();
  }

  void Run::SetFontStyle(FontStyle f)
  {
    auto b = rPr_.child("w:b");
    if (f & Bold) {
      if (b.empty()) rPr_.append_child("w:b");
    } else {
      rPr_.remove_child(b);
    }

    auto i = rPr_.child("w:i");
    if (f & Italic) {
      if (i.empty()) rPr_.append_child("w:i");
    } else {
      rPr_.remove_child(i);
    }

    auto u = rPr_.child("w:u");
    if (f & Underline) {
      if (u.empty())
        rPr_.append_child("w:u").append_attribute("w:val") = "single";
    } else {
      rPr_.remove_child(u);
    }

    auto strike = rPr_.child("w:strike");
    if (f & Strikethrough) {
      if (strike.empty())
        rPr_.append_child("w:strike").append_attribute("w:val") = "true";
    } else {
      rPr_.remove_child(strike);
    }
  }

  Run::FontStyle Run::GetFontStyle()
  {
    FontStyle fontStyle = 0;

    if (rPr_.child("w:b")) fontStyle |= Bold;
    if (rPr_.child("w:i")) fontStyle |= Italic;
    if (rPr_.child("w:u")) fontStyle |= Underline;
    if (rPr_.child("w:strike")) fontStyle |= Strikethrough;

    return fontStyle;
  }

  void Run::SetCharacterSpacing(const int characterSpacing)
  {
    auto spacing = rPr_.child("w:spacing");
    if (!spacing) {
      spacing = rPr_.append_child("w:spacing");
    }
    auto spacingVal = spacing.attribute("w:val");
    if (!spacingVal) {
      spacingVal = spacing.append_attribute("w:val");
    }
    spacingVal.set_value(characterSpacing);
  }

  int Run::GetCharacterSpacing()
  {
    return rPr_.child("w:spacing").attribute("w:val").as_int();
  }

  bool Run::IsPageBreak()
  {
    return r_.find_child_by_attribute("w:br", "w:type", "page");
  }

  void Run::Remove()
  {
    p_.remove_child(r_);
  }

  Run Run::Next()
  {
    auto r = r_.next_sibling("w:r");
    auto rPr = r.child("w:rPr");
    return Run(p_, r, rPr);
  }

  Run::operator bool()
  {
    return r_;
  }


} // namespace docx
