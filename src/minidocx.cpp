
#include "minidocx.hpp"
#include <cstring> // std::strlen(), std::strcmp()
#include <cstdlib> // std::free()
#include "zip.h"

#define _RELS R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>)"
#define DOCUMENT_XML R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex" xmlns:cx1="http://schemas.microsoft.com/office/drawing/2015/9/8/chartex" xmlns:cx2="http://schemas.microsoft.com/office/drawing/2015/10/21/chartex" xmlns:cx3="http://schemas.microsoft.com/office/drawing/2016/5/9/chartex" xmlns:cx4="http://schemas.microsoft.com/office/drawing/2016/5/10/chartex" xmlns:cx5="http://schemas.microsoft.com/office/drawing/2016/5/11/chartex" xmlns:cx6="http://schemas.microsoft.com/office/drawing/2016/5/12/chartex" xmlns:cx7="http://schemas.microsoft.com/office/drawing/2016/5/13/chartex" xmlns:cx8="http://schemas.microsoft.com/office/drawing/2016/5/14/chartex" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:aink="http://schemas.microsoft.com/office/drawing/2016/ink" xmlns:am3d="http://schemas.microsoft.com/office/drawing/2017/model3d" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:oel="http://schemas.microsoft.com/office/2019/extlst" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" xmlns:w16cex="http://schemas.microsoft.com/office/word/2018/wordml/cex" xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid" xmlns:w16="http://schemas.microsoft.com/office/word/2018/wordml" xmlns:w16sdtdh="http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash" xmlns:w16se="http://schemas.microsoft.com/office/word/2015/wordml/symex" xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk" xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" mc:Ignorable="w14 w15 w16se w16cid w16 w16cex w16sdtdh wp14"><w:body><w:sectPr><w:pgSz w:w="11906" w:h="16838" /><w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800" w:header="851" w:footer="992" w:gutter="0" /><w:cols w:space="425" /><w:docGrid w:type="lines" w:linePitch="312" /></w:sectPr></w:body></w:document>)"
#define CONTENT_TYPES_XML R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>)"

namespace docx
{
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
    while (!child.empty() && std::strcmp(name, child.name()) != 0) {
      child = child.previous_sibling(name);
    }
    return child;
  }

  std::ostream& operator<<(std::ostream &out, const Document &doc)
  {
    xml_string_writer writer;
    doc.w_body_.print(writer, "  ");
    out << writer.result;
    return out;
  }

  // class Document
  Document::Document(const std::string path): path_(path)
  {
    doc_.load_buffer(DOCUMENT_XML, std::strlen(DOCUMENT_XML), pugi::parse_declaration);
    w_body_ = doc_.child("w:document").child("w:body");
    w_sectPr_ = w_body_.child("w:sectPr");
  }

  bool Document::Save()
  {
    xml_string_writer writer;
    doc_.save(writer, "", pugi::format_raw);
    const char *buf = writer.result.c_str();

    struct zip_t *zip = zip_open(path_.c_str(), ZIP_DEFAULT_COMPRESSION_LEVEL, 'w');
    if (zip == nullptr) {
      return false;
    }

    zip_entry_open(zip, "_rels/.rels");
    zip_entry_write(zip, _RELS, std::strlen(_RELS));
    zip_entry_close(zip);

    zip_entry_open(zip, "word/document.xml");
    zip_entry_write(zip, buf, std::strlen(buf));
    zip_entry_close(zip);

    zip_entry_open(zip, "[Content_Types].xml");
    zip_entry_write(zip, CONTENT_TYPES_XML, std::strlen(CONTENT_TYPES_XML));
    zip_entry_close(zip);

    zip_close(zip);
    return true;
  }

  bool Document::Open(const std::string path)
  {
    struct zip_t *zip = zip_open(path.c_str(), ZIP_DEFAULT_COMPRESSION_LEVEL, 'r');
    if (zip == nullptr) {
      return false;
    }

    if (zip_entry_open(zip, "word/document.xml") < 0) {
      zip_close(zip);
      return false;
    }
    void *buf = nullptr;
    size_t bufsize;
    zip_entry_read(zip, &buf, &bufsize);
    zip_entry_close(zip);
    zip_close(zip);

    doc_.load_buffer(buf, bufsize, pugi::parse_declaration);
    w_body_ = doc_.child("w:document").child("w:body");
    w_sectPr_ = w_body_.child("w:sectPr");
    std::free(buf);
    return true;
  }

  Paragraph Document::FirstParagraph()
  {
    auto w_p = w_body_.child("w:p");
    auto w_pPr = w_p.child("w:pPr");
    return Paragraph(w_body_, w_p, w_pPr);
  }

  Paragraph Document::LastParagraph()
  {
    auto w_p = GetLastChild(w_body_, "w:p");
    auto w_pPr = w_p.child("w:pPr");
    return Paragraph(w_body_, w_p, w_pPr);
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
    auto w_p = w_body_.insert_child_before("w:p", w_sectPr_);
    auto w_pPr = w_p.append_child("w:pPr");
    return Paragraph(w_body_, w_p, w_pPr);
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
    auto w_p = w_body_.prepend_child("w:p");
    auto w_pPr = w_p.append_child("w:pPr");
    return Paragraph(w_body_, w_p, w_pPr);
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

  Paragraph Document::InsertParagraphBefore(Paragraph &p)
  {
    auto w_p = w_body_.insert_child_before("w:p", p.w_p_);
    auto w_pPr = w_p.append_child("w:pPr");
    return Paragraph(w_body_, w_p, w_pPr);
  }

  Paragraph Document::InsertParagraphAfter(Paragraph &p)
  {
    auto w_p = w_body_.insert_child_after("w:p", p.w_p_);
    auto w_pPr = w_p.append_child("w:pPr");
    return Paragraph(w_body_, w_p, w_pPr);
  }

  bool Document::RemoveParagraph(Paragraph &p)
  {
    return w_body_.remove_child(p.w_p_);
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

  Table Document::AppendTable(const int rows, const int cols)
  {
    auto w_tbl = w_body_.insert_child_before("w:tbl", w_sectPr_);
    auto w_tblPr = w_tbl.append_child("w:tblPr");
    auto w_tblGrid = w_tbl.append_child("w:tblGrid");
    auto tbl = Table(w_body_, w_tbl, w_tblPr, w_tblGrid);
    tbl.SetGrid(rows, cols);
    tbl.SetWidthPercent(100);
    tbl.SetAllBorders();
    return tbl;
  }


  // class Paragraph
  Paragraph::Paragraph()
  {
  }

  Paragraph::Paragraph(pugi::xml_node w_body, 
                       pugi::xml_node w_p, 
                       pugi::xml_node w_pPr): w_body_(w_body), 
                                              w_p_(w_p), 
                                              w_pPr_(w_pPr)
  {
  }

  Run Paragraph::FirstRun()
  {
    auto w_r = w_p_.child("w:r");
    auto w_rPr = w_r.child("w:rPr");
    return Run(w_p_, w_r, w_rPr);
  }

  Run Paragraph::AppendRun()
  {
    auto w_r = w_p_.append_child("w:r");
    auto w_rPr = w_r.append_child("w:rPr");
    return Run(w_p_, w_r, w_rPr);
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
    auto w_r = w_p_.append_child("w:r");
    auto w_br = w_r.append_child("w:br");
    w_br.append_attribute("w:type") = "page";
    return Run(w_p_, w_r, w_br);
  }

  void Paragraph::SetAlignment(const Alignment alignment)
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

    auto jc = w_pPr_.child("w:jc");
    if (!jc) {
      jc = w_pPr_.append_child("w:jc");
    }
    auto jcVal = jc.attribute("w:val");
    if (!jcVal) {
      jcVal = jc.append_attribute("w:val");
    }
    jcVal.set_value(val);
  }

  void Paragraph::SetLineSpacingSingle()
  {
    auto spacing = w_pPr_.child("w:spacing");
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

  void Paragraph::SetLineSpacingLines(const double at)
  {
    // A normal single-spaced paragaph has a w:line value of 240, or 12 points.
    // 
    // If the value of lineRule is auto, then the value of line 
    // is interpreted as 240th of a line, e.g. 360 = 1.5 lines.
    SetLineSpacing(at * 240, "auto");
  }

  void Paragraph::SetLineSpacingAtLeast(const int at)
  {
    // If the value of the lineRule attribute is atLeast or exactly, 
    // then the value of the line attribute is interpreted as 240th of a point.
    // (Not really. Actually, values are in twentieths of a point, e.g. 240 = 12 pt.)
    SetLineSpacing(at, "atLeast");
  }

  void Paragraph::SetLineSpacingExactly(const int at)
  {
    SetLineSpacing(at, "exact");
  }

  void Paragraph::SetLineSpacing(const int at, const char *lineRule)
  {
    auto spacing = w_pPr_.child("w:spacing");
    if (!spacing) {
      spacing = w_pPr_.append_child("w:spacing");
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
    auto spacing = w_pPr_.child("w:spacing");
    if (!spacing) {
      spacing = w_pPr_.append_child("w:spacing");
    }
    auto spacingAuto = spacing.attribute(attrNameAuto);
    if (!spacingAuto) {
      spacingAuto = spacing.append_attribute(attrNameAuto);
    }
    // Any value for before or beforeLines is ignored.
    spacingAuto.set_value("true");
  }

  void Paragraph::SetBeforeSpacingLines(const double beforeSpacing)
  {
    // To specify units in hundreths of a line, 
    // use attributes 'afterLines'/'beforeLines'.
    SetSpacing(beforeSpacing * 100, "w:beforeAutospacing", "w:beforeLines");
  }

  void Paragraph::SetAfterSpacingLines(const double afterSpacing)
  {
    SetSpacing(afterSpacing * 100, "w:afterAutospacing", "w:afterLines");
  }

  void Paragraph::SetBeforeSpacing(const int beforeSpacing)
  {
    SetSpacing(beforeSpacing, "w:beforeAutospacing", "w:before");
  }

  void Paragraph::SetAfterSpacing(const int afterSpacing)
  {
    SetSpacing(afterSpacing, "w:afterAutospacing", "w:after");
  }

  void Paragraph::SetSpacing(const int twip, const char *attrNameAuto, const char *attrName)
  {
    auto elemSpacing = w_pPr_.child("w:spacing");
    if (!elemSpacing) {
      elemSpacing = w_pPr_.append_child("w:spacing");
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

  void Paragraph::SetLeftIndentChars(const double leftIndent)
  {
    // To specify units in hundreths of a character, 
    // use attributes leftChars/endChars, rightChars/endChars, etc. 
    SetIndent(leftIndent * 100, "w:leftChars");
  }

  void Paragraph::SetRightIndentChars(const double rightIndent)
  {
    SetIndent(rightIndent * 100, "w:rightChars");
  }

  void Paragraph::SetLeftIndent(const int leftIndent)
  {
    SetIndent(leftIndent, "w:left");
  }

  void Paragraph::SetRightIndent(const int rightIndent)
  {
    SetIndent(rightIndent, "w:right");
  }

  void Paragraph::SetFirstLineChars(const double indent)
  {
    SetIndent(indent * 100, "w:firstLineChars");
  }

  void Paragraph::SetHangingChars(const double indent)
  {
    SetIndent(indent * 100, "w:hangingChars");
  }

  void Paragraph::SetFirstLine(const int indent)
  {
    SetIndent(indent, "w:firstLine");
  }

  void Paragraph::SetHanging(const int indent)
  {
    SetIndent(indent, "w:hanging");
    SetLeftIndent(indent);
  }

  void Paragraph::SetIndent(const int indent, const char *attrName)
  {
    auto elemIndent = w_pPr_.child("w:ind");
    if (!elemIndent) {
      elemIndent = w_pPr_.append_child("w:ind");
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

  void Paragraph::SetFontStyle(const Run::FontStyle fontStyle)
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

  Paragraph Paragraph::Next()
  {
    auto w_p = w_p_.next_sibling("w:p");
    auto w_pPr = w_p.child("w:pPr");
    return Paragraph(w_body_, w_p, w_pPr);
  }

  Paragraph Paragraph::Prev()
  {
    auto w_p = w_p_.previous_sibling("w:p");
    auto w_pPr = w_p.child("w:pPr");
    return Paragraph(w_body_, w_p, w_pPr);
  }

  Paragraph::operator bool()
  {
    return w_p_;
  }

  bool Paragraph::operator==(const Paragraph &p)
  {
    return w_p_ == p.w_p_;
  }

  Section Paragraph::GetSection()
  {
    return Section(w_body_, w_p_, w_pPr_);
  }

  Section Paragraph::InsertSectionBreak()
  {
    auto s = Section(w_body_, w_p_, w_pPr_);
    // this paragraph will be the last paragraph of the new section
    s.Split();
    return s;
  }

  Section Paragraph::RemoveSectionBreak()
  {
    auto s = Section(w_body_, w_p_, w_pPr_);
    if (s.IsSplit()) s.Merge();
    return s;
  }

  bool Paragraph::HasSectionBreak()
  {
    return GetSection().IsSplit();
  }


  // class Section
  Section::Section()
  {}

  Section::Section(pugi::xml_node w_body, 
                   pugi::xml_node w_p, 
                   pugi::xml_node w_pPr): w_body_(w_body), 
                                          w_p_(w_p), 
                                          w_pPr_(w_pPr)
  {
    GetSectPr();
  }

  Section::Section(pugi::xml_node w_body, 
                   pugi::xml_node w_p, 
                   pugi::xml_node w_pPr, 
                   pugi::xml_node w_sectPr): w_body_(w_body), 
                                             w_p_(w_p), 
                                             w_p_last_(w_p), 
                                             w_pPr_(w_pPr), 
                                             w_pPr_last_(w_pPr), 
                                             w_sectPr_(w_sectPr)
  {}

  void Section::GetSectPr()
  {
    pugi::xml_node w_p_next = w_p_, w_p, w_pPr, w_sectPr;
    do {
      w_p = w_p_next;
      w_pPr = w_p.child("w:pPr");
      w_sectPr = w_pPr.child("w:sectPr");
      w_p_next = w_p.next_sibling();
    } while (w_sectPr.empty() && !w_p_next.empty());

    w_p_last_   = w_p;
    w_pPr_last_ = w_pPr;
    w_sectPr_   = w_sectPr;

    if (w_sectPr_.empty()) w_sectPr_ = w_body_.child("w:sectPr");
  }

  void Section::Split()
  {
    if (IsSplit()) return;
    w_p_last_ = w_p_;
    w_pPr_last_ = w_pPr_;
    w_sectPr_ = w_pPr_.append_copy(w_sectPr_);
  }

  bool Section::IsSplit()
  {
    return w_pPr_.child("w:sectPr");
  }

  void Section::Merge()
  {
    if (w_pPr_.child("w:sectPr").empty()) return;
    w_pPr_last_.remove_child(w_sectPr_);
    GetSectPr();
  }

  void Section::SetPageSize(const int w, const int h)
  {
    auto pgSz = w_sectPr_.child("w:pgSz");
    if (!pgSz) {
      pgSz = w_sectPr_.append_child("w:pgSz");
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
    auto pgSz = w_sectPr_.child("w:pgSz");
    if (!pgSz) return;
    auto pgSzW = pgSz.attribute("w:w");
    if (!pgSzW) return;
    auto pgSzH = pgSz.attribute("w:h");
    if (!pgSzH) return;
    w = pgSzW.as_int();
    h = pgSzH.as_int();
  }

  void Section::SetPageOrient(const Orientation orient)
  {
    auto pgSz = w_sectPr_.child("w:pgSz");
    if (!pgSz) {
      pgSz = w_sectPr_.append_child("w:pgSz");
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
    auto pgSz = w_sectPr_.child("w:pgSz");
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
    auto pgMar = w_sectPr_.child("w:pgMar");
    if (!pgMar) {
      pgMar = w_sectPr_.append_child("w:pgMar");
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
    auto pgMar = w_sectPr_.child("w:pgMar");
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
    auto pgMar = w_sectPr_.child("w:pgMar");
    if (!pgMar) {
      pgMar = w_sectPr_.append_child("w:pgMar");
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
    auto pgMar = w_sectPr_.child("w:pgMar");
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
    return Prev().LastParagraph().Next();
  }

  Paragraph Section::LastParagraph()
  {
    return Paragraph(w_body_, w_p_last_, w_pPr_last_);
  }

  Section Section::Next()
  {
    auto w_p = w_p_last_.next_sibling();
    if (w_p.empty()) return Section();
    return Section(w_body_, w_p, w_p.child("w:pPr"));
  }

  Section Section::Prev()
  {
    pugi::xml_node w_p_prev, w_p, w_pPr, w_sectPr;

    w_p_prev = w_p_.previous_sibling();
    if (w_p_prev.empty()) return Section();

    do {
      w_p = w_p_prev;
      w_pPr = w_p.child("w:pPr");
      w_sectPr = w_pPr.child("w:sectPr");
      w_p_prev = w_p.previous_sibling();
    } while (w_sectPr.empty() && !w_p_prev.empty());

    return Section(w_body_, w_p, w_pPr, w_sectPr);
  }

  Section::operator bool()
  {
    return w_sectPr_;
  }

  bool Section::operator==(const Section &s)
  {
    return w_sectPr_ == s.w_sectPr_;
  }


  // class Run
  Run::Run(pugi::xml_node w_p, 
           pugi::xml_node w_r, 
           pugi::xml_node w_rPr): w_p_(w_p), 
                                  w_r_(w_r), 
                                  w_rPr_(w_rPr)
  {}

  void Run::AppendText(const std::string text)
  {
    auto t = w_r_.append_child("w:t");
    if (isspace(text.front()) || isspace(text.back())) {
      t.append_attribute("xml:space") = "preserve";
    }
    t.text().set(text.c_str());
  }

  std::string Run::GetText()
  {
    std::string text;
    for (auto t = w_r_.child("w:t"); t; t = t.next_sibling("w:t")) {
      text += t.text().get();
    }
    return text;
  }

  void Run::ClearText()
  {
    w_r_.remove_children();
  }

  void Run::AppendLineBreak()
  {
    w_r_.append_child("w:br");
  }

  void Run::SetFontSize(const double fontSize)
  {
    auto sz = w_rPr_.child("w:sz");
    if (!sz) {
      sz = w_rPr_.append_child("w:sz");
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
    auto sz = w_rPr_.child("w:sz");
    if (!sz) return 0;
    auto szVal = sz.attribute("w:val");
    if (!szVal) return 0;
    return szVal.as_int() / 2.0;
  }

  void Run::SetFont(const std::string fontAscii, 
                    const std::string fontEastAsia)
  {
    auto rFonts = w_rPr_.child("w:rFonts");
    if (!rFonts) {
      rFonts = w_rPr_.append_child("w:rFonts");
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
    auto rFonts = w_rPr_.child("w:rFonts");
    if (!rFonts) return;

    auto rFontsAscii = rFonts.attribute("w:ascii");
    if (rFontsAscii) fontAscii = rFontsAscii.value();

    auto rFontsEastAsia = rFonts.attribute("w:eastAsia");
    if (rFontsEastAsia) fontEastAsia = rFontsEastAsia.value();
  }

  void Run::SetFontStyle(const FontStyle f)
  {
    auto b = w_rPr_.child("w:b");
    if (f & Bold) {
      if (b.empty()) w_rPr_.append_child("w:b");
    } else {
      w_rPr_.remove_child(b);
    }

    auto i = w_rPr_.child("w:i");
    if (f & Italic) {
      if (i.empty()) w_rPr_.append_child("w:i");
    } else {
      w_rPr_.remove_child(i);
    }

    auto u = w_rPr_.child("w:u");
    if (f & Underline) {
      if (u.empty())
        w_rPr_.append_child("w:u").append_attribute("w:val") = "single";
    } else {
      w_rPr_.remove_child(u);
    }

    auto strike = w_rPr_.child("w:strike");
    if (f & Strikethrough) {
      if (strike.empty())
        w_rPr_.append_child("w:strike").append_attribute("w:val") = "true";
    } else {
      w_rPr_.remove_child(strike);
    }
  }

  Run::FontStyle Run::GetFontStyle()
  {
    FontStyle fontStyle = 0;
    if (w_rPr_.child("w:b")) fontStyle |= Bold;
    if (w_rPr_.child("w:i")) fontStyle |= Italic;
    if (w_rPr_.child("w:u")) fontStyle |= Underline;
    if (w_rPr_.child("w:strike")) fontStyle |= Strikethrough;
    return fontStyle;
  }

  void Run::SetCharacterSpacing(const int characterSpacing)
  {
    auto spacing = w_rPr_.child("w:spacing");
    if (!spacing) {
      spacing = w_rPr_.append_child("w:spacing");
    }
    auto spacingVal = spacing.attribute("w:val");
    if (!spacingVal) {
      spacingVal = spacing.append_attribute("w:val");
    }
    spacingVal.set_value(characterSpacing);
  }

  int Run::GetCharacterSpacing()
  {
    return w_rPr_.child("w:spacing").attribute("w:val").as_int();
  }

  bool Run::IsPageBreak()
  {
    return w_r_.find_child_by_attribute("w:br", "w:type", "page");
  }

  void Run::Remove()
  {
    w_p_.remove_child(w_r_);
  }

  Run Run::Next()
  {
    auto w_r = w_r_.next_sibling("w:r");
    auto w_rPr = w_r.child("w:rPr");
    return Run(w_p_, w_r, w_rPr);
  }

  Run::operator bool()
  {
    return w_r_;
  }

  // class Table
  Table::Table(pugi::xml_node w_body, 
               pugi::xml_node w_tbl, 
               pugi::xml_node w_tblPr, 
               pugi::xml_node w_tblGrid): w_body_(w_body),
                                          w_tbl_(w_tbl), 
                                          w_tblPr_(w_tblPr), 
                                          w_tblGrid_(w_tblGrid)
  {}

  void Table::SetGrid(const int rows, const int cols)
  {
    rows_ = rows;
    cols_ = cols;

    for (int i = 0; i < rows; i++) {
      Row row;
      for (int j = 0; j < cols; j++) {
        Cell cell = { i, j, 1, 1 };
        row.push_back(cell);
      }
      grid_.push_back(row);
    }

    for (int i = 0; i < rows; i++) {
      auto w_gridCol = w_tblGrid_.append_child("w:gridCol");

      auto w_tr = w_tbl_.append_child("w:tr");
      for (int j = 0; j < cols; j++) {
        auto w_tc = w_tr.append_child("w:tc");
        auto w_tcPr = w_tc.append_child("w:tcPr");
        auto c = TableCell(i, j, w_tr, w_tc, w_tcPr);
        // A table cell must contain at least one block-level element, 
        // even if it is an empty <p/>.
        c.AppendParagraph();
      }
    }
  }

  bool Table::MergCells(TableCell c1, TableCell c2)
  {
    return false;
  }

  void Table::SetWidthAuto()
  {
    SetWidth(0, "auto");
  }

  void Table::SetWidthPercent(const double w)
  {
    SetWidth(w / 0.02, "pct");
  }

  void Table::SetWidth(const int w, const char *units)
  {
    auto w_tblW = w_tblPr_.child("w:tblW");
    if (!w_tblW) {
      w_tblW = w_tblPr_.append_child("w:tblW");
    }

    auto w_w = w_tblW.attribute("w:w");
    if (!w_w) {
      w_w = w_tblW.append_attribute("w:w");
    }

    auto w_type = w_tblW.attribute("w:type");
    if (!w_type) {
      w_type = w_tblW.append_attribute("w:type");
    }

    w_w.set_value(w);
    w_type.set_value(units);
  }

  void Table::SetCellMarginTop(const int w, const char *units)
  {
    SetCellMargin("w:top", w, units);
  }

  void Table::SetCellMarginBottom(const int w, const char *units)
  {
    SetCellMargin("w:bottom", w, units);
  }

  void Table::SetCellMarginLeft(const int w, const char *units)
  {
    SetCellMargin("w:start", w, units);
  }

  void Table::SetCellMarginRight(const int w, const char *units)
  {
    SetCellMargin("w:end", w, units);
  }

  void Table::SetCellMargin(const char *name, const int w, const char *units)
  {
    auto w_tblCellMar = w_tblPr_.child("w:tblCellMar");
    if (!w_tblCellMar) {
      w_tblCellMar = w_tblPr_.append_child("w:tblCellMar");
    }

    auto w_tblCellMarChild = w_tblCellMar.child(name);
    if (!w_tblCellMarChild) {
      w_tblCellMarChild = w_tblCellMar.append_child(name);
    }

    auto w_w = w_tblCellMarChild.attribute("w:w");
    if (!w_w) {
      w_w = w_tblCellMarChild.append_attribute("w:w");
    }

    auto w_type = w_tblCellMarChild.attribute("w:type");
    if (!w_type) {
      w_type = w_tblCellMarChild.append_attribute("w:type");
    }

    w_w.set_value(w);
    w_type.set_value(units);
  }

  TableCell Table::GetCell(const int row, const int col)
  {
    if (row < 0 || row >= rows_ || col < 0 || col >= cols_) {
      return TableCell();
    }

    int i = 0;
    auto w_tr = w_tbl_.child("w:tr");
    while (i < row && !w_tr.empty()) {
      w_tr = w_tr.next_sibling("w:tr");
      i++;
    }
    if (w_tr.empty()) {
      return TableCell();
    }

    int j = 0;
    auto w_tc = w_tr.child("w:tc");
    while (j < col && !w_tc.empty()) {
      w_tc = w_tc.next_sibling("w:tc");
      j++;
    }
    if (w_tc.empty()) {
      return TableCell();
    }

    auto w_tcPr = w_tc.child("w:tcPr");
    return TableCell(row, col, w_tr, w_tc, w_tcPr);
  }

  void Table::SetAlignment(const Alignment alignment)
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
    }

    auto w_jc = w_tblPr_.child("w:jc");
    if (!w_jc) {
      w_jc = w_tblPr_.append_child("w:jc");
    }
    auto w_val = w_jc.attribute("w:val");
    if (!w_val) {
      w_val = w_jc.append_attribute("w:val");
    }
    w_val.set_value(val);
  }

  void Table::SetTopBorders(const BorderStyle style, const double width, const char *color)
  {
    SetBorders("w:top", style, width, color);
  }

  void Table::SetBottomBorders(const BorderStyle style, const double width, const char *color)
  {
    SetBorders("w:bottom", style, width, color);
  }

  void Table::SetLeftBorders(const BorderStyle style, const double width, const char *color)
  {
    SetBorders("w:start", style, width, color);
  }

  void Table::SetRightBorders(const BorderStyle style, const double width, const char *color)
  {
    SetBorders("w:end", style, width, color);
  }

  void Table::SetInsideHBorders(const BorderStyle style, const double width, const char *color)
  {
    SetBorders("w:insideH", style, width, color);
  }

  void Table::SetInsideVBorders(const BorderStyle style, const double width, const char *color)
  {
    SetBorders("w:insideV", style, width, color);
  }

  void Table::SetInsideBorders(const BorderStyle style, const double width, const char *color)
  {
    SetInsideHBorders(style, width, color);
    SetInsideVBorders(style, width, color);
  }

  void Table::SetOutsideBorders(const BorderStyle style, const double width, const char *color)
  {
    SetTopBorders(style, width, color);
    SetBottomBorders(style, width, color);
    SetLeftBorders(style, width, color);
    SetRightBorders(style, width, color);
  }

  void Table::SetAllBorders(const BorderStyle style, const double width, const char *color)
  {
    SetOutsideBorders(style, width, color);
    SetInsideBorders(style, width, color);
  }

  void Table::SetBorders(const char *name, const BorderStyle style, const double width, const char *color)
  {
    auto w_tblBorders = w_tblPr_.child("w:tblBorders");
    if (!w_tblBorders) {
      w_tblBorders = w_tblPr_.append_child("w:tblBorders");
    }

    auto w_tblBordersChild = w_tblBorders.child(name);
    if (!w_tblBordersChild) {
      w_tblBordersChild = w_tblBorders.append_child(name);
    }

    const char *val;
    switch (style) {
      case BorderStyle::Single:
        val = "single";
        break;
      case BorderStyle::Dotted:
        val = "dotted";
        break;
      case BorderStyle::DotDash:
        val = "dotDash";
        break;
      case BorderStyle::Dashed:
        val = "dashed";
        break;
      case BorderStyle::Double:
        val = "double";
        break;
      case BorderStyle::None:
        val = "none";
        break;
    }

    auto w_val = w_tblBordersChild.attribute("w:val");
    if (!w_val) {
      w_val = w_tblBordersChild.append_attribute("w:val");
    }
    w_val.set_value(val);

    auto w_sz = w_tblBordersChild.attribute("w:sz");
    if (!w_sz) {
      w_sz = w_tblBordersChild.append_attribute("w:sz");
    }
    w_sz.set_value(width * 8);

    auto w_color = w_tblBordersChild.attribute("w:color");
    if (!w_color) {
      w_color = w_tblBordersChild.append_attribute("w:color");
    }
    w_color.set_value(color);
  }

  // class TableCell
  TableCell::TableCell()
  {}

  TableCell::TableCell(const int row, 
                       const int col, 
                       pugi::xml_node w_tr, 
                       pugi::xml_node w_tc, 
                       pugi::xml_node w_tcPr): row_(row), 
                                               col_(col), 
                                               w_tr_(w_tr), 
                                               w_tc_(w_tc), 
                                               w_tcPr_(w_tcPr)
  {}

  void TableCell::SetWidth(const int w, const char *units)
  {
    auto w_tcW = w_tcPr_.child("w:tcW");
    if (!w_tcW) {
      w_tcW = w_tcPr_.append_child("w:tcW");
    }

    auto w_w = w_tcW.attribute("w:w");
    if (!w_w) {
      w_w = w_tcW.append_attribute("w:w");
    }

    auto w_type = w_tcW.attribute("w:type");
    if (!w_type) {
      w_type = w_tcW.append_attribute("w:type");
    }

    w_w.set_value(w);
    w_type.set_value(units);
  }

  TableCell::operator bool()
  {
    return w_tc_;
  }

  Paragraph TableCell::AppendParagraph()
  {
    auto w_p = w_tc_.append_child("w:p");
    auto w_pPr = w_p.append_child("w:pPr");
    return Paragraph(w_tc_, w_p, w_pPr);
  }

  Paragraph TableCell::FirstParagraph()
  {
    auto w_p = w_tc_.child("w:p");
    auto w_pPr = w_p.child("w:pPr");
    return Paragraph(w_tc_, w_p, w_pPr);
  }

} // namespace docx
