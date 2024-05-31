﻿/**
 * minidocx 0.5.0 - C++ library for creating Microsoft Word Document (.docx).
 * --------------------------------------------------------
 * Copyright (C) 2022-2024, by Xie Zequn (totravel@foxmail.com)
 * Report bugs and download new versions at https://github.com/totravel/minidocx
 */

#include "minidocx.hpp"
#include <algorithm>
#include <cstring> // std::strlen(), std::strcmp()
#include <cstdlib> // std::free()
#include <cctype> // std::isspace()
#include "zip.h"
#include "pugixml.hpp"

 // Raw string literal R is danger removed Borland not supported him
#define _RELS "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"word/document.xml\"/></Relationships>"
#define DOCUMENT_XML "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><w:document xmlns:wpc=\"http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas\" xmlns:cx=\"http://schemas.microsoft.com/office/drawing/2014/chartex\" xmlns:cx1=\"http://schemas.microsoft.com/office/drawing/2015/9/8/chartex\" xmlns:cx2=\"http://schemas.microsoft.com/office/drawing/2015/10/21/chartex\" xmlns:cx3=\"http://schemas.microsoft.com/office/drawing/2016/5/9/chartex\" xmlns:cx4=\"http://schemas.microsoft.com/office/drawing/2016/5/10/chartex\" xmlns:cx5=\"http://schemas.microsoft.com/office/drawing/2016/5/11/chartex\" xmlns:cx6=\"http://schemas.microsoft.com/office/drawing/2016/5/12/chartex\" xmlns:cx7=\"http://schemas.microsoft.com/office/drawing/2016/5/13/chartex\" xmlns:cx8=\"http://schemas.microsoft.com/office/drawing/2016/5/14/chartex\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:aink=\"http://schemas.microsoft.com/office/drawing/2016/ink\" xmlns:am3d=\"http://schemas.microsoft.com/office/drawing/2017/model3d\" xmlns:o=\"urn:schemas-microsoft-com:office:office\" xmlns:oel=\"http://schemas.microsoft.com/office/2019/extlst\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\" xmlns:v=\"urn:schemas-microsoft-com:vml\" xmlns:wp14=\"http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing\" xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" xmlns:w10=\"urn:schemas-microsoft-com:office:word\" xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\" xmlns:w15=\"http://schemas.microsoft.com/office/word/2012/wordml\" xmlns:w16cex=\"http://schemas.microsoft.com/office/word/2018/wordml/cex\" xmlns:w16cid=\"http://schemas.microsoft.com/office/word/2016/wordml/cid\" xmlns:w16=\"http://schemas.microsoft.com/office/word/2018/wordml\" xmlns:w16sdtdh=\"http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash\" xmlns:w16se=\"http://schemas.microsoft.com/office/word/2015/wordml/symex\" xmlns:wpg=\"http://schemas.microsoft.com/office/word/2010/wordprocessingGroup\" xmlns:wpi=\"http://schemas.microsoft.com/office/word/2010/wordprocessingInk\" xmlns:wne=\"http://schemas.microsoft.com/office/word/2006/wordml\" xmlns:wps=\"http://schemas.microsoft.com/office/word/2010/wordprocessingShape\" mc:Ignorable=\"w14 w15 w16se w16cid w16 w16cex w16sdtdh wp14\"><w:body><w:sectPr><w:pgSz w:w=\"11906\" w:h=\"16838\" /><w:pgMar w:top=\"1440\" w:right=\"1800\" w:bottom=\"1440\" w:left=\"1800\" w:header=\"851\" w:footer=\"992\" w:gutter=\"0\" /><w:cols w:space=\"425\" /><w:docGrid w:type=\"lines\" w:linePitch=\"312\" /></w:sectPr></w:body></w:document>"
#define CONTENT_TYPES_XML "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\" /><Default Extension=\"xml\" ContentType=\"application/xml\" /><Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\" /><Override PartName=\"/word/footer1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml\" /><Override PartName=\"/word/settings.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml\" /></Types>"
#define DOCUMENT_XML_RELS "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer\" Target=\"footer1.xml\" /><Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings\" Target=\"settings.xml\" /></Relationships>"
#define FOOTER1_XML "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><w:ftr xmlns:wpc=\"http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas\" xmlns:cx=\"http://schemas.microsoft.com/office/drawing/2014/chartex\" xmlns:cx1=\"http://schemas.microsoft.com/office/drawing/2015/9/8/chartex\" xmlns:cx2=\"http://schemas.microsoft.com/office/drawing/2015/10/21/chartex\" xmlns:cx3=\"http://schemas.microsoft.com/office/drawing/2016/5/9/chartex\" xmlns:cx4=\"http://schemas.microsoft.com/office/drawing/2016/5/10/chartex\" xmlns:cx5=\"http://schemas.microsoft.com/office/drawing/2016/5/11/chartex\" xmlns:cx6=\"http://schemas.microsoft.com/office/drawing/2016/5/12/chartex\" xmlns:cx7=\"http://schemas.microsoft.com/office/drawing/2016/5/13/chartex\" xmlns:cx8=\"http://schemas.microsoft.com/office/drawing/2016/5/14/chartex\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:aink=\"http://schemas.microsoft.com/office/drawing/2016/ink\" xmlns:am3d=\"http://schemas.microsoft.com/office/drawing/2017/model3d\" xmlns:o=\"urn:schemas-microsoft-com:office:office\" xmlns:oel=\"http://schemas.microsoft.com/office/2019/extlst\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\" xmlns:v=\"urn:schemas-microsoft-com:vml\" xmlns:wp14=\"http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing\" xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" xmlns:w10=\"urn:schemas-microsoft-com:office:word\" xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\" xmlns:w15=\"http://schemas.microsoft.com/office/word/2012/wordml\" xmlns:w16cex=\"http://schemas.microsoft.com/office/word/2018/wordml/cex\" xmlns:w16cid=\"http://schemas.microsoft.com/office/word/2016/wordml/cid\" xmlns:w16=\"http://schemas.microsoft.com/office/word/2018/wordml\" xmlns:w16du=\"http://schemas.microsoft.com/office/word/2023/wordml/word16du\" xmlns:w16sdtdh=\"http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash\" xmlns:w16se=\"http://schemas.microsoft.com/office/word/2015/wordml/symex\" xmlns:wpg=\"http://schemas.microsoft.com/office/word/2010/wordprocessingGroup\" xmlns:wpi=\"http://schemas.microsoft.com/office/word/2010/wordprocessingInk\" xmlns:wne=\"http://schemas.microsoft.com/office/word/2006/wordml\" xmlns:wps=\"http://schemas.microsoft.com/office/word/2010/wordprocessingShape\" mc:Ignorable=\"w14 w15 w16se w16cid w16 w16cex w16sdtdh wp14\"><w:p><w:pPr><w:jc w:val=\"center\" /></w:pPr><w:r><w:fldChar w:fldCharType=\"begin\" /></w:r><w:r><w:instrText>PAGE \\* MERGEFORMAT</w:instrText></w:r><w:r><w:fldChar w:fldCharType=\"end\" /></w:r></w:p></w:ftr>"
#define SETTINGS_XML "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><w:settings xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:o=\"urn:schemas-microsoft-com:office:office\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\" xmlns:v=\"urn:schemas-microsoft-com:vml\" xmlns:w10=\"urn:schemas-microsoft-com:office:word\" xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\" xmlns:w15=\"http://schemas.microsoft.com/office/word/2012/wordml\" xmlns:w16cex=\"http://schemas.microsoft.com/office/word/2018/wordml/cex\" xmlns:w16cid=\"http://schemas.microsoft.com/office/word/2016/wordml/cid\" xmlns:w16=\"http://schemas.microsoft.com/office/word/2018/wordml\" xmlns:w16sdtdh=\"http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash\" xmlns:w16se=\"http://schemas.microsoft.com/office/word/2015/wordml/symex\" xmlns:sl=\"http://schemas.openxmlformats.org/schemaLibrary/2006/main\" mc:Ignorable=\"w14 w15 w16se w16cid w16 w16cex w16sdtdh\"></w:settings>"

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


  struct xml_string_writer : pugi::xml_writer
  {
    std::string result;

    virtual void write(const void* data, size_t size)
    {
      result.append(static_cast<const char*>(data), size);
    }
  };

  pugi::xml_node GetLastChild(pugi::xml_node node, const char* name)
  {
    pugi::xml_node child = node.last_child();
    while (!child.empty() && std::strcmp(name, child.name()) != 0) {
      child = child.previous_sibling(name);
    }
    return child;
  }


  void SetBorders(pugi::xml_node& w_bdrs, const char* elemName, const Box::BorderStyle style, const double width, const char* color)
  {
    pugi::xml_node w_bdr = w_bdrs.child(elemName);
    if (!w_bdr) {
      w_bdr = w_bdrs.append_child(elemName);
    }

    const char* val;
    switch (style) {
    case Box::BorderStyle::Single:
      val = "single";
      break;
    case Box::BorderStyle::Dotted:
      val = "dotted";
      break;
    case Box::BorderStyle::DotDash:
      val = "dotDash";
      break;
    case Box::BorderStyle::Dashed:
      val = "dashed";
      break;
    case Box::BorderStyle::Double:
      val = "double";
      break;
    case Box::BorderStyle::Wave:
      val = "wave";
      break;
    case Box::BorderStyle::None:
      val = "none";
      break;
    }

    pugi::xml_attribute w_val = w_bdr.attribute("w:val");
    if (!w_val) {
      w_val = w_bdr.append_attribute("w:val");
    }
    w_val.set_value(val);

    pugi::xml_attribute w_sz = w_bdr.attribute("w:sz");
    if (!w_sz) {
      w_sz = w_bdr.append_attribute("w:sz");
    }
    w_sz.set_value(width * 8);

    pugi::xml_attribute w_color = w_bdr.attribute("w:color");
    if (!w_color) {
      w_color = w_bdr.append_attribute("w:color");
    }
    w_color.set_value(color);
  }


  struct Document::Impl
  {
    pugi::xml_document doc_;
    pugi::xml_node     w_body_;
    pugi::xml_node     w_sectPr_;

    pugi::xml_document settings_;
    pugi::xml_node     w_settings_;

    unsigned int nextBookmarkId_;
    std::vector<Bookmark> bookmarks_;
  };

  struct Bookmark::Impl
  {
    unsigned int id_;
    std::string name_;
    pugi::xml_node w_bookmarkStart_;
    pugi::xml_node w_bookmarkEnd_;
  };

  struct Paragraph::Impl
  {
    pugi::xml_node w_body_;
    pugi::xml_node w_p_;
    pugi::xml_node w_pPr_;
  };

  struct TextFrame::Impl
  {
    pugi::xml_node w_framePr_;
  };

  struct Section::Impl
  {
    pugi::xml_node w_body_;
    pugi::xml_node w_p_;      // current paragraph
    pugi::xml_node w_pPr_;
    pugi::xml_node w_p_last_; // the last paragraph of the section
    pugi::xml_node w_pPr_last_;
    pugi::xml_node w_sectPr_;
  };

  struct Table::Impl
  {
    pugi::xml_node w_body_;
    pugi::xml_node w_tbl_;
    pugi::xml_node w_tblPr_;
    pugi::xml_node w_tblGrid_;

    int rows_;
    int cols_;
    Grid grid_; // logical grid
    Impl() : rows_(0), cols_(0) {}
  };

  struct Run::Impl
  {
    pugi::xml_node w_p_;
    pugi::xml_node w_r_;
    pugi::xml_node w_rPr_;
  };

  struct TableCell::Impl
  {
    Cell* c_;
    pugi::xml_node w_tr_;
    pugi::xml_node w_tc_;
    pugi::xml_node w_tcPr_;
  };


  std::ostream& operator<<(std::ostream& out, const Document& doc)
  {
    if (doc.impl_) {
      xml_string_writer writer;
      doc.impl_->w_body_.print(writer, "  ");
      out << writer.result;
    }
    else {
      out << "<document />";
    }
    return out;
  }

  std::ostream& operator<<(std::ostream& out, const Paragraph& p)
  {
    if (p.impl_) {
      xml_string_writer writer;
      p.impl_->w_p_.print(writer, "  ");
      out << writer.result;
    }
    else {
      out << "<paragraph />";
    }
    return out;
  }

  std::ostream& operator<<(std::ostream& out, const Run& r)
  {
    if (r.impl_) {
      xml_string_writer writer;
      r.impl_->w_r_.print(writer, "  ");
      out << writer.result;
    }
    else {
      out << "<run />";
    }
    return out;
  }

  std::ostream& operator<<(std::ostream& out, const Section& s)
  {
    if (s.impl_) {
      xml_string_writer writer;
      s.impl_->w_p_.print(writer, "  ");
      out << writer.result;
    }
    else {
      out << "<section />";
    }
    return out;
  }


  // class Bookmark
  Bookmark::Bookmark() : impl_(NULL)
  {

  }

  Bookmark::Bookmark(Impl* impl) : impl_(impl)
  {

  }

  Bookmark::Bookmark(const Bookmark& rhs)
  {
    impl_ = new Impl;
    impl_->id_ = rhs.impl_->id_;
    impl_->name_ = rhs.impl_->name_;
    impl_->w_bookmarkStart_ = rhs.impl_->w_bookmarkStart_;
    impl_->w_bookmarkEnd_ = rhs.impl_->w_bookmarkEnd_;
  }

  Bookmark::~Bookmark()
  {
    if (impl_ != NULL) {
      delete impl_;
      impl_ = NULL;
    }
  }

  bool Bookmark::operator==(const Bookmark& rhs)
  {
    if (!impl_ && !rhs.impl_) return true;
    if (impl_ && rhs.impl_) return impl_->id_ == rhs.impl_->id_;
    return false;
  }

  unsigned int Bookmark::GetId() const
  {
    return impl_->id_;
  }

  std::string Bookmark::GetName() const
  {
    return impl_->name_;
  }


  // class Document
  Document::Document()
  {
    impl_ = new Impl;
    impl_->doc_.load_buffer(DOCUMENT_XML, std::strlen(DOCUMENT_XML), pugi::parse_declaration);
    impl_->w_body_ = impl_->doc_.child("w:document").child("w:body");
    impl_->w_sectPr_ = impl_->w_body_.child("w:sectPr");
    impl_->settings_.load_buffer(SETTINGS_XML, std::strlen(SETTINGS_XML), pugi::parse_declaration);
    impl_->w_settings_ = impl_->settings_.child("w:settings");
    impl_->nextBookmarkId_ = 0;
  }

  Document::~Document()
  {
    if (impl_ != NULL) {
      delete impl_;
      impl_ = NULL;
    }
  }

  bool Document::Save(const std::string& path)
  {
    if (!impl_) return false;

    struct zip_t* zip = zip_open(path.c_str(), ZIP_DEFAULT_COMPRESSION_LEVEL, 'w');
    if (zip == NULL) {
      return false;
    }

    xml_string_writer writer;
    const char* buf = NULL;

    zip_entry_open(zip, "_rels/.rels");
    zip_entry_write(zip, _RELS, std::strlen(_RELS));
    zip_entry_close(zip);

    zip_entry_open(zip, "word/document.xml");
    impl_->doc_.save(writer, "", pugi::format_raw);
    buf = writer.result.c_str();
    zip_entry_write(zip, buf, std::strlen(buf));
    zip_entry_close(zip);

    zip_entry_open(zip, "word/settings.xml");
    writer.result = "";
    impl_->settings_.save(writer, "", pugi::format_raw);
    buf = writer.result.c_str();
    zip_entry_write(zip, buf, std::strlen(buf));
    zip_entry_close(zip);

    zip_entry_open(zip, "word/_rels/document.xml.rels");
    zip_entry_write(zip, DOCUMENT_XML_RELS, std::strlen(DOCUMENT_XML_RELS));
    zip_entry_close(zip);

    zip_entry_open(zip, "word/footer1.xml");
    zip_entry_write(zip, FOOTER1_XML, std::strlen(FOOTER1_XML));
    zip_entry_close(zip);

    zip_entry_open(zip, "[Content_Types].xml");
    zip_entry_write(zip, CONTENT_TYPES_XML, std::strlen(CONTENT_TYPES_XML));
    zip_entry_close(zip);

    zip_close(zip);
    return true;
  }

  bool Document::Open(const std::string& path)
  {
    if (!impl_) return false;

    struct zip_t* zip = zip_open(path.c_str(), ZIP_DEFAULT_COMPRESSION_LEVEL, 'r');
    if (zip == NULL) {
      return false;
    }

    if (zip_entry_open(zip, "word/document.xml") < 0) {
      zip_close(zip);
      return false;
    }
    void* buf = NULL;
    size_t bufsize;
    zip_entry_read(zip, &buf, &bufsize);
    zip_entry_close(zip);

    impl_->doc_.load_buffer(buf, bufsize, pugi::parse_declaration);
    impl_->w_body_ = impl_->doc_.child("w:document").child("w:body");
    impl_->w_sectPr_ = impl_->w_body_.child("w:sectPr");
    std::free(buf);

    if (zip_entry_open(zip, "word/settings.xml") >= 0) {
      zip_entry_read(zip, &buf, &bufsize);
      zip_entry_close(zip);

      impl_->settings_.load_buffer(buf, bufsize, pugi::parse_declaration);
      impl_->w_settings_ = impl_->settings_.child("w:settings");
      std::free(buf);
    }

    zip_close(zip);
    FindBookmarks();
    return true;
  }

  Paragraph Document::FirstParagraph()
  {
    if (!impl_) return Paragraph();
    pugi::xml_node w_p = impl_->w_body_.child("w:p");
    if (!w_p) return Paragraph();

    Paragraph::Impl* impl = new Paragraph::Impl;
    impl->w_body_ = impl_->w_body_;
    impl->w_p_ = w_p;
    impl->w_pPr_ = w_p.child("w:pPr");
    return Paragraph(impl);
  }

  Paragraph Document::LastParagraph()
  {
    if (!impl_) return Paragraph();
    pugi::xml_node w_p = GetLastChild(impl_->w_body_, "w:p");
    if (!w_p) return Paragraph();

    Paragraph::Impl* impl = new Paragraph::Impl;
    impl->w_body_ = impl_->w_body_;
    impl->w_p_ = w_p;
    impl->w_pPr_ = w_p.child("w:pPr");
    return Paragraph(impl);
  }

  Section Document::FirstSection()
  {
    Paragraph firstParagraph = FirstParagraph();
    if (firstParagraph) return firstParagraph.GetSection();

    Section::Impl* impl = new Section::Impl;
    impl->w_body_ = impl_->w_body_;
    Section section(impl);
    section.FindSectionProperties();
    return section;
  }

  Section Document::LastSection()
  {
    Paragraph lastParagraph = LastParagraph();
    if (lastParagraph) return lastParagraph.GetSection();

    Section::Impl* impl = new Section::Impl;
    impl->w_body_ = impl_->w_body_;
    Section section(impl);
    section.FindSectionProperties();
    return section;
  }

  Paragraph Document::AppendParagraph()
  {
    if (!impl_) return Paragraph();

    pugi::xml_node w_p = impl_->w_body_.insert_child_before("w:p", impl_->w_sectPr_);
    pugi::xml_node w_pPr = w_p.append_child("w:pPr");

    Paragraph::Impl* impl = new Paragraph::Impl;
    impl->w_body_ = impl_->w_body_;
    impl->w_p_ = w_p;
    impl->w_pPr_ = w_pPr;
    return Paragraph(impl);
  }

  Paragraph Document::AppendParagraph(const std::string& text)
  {
    Paragraph p = AppendParagraph();
    p.AppendRun(text);
    return p;
  }

  Paragraph Document::AppendParagraph(const std::string& text,
    const double fontSize)
  {
    Paragraph p = AppendParagraph();
    p.AppendRun(text, fontSize);
    return p;
  }

  Paragraph Document::AppendParagraph(const std::string& text,
    const double fontSize,
    const std::string& fontAscii,
    const std::string& fontEastAsia)
  {
    Paragraph p = AppendParagraph();
    p.AppendRun(text, fontSize, fontAscii, fontEastAsia);
    return p;
  }

  Paragraph Document::PrependParagraph()
  {
    if (!impl_) return Paragraph();

    pugi::xml_node w_p = impl_->w_body_.prepend_child("w:p");
    pugi::xml_node w_pPr = w_p.append_child("w:pPr");

    Paragraph::Impl* impl = new Paragraph::Impl;
    impl->w_body_ = impl_->w_body_;
    impl->w_p_ = w_p;
    impl->w_pPr_ = w_pPr;
    return Paragraph(impl);
  }

  Paragraph Document::PrependParagraph(const std::string& text)
  {
    Paragraph p = PrependParagraph();
    p.AppendRun(text);
    return p;
  }

  Paragraph Document::PrependParagraph(const std::string& text,
    const double fontSize)
  {
    Paragraph p = PrependParagraph();
    p.AppendRun(text, fontSize);
    return p;
  }

  Paragraph Document::PrependParagraph(const std::string& text,
    const double fontSize,
    const std::string& fontAscii,
    const std::string& fontEastAsia)
  {
    Paragraph p = PrependParagraph();
    p.AppendRun(text, fontSize, fontAscii, fontEastAsia);
    return p;
  }

  Paragraph Document::InsertParagraphBefore(const Paragraph& p)
  {
    if (!impl_) return Paragraph();

    pugi::xml_node w_p = impl_->w_body_.insert_child_before("w:p", p.impl_->w_p_);
    pugi::xml_node w_pPr = w_p.append_child("w:pPr");

    Paragraph::Impl* impl = new Paragraph::Impl;
    impl->w_body_ = impl_->w_body_;
    impl->w_p_ = w_p;
    impl->w_pPr_ = w_pPr;
    return Paragraph(impl);
  }

  Paragraph Document::InsertParagraphAfter(const Paragraph& p)
  {
    if (!impl_) return Paragraph();

    pugi::xml_node w_p = impl_->w_body_.insert_child_after("w:p", p.impl_->w_p_);
    pugi::xml_node w_pPr = w_p.append_child("w:pPr");

    Paragraph::Impl* impl = new Paragraph::Impl;
    impl->w_body_ = impl_->w_body_;
    impl->w_p_ = w_p;
    impl->w_pPr_ = w_pPr;
    return Paragraph(impl);
  }

  bool Document::RemoveParagraph(Paragraph& p)
  {
    if (!impl_) return false;
    return impl_->w_body_.remove_child(p.impl_->w_p_);
  }

  Paragraph Document::AppendPageBreak()
  {
    Paragraph p = AppendParagraph();
    p.AppendPageBreak();
    return p;
  }

  Paragraph Document::AppendSectionBreak()
  {
    Paragraph p = AppendParagraph();
    p.InsertSectionBreak();
    return p;
  }

  Table Document::AppendTable(const int rows, const int cols)
  {
    if (!impl_) return Table();

    pugi::xml_node w_tbl = impl_->w_body_.insert_child_before("w:tbl", impl_->w_sectPr_);
    pugi::xml_node w_tblPr = w_tbl.append_child("w:tblPr");
    pugi::xml_node w_tblGrid = w_tbl.append_child("w:tblGrid");

    Table::Impl* impl = new Table::Impl;
    impl->w_body_ = impl_->w_body_;
    impl->w_tbl_ = w_tbl;
    impl->w_tblPr_ = w_tblPr;
    impl->w_tblGrid_ = w_tblGrid;
    Table tbl(impl);

    tbl.Create_(rows, cols);
    tbl.SetAllBorders();
    tbl.SetWidthPercent(100);
    tbl.SetCellMarginLeft(CM2Twip(0.19));
    tbl.SetCellMarginRight(CM2Twip(0.19));
    tbl.SetAlignment(Table::Alignment::Centered);
    return tbl;
  }

  void Document::RemoveTable(Table& tbl)
  {
    if (!impl_) return;
    impl_->w_body_.remove_child(tbl.impl_->w_tbl_);
  }

  TextFrame Document::AppendTextFrame(const int w, const int h)
  {
    if (!impl_) return TextFrame();

    pugi::xml_node w_p = impl_->w_body_.insert_child_before("w:p", impl_->w_sectPr_);
    pugi::xml_node w_pPr = w_p.append_child("w:pPr");
    pugi::xml_node w_framePr = w_pPr.append_child("w:framePr");


    Paragraph::Impl* paragraph_impl = new Paragraph::Impl;
    paragraph_impl->w_body_ = impl_->w_body_;
    paragraph_impl->w_p_ = w_p;
    paragraph_impl->w_pPr_ = w_pPr;

    TextFrame::Impl* impl = new TextFrame::Impl;
    impl->w_framePr_ = w_framePr;
    TextFrame textFrame = TextFrame(impl, paragraph_impl);

    textFrame.SetSize(w, h);
    textFrame.SetBorders();
    return textFrame;
  }

  void Document::SetReadOnly(const bool enabled)
  {
    if (!impl_) return;

    if (!enabled) {
      pugi::xml_node documentProtection = impl_->w_settings_.child("w:documentProtection");
      if (documentProtection) {
        impl_->w_settings_.remove_child(documentProtection);
      }
      return;
    }

    pugi::xml_node documentProtection = impl_->w_settings_.child("w:documentProtection");
    if (!documentProtection) {
      documentProtection = impl_->w_settings_.append_child("w:documentProtection");
    }

    pugi::xml_attribute documentProtectionEdit = documentProtection.attribute("w:edit");
    if (!documentProtectionEdit) {
      documentProtectionEdit = documentProtection.append_attribute("w:edit");
    }
    documentProtectionEdit.set_value("readOnly");

    pugi::xml_attribute documentProtectionEnforcement = documentProtection.attribute("w:enforcement");
    if (!documentProtectionEnforcement) {
      documentProtectionEnforcement = documentProtection.append_attribute("w:enforcement");
    }
    documentProtectionEnforcement.set_value("1");
  }

  std::map<std::string, std::string> Document::GetVars()
  {
    std::map<std::string, std::string> vars;
    if (!impl_) return vars;

    pugi::xml_node docVars = impl_->w_settings_.child("w:docVars");
    if (!docVars) {
      return vars;
    }

    for (pugi::xml_node p = docVars.child("w:docVar"); p; p = p.next_sibling("w:docVar")) {
      vars.insert(std::pair<std::string, std::string>(p.attribute("w:name").value(), p.attribute("w:val").value()));
    }

    return vars;
  }

  void Document::SetVars(const std::map<std::string, std::string>& vars)
  {
    if (!impl_) return;

    pugi::xml_node docVars = impl_->w_settings_.child("w:docVars");
    if (!docVars) {
      docVars = impl_->w_settings_.append_child("w:docVars");
    }
    else {
      docVars.remove_children();
    }

    std::map<std::string, std::string>::const_iterator it;
    for (it = vars.begin(); it != vars.end(); it++) {
      pugi::xml_node docVar = docVars.append_child("w:docVar");
      docVar.append_attribute("w:name") = it->first.c_str();
      docVar.append_attribute("w:val") = it->second.c_str();
    }
  }

  void Document::AddVars(const std::map<std::string, std::string>& vars)
  {
    if (!impl_) return;

    pugi::xml_node docVars = impl_->w_settings_.child("w:docVars");
    if (!docVars) {
      docVars = impl_->w_settings_.append_child("w:docVars");
    }

    std::map<std::string, std::string>::const_iterator it;
    for (it = vars.begin(); it != vars.end(); it++) {
      pugi::xml_node docVar = docVars.append_child("w:docVar");
      docVar.append_attribute("w:name") = it->first.c_str();
      docVar.append_attribute("w:val") = it->second.c_str();
    }
  }

  void Document::FindBookmarks()
  {
    if (!impl_) return;

    unsigned int maxBookmarkId = 0;

    for (pugi::xml_node i = impl_->w_body_.first_child(); i; i = i.next_sibling()) {
      if (std::strcmp(i.name(), "w:p") != 0) continue;

      for (pugi::xml_node j = i.first_child(); j; j = j.next_sibling()) {

        if (std::strcmp(j.name(), "w:bookmarkStart") == 0) {

          Bookmark::Impl* impl = new Bookmark::Impl;
          impl->id_ = j.attribute("w:id").as_uint();
          impl->name_ = j.attribute("w:name").as_string();
          impl->w_bookmarkStart_ = j;

          if (maxBookmarkId < impl->id_)
            maxBookmarkId = impl->id_;

          impl_->bookmarks_.push_back(impl);
        }
        else if (std::strcmp(j.name(), "w:bookmarkEnd") == 0) {

          const unsigned int id = j.attribute("w:id").as_uint();
          for (std::vector<Bookmark>::iterator it = impl_->bookmarks_.begin(); it != impl_->bookmarks_.end(); ++it) {
            if ((*it).GetId() == id) {
              it->impl_->w_bookmarkEnd_ = j;
              break;
            }
          }
        }
        else {
          continue;
        }
      }
    }

    if (impl_->bookmarks_.size() > 0)
      impl_->nextBookmarkId_ = maxBookmarkId++;
  }

  std::vector<Bookmark> Document::GetBookmarks()
  {
    if (!impl_) return std::vector<Bookmark>();
    return impl_->bookmarks_;
  }

  Bookmark Document::AddBookmark(const std::string& name, const Run& start, const Run& end)
  {
    if (!impl_) return Bookmark();

    Bookmark::Impl* impl = new Bookmark::Impl;
    impl->id_ = impl_->nextBookmarkId_++;
    impl->name_ = name;

    pugi::xml_node w_bookmarkStart = start.impl_->w_p_.insert_child_before("w:bookmarkStart", start.impl_->w_r_);
    w_bookmarkStart.append_attribute("w:id") = impl->id_;
    w_bookmarkStart.append_attribute("w:name") = impl->name_.c_str();

    pugi::xml_node w_bookmarkEnd = end.impl_->w_p_.insert_child_after("w:bookmarkEnd", end.impl_->w_r_);
    w_bookmarkEnd.append_attribute("w:id") = impl->id_;

    impl->w_bookmarkStart_ = w_bookmarkStart;
    impl->w_bookmarkStart_ = w_bookmarkEnd;

    Bookmark b(impl);
    impl_->bookmarks_.push_back(b);
    return b;
  }


  void Document::RemoveBookmark(Bookmark& bookmark)
  {
    if (!impl_) return;

    std::vector<Bookmark>::iterator it = std::find(
      impl_->bookmarks_.begin(), impl_->bookmarks_.end(), bookmark);
    if (it != impl_->bookmarks_.end())
      impl_->bookmarks_.erase(it);

    bookmark.impl_->w_bookmarkStart_.parent().remove_child(bookmark.impl_->w_bookmarkStart_);
    bookmark.impl_->w_bookmarkEnd_.parent().remove_child(bookmark.impl_->w_bookmarkEnd_);
  }


  // class Paragraph
  Paragraph::Paragraph() : impl_(NULL)
  {

  }

  Paragraph::Paragraph(Impl* impl) : impl_(impl)
  {

  }

  Paragraph::Paragraph(const Paragraph& p)
  {
    impl_ = new Impl;
    impl_->w_body_ = p.impl_->w_body_;
    impl_->w_p_ = p.impl_->w_p_;
    impl_->w_pPr_ = p.impl_->w_pPr_;
  }

  Paragraph::~Paragraph()
  {
    if (impl_ != NULL) {
      delete impl_;
      impl_ = NULL;
    }
  }

  Run Paragraph::FirstRun()
  {
    if (!impl_) return Run();
    pugi::xml_node w_r = impl_->w_p_.child("w:r");
    if (!w_r) return Run();

    Run::Impl* impl = new Run::Impl;
    impl->w_p_ = impl_->w_p_;
    impl->w_r_ = w_r;
    impl->w_rPr_ = w_r.child("w:rPr");
    return Run(impl);
  }

  Run Paragraph::AppendRun()
  {
    if (!impl_) return Run();
    pugi::xml_node w_r = impl_->w_p_.append_child("w:r");
    pugi::xml_node w_rPr = w_r.append_child("w:rPr");

    Run::Impl* impl = new Run::Impl;
    impl->w_p_ = impl_->w_p_;
    impl->w_r_ = w_r;
    impl->w_rPr_ = w_rPr;
    return Run(impl);
  }

  Run Paragraph::AppendRun(const std::string& text)
  {
    Run r = AppendRun();
    if (!text.empty()) {
      r.AppendText(text);
    }
    return r;
  }

  Run Paragraph::AppendRun(const std::string& text, const double fontSize)
  {
    Run r = AppendRun(text);
    if (fontSize != 0) {
      r.SetFontSize(fontSize);
    }
    return r;
  }

  Run Paragraph::AppendRun(const std::string& text, const double fontSize, const std::string& fontAscii, const std::string& fontEastAsia)
  {
    Run r = AppendRun(text, fontSize);
    if (!fontAscii.empty()) {
      r.SetFont(fontAscii, fontEastAsia);
    }
    return r;
  }

  Run Paragraph::AppendPageBreak()
  {
    if (!impl_) return Run();
    pugi::xml_node w_r = impl_->w_p_.append_child("w:r");
    pugi::xml_node w_br = w_r.append_child("w:br");
    w_br.append_attribute("w:type") = "page";

    Run::Impl* impl = new Run::Impl;
    impl->w_p_ = impl_->w_p_;
    impl->w_r_ = w_r;
    impl->w_rPr_ = w_br;
    return Run(impl);
  }

  void Paragraph::SetAlignment(const Alignment alignment)
  {
    if (!impl_) return;

    const char* val;
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

    pugi::xml_node jc = impl_->w_pPr_.child("w:jc");
    if (!jc) {
      jc = impl_->w_pPr_.append_child("w:jc");
    }
    pugi::xml_attribute jcVal = jc.attribute("w:val");
    if (!jcVal) {
      jcVal = jc.append_attribute("w:val");
    }
    jcVal.set_value(val);
  }

  void Paragraph::SetLineSpacingSingle()
  {
    SetLineSpacing(240, "auto");
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

  void Paragraph::SetLineSpacing(const int at, const char* lineRule)
  {
    if (!impl_) return;
    pugi::xml_node spacing = impl_->w_pPr_.child("w:spacing");
    if (!spacing) {
      spacing = impl_->w_pPr_.append_child("w:spacing");
    }

    pugi::xml_attribute spacingLineRule = spacing.attribute("w:lineRule");
    if (!spacingLineRule) {
      spacingLineRule = spacing.append_attribute("w:lineRule");
    }

    pugi::xml_attribute spacingLine = spacing.attribute("w:line");
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

  void Paragraph::SetSpacingAuto(const char* attrNameAuto)
  {
    if (!impl_) return;
    pugi::xml_node spacing = impl_->w_pPr_.child("w:spacing");
    if (!spacing) {
      spacing = impl_->w_pPr_.append_child("w:spacing");
    }
    pugi::xml_attribute spacingAuto = spacing.attribute(attrNameAuto);
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

  void Paragraph::SetSpacing(const int twip, const char* attrNameAuto, const char* attrName)
  {
    if (!impl_) return;
    pugi::xml_node elemSpacing = impl_->w_pPr_.child("w:spacing");
    if (!elemSpacing) {
      elemSpacing = impl_->w_pPr_.append_child("w:spacing");
    }

    pugi::xml_attribute attrSpacingAuto = elemSpacing.attribute(attrNameAuto);
    if (attrSpacingAuto) {
      elemSpacing.remove_attribute(attrSpacingAuto);
    }

    pugi::xml_attribute attrSpacing = elemSpacing.attribute(attrName);
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

  void Paragraph::SetIndent(const int indent, const char* attrName)
  {
    if (!impl_) return;
    pugi::xml_node elemIndent = impl_->w_pPr_.child("w:ind");
    if (!elemIndent) {
      elemIndent = impl_->w_pPr_.append_child("w:ind");
    }

    pugi::xml_attribute attrIndent = elemIndent.attribute(attrName);
    if (!attrIndent) {
      attrIndent = elemIndent.append_attribute(attrName);
    }
    attrIndent.set_value(indent);
  }

  void Paragraph::SetTopBorder(const BorderStyle style, const double width, const char* color)
  {
    SetBorders_("w:top", style, width, color);
  }

  void Paragraph::SetBottomBorder(const BorderStyle style, const double width, const char* color)
  {
    SetBorders_("w:bottom", style, width, color);
  }

  void Paragraph::SetLeftBorder(const BorderStyle style, const double width, const char* color)
  {
    SetBorders_("w:left", style, width, color);
  }

  void Paragraph::SetRightBorder(const BorderStyle style, const double width, const char* color)
  {
    SetBorders_("w:right", style, width, color);
  }

  void Paragraph::SetBorders(const BorderStyle style, const double width, const char* color)
  {
    SetTopBorder(style, width, color);
    SetBottomBorder(style, width, color);
    SetLeftBorder(style, width, color);
    SetRightBorder(style, width, color);
  }

  void Paragraph::SetBorders_(const char* elemName, const BorderStyle style, const double width, const char* color)
  {
    if (!impl_) return;
    pugi::xml_node w_pBdr = impl_->w_pPr_.child("w:pBdr");
    if (!w_pBdr) {
      w_pBdr = impl_->w_pPr_.append_child("w:pBdr");
    }
    docx::SetBorders(w_pBdr, elemName, style, width, color);
  }

  void Paragraph::SetFontSize(const double fontSize)
  {
    for (Run r = FirstRun(); r; r = r.Next()) {
      r.SetFontSize(fontSize);
    }
  }

  void Paragraph::SetFont(const std::string& fontAscii, const std::string& fontEastAsia)
  {
    for (Run r = FirstRun(); r; r = r.Next()) {
      r.SetFont(fontAscii, fontEastAsia);
    }
  }

  void Paragraph::SetFontStyle(const Run::FontStyle fontStyle)
  {
    for (Run r = FirstRun(); r; r = r.Next()) {
      r.SetFontStyle(fontStyle);
    }
  }

  void Paragraph::SetCharacterSpacing(const int characterSpacing)
  {
    for (Run r = FirstRun(); r; r = r.Next()) {
      r.SetCharacterSpacing(characterSpacing);
    }
  }

  void Paragraph::SetFontColor(const std::string &color)
  {
    for (Run r = FirstRun();r ;r = r.Next()) {
        r.SetFontColor(color);
    }
  }

  std::string Paragraph::GetText()
  {
    std::string text;
    for (Run r = FirstRun(); r; r = r.Next()) {
      text += r.GetText();
    }
    return text;
  }

  Paragraph Paragraph::Next()
  {
    if (!impl_) return Paragraph();
    pugi::xml_node w_p = impl_->w_p_.next_sibling("w:p");
    if (!w_p) return Paragraph();

    Paragraph::Impl* impl = new Paragraph::Impl;
    impl->w_body_ = impl_->w_body_;
    impl->w_p_ = w_p;
    impl->w_pPr_ = w_p.child("w:pPr");
    return Paragraph(impl);
  }

  Paragraph Paragraph::Prev()
  {
    if (!impl_) return Paragraph();
    pugi::xml_node w_p = impl_->w_p_.previous_sibling("w:p");
    if (!w_p) return Paragraph();

    Paragraph::Impl* impl = new Paragraph::Impl;
    impl->w_body_ = impl_->w_body_;
    impl->w_p_ = w_p;
    impl->w_pPr_ = w_p.child("w:pPr");
    return Paragraph(impl);
  }

  Paragraph::operator bool()
  {
    return impl_ != NULL && impl_->w_p_;
  }

  bool Paragraph::operator==(const Paragraph& p)
  {
    if (!impl_ && !p.impl_) return true;
    if (impl_ && p.impl_) return impl_->w_p_ == p.impl_->w_p_;
    return false;
  }

  void Paragraph::operator=(const Paragraph& right)
  {
    if (this == &right) return;
    if (impl_ != NULL) delete impl_;
    if (right.impl_ != NULL) {
      impl_ = new Impl;
      impl_->w_body_ = right.impl_->w_body_;
      impl_->w_p_ = right.impl_->w_p_;
      impl_->w_pPr_ = right.impl_->w_pPr_;
    }
    else {
      impl_ = NULL;
    }
  }

  Section Paragraph::GetSection()
  {
    if (!impl_) return Section();
    Section::Impl* impl = new Section::Impl;
    impl->w_body_ = impl_->w_body_;
    impl->w_p_ = impl_->w_p_;
    impl->w_pPr_ = impl_->w_pPr_;
    Section section(impl);
    section.FindSectionProperties();
    return section;
  }

  Section Paragraph::InsertSectionBreak()
  {
    Section section = GetSection();
    // this paragraph will be the last paragraph of the new section
    if (section) section.Split();
    return section;
  }

  Section Paragraph::RemoveSectionBreak()
  {
    Section section = GetSection();
    if (section && section.IsSplit()) section.Merge();
    return section;
  }

  bool Paragraph::HasSectionBreak()
  {
    return GetSection().IsSplit();
  }


  // class Section
  Section::Section() : impl_(NULL)
  {

  }

  Section::Section(Impl* impl) : impl_(impl)
  {

  }

  Section::Section(const Section& s)
  {
    impl_ = new Impl;
    impl_->w_body_ = s.impl_->w_body_;
    impl_->w_p_ = s.impl_->w_p_;
    impl_->w_p_last_ = s.impl_->w_p_last_;
    impl_->w_pPr_ = s.impl_->w_pPr_;
    impl_->w_pPr_last_ = s.impl_->w_pPr_last_;
    impl_->w_sectPr_ = s.impl_->w_sectPr_;
  }

  Section::~Section()
  {
    if (impl_ != NULL) {
      delete impl_;
      impl_ = NULL;
    }
  }

  void Section::FindSectionProperties()
  {
    if (!impl_) return;
    pugi::xml_node w_p_next = impl_->w_p_, w_p, w_pPr, w_sectPr;
    do {
      w_p = w_p_next;
      w_pPr = w_p.child("w:pPr");
      w_sectPr = w_pPr.child("w:sectPr");

      w_p_next = w_p.next_sibling("w:p");
    } while (w_sectPr.empty() && !w_p_next.empty());

    impl_->w_p_last_ = w_p;
    impl_->w_pPr_last_ = w_pPr;
    impl_->w_sectPr_ = w_sectPr;

    if (impl_->w_sectPr_.empty()) {
      impl_->w_sectPr_ = impl_->w_body_.child("w:sectPr");
    }
  }

  void Section::Split()
  {
    if (!impl_) return;
    if (IsSplit()) return;
    impl_->w_p_last_ = impl_->w_p_;
    impl_->w_pPr_last_ = impl_->w_pPr_;
    impl_->w_sectPr_ = impl_->w_pPr_.append_copy(impl_->w_sectPr_);
  }

  bool Section::IsSplit()
  {
    if (!impl_) return false;
    return impl_->w_pPr_.child("w:sectPr");
  }

  void Section::Merge()
  {
    if (!impl_) return;
    if (impl_->w_pPr_.child("w:sectPr").empty()) return;
    impl_->w_pPr_last_.remove_child(impl_->w_sectPr_);
    FindSectionProperties();
  }

  void Section::SetPageSize(const int w, const int h)
  {
    if (!impl_) return;
    pugi::xml_node pgSz = impl_->w_sectPr_.child("w:pgSz");
    if (!pgSz) {
      pgSz = impl_->w_sectPr_.append_child("w:pgSz");
    }
    pugi::xml_attribute pgSzW = pgSz.attribute("w:w");
    if (!pgSzW) {
      pgSzW = pgSz.append_attribute("w:w");
    }
    pugi::xml_attribute pgSzH = pgSz.attribute("w:h");
    if (!pgSzH) {
      pgSzH = pgSz.append_attribute("w:h");
    }
    pgSzW.set_value(w);
    pgSzH.set_value(h);
  }

  void Section::GetPageSize(int& w, int& h)
  {
    if (!impl_) return;
    w = h = 0;
    pugi::xml_node pgSz = impl_->w_sectPr_.child("w:pgSz");
    if (!pgSz) return;
    pugi::xml_attribute pgSzW = pgSz.attribute("w:w");
    if (!pgSzW) return;
    pugi::xml_attribute pgSzH = pgSz.attribute("w:h");
    if (!pgSzH) return;
    w = pgSzW.as_int();
    h = pgSzH.as_int();
  }

  void Section::SetPageOrient(const Orientation orient)
  {
    if (!impl_) return;
    pugi::xml_node pgSz = impl_->w_sectPr_.child("w:pgSz");
    if (!pgSz) {
      pgSz = impl_->w_sectPr_.append_child("w:pgSz");
    }
    pugi::xml_attribute pgSzH = pgSz.attribute("w:h");
    if (!pgSzH) {
      pgSzH = pgSz.append_attribute("w:h");
    }
    pugi::xml_attribute pgSzW = pgSz.attribute("w:w");
    if (!pgSzW) {
      pgSzW = pgSz.append_attribute("w:w");
    }
    pugi::xml_attribute pgSzOrient = pgSz.attribute("w:orient");
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
    if (!impl_) return Orientation::Unknown;
    Orientation orient = Orientation::Portrait;
    pugi::xml_node pgSz = impl_->w_sectPr_.child("w:pgSz");
    if (!pgSz) return orient;
    pugi::xml_attribute pgSzOrient = pgSz.attribute("w:orient");
    if (!pgSzOrient) return orient;
    if (std::string(pgSzOrient.value()).compare("landscape") == 0) {
      orient = Orientation::Landscape;
    }
    return orient;
  }

  void Section::SetPageMargin(const int top, const int bottom,
    const int left, const int right)
  {
    if (!impl_) return;
    pugi::xml_node pgMar = impl_->w_sectPr_.child("w:pgMar");
    if (!pgMar) {
      pgMar = impl_->w_sectPr_.append_child("w:pgMar");
    }
    pugi::xml_attribute pgMarTop = pgMar.attribute("w:top");
    if (!pgMarTop) {
      pgMarTop = pgMar.append_attribute("w:top");
    }
    pugi::xml_attribute pgMarBottom = pgMar.attribute("w:bottom");
    if (!pgMarBottom) {
      pgMarBottom = pgMar.append_attribute("w:bottom");
    }
    pugi::xml_attribute pgMarLeft = pgMar.attribute("w:left");
    if (!pgMarLeft) {
      pgMarLeft = pgMar.append_attribute("w:left");
    }
    pugi::xml_attribute pgMarRight = pgMar.attribute("w:right");
    if (!pgMarRight) {
      pgMarRight = pgMar.append_attribute("w:right");
    }
    pgMarTop.set_value(top);
    pgMarBottom.set_value(bottom);
    pgMarLeft.set_value(left);
    pgMarRight.set_value(right);
  }

  void Section::GetPageMargin(int& top, int& bottom,
    int& left, int& right)
  {
    if (!impl_) return;
    top = bottom = left = right = 0;
    pugi::xml_node pgMar = impl_->w_sectPr_.child("w:pgMar");
    if (!pgMar) return;
    pugi::xml_attribute pgMarTop = pgMar.attribute("w:top");
    if (!pgMarTop) return;
    pugi::xml_attribute pgMarBottom = pgMar.attribute("w:bottom");
    if (!pgMarBottom) return;
    pugi::xml_attribute pgMarLeft = pgMar.attribute("w:left");
    if (!pgMarLeft) return;
    pugi::xml_attribute pgMarRight = pgMar.attribute("w:right");
    if (!pgMarRight) return;
    top = pgMarTop.as_int();
    bottom = pgMarBottom.as_int();
    left = pgMarLeft.as_int();
    right = pgMarRight.as_int();
  }

  void Section::SetPageMargin(const int header, const int footer)
  {
    if (!impl_) return;
    pugi::xml_node pgMar = impl_->w_sectPr_.child("w:pgMar");
    if (!pgMar) {
      pgMar = impl_->w_sectPr_.append_child("w:pgMar");
    }
    pugi::xml_attribute pgMarHeader = pgMar.attribute("w:header");
    if (!pgMarHeader) {
      pgMarHeader = pgMar.append_attribute("w:header");
    }
    pugi::xml_attribute pgMarFooter = pgMar.attribute("w:footer");
    if (!pgMarFooter) {
      pgMarFooter = pgMar.append_attribute("w:footer");
    }
    pgMarHeader.set_value(header);
    pgMarFooter.set_value(footer);
  }

  void Section::GetPageMargin(int& header, int& footer)
  {
    if (!impl_) return;
    header = footer = 0;
    pugi::xml_node pgMar = impl_->w_sectPr_.child("w:pgMar");
    if (!pgMar) return;
    pugi::xml_attribute pgMarHeader = pgMar.attribute("w:header");
    if (!pgMarHeader) return;
    pugi::xml_attribute pgMarFooter = pgMar.attribute("w:footer");
    if (!pgMarFooter) return;
    header = pgMarHeader.as_int();
    footer = pgMarFooter.as_int();
  }

  void Section::SetPageNumber(const PageNumberFormat fmt, const unsigned int start)
  {
    if (!impl_) return;

    pugi::xml_node footerReference = impl_->w_sectPr_.child("w:footerReference");
    if (!footerReference) {
      footerReference = impl_->w_sectPr_.append_child("w:footerReference");
    }

    pugi::xml_attribute footerReferenceId = footerReference.attribute("r:id");
    if (!footerReferenceId) {
      footerReferenceId = footerReference.append_attribute("r:id");
    }
    footerReferenceId.set_value("rId1");

    pugi::xml_attribute footerReferenceType = footerReference.attribute("w:type");
    if (!footerReferenceType) {
      footerReferenceType = footerReference.append_attribute("w:type");
    }
    footerReferenceType.set_value("default");

    pugi::xml_node pgNumType = impl_->w_sectPr_.child("w:pgNumType");
    if (!pgNumType) {
      pgNumType = impl_->w_sectPr_.append_child("w:pgNumType");
    }

    pugi::xml_attribute pgNumTypeFmt = pgNumType.attribute("w:fmt");
    if (!pgNumTypeFmt) {
      pgNumTypeFmt = pgNumType.append_attribute("w:fmt");
    }
    const char* fmtVal = "";
    switch (fmt) {
    case PageNumberFormat::Decimal:
      fmtVal = "decimal";
      break;
    case PageNumberFormat::NumberInDash:
      fmtVal = "numberInDash";
      break;
    case PageNumberFormat::CardinalText:
      fmtVal = "cardinalText";
      break;
    case PageNumberFormat::OrdinalText:
      fmtVal = "ordinalText";
      break;
    case PageNumberFormat::LowerLetter:
      fmtVal = "lowerLetter";
      break;
    case PageNumberFormat::UpperLetter:
      fmtVal = "upperLetter";
      break;
    case PageNumberFormat::LowerRoman:
      fmtVal = "lowerRoman";
      break;
    case PageNumberFormat::UpperRoman:
      fmtVal = "upperRoman";
      break;
    }
    pgNumTypeFmt.set_value(fmtVal);

    pugi::xml_attribute pgNumTypeStart = pgNumType.attribute("w:start");
    if (start > 0) {
      if (!pgNumTypeStart) {
        pgNumTypeStart = pgNumType.append_attribute("w:start");
        pgNumTypeStart.set_value(start);
      }
    }
    else {
      if (pgNumTypeStart) {
        pgNumType.remove_attribute(pgNumTypeStart);
      }
    }
  }

  void Section::RemovePageNumber()
  {
    if (!impl_) return;

    pugi::xml_node footerReference = impl_->w_sectPr_.child("w:footerReference");
    if (footerReference) {
      impl_->w_sectPr_.remove_child(footerReference);
    }

    pugi::xml_node pgNumType = impl_->w_sectPr_.child("w:pgNumType");
    if (pgNumType) {
      impl_->w_sectPr_.remove_child(pgNumType);
    }
  }

  Paragraph Section::FirstParagraph()
  {
    return Prev().LastParagraph().Next();
  }

  Paragraph Section::LastParagraph()
  {
    if (!impl_) return Paragraph();
    Paragraph::Impl* impl = new Paragraph::Impl;
    impl->w_body_ = impl_->w_body_;
    impl->w_p_ = impl_->w_p_last_;
    impl->w_pPr_ = impl_->w_pPr_last_;
    return Paragraph(impl);
  }

  Section Section::Next()
  {
    if (!impl_) return Section();
    pugi::xml_node w_p = impl_->w_p_last_.next_sibling();
    if (w_p.empty()) return Section();

    Section::Impl* impl = new Section::Impl;
    impl->w_body_ = impl_->w_body_;
    impl->w_p_ = w_p;
    impl->w_pPr_ = w_p.child("w:pPr");
    Section s(impl);
    s.FindSectionProperties();
    return s;
  }

  Section Section::Prev()
  {
    if (!impl_) return Section();

    pugi::xml_node w_p_prev, w_p, w_pPr, w_sectPr;

    w_p_prev = impl_->w_p_.previous_sibling();
    if (w_p_prev.empty()) return Section();

    do {
      w_p = w_p_prev;
      w_pPr = w_p.child("w:pPr");
      w_sectPr = w_pPr.child("w:sectPr");
      w_p_prev = w_p.previous_sibling();
    } while (w_sectPr.empty() && !w_p_prev.empty());

    Section::Impl* impl = new Section::Impl;
    impl->w_body_ = impl_->w_body_;
    impl->w_p_ = impl->w_p_last_ = w_p;
    impl->w_pPr_ = impl->w_pPr_last_ = w_pPr;
    impl->w_sectPr_ = w_sectPr;
    return Section(impl);
  }

  Section::operator bool()
  {
    return impl_ != NULL && impl_->w_sectPr_;
  }

  bool Section::operator==(const Section& s)
  {
    if (!impl_ && !s.impl_) return true;
    if (impl_ && s.impl_) return impl_->w_sectPr_ == s.impl_->w_sectPr_;
    return false;
  }

  void Section::operator=(const Section& right)
  {
    if (this == &right) return;
    if (impl_ != NULL) delete impl_;
    if (right.impl_ != NULL) {
      impl_ = new Impl;
      impl_->w_body_ = right.impl_->w_body_;
      impl_->w_p_ = right.impl_->w_p_;
      impl_->w_p_last_ = right.impl_->w_p_last_;
      impl_->w_pPr_ = right.impl_->w_pPr_;
      impl_->w_pPr_last_ = right.impl_->w_pPr_last_;
      impl_->w_sectPr_ = right.impl_->w_sectPr_;
    }
    else {
      impl_ = NULL;
    }
  }


  // class Run
  Run::Run() : impl_(NULL)
  {

  }

  Run::Run(Impl* impl) : impl_(impl)
  {

  }

  Run::Run(const Run& r)
  {
    impl_ = new Impl;
    impl_->w_p_ = r.impl_->w_p_;
    impl_->w_r_ = r.impl_->w_r_;
    impl_->w_rPr_ = r.impl_->w_rPr_;
  }

  Run::~Run()
  {
    if (impl_ != NULL) {
      delete impl_;
      impl_ = NULL;
    }
  }

  void Run::AppendText(const std::string& text)
  {
    if (!impl_) return;
    pugi::xml_node t = impl_->w_r_.append_child("w:t");
    if (std::isspace(static_cast<unsigned char>(text[0]))) {
      t.append_attribute("xml:space") = "preserve";
    }
    t.text().set(text.c_str());
  }

  std::string Run::GetText()
  {
    if (!impl_) return "";
    std::string text;
    for (pugi::xml_node t = impl_->w_r_.child("w:t"); t; t = t.next_sibling("w:t")) {
      text += t.text().get();
    }
    return text;
  }

  void Run::ClearText()
  {
    if (!impl_) return;
    impl_->w_r_.remove_children();
  }

  void Run::AppendLineBreak()
  {
    if (!impl_) return;
    impl_->w_r_.append_child("w:br");
  }

  void Run::AppendTabs(const unsigned int count)
  {
    if (!impl_) return;
    for (unsigned int i = 0; i < count; i++)
      impl_->w_r_.append_child("w:tab");
  }

  void Run::SetFontSize(const double fontSize)
  {
    if (!impl_) return;
    pugi::xml_node sz = impl_->w_rPr_.child("w:sz");
    if (!sz) {
      sz = impl_->w_rPr_.append_child("w:sz");
    }
    pugi::xml_attribute szVal = sz.attribute("w:val");
    if (!szVal) {
      szVal = sz.append_attribute("w:val");
    }
    // font size in half-points (1/144 of an inch)
    szVal.set_value(fontSize * 2);
  }

  double Run::GetFontSize()
  {
    if (!impl_) return -1;
    pugi::xml_node sz = impl_->w_rPr_.child("w:sz");
    if (!sz) return 0;
    pugi::xml_attribute szVal = sz.attribute("w:val");
    if (!szVal) return 0;
    return szVal.as_int() / 2.0;
  }

  void Run::SetFont(
    const std::string& fontAscii,
    const std::string& fontEastAsia)
  {
    if (!impl_) return;
    pugi::xml_node rFonts = impl_->w_rPr_.child("w:rFonts");
    if (!rFonts) {
      rFonts = impl_->w_rPr_.append_child("w:rFonts");
    }
    pugi::xml_attribute rFontsAscii = rFonts.attribute("w:ascii");
    if (!rFontsAscii) {
      rFontsAscii = rFonts.append_attribute("w:ascii");
    }
    pugi::xml_attribute rFontsEastAsia = rFonts.attribute("w:eastAsia");
    if (!rFontsEastAsia) {
      rFontsEastAsia = rFonts.append_attribute("w:eastAsia");
    }
    rFontsAscii.set_value(fontAscii.c_str());
    rFontsEastAsia.set_value(fontEastAsia.empty()
      ? fontAscii.c_str()
      : fontEastAsia.c_str());
  }

  void Run::GetFont(
    std::string& fontAscii,
    std::string& fontEastAsia)
  {
    if (!impl_) return;
    pugi::xml_node rFonts = impl_->w_rPr_.child("w:rFonts");
    if (!rFonts) return;

    pugi::xml_attribute rFontsAscii = rFonts.attribute("w:ascii");
    if (rFontsAscii) fontAscii = rFontsAscii.value();

    pugi::xml_attribute rFontsEastAsia = rFonts.attribute("w:eastAsia");
    if (rFontsEastAsia) fontEastAsia = rFontsEastAsia.value();
  }

  void Run::SetFontStyle(const FontStyle f)
  {
    if (!impl_) return;
    pugi::xml_node b = impl_->w_rPr_.child("w:b");
    if (f & Bold) {
      if (b.empty()) impl_->w_rPr_.append_child("w:b");
    }
    else {
      impl_->w_rPr_.remove_child(b);
    }

    pugi::xml_node i = impl_->w_rPr_.child("w:i");
    if (f & Italic) {
      if (i.empty()) impl_->w_rPr_.append_child("w:i");
    }
    else {
      impl_->w_rPr_.remove_child(i);
    }

    pugi::xml_node u = impl_->w_rPr_.child("w:u");
    if (f & Underline) {
      if (u.empty())
        impl_->w_rPr_.append_child("w:u").append_attribute("w:val") = "single";
    }
    else {
      impl_->w_rPr_.remove_child(u);
    }

    pugi::xml_node strike = impl_->w_rPr_.child("w:strike");
    if (f & Strikethrough) {
      if (strike.empty())
        impl_->w_rPr_.append_child("w:strike").append_attribute("w:val") = "true";
    }
    else {
      impl_->w_rPr_.remove_child(strike);
    }
  }

  Run::FontStyle Run::GetFontStyle()
  {
    FontStyle fontStyle = 0;
    if (!impl_) return fontStyle;
    if (impl_->w_rPr_.child("w:b")) fontStyle |= Bold;
    if (impl_->w_rPr_.child("w:i")) fontStyle |= Italic;
    if (impl_->w_rPr_.child("w:u")) fontStyle |= Underline;
    if (impl_->w_rPr_.child("w:strike")) fontStyle |= Strikethrough;
    return fontStyle;
  }

  void Run::SetFontColor(const std::string &color)
  {
    if (!impl_) return;
    auto n = impl_->w_rPr_.child("w:color");
    if (n) {
        n.attribute("w:val") = color.c_str();
    }
    else impl_->w_rPr_.append_child("w:color").append_attribute("w:val") = color.c_str();
  }

  void Run::SetCharacterSpacing(const int characterSpacing)
  {
    if (!impl_) return;
    pugi::xml_node spacing = impl_->w_rPr_.child("w:spacing");
    if (!spacing) {
      spacing = impl_->w_rPr_.append_child("w:spacing");
    }
    pugi::xml_attribute spacingVal = spacing.attribute("w:val");
    if (!spacingVal) {
      spacingVal = spacing.append_attribute("w:val");
    }
    spacingVal.set_value(characterSpacing);
  }

  int Run::GetCharacterSpacing()
  {
    if (!impl_) return -1;
    return impl_->w_rPr_.child("w:spacing").attribute("w:val").as_int();
  }

  bool Run::IsPageBreak()
  {
    if (!impl_) return false;
    return impl_->w_r_.find_child_by_attribute("w:br", "w:type", "page");
  }

  void Run::Remove()
  {
    if (!impl_) return;
    impl_->w_p_.remove_child(impl_->w_r_);
  }

  Run Run::Next()
  {
    if (!impl_) return Run();
    pugi::xml_node w_r = impl_->w_r_.next_sibling("w:r");
    if (!w_r) return Run();

    Run::Impl* impl = new Run::Impl;
    impl->w_p_ = impl_->w_p_;
    impl->w_r_ = w_r;
    impl->w_rPr_ = w_r.child("w:rPr");
    return Run(impl);
  }

  Run::operator bool()
  {
    return impl_ != NULL && impl_->w_r_;
  }

  void Run::operator=(const Run& right)
  {
    if (this == &right) return;
    if (impl_ != NULL) delete impl_;
    if (right.impl_ != NULL) {
      impl_ = new Impl;
      impl_->w_p_ = right.impl_->w_p_;
      impl_->w_r_ = right.impl_->w_r_;
      impl_->w_rPr_ = right.impl_->w_rPr_;
    }
    else {
      impl_ = NULL;
    }
  }

  // class Table
  Table::Table(Impl* impl) : impl_(impl)
  {

  }

  Table::Table() : impl_(NULL)
  {

  }

  Table::Table(const Table& t)
  {
    impl_ = new Impl;
    impl_->w_body_ = t.impl_->w_body_;
    impl_->w_tbl_ = t.impl_->w_tbl_;
    impl_->w_tblPr_ = t.impl_->w_tblPr_;
    impl_->w_tblGrid_ = t.impl_->w_tblGrid_;
    impl_->rows_ = t.impl_->rows_;
    impl_->cols_ = t.impl_->cols_;
    impl_->grid_ = t.impl_->grid_;
  }

  Table::~Table()
  {
    if (impl_ != NULL) {
      delete impl_;
      impl_ = NULL;
    }
  }

  void Table::operator=(const Table& right)
  {
    if (this == &right) return;
    if (impl_ != NULL) delete impl_;
    if (right.impl_ != NULL) {
      impl_ = new Impl;
      impl_->w_body_ = right.impl_->w_body_;
      impl_->w_tbl_ = right.impl_->w_tbl_;
      impl_->w_tblPr_ = right.impl_->w_tblPr_;
      impl_->w_tblGrid_ = right.impl_->w_tblGrid_;
      impl_->rows_ = right.impl_->rows_;
      impl_->cols_ = right.impl_->cols_;
      impl_->grid_ = right.impl_->grid_;
    }
    else {
      impl_ = NULL;
    }
  }

  void Table::Create_(const int rows, const int cols)
  {
    if (!impl_) return;
    impl_->rows_ = rows;
    impl_->cols_ = cols;

    // init grid
    for (int i = 0; i < rows; i++) {
      Row row;
      for (int j = 0; j < cols; j++) {
        Cell cell = { i, j, 1, 1 };
        row.push_back(cell);
      }
      impl_->grid_.push_back(row);
    }

    // init table
    for (int i = 0; i < rows; i++) {
      pugi::xml_node w_gridCol = impl_->w_tblGrid_.append_child("w:gridCol");
      pugi::xml_node w_tr = impl_->w_tbl_.append_child("w:tr");

      for (int j = 0; j < cols; j++) {
        pugi::xml_node w_tc = w_tr.append_child("w:tc");
        pugi::xml_node w_tcPr = w_tc.append_child("w:tcPr");

        TableCell::Impl* impl = new TableCell::Impl;
        impl->c_ = &impl_->grid_[i][j];
        impl->w_tr_ = w_tr;
        impl->w_tc_ = w_tc;
        impl->w_tcPr_ = w_tcPr;
        TableCell tc(impl);
        // A table cell must contain at least one block-level element,
        // even if it is an empty <p/>.
        tc.AppendParagraph();
      }
    }
  }

  TableCell Table::GetCell(const int row, const int col)
  {
    if (!impl_) return TableCell();
    if (row < 0 || row >= impl_->rows_ || col < 0 || col >= impl_->cols_) {
      return TableCell();
    }

    Cell* c = &impl_->grid_[row][col];
    return GetCell_(c->row, c->col);
  }

  TableCell Table::GetCell_(const int row, const int col)
  {
    if (!impl_) return TableCell();
    int i = 0;
    pugi::xml_node w_tr = impl_->w_tbl_.child("w:tr");
    while (i < row && !w_tr.empty()) {
      i++;
      w_tr = w_tr.next_sibling("w:tr");
    }
    if (w_tr.empty()) {
      return TableCell();
    }

    int j = 0;
    pugi::xml_node w_tc = w_tr.child("w:tc");
    while (j < col && !w_tc.empty()) {
      j += impl_->grid_[row][j].cols;
      w_tc = w_tc.next_sibling("w:tc");
    }
    if (w_tc.empty()) {
      return TableCell();
    }

    TableCell::Impl* impl = new TableCell::Impl;
    impl->c_ = &impl_->grid_[row][col];
    impl->w_tr_ = w_tr;
    impl->w_tc_ = w_tc;
    impl->w_tcPr_ = w_tc.child("w:tcPr");
    return TableCell(impl);
  }

  bool Table::MergeCells(TableCell tc1, TableCell tc2)
  {
    if (tc1.empty() || tc2.empty()) {
      return false;
    }

    Cell* c1 = tc1.impl_->c_;
    Cell* c2 = tc2.impl_->c_;

    if (c1->row == c2->row && c1->col != c2->col && c1->rows == c2->rows) {
      Cell* left_cell, * right_cell;
      TableCell* left_tc, * right_tc;
      if (c1->col < c2->col) {
        left_cell = c1;
        left_tc = &tc1;
        right_cell = c2;
        right_tc = &tc2;
      }
      else {
        left_cell = c2;
        left_tc = &tc2;
        right_cell = c1;
        right_tc = &tc1;
      }

      const int col = left_cell->col;
      if ((right_cell->col - col) == left_cell->cols) {
        const int cols = left_cell->cols + right_cell->cols;

        // update right grid
        const int right_col = right_cell->col;
        const int right_cols = right_cell->cols;
        for (int i = 0; i < right_cell->rows; i++) {
          const int y = right_cell->row + i;
          for (int j = 0; j < right_cols; j++) {
            Cell& c = impl_->grid_[y][right_col + j];
            c.col = col;
            c.cols = cols;
          }
        }

        // update cells
        for (int i = 0; i < right_cell->rows; i++) {
          RemoveCell_(GetCell_(right_cell->row + i, right_cell->col));
        }
        for (int i = 0; i < left_cell->rows; i++) {
          GetCell_(left_cell->row + i, left_cell->col).SetCellSpanning_(cols);
        }

        // update left grid
        const int left_cols = left_cell->cols;
        for (int i = 0; i < left_cell->rows; i++) {
          const int y = left_cell->row + i;
          for (int j = 0; j < left_cols; j++) {
            Cell& c = impl_->grid_[y][left_cell->col + j];
            c.cols = cols;
          }
        }

        right_tc->impl_->c_ = left_cell;
        right_tc->impl_->w_tc_ = left_tc->impl_->w_tc_;
        right_tc->impl_->w_tcPr_ = left_tc->impl_->w_tcPr_;
        return true;
      }
    }
    else if (c1->col == c2->col && c1->row != c2->row && c1->cols == c2->cols) {
      Cell* top_cell, * bottom_cell;
      TableCell* top_tc, * bottom_tc;
      if (c1->row < c2->row) {
        top_cell = c1;
        top_tc = &tc1;
        bottom_cell = c2;
        bottom_tc = &tc2;
      }
      else {
        top_cell = c2;
        top_tc = &tc2;
        bottom_cell = c1;
        bottom_tc = &tc1;
      }

      const int row = top_cell->row;
      if ((bottom_cell->row - top_cell->row) == top_cell->rows) {
        const int rows = top_cell->rows + bottom_cell->rows;

        // update cells
        if (top_cell->rows == 1) {
          pugi::xml_node w_vMerge = top_tc->impl_->w_tcPr_.append_child("w:vMerge");
          pugi::xml_attribute w_val = w_vMerge.append_attribute("w:val");
          w_val.set_value("restart");
        }
        if (bottom_cell->rows == 1) {
          bottom_tc->impl_->w_tcPr_.append_child("w:vMerge");
        }
        else {
          bottom_tc->impl_->w_tcPr_.remove_child("w:vMerge");
          bottom_tc->impl_->w_tcPr_.append_child("w:vMerge");
        }

        // update top grid
        const int top_rows = top_cell->rows;
        for (int i = 0; i < top_rows; i++) {
          const int x = top_cell->row + i;
          for (int j = 0; j < top_cell->cols; j++) {
            Cell& c = impl_->grid_[x][top_cell->col + j];
            c.rows = rows;
          }
        }

        // update bottom grid
        const int bottom_row = bottom_cell->row;
        const int bottom_rows = bottom_cell->rows;
        for (int i = 0; i < bottom_rows; i++) {
          const int x = bottom_row + i;
          for (int j = 0; j < bottom_cell->cols; j++) {
            Cell& c = impl_->grid_[x][bottom_cell->col + j];
            c.row = row;
            c.rows = rows;
          }
        }

        bottom_tc->impl_->c_ = top_cell;
        bottom_tc->impl_->w_tc_ = top_tc->impl_->w_tc_;
        bottom_tc->impl_->w_tcPr_ = top_tc->impl_->w_tcPr_;
        return true;
      }
    }

    return false;
  }

  void Table::RemoveCell_(TableCell tc)
  {
    if (!impl_) return;
    tc.impl_->w_tr_.remove_child(tc.impl_->w_tc_);
  }

  void Table::SetWidthAuto()
  {
    SetWidth(0, "auto");
  }

  void Table::SetWidthPercent(const double w)
  {
    SetWidth(w / 0.02, "pct");
  }

  void Table::SetWidth(const int w, const char* units)
  {
    if (!impl_) return;
    pugi::xml_node w_tblW = impl_->w_tblPr_.child("w:tblW");
    if (!w_tblW) {
      w_tblW = impl_->w_tblPr_.append_child("w:tblW");
    }

    pugi::xml_attribute w_w = w_tblW.attribute("w:w");
    if (!w_w) {
      w_w = w_tblW.append_attribute("w:w");
    }

    pugi::xml_attribute w_type = w_tblW.attribute("w:type");
    if (!w_type) {
      w_type = w_tblW.append_attribute("w:type");
    }

    w_w.set_value(w);
    w_type.set_value(units);
  }

  void Table::SetCellMarginTop(const int w, const char* units)
  {
    SetCellMargin("w:top", w, units);
  }

  void Table::SetCellMarginBottom(const int w, const char* units)
  {
    SetCellMargin("w:bottom", w, units);
  }

  void Table::SetCellMarginLeft(const int w, const char* units)
  {
    SetCellMargin("w:start", w, units);
  }

  void Table::SetCellMarginRight(const int w, const char* units)
  {
    SetCellMargin("w:end", w, units);
  }

  void Table::SetCellMargin(const char* elemName, const int w, const char* units)
  {
    if (!impl_) return;
    pugi::xml_node w_tblCellMar = impl_->w_tblPr_.child("w:tblCellMar");
    if (!w_tblCellMar) {
      w_tblCellMar = impl_->w_tblPr_.append_child("w:tblCellMar");
    }

    pugi::xml_node w_tblCellMarChild = w_tblCellMar.child(elemName);
    if (!w_tblCellMarChild) {
      w_tblCellMarChild = w_tblCellMar.append_child(elemName);
    }

    pugi::xml_attribute w_w = w_tblCellMarChild.attribute("w:w");
    if (!w_w) {
      w_w = w_tblCellMarChild.append_attribute("w:w");
    }

    pugi::xml_attribute w_type = w_tblCellMarChild.attribute("w:type");
    if (!w_type) {
      w_type = w_tblCellMarChild.append_attribute("w:type");
    }

    w_w.set_value(w);
    w_type.set_value(units);
  }

  void Table::SetAlignment(const Alignment alignment)
  {
    if (!impl_) return;
    const char* val;
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

    pugi::xml_node w_jc = impl_->w_tblPr_.child("w:jc");
    if (!w_jc) {
      w_jc = impl_->w_tblPr_.append_child("w:jc");
    }
    pugi::xml_attribute w_val = w_jc.attribute("w:val");
    if (!w_val) {
      w_val = w_jc.append_attribute("w:val");
    }
    w_val.set_value(val);
  }

  void Table::SetTopBorders(const BorderStyle style, const double width, const char* color)
  {
    SetBorders_("w:top", style, width, color);
  }

  void Table::SetBottomBorders(const BorderStyle style, const double width, const char* color)
  {
    SetBorders_("w:bottom", style, width, color);
  }

  void Table::SetLeftBorders(const BorderStyle style, const double width, const char* color)
  {
    SetBorders_("w:start", style, width, color);
  }

  void Table::SetRightBorders(const BorderStyle style, const double width, const char* color)
  {
    SetBorders_("w:end", style, width, color);
  }

  void Table::SetInsideHBorders(const BorderStyle style, const double width, const char* color)
  {
    SetBorders_("w:insideH", style, width, color);
  }

  void Table::SetInsideVBorders(const BorderStyle style, const double width, const char* color)
  {
    SetBorders_("w:insideV", style, width, color);
  }

  void Table::SetInsideBorders(const BorderStyle style, const double width, const char* color)
  {
    SetInsideHBorders(style, width, color);
    SetInsideVBorders(style, width, color);
  }

  void Table::SetOutsideBorders(const BorderStyle style, const double width, const char* color)
  {
    SetTopBorders(style, width, color);
    SetBottomBorders(style, width, color);
    SetLeftBorders(style, width, color);
    SetRightBorders(style, width, color);
  }

  void Table::SetAllBorders(const BorderStyle style, const double width, const char* color)
  {
    SetOutsideBorders(style, width, color);
    SetInsideBorders(style, width, color);
  }

  void Table::SetBorders_(const char* elemName, const BorderStyle style, const double width, const char* color)
  {
    if (!impl_) return;
    pugi::xml_node w_tblBorders = impl_->w_tblPr_.child("w:tblBorders");
    if (!w_tblBorders) {
      w_tblBorders = impl_->w_tblPr_.append_child("w:tblBorders");
    }
    SetBorders(w_tblBorders, elemName, style, width, color);
  }

  // class TableCell
  TableCell::TableCell() : impl_(NULL)
  {

  }

  TableCell::TableCell(Impl* impl) : impl_(impl)
  {

  }

  TableCell::TableCell(const TableCell& tc)
  {
    impl_ = new Impl;
    impl_->c_ = tc.impl_->c_;
    impl_->w_tr_ = tc.impl_->w_tr_;
    impl_->w_tc_ = tc.impl_->w_tc_;
    impl_->w_tcPr_ = tc.impl_->w_tcPr_;
  }

  TableCell::~TableCell()
  {
    if (impl_ != NULL) {
      delete impl_;
      impl_ = NULL;
    }
  }

  void TableCell::operator=(const TableCell& right)
  {
    if (this == &right) return;
    if (impl_ != NULL) delete impl_;
    if (right.impl_ != NULL) {
      impl_ = new Impl;
      impl_->c_ = right.impl_->c_;
      impl_->w_tr_ = right.impl_->w_tr_;
      impl_->w_tc_ = right.impl_->w_tc_;
      impl_->w_tcPr_ = right.impl_->w_tcPr_;
    }
    else {
      impl_ = NULL;
    }
  }

  TableCell::operator bool()
  {
    return impl_ != NULL && impl_->w_tc_;
  }

  bool TableCell::empty() const
  {
    return impl_ == NULL || !impl_->w_tc_;
  }

  void TableCell::SetWidth(const int w, const char* units)
  {
    if (!impl_) return;
    pugi::xml_node w_tcW = impl_->w_tcPr_.child("w:tcW");
    if (!w_tcW) {
      w_tcW = impl_->w_tcPr_.append_child("w:tcW");
    }

    pugi::xml_attribute w_w = w_tcW.attribute("w:w");
    if (!w_w) {
      w_w = w_tcW.append_attribute("w:w");
    }

    pugi::xml_attribute w_type = w_tcW.attribute("w:type");
    if (!w_type) {
      w_type = w_tcW.append_attribute("w:type");
    }

    w_w.set_value(w);
    w_type.set_value(units);
  }

  void TableCell::SetVerticalAlignment(const Alignment align)
  {
    if (!impl_) return;
    pugi::xml_node w_vAlign = impl_->w_tcPr_.child("w:vAlign");
    if (!w_vAlign) {
      w_vAlign = impl_->w_tcPr_.append_child("w:vAlign");
    }

    pugi::xml_attribute w_val = w_vAlign.attribute("w:val");
    if (!w_val) {
      w_val = w_vAlign.append_attribute("w:val");
    }

    const char* val;
    switch (align) {
    case Alignment::Top:
      val = "top";
      break;
    case Alignment::Center:
      val = "center";
      break;
    case Alignment::Bottom:
      val = "bottom";
      break;
    }
    w_val.set_value(val);
  }

  void TableCell::SetCellSpanning_(const int cols)
  {
    if (!impl_) return;
    pugi::xml_node w_gridSpan = impl_->w_tcPr_.child("w:gridSpan");
    if (cols == 1) {
      if (w_gridSpan) {
        impl_->w_tcPr_.remove_child(w_gridSpan);
      }
      return;
    }
    if (!w_gridSpan) {
      w_gridSpan = impl_->w_tcPr_.append_child("w:gridSpan");
    }

    pugi::xml_attribute w_val = w_gridSpan.attribute("w:val");
    if (!w_val) {
      w_val = w_gridSpan.append_attribute("w:val");
    }

    w_val.set_value(cols);
  }

  Paragraph TableCell::AppendParagraph()
  {
    if (!impl_) return Paragraph();
    pugi::xml_node w_p = impl_->w_tc_.append_child("w:p");
    pugi::xml_node w_pPr = w_p.append_child("w:pPr");

    Paragraph::Impl* impl = new Paragraph::Impl;
    impl->w_body_ = impl_->w_tc_;
    impl->w_p_ = w_p;
    impl->w_pPr_ = w_pPr;
    return Paragraph(impl);
  }

  Paragraph TableCell::FirstParagraph()
  {
    if (!impl_) return Paragraph();
    pugi::xml_node w_p = impl_->w_tc_.child("w:p");
    pugi::xml_node w_pPr = w_p.child("w:pPr");

    Paragraph::Impl* impl = new Paragraph::Impl;
    impl->w_body_ = impl_->w_tc_;
    impl->w_p_ = w_p;
    impl->w_pPr_ = w_pPr;
    return Paragraph(impl);
  }

  // class TextFrame
  TextFrame::TextFrame() : impl_(NULL)
  {

  }

  TextFrame::TextFrame(Impl* impl, Paragraph::Impl* paragraph_impl)
    : Paragraph(paragraph_impl), impl_(impl)
  {

  }

  TextFrame::TextFrame(const TextFrame& tf)
    : Paragraph(tf)
  {
    impl_ = new Impl;
    impl_->w_framePr_ = tf.impl_->w_framePr_;
  }

  TextFrame::~TextFrame()
  {
    if (impl_ != NULL) {
      delete impl_;
      impl_ = NULL;
    }
  }

  void TextFrame::SetSize(const int w, const int h)
  {
    if (!impl_) return;
    pugi::xml_attribute w_w = impl_->w_framePr_.attribute("w:w");
    if (!w_w) {
      w_w = impl_->w_framePr_.append_attribute("w:w");
    }
    pugi::xml_attribute w_h = impl_->w_framePr_.attribute("w:h");
    if (!w_h) {
      w_h = impl_->w_framePr_.append_attribute("w:h");
    }

    w_w.set_value(w);
    w_h.set_value(h);
  }

  void TextFrame::SetPositionX(const Position align, const Anchor ralativeTo)
  {
    SetAnchor_("w:hAnchor", ralativeTo);
    SetPosition_("w:xAlign", align);
  }

  void TextFrame::SetPositionY(const Position align, const Anchor ralativeTo)
  {
    SetAnchor_("w:vAnchor", ralativeTo);
    SetPosition_("w:yAlign", align);
  }

  void TextFrame::SetPositionX(const int x, const Anchor ralativeTo)
  {
    SetAnchor_("w:hAnchor", ralativeTo);
    SetPosition_("w:x", x);
  }

  void TextFrame::SetPositionY(const int y, const Anchor ralativeTo)
  {
    SetAnchor_("w:vAnchor", ralativeTo);
    SetPosition_("w:y", y);
  }

  void TextFrame::SetAnchor_(const char* attrName, const Anchor anchor)
  {
    if (!impl_) return;
    pugi::xml_attribute w_anchor = impl_->w_framePr_.attribute(attrName);
    if (!w_anchor) {
      w_anchor = impl_->w_framePr_.append_attribute(attrName);
    }

    const char* val;
    switch (anchor) {
    case Anchor::Page:
      val = "page";
      break;
    case Anchor::Margin:
      val = "margin";
      break;
    }
    w_anchor.set_value(val);
  }

  void TextFrame::SetPosition_(const char* attrName, const Position align)
  {
    if (!impl_) return;
    pugi::xml_attribute w_align = impl_->w_framePr_.attribute(attrName);
    if (!w_align) {
      w_align = impl_->w_framePr_.append_attribute(attrName);
    }

    const char* val;
    switch (align) {
    case Position::Left:
      val = "left";
      break;
    case Position::Center:
      val = "center";
      break;
    case Position::Right:
      val = "right";
      break;
    case Position::Top:
      val = "top";
      break;
    case Position::Bottom:
      val = "bottom";
      break;
    }
    w_align.set_value(val);
  }

  void TextFrame::SetPosition_(const char* attrName, const int twip)
  {
    if (!impl_) return;
    pugi::xml_attribute w_Align = impl_->w_framePr_.attribute(attrName);
    if (!w_Align) {
      w_Align = impl_->w_framePr_.append_attribute(attrName);
    }
    w_Align.set_value(twip);
  }

  void TextFrame::SetTextWrapping(const Wrapping wrapping)
  {
    if (!impl_) return;
    pugi::xml_attribute w_wrap = impl_->w_framePr_.attribute("w:wrap");
    if (!w_wrap) {
      w_wrap = impl_->w_framePr_.append_attribute("w:wrap");
    }

    const char* val;
    switch (wrapping) {
    case Wrapping::Around:
      val = "around";
      break;
    case Wrapping::None:
      val = "none";
      break;
    }
    w_wrap.set_value(val);
  }


} // namespace docx
