/**
 * minidocx - C++ library for creating Microsoft Word Document (.docx).
 * 
 * Copyright (C) 2022 Xie Zequn <totravel@foxmail.com>
 *
 * Units: 
 *   Values are in twentieths of a point, e.g. 1440 = 72 points.
 *   One point is equal to 1/72 inch, e.g. 72 points = 1 inch.
 */

#include <string>
#include "pugixml.hpp"

namespace docx
{
  const int PPI = 72;

  const double A0_W = 33.1; // inches
  const double A0_H = 46.8;

  const double A1_W = 23.4;
  const double A1_H = 33.1;

  const double A2_W = 16.5;
  const double A2_H = 23.4;

  const double A3_W = 11.7;
  const double A3_H = 16.5;

  const double A4_W = 8.3;
  const double A4_H = 11.7;

  const double A5_W = 5.8;
  const double A5_H = 8.3;

  const double A6_W = 4.1;
  const double A6_H = 5.8;

  const double Letter_W = 8.5;
  const double Letter_H = 11;

  const double Legal_W = 8.5;
  const double Legal_H = 14;

  const double Tabloid_W = 11;
  const double Tabloid_H = 17;

  const double A0_COLS = 2384; // pixels
  const double A0_ROWS = 3370;

  const double A1_COLS = 1684;
  const double A1_ROWS = 2384;

  const double A2_COLS = 1191;
  const double A2_ROWS = 1684;

  const double A3_COLS = 842;
  const double A3_ROWS = 1190;

  const double A4_COLS = 595;
  const double A4_ROWS = 842;

  const double A5_COLS = 420;
  const double A5_ROWS = 595;

  const double A6_COLS = 297;
  const double A6_ROWS = 420;

  const double Letter_COLS = 612;
  const double Letter_ROWS = 792;

  const double Legal_COLS = 612;
  const double Legal_ROWS = 1008;

  const double Tabloid_COLS = 792;
  const double Tabloid_ROWS = 1224;

  void GetCharLen(const std::string text, int &ascii, int &eastAsia);
  int GetCharLen(const std::string s);
  
  int Pt2Twip(const double pt);
  double Twip2Pt(const int twip);

  double Inch2Pt(const double inch);
  double Pt2Inch(const double pt);

  double MM2Inch(const int mm);
  int Inch2MM(const double inch);

  double CM2Inch(const double cm);
  double Inch2CM(const double inch);

  int Inch2Twip(const double inch);
  double Twip2Inch(const int twip);

  int MM2Twip(const int mm);
  int Twip2MM(const int twip);

  int CM2Twip(const double cm);
  double Twip2CM(const int twip);

  class Document;
  class Paragraph;
  class Section;
  class Run;

  class Run
  {
  public:
    Run(pugi::xml_node p, 
        pugi::xml_node r, 
        pugi::xml_node rPr): p_(p), 
                             r_(r), 
                             rPr_(rPr) {}

    // text
    void AppendText(const std::string text);
    std::string GetText();
    void ClearText();
    void AppendLineBreak();

    // text formatting
    using FontStyle = unsigned int;
    enum : FontStyle
    {
      Bold          = 1 << 0, 
      Italic        = 1 << 1, 
      Underline     = 1 << 2, 
      Strikethrough = 1 << 3
    };
    void SetFontSize(const double fontSize);
    double GetFontSize();

    void SetFont(const std::string fontAscii, 
                 const std::string fontEastAsia = "");
    void GetFont(std::string &fontAscii, 
                 std::string &fontEastAsia);

    void SetFontStyle(const FontStyle fontStyle);
    FontStyle GetFontStyle();

    void SetCharacterSpacing(const int characterSpacing);
    int GetCharacterSpacing();

    // Run
    void Remove();
    bool IsPageBreak();

    // traverse
    Run Next();
    operator bool();

  private:
    pugi::xml_node p_;
    pugi::xml_node r_;
    pugi::xml_node rPr_;
  }; // class Run


  class Section
  {
  public:
    Section() {};
    Section(pugi::xml_node body, 
            pugi::xml_node p, 
            pugi::xml_node pPr);
    Section(pugi::xml_node body, 
            pugi::xml_node p, 
            pugi::xml_node pPr, 
            pugi::xml_node sectPr): body_(body), 
                                    p_(p), 
                                    pLast_(p), 
                                    pPr_(pPr), 
                                    pPrLast_(pPr), 
                                    sectPr_(sectPr) {}

    // section
    void Split();
    bool IsSplit();
    void Merge();

    // page formatting
    enum class Orientation { Landscape, Portrait };
    void SetPageSize(const int w, const int h);
    void GetPageSize(int &w, int &h);

    void SetPageOrient(const Orientation orient);
    Orientation GetPageOrient();

    void SetPageMargin(const int top, const int bottom, 
                       const int left, const int right);
    void GetPageMargin(int &top, int &bottom, 
                       int &left, int &right);

    void SetPageMargin(const int header, const int footer);
    void GetPageMargin(int &header, int &footer);

    void SetColumn(const int num, const int space = 425);


    // paragraph
    Paragraph FirstParagraph();
    Paragraph LastParagraph();

    // traverse
    Section Next();
    Section Prev();
    operator bool();
    bool operator==(const Section &s);

  private:
    pugi::xml_node body_;
    pugi::xml_node p_;     // current paragraph
    pugi::xml_node pLast_; // the last paragraph of the section
    pugi::xml_node pPr_;
    pugi::xml_node pPrLast_;
    pugi::xml_node sectPr_;

    void GetSectPr();
  }; // class Section


  class Paragraph
  {
  public:
    Paragraph(pugi::xml_node body): body_(body), 
                                    p_(body.append_child("w:p")), 
                                    pPr_(p_.append_child("w:pPr")) {}
    Paragraph(pugi::xml_node body, 
              pugi::xml_node p, 
              pugi::xml_node pPr): body_(body), 
                                   p_(p), 
                                   pPr_(pPr) {}

    // run
    Run FirstRun();

    Run AppendRun();
    Run AppendRun(const std::string text);
    Run AppendRun(const std::string text, 
                  const double fontSize);
    Run AppendRun(const std::string text, 
                  const double fontSize, 
                  const std::string fontAscii, 
                  const std::string fontEastAsia = "");
    Run AppendPageBreak();

    // paragraph formatting
    enum class Alignment { Left, Centered, Right, Justified, Distributed };
    void SetAlignment(const Alignment alignment);

    void SetLineSpacingSingle();         // Single
    void SetLineSpacingLines(const double at); // 1.5 lines, Double (2 lines), Multiple (3 lines)
    void SetLineSpacingAtLeast(const int at);  // At Least
    void SetLineSpacingExactly(const int at);  // Exactly
    void SetLineSpacing(const int at, const char *lineRule);

    void SetBeforeSpacingAuto();
    void SetAfterSpacingAuto();
    void SetSpacingAuto(const char *attrNameAuto);
    void SetBeforeSpacingLines(const double beforeSpacing);
    void SetAfterSpacingLines(const double afterSpacing);
    void SetBeforeSpacing(const int beforeSpacing);
    void SetAfterSpacing(const int afterSpacing);
    void SetSpacing(const int twip, const char *attrNameAuto, const char *attrName);

    void SetLeftIndentChars(const double leftIndent);
    void SetRightIndentChars(const double rightIndent);
    void SetLeftIndent(const int leftIndent);
    void SetRightIndent(const int rightIndent);
    void SetFirstLineChars(const double indent);
    void SetHangingChars(const double indent);
    void SetFirstLine(const int indent);
    void SetHanging(const int indent);
    void SetIndent(const int indent, const char *attrName);

    // helper
    void SetFontSize(const double fontSize);
    void SetFont(const std::string fontAscii, 
                 const std::string fontEastAsia = "");
    void SetFontStyle(const Run::FontStyle fontStyle);
    void SetCharacterSpacing(const int characterSpacing);
    std::string GetText();

    // paragraph
    Paragraph InsertBefore();
    Paragraph InsertAfter();
    void Remove();

    // traverse
    Paragraph Next();
    Paragraph Prev();
    operator bool();
    bool operator==(const Paragraph &p);

    // section
    Section GetSection();
    Section InsertSectionBreak();
    Section RemoveSectionBreak();
    bool HasSectionBreak();

  private:
    pugi::xml_node body_;
    pugi::xml_node p_;
    pugi::xml_node pPr_;
  }; // class Paragraph


  class Document
  {
  public:
    Document(const std::string path);
    void Save();

    std::string GetFormatedBody();
    
    // paragraph
    Paragraph FirstParagraph();
    Paragraph LastParagraph();

    // section
    Section FirstSection();
    Section LastSection();

    Paragraph AppendParagraph();
    Paragraph AppendParagraph(const std::string text);
    Paragraph AppendParagraph(const std::string text, 
                              const double fontSize);
    Paragraph AppendParagraph(const std::string text, 
                              const double fontSize, 
                              const std::string fontAscii, 
                              const std::string fontEastAsia = "");
    Paragraph PrependParagraph();
    Paragraph PrependParagraph(const std::string text);
    Paragraph PrependParagraph(const std::string text, 
                               const double fontSize);
    Paragraph PrependParagraph(const std::string text, 
                               const double fontSize, 
                               const std::string fontAscii, 
                               const std::string fontEastAsia = "");
    Paragraph AppendPageBreak();
    Paragraph AppendSectionBreak();

  private:
    std::string        path_;
    pugi::xml_document doc_;
    pugi::xml_node     body_;
    pugi::xml_node     sectPr_;
  }; // class Document


} // namespace docx
