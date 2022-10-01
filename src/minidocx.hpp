/**
 * minidocx - C++ library for creating Microsoft Word Document (.docx).
 * 
 * Copyright (C) 2022 Xie Zequn <totravel@foxmail.com>
 *
 * Units: 
 *   Values are in twentieths of a point, e.g. 1440 = 72 points.
 *   One point is equal to 1/72 inch, e.g. 72 points = 1 inch.
 */

#include <iostream>
#include <string>
#include "pugixml.hpp"

namespace docx
{
  const int PPI = 72;

  // inches
  const double A0_W = 33.1;
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

  const double LETTER_W = 8.5;
  const double LETTER_H = 11;

  const double LEGAL_W = 8.5;
  const double LEGAL_H = 14;

  const double TABLOID_W = 11;
  const double TABLOID_H = 17;

  // pixels
  const unsigned int A0_COLS = 2384;
  const unsigned int A0_ROWS = 3370;

  const unsigned int A1_COLS = 1684;
  const unsigned int A1_ROWS = 2384;

  const unsigned int A2_COLS = 1191;
  const unsigned int A2_ROWS = 1684;

  const unsigned int A3_COLS = 842;
  const unsigned int A3_ROWS = 1190;

  const unsigned int A4_COLS = 595;
  const unsigned int A4_ROWS = 842;

  const unsigned int A5_COLS = 420;
  const unsigned int A5_ROWS = 595;

  const unsigned int A6_COLS = 297;
  const unsigned int A6_ROWS = 420;

  const unsigned int LETTER_COLS = 612;
  const unsigned int LETTER_ROWS = 792;

  const unsigned int LEGAL_COLS = 612;
  const unsigned int LEGAL_ROWS = 1008;

  const unsigned int TABLOID_COLS = 792;
  const unsigned int TABLOID_ROWS = 1224;
  
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
    // constructs run from existing xml node
    Run(pugi::xml_node w_p, 
        pugi::xml_node w_r, 
        pugi::xml_node w_rPr);

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
    pugi::xml_node w_p_;
    pugi::xml_node w_r_;
    pugi::xml_node w_rPr_;
  }; // class Run


  class Section
  {
  public:
    // constructs an empty section
    Section();
    // constructs a new section
    Section(pugi::xml_node w_body, 
            pugi::xml_node w_p, 
            pugi::xml_node w_pPr);
    // constructs section from existing xml node
    Section(pugi::xml_node w_body, 
            pugi::xml_node w_p, 
            pugi::xml_node w_pPr, 
            pugi::xml_node w_sectPr);

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
    pugi::xml_node w_body_;
    pugi::xml_node w_p_;      // current paragraph
    pugi::xml_node w_p_last_; // the last paragraph of the section
    pugi::xml_node w_pPr_;
    pugi::xml_node w_pPr_last_;
    pugi::xml_node w_sectPr_;

    void GetSectPr();
  }; // class Section


  class Paragraph
  {
    friend class Document;
    friend class Section;

  public:
    // constructs an empty paragraph
    Paragraph();
    // constructs a new paragraph
    Paragraph(pugi::xml_node w_body);
    // constructs paragraph from existing xml node
    Paragraph(pugi::xml_node w_body, 
              pugi::xml_node w_p, 
              pugi::xml_node w_pPr);

    // get run
    Run FirstRun();

    // add run
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

    // traverse paragraph
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
    pugi::xml_node w_body_;
    pugi::xml_node w_p_;
    pugi::xml_node w_pPr_;
  }; // class Paragraph


  class Document
  {
  public:
    // constructs an empty document
    Document(const std::string path);
    // save document to file
    bool Save();
    bool Open(const std::string path);

    friend std::ostream& operator<<(std::ostream &out, const Document &doc);

    // get paragraph
    Paragraph FirstParagraph();
    Paragraph LastParagraph();

    // add paragraph
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

    Paragraph InsertParagraphBefore(Paragraph &p);
    Paragraph InsertParagraphAfter(Paragraph &p);
    bool RemoveParagraph(Paragraph &p);

    Paragraph AppendPageBreak();

    // get section
    Section FirstSection();
    Section LastSection();

    // add section
    Paragraph AppendSectionBreak();

  private:
    std::string        path_;
    pugi::xml_document doc_;
    pugi::xml_node     w_body_;
    pugi::xml_node     w_sectPr_;
  }; // class Document


} // namespace docx
