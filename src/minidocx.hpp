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
  void GetCharLen(std::string text, int &ascii, int &eastAsia);
  int GetCharLen(std::string s);
  
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

    void SetFontStyle(FontStyle fontStyle);
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

    void SetPageOrient(Orientation orient);
    Orientation GetPageOrient();

    void SetPageMargin(const int top, const int bottom, 
                       const int left, const int right);
    void GetPageMargin(int &top, int &bottom, 
                       int &left, int &right);

    void SetPageMargin(const int header, const int footer);
    void GetPageMargin(int &header, int &footer);

    void SetColumn(int num, int space = 425);


    // paragraph
    Paragraph FirstParagraph();
    Paragraph LastParagraph();

    // traverse
    Section Next();
    Section Prev();
    operator bool();

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
    void SetAlignment(Alignment alignment);

    void SetLineSpacingSingle();         // Single
    void SetLineSpacingLines(double at); // 1.5 lines, Double (2 lines), Multiple (3 lines)
    void SetLineSpacingAtLeast(int at);  // At Least
    void SetLineSpacingExactly(int at);  // Exactly
    void SetLineSpacing(int at, const char *lineRule);

    void SetBeforeSpacingAuto();
    void SetAfterSpacingAuto();
    void SetSpacingAuto(const char *attrNameAuto);
    void SetBeforeSpacingLines(double beforeSpacing);
    void SetAfterSpacingLines(double afterSpacing);
    void SetBeforeSpacing(int beforeSpacing);
    void SetAfterSpacing(int afterSpacing);
    void SetSpacing(int twip, const char *attrNameAuto, const char *attrName);

    void SetLeftIndentChars(double leftIndent);
    void SetRightIndentChars(double rightIndent);
    void SetLeftIndent(int leftIndent);
    void SetRightIndent(int rightIndent);
    void SetFirstLineChars(double indent);
    void SetHangingChars(double indent);
    void SetFirstLine(int indent);
    void SetHanging(int indent);
    void SetIndent(int indent, const char *attrName);

    // helper
    void SetFontSize(const double fontSize);
    void SetFont(const std::string fontAscii, 
                 const std::string fontEastAsia = "");
    void SetFontStyle(Run::FontStyle fontStyle);
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
  }; // class Document


} // namespace docx
