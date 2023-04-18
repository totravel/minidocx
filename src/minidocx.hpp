/**
 * minidocx 0.5.0 - C++ library for creating Microsoft Word Document (.docx).
 * --------------------------------------------------------
 * Copyright (C) 2022-2023, by Xie Zequn (totravel@foxmail.com)
 * Report bugs and download new versions at https://github.com/totravel/minidocx
 */

#include <iostream> // std::ostream
#include <string>
#include <vector>


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
  class Table;
  class TableCell;
  class TextFrame;


  class Box
  {
  public:
    enum class BorderStyle { Single, Dotted, Dashed, DotDash, Double, Wave, None };
  };


  struct Cell {
    int row, col; // position
    int rows, cols; // size
  };
  using Row = std::vector<Cell>;
  using Grid = std::vector<Row>;


  class TableCell
  {
    friend class Table;

  public:
    // constructs an empty cell
    TableCell();
    TableCell(const TableCell& tc);
    ~TableCell();
    void operator=(const TableCell& right);

    operator bool();
    bool empty() const;

    void SetWidth(const int w, const char* units = "dxa");

    enum class Alignment { Top, Center, Bottom };
    void SetVerticalAlignment(const Alignment align);

    void SetCellSpanning_(const int cols);

    Paragraph AppendParagraph();
    Paragraph FirstParagraph();

  private:
    struct Impl;
    Impl* impl_ = nullptr;

    // constructs a table from existing xml node
    TableCell(Impl* impl);
  }; // class TableCell


  class Table : public Box
  {
    friend class Document;

  public:
    // constructs an empty table
    Table();
    Table(const Table& t);
    ~Table();
    void operator=(const Table& right);

    void Create_(const int rows, const int cols);

    TableCell GetCell(const int row, const int col);
    TableCell GetCell_(const int row, const int col);
    bool MergeCells(TableCell tc1, TableCell tc2);
    bool SplitCell();

    void RemoveCell_(TableCell tc);

    // units: 
    //   auto - Specifies that width is determined by the overall table layout algorithm.
    //   dxa  - Specifies that the value is in twentieths of a point (1/1440 of an inch).
    //   nil  - Specifies a value of zero.
    //   pct  - Specifies a value as a percent of the table width.
    void SetWidthAuto();
    void SetWidthPercent(const double w); // 0-100
    void SetWidth(const int w, const char* units = "dxa");

    // the distance between the cell contents and the cell borders
    void SetCellMarginTop(const int w, const char* units = "dxa");
    void SetCellMarginBottom(const int w, const char* units = "dxa");
    void SetCellMarginLeft(const int w, const char* units = "dxa");
    void SetCellMarginRight(const int w, const char* units = "dxa");
    void SetCellMargin(const char* elemName, const int w, const char* units = "dxa");

    // table formatting
    enum class Alignment { Left, Centered, Right };
    void SetAlignment(const Alignment alignment);

    // style - Specifies the style of the border.
    // width - Specifies the width of the border in points.
    // color - Specifies the color of the border. 
    //         Values are given as hex values (in RRGGBB format). 
    //         No #, unlike hex values in HTML/CSS. E.g., color="FFFF00". 
    //         A value of auto is also permitted and will allow the 
    //         consuming word processor to determine the color.
    void SetTopBorders(const BorderStyle style = BorderStyle::Single, const double width = 0.5, const char* color = "auto");
    void SetBottomBorders(const BorderStyle style = BorderStyle::Single, const double width = 0.5, const char* color = "auto");
    void SetLeftBorders(const BorderStyle style = BorderStyle::Single, const double width = 0.5, const char* color = "auto");
    void SetRightBorders(const BorderStyle style = BorderStyle::Single, const double width = 0.5, const char* color = "auto");
    void SetInsideHBorders(const BorderStyle style = BorderStyle::Single, const double width = 0.5, const char* color = "auto");
    void SetInsideVBorders(const BorderStyle style = BorderStyle::Single, const double width = 0.5, const char* color = "auto");
    void SetInsideBorders(const BorderStyle style = BorderStyle::Single, const double width = 0.5, const char* color = "auto");
    void SetOutsideBorders(const BorderStyle style = BorderStyle::Single, const double width = 0.5, const char* color = "auto");
    void SetAllBorders(const BorderStyle style = BorderStyle::Single, const double width = 0.5, const char* color = "auto");
    void SetBorders_(const char* elemName, const BorderStyle style, const double width, const char* color);

  private:
    struct Impl;
    Impl* impl_ = nullptr;

    // constructs a table from existing xml node
    Table(Impl* impl);
  }; // class Table


  class Run
  {
    friend class Paragraph;
    friend std::ostream& operator<<(std::ostream& out, const Run& r);

  public:
    // constructs an empty run
    Run();
    Run(const Run& r);
    ~Run();
    void operator=(const Run& right);

    operator bool();
    Run Next();
    
    // text
    void AppendText(const std::string& text);
    std::string GetText();
    void ClearText();
    void AppendLineBreak();

    // text formatting
    using FontStyle = unsigned int;
    enum : FontStyle
    {
      Bold = 1,
      Italic = 2,
      Underline = 4,
      Strikethrough = 8
    };
    void SetFontSize(const double fontSize);
    double GetFontSize();

    void SetFont(const std::string& fontAscii, const std::string& fontEastAsia = "");
    void GetFont(std::string& fontAscii, std::string& fontEastAsia);

    void SetFontStyle(const FontStyle fontStyle);
    FontStyle GetFontStyle();

    void SetCharacterSpacing(const int characterSpacing);
    int GetCharacterSpacing();

    // Run
    void Remove();
    bool IsPageBreak();

  private:
    struct Impl;
    Impl* impl_ = nullptr;

    // constructs run from existing xml node
    Run(Impl* impl);
  }; // class Run


  class Section
  {
    friend class Document;
    friend class Paragraph;
    friend std::ostream& operator<<(std::ostream& out, const Section& s);

  public:
    // constructs an empty section
    Section();
    Section(const Section& s);
    ~Section();
    void operator=(const Section& right);
    bool operator==(const Section& s);

    operator bool();
    Section Next();
    Section Prev();

    // section
    void Split();
    bool IsSplit();
    void Merge();

    // page formatting
    enum class Orientation { Landscape, Portrait, Unknown };
    void SetPageSize(const int w, const int h);
    void GetPageSize(int& w, int& h);

    void SetPageOrient(const Orientation orient);
    Orientation GetPageOrient();

    void SetPageMargin(const int top, const int bottom, const int left, const int right);
    void GetPageMargin(int& top, int& bottom, int& left, int& right);

    void SetPageMargin(const int header, const int footer);
    void GetPageMargin(int& header, int& footer);

    void SetColumn(const int num, const int space = 425);
    
    enum class PageNumberFormat {
      Decimal,      // e.g., 1, 2, 3, 4, etc.
      NumberInDash, // e.g., -1-, -2-, -3-, -4-, etc.
      CardinalText, // In English, One, Two, Three, etc.
      OrdinalText,  // In English, First, Second, Third, etc.
      LowerLetter,  // e.g., a, b, c, etc.
      UpperLetter,  // e.g., A, B, C, etc.
      LowerRoman,   // e.g., i, ii, iii, iv, etc.
      UpperRoman    // e.g., I, II, III, IV, etc.
    };

    /**
     * Specifies the page numbers for pages in the section.
     * 
     * @param fmt   Specifies the number format to be used for page numbers in the section.
     * 
     * @param start Specifies the page number that appears on the first page of the section.
     *              If the value is omitted, numbering continues from the highest page number in the previous section.
     */
    void SetPageNumber(const PageNumberFormat fmt = PageNumberFormat::Decimal, const unsigned int start = 0);
    void RemovePageNumber();

    // paragraph
    Paragraph FirstParagraph();
    Paragraph LastParagraph();

  private:
    struct Impl;
    Impl* impl_ = nullptr;

    // constructs section from existing xml node
    Section(Impl* impl);
    void FindSectionProperties();
  }; // class Section


  class Paragraph : public Box
  {
    friend class Document;
    friend class Section;
    friend class TableCell;
    friend std::ostream& operator<<(std::ostream& out, const Paragraph& p);

  public:
    // constructs an empty paragraph
    Paragraph();
    Paragraph(const Paragraph& p);
    ~Paragraph();
    void operator=(const Paragraph& right);
    bool operator==(const Paragraph& p);

    operator bool();
    Paragraph Next();
    Paragraph Prev();

    // get run
    Run FirstRun();

    // add run
    Run AppendRun();
    Run AppendRun(const std::string& text);
    Run AppendRun(const std::string& text, const double fontSize);
    Run AppendRun(const std::string& text, const double fontSize, const std::string& fontAscii, const std::string& fontEastAsia = "");
    Run AppendPageBreak();

    // paragraph formatting
    enum class Alignment { Left, Centered, Right, Justified, Distributed };
    void SetAlignment(const Alignment alignment);

    void SetLineSpacingSingle();               // Single
    void SetLineSpacingLines(const double at); // 1.5 lines, Double (2 lines), Multiple (3 lines)
    void SetLineSpacingAtLeast(const int at);  // At Least
    void SetLineSpacingExactly(const int at);  // Exactly
    void SetLineSpacing(const int at, const char* lineRule);

    void SetBeforeSpacingAuto();
    void SetAfterSpacingAuto();
    void SetSpacingAuto(const char* attrNameAuto);
    void SetBeforeSpacingLines(const double beforeSpacing);
    void SetAfterSpacingLines(const double afterSpacing);
    void SetBeforeSpacing(const int beforeSpacing);
    void SetAfterSpacing(const int afterSpacing);
    void SetSpacing(const int twip, const char* attrNameAuto, const char* attrName);

    void SetLeftIndentChars(const double leftIndent);
    void SetRightIndentChars(const double rightIndent);
    void SetLeftIndent(const int leftIndent);
    void SetRightIndent(const int rightIndent);
    void SetFirstLineChars(const double indent);
    void SetHangingChars(const double indent);
    void SetFirstLine(const int indent);
    void SetHanging(const int indent);
    void SetIndent(const int indent, const char* attrName);

    void SetTopBorder(const BorderStyle style = BorderStyle::Single, const double width = 0.5, const char* color = "auto");
    void SetBottomBorder(const BorderStyle style = BorderStyle::Single, const double width = 0.5, const char* color = "auto");
    void SetLeftBorder(const BorderStyle style = BorderStyle::Single, const double width = 0.5, const char* color = "auto");
    void SetRightBorder(const BorderStyle style = BorderStyle::Single, const double width = 0.5, const char* color = "auto");
    void SetBorders(const BorderStyle style = BorderStyle::Single, const double width = 0.5, const char* color = "auto");
    void SetBorders_(const char* elemName, const BorderStyle style, const double width, const char* color);

    // helper
    void SetFontSize(const double fontSize);
    void SetFont(const std::string& fontAscii, const std::string& fontEastAsia = "");
    void SetFontStyle(const Run::FontStyle fontStyle);
    void SetCharacterSpacing(const int characterSpacing);
    std::string GetText();

    // section
    Section GetSection();
    Section InsertSectionBreak();
    Section RemoveSectionBreak();
    bool HasSectionBreak();

  protected:
    struct Impl;
    Impl* impl_ = nullptr;

    // constructs paragraph from existing xml node
    Paragraph(Impl* impl);
  }; // class Paragraph


  class TextFrame : public Paragraph
  {
    friend class Document;

  public:
    // constructs an empty text frame
    TextFrame();
    TextFrame(const TextFrame& tf);
    ~TextFrame();

    void SetSize(const int w, const int h);

    enum class Anchor { Page, Margin };
    enum class Position { Left, Center, Right, Top, Bottom };
    void SetAnchor_(const char* attrName, const Anchor anchor);
    void SetPosition_(const char* attrName, const Position align);
    void SetPosition_(const char* attrName, const int twip);

    void SetPositionX(const Position align, const Anchor ralativeTo);
    void SetPositionY(const Position align, const Anchor ralativeTo);
    void SetPositionX(const int x, const Anchor ralativeTo);
    void SetPositionY(const int y, const Anchor ralativeTo);

    enum class Wrapping { Around, None };
    void SetTextWrapping(const Wrapping wrapping);

  private:
    struct Impl;
    Impl* impl_ = nullptr;

    // constructs text frame from existing xml node
    TextFrame(Impl* impl, Paragraph::Impl* p_impl);
  }; // class TextFrame


  class Document
  {
    friend std::ostream& operator<<(std::ostream& out, const Document& doc);

  public:
    // constructs an empty document
    Document(const std::string& path);
    ~Document();

    // save document to file
    bool Save();
    bool Open(const std::string& path);

    // get paragraph
    Paragraph FirstParagraph();
    Paragraph LastParagraph();

    // add paragraph
    Paragraph AppendParagraph();
    Paragraph AppendParagraph(const std::string& text);
    Paragraph AppendParagraph(const std::string& text, const double fontSize);
    Paragraph AppendParagraph(const std::string& text, const double fontSize, const std::string& fontAscii, const std::string& fontEastAsia = "");
    Paragraph PrependParagraph();
    Paragraph PrependParagraph(const std::string& text);
    Paragraph PrependParagraph(const std::string& text, const double fontSize);
    Paragraph PrependParagraph(const std::string& text, const double fontSize, const std::string& fontAscii, const std::string& fontEastAsia = "");

    Paragraph InsertParagraphBefore(const Paragraph& p);
    Paragraph InsertParagraphAfter(const Paragraph& p);
    bool RemoveParagraph(Paragraph& p);

    Paragraph AppendPageBreak();

    // get section
    Section FirstSection();
    Section LastSection();

    // add section
    Paragraph AppendSectionBreak();

    // add table
    Table AppendTable(const int rows, const int cols);
    void RemoveTable(Table& tbl);

    // add text frame
    TextFrame AppendTextFrame(const int w, const int h);

  private:
    struct Impl;
    Impl* impl_ = nullptr;
  }; // class Document


} // namespace docx
