
# minidocx

![License](https://img.shields.io/github/license/totravel/minidocx)
![Documentation Status](https://img.shields.io/badge/中文文档-最新-brightgreen.svg)
![GitHub code size in bytes](https://img.shields.io/github/languages/code-size/totravel/minidocx)
![GitHub last commit](https://img.shields.io/github/last-commit/totravel/minidocx)

English | [简体中文](./README-zh_CN.md)

minidocx is a portable and easy-to-use C++ library for creating Microsoft Word Document (.docx) from scratch. It is designed to be simple and small enough. So, you can grab these files and compile them into your project.

## Requirements

To build minidocx you'll need a C++11 compiler and the following libraries:

- [zip](https://github.com/kuba--/zip) <= 0.2.1
- [pugixml](https://github.com/zeux/pugixml) >= 1.13

It is tested for the following platform:

- CMake 3.21
- Visual Studio 16 2019 / Visual Studio 14 2015 / GCC 8.2+
- Windows 7/10 / Ubuntu 18.04 / Kylin-Desktop V10-SP1

## References

To manipulate a docx file, you need to understand at least the following three concepts:

- `Paragraph` Text with the same paragraph formatting.
- `Run` Text with the same font formatting.
- `Section` A division of a document having the same page layout settings, such as margins and page orientation.

Documents contain sections, you can have multiple sections per document. This simple example will only contain one section. More informations are available in [this site](http://officeopenxml.com/).

## Examples

Here's an example of how to use minidocx to create a .docx file.

```cpp
#include "minidocx.hpp"

using namespace docx;

int main()
{
  Document doc("./a.docx");

  auto p1 = doc.AppendParagraph("Hello, World!", 12, "Times New Roman");
  auto p2 = doc.AppendParagraph(u8"你好，世界！", 14, u8"宋体");
  auto p3 = doc.AppendParagraph(u8"你好，World!", 16, "Times New Roman", u8"宋体");

  auto p4 = doc.AppendParagraph();
  p4.SetAlignment(Paragraph::Alignment::Centered);

  auto p4r1 = p4.AppendRun("This is a simple sentence. ", 12, "Arial");
  p4r1.SetCharacterSpacing(Pt2Twip(2));

  auto p4r2 = p4.AppendRun(u8"这是一个简单的句子。");
  p4r2.SetFontSize(14);
  p4r2.SetFont(u8"黑体");
  p4r2.SetFontStyle(Run::Bold | Run::Italic);

  doc.Save();
  return 0;
}
```

See other [examples](./examples).

## Build Instructions

```bash
git clone git@github.com:totravel/minidocx.git
cd minidocx

# Windows
cmake -S . -B build -DBUILD_EXAMPLES=ON -DWITH_STATIC_CRT=OFF
cmake --build build --config Release -j
cmake --install build --prefix install --config Release

# Linux
cmake -S . -B build -DCMAKE_BUILD_TYPE=Release -DBUILD_EXAMPLES=ON -DWITH_STATIC_CRT=OFF
cmake --build build -j
cmake --install build --prefix install
```

## User Guide

`minidocx.hpp` is the only header which you need to include in order to use minidocx classes/functions. All minidocx classes/functions are member of the `docx` namespace.

```cpp
#include "minidocx.hpp"

using namespace docx;
```

Following sections describe the functionality supported by minidocx. Please note this description may not be complete but limited to the most useful ones. If you want to find less common features, please check header files under `src` directory.

### Units

The main unit in OOXML are twentieths of a point (Twip) from the [OOXML](http://officeopenxml.com/) specification. This is used for specifying page dimensions, margins, tabs, etc. minidocx provided some helper functions for unit conversions:

```cpp
int Pt2Twip(const double pt);       // points to twip
double Twip2Pt(const int twip);     // twip to points

int Inch2Twip(const double inch);   // inches to twip
double Twip2Inch(const int twip);   // twip to inches

int MM2Twip(const int mm);          // mm to twip
int Twip2MM(const int twip);        // twip to mm

int CM2Twip(const double cm);       // cm to twip
double Twip2CM(const int twip);     // twip to cm

double Inch2Pt(const double inch);  // inches to points
double Pt2Inch(const double pt);    // points to inches

double MM2Inch(const int mm);       // mm to inches
int Inch2MM(const double inch);     // inches to mm

double CM2Inch(const double cm);    // cm to inches
double Inch2CM(const double inch);  // inches to cm
```

See [Lars Corneliussen's blog post](https://startbigthinksmall.wordpress.com/2010/01/04/points-inches-and-emus-measuring-units-in-office-open-xml/) for more information and how to convert units.

### Document

`Document` is the class which is able to write a .docx file. To create a new document, it is very easy:

```cpp
Document doc("./a.docx");
```

The last step is to save the document:

```cpp
doc.Save();
```

### Paragraph

`Paragraph` is the class that represents a paragraph. You can create paragraphs in the following ways:

```cpp
auto p1 = doc.AppendParagraph(); // Append a new document to the document
auto p3 = doc.InsertParagraphAfter(p1);  // Insert a new document after p1
auto p2 = doc.InsertParagraphBefore(p3); // Insert a new document before p3
```

It is easy to get every paragraph of the document:

```cpp
auto p1 = doc.FirstParagraph();
auto p3 = doc.LastParagraph();
auto p2 = p3.Prev(); // Next() also available
```

Paragraphs can be removed:

```cpp
doc.RemoveParagraph(p3); // Remove p3
```

You can check if two `Paragraph` instances represent the same paragraph:

```cpp
if (p1 == p2) {
  std::cout << "They're the same paragraph\n";
}
```

### Run

You can add multiple runs in paragraph:

```cpp
auto p4 = doc.AppendParagraph();
auto p4r1 = p4.AppendRun("Hello, World!");
auto p4r2 = p4.AppendRun(u8"你好，世界！");
auto p4r3 = p4.AppendRun(u8"你好，World!");
```

When you create a new paragraph by providing text to the `AppendParagraph()` method, it gets put into a single run.

```cpp
auto p5 = doc.AppendParagraph("Hello, World!");
// is equivalent to:
auto p5 = doc.AppendParagraph();
auto p5r1 = p5.AppendRun("Hello, World!");
```

Font size and typeface can be specified when adding a new run.

```cpp
auto p5r2 = p5.AppendRun("Hello, World!", 12, "Times New Roman");
```

Font size is specified in points.

You can create an empty run and add text to it later.

```cpp
auto p5r3 = p5.AppendRun();
p5r3.AppendText("Hello, World!");
p5r3.AppendText(u8"你好，世界！");
p5r3.AppendText(u8"你好，World!");
```

You can get the text contained in the run:

```cpp
auto text = p5r3.GetText(); // "Hello, World!你好，世界！你好，World!"
```

You can set character formatting for a run after it is created.

```cpp
p5r3.SetFontSize(14);
p5r3.SetFont("Times New Roman");
p5r3.SetFontStyle(Run::Bold | Run::Italic); // Run::Underline and 
                                            // Run::Strikethrough also available
p5r3.SetCharacterSpacing(Pt2Twip(2));
```

Helper function `Pt2Twip()` can be used to specify a character spacing in points.

It is easy to get each run in a paragraph:

```cpp
auto p4r1 = p4.FirstRun(); // no LastRun()
auto p4r2 = p4r1.Next(); // no Prev()
```

Run can be removed:

```cpp
p4r2.Remove();
```

#### Line Break

You can append a line break into a run. 

```cpp
auto r = p.AppendRun();
r.AppendText("This is");
r.AppendLineBreak();
r.AppendText("a simple sentence.");
```

#### Page Break

Page breaks are special run that can only be appended to a paragraph by calling `AppendPageBreak()`.

```cpp
auto r = p.AppendPageBreak();
```

### Section

Each document contains at least one section and you can't remove it.

To create a new section, you need to insert a section break. Section break needs to be inserted into a paragraph. So, you need to prepare a paragraph first, and then insert a section break to it.

```cpp
p3.InsertSectionBreak();
```

The paragraph containing a section break will be the last paragraph of the new section.

You can check if the paragraph contain a section break.

```cpp
if (p3.HasSectionBreak()) {
  std::cout << "p3 is the last paragraph of this section\n";
}
```

Section break can be removed:

```cpp
p3.RemoveSectionBreak();
```

`Section` is the class that represents a Section. You can get every section of the document:

```cpp
auto s1 = doc.FirstSection();
auto s2 = s1.Next(); // Prev() also available
```

You can get the first/last paragraph of a section.

```cpp
auto p1 = s1.FirstParagraph();
```

You can check if two `Section` instances represent the same section:

```cpp
if (s1 == s2) {
  std::cout << "They're the same Section\n";
}
```

You can set page formatting for a section.

```cpp
s1.SetPageSize(MM2Twip(297), MM2Twip(420)); // A3
s1.SetPageOrient(Section::Orientation::Landscape);
```

#### Page Number

The page numbers for pages in the section can be added and removed by using `SetPageNumber()` and `RemovePageNumber()`:

```cpp
s1.SetPageNumber(docx::Section::PageNumberFormat::Decimal);         // 1, 2, 3, ...
s1.SetPageNumber(docx::Section::PageNumberFormat::NumberInDash, 3); // -3-, -4-, -5-, ...
s1.RemovePageNumber();
```

### Table

You can insert a table by using `Document::AppendTable()`.

```cpp
auto tbl = doc.AppendTable(2, 3);
```

Each cell already contains a paragraph.

```cpp
tbl.GetCell(0, 0).FirstParagraph().AppendRun("AAA");
tbl.GetCell(0, 1).FirstParagraph().AppendRun("BBB");
tbl.GetCell(0, 2).FirstParagraph().AppendRun("CCC");

tbl.GetCell(1, 0).FirstParagraph().AppendRun("DDD");
tbl.GetCell(1, 1).FirstParagraph().AppendRun("EEE");
```

| AAA  | BBB  | CCC  |
| ---- | ---- | ---- |
| DDD  | EEE  |      |

You can change the style, width (points) and color (hex) of borders.

```cpp
tbl.SetTopBorders(Table::BorderStyle::Single, 1, "FF0000");      // a single line, 1 pt, Red
tbl.SetBottomBorders(Table::BorderStyle::Dotted, 2, "00FF00");   // a dotted line, 2 pt, Green
tbl.SetLeftBorders(Table::BorderStyle::Dashed, 3, "0000FF");     // a dashed line, 3 pt, Blue
tbl.SetRightBorders(Table::BorderStyle::DotDash, 0.5, "FFFF00"); // a line with alternating dots and dashes, 1/2 pt, yellow
tbl.SetInsideHBorders(Table::BorderStyle::Double, 1, "FF00FF");  // a double line, 1 pt, purple
```

#### Merge cells

You can merge adjacent cells with the same number of rows or columns.

```cpp
auto c00 = tbl.GetCell(0, 0);
auto c01 = tbl.GetCell(0, 1);
if (tbl.MergeCells(c00, c01)) {
  std::cout << "c00 c01 merged\n";
}
```

### Text Frame

A text frame is similar to a text box. Both are containers for text that can be positioned on a page and sized. Text boxes have more flexibility for formatting.

A text frame paragraph is simply a paragraph.

```cpp
auto frame = doc.AppendTextFrame(CM2Twip(4), CM2Twip(5));
frame.AppendRun("This is the text frame paragraph.");

frame.SetPositionX(TextFrame::Position::Left, TextFrame::Anchor::Page);
frame.SetPositionY(TextFrame::Position::Top,  TextFrame::Anchor::Margin);

frame.SetTextWrapping(TextFrame::Wrapping::Around);
``` 

## Contact

Do you have an issue using minidocx? Feel free to let me know on [issue tracker](https://github.com/totravel/minidocx/issues).

## Licensing

This library is available to anybody free of charge, under the terms of MIT License (see LICENSE.md).
