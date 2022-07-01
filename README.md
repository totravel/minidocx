
# minidocx

![License](https://img.shields.io/github/license/totravel/minidocx)
![Lines of code](https://img.shields.io/tokei/lines/github/totravel/minidocx)
![GitHub code size in bytes](https://img.shields.io/github/languages/code-size/totravel/minidocx)
![GitHub last commit](https://img.shields.io/github/last-commit/totravel/minidocx)

English | [简体中文](./README-zh_CN.md)

minidocx is a portable and easy-to-use C++ library for creating Microsoft Word Document (.docx). It is designed to be simple and small enough. So, you can grab these 2 files and compile them into your project.

## Requirements

To build minidocx you'll need a C++11 compiler and the following libraries:

- [zip](https://github.com/kuba--/zip) 0.2.1
- [pugixml](https://github.com/zeux/pugixml) 1.12.1

It is tested with the following IDE/toolchains versions:

- Visual Studio 16 2019
- GNU 8.2.0

## References

To manipulate a docx file, you need to understand the following concepts:

- `Paragraph` Text with the same paragraph formatting 
- `Run` Text with the same font formatting. Paragraph contain 
- `Section` A division of a document having the same page layout settings, such as margins and page orientation

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
  auto p2 = doc.AppendParagraph("你好，世界！", 14, "宋体");
  auto p3 = doc.AppendParagraph("你好，World!", 16, "Times New Roman", "宋体");
  
  auto p4 = doc.AppendParagraph();
  p4.SetAlignment(Paragraph::Alignment::Centered);

  auto p4r1 = p4.AppendRun("This is a simple sentence. ", 12, "Arial");
  p4r1.SetCharacterSpacing(Pt2Twip(2));

  auto p4r2 = p4.AppendRun("这是一个简单的句子。");
  p4r2.SetFontSize(14);
  p4r2.SetFont("黑体");
  p4r2.SetFontStyle(Run::Bold | Run::Italic);

  doc.Save();
  return 0;
}
```

See other [examples](./examples).

## Building From Source

The minidocx source consists of 2 files - one source file, `minidocx.cpp`, and one header file, `minidocx.hpp`.

The easiest way to build minidocx is to compile the source file, `minidocx.cpp`, along with the existing library/executable. If you’re using CMake, just add the following commands to the `CMakeLists.txt` file in your projects.

```cmake
project(myproj VERSION 0.1.0 LANGUAGES C CXX) # C needed by zip.c

add_library(zip INTERFACE)
set_target_properties(zip PROPERTIES
  INTERFACE_INCLUDE_DIRECTORIES "${ZIP_DIR}/src"
  INTERFACE_SOURCES             "${ZIP_DIR}/src/zip.c"
)

add_library(pugixml INTERFACE)
set_target_properties(pugixml PROPERTIES
  INTERFACE_INCLUDE_DIRECTORIES "${PUGIXML_DIR}/src"
  INTERFACE_SOURCES             "${PUGIXML_DIR}/src/pugixml.cpp"
)

add_library(minidocx INTERFACE)
set_target_properties(minidocx PROPERTIES
  INTERFACE_INCLUDE_DIRECTORIES "${MINIDOCX_DIR}/src"
  INTERFACE_SOURCES             "${MINIDOCX_DIR}/src/minidocx.cpp"
  INTERFACE_COMPILE_OPTIONS     "$<$<CXX_COMPILER_ID:MSVC>:/utf-8>"
  INTERFACE_LINK_LIBRARIES      "zip;pugixml"
)

target_link_libraries(myapp PRIVATE minidocx)
```

When running CMake to configure the build tree, the following variables need to be set correctly:

- `ZIP_DIR` zip directory
- `PUGIXML_DIR` pugixml directory
- `MINIDOCX_DIR` minidocx directory

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
auto p3 = p1.InsertAfter();      // Insert a new document after p1
auto p2 = p3.InsertBefore();     // Insert a new document before p3
```

It is easy to get every paragraph of the document:

```cpp
auto p1 = doc.FirstParagraph();
auto p3 = doc.LastParagraph();
auto p2 = p3.Prev(); // Next() also available
```

Paragraphs can be removed:

```cpp
p.Remove();
```

### Run

You can add multiple runs in paragraph:

```cpp
auto p4 = doc.AppendParagraph();
auto p4r1 = p4.AppendRun("Hello, World!");
auto p4r2 = p4.AppendRun("你好，世界！");
auto p4r3 = p4.AppendRun("你好，World!");
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
p5r3.AppendText("你好，世界！");
p5r3.AppendText("你好，World!");
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

You can set page formatting for a section.

```cpp
s1.SetPageSize(MM2Twip(297), MM2Twip(420)); // A3
s1.SetPageOrient(Section::Orientation::Landscape);
```

## Contact

Do you have an issue using minidocx? Feel free to let me know on [issue tracker](https://github.com/totravel/minidocx/issues).

## Licensing

This library is available to anybody free of charge, under the terms of MIT License (see LICENSE.md).
