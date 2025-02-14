
English | [简体中文](./README-zh_CN.md)

# minidocx

minidocx is a free, open-source, cross-platform, modern, light-weight and user-friendly C++20 library for creating Microsoft Word Document (.docx file) as described in [ECMA 376 5th edition](https://www.ecma-international.org/publications-and-standards/standards/ecma-376) or [ISO/IEC 29500-1:2016](https://www.iso.org/standard/71691.html) without installing MS Office or WPS Office.

> [!WARNING]
> minidocx 1.0 is currently in beta and should not be used in production.

> [!NOTE]
> Check out the master branch to view minidocx 0.6.

## Preview

Light Mode | Dark Mode
---------- | ---------
![](./screenshots/20250214232857.png) | ![](./screenshots/20250214233038.png)

## Features

- Section
- Paragraph
- Rich text
- Table
- Picture
- Style
- List

## Example

Here's an example of how to use minidocx to create a .docx file.

```cpp
#include "minidocx/minidocx.hpp"
#include <iostream>

int main()
{
  using namespace md;
  try {
    Document doc;
    SectionPointer sect = doc.addSection();

    ParagraphPointer para = sect->addParagraph();
    para->properties().align_ = Alignment::Centered;

    RichTextPointer rich = para->addRichText("Happy Chinese New Year!");
    rich->properties().fontSize_ = 32;
    rich->properties().color_ = "FF0000";

    doc.saveAs("a.docx");
  }
  catch (const Exception& ex) {
    std::cerr << ex.what() << std::endl;
  }
  return 0;
}
```

## Building

To build minidocx lib you'll need a C++20 compiler and CMake 3.28.

```bash
git clone git@github.com:totravel/minidocx.git
cd minidocx

# Windows
cmake --preset x64-win-msbuild-v143               # Configure
cmake --build --preset x64-win-msbuild-v143-debug # Build
out/x64-win-msbuild-v143/examples/Debug/myapp.exe # Run

# Linux
cmake --preset x64-linux-ninja-gcc
cmake --build --preset x64-linux-ninja-gcc-debug
out/x64-linux-ninja-gcc/examples/myapp
```

## User Guide

### Document Structure

A document consists of one or more sections. A section is a special container that have a specific set of properties used to define the pages on which its contents will appear, such as page size, page orientation, and page margins.

A container can contain two different types of block-level content: paragraphs and tables.
 
A paragraph is a division of content that begins on a new line with a common set of properties, such as outline level, alignment, indentation, spacing, and borders. A paragraph can contain two different types of non-block content: texts and pictures.

Tables are another type of block-level content. A table is composed of a collection of cells. Cells are also containers.

- Document
  - Section (Container)
    - Paragraph
      - Text
      - Picture
    - Table
      - Cell (Container)

### Measuring units

The main unit in OOXML is a twentieth of a point. This is used for specifying page dimensions, margins, tabs, etc.

- Inch (in)
- Point (pt)
- Twentieth of a point (tw)
- English Metric Unit (emu)

The relationship between them is shown in the table below.

|   mm |   cm |   in |   pt |   tw |    emu |
| ---: | ---: | ---: | ---: | ---: | -----: | 
|    1 |      |      |      |      |  36000 |
|      |    1 |      |      |      | 360000 |
| 25.4 | 2.54 |    1 |   72 | 1440 | 914400 |
|      |      |      |    1 |   20 |  12700 |
|      |      |      |      |    1 |    635 |

For more information, see [Lars Corneliussen's blog post](https://startbigthinksmall.wordpress.com/2010/01/04/points-inches-and-emus-measuring-units-in-office-open-xml/).

### Quick Start

`minidocx.hpp` is the only one header you need to include in order to have access to all functions of minidocx and so that you do not have to care about the order of includes. All minidocx classes are member of the `md` namespace.

```cpp
#include "minidocx/minidocx.hpp"
using namespace md;
```

All minidocx functions will throw an exception in case of an error. You should catch the exception to either fix it or report back to the user. All exceptions minidocx throws are objects of the class `Exception`. That's why we simply catch `Exception` objects.

```cpp
try
{
  // Do something
}
catch (const Exception& ex)
{
  std::cerr << ex.what() << std::endl;
}
```

A Document is represented by a `Document` object. To create a new document and save it as `example.docx`:

```cpp
Document doc;
doc.properties().title_ = "Chinese New Year";

SectionPointer sect = doc.addSection();
sect->properties().landscape_ = true;

ParagraphPointer para = sect->addParagraph();
para->properties().align_ = Alignment::Centered;

RichTextPointer rich = para->addRichText("Happy Chinese New Year!");
rich->properties().color_ = "FF0000";

doc.saveAs("example.docx");
```

All properties can be access via `properties()` or `setProperties()` method. See other avaliable properties in `include/properties` directory.

Following sections describe other features supported by minidocx. Please note this description may not be complete but limited to the most useful ones. If you want to find less common features, please check header files under `include` directory.

### Tables

...

### Pictures

...

### Styles

...

### Lists

...

## Donation

...

## Sponsor

You can sponsor this library at [AFDIAN](https://afdian.com/a/totravel).

Your sponsorship means a lot to me. It will help me sustain my projects actively and make more of my ideas come true. Much appreciated! 💖 🙏

## License

MIT
