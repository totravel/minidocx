
# User Guide

Following sections describe necessary information and features supported by minidocx. Please note this description may not be complete but limited to the most useful ones. If you want to find less common features, please check header files under `include` directory.

## Measuring Units

The measuring units used in the document mainly include point (pt), twentieth of a point (tw), and English Metric Unit (emu), which are used to specify font size, page size, table width, etc. The relationship between them is shown in the table below.

|   mm |   cm |   in |   pt |   tw |    emu |
| ---: | ---: | ---: | ---: | ---: | -----: | 
|    1 |      |      |      |      |  36000 |
|      |    1 |      |      |      | 360000 |
| 25.4 | 2.54 |    1 |   72 | 1440 | 914400 |
|      |      |      |    1 |   20 |  12700 |
|      |      |      |      |    1 |    635 |

For more information, see [Lars Corneliussen's blog post](https://startbigthinksmall.wordpress.com/2010/01/04/points-inches-and-emus-measuring-units-in-office-open-xml/).

## Data Structure

A document consists of the following objects:

- Document
  - Section (Container)
    - Paragraph (Block)
      - Text (Inline)
      - Picture (Inline)
    - Table (Block)
      - Cell (Container)

A document consists of one or more sections. A section is a special container that have a specific set of properties used to define the pages on which its contents will appear, such as page size, page orientation, and page margins.

A container can contain two different types of block-level objects: paragraphs and tables.
 
A paragraph is a division of content that begins on a new line with a common set of properties, such as outline level, alignment, indentation, spacing, and borders. A paragraph can contain two different types of inline objects: texts and pictures.

Tables are another type of block-level objects. A table is composed of a collection of cells. Cells are also containers.

## Headers and Namespace

`minidocx.hpp` is the only one header you need to include in order to have access to all functions of minidocx and so that you do not have to care about the order of includes. All minidocx classes are member of the `md` namespace.

```cpp
#include "minidocx/minidocx.hpp"
using namespace md;
```

If you are linking against a precompiled shared build of minidocx, add `MINIDOCX_SHARED` compilation definition before including `minidocx.hpp`.

## Error Handling

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

## Documents

A document is represented by a `Document` object. To create a new document and save it as `example.docx`:

```cpp
Document doc;
// Do something
doc.saveAs("example.docx");
```

The `prop_` public data member of the `Document` object is a `PackageProperties` object, which is used to store additional information about the document, such as title, subject, author, and company.

```cpp
doc.prop_.title_ = "Chinese New Year";
doc.prop_.author_ = "John";
doc.prop_.lastModifiedBy_ = "Peter";
```

See other avaliable document properties in [PackageProperties](./include/minidocx/packaging/package.hpp).

## Sections

A section is represented by a `Section` object which can be created by making a call to the `addSection()` method on a `Document` object:

```cpp
SectionPointer sect = doc.addSection();
```

The `prop_` public data member of the `Section` object is a `SectionProperties` object, which is used to store formatting properties for all pages in the section, such as page size, page orientaion, page margins, etc.

```cpp
sect->prop_.size_.width_ = A3_W;
sect->prop_.size_.height_ = A3_H;
sect->prop_.landscape_ = true;
```

See other avaliable section properties in [SectionProperties](./include/minidocx/word/main/properties/section.hpp).

## Paragraphs

A paragraph is represented by a `Paragraph` object which can be created by calling the `addParagraph()` method on a `Section` object:

```cpp
ParagraphPointer para = sect->addParagraph();
```

The `prop_` public data member of the `Paragraph` object is a `ParagraphProperties` object, which is used to store formatting properties for the paragraph, such as alignment, outline level, indentation, spacing, etc.

```cpp
para->prop_.align_ = Alignment::Centered;
para->prop_.outlineLevel_ = OutlineLevel::Level1;
```

See other avaliable paragraph properties in [ParagraphProperties](./include/minidocx/word/main/properties/paragraph.hpp).

## Rich Text

A sequence of characters with a set of properties is represented by a `RichText` object which can be created by calling the `addRichText()` method on a `Paragraph` object with a piece of text encoded in UTF-8 as argument. Note that all characters, including font names mentioned below, should be encoded in UTF-8.

```cpp
RichTextPointer rich = para->addRichText(u8"Happy Chinese New Year!\n中国新年快乐！");
```

As you can see, the escape character `\n` (line break) is allowed. Note that the tab character `\t` is also allowed but the carriage return character `\r` is omitted. 

The `prop_` public data member of the `RichText` object is a `RichTextProperties` object, which is used to store formatting properties for the text, such as font family, font size, font color, highlight, spacing, etc.

```cpp
rich->prop_.font_ = { .ascii_ = "Aria", .eastAsia_ = "Simsun" };
rich->prop_.fontSize_ = 32;
rich->prop_.color_ = "FF0000";
```

See other avaliable properties in [RichTextProperties](./include/minidocx/word/main/properties/richtext.hpp).
