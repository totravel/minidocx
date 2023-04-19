
# minidocx

![License](https://img.shields.io/github/license/totravel/minidocx)
![Documentation Status](https://img.shields.io/badge/中文文档-最新-brightgreen.svg)
![GitHub code size in bytes](https://img.shields.io/github/languages/code-size/totravel/minidocx)
![GitHub last commit](https://img.shields.io/github/last-commit/totravel/minidocx)

[English](./README.md) | 简体中文

minidocx 是一个跨平台且易于使用的 C++ 库，用于从零开始创建 Microsoft Word 文档 (.docx)。它被设计为足够精简和小巧。因此，你只需要将几个文件与你项目的其他源文件一同编译即可。

## 环境要求

要构建 minidocx，你需要一个支持 C++ 11 的编译器和下列两个第三方库：

- [zip](https://github.com/kuba--/zip) <= 0.2.1
- [pugixml](https://github.com/zeux/pugixml) >= 1.13

已在下列平台测试通过：

- CMake 3.21
- Visual Studio 14 2015 / Visual Studio 16 2019 / GCC 8.2+
- Windows 7/10 / Ubuntu 18.04 / Kylin-Desktop V10-SP1

## 参考文献

要处理 `.docx` 文档，你至少要了解下列 3 个概念：

- `Paragraph` 段落，具有相同段落格式的文本
- `Run` 富文本，段落中具有相同字体格式的文本
- `Section` 分节，具有相同页面设置的一个或多个页面

更多信息见 [此网站](http://officeopenxml.com/)。

## 示例

下面是一个使用 minidocx 创建 `.docx` 文件的示例。

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

更多示例见 [examples](./examples) 文件夹。

## 构建命令

```bash
git clone git@gitee.com:totravel/minidocx.git
cd minidocx

# Windows
cmake -S . -B build -DBUILD_EXAMPLES=ON -DWITH_STATIC_CRT=OFF
cmake --build build -j4 --config Release
cmake --install build --prefix install --config Release

# Linux
cmake -S . -B build -DBUILD_EXAMPLES=ON -DWITH_STATIC_CRT=OFF -DCMAKE_BUILD_TYPE=Release
cmake --build build -j4
cmake --install build --prefix install
```

## 指南

使用 minidocx 的程序只需包含头文件 `minidocx.hpp`。minidocx 提供的类和函数都定义在命名空间 `docx` 中。

```cpp
#include "minidocx.hpp"

using namespace docx;
```

下文将介绍常用 API。如需了解所有 API，请查阅头文件。

### 单位

文档的基本单位是缇（Twip），它等于磅的二十分之一。minidocx 提供了一些辅助函数用于单位换算：

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

有关单位转换的更多信息，请参阅 [Lars Corneliussen 的博文](https://startbigthinksmall.wordpress.com/2010/01/04/points-inches-and-emus-measuring-units-in-office-open-xml/)。

### 文档

类 `Document` 表示一个 Microsoft Word 文档。

```cpp
Document doc("./a.docx");
```

保存文档的内容：

```cpp
doc.Save();
```

### 段落

类 `Paragraph` 表示一个段落。有多种方法可以新建段落：

```cpp
auto p1 = doc.AppendParagraph();         // 在文档结尾追加新段落
auto p3 = doc.InsertParagraphAfter(p1);  // 在 p1 之后插入新段落
auto p2 = doc.InsertParagraphBefore(p3); // 在 p3 之前插入新段落
```

下列方法用于遍历文档的段落：

```cpp
auto p1 = doc.FirstParagraph();
auto p3 = doc.LastParagraph();
auto p2 = p3.Prev(); // 还可以用 Next() 方法
```

段落可以被移除：

```cpp
doc.RemoveParagraph(p3); // 移除 p3
```

可以检查两个 `Paragraph` 对象是否是同一个段落：

```cpp
if (p1 == p2) {
  std::cout << "They're the same paragraph\n";
}
```

### 富文本

向一个段落添加富文本：

```cpp
auto p4 = doc.AppendParagraph();
auto p4r1 = p4.AppendRun("Hello, World!");
auto p4r2 = p4.AppendRun(u8"你好，世界！");
auto p4r3 = p4.AppendRun(u8"你好，World!");
```

可以在新建段落的同时添加富文本：

```cpp
auto p5 = doc.AppendParagraph("Hello, World!");
// is equivalent to:
auto p5 = doc.AppendParagraph();
auto p5r1 = p5.AppendRun("Hello, World!");
```

添加富文本时可以指定字号和字体（西文和中文字体可以分别设置）：

```cpp
auto p5r2 = p5.AppendRun("Hello, World!", 12, "Times New Roman");
```

字号以磅为单位。

可以创建空的富文本。富文本在创建之后都可以继续添加更多文本。

```cpp
auto p5r3 = p5.AppendRun();
p5r3.AppendText("Hello, World!");
p5r3.AppendText(u8"你好，世界！");
p5r3.AppendText(u8"你好，World!");
```

可以获取富文本包含的文本：

```cpp
auto text = p5r3.GetText(); // "Hello, World!你好，世界！你好，World!"
```

设置字体格式:

```cpp
p5r3.SetFontSize(14);
p5r3.SetFont("Times New Roman");
p5r3.SetFontStyle(Run::Bold | Run::Italic); // 加粗、倾斜
                                            // 还可以用下划线 Run::Underline 和删除线 Run::Strikethrough
p5r3.SetCharacterSpacing(Pt2Twip(2));
```

要以磅为单位指定字间距，可用辅助函数 `Pt2Twip()`。

下列方法用于遍历段落的富文本：

```cpp
auto p4r1 = p4.FirstRun(); // 没有 LastRun() 方法
auto p4r2 = p4r1.Next(); // 没有 Prev() 方法
```

富文本可以被移除：

```cpp
p4r2.Remove();
```

#### 换行

可以添加换行符到富文本中：

```cpp
auto r = p.AppendRun();
r.AppendText("This is");
r.AppendLineBreak();
r.AppendText("a simple sentence.");
```

#### 分页

分页符是特殊的富文本，只能通过调用 `AppendPageBreak()` 函数将分页符添加到段落中。

```cpp
auto r = p.AppendPageBreak();
```

### 分节

任何文档都至少包含一个分节且不可删除。

新建分节需要插入分节符。分节符需要插入到某个段落中：

```cpp
p3.InsertSectionBreak();
```

包含分节符的段落将成为分节的最后一个段落。

可以检查一个段落是否包含分节符：

```cpp
if (p3.HasSectionBreak()) {
  std::cout << "p3 是这一节的最后一个段落\n";
}
```

段落中的分节符可以被移除：

```cpp
p3.RemoveSectionBreak();
```

类 `Section` 表示一个分节。下列方法用于遍历文档的分节：

```cpp
auto s1 = doc.FirstSection();
auto s2 = s1.Next(); // 还可以用 Prev() 方法
```

下列方法用于遍历分节的段落：

```cpp
auto p1 = s1.FirstParagraph();
```

可以检查两个 `Section` 对象是否是同一分节：

```cpp
if (s1 == s2) {
  std::cout << "They're the same Section\n";
}
```

可以设置分节的页面格式：

```cpp
s1.SetPageSize(MM2Twip(297), MM2Twip(420));        // 纸张大小为 A3
s1.SetPageOrient(Section::Orientation::Landscape); // 纸张方向为横向
```

#### 页码

底部居中的页码可以用 `SetPageNumber()` 方法添加，用 `RemovePageNumber()` 方法移除：

```cpp
s1.SetPageNumber(docx::Section::PageNumberFormat::Decimal);         // 1, 2, 3, ...
s1.SetPageNumber(docx::Section::PageNumberFormat::NumberInDash, 3); // -3-, -4-, -5-, ...
s1.RemovePageNumber();
```

### 表格

要插入表格，可以用 `Document` 类的 `AppendTable()` 方法。比如，插入一个 2 行 3 列的表格：

```cpp
auto tbl = doc.AppendTable(2, 3);
```

每个单元格都已经包含一个段落。

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

可以设置边框的样式、宽度（磅）和颜色（十六进制）：

```cpp
tbl.SetTopBorders(Table::BorderStyle::Single, 1, "FF0000");      // 单直线, 1 磅, 红色
tbl.SetBottomBorders(Table::BorderStyle::Dotted, 2, "00FF00");   // 点虚线, 2 磅, 绿色
tbl.SetLeftBorders(Table::BorderStyle::Dashed, 3, "0000FF");     // 短划线, 3 磅, 蓝色
tbl.SetRightBorders(Table::BorderStyle::DotDash, 0.5, "FFFF00"); // 点划线, 1/2 磅，黄色
tbl.SetInsideHBorders(Table::BorderStyle::Double, 1, "FF00FF");  // 双直线, 1 磅, 紫色
```

#### 合并单元格

可以合并相邻且具有相同行数或列数的单元格：

```cpp
auto c00 = tbl.GetCell(0, 0);
auto c01 = tbl.GetCell(0, 1);
if (tbl.MergeCells(c00, c01)) {
  std::cout << "c00 c01 merged\n";
}
```

### 图文框

图文框类似于文本框，但比文本框简单。图文框是特殊的段落，可以 `Document::AppendTextFrame()` 方法添加。

```cpp
auto frame = doc.AppendTextFrame(CM2Twip(4), CM2Twip(5));
frame.AppendRun("This is the text frame paragraph.");
```

设置图文框的位置：

```cpp
frame.SetPositionX(TextFrame::Position::Left, TextFrame::Anchor::Page);   // 相对于页面
frame.SetPositionY(TextFrame::Position::Top,  TextFrame::Anchor::Margin); // 相对于页边距
```

设置文本环绕：

```cpp
frame.SetTextWrapping(TextFrame::Wrapping::Around);
```

## 反馈

有任何疑问，可随时在 [此处](https://github.com/totravel/minidocx/issues) 提问。

## 许可

根据 MIT 许可条款，任何人都可以免费使用这个库（参见 License.md）。
