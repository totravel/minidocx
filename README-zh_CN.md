
# minidocx

![License](https://img.shields.io/github/license/totravel/minidocx)
![Documentation Status](https://img.shields.io/badge/中文文档-最新-brightgreen.svg)
![Lines of code](https://img.shields.io/tokei/lines/github/totravel/minidocx)
![GitHub code size in bytes](https://img.shields.io/github/languages/code-size/totravel/minidocx)
![GitHub last commit](https://img.shields.io/github/last-commit/totravel/minidocx)

[English](./README.md) | 简体中文

minidocx 是一个跨平台且易于使用的 C++ 库，用于从零开始创建 Microsoft Word 文档 (.docx)。它被设计为足够精简和小巧。因此，你只需将它的两个文件与你的项目的其他源文件一同编译即可。

## 环境要求

要使用 minidocx，你需要一个支持 C++ 11 的编译器和下列两个第三方库：

- [zip](https://github.com/kuba--/zip) 0.2.1
- [pugixml](https://github.com/zeux/pugixml) 1.12.1

已测试的开发环境：

- Visual Studio 16 2019
- GNU 8.2.0

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

更多示例见 [examples](./examples) 文件夹。

## 构建

minidocx 包含 2 个文件——一个源文件 `minidocx.cpp` 和一个头文件 `minidocx.hpp`。

构建 minidocx 最简单的方式是将它的源文件与你的项目的其他源文件一同编译。如果你在使用 CMake，只需将下列命令添加到项目的 `CMakeLists.txt` 文件中。

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

首次运行 CMake 时，需要正确设置下列变量：

- `ZIP_DIR` zip 项目的根目录
- `PUGIXML_DIR` pugixml 项目的根目录
- `MINIDOCX_DIR` minidocx 项目的根目录

## 指南

使用 minidocx 的程序只需包含头文件 `minidocx.hpp`。minidocx 提供的类和函数都定义在命名空间 `docx` 中。

```cpp
#include "minidocx.hpp"

using namespace docx;
```

### 单位

文档的基本单位是磅的二十分之一（缇）。minidocx 提供了一些辅助函数用于单位换算：

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

类 `Paragraph` 表示一个段落。有三种方法可以新建段落：

```cpp
auto p1 = doc.AppendParagraph(); // 向文档追加新段落
auto p3 = p1.InsertAfter();      // 在 p1 之后插入新段落
auto p2 = p3.InsertBefore();     // 在 p1 之前插入新段落
```

下列方法用于遍历文档的段落：

```cpp
auto p1 = doc.FirstParagraph();
auto p3 = doc.LastParagraph();
auto p2 = p3.Prev(); // 还可以用 Next() 方法
```

段落可以被移除：

```cpp
p3.Remove();
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
auto p4r2 = p4.AppendRun("你好，世界！");
auto p4r3 = p4.AppendRun("你好，World!");
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
p5r3.AppendText("你好，世界！");
p5r3.AppendText("你好，World!");
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

## 联系

有任何疑问，可随时在 [此处](https://github.com/totravel/minidocx/issues) 提问。

## 许可

根据 MIT 许可条款，任何人都可以免费使用这个库（参见 License.md）。
