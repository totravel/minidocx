
[English](./README.md) | 简体中文

# minidocx

minidocx 是一个免费、开源、跨平台、现代、轻量、易用的 C++20 库，用于生成 Word 文档（.docx 文件），不依赖 Word 或 WPS。

> **警告**
> minidocx 1.0 仍为预览版，不建议在生产环境中使用。

> **备注**
> minidocx 0.6 可在 master 分支查看。

## 特性

- 分节
- 段落
- 富文本
- 表格
- 图片
- 列表
- 样式

## 截图

以下是一些由 minidocx 生成的样本的截图。

Sources              | Screenshots
-------------------- | -------------------------------------
examples/main.cpp    | ![](./screenshots/20250203000323.png)
examples/picture.cpp | ![](./screenshots/20250203000246.png)

## 示例

以下是使用 minidocx 创建 Word 文档的示例。

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
  catch (const exception& ex) {
    std::cerr << ex.what() << std::endl;
  }
  return 0;
}
```

## 编译

编译 minidocx 需要 C++20 编译器和 CMake 3.28。

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

## 快速开始

施工中……

## 捐赠

如果我的项目对你有所帮助，请考虑给我一些鼓励！

- [爱发电](https://afdian.com/a/totravel)

## 许可

MIT
