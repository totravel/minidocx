
English | [简体中文](./README-zh_CN.md)

<div align="center">
  <img src="./assets/logo.png" width="100px">
  <h1>minidocx</h1>
  <p>C++ library for manipulating Microsoft Word Document</p>
</div>

## About

minidocx is a modern, free, open-source, cross-platform, light-weight, and user-friendly C++20 library for manipulating Microsoft Word Document (.docx file) as described in [ECMA 376 5th edition](https://www.ecma-international.org/publications-and-standards/standards/ecma-376) or [ISO/IEC 29500-1:2016](https://www.iso.org/standard/71691.html) without installing MS Office or WPS Office.

> [!WARNING]
> minidocx 1.0 is currently in beta and should not be used in production.

> [!NOTE]
> Check out the master branch to view minidocx 0.6.

## Features

- Section
- Paragraph
- Rich text
- Table
- Picture
- Style
- List

## Preview

Light Mode | Dark Mode
---------- | ---------
![](./assets/screenshots/20250214232857.png) | ![](./assets/screenshots/20250214233038.png)

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
    para->prop_.align_ = Alignment::Centered;

    RichTextPointer rich = para->addRichText("Happy Chinese New Year!");
    rich->prop_.fontSize_ = 32;
    rich->prop_.color_ = "FF0000";

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
cmake --preset x64-win-msbuild-v143
cmake --build --preset x64-win-msbuild-v143-debug
./out/x64-win-msbuild-v143/bin/exe/Debug/myapp.exe

# Linux
cmake --preset x64-linux-ninja-gcc
cmake --build --preset x64-linux-ninja-gcc-debug
./out/x64-linux-ninja-gcc/bin/exe/myapp
```

A static library is built by default. If you want to use a shared build of minidocx, set the `BUILD_SHARED` CMake option to `true`.

## Documentation

- [User Guide](./guide.md)

## Donation

If you benefit from this project, please consider donating to help me sustain my projects actively and make more of my ideas come true.

Alipay | WeChat Pay
------ | ----------
![](./assets/qrcode/alipay.png) | ![](./assets/qrcode/wechat.png)

## Sponsor

You can sponsor this library at [AFDIAN](https://afdian.com/a/totravel).

Your sponsorship means a lot to me. It will help me sustain my projects actively and make more of my ideas come true. Much appreciated! 💖 🙏

## License

minidocx is released to the public for free under the terms of the MIT License. See [LICENSE](./LICENSE) for the full text of the license. [LICENSE](./LICENSE) should be distributed alongside any assemblies that use minidocx in source or compiled form.
