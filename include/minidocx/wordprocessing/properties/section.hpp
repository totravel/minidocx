
#pragma once

#include "wordprocessing/properties/base.hpp"
#include "wordprocessing/constants.hpp"


namespace NAMESPACE
{
  struct SectionProperties
  {
    // Section Type 分节符类型
    enum class Type {
      NextPage, // 下一页
    } type_ = Type::NextPage;


    // Page Size 纸张大小
    // Specifies the size and orientation for all pages in the section.
    // 指定分节中所有页面的大小和方向。
    struct Size {
      // Specifies the page size in twentieths of a point (tw).
      // 以缇为单位指定页面大小。
      size_t width_ = A4_W;
      size_t height_ = A4_H;
      bool landscape_ = false; // Landscape 横向
    } size_;


    // Page Margins 页边距
    struct Margins {
      size_t top_ = 1440;
      size_t bottom_ = 1440;

      size_t left_ = 1800;
      size_t right_ = 1800;

      size_t header_ = 992;
      size_t footer_ = 851;

      // Page Gutter Spacing 装订线
      size_t gutter_ = 0;
    } margins_;
    // top_ = bottom_ = 2.54 cm
    // left_ = right_ = 3.17 cm
    // header_ = 1.75 cm
    // footer_ = 1.5 cm


    // Document grid 文档网格
    // This element specifies the settings for the document grid, 
    // which enables precise layout of full-width East Asian
    // language characters within a document by specifying the 
    // desired number of characters per line and lines per page
    // for all East Asian text content in this section.
    // 通过指定每行字符数和每页行数为全角东亚语言符号提供精确布局。
    struct DocGrid {

      // Document Grid Type 文档网格类型
      enum class Type {
        Default,       // No Document Grid 无网格
        Lines,         // Line Grid Only 只指定行网格
        LinesAndChars, // Line and Character Grid 指定行和字符网格
        SnapToChars,   // Character Grid Only 文字对齐字符网格
      } type_ = Type::Lines;

      // Specifies the number of lines to be allowed on the document
      // grid for the current page assuming all lines have equal line
      // pitch applied to them. 由行间距表示的每页行数。
      size_t linePitch_ = 312;
      // 312 = 15.6 pt (about 44 lines)
      // 312 = 15.6 磅（约 44 行）
    };

    std::optional<DocGrid> docGrid_;
  };
}