/**
 * Copyright (C) 2022-2025, Xie Zequn <totravel@foxmail.com>. All rights reserved.
 * Distributed under the MIT License (http://opensource.org/licenses/MIT)
 * Report bugs and download new versions at https://github.com/totravel/minidocx
 */

#pragma once

#include "word/main/container.hpp"
#include "word/main/properties/section.hpp"


namespace MINIDOCX_NAMESPACE
{
  class MINIDOCX_API Section : public Container
  {
  public:
    SectionProperties prop_;
  };
}
