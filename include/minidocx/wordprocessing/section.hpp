/**
 * Copyright (C) 2022-2025, Xie Zequn <totravel@foxmail.com>. All rights reserved.
 * Distributed under the MIT License (http://opensource.org/licenses/MIT)
 * Report bugs and download new versions at https://github.com/totravel/minidocx
 */

#pragma once

#include "wordprocessing/container.hpp"
#include "wordprocessing/properties/section.hpp"


namespace NAMESPACE
{
  class Section : public Container, public Configurable<SectionProperties>
  {
  };
}
