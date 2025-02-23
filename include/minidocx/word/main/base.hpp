/**
 * Copyright (C) 2022-2025, Xie Zequn <totravel@foxmail.com>. All rights reserved.
 * Distributed under the MIT License (http://opensource.org/licenses/MIT)
 * Report bugs and download new versions at https://github.com/totravel/minidocx
 */

#pragma once

#include "utils/base.hpp"


namespace MINIDOCX_NAMESPACE
{
  enum class BlockType { Paragraph, Table };
  enum class RunType { RichText, Picture };
  
  using Block = Node<BlockType>;
  using Run = Node<RunType>;

}
