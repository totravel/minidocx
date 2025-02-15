/**
 * Copyright (C) 2022-2025, Xie Zequn <totravel@foxmail.com>. All rights reserved.
 * Distributed under the MIT License (http://opensource.org/licenses/MIT)
 * Report bugs and download new versions at https://github.com/totravel/minidocx
 */

#pragma once

#include "utils/base.hpp"


namespace NAMESPACE
{
  enum class BlockType { Paragraph, Table };
  using Block = Node<BlockType>;

  enum class RunType { RichText, Picture };
  using Run = Node<RunType>;

}
