/**
 * Copyright (C) 2022-2025, Xie Zequn <totravel@foxmail.com>. All rights reserved.
 * Distributed under the MIT License (http://opensource.org/licenses/MIT)
 * Report bugs and download new versions at https://github.com/totravel/minidocx
 */

#include "utils/string.hpp"


namespace MINIDOCX_NAMESPACE
{
  std::string removeSpaces(std::string str) {
    std::string tmp{ std::move(str) };
    tmp.erase(std::remove_if(tmp.begin(), tmp.end(), std::isspace), tmp.end());
    return tmp;
  }
}
