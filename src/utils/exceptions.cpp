/**
 * Copyright (C) 2022-2025, Xie Zequn <totravel@foxmail.com>. All rights reserved.
 * Distributed under the MIT License (http://opensource.org/licenses/MIT)
 * Report bugs and download new versions at https://github.com/totravel/minidocx
 */

#include "utils/exceptions.hpp"


namespace MINIDOCX_NAMESPACE
{
  Exception::Exception(const std::string& message, const std::string& sender)
    : runtime_error{ sender + ": "  + message}
  {
  }

  unsupported_feature::unsupported_feature()
    : Exception("unsupported feature")
  {
  }

  invalid_parameter::invalid_parameter()
    : Exception("invalid parameter")
  {
  }

  invalid_operation::invalid_operation()
    : Exception("invalid operation")
  {
  }

  io_error::io_error(const std::string& filename, const std::string& message)
    : Exception(message + ": '" + filename + "'")
  {
  }
}
