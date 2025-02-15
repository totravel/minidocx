/**
 * Copyright (C) 2022-2025, Xie Zequn <totravel@foxmail.com>. All rights reserved.
 * Distributed under the MIT License (http://opensource.org/licenses/MIT)
 * Report bugs and download new versions at https://github.com/totravel/minidocx
 */

#pragma once

#include <stdexcept>
#include <string>


namespace NAMESPACE
{
  class Exception : public std::runtime_error
  {
  public:
    Exception(const std::string& message, const std::string& sender = "minidocx");
  };

  class unsupported_feature : public Exception
  {
  public:
    unsupported_feature();
  };

  class invalid_parameter : public Exception
  {
  public:
    invalid_parameter();
  };

  class invalid_operation : public Exception
  {
  public:
    invalid_operation();
  };

  class io_error : public Exception
  {
  public:
    io_error(const std::string& filename, const std::string& message);
  };
}
