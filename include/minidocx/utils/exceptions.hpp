
#pragma once

#include <stdexcept>
#include <string>


namespace NAMESPACE
{
  class exception : public std::runtime_error
  {
  public:
    exception(const std::string& message, const std::string& sender = "minidocx");
  };

  class unsupported_feature : public exception
  {
  public:
    unsupported_feature();
  };

  class invalid_parameter : public exception
  {
  public:
    invalid_parameter();
  };

  class invalid_operation : public exception
  {
  public:
    invalid_operation();
  };

  class io_error : public exception
  {
  public:
    io_error(const std::string& filename, const std::string& message);
  };
}
