
#include "utils/exceptions.hpp"


namespace NAMESPACE
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
