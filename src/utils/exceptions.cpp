
#include "utils/exceptions.hpp"


namespace NAMESPACE
{
  exception::exception(const std::string& message, const std::string& sender)
    : runtime_error{ sender + ": "  + message}
  {
  }

  unsupported_feature::unsupported_feature()
    : exception("unsupported feature")
  {
  }

  invalid_parameter::invalid_parameter()
    : exception("invalid parameter")
  {
  }

  invalid_operation::invalid_operation()
    : exception("invalid operation")
  {
  }

  io_error::io_error(const std::string& filename, const std::string& message)
    : exception(message + ": '" + filename + "'")
  {
  }
}
