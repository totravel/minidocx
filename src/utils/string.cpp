
#include "utils/string.hpp"


namespace NAMESPACE
{
  std::string removeSpaces(std::string str) {
    std::string tmp{ std::move(str) };
    tmp.erase(std::remove_if(tmp.begin(), tmp.end(), std::isspace), tmp.end());
    return tmp;
  }
}
