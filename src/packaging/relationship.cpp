
#include "packaging/relationship.hpp"


namespace NAMESPACE
{
  RelationshipId Relationship::parseId(const std::string& id)
  {
    return std::stoi(id.c_str() + 3);
  }

  std::string Relationship::stringifyId(const RelationshipId id)
  {
    return "rId" + std::to_string(id);
  }

  Relationship::TargetMode Relationship::parseTargetMode(const std::string& mode)
  {
    if (mode == "Internal")
      return TargetMode::Internal;
    else if (mode == "External")
      return TargetMode::External;
    else
      return TargetMode::Unknown;
  }
}
