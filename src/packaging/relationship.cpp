/**
 * Copyright (C) 2022-2025, Xie Zequn <totravel@foxmail.com>. All rights reserved.
 * Distributed under the MIT License (http://opensource.org/licenses/MIT)
 * Report bugs and download new versions at https://github.com/totravel/minidocx
 */

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
