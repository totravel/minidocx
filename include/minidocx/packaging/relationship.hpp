/**
 * Copyright (C) 2022-2025, Xie Zequn <totravel@foxmail.com>. All rights reserved.
 * Distributed under the MIT License (http://opensource.org/licenses/MIT)
 * Report bugs and download new versions at https://github.com/totravel/minidocx
 */

#pragma once

#include "packaging/part.hpp"

#include <string>
#include <map>


namespace MINIDOCX_NAMESPACE
{
  using RelationshipId = size_t;

  struct Relationship
  {
    enum class TargetMode { Unknown, Internal, External };

    RelationshipId id_ = 0;
    PartType type_ = PartType::Unknown;
    PartName target_;
    TargetMode targetMode_ = TargetMode::Unknown;

    static RelationshipId parseId(const std::string& id)
    {
      return std::stoull(id.data() + 3);
    }

    static std::string stringifyId(const RelationshipId id)
    {
      return "rId" + std::to_string(id);
    }

    static TargetMode parseTargetMode(const std::string& mode)
    {
      if (mode == "Internal")
        return TargetMode::Internal;
      else if (mode == "External")
        return TargetMode::External;
      else
        return TargetMode::Unknown;
    }
  };

  using RelationshipMap = std::map<RelationshipId, Relationship>;

  struct Relationships
  {
    RelationshipId maxId_ = 0;
    RelationshipMap map_;

    void clear()
    {
      maxId_ = 0;
      map_.clear();
    }
  };
}
