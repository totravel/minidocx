/**
 * Copyright (C) 2022-2025, Xie Zequn <totravel@foxmail.com>. All rights reserved.
 * Distributed under the MIT License (http://opensource.org/licenses/MIT)
 * Report bugs and download new versions at https://github.com/totravel/minidocx
 */

#pragma once

#include "utils/zip.hpp"
#include "utils/file.hpp"
#include "packaging/part.hpp"
#include "packaging/relationship.hpp"

#include <string>
#include <map>
#include <vector>


namespace pugi { class xml_document; }

namespace MINIDOCX_NAMESPACE
{
  struct PackageProperties
  {
    std::string title_;
    std::string subject_;
    std::string author_;
    std::string company_;
    std::string lastModifiedBy_;
  };
  

  class Package : protected Zip
  {
  public:
    PackageProperties prop_;

  protected:
    void init();
    void flush();
    void clear();
    void load();

    void writePart(const PartName& name, pugi::xml_document& doc);

  private:
    std::map<PartName, Buffer> buffered_;

    void writeBufferedParts();

  protected:
    void addBufferedPart(const PartName& name, Buffer buf)
    {
      buffered_[name] = std::move(buf);
    }

  private:
    using MediaType = std::string;

    std::map<PartExtension, MediaType> defaultContentTypes_;
    std::map<PartName, MediaType> overrideContentTypes_;

    void initContentTypes();
    void writeContentTypes();
    void readContentTypes();

  protected:
    inline void registerDefaultContentType(
      const PartExtension& ext, const MediaType& type)
    {
      defaultContentTypes_[ext] = type;
    }

    inline void registerOverrideContentType(
      const PartName& name, const MediaType& type)
    {
      overrideContentTypes_[name] = type;
    }

    inline void unregisterOverrideContentType(const PartName& name)
    {
      overrideContentTypes_.erase(name);
    }

  private:
    std::map<PartName, Relationships> relationships_;
    Relationships packageRelationships_;

    inline RelationshipId addRelationshipTo(Relationships& rels,
      const PartType type, const PartName& target, const Relationship::TargetMode mode)
    {
      // TODO: Check if the target already exists.
      const RelationshipId id = ++rels.maxId_;
      rels.map_[id] = { id, type, target, mode };
      return id;
    }

    inline void deleteRelationshipFrom(Relationships& rels, const RelationshipId id)
    {
      rels.map_.erase(id);
    }

    void writeRelationships(Relationships& rels, const PartName& name);
    void writeRelationships();

    void readRelationships(const PartName& name, Relationships& rels);
    inline void readPkgRelationships()
    {
      readRelationships(RELS_PART, packageRelationships_);
    }

  protected:
    inline RelationshipId addRelationshipFor(const PartName& src,
      const PartType type, const PartName& target, const Relationship::TargetMode mode)
    {
      return addRelationshipTo(relationships_[src], type, fs::relative(target, src.parent_path()), mode);
    }

    inline RelationshipId addPkgRelationship(
      const PartType type, const PartName& target, const Relationship::TargetMode mode)
    {
      return addRelationshipTo(packageRelationships_, type, target.relative_path(), mode);
    }

    inline void deleteRelationshipFor(const PartName& src, const RelationshipId id)
    {
      deleteRelationshipFrom(relationships_[src], id);
    }

    inline void deletePkgRelationship(const RelationshipId id)
    {
      deleteRelationshipFrom(packageRelationships_, id);
    }

    inline void initRelationshipsFor(const PartName& src)
    {
      relationships_[src] = Relationships();
    }

    inline void readRelationshipsFor(const PartName& src)
    {
      initRelationshipsFor(src);
      readRelationships(src, relationships_[src]);
    }

  private:
    PartName corePart_{ CORE_PART };
    PartName appPart_{ APP_PART };

    void writeCoreProperties();
    void writeExtendedProperties();

    //void readCoreProperties();
    //void readExtendedProperties();
  };
}
