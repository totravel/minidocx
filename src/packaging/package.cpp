/**
 * Copyright (C) 2022-2025, Xie Zequn <totravel@foxmail.com>. All rights reserved.
 * Distributed under the MIT License (http://opensource.org/licenses/MIT)
 * Report bugs and download new versions at https://github.com/totravel/minidocx
 */

#include "packaging/package.hpp"
#include "utils/exceptions.hpp"

#include "pugixml.hpp"

#ifndef NDEBUG
#include <iostream>
#endif
#include <sstream>


namespace MINIDOCX_NAMESPACE
{
  void Package::init()
  {
    initContentTypes();
  }

  void Package::flush()
  {
    writeCoreProperties();
    writeExtendedProperties();
    writeContentTypes();
    writeRelationships();
    writeBufferedParts();
  }

  void Package::clear()
  {
    defaultContentTypes_.clear();
    overrideContentTypes_.clear();
    relationships_.clear();
    packageRelationships_.clear();
  }

  void Package::load()
  {
    readContentTypes();
    readPkgRelationships();
  }


  void Package::writePart(const PartName& name, pugi::xml_document& doc)
  {
#ifndef NDEBUG
    std::clog << "-----> " << name.generic_string() << std::endl;
    doc.save(std::clog, "  ", pugi::format_indent | pugi::format_no_declaration);
#endif
    pugi::xml_node decl = doc.prepend_child(pugi::node_declaration);
    decl.append_attribute("version") = "1.0";
    decl.append_attribute("encoding") = "UTF-8";
    decl.append_attribute("standalone") = "yes";

    std::stringstream ss;
    doc.save(ss, "", pugi::format_raw);
    addFileFromStream(name, ss);
  }


  void Package::writeBufferedParts()
  {
    for (auto& ref : buffered_) {
#ifndef NDEBUG
      std::clog << "-----> " << ref.first.generic_string() << std::endl;
#endif
      addFileFromMem(ref.first, ref.second.data(), ref.second.size());
    }
  }

  
  void Package::initContentTypes()
  {
    defaultContentTypes_["xml"] = "application/xml";
    defaultContentTypes_["rels"] = toPartMediaType(PartType::Relationships);
  }

  void Package::writeContentTypes()
  {
    pugi::xml_document doc;
    pugi::xml_node root = doc.append_child("Types");
    root.append_attribute("xmlns")
      .set_value(toPartRootNamespace(PartType::ContentType));

    for (auto& ref : defaultContentTypes_) {
      pugi::xml_node child = root.append_child("Default");
      child.append_attribute("Extension") = ref.first.c_str();
      child.append_attribute("ContentType") = ref.second.c_str();
    }

    for (auto& ref : overrideContentTypes_) {
      pugi::xml_node child = root.append_child("Override");
      child.append_attribute("PartName") = ref.first.generic_string().c_str();
      child.append_attribute("ContentType") = ref.second.c_str();
    }

    writePart(TYPES_PART, doc);
  }

  void Package::readContentTypes()
  {
    pugi::xml_document doc;
    if (!doc.load_string(extractFileToString(TYPES_PART).c_str()))
      throw Exception("Cannot load xml");

    pugi::xml_node root = doc.child("Types");
    if (!root)
      throw io_error(TYPES_PART, "Invalid part");

    for (pugi::xml_node el : root.children()) {
      pugi::xml_attribute typeAttr = el.attribute("ContentType");
      if (!typeAttr)
        continue;

      if (std::strcmp(el.name(), "Default") == 0) {
        pugi::xml_attribute extAttr = el.attribute("Extension");
        if (!extAttr)
          continue;
        defaultContentTypes_[extAttr.value()] = typeAttr.value();
      }
      else if (std::strcmp(el.name(), "Override") == 0) {
        pugi::xml_attribute nameAttr = el.attribute("PartName");
        if (!nameAttr)
          continue;
        overrideContentTypes_[nameAttr.value()] = typeAttr.value();
      }
      else {
        continue;
      }
    }
  }


  void Package::writeRelationships(Relationships& rels, const PartName& name)
  {
    if (rels.maxId_ == 0)
      return;

    pugi::xml_document doc;
    pugi::xml_node root = doc.append_child("Relationships");
    root.append_attribute("xmlns")
      .set_value(toPartRootNamespace(PartType::Relationships));

    for (auto& ref : rels.map_) {
      auto& rel = ref.second;

      pugi::xml_node child = root.append_child("Relationship");
      child.append_attribute("Id") = Relationship::stringifyId(rel.id_).c_str();
      child.append_attribute("Type") = toPartRelationshipType(rel.type_);
      child.append_attribute("Target") = rel.target_.generic_string().c_str();
      if (rel.targetMode_ == Relationship::TargetMode::External)
        child.append_attribute("TargetMode") = "External";
    }

    writePart(name, doc);
  }

  void Package::writeRelationships()
  {
    writeRelationships(packageRelationships_, RELS_PART);

    for (auto& ref : relationships_)
      writeRelationships(ref.second, toRelationshipsPartName(ref.first));
  }

  void Package::readRelationships(const PartName& name, Relationships& rels)
  {
    pugi::xml_document doc;
    if (!doc.load_string(extractFileToString(name).c_str()))
      throw Exception("Could not read relationships part");

    pugi::xml_node root = doc.child("Relationships");
    if (!root)
      throw io_error(name.string(), "Invalid relationships part");

    for (pugi::xml_node el : root.children()) {
      if (std::strcmp(el.name(), "Relationship") != 0)
        continue;

      pugi::xml_attribute idAttr = el.attribute("Id");
      if (!idAttr)
        continue;

      pugi::xml_attribute typeAttr = el.attribute("Type");
      if (!typeAttr)
        continue;

      pugi::xml_attribute targetAttr = el.attribute("Target");
      if (!targetAttr)
        continue;

      pugi::xml_attribute modeAttr = el.attribute("TargetMode");

      RelationshipId id = Relationship::parseId(idAttr.value());
      PartType type = fromRelationshipType(typeAttr.value());
      PartName target = targetAttr.value();
      Relationship::TargetMode mode = modeAttr
        ? Relationship::parseTargetMode(modeAttr.value())
        : Relationship::TargetMode::Internal;

      if (type == PartType::Unknown || mode == Relationship::TargetMode::Unknown)
        continue;

      rels.map_[id] = { id, type, target, mode };
      if (id > rels.maxId_)
        rels.maxId_ = id;
    }
  }


  void Package::writeCoreProperties()
  {
    registerOverrideContentType(corePart_, toPartMediaType(PartType::CoreProperties));
    addPkgRelationship(PartType::CoreProperties, corePart_, Relationship::TargetMode::Internal);

    pugi::xml_document doc;
    pugi::xml_node root = doc.append_child("cp:coreProperties");
    root.append_attribute("xmlns:cp")
      .set_value(toPartRootNamespace(PartType::CoreProperties));
    root.append_attribute("xmlns:dc")
      .set_value("http://purl.org/dc/elements/1.1/");

    if (prop_.title_.size() > 0)
      root.append_child("dc:title").append_child(pugi::node_pcdata).set_value(prop_.title_.c_str());

    if (prop_.subject_.size() > 0)
      root.append_child("dc:subject").append_child(pugi::node_pcdata).set_value(prop_.subject_.c_str());

    if (prop_.author_.size() > 0)
      root.append_child("dc:creator").append_child(pugi::node_pcdata).set_value(prop_.author_.c_str());

    if (prop_.lastModifiedBy_.size() > 0)
      root.append_child("cp:lastModifiedBy").append_child(pugi::node_pcdata).set_value(prop_.lastModifiedBy_.c_str());

    if (prop_.company_.size() > 0)
      root.append_child("cp:company").append_child(pugi::node_pcdata).set_value(prop_.company_.c_str());

    writePart(corePart_, doc);
  }

  void Package::writeExtendedProperties()
  {
    registerOverrideContentType(appPart_, toPartMediaType(PartType::ExtendedProperties));
    addPkgRelationship(PartType::ExtendedProperties, appPart_, Relationship::TargetMode::Internal);

    pugi::xml_document doc;
    pugi::xml_node root = doc.append_child("Properties");
    root.append_attribute("xmlns")
      .set_value(toPartRootNamespace(PartType::ExtendedProperties));

    root.append_child("Template").append_child(pugi::node_pcdata).set_value("Normal.dotm");
    root.append_child("Application").append_child(pugi::node_pcdata).set_value("minidocx");
    root.append_child("AppVersion").append_child(pugi::node_pcdata).set_value("10.0000"); // XX.YYYY

    writePart(appPart_, doc);
  }

}
