/**
 * Copyright (C) 2022-2025, Xie Zequn <totravel@foxmail.com>. All rights reserved.
 * Distributed under the MIT License (http://opensource.org/licenses/MIT)
 * Report bugs and download new versions at https://github.com/totravel/minidocx
 */

#include "packaging/part.hpp"


namespace MINIDOCX_NAMESPACE
{
  PartName toRelationshipsPartName(const PartName& src)
  {
    PartName rels(src.parent_path() / "_rels" / src.filename());
    rels += ".rels";
    return rels;
  }

  const char* toPartMediaType(const PartType type)
  {
    switch (type)
    {
    case PartType::Relationships:
      return "application/vnd.openxmlformats-package.relationships+xml";

    case PartType::Setting:
      return "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml";

    case PartType::FontTable:
      return "application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml";

    case PartType::Footer:
      return "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml";

    case PartType::Header:
      return "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml";

    case PartType::OfficeDocument:
      return "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";

    case PartType::Numbering:
      return "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml";

    case PartType::Style:
      return "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml";

    case PartType::CoreProperties:
      return "application/vnd.openxmlformats-package.core-properties+xml";

    case PartType::ExtendedProperties:
      return "application/vnd.openxmlformats-officedocument.extended-properties+xml";

    default:
      return nullptr;
    }
  }

  const char* toPartRelationshipType(const PartType type)
  {
    switch (type)
    {
    case PartType::Setting:
      return "http://purl.oclc.org/ooxml/officeDocument/relationships/settings";

    case PartType::FontTable:
      return "http://purl.oclc.org/ooxml/officeDocument/relationships/fontTable";

    case PartType::Footer:
      return "http://purl.oclc.org/ooxml/officeDocument/relationships/footer";

    case PartType::Header:
      return "http://purl.oclc.org/ooxml/officeDocument/relationships/header";

    case PartType::OfficeDocument:
      return "http://purl.oclc.org/ooxml/officeDocument/relationships/officeDocument";

    case PartType::Numbering:
      return "http://purl.oclc.org/ooxml/officeDocument/relationships/numbering";

    case PartType::Style:
      return "http://purl.oclc.org/ooxml/officeDocument/relationships/styles";

    case PartType::CoreProperties:
      return "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties";

    case PartType::ExtendedProperties:
      return "http://purl.oclc.org/ooxml/officeDocument/relationships/extendedProperties";

    case PartType::Image:
      return "http://purl.oclc.org/ooxml/officeDocument/relationships/image";

    case PartType::Thumbnail:
      return "http://purl.oclc.org/ooxml/officeDocument/relationships/metadata/thumbnail";

    default:
      return nullptr;
    }

  }

  const char* toPartRootNamespace(const PartType type)
  {
    switch (type)
    {
    case PartType::ContentType:
      return "http://schemas.openxmlformats.org/package/2006/content-types";

    case PartType::Relationships:
      return "http://schemas.openxmlformats.org/package/2006/relationships";

    case PartType::Setting:
    case PartType::FontTable:
    case PartType::Footer:
    case PartType::Header:
    case PartType::OfficeDocument:
    case PartType::Numbering:
    case PartType::Style:
      return "http://purl.oclc.org/ooxml/wordprocessingml/main";

    case PartType::CoreProperties:
      return "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";

    case PartType::ExtendedProperties:
      return "http://purl.oclc.org/ooxml/officeDocument/extendedProperties";

    default:
      return nullptr;
    }
  }

  PartType fromRelationshipType(const std::string& type)
  {
    if (
      type == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" ||
      type == "http://purl.oclc.org/ooxml/officeDocument/relationships/settings")
      return PartType::Setting;

    if (
      type == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable" ||
      type == "http://purl.oclc.org/ooxml/officeDocument/relationships/fontTable")
      return PartType::FontTable;

    if (
      type == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" ||
      type == "http://purl.oclc.org/ooxml/officeDocument/relationships/footer")
      return PartType::Footer;

    if (
      type == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" ||
      type == "http://purl.oclc.org/ooxml/officeDocument/relationships/header")
      return PartType::Header;

    if (
      type == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" ||
      type == "http://purl.oclc.org/ooxml/officeDocument/relationships/officeDocument")
      return PartType::OfficeDocument;

    if (
      type == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" ||
      type == "http://purl.oclc.org/ooxml/officeDocument/relationships/numbering")
      return PartType::Numbering;

    if (
      type == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" ||
      type == "http://purl.oclc.org/ooxml/officeDocument/relationships/styles")
      return PartType::Style;

    if (
      type == "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties")
      return PartType::CoreProperties;

    if (
      type == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" ||
      type == "http://purl.oclc.org/ooxml/officeDocument/relationships/extendedProperties")
      return PartType::ExtendedProperties;

    if (
      type == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" ||
      type == "http://purl.oclc.org/ooxml/officeDocument/relationships/image")
      return PartType::Image;

    if (
      type == "http://schemas.openxmlformats.org/package/2006/relationships/metadata/thumbnail" ||
      type == "http://purl.oclc.org/ooxml/officeDocument/relationships/metadata/thumbnail")
      return PartType::Thumbnail;

    return PartType::Unknown;
  }

}
