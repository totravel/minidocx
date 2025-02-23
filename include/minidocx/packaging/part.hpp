/**
 * Copyright (C) 2022-2025, Xie Zequn <totravel@foxmail.com>. All rights reserved.
 * Distributed under the MIT License (http://opensource.org/licenses/MIT)
 * Report bugs and download new versions at https://github.com/totravel/minidocx
 */

#pragma once

#include <string>
#include <filesystem>


// Default part names
#define TYPES_PART "/[Content_Types].xml"
#define RELS_PART  "/_rels/.rels"
#define CORE_PART  "/docProps/core.xml"
#define APP_PART   "/docProps/app.xml"
#define MAIN_PART  "/word/document.xml"
#define STYLE_PART "/word/styles.xml"
#define NUM_PART   "/word/numbering.xml"
#define IMG_PREFIX "/word/media/image"


namespace MINIDOCX_NAMESPACE
{
  namespace fs = std::filesystem;

  // Part extension without a leading dot.
  using PartExtension = std::string;

  // Absolute path of the part in package.
  using PartName = fs::path;

  PartName toRelationshipsPartName(const PartName& src);

  enum class PartType
  {
    Unknown,

    // OPC
    ContentType,
    Relationships,

    // WordprocessingML
    Setting,
    FontTable,
    Footer,
    Header,
    OfficeDocument,
    Numbering,
    Style,

    // Shared
    CoreProperties,
    ExtendedProperties,
    Image,
    Thumbnail
  };

  const char* toPartMediaType(const PartType type);
  const char* toPartRelationshipType(const PartType type);
  const char* toPartRootNamespace(const PartType type);

  PartType fromRelationshipType(const std::string& type);
}
