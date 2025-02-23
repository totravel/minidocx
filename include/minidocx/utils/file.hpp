/**
 * Copyright (C) 2022-2025, Xie Zequn <totravel@foxmail.com>. All rights reserved.
 * Distributed under the MIT License (http://opensource.org/licenses/MIT)
 * Report bugs and download new versions at https://github.com/totravel/minidocx
 */

#pragma once

#include <filesystem>
#include <vector>


namespace MINIDOCX_NAMESPACE
{
  namespace fs = std::filesystem;

  using FileName = fs::path;

  enum class FileType
  {
    Unknown,
    JPG,
    PNG,
    GIF,
    SVG,
    WEBP
  };

  FileType getFileType(const FileName& ext);

  const char* getFileExtension(const FileType type);

  const char* getFileMediaType(const FileType type);

  using Byte = unsigned char;
  using Buffer = std::vector<Byte>;

  void readFile(Buffer& buf, const FileName& filename);

  void writeFile(const Buffer& buf, const FileName& filename);
}
