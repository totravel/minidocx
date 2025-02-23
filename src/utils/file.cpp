/**
 * Copyright (C) 2022-2025, Xie Zequn <totravel@foxmail.com>. All rights reserved.
 * Distributed under the MIT License (http://opensource.org/licenses/MIT)
 * Report bugs and download new versions at https://github.com/totravel/minidocx
 */

#include "utils/file.hpp"
#include "utils/exceptions.hpp"

#include <fstream>


namespace MINIDOCX_NAMESPACE
{
  FileType getFileType(const FileName& filename)
  {
    const std::string ext = filename.extension().string();

    if (ext == ".jpg" || ext == ".jpeg")
      return FileType::JPG;

    if (ext == ".png")
      return FileType::PNG;

    if (ext == ".gif")
      return FileType::GIF;

    if (ext == ".svg")
      return FileType::SVG;

    if (ext == ".webp")
      return FileType::WEBP;

    return FileType::Unknown;
  }

  const char* getFileExtension(const FileType type)
  {
    switch (type)
    {
    case FileType::JPG:
      return ".jpg";

    case FileType::PNG:
      return ".png";

    case FileType::GIF:
      return ".gif";

    case FileType::SVG:
      return ".svg";

    case FileType::WEBP:
      return ".webp";

    default:
      return nullptr;
    }
  }

  const char* getFileMediaType(const FileType type)
  {
    switch (type)
    {
    case FileType::JPG:
      return "image/jpeg";

    case FileType::PNG:
      return "image/png";

    case FileType::GIF:
      return "image/gif";

    case FileType::SVG:
      return "image/svg+xml";

    case FileType::WEBP:
      return "image/webp";

    default:
      return nullptr;
    }
  }

  void readFile(Buffer& buf, const FileName& filename)
  {
    std::ifstream fin;
    fin.open(filename.string(), std::ios::binary | std::ios::in | std::ios::ate);
    if (fin.fail())
      throw io_error(filename.string(), "cannot open file");

    const std::streampos fileSize = fin.tellg();
    if (fin.fail())
      throw io_error(filename.string(), "cannot seek file");
    fin.seekg(0, std::ios::beg);

    buf.resize(fileSize);
    fin.read(reinterpret_cast<std::ifstream::char_type*>(buf.data()), fileSize);
    if (fin.fail())
      throw io_error(filename.string(), "cannot read file");
  }

  void writeFile(const Buffer& buf, const FileName& filename)
  {
    std::ofstream fout;
    fout.open(filename, std::ios::binary | std::ios::out);
    if (fout.fail())
      throw io_error(filename.string(), "cannot open file");

    fout.write(reinterpret_cast<const std::ofstream::char_type*>(buf.data()), buf.size());
    if (fout.fail())
      throw io_error(filename.string(), "cannot write file");
  }

}
