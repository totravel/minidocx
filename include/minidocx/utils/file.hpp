
#pragma once

#include <filesystem>
#include <vector>


namespace NAMESPACE
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
