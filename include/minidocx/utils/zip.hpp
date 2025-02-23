/**
 * Copyright (C) 2024-2025, Xie Zequn <totravel@foxmail.com>. All rights reserved.
 * Distributed under the MIT License (http://opensource.org/licenses/MIT)
 * Report bugs and download new versions at https://github.com/totravel/zip
 */

#pragma once

#include <iosfwd>
#include <filesystem>
#include <string>
#include <vector>
#include <utility>
#include <chrono>


namespace MINIDOCX_NAMESPACE
{
  namespace fs = std::filesystem;

  class Zip
  {
  public:
    Zip();
    virtual ~Zip();

    Zip(const Zip& src) = delete;
    Zip& operator=(const Zip& rhs) = delete;

    Zip(Zip&& src) noexcept;
    Zip& operator=(Zip&& rhs) noexcept;


    enum class OpenMode { None, ReadOnly, Update, Create, Create64 };

    // Opens an archive with given mode.
    void open(const std::string& filename, const OpenMode openMode = OpenMode::ReadOnly);

    // Closes the archive and releases resources.
    void close() noexcept;

    // Returns true if the archive is in zip64 format.
    bool isZip64() const;


    // High-level helper methods


  public:
    // Extracts file or folder to disk.
    void extractTo(const fs::path& name, const fs::path& dst);

    // Adds file or folder from disk.
    void add(const fs::path& name, const fs::path& src);

    // Deletes files and/or folders from the archive.
    void deleteFiles(const std::vector<fs::path>& names);


    // Low-level methods


  public:
    // Returns the total number of entries in the archive.
    size_t countEntries() const;

    // List of all entries.
    std::vector<fs::path> listEntries();

    // Returns false if the entry cannot be found.
    bool hasEntry(const fs::path& name);

    // Returns the uncompressed size of entry in bytes.
    size_t entrySize(const fs::path& name);


  public:
    // Extracts an file into output stream and returns its modified time.
    std::chrono::system_clock::time_point
      extractFileToStream(const fs::path& name, std::ostream& dst);

    // Extracts an file into disk file and sets its modified time.
    void extractFileToDisk(const fs::path& name, const fs::path& dst);

    // Extracts an file and returns a string.
    std::string extractFileToString(const fs::path& name);


  public:
    // Adds an file from memory and records its modified time.
    void addFileFromMem(const fs::path& name, const void* buf, const size_t bufsize,
      const std::chrono::system_clock::time_point& modifiedTime = std::chrono::system_clock::now());

    // Adds an file from input stream and records its modified time.
    void addFileFromStream(const fs::path& name, std::istream& src,
      const std::chrono::system_clock::time_point& modified_time = std::chrono::system_clock::now());

    // Adds an file from disk file and records its modified time.
    void addFileFromDisk(const fs::path& name, const fs::path& src);

    // Adds an file from string.
    void addFileFromString(const fs::path& name, const std::string& data);

    // Creates an empty folder in the archive.
    void addFolder(const fs::path& name);


    inline void swap(Zip& other) noexcept
    {
      std::swap(zip_, other.zip_);
      std::swap(file_, other.file_);
      std::swap(filename_, other.filename_);
      std::swap(openMode_, other.openMode_);
    }

  private:
    struct mz_zip_archive_ex;

    mz_zip_archive_ex* zip_{ nullptr };
    std::fstream* file_{ nullptr };

  protected:
    std::string filename_;
    OpenMode openMode_{ OpenMode::None };
  };
}
