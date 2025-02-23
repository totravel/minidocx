/**
 * Copyright (C) 2024-2025, Xie Zequn <totravel@foxmail.com>. All rights reserved.
 * Distributed under the MIT License (http://opensource.org/licenses/MIT)
 * Report bugs and download new versions at https://github.com/totravel/zip
 */

#include "utils/zip.hpp"
#include "utils/exceptions.hpp"

#include "miniz.h"

#include <fstream>
#include <algorithm>
#include <memory>
#include <sstream>


namespace MINIDOCX_NAMESPACE
{
  struct Zip::mz_zip_archive_ex : mz_zip_archive
  {
    mz_uint64 m_orig_archive_size{ 0 };
  };

  static size_t readStream(void* pOpaque, size_t file_ofs, void* pBuf, size_t n) noexcept
  {
    auto* in = static_cast<std::istream*>(pOpaque);
    in->seekg(file_ofs);
    if (in->good()) {
      in->read(static_cast<char*>(pBuf), n);
      if (in->good() || in->eof()) {
        return in->gcount();
      }
    }
    return 0;
  }

  static size_t writeStream(void* pOpaque, size_t file_ofs, const void* pBuf, size_t n) noexcept
  {
    auto* out = static_cast<std::ostream*>(pOpaque);
    out->seekp(file_ofs);
    if (out->good()) {
      out->write(static_cast<const char*>(pBuf), n);
      if (out->good()) {
        return n;
      }
    }
    return 0;
  }


  Zip::Zip()
    : zip_{ new mz_zip_archive_ex }, file_{ new std::fstream }
  {
    mz_zip_zero_struct(zip_);
    zip_->m_pIO_opaque = file_;
    zip_->m_pRead = readStream;
    zip_->m_pWrite = writeStream;
  }

  Zip::~Zip()
  {
    close();
    if (file_ != nullptr)
      delete file_;
    if (zip_ != nullptr)
      delete zip_;
  }

  Zip::Zip(Zip&& src) noexcept
  {
    swap(src);
  }

  Zip& Zip::operator=(Zip&& rhs) noexcept
  {
    if (this == &rhs)
      return *this;
    Zip tmp{ std::move(rhs) };
    swap(tmp);
    return *this;
  }


  void Zip::open(const std::string& filename, const OpenMode openMode)
  {
    if (openMode_ != OpenMode::None)
      return;

    if (filename.size() == 0)
      throw invalid_parameter();

    std::ios::openmode mode = std::ios::binary;
    switch (openMode) {
    case OpenMode::Create:
    case OpenMode::Create64:
      mode |= std::ios::out;
      break;

    case OpenMode::Update:
      mode |= std::ios::out;
      [[fallthrough]];

    case OpenMode::ReadOnly:
      mode |= std::ios::in | std::ios::ate;
      break;

    case OpenMode::None:
    default:
      throw invalid_parameter();
    }

    file_->open(filename, mode);
    if (file_->fail())
      throw io_error(filename, "cannot open archive");

    uint64_t size{ 0 };
    mz_uint flags{ 0 };
    switch (openMode) {
    case OpenMode::Update:
      flags |= MZ_ZIP_FLAG_DO_NOT_SORT_CENTRAL_DIRECTORY;
      [[fallthrough]];

    case OpenMode::ReadOnly:
      size = file_->tellg();
      if (file_->fail()) {
        file_->close();
        throw io_error(filename, "cannot obtain archive size");
      }

      if (!mz_zip_reader_init(zip_, size, flags)) {
        file_->close();
        throw Exception(mz_zip_get_error_string(mz_zip_get_last_error(zip_)), "miniz");
      }

      if (openMode == OpenMode::ReadOnly)
        break;

      if (!mz_zip_writer_init_from_reader(zip_, NULL)) {
        file_->close();
        throw Exception(mz_zip_get_error_string(mz_zip_get_last_error(zip_)), "miniz");
      }
      break;

    case OpenMode::Create64:
      flags |= MZ_ZIP_FLAG_WRITE_ZIP64;
      [[fallthrough]];

    case OpenMode::Create:
      if (!mz_zip_writer_init_v2(zip_, 0, flags)) {
        file_->close();
        throw Exception(mz_zip_get_error_string(mz_zip_get_last_error(zip_)), "miniz");
      }
    }

    openMode_ = openMode;
    filename_ = filename;
    zip_->m_orig_archive_size = size;
  }

  void Zip::close() noexcept
  {
    switch (openMode_) {
    case OpenMode::ReadOnly:
      mz_zip_reader_end(zip_);
      break;

    case OpenMode::Update:
    case OpenMode::Create:
    case OpenMode::Create64:
      mz_zip_writer_finalize_archive(zip_);
      mz_zip_writer_end(zip_);
      break;

    case OpenMode::None:
    default:
      return;
    }

    file_->close();
    if (zip_->m_archive_size < zip_->m_orig_archive_size)
      fs::resize_file(filename_, zip_->m_archive_size);

    openMode_ = OpenMode::None;
    zip_->m_orig_archive_size = 0;
  }

  bool Zip::isZip64() const
  {
    if (openMode_ != OpenMode::ReadOnly && openMode_ != OpenMode::Update)
      throw invalid_operation();

    return zip_->m_pState->m_zip64;
  }


  struct EntryPosition {
    uint64_t lf_ofs_, lf_len_;
    uint64_t cd_ofs_, cd_len_;
  };

  enum class EntryCommand { Keep, Delete, Move };

  static void iosmove(std::iostream* ios,
    const std::streampos dst, const std::streampos src, const std::streamsize len)
  {
    static const size_t pageSize{ 20 };
    const size_t numPage{ len / pageSize };

    auto buf{ std::make_unique<char[]>(numPage ? pageSize : len) };
    auto ptr = buf.get();

    std::streampos pptr = dst, gptr = src;
    for (size_t i = 0; i < numPage; i++) {
      gptr = ios->seekg(gptr).read(ptr, pageSize).tellg();
      pptr = ios->seekp(pptr).write(ptr, pageSize).tellp();
    }
    ios->seekg(gptr).read(ptr, len % pageSize);
    ios->seekp(pptr).write(ptr, len % pageSize);
  }

  void Zip::deleteFiles(const std::vector<fs::path>& names)
  {
    if (openMode_ != OpenMode::Update)
      throw invalid_operation();

    if (names.size() == 0)
      return;

    std::vector<std::string> selectedNames(names.size());
    for (uint32_t i{ 0 }; i < names.size(); i++) {
      const fs::path name = names[i].lexically_normal().relative_path();
      if (name.empty())
        throw invalid_parameter();
      selectedNames[i] = name.generic_string();
    }

    const uint32_t numEntries{ zip_->m_total_files };
    if (!numEntries)
      return;

    std::vector<std::string> entryNames(numEntries);
    std::vector<EntryPosition> entryPos(numEntries);
    for (uint32_t i{ 0 }; i < numEntries; i++) {
      mz_zip_archive_file_stat entryStat;
      if (!mz_zip_reader_file_stat(zip_, i, &entryStat))
        throw Exception(mz_zip_get_error_string(mz_zip_get_last_error(zip_)), "miniz");
      entryNames[i] = entryStat.m_filename;
      entryPos[i].lf_ofs_ = entryStat.m_local_header_ofs;
      entryPos[i].cd_ofs_ = entryStat.m_central_dir_ofs;
    }

    std::vector<uint32_t> indices(numEntries);
    for (uint32_t i{ 0 }; i < numEntries; i++)
      indices[i] = i;
    std::sort(indices.begin(), indices.end(),
      [&entryPos](const int pos1, const int pos2)
      {
        return entryPos[pos1].lf_ofs_ < entryPos[pos2].lf_ofs_;
      });

    std::vector<EntryCommand> entryCmd(numEntries);
    uint32_t numDel{ 0 };
    for (uint32_t i{ 0 }; i < numEntries; i++) {
      const std::string& entryName = entryNames[indices[i]];
      const auto it = std::find_if(selectedNames.begin(), selectedNames.end(),
        [entryName](const std::string& selectedName) {
          return entryName.starts_with(selectedName) &&
            (entryName.size() == selectedName.size() || entryName[selectedName.size()] == '/');
        });
      if (it != selectedNames.end()) {
        entryCmd[indices[i]] = EntryCommand::Delete;
        numDel++;
      }
      else {
        entryCmd[indices[i]] = numDel ? EntryCommand::Move : EntryCommand::Keep;
      }
    }

    if (!numDel)
      return;

    const uint32_t maxIndex{ numEntries - 1 };
    mz_zip_array* cdirPtr = &zip_->m_pState->m_central_dir;
    for (uint32_t i{ 0 }; i < maxIndex; i++) {
      entryPos[indices[i]].lf_len_ = entryPos[indices[i + 1]].lf_ofs_ - entryPos[indices[i]].lf_ofs_;
      entryPos[i].cd_len_ = entryPos[i + 1].cd_ofs_ - entryPos[i].cd_ofs_;
    }
    entryPos[indices[maxIndex]].lf_len_ = zip_->m_archive_size - entryPos[indices[maxIndex]].lf_ofs_;
    entryPos[maxIndex].cd_len_ = cdirPtr->m_size - entryPos[maxIndex].cd_ofs_;

    uint32_t index{ 0 };
    uint64_t wrPos{ 0 };
    uint64_t rdPos{ 0 };
    uint64_t delLen{ 0 };
    uint64_t movLen{ 0 };
    while (index < numEntries && entryCmd[indices[index]] == EntryCommand::Keep)
      wrPos += entryPos[indices[index++]].lf_len_;
    rdPos = wrPos;

    do {
      while (index < numEntries && entryCmd[indices[index]] == EntryCommand::Delete)
        rdPos += entryPos[indices[index++]].lf_len_;
      delLen = rdPos - wrPos;

      while (index < numEntries && entryCmd[indices[index]] == EntryCommand::Move) {
        movLen += entryPos[indices[index]].lf_len_;
        entryPos[indices[index]].lf_ofs_ -= delLen;

        MZ_WRITE_LE32(
          &MZ_ZIP_ARRAY_ELEMENT(cdirPtr, mz_uint8, entryPos[indices[index]].cd_ofs_) + MZ_ZIP_CDH_LOCAL_HEADER_OFS,
          entryPos[indices[index]].lf_ofs_);

        index++;
      }
      if (!movLen)
        break;

      iosmove(file_, wrPos, rdPos, movLen);
      wrPos += movLen;
      rdPos += movLen;
      movLen = 0;
    } while (index < numEntries);

    zip_->m_archive_size -= delLen;
    zip_->m_total_files -= numDel;

    mz_zip_array* cdirOfsPtr = &zip_->m_pState->m_central_dir_offsets;
    cdirOfsPtr->m_size = 0;
    index = wrPos = rdPos = delLen = 0;
    do {
      while (index < numEntries && entryCmd[index] == EntryCommand::Delete)
        rdPos += entryPos[index++].cd_len_;
      delLen = rdPos - wrPos;

      while (index < numEntries && entryCmd[index] != EntryCommand::Delete) {
        movLen += entryPos[index].cd_len_;
        entryPos[index].cd_ofs_ -= delLen;

        MZ_ZIP_ARRAY_ELEMENT(cdirOfsPtr, mz_uint32, cdirOfsPtr->m_size++) = entryPos[index].cd_ofs_;

        index++;
      }
      if (!movLen)
        break;

      if (delLen) {
        std::memmove(
          &MZ_ZIP_ARRAY_ELEMENT(cdirPtr, mz_uint8, wrPos),
          &MZ_ZIP_ARRAY_ELEMENT(cdirPtr, mz_uint8, rdPos),
          movLen);
      }
      wrPos += movLen;
      rdPos += movLen;
      movLen = 0;
    } while (index < numEntries);

    cdirPtr->m_size -= delLen;
  }


  size_t Zip::countEntries() const
  {
    if (openMode_ != OpenMode::ReadOnly && openMode_ != OpenMode::Update)
      throw invalid_operation();

    return zip_->m_total_files;
  }

  std::vector<fs::path> Zip::listEntries()
  {
    if (openMode_ != OpenMode::ReadOnly && openMode_ != OpenMode::Update)
      throw invalid_operation();

    std::vector<fs::path> names(zip_->m_total_files);

    for (uint32_t i{ 0 }; i < zip_->m_total_files; i++) {

      mz_zip_archive_file_stat stat;
      if (!mz_zip_reader_file_stat(zip_, i, &stat))
        throw Exception(mz_zip_get_error_string(mz_zip_get_last_error(zip_)), "miniz");

      names[i] = stat.m_filename;
    }

    return names;
  }

  bool Zip::hasEntry(const fs::path& name)
  {
    if (openMode_ != OpenMode::ReadOnly && openMode_ != OpenMode::Update)
      throw invalid_operation();

    const fs::path entryName = name.lexically_normal().relative_path();
    if (entryName.empty())
      throw invalid_parameter();

    mz_uint32 entryIndex;
    return mz_zip_reader_locate_file_v2(zip_, entryName.generic_string().c_str(), 0, 0, &entryIndex);
  }

  size_t Zip::entrySize(const fs::path& name)
  {
    if (openMode_ != OpenMode::ReadOnly && openMode_ != OpenMode::Update)
      throw invalid_operation();

    const fs::path entryName = name.lexically_normal().relative_path();
    if (entryName.empty())
      throw invalid_parameter();

    mz_uint32 entryIndex;
    if (!mz_zip_reader_locate_file_v2(zip_, entryName.generic_string().c_str(), 0, 0, &entryIndex))
      throw Exception(mz_zip_get_error_string(mz_zip_get_last_error(zip_)), "miniz");

    mz_zip_archive_file_stat entryStat;
    if (!mz_zip_reader_file_stat(zip_, entryIndex, &entryStat))
      throw Exception(mz_zip_get_error_string(mz_zip_get_last_error(zip_)), "miniz");

    return entryStat.m_is_directory ? 0 : entryStat.m_uncomp_size;
  }


  static inline bool isFilename(const fs::path& name)
  {
    return name.has_filename();
  }

  std::chrono::system_clock::time_point
    Zip::extractFileToStream(const fs::path& name, std::ostream& dst)
  {
    if (openMode_ != OpenMode::ReadOnly && openMode_ != OpenMode::Update)
      throw invalid_operation();

    const fs::path filename = name.lexically_normal().relative_path();
    if (filename.empty() || !isFilename(filename))
      throw invalid_parameter();

    mz_uint32 fileIndex;
    if (!mz_zip_reader_locate_file_v2(zip_, filename.generic_string().c_str(), 0, 0, &fileIndex))
      throw Exception(mz_zip_get_error_string(mz_zip_get_last_error(zip_)), "miniz");

    if (!mz_zip_reader_extract_to_callback(zip_, fileIndex, writeStream, &dst, 0))
      throw Exception(mz_zip_get_error_string(mz_zip_get_last_error(zip_)), "miniz");

    mz_zip_archive_file_stat fileStat;
    if (!mz_zip_reader_file_stat(zip_, fileIndex, &fileStat))
      throw Exception(mz_zip_get_error_string(mz_zip_get_last_error(zip_)), "miniz");

    return std::chrono::system_clock::from_time_t(fileStat.m_time);
  }

  void Zip::extractFileToDisk(const fs::path& name, const fs::path& dst)
  {
    std::ofstream fout;
    fout.open(dst, std::ios::binary);
    if (fout.fail())
      throw io_error(dst.string(), "cannot open file");

    const auto mtime = extractFileToStream(name, fout);

    fout.close();
    fs::last_write_time(dst, std::chrono::clock_cast<std::chrono::file_clock>(mtime));
  }

  std::string Zip::extractFileToString(const fs::path& name)
  {
    std::ostringstream sout;
    extractFileToStream(name, sout);
    return sout.str();
  }


  void Zip::addFileFromMem(const fs::path& name, const void* buf, const size_t bufsize,
    const std::chrono::system_clock::time_point& modifiedTime)
  {
    if (openMode_ != OpenMode::Create && openMode_ != OpenMode::Create64 && openMode_ != OpenMode::Update)
      throw invalid_operation();

    const fs::path filename = name.lexically_normal().relative_path();
    if (filename.empty() || !isFilename(filename))
      throw invalid_parameter();

    std::time_t mtime = std::chrono::system_clock::to_time_t(modifiedTime);

    if (!mz_zip_writer_add_mem_ex_v2(
      zip_, filename.generic_string().data(),
      buf, bufsize, NULL, 0, MZ_DEFAULT_COMPRESSION,
      0, 0, &mtime, NULL, 0, NULL, 0))
      throw Exception(mz_zip_get_error_string(mz_zip_get_last_error(zip_)), "miniz");
  }

  void Zip::addFileFromStream(const fs::path& name, std::istream& src,
    const std::chrono::system_clock::time_point& modifiedTime)
  {
    if (openMode_ != OpenMode::Create && openMode_ != OpenMode::Create64 && openMode_ != OpenMode::Update)
      throw invalid_operation();

    const fs::path filename = name.lexically_normal().relative_path();
    if (filename.empty() || !isFilename(filename))
      throw invalid_parameter();

    std::streampos maxSize = 0;
    src.seekg(0, std::ios::end);
    if (src.good()) {
      std::streampos end = src.tellg();
      if (src.good())
        maxSize = end;
    }
    if (!maxSize)
      throw io_error("istream", "cannot obtain associated input sequence size");

    const std::time_t mtime = std::chrono::system_clock::to_time_t(modifiedTime);

    if (!mz_zip_writer_add_read_buf_callback(
      zip_, filename.generic_string().data(),
      readStream, &src, maxSize, &mtime,
      NULL, 0, MZ_DEFAULT_COMPRESSION, NULL, 0, NULL, 0))
      throw Exception(mz_zip_get_error_string(mz_zip_get_last_error(zip_)), "miniz");
  }

  void Zip::addFileFromDisk(const fs::path& name, const fs::path& src)
  {
    std::ifstream fin;
    fin.open(src, std::ios::binary);
    if (fin.fail())
      throw io_error(src.string(), "cannot open file");

    addFileFromStream(name, fin,
      std::chrono::clock_cast<std::chrono::system_clock>(fs::last_write_time(src)));
  }

  void Zip::addFileFromString(const fs::path& name, const std::string& data)
  {
    if (!data.size())
      throw invalid_parameter();

    std::istringstream sin(data);
    addFileFromStream(name, sin);
  }

  void Zip::addFolder(const fs::path& name)
  {
    if (openMode_ != OpenMode::Create && openMode_ != OpenMode::Create64 && openMode_ != OpenMode::Update)
      throw invalid_operation();

    const fs::path dirname = name.lexically_normal().relative_path();
    if (dirname.empty() || isFilename(dirname))
      throw invalid_parameter();

    if (!mz_zip_writer_add_mem(zip_, dirname.generic_string().data(), NULL, 0, MZ_DEFAULT_COMPRESSION))
      throw Exception(mz_zip_get_error_string(mz_zip_get_last_error(zip_)), "miniz");
  }
}
