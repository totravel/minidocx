/**
 * Copyright (C) 2022-2025, Xie Zequn <totravel@foxmail.com>. All rights reserved.
 * Distributed under the MIT License (http://opensource.org/licenses/MIT)
 * Report bugs and download new versions at https://github.com/totravel/minidocx
 */

#pragma once

#include "config.hpp"
#include "packaging/package.hpp"
#include "word/styles.hpp"
#include "word/numbering.hpp"

#include <memory>
#include <string>
#include <list>
#include <map>


namespace MINIDOCX_NAMESPACE
{
  class Section;
  class Paragraph;
  class Table;
  class Picture;

  using SectionPointer = std::shared_ptr<Section>;

  class MINIDOCX_API Document : public Package
  {
  public:
    Document();
    void saveAs(const std::string& filename);

    void load(const std::string& filename);
    inline void save() { saveAs(filename_); }

    inline void reset() { clear(); init(); }

  private:
    void init();
    void flush();
    void clear();
    void load();

  private:
    PartName mainPart_{ MAIN_PART };

    void initContentTypes();
    void initRelationships();

    void writeOfficeDocument();
    //void readOfficeDocument();

  private:
    std::list<SectionPointer> sections_;

  public:
    inline std::list<SectionPointer> sections() const { return sections_; }
    SectionPointer addSection();
    void deleteSection(const SectionPointer& section);
    void clearSections();

  private:
    size_t imgCount_ = 0;

  public:
    RelationshipId addImage(Buffer buf, const FileType type);
    RelationshipId addImage(const FileName& filename);

  private:
    PartName stylePart_{ STYLE_PART };

    std::map<std::string, ParagraphStyle> paragraphStyles_;
    std::map<std::string, CharacterStyle> characterStyles_;

    void writeStyles();

  public:
    void addParagraphStyle(const ParagraphStyle& style);
    void addCharacterStyle(const CharacterStyle& style);

  private:
    PartName numPart_{ NUM_PART };

    NumberingId nextAbstractNumId_ = 0;
    NumberingId nextNumId_ = 1;
    std::map<NumberingId, AbstractNumberingDefinition> abstractNumDefinitions_;
    std::map<NumberingId, NumberingDefinition> numDefinitions_;

    void writeNumDefinitions();

  public:
    // Adds abstract numbering definition.
    NumberingId addAbstractNumDefinition(AbstractNumberingDefinition def)
    {
      const NumberingId numId = nextAbstractNumId_++;
      abstractNumDefinitions_[numId] = std::move(def);
      return numId;
    }

    // Adds numbering definition instance.
    NumberingId addNumDefinition(NumberingDefinition def)
    {
      const NumberingId numId = nextNumId_++;
      numDefinitions_[numId] = std::move(def);
      return numId;
    }

    // Adds numbering definition instance for numbered list.
    inline NumberingId addNumberedListDefinition()
    {
      return addNumDefinition(
        addAbstractNumDefinition(AbstractNumberingDefinition::makeNumberedList()));
    }

    // Adds numbering definition instance for bulleted list.
    inline NumberingId addBulletedListDefinition()
    {
      return addNumDefinition(
        addAbstractNumDefinition(AbstractNumberingDefinition::makeBulletedList()));
    }
  };
}
