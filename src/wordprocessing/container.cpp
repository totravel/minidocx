
#include "wordprocessing/container.hpp"
#include "wordprocessing/paragraph.hpp"
#include "wordprocessing/table.hpp"
#include "utils/exceptions.hpp"


namespace NAMESPACE
{
  ParagraphPointer Container::addParagraph()
  {
    auto block{ std::make_shared<Paragraph>() };
    blocks_.push_back(block);
    return block;
  }

  TablePointer Container::addTable(const size_t rows, const size_t cols)
  {
    auto block{ std::make_shared<Table>(rows, cols) };
    blocks_.push_back(block);
    return block;
  }

  void Container::deleteBlock(const BlockPointer& block)
  {
    block->destroy();
    blocks_.remove(block);
  }

  void Container::clear()
  {
    for (auto& block : blocks_)
      block->destroy();
    blocks_.clear();
  }
}
