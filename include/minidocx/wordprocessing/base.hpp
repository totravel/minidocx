
#pragma once

#include "utils/base.hpp"


namespace NAMESPACE
{
  enum class BlockType { Paragraph, Table };
  using Block = Node<BlockType>;

  enum class RunType { RichText, Picture };
  using Run = Node<RunType>;

}
