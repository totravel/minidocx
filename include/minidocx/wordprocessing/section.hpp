
#pragma once

#include "wordprocessing/container.hpp"
#include "wordprocessing/properties/section.hpp"


namespace NAMESPACE
{
  class Section : public Container, public Configurable<SectionProperties>
  {
  };
}
