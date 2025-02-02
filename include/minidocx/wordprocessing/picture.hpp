
#pragma once

#include "wordprocessing/base.hpp"
#include "wordprocessing/properties/picture.hpp"
#include "packaging/relationship.hpp"


namespace NAMESPACE
{
  class Picture : public Run, public Configurable<PictureProperties>
  {
  public:
    Picture(const RelationshipId id) : Run(RunType::Picture), id_{ id } {}
    ~Picture() override = default;

    RelationshipId id_;
  };
}
