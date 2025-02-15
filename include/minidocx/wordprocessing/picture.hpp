/**
 * Copyright (C) 2022-2025, Xie Zequn <totravel@foxmail.com>. All rights reserved.
 * Distributed under the MIT License (http://opensource.org/licenses/MIT)
 * Report bugs and download new versions at https://github.com/totravel/minidocx
 */

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
