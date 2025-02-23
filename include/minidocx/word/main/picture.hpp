/**
 * Copyright (C) 2022-2025, Xie Zequn <totravel@foxmail.com>. All rights reserved.
 * Distributed under the MIT License (http://opensource.org/licenses/MIT)
 * Report bugs and download new versions at https://github.com/totravel/minidocx
 */

#pragma once

#include "word/main/base.hpp"
#include "word/main/properties/picture.hpp"
#include "packaging/relationship.hpp"


namespace MINIDOCX_NAMESPACE
{
  class MINIDOCX_API Picture : public Run
  {
  public:
    Picture(const RelationshipId id) : Run(RunType::Picture), id_{ id } {}
    ~Picture() override = default;

    PictureProperties prop_;
    RelationshipId id_;
  };
}
