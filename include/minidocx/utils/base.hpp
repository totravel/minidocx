/**
 * Copyright (C) 2022-2025, Xie Zequn <totravel@foxmail.com>. All rights reserved.
 * Distributed under the MIT License (http://opensource.org/licenses/MIT)
 * Report bugs and download new versions at https://github.com/totravel/minidocx
 */

#pragma once

#include "config.hpp"


namespace MINIDOCX_NAMESPACE
{
  class MINIDOCX_API Destroyable
  {
  private:
    bool destroyed_{ false };

  protected:
    virtual void clear() {}

  public:
    inline bool destroyed() const
    {
      return destroyed_;
    }

    virtual void destroy()
    {
      clear();
      destroyed_ = true;
    }
  };

  template<class T>
  class MINIDOCX_API Node : public Destroyable
  {
  public:
    Node(const T type) : type_{ type } {}
    virtual ~Node() = default;

    inline T type() const
    {
      return type_;
    };

  private:
    const T type_;
  };

}
