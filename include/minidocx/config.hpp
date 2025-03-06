/**
 * Copyright (C) 2022-2025, Xie Zequn <totravel@foxmail.com>. All rights reserved.
 * Distributed under the MIT License (http://opensource.org/licenses/MIT)
 * Report bugs and download new versions at https://github.com/totravel/minidocx
 */

#pragma once

#ifndef MINIDOCX_NAMESPACE
#define MINIDOCX_NAMESPACE md
#endif

#ifdef WIN32
# ifdef MINIDOCX_SHARED
#   ifdef MINIDOCX_EXPORTS
#     define MINIDOCX_API __declspec(dllexport)
#   else
#     define MINIDOCX_API __declspec(dllimport)
#   endif
# else
#  define MINIDOCX_API
# endif
#else
# define MINIDOCX_API
#endif
