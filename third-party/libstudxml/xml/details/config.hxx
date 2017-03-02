// file      : xml/details/config.hxx
// copyright : Copyright (c) 2013-2017 Code Synthesis Tools CC
// license   : MIT; see accompanying LICENSE file

#ifndef XML_DETAILS_CONFIG_HXX
#define XML_DETAILS_CONFIG_HXX

// C++11 support.
//
#ifdef _MSC_VER
#  if _MSC_VER >= 1900
#    define STUDXML_CXX11_NOEXCEPT
#  endif
#else
#  if defined(__GXX_EXPERIMENTAL_CXX0X__) || __cplusplus >= 201103L
#    ifdef __clang__ // Pretends to be a really old __GNUC__ on some platforms.
#      define STUDXML_CXX11_NOEXCEPT
#    elif defined(__GNUC__)
#      if (__GNUC__ == 4 && __GNUC_MINOR__ >= 6) || __GNUC__ > 4
#        define STUDXML_CXX11_NOEXCEPT
#      endif
#    else
#      define STUDXML_CXX11_NOEXCEPT
#    endif
#  endif
#endif

#ifdef STUDXML_CXX11_NOEXCEPT
#  define STUDXML_NOTHROW_NOEXCEPT noexcept
#else
#  define STUDXML_NOTHROW_NOEXCEPT throw()
#endif

// Note: the same in expat/config.h
//
#ifdef LIBSTUDXML_BUILD2
#  ifdef _MSC_VER
#    include <xml/details/build2/config-vc.h>
#  else
#    include <xml/details/build2/config.h>
#  endif
#else
#  ifdef _MSC_VER
#    include <xml/details/config-vc.h>
#  else
#    include <xml/details/config.h>
#  endif
#endif

#endif // XML_DETAILS_CONFIG_HXX
