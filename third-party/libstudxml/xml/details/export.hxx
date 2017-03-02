// file      : xml/details/export.hxx
// copyright : Copyright (c) 2013-2017 Code Synthesis Tools CC
// license   : MIT; see accompanying LICENSE file

#ifndef XML_DETAILS_EXPORT_HXX
#define XML_DETAILS_EXPORT_HXX

#include <xml/details/config.hxx>

// Normally we don't export class templates (but do complete specializations),
// inline functions, and classes with only inline member functions. Exporting
// classes that inherit from non-exported/imported bases (e.g., std::string)
// will end up badly. The only known workarounds are to not inherit or to not
// export. Also, MinGW GCC doesn't like seeing non-exported function being
// used before their inline definition. The workaround is to reorder code. In
// the end it's all trial and error.

#ifdef LIBSTUDXML_BUILD2

#if defined(LIBSTUDXML_STATIC)         // Using static.
#  define LIBSTUDXML_EXPORT
#elif defined(LIBSTUDXML_STATIC_BUILD) // Building static.
#  define LIBSTUDXML_EXPORT
#elif defined(LIBSTUDXML_SHARED)       // Using shared.
#  ifdef _WIN32
#    define LIBSTUDXML_EXPORT __declspec(dllimport)
#  else
#    define LIBSTUDXML_EXPORT
#  endif
#elif defined(LIBSTUDXML_SHARED_BUILD) // Building shared.
#  ifdef _WIN32
#    define LIBSTUDXML_EXPORT __declspec(dllexport)
#  else
#    define LIBSTUDXML_EXPORT
#  endif
#else
// If none of the above macros are defined, then we assume we are being used
// by some third-party build system that cannot/doesn't signal the library
// type. Note that this fallback works for both static and shared but in case
// of shared will be sub-optimal compared to having dllimport.
//
#  define LIBSTUDXML_EXPORT            // Using static or shared.
#endif

#else // LIBSTUDXML_BUILD2

#ifdef LIBSTUDXML_STATIC_LIB
#  define LIBSTUDXML_EXPORT
#else
#  ifdef _WIN32
#    ifdef _MSC_VER
#      ifdef LIBSTUDXML_DYNAMIC_LIB
#        define LIBSTUDXML_EXPORT __declspec(dllexport)
#      else
#        define LIBSTUDXML_EXPORT __declspec(dllimport)
#      endif
#    else
#      ifdef LIBSTUDXML_DYNAMIC_LIB
#        ifdef DLL_EXPORT
#          define LIBSTUDXML_EXPORT __declspec(dllexport)
#        else
#          define LIBSTUDXML_EXPORT
#        endif
#      else
#        define LIBSTUDXML_EXPORT __declspec(dllimport)
#      endif
#    endif
#  else
#    define LIBSTUDXML_EXPORT
#  endif
#endif

#endif // LIBSTUDXML_BUILD2

#endif // XML_DETAILS_EXPORT_HXX
