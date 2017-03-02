/* file      : xml/details/build2/config.h
 * copyright : Copyright (c) 2013-2017 Code Synthesis Tools CC
 * license   : MIT; see accompanying LICENSE file
 */

/* Static configuration file for the build2 build. */

#ifndef XML_DETAILS_CONFIG_H
#define XML_DETAILS_CONFIG_H

/* Define LIBSTUDXML_BUILD2 for the installed case. */
#ifndef LIBSTUDXML_BUILD2
#  define LIBSTUDXML_BUILD2
#endif

#ifdef __FreeBSD__
#  include <sys/endian.h> /* BYTE_ORDER */
#else
#  if defined(_WIN32)
#    ifndef BYTE_ORDER
#      define BIG_ENDIAN    4321
#      define LITTLE_ENDIAN 1234
#      define BYTE_ORDER    LITTLE_ENDIAN
#    endif
#  else
#    include <sys/param.h>  /* BYTE_ORDER/__BYTE_ORDER */
#    ifndef BYTE_ORDER
#      ifdef __BYTE_ORDER
#        define BYTE_ORDER    __BYTE_ORDER
#        define BIG_ENDIAN    __BIG_ENDIAN
#        define LITTLE_ENDIAN __LITTLE_ENDIAN
#      else
#        error no BYTE_ORDER/__BYTE_ORDER define
#      endif
#    endif
#  endif
#endif

#if BYTE_ORDER == BIG_ENDIAN
#  define LIBSTUDXML_BYTEORDER 4321
#else
#  define LIBSTUDXML_BYTEORDER 1234
#endif

#endif /* XML_DETAILS_CONFIG_H */
