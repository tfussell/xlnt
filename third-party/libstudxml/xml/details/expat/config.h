#ifndef EXPAT_CONFIG_H
#define EXPAT_CONFIG_H

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

#define BYTEORDER LIBSTUDXML_BYTEORDER

#define XML_NS 1
#define XML_DTD 1
#define XML_CONTEXT_BYTES 1024

#define UNUSED(x) (void)x;

#ifdef _WIN32
/* Windows
 *
 */
#define WIN32_LEAN_AND_MEAN
#include <windows.h>
#undef WIN32_LEAN_AND_MEAN

#define HAVE_MEMMOVE 1

#else
/* POSIX
 *
 */
#define HAVE_MEMMOVE 1
#endif

#endif /* EXPAT_CONFIG_H */
