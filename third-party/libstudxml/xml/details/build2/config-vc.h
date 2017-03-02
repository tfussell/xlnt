/* file      : xml/details/build2/config-vc.h
 * copyright : Copyright (c) 2013-2017 Code Synthesis Tools CC
 * license   : MIT; see accompanying LICENSE file
 */

/* Configuration file for Windows/VC++ for the build2 build. */

#ifndef XML_DETAILS_CONFIG_VC_H
#define XML_DETAILS_CONFIG_VC_H

/* Define LIBSTUDXML_BUILD2 for the installed case. */
#ifndef LIBSTUDXML_BUILD2
#  define LIBSTUDXML_BUILD2
#endif

// Always little-endian, at least on i686 and x86_64.
//
#define LIBSTUDXML_BYTEORDER 1234

#endif /* XML_DETAILS_CONFIG_VC_H */
