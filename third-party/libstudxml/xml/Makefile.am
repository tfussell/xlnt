# file      : xml/Makefile.am
# copyright : Copyright (c) 2013-2017 Code Synthesis Tools CC
# license   : MIT; see accompanying LICENSE file

EXTRA_DIST = __file__(extra_dist)

studxmlincludedir = $(includedir)/xml

lib_LTLIBRARIES = libstudxml.la
libstudxml_la_SOURCES = __path__(sources)

nobase_studxmlinclude_HEADERS = __path__(headers)
nobase_nodist_studxmlinclude_HEADERS = details/config.h

if !LIBSTUDXML_EXTERNAL_EXPAT
libstudxml_la_SOURCES += __path__(expat_sources)
nobase_studxmlinclude_HEADERS += __path__(expat_headers)
endif

libstudxml_la_SOURCES += __path__(genx_sources)
nobase_studxmlinclude_HEADERS += __path__(genx_headers)

nobase_studxmlinclude_HEADERS += __path__(xml_headers)

# Make sure make doesn't try to build the no-extension headers
# thinking that they are executable.
#
__foreach_w__(__f,__path__(xml_headers),__f $(top_srcdir)/xml/__f ): ; @:

AM_CPPFLAGS = -I$(top_builddir) -I$(top_srcdir) -DLIBSTUDXML_DYNAMIC_LIB
AM_LDFLAGS = -release __value__(interface_version) -no-undefined
