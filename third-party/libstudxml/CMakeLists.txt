cmake_minimum_required(VERSION 3.1)
project(libstudxml)

# Require C++11 compiler
set(CMAKE_CXX_STANDARD 11)
set(CMAKE_CXX_STANDARD_REQUIRED ON)

set(LIBSTUDXML_ROOT_DIR ${CMAKE_CURRENT_SOURCE_DIR})
set(LIBSTUDXML_INCLUDE_DIR ${LIBSTUDXML_ROOT_DIR})

if(STATIC_CRT)
  include(${CMAKE_CURRENT_SOURCE_DIR}/../../cmake/ucm.cmake)
  ucm_set_runtime(STATIC)
endif()

set(LIBSTUDXML
  ${LIBSTUDXML_ROOT_DIR}/libstudxml/parser.cxx
  ${LIBSTUDXML_ROOT_DIR}/libstudxml/qname.cxx
  ${LIBSTUDXML_ROOT_DIR}/libstudxml/serializer.cxx
  ${LIBSTUDXML_ROOT_DIR}/libstudxml/value-traits.cxx
  ${LIBSTUDXML_ROOT_DIR}/libstudxml/content
  ${LIBSTUDXML_ROOT_DIR}/libstudxml/exception
  ${LIBSTUDXML_ROOT_DIR}/libstudxml/forward
  ${LIBSTUDXML_ROOT_DIR}/libstudxml/parser
  ${LIBSTUDXML_ROOT_DIR}/libstudxml/qname
  ${LIBSTUDXML_ROOT_DIR}/libstudxml/serializer
  ${LIBSTUDXML_ROOT_DIR}/libstudxml/value-traits)

set(GENX
  ${LIBSTUDXML_ROOT_DIR}/libstudxml/details/genx/char-props.c
  ${LIBSTUDXML_ROOT_DIR}/libstudxml/details/genx/genx.c
  ${LIBSTUDXML_ROOT_DIR}/libstudxml/details/genx/genx.h)

set(EXPAT
  ${LIBSTUDXML_ROOT_DIR}/libstudxml/details/expat/xmlparse.c
  ${LIBSTUDXML_ROOT_DIR}/libstudxml/details/expat/xmlrole.c
  ${LIBSTUDXML_ROOT_DIR}/libstudxml/details/expat/xmltok.c
  ${LIBSTUDXML_ROOT_DIR}/libstudxml/details/expat/ascii.h
  ${LIBSTUDXML_ROOT_DIR}/libstudxml/details/expat/asciitab.h
  ${LIBSTUDXML_ROOT_DIR}/libstudxml/details/expat/config.h
  ${LIBSTUDXML_ROOT_DIR}/libstudxml/details/expat/expat_external.h
  ${LIBSTUDXML_ROOT_DIR}/libstudxml/details/expat/expat.h
  ${LIBSTUDXML_ROOT_DIR}/libstudxml/details/expat/iasciitab.h
  ${LIBSTUDXML_ROOT_DIR}/libstudxml/details/expat/internal.h
  ${LIBSTUDXML_ROOT_DIR}/libstudxml/details/expat/latin1tab.h
  ${LIBSTUDXML_ROOT_DIR}/libstudxml/details/expat/nametab.h
  ${LIBSTUDXML_ROOT_DIR}/libstudxml/details/expat/utf8tab.h
  ${LIBSTUDXML_ROOT_DIR}/libstudxml/details/expat/xmlrole.h
  ${LIBSTUDXML_ROOT_DIR}/libstudxml/details/expat/xmltok_impl.h
  ${LIBSTUDXML_ROOT_DIR}/libstudxml/details/expat/xmltok.h)

add_library(libstudxml OBJECT ${LIBSTUDXML} ${GENX} ${EXPAT})

target_compile_definitions(libstudxml PUBLIC LIBSTUDXML_STATIC_LIB=1)
target_include_directories(libstudxml
  PUBLIC ${LIBSTUDXML_ROOT_DIR}
  PUBLIC ${EXPAT_INCLUDE_DIRS})

if(STATIC)
  target_compile_definitions(libstudxml PUBLIC XML_STATIC=1)
endif()

# Prevent warning C4996 caused by strcpy, strncpy, sprintf in genx
# TODO: would it be better to define this only in genx.c?
if(MSVC)
  target_compile_definitions(libstudxml PRIVATE _CRT_SECURE_NO_WARNINGS=1)
endif()

# Build with -fPIC when xlnt is a shared library
# TODO: is this still necessary? try removing
if(NOT STATIC)
  set_target_properties(libstudxml PROPERTIES POSITION_INDEPENDENT_CODE 1)
endif()

# When xlnt is a static library, assume expat will be linked statically too
# TODO: is this a valid assumption?
if(STATIC)
  target_compile_definitions(libstudxml PUBLIC XML_STATIC=1)
endif()
