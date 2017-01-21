cmake_minimum_required(VERSION 3.2)
project(xlnt.third-party)
list(APPEND CMAKE_MODULE_PATH "${CMAKE_CURRENT_SOURCE_DIR}")

find_package(ZLIB REQUIRED)
find_package(CryptoPP REQUIRED)
find_package(CxxTest REQUIRED)

include(ExternalProject)
ExternalProject_Add(libstudxml-download
    GIT_REPOSITORY http://scm.codesynthesis.com/libstudxml/libstudxml.git
    PREFIX ${CMAKE_CURRENT_BINARY_DIR}/libstudxml
    UPDATE_COMMAND ""
    PATCH_COMMAND ""
    CONFIGURE_COMMAND ""
    BUILD_COMMAND ""
    INSTALL_COMMAND ""
    STEP_TARGETS build
    EXCLUDE_FROM_ALL TRUE)

set(LIBSTUDXML_ROOT_DIR ${CMAKE_CURRENT_BINARY_DIR}/libstudxml/src/libstudxml)

include_directories(${LIBSTUDXML_ROOT_DIR})

set(LIBSTUDXML 
    ${LIBSTUDXML_ROOT_DIR}/xml/parser.cxx
    ${LIBSTUDXML_ROOT_DIR}/xml/qname.cxx 
    ${LIBSTUDXML_ROOT_DIR}/xml/serializer.cxx
    ${LIBSTUDXML_ROOT_DIR}/xml/value-traits.cxx
    ${LIBSTUDXML_ROOT_DIR}/xml/content
    ${LIBSTUDXML_ROOT_DIR}/xml/exception
    ${LIBSTUDXML_ROOT_DIR}/xml/forward
    ${LIBSTUDXML_ROOT_DIR}/xml/parser
    ${LIBSTUDXML_ROOT_DIR}/xml/qname
    ${LIBSTUDXML_ROOT_DIR}/xml/serializer
    ${LIBSTUDXML_ROOT_DIR}/xml/value-traits)

set(EXPAT
    ${LIBSTUDXML_ROOT_DIR}/xml/details/expat/xmlparse.c
    ${LIBSTUDXML_ROOT_DIR}/xml/details/expat/xmlrole.c
    ${LIBSTUDXML_ROOT_DIR}/xml/details/expat/xmltok_impl.c
    ${LIBSTUDXML_ROOT_DIR}/xml/details/expat/xmltok_ns.c
    ${LIBSTUDXML_ROOT_DIR}/xml/details/expat/xmltok.c
    ${LIBSTUDXML_ROOT_DIR}/xml/details/expat/ascii.h
    ${LIBSTUDXML_ROOT_DIR}/xml/details/expat/asciitab.h
    ${LIBSTUDXML_ROOT_DIR}/xml/details/expat/config.h
    ${LIBSTUDXML_ROOT_DIR}/xml/details/expat/expat_external.h
    ${LIBSTUDXML_ROOT_DIR}/xml/details/expat/expat.h
    ${LIBSTUDXML_ROOT_DIR}/xml/details/expat/iasciitab.h
    ${LIBSTUDXML_ROOT_DIR}/xml/details/expat/internal.h
    ${LIBSTUDXML_ROOT_DIR}/xml/details/expat/latin1tab.h
    ${LIBSTUDXML_ROOT_DIR}/xml/details/expat/nametab.h
    ${LIBSTUDXML_ROOT_DIR}/xml/details/expat/utf8tab.h
    ${LIBSTUDXML_ROOT_DIR}/xml/details/expat/xmlrole.h
    ${LIBSTUDXML_ROOT_DIR}/xml/details/expat/xmltok_impl.h
    ${LIBSTUDXML_ROOT_DIR}/xml/details/expat/xmltok.h)

set(GENX
    ${LIBSTUDXML_ROOT_DIR}/xml/details/genx/char-props.c
    ${LIBSTUDXML_ROOT_DIR}/xml/details/genx/genx.c
    ${LIBSTUDXML_ROOT_DIR}/xml/details/genx/genx.h)

add_library(libstudxml STATIC ${LIBSTUDXML} ${EXPAT} ${GENX})
target_compile_definitions(libstudxml PRIVATE LIBSTUDXML_STATIC_LIB=1)
add_dependencies(libstudxml libstudxml-download)
