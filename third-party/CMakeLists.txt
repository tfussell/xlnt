include(${CMAKE_CURRENT_SOURCE_DIR}/../cmake/common.cmake)
project(${LIBRARY_NAME}.third-party VERSION ${LIBRARY_VERSION} LANGUAGES CXX C)

# Includes
include_directories(libstudxml
    utfcpp/source
    pole
    botan
    zlib
    ${LIBRARY_SOURCE_DIR}/detail)

set(LIBSTUDXML 
    ${CMAKE_CURRENT_SOURCE_DIR}/libstudxml/xml/parser.cxx
    ${CMAKE_CURRENT_SOURCE_DIR}/libstudxml/xml/qname.cxx 
    ${CMAKE_CURRENT_SOURCE_DIR}/libstudxml/xml/serializer.cxx
    ${CMAKE_CURRENT_SOURCE_DIR}/libstudxml/xml/value-traits.cxx 
    ${CMAKE_CURRENT_SOURCE_DIR}/libstudxml/xml/details/expat/xmlparse.c
    ${CMAKE_CURRENT_SOURCE_DIR}/libstudxml/xml/details/expat/xmlrole.c 
    ${CMAKE_CURRENT_SOURCE_DIR}/libstudxml/xml/details/expat/xmltok_impl.c
    ${CMAKE_CURRENT_SOURCE_DIR}/libstudxml/xml/details/expat/xmltok_ns.c 
    ${CMAKE_CURRENT_SOURCE_DIR}/libstudxml/xml/details/expat/xmltok.c
    ${CMAKE_CURRENT_SOURCE_DIR}/libstudxml/xml/details/genx/char-props.c 
    ${CMAKE_CURRENT_SOURCE_DIR}/libstudxml/xml/details/genx/genx.c)
set(POLE pole/pole.cpp)
set(BOTAN ${CMAKE_CURRENT_SOURCE_DIR}/botan/botan_all.cpp)
set(ZLIB ${CMAKE_CURRENT_SOURCE_DIR}/zlib/adler32.c
    ${CMAKE_CURRENT_SOURCE_DIR}/zlib/compress.c
    ${CMAKE_CURRENT_SOURCE_DIR}/zlib/crc32.c
    ${CMAKE_CURRENT_SOURCE_DIR}/zlib/deflate.c
    ${CMAKE_CURRENT_SOURCE_DIR}/zlib/gzclose.c
    ${CMAKE_CURRENT_SOURCE_DIR}/zlib/gzlib.c
    ${CMAKE_CURRENT_SOURCE_DIR}/zlib/gzread.c
    ${CMAKE_CURRENT_SOURCE_DIR}/zlib/gzwrite.c
    ${CMAKE_CURRENT_SOURCE_DIR}/zlib/infback.c
    ${CMAKE_CURRENT_SOURCE_DIR}/zlib/inffast.c
    ${CMAKE_CURRENT_SOURCE_DIR}/zlib/inflate.c
    ${CMAKE_CURRENT_SOURCE_DIR}/zlib/inftrees.c
    ${CMAKE_CURRENT_SOURCE_DIR}/zlib/trees.c
    ${CMAKE_CURRENT_SOURCE_DIR}/zlib/uncompr.c
    ${CMAKE_CURRENT_SOURCE_DIR}/zlib/zutil.c)

if(MSVC)
    set_source_files_properties(${BOTAN} PROPERTIES COMPILE_FLAGS "/wd\"4244\"")
endif()

if (CMAKE_SIZEOF_VOID_P EQUAL 8)
    set(CPU "x64")
else()
    set(CPU "x86")
endif()

add_custom_command(OUTPUT ${BOTAN}
    COMMAND python configure.py --minimized-build --enable-modules=sha1,aes,filters,codec_filt,cbc,ecb,sha2_32,sha2_64 --without-sphinx --disable-shared --amalgamation --cpu=${CPU}
    WORKING_DIRECTORY ${CMAKE_CURRENT_SOURCE_DIR}/botan
    COMMENT "Generating botan amalgamation ${BOTAN}")

add_library(xlnt.third-party OBJECT ${LIBSTUDXML} ${POLE} ${BOTAN} ${ZLIB})
target_compile_definitions(xlnt.third-party PRIVATE LIBSTUDXML_STATIC_LIB=1)

if(NOT STATIC)
    set_target_properties(xlnt.third-party
        PROPERTIES POSITION_INDEPENDENT_CODE 1)
endif()

if(MSVC)
    target_compile_definitions(xlnt.third-party PRIVATE _CRT_SECURE_NO_WARNINGS=1 NOMINMAX)
    set_target_properties(xlnt.third-party PROPERTIES COMPILE_FLAGS "/MP")
endif()

source_group(botan FILES ${BOTAN})
source_group(libstudxml FILES ${LIBSTUDXML})
source_group(pole FILES ${LIBSTUDXML})
source_group(zlib FILES ${ZLIB})
