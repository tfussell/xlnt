cmake_minimum_required(VERSION 3.2)
project(libstudxml)

set(CMAKE_CXX_STANDARD 14)
set(CMAKE_CXX_STANDARD_REQUIRED ON)

set(LIBSTUDXML_ROOT_DIR ${CMAKE_CURRENT_SOURCE_DIR})
set(LIBSTUDXML_INCLUDE_DIR ${LIBSTUDXML_ROOT_DIR})

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

set(GENX
    ${LIBSTUDXML_ROOT_DIR}/xml/details/genx/char-props.c
    ${LIBSTUDXML_ROOT_DIR}/xml/details/genx/genx.c
    ${LIBSTUDXML_ROOT_DIR}/xml/details/genx/genx.h)

find_package(EXPAT REQUIRED)

add_library(libstudxml STATIC ${LIBSTUDXML} ${GENX})
target_compile_definitions(libstudxml
    PUBLIC LIBSTUDXML_STATIC_LIB=1
    PUBLIC LIBSTUDXML_EXTERNAL_EXPAT=1)
target_include_directories(libstudxml
    PUBLIC ${LIBSTUDXML_ROOT_DIR}
    PUBLIC ${EXPAT_INCLUDE_DIRS})
target_link_libraries(libstudxml
    PRIVATE ${EXPAT_LIBRARIES})

if(NOT STATIC)
    set_target_properties(libstudxml PROPERTIES POSITION_INDEPENDENT_CODE 1)
else()
    target_compile_definitions(libstudxml PUBLIC XML_STATIC=1)
endif()
