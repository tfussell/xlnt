cmake_minimum_required(VERSION 3.2)
project(zlib)

set(CMAKE_CXX_STANDARD 14)
set(CMAKE_CXX_STANDARD_REQUIRED ON)

include(ExternalProject)
ExternalProject_Add(zlib-download
    GIT_REPOSITORY https://github.com/madler/zlib
    INSTALL_COMMAND "")

set(ZLIB_ROOT_DIR ${CMAKE_CURRENT_BINARY_DIR}/zlib-download-prefix/src/zlib-download)
set(ZLIB_INCLUDE_DIR ${ZLIB_ROOT_DIR})
set(ZLIB_BUILD_DIR ${CMAKE_CURRENT_BINARY_DIR}/zlib-download-prefix/src/zlib-download-build)

add_library(zlib STATIC IMPORTED)
set_property(TARGET zlib PROPERTY IMPORTED_LOCATION ${ZLIB_BUILD_DIR}/libz.a)
add_dependencies(zlib zlib-download)
