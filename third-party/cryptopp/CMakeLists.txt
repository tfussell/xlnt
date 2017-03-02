cmake_minimum_required(VERSION 3.7)
project(cryptopp)

include(ExternalProject)
ExternalProject_Add(cryptopp-download
    GIT_REPOSITORY https://github.com/weidai11/cryptopp
    INSTALL_COMMAND "")

set(CRYPTOPP_ROOT_DIR ${CMAKE_CURRENT_BINARY_DIR}/cryptopp-download-prefix/src/cryptopp-download)
set(CRYPTOPP_INCLUDE_DIR ${CRYPTOPP_ROOT_DIR})
set(CRYPTOPP_BUILD_DIR ${CMAKE_CURRENT_BINARY_DIR}/cryptopp-download-prefix/src/cryptopp-download-build)

add_library(cryptopp STATIC IMPORTED)
set_property(TARGET cryptopp PROPERTY IMPORTED_LOCATION ${CRYPTOPP_BUILD_DIR}/libcryptopp.a)
add_dependencies(cryptopp cryptopp-download)

if(NOT STATIC)
    set_target_properties(cryptopp PROPERTIES POSITION_INDEPENDENT_CODE 1)
endif()
