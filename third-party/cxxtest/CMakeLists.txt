cmake_minimum_required(VERSION 3.7)
project(cxxtest)

include(ExternalProject)
ExternalProject_Add(cxxtest-download
    GIT_REPOSITORY https://github.com/CxxTest/cxxtest
    INSTALL_COMMAND "")

set(CXXTEST_ROOT_DIR ${CMAKE_CURRENT_BINARY_DIR}/cxxtest-download-prefix/src/cxxtest-download)
set(CXXTEST_INCLUDE_DIR ${CXXTEST_ROOT_DIR})

