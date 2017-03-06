cmake_minimum_required(VERSION 3.2)
project(cxxtest)

set(CMAKE_CXX_STANDARD 14)
set(CMAKE_CXX_STANDARD_REQUIRED ON)

include(ExternalProject)
ExternalProject_Add(cxxtest-download
    GIT_REPOSITORY https://github.com/CxxTest/cxxtest
    INSTALL_COMMAND "")

set(CXXTEST_ROOT_DIR ${CMAKE_CURRENT_BINARY_DIR}/cxxtest-download-prefix/src/cxxtest-download)
set(CXXTEST_INCLUDE_DIR ${CXXTEST_ROOT_DIR})

