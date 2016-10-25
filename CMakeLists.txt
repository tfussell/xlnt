cmake_minimum_required(VERSION 3.1)

project(xlnt)

option(SHARED "Set to OFF to not build shared libraries" ON)
option(STATIC "Set to ON to build static libraries" OFF)
option(COVERAGE "Generate coverage data for use in Coveralls" OFF)

option(WITH_EXAMPLES "Build examples" OFF)
option(WITH_TESTS "Build tests" OFF)
option(WITH_BENCHMARKS "Build performance benchmarks" OFF)

if(APPLE)
  option(FRAMEWORK "Set to ON to package dylib and headers into a .framework, OSX only" OFF)
  execute_process(COMMAND "sw_vers -productVersion | awk -F'.' '{print $1\".\"$2}'"
                  OUTPUT_VARIABLE OSX_VERSION)
  set(CMAKE_OSX_DEPLOYMENT_TARGET ${OSX_VERSION})
endif()

if(COVERAGE)
  set(CMAKE_CXX_FLAGS "${CMAKE_CXX_FLAGS} -fprofile-arcs -ftest-coverage")
endif()

set(CMAKE_CXX_STANDARD 14)
set(CMAKE_CXX_STANDARD_REQUIRED ON)

if(MSVC)
    add_definitions(-DUNICODE -D_UNICODE)
endif()

set(CMAKE_ARCHIVE_OUTPUT_DIRECTORY ${CMAKE_BINARY_DIR}/lib)
set(CMAKE_LIBRARY_OUTPUT_DIRECTORY ${CMAKE_BINARY_DIR}/lib)
set(CMAKE_RUNTIME_OUTPUT_DIRECTORY ${CMAKE_BINARY_DIR}/bin)

foreach(OUTPUTCONFIG ${CMAKE_CONFIGURATION_TYPES})
    string(TOUPPER ${OUTPUTCONFIG} OUTPUTCONFIG)
    set(CMAKE_LIBRARY_OUTPUT_DIRECTORY_${OUTPUTCONFIG} ${CMAKE_BINARY_DIR}/lib)
    set(CMAKE_ARCHIVE_OUTPUT_DIRECTORY_${OUTPUTCONFIG} ${CMAKE_BINARY_DIR}/lib)
    set(CMAKE_RUNTIME_OUTPUT_DIRECTORY_${OUTPUTCONFIG} ${CMAKE_BINARY_DIR}/bin)
endforeach()

if(WITH_TESTS)
  include(cmake/xlnt.test.cmake)
endif()

include(cmake/xlnt.cmake)
