cmake_minimum_required(VERSION 2.8.7)

if(NOT DEFINED CMAKE_SUPPRESS_DEVELOPER_WARNINGS)
    set(CMAKE_SUPPRESS_DEVELOPER_WARNINGS 1 CACHE INTERNAL "No dev warnings")
endif()

project(xlnt)

option(SHARED "Set to OFF to not build shared libraries" ON)
option(STATIC "Set to ON to build static libraries" OFF)
option(DEBUG "Set to ON to for debug configuration" OFF)
option(EXAMPLES "Build examples" OFF)
option(TESTS "Build tests" OFF)
option(BENCHMARKS "Build performance benchmarks" OFF)
option(COVERAGE "Generate coverage data for use in Coveralls" OFF)

if(APPLE)
    option(FRAMEWORK "Set to ON to package dylib and headers into a .framework, OSX only" OFF)
endif()

if(COVERAGE)
    set(CMAKE_CXX_FLAGS "${CMAKE_CXX_FLAGS} -fprofile-arcs -ftest-coverage")
endif()

if(${CMAKE_GENERATOR} STREQUAL "Unix Makefiles")
    if(DEBUG)
        set(CMAKE_BUILD_TYPE "Debug")
    else()
        set(CMAKE_BUILD_TYPE "Release")
    endif()
endif()

if(APPLE)
    execute_process(COMMAND "sw_vers -productVersion | awk -F'.' '{print $1\".\"$2}'"
                    OUTPUT_VARIABLE OSX_VERSION
    )
    set(CMAKE_OSX_DEPLOYMENT_TARGET ${OSX_VERSION})
endif(APPLE)

if(CMAKE_VERSION VERSION_LESS 3.1)
    if(${CMAKE_CXX_COMPILER_ID} STREQUAL "Clang")
        if(CMAKE_CXX_COMPILER_VERSION VERSION_LESS 3.5)
            set(CMAKE_CXX_FLAGS "${CMAKE_CXX_FLAGS} -std=c++1y")
        else()
            set(CMAKE_CXX_FLAGS "${CMAKE_CXX_FLAGS} -std=c++14")
        endif()
    elseif(${CMAKE_CXX_COMPILER_ID} STREQUAL "GNU")
        set(CMAKE_CXX_FLAGS "${CMAKE_CXX_FLAGS} -std=c++14")
    endif()
else()
    set(CMAKE_CXX_STANDARD 14)
    set(CMAKE_CXX_STANDARD_REQUIRED ON)
endif()

if(${CMAKE_CXX_COMPILER_ID} STREQUAL "MSVC")
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
endforeach(OUTPUTCONFIG CMAKE_CONFIGURATION_TYPES)

if(TESTS)
    include(cmake/xlnt.test.cmake)
endif()

include(cmake/xlnt.cmake)
