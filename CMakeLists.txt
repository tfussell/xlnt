# Top level cmake script
cmake_minimum_required(VERSION 3.1)

project(xlnt
        VERSION 0.9.0
        LANGUAGES CXX C)


# Library type
option(SHARED "Set to OFF to not build shared libraries" ON)
option(STATIC "Set to ON to build static libraries" OFF)


# Optional components
option(WITH_SAMPLES "Build executable code samples" OFF)
option(WITH_BENCHMARKS "Build performance benchmarks" OFF)


# Platform specific options
if(NOT MSVC)
    option(COVERAGE "Generate coverage data for use in Coveralls" OFF)
endif()

if(APPLE)
    option(FRAMEWORK "Set to ON to package dylib and headers into a .framework, OSX only" OFF)
endif()


# Global language settings
set(CMAKE_CXX_STANDARD 14)
set(CMAKE_CXX_STANDARD_REQUIRED ON)


# Platform specific settings
if(COVERAGE)
  set(CMAKE_CXX_FLAGS "${CMAKE_CXX_FLAGS} -fprofile-arcs -ftest-coverage")
endif()

if(APPLE)
    # Prevent a warning about deployment target not being set
    execute_process(COMMAND "sw_vers -productVersion | awk -F'.' '{print $1\".\"$2}'"
                    OUTPUT_VARIABLE OSX_VERSION)
    set(CMAKE_OSX_DEPLOYMENT_TARGET ${OSX_VERSION})
endif()

if(MSVC)
    add_definitions(-DUNICODE -D_UNICODE)
endif()


# Define project directories
set(XLNT_INCLUDE_DIR ${CMAKE_SOURCE_DIR}/include)
set(XLNT_SOURCE_DIR ${CMAKE_SOURCE_DIR}/source)
set(THIRD_PARTY_DIR ${CMAKE_SOURCE_DIR}/third-party)


# Tell CMake where to find .cmake files
set(CMAKE_MODULE_PATH ${CMAKE_SOURCE_DIR}/cmake ${CMAKE_MODULE_PATH})


# Add components according to selected options
if (WITH_SAMPLES)
    add_subdirectory(samples)
endif()

if (WITH_BENCHMARKS)
    add_subdirectory(benchmarks)
endif()

add_subdirectory(tests)
add_subdirectory(source)
