cmake_minimum_required(VERSION 3.1)
project(xlnt_all)

# This indicates to CMakeLists in subdirectories that they are part of a larger project
set(COMBINED_PROJECT TRUE)

# Library type
option(STATIC "Set to ON to build xlnt as a static library instead of a shared library" OFF)

# c++ language standard to use
set(XLNT_VALID_LANGS 11 14 17)
set(XLNT_CXX_LANG "14" CACHE STRING "c++ language features to compile with")
# enumerate allowed values for cmake gui
set_property(CACHE XLNT_CXX_LANG PROPERTY STRINGS ${XLNT_VALID_LANGS})
# validate value is in XLNT_VALID_LANGS
list(FIND XLNT_VALID_LANGS ${XLNT_CXX_LANG} index)
if(index EQUAL -1)
    message(FATAL_ERROR "XLNT_CXX_LANG must be one of ${XLNT_VALID_LANGS}")
endif()


# Optional components
option(TESTS "Set to OFF to skip building test executable (in ./tests)" ON)
option(SAMPLES "Set to ON to build executable code samples (in ./samples)" OFF)
option(BENCHMARKS "Set to ON to build performance benchmarks (in ./benchmarks)" OFF)
option(PYTHON "Set to ON to build Arrow conversion functions (in ./python)" OFF)

# Platform specific options
if(MSVC)
  option(STATIC_CRT "Link with the static version of MSVCRT (/MD[d])" OFF)
else()
  option(COVERAGE "Generate coverage data using gcov" OFF)
endif()

# Add components according to selected options
if(SAMPLES)
  add_subdirectory(samples)
endif()

if(BENCHMARKS)
  add_subdirectory(benchmarks)
endif()

if(TESTS)
  add_subdirectory(tests)
endif()

if(PYTHON)
  add_subdirectory(python)
endif()

add_subdirectory(source)
