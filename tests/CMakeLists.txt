cmake_minimum_required(VERSION 3.1)
project(xlnt.test)

# Require C++11 compiler
set(CMAKE_CXX_STANDARD 11)
set(CMAKE_CXX_STANDARD_REQUIRED ON)

if(NOT COMBINED_PROJECT)
  add_subdirectory(${CMAKE_CURRENT_SOURCE_DIR}/../source ${CMAKE_CURRENT_BINARY_DIR}/source)
endif()

if(STATIC_CRT)
  include(${CMAKE_CURRENT_SOURCE_DIR}/../cmake/ucm.cmake)
  ucm_set_runtime(STATIC)
endif()

file(GLOB CELL_TESTS ${CMAKE_CURRENT_SOURCE_DIR}/cell/*.hpp)
file(GLOB PACKAGING_TESTS ${CMAKE_CURRENT_SOURCE_DIR}/packaging/*.hpp)
file(GLOB STYLES_TESTS ${CMAKE_CURRENT_SOURCE_DIR}/styles/*.hpp)
file(GLOB UTILS_TESTS ${CMAKE_CURRENT_SOURCE_DIR}/utils/*.hpp)
file(GLOB WORKBOOK_TESTS ${CMAKE_CURRENT_SOURCE_DIR}/workbook/*.hpp)
file(GLOB WORKSHEET_TESTS ${CMAKE_CURRENT_SOURCE_DIR}/worksheet/*.hpp)

set(TESTS
  ${CELL_TESTS}
  ${PACKAGING_TESTS}
  ${STYLES_TESTS}
  ${UTILS_TESTS}
  ${WORKBOOK_TESTS}
  ${WORKSHEET_TESTS})

file(GLOB HELPERS_HEADERS ${CMAKE_CURRENT_SOURCE_DIR}/helpers/*.hpp)
file(GLOB HELPERS_SOURCES ${CMAKE_CURRENT_SOURCE_DIR}/helpers/*.cpp)

set(HELPERS ${HELPERS_HEADERS} ${HELPERS_SOURCES})
set(RUNNER ${CMAKE_CURRENT_SOURCE_DIR}/runner.cpp)

if(COVERAGE)
  set(CMAKE_EXE_LINKER_FLAGS "${CMAKE_EXE_LINKER_FLAGS} -fprofile-arcs -ftest-coverage")
endif()

add_executable(xlnt.test ${RUNNER} ${TESTS} ${HELPERS} $<TARGET_OBJECTS:libstudxml>)
target_link_libraries(xlnt.test PRIVATE xlnt)
target_include_directories(xlnt.test
  PRIVATE ${CMAKE_CURRENT_SOURCE_DIR}
  PRIVATE ${CMAKE_CURRENT_SOURCE_DIR}/../source
  PRIVATE ${CMAKE_CURRENT_SOURCE_DIR}/../third-party/libstudxml)

set(XLNT_TEST_DATA_DIR ${CMAKE_CURRENT_SOURCE_DIR}/data)
target_compile_definitions(xlnt.test PRIVATE XLNT_TEST_DATA_DIR=${XLNT_TEST_DATA_DIR})

if(MSVC)
  # bigobj because there are so many headers in one source file
  set_target_properties(xlnt.test PROPERTIES COMPILE_FLAGS "/wd\"4068\" /bigobj")
endif()

source_group(helpers FILES ${HELPERS})
source_group(runner FILES ${RUNNER})
source_group(tests\\cell FILES ${CELL_TESTS})
source_group(tests\\packaging FILES ${PACKAGING_TESTS})
source_group(tests\\serialization FILES ${SERIALIZATION_TESTS})
source_group(tests\\styles FILES ${STYLES_TESTS})
source_group(tests\\utils FILES ${UTILS_TESTS})
source_group(tests\\workbook FILES ${WORKBOOK_TESTS})
source_group(tests\\worksheet FILES ${WORKSHEET_TESTS})

if(MSVC AND NOT STATIC)
  # copy xlnt DLL into xlnt.test directory
  add_custom_command(TARGET xlnt.test POST_BUILD
    COMMAND ${CMAKE_COMMAND} -E copy_if_different
    $<TARGET_FILE:xlnt>
    $<TARGET_FILE_DIR:xlnt.test>)
endif()
