cmake_minimum_required(VERSION 3.1)
project(xlnt.test)

set(CMAKE_CXX_STANDARD 14)
set(CMAKE_CXX_STANDARD_REQUIRED ON)

if(NOT COMBINED_PROJECT)
    add_subdirectory(${LIBRARY_SOURCE_DIR} ${CMAKE_CURRENT_BINARY_DIR}/source)
endif()

file(GLOB CELL_TESTS ${CMAKE_CURRENT_SOURCE_DIR}/../source/cell/tests/test_*.hpp)
file(GLOB CHARTS_TESTS ${CMAKE_CURRENT_SOURCE_DIR}/../source/charts/tests/test_*.hpp)
file(GLOB CHARTSHEET_TESTS ${CMAKE_CURRENT_SOURCE_DIR}/../source/chartsheet/tests/test_*.hpp)
file(GLOB DRAWING_TESTS ${CMAKE_CURRENT_SOURCE_DIR}/../source/drawing/tests/test_*.hpp)
file(GLOB FORMULA_TESTS ${CMAKE_CURRENT_SOURCE_DIR}/../source/formula/tests/test_*.hpp)
file(GLOB PACKAGING_TESTS ${CMAKE_CURRENT_SOURCE_DIR}/../source/packaging/tests/test_*.hpp)
file(GLOB STYLES_TESTS ${CMAKE_CURRENT_SOURCE_DIR}/../source/styles/tests/test_*.hpp)
file(GLOB UTILS_TESTS ${CMAKE_CURRENT_SOURCE_DIR}/../source/utils/tests/test_*.hpp)
file(GLOB WORKBOOK_TESTS ${CMAKE_CURRENT_SOURCE_DIR}/../source/workbook/tests/test_*.hpp)
file(GLOB WORKSHEET_TESTS ${CMAKE_CURRENT_SOURCE_DIR}/../source/worksheet/tests/test_*.hpp)

set(XLNT_ZIP
    ${CMAKE_CURRENT_SOURCE_DIR}/../source/detail/zstream.cpp
    ${CMAKE_CURRENT_SOURCE_DIR}/../source//detail/zstream.hpp)

set(TESTS ${CELL_TESTS} ${CHARTS_TESTS} ${CHARTSHEET_TESTS} ${DRAWING_TESTS} 
    ${FORMULA_TESTS} ${PACKAGING_TESTS} ${STYLES_TESTS} ${UTILS_TESTS} 
	${WORKBOOK_TESTS} ${WORKSHEET_TESTS})

file(GLOB TEST_HELPERS_HEADERS ${CMAKE_CURRENT_SOURCE_DIR}/helpers/*.hpp)
file(GLOB TEST_HELPERS_SOURCES ${CMAKE_CURRENT_SOURCE_DIR}/helpers/*.cpp)

set(TEST_HELPERS ${TEST_HELPERS_HEADERS} ${TEST_HELPERS_SOURCES})

file(MAKE_DIRECTORY "${CMAKE_CURRENT_BINARY_DIR}/tests")
set(RUNNER "${CMAKE_CURRENT_BINARY_DIR}/runner-autogen.cpp")
set_source_files_properties(${RUNNER} PROPERTIES GENERATED TRUE)

if(COVERAGE)
    set(CMAKE_EXE_LINKER_FLAGS "${CMAKE_EXE_LINKER_FLAGS} -lgcov --coverage")
endif()

add_executable(xlnt.test ${TEST_HELPERS} ${TESTS} ${RUNNER} ${XLNT_ZIP})
target_link_libraries(xlnt.test xlnt)
target_include_directories(xlnt.test
    PRIVATE ${CMAKE_CURRENT_SOURCE_DIR}
    PRIVATE ${CMAKE_CURRENT_SOURCE_DIR}/../source
    PRIVATE ${CXXTEST_INCLUDE_DIR})
set(XLNT_TEST_DATA_DIR ${CMAKE_CURRENT_SOURCE_DIR}/data)
target_compile_definitions(xlnt.test PRIVATE XLNT_TEST_DATA_DIR=${XLNT_TEST_DATA_DIR})

source_group(helpers FILES ${TEST_HELPERS})
source_group(xlnt\\detail FILES ${XLNT_ZIP})
source_group(tests\\cell FILES ${CELL_TESTS})
source_group(tests\\charts FILES ${CHARTS_TESTS})
source_group(tests\\chartsheet FILES ${CHARTSHEET_TESTS})
source_group(tests\\drawing FILES ${DRAWING_TESTS})
source_group(tests\\formula FILES ${FORMULA_TESTS})
source_group(tests\\packaging FILES ${PACKAGING_TESTS})
source_group(tests\\serialization FILES ${SERIALIZATION_TESTS})
source_group(tests\\styles FILES ${STYLES_TESTS})
source_group(tests\\utils FILES ${UTILS_TESTS})
source_group(tests\\workbook FILES ${WORKBOOK_TESTS})
source_group(tests\\worksheet FILES ${WORKSHEET_TESTS})
source_group(runner FILES ${RUNNER})

if(MSVC)
    set_target_properties(xlnt.test PROPERTIES COMPILE_FLAGS "/wd\"4251\" /wd\"4275\" /wd\"4068\"")
else()
    set(THREADS_PREFER_PTHREAD_FLAG ON)
    find_package(Threads REQUIRED)
    target_link_libraries(xlnt.test Threads::Threads)
endif()

find_package(PythonInterp REQUIRED)
find_package(CxxTest REQUIRED)

add_custom_command(OUTPUT ${RUNNER}
    COMMAND ${CXXTEST_PYTHON_TESTGEN_EXECUTABLE} --runner=ErrorPrinter -o ${RUNNER} ${TESTS}
    DEPENDS ${TESTS}
    COMMENT "Generating test runner ${RUNNER}")

add_custom_target(generate-test-runner DEPENDS ${RUNNER})
add_dependencies(xlnt.test generate-test-runner)
