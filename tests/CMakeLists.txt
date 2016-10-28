cmake_minimum_required(VERSION 3.1)
include(${CMAKE_CURRENT_SOURCE_DIR}/../cmake/common.cmake)
project(${LIBRARY_NAME}.test VERSION ${LIBRARY_VERSION} LANGUAGES CXX C)

if(NOT COMBINED_PROJECT)
    add_subdirectory(${LIBRARY_SOURCE_DIR} ${CMAKE_CURRENT_BINARY_DIR}/source)
endif()

include_directories(${LIBRARY_INCLUDE_DIR})
include_directories(${LIBRARY_SOURCE_DIR})
include_directories(${CMAKE_CURRENT_SOURCE_DIR})
include_directories(${THIRD_PARTY_DIR}/cxxtest)
include_directories(${THIRD_PARTY_DIR}/utfcpp/source)
include_directories(${THIRD_PARTY_DIR}/pugixml/src)

file(GLOB CELL_TESTS ${LIBRARY_SOURCE_DIR}/cell/tests/test_*.hpp)
file(GLOB CHARTS_TESTS ${LIBRARY_SOURCE_DIR}/charts/tests/test_*.hpp)
file(GLOB CHARTSHEET_TESTS ${LIBRARY_SOURCE_DIR}/chartsheet/tests/test_*.hpp)
file(GLOB DRAWING_TESTS ${LIBRARY_SOURCE_DIR}/drawing/tests/test_*.hpp)
file(GLOB FORMULA_TESTS ${LIBRARY_SOURCE_DIR}/formula/tests/test_*.hpp)
file(GLOB PACKAGING_TESTS ${LIBRARY_SOURCE_DIR}/packaging/tests/test_*.hpp)
file(GLOB STYLES_TESTS ${LIBRARY_SOURCE_DIR}/styles/tests/test_*.hpp)
file(GLOB UTILS_TESTS ${LIBRARY_SOURCE_DIR}/utils/tests/test_*.hpp)
file(GLOB WORKBOOK_TESTS ${LIBRARY_SOURCE_DIR}/workbook/tests/test_*.hpp)
file(GLOB WORKSHEET_TESTS ${LIBRARY_SOURCE_DIR}/worksheet/tests/test_*.hpp)

set(TESTS ${CELL_TESTS} ${CHARTS_TESTS} ${CHARTSHEET_TESTS} ${DRAWING_TESTS} 
    ${FORMULA_TESTS} ${PACKAGING_TESTS} ${STYLES_TESTS} ${UTILS_TESTS} 
	${WORKBOOK_TESTS} ${WORKSHEET_TESTS})
set(PUGIXML ${THIRD_PARTY_DIR}/pugixml/src/pugixml.cpp)

file(GLOB TEST_HELPERS_HEADERS ${XLNT_TESTS_DIR}/helpers/*.hpp)
file(GLOB TEST_HELPERS_SOURCES ${XLNT_TESTS_DIR}/helpers/*.cpp)

SET(TEST_HELPERS ${TEST_HELPERS_HEADERS} ${TEST_HELPERS_SOURCES})

file(MAKE_DIRECTORY "${CMAKE_CURRENT_BINARY_DIR}/tests")
set(RUNNER "${CMAKE_CURRENT_BINARY_DIR}/runner-autogen.cpp")
set_source_files_properties(${RUNNER} PROPERTIES GENERATED TRUE)

add_executable(${PROJECT_NAME} ${TEST_HELPERS} ${TESTS} ${RUNNER} ${PUGIXML})

source_group(helpers FILES ${TEST_HELPERS})
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

target_link_libraries(${PROJECT_NAME} ${LIBRARY_NAME})

if(MSVC)
    set_target_properties(${PROJECT_NAME} PROPERTIES COMPILE_FLAGS "/wd\"4251\" /wd\"4275\"")
    # Needed for PathFileExists in path_helper
    target_link_libraries(${PROJECT_NAME} Shlwapi)
else()
    set(THREADS_PREFER_PTHREAD_FLAG ON)
    find_package(Threads REQUIRED)
    target_link_libraries(${PROJECT_NAME} Threads::Threads)
endif()

add_custom_command(OUTPUT ${RUNNER}
    COMMAND python ${CMAKE_CURRENT_SOURCE_DIR}/../third-party/cxxtest/bin/cxxtestgen --runner=ErrorPrinter -o ${RUNNER} ${TESTS} 
    DEPENDS ${TESTS}
    COMMENT "Generating test runner ${RUNNER}")

add_custom_target(generate-test-runner DEPENDS ${RUNNER})
add_dependencies(${PROJECT_NAME} generate-test-runner)

file(COPY ${CMAKE_CURRENT_SOURCE_DIR}/data DESTINATION ${CMAKE_BINARY_DIR}/bin)
