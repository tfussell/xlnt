project(xlnt.test)

include_directories(include)
include_directories(source)
include_directories(tests)
include_directories(third-party/cxxtest)
include_directories(third-party/pugixml/src)

FILE(GLOB CELL_TESTS source/cell/tests/test_*.hpp)
FILE(GLOB CHARTS_TESTS source/charts/tests/test_*.hpp)
FILE(GLOB CHARTSHEET_TESTS source/chartsheet/tests/test_*.hpp)
FILE(GLOB DRAWING_TESTS source/drawing/tests/test_*.hpp)
FILE(GLOB FORMULA_TESTS source/formula/tests/test_*.hpp)
FILE(GLOB PACKAGING_TESTS source/packaging/tests/test_*.hpp)
FILE(GLOB SERIALIZATION_TESTS source/serialization/tests/test_*.hpp)
FILE(GLOB STYLES_TESTS source/styles/tests/test_*.hpp)
FILE(GLOB UTILS_TESTS source/utils/tests/test_*.hpp)
FILE(GLOB WORKBOOK_TESTS source/workbook/tests/test_*.hpp)
FILE(GLOB WORKSHEET_TESTS source/worksheet/tests/test_*.hpp)

SET(TESTS ${CELL_TESTS} ${CHARTS_TESTS} ${CHARTSHEET_TESTS} ${DRAWING_TESTS} ${FORMULA_TESTS} ${PACKAGING_TESTS} ${SERIALIZATION_TESTS} ${STYLES_TESTS} ${UTILS_TESTS} ${WORKBOOK_TESTS} ${WORKSHEET_TESTS})
SET(PUGIXML ../third-party/pugixml/src/pugixml.cpp)

FILE(GLOB TEST_HELPERS_HEADERS tests/helpers/*.hpp)
FILE(GLOB TEST_HELPERS_SOURCES tests/helpers/*.cpp)

SET(TEST_HELPERS ${TEST_HELPERS_HEADERS} ${TEST_HELPERS_SOURCES})

file(MAKE_DIRECTORY "${CMAKE_CURRENT_BINARY_DIR}/tests")
SET(RUNNER "${CMAKE_CURRENT_BINARY_DIR}/tests/runner-autogen.cpp")
SET_SOURCE_FILES_PROPERTIES(${RUNNER} PROPERTIES GENERATED TRUE)

add_executable(xlnt.test ${TEST_HELPERS} ${TESTS} ${RUNNER} ${PUGIXML})

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

if(SHARED)
    target_link_libraries(xlnt.test xlnt.shared)
else()
    target_link_libraries(xlnt.test xlnt.static)
endif()

if(MSVC)
    set_target_properties(xlnt.test PROPERTIES COMPILE_FLAGS "/wd\"4251\" /wd\"4275\"")
endif()

# Needed for PathFileExists in path_helper (test helper)
if(${CMAKE_CXX_COMPILER_ID} STREQUAL "MSVC")
    target_link_libraries(xlnt.test Shlwapi)
endif()

add_custom_command(OUTPUT ${RUNNER}
    COMMAND ${CMAKE_CURRENT_SOURCE_DIR}/cmake/generate-tests ${CMAKE_CURRENT_BINARY_DIR}
    DEPENDS ${TESTS}
    COMMENT "Generating test runner ${RUNNER}"
)

add_custom_target(generate-test-runner DEPENDS ${RUNNER})

add_dependencies(xlnt.test generate-test-runner)
