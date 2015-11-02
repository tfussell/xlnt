project(xlnt.test)

include_directories(../include)
include_directories(../source)
include_directories(../third-party/pugixml/src)
include_directories(../third-party/cxxtest)

FILE(GLOB TEST_HEADERS ../tests/*.hpp)
FILE(GLOB TEST_HELPERS_HEADERS ../tests/helpers/*.hpp)
FILE(GLOB TEST_HELPERS_SOURCES ../tests/helpers/*.cpp)

add_executable(xlnt.test ../tests/runner-autogen.cpp ${TEST_HEADERS} ${TEST_HELPERS_HEADERS} ${TEST_HELPERS_SOURCES})

source_group(runner FILES ../tests/runner-autogen.cpp)
source_group(tests FILES ${TEST_HEADERS})
source_group(helpers FILES ${TEST_HELPERS_HEADERS} ${TEST_HELPERS_SOURCES})

target_link_libraries(xlnt.test xlnt)

if(${CMAKE_CXX_COMPILER_ID} STREQUAL "MSVC")
    target_link_libraries(xlnt.test Shlwapi)
endif()

add_custom_target (generate-test-runner
    COMMAND ${CMAKE_CURRENT_SOURCE_DIR}/generate-tests
    COMMENT "Generating test runner tests/runner-autogen.cpp"
)

add_dependencies(xlnt.test generate-test-runner)

if(${CMAKE_GENERATOR} STREQUAL "Unix Makefiles")
  add_custom_command(
    TARGET xlnt.test
    POST_BUILD
    COMMAND ${CMAKE_CURRENT_SOURCE_DIR}/../bin/xlnt.test
    VERBATIM
  )
endif()
