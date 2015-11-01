project(xlnt.test)

include_directories(../../../include)
include_directories(../../../source)
include_directories(../../../third-party/pugixml/src)
include_directories(../../../third-party/cxxtest)

add_executable(xlnt.test ../../../tests/runner-autogen.cpp)

target_link_libraries(xlnt.test xlnt)

add_custom_target (test-runner 
    COMMAND ${CMAKE_CURRENT_SOURCE_DIR}/../../generate-tests.sh
    COMMENT "Generating test runner tests/runner-autogen.cpp"
)

add_dependencies(xlnt.test test-runner)

add_custom_command(
    TARGET xlnt.test
    POST_BUILD
    COMMAND ${CMAKE_CURRENT_SOURCE_DIR}/../../../bin/xlnt.test
    VERBATIM
)
