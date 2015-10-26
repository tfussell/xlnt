project(xlnt.test)

include_directories(../../../include)
include_directories(../../../third-party/pugixml/src)
include_directories(../../../third-party/cxxtest)

add_executable(xlnt.test ../../../tests/runner-autogen.cpp)

target_link_libraries(xlnt.test xlnt)

add_custom_target (generate 
    COMMAND ../../../generate-tests.sh
    COMMENT "Generating test runner tests/runner-autogen.cpp"
)

add_dependencies(xlnt.test generate)

add_custom_command(
    TARGET xlnt.test
    POST_BUILD
    COMMAND ../../../../bin/xlnt.test
    VERBATIM
)
