project(xlnt.samples)

include_directories(${XLNT_INCLUDE_DIR})

if(SHARED)
    set(XLNT_LIB xlnt.shared)
else()
    set(XLNT_LIB xlnt.static)
endif()

add_executable(sample-basic ../samples/sample.cpp)
target_link_libraries(sample-basic ${XLNT_LIB})
set_target_properties(sample-basic
    PROPERTIES
    RUNTIME_OUTPUT_DIRECTORY_DEBUG "${CMAKE_BINARY_DIR}/bin"
    RUNTIME_OUTPUT_DIRECTORY_MINSIZEREL "${CMAKE_BINARY_DIR}/bin"
    RUNTIME_OUTPUT_DIRECTORY_RELWITHDEBINFO "${CMAKE_BINARY_DIR}/bin"
    RUNTIME_OUTPUT_DIRECTORY_RELEASE "${CMAKE_BINARY_DIR}/bin")

add_executable(sample-decrypt ../samples/decrypt.cpp)
target_link_libraries(sample-decrypt ${XLNT_LIB})
set_target_properties(sample-decrypt
    PROPERTIES
    RUNTIME_OUTPUT_DIRECTORY_DEBUG "${CMAKE_BINARY_DIR}/bin"
    RUNTIME_OUTPUT_DIRECTORY_MINSIZEREL "${CMAKE_BINARY_DIR}/bin"
    RUNTIME_OUTPUT_DIRECTORY_RELWITHDEBINFO "${CMAKE_BINARY_DIR}/bin"
    RUNTIME_OUTPUT_DIRECTORY_RELEASE "${CMAKE_BINARY_DIR}/bin")

file(COPY ${CMAKE_CURRENT_SOURCE_DIR}/data
    DESTINATION ${CMAKE_BINARY_DIR}/bin)
