project(xlnt.samples)

include_directories(${XLNT_INCLUDE_DIR})

if(SHARED)
    set(XLNT_LIB xlnt.shared)
else()
    set(XLNT_LIB xlnt.static)
endif()

add_executable(sample-basic ../samples/sample.cpp)
target_link_libraries(sample-basic ${XLNT_LIB})

add_executable(sample-decrypt ../samples/decrypt.cpp)
target_link_libraries(sample-decrypt ${XLNT_LIB})

file(COPY ${CMAKE_CURRENT_SOURCE_DIR}/data
    DESTINATION ${CMAKE_BINARY_DIR}/bin)
