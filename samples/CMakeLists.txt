include(${CMAKE_CURRENT_SOURCE_DIR}/../cmake/common.cmake)
project(${LIBRARY_NAME}.samples VERSION ${LIBRARY_VERSION} LANGUAGES CXX C)

if(NOT COMBINED_PROJECT)
    add_subdirectory(${CMAKE_CURRENT_SOURCE_DIR}/../source ${CMAKE_CURRENT_BINARY_DIR}/source)
endif()

add_executable(sample-basic ${CMAKE_CURRENT_SOURCE_DIR}/samples/sample.cpp)
target_link_libraries(sample-basic ${LIBRARY_NAME})

add_executable(sample-decrypt ${CMAKE_CURRENT_SOURCE_DIR}/samples/decrypt.cpp)
target_link_libraries(sample-decrypt ${LIBRARY_NAME})

file(COPY ${CMAKE_CURRENT_SOURCE_DIR}/data
    DESTINATION ${CMAKE_CURRENT_BINARY_DIR}/bin)
