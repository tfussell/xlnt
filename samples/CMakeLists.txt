include(${CMAKE_CURRENT_SOURCE_DIR}/../cmake/common.cmake)
project(${LIBRARY_NAME}.samples VERSION ${LIBRARY_VERSION} LANGUAGES CXX C)

if(NOT COMBINED_PROJECT)
    add_subdirectory(${CMAKE_CURRENT_SOURCE_DIR}/../source ${CMAKE_CURRENT_BINARY_DIR}/source)
endif()

include_directories(${LIBRARY_INCLUDE_DIR})

add_executable(sample-basic ${CMAKE_CURRENT_SOURCE_DIR}/sample.cpp)
target_link_libraries(sample-basic ${LIBRARY_NAME})

add_executable(sample-decrypt ${CMAKE_CURRENT_SOURCE_DIR}/decrypt.cpp)
target_link_libraries(sample-decrypt ${LIBRARY_NAME})

if(NOT MSVC)
    set(THREADS_PREFER_PTHREAD_FLAG ON)
    find_package(Threads REQUIRED)
    target_link_libraries(sample-basic Threads::Threads)
    target_link_libraries(sample-decrypt Threads::Threads)
endif()

file(COPY ${CMAKE_CURRENT_SOURCE_DIR}/data
    DESTINATION ${CMAKE_BINARY_DIR}/bin)
