cmake_minimum_required(VERSION 3.10)
project(XlsToCsvConverter)

set(CMAKE_CXX_STANDARD 17)
set(CMAKE_CXX_STANDARD_REQUIRED ON)

# 手动查找libxls
find_path(LIBXLS_INCLUDE_DIR NAMES xls.h)
find_library(LIBXLS_LIBRARY NAMES xlsreader xls)

if(LIBXLS_INCLUDE_DIR AND LIBXLS_LIBRARY)
    message(STATUS "Found libxls: ${LIBXLS_LIBRARY}")
    message(STATUS "Include dir: ${LIBXLS_INCLUDE_DIR}")
else()
    message(FATAL_ERROR "libxls not found! Please install libxls-dev package")
endif()

add_executable(xls_to_csv src/excelReader.cpp src/utils.cpp)
target_include_directories(xls_to_csv PRIVATE ${LIBXLS_INCLUDE_DIR})
target_link_libraries(xls_to_csv ${LIBXLS_LIBRARY})
