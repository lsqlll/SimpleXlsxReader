cmake_minimum_required(VERSION 3.10)
project(simpleTableReader)

set(CMAKE_CXX_STANDARD 17)
set(CMAKE_CXX_STANDARD_REQUIRED ON)



# 设置第三方库路径
set(THIRD_PARTY_DIR ${CMAKE_SOURCE_DIR}/third-party)
# 查找libxls头文件
find_path(LIBXLS_INCLUDE_DIR
    NAMES xls.h
    PATHS ${THIRD_PARTY_DIR}/libxls/include
    NO_DEFAULT_PATH
)

# 查找libxls库文件（假设库文件在lib目录下）
find_library(LIBXLS_LIBRARY
    NAMES xlsreader
    PATHS ${THIRD_PARTY_DIR}/libxls/lib
    NO_DEFAULT_PATH
)

# 如果没找到库文件，尝试其他可能的名称
if(NOT LIBXLS_LIBRARY)
    find_library(LIBXLS_LIBRARY
        NAMES xls
        PATHS ${THIRD_PARTY_DIR}/libxls/lib
        NO_DEFAULT_PATH
    )
endif()

if(LIBXLS_INCLUDE_DIR AND LIBXLS_LIBRARY)
    message(STATUS "Found libxls: ${LIBXLS_LIBRARY}")
    message(STATUS "Include dir: ${LIBXLS_INCLUDE_DIR}")
else()

    message(FATAL_ERROR "libxls not found! Please check the third-party/libxls directory")
    message(STATUS "Searched in: ${THIRD_PARTY_DIR}/libxls")
endif()

add_executable(xls_to_csv src/excelReader.cpp src/utils.cpp)
target_include_directories(xls_to_csv PRIVATE ${LIBXLS_INCLUDE_DIR})
target_link_libraries(xls_to_csv ${LIBXLS_LIBRARY})

    NO_DEFAULT_PATH
)

# 如果没找到库文件，尝试其他可能的名称
if(NOT LIBXLS_LIBRARY)
    find_library(LIBXLS_LIBRARY
        NAMES xls
        PATHS ${THIRD_PARTY_DIR}/libxls/lib
        NO_DEFAULT_PATH
    )
endif()

if(LIBXLS_INCLUDE_DIR AND LIBXLS_LIBRARY)
    message(STATUS "Found libxls: ${LIBXLS_LIBRARY}")
    message(STATUS "Include dir: ${LIBXLS_INCLUDE_DIR}")
else()
    message(FATAL_ERROR "libxls not found! Please check the third-party/libxls directory")
    message(STATUS "Searched in: ${THIRD_PARTY_DIR}/libxls")
endif()

add_executable(xls_to_csv src/excelReader.cpp src/utils.cpp)
target_include_directories(xls_to_csv PRIVATE ${LIBXLS_INCLUDE_DIR})
target_link_libraries(xls_to_csv ${LIBXLS_LIBRARY})cmake_minimum_required(VERSION 3.10)
project(simpleTableReader)

set(CMAKE_CXX_STANDARD 17)
set(CMAKE_CXX_STANDARD_REQUIRED ON)

# 手动查找libxls
find_path(LIBXLS_INCLUDE_DIR )

if(LIBXLS_INCLUDE_DIR AND LIBXLS_LIBRARY)
    message(STATUS "Found libxls: ${LIBXLS_LIBRARY}")
    message(STATUS "Include dir: ${LIBXLS_INCLUDE_DIR}")
else()
    message(FATAL_ERROR "libxls not found! Please install libxls-dev package")
endif()

add_executable(xls_to_csv src/excelReader.cpp src/utils.cpp)
target_include_directories(xls_to_csv PRIVATE ${LIBXLS_INCLUDE_DIR})
target_link_libraries(xls_to_csv ${LIBXLS_LIBRARY})
