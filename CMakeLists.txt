cmake_minimum_required(VERSION 3.10)
project(simpleTableReader)

set(CMAKE_CXX_STANDARD 17)
set(CMAKE_CXX_STANDARD_REQUIRED ON)

set(CMAKE_C_COMPILER "E:/msys64/mingw64/bin/gcc.exe")
set(CMAKE_CXX_COMPILER "E:/msys64/mingw64/bin/g++.exe")

# 设置第三方库路径
set(THIRD_PARTY_DIR ${CMAKE_SOURCE_DIR}/third-party)
set(GTEST_DIR ${THIRD_PARTY_DIR}/googletest)

add_subdirectory(${GTEST_DIR} ${CMAKE_BINARY_DIR}/googletest EXCLUDE_FROM_ALL)

# 查找 libxls（如果已经编译好）
set(LIBXLS_INCLUDE_DIR ${THIRD_PARTY_DIR}/libxls/include)
set(LIBXLS_LIBRARY_DIR ${THIRD_PARTY_DIR}/libxls/lib/mingw64/lib)

# 查找 libxls 库文件
find_library(LIBXLS_LIBRARY
    NAMES xls libxls xlsreader
    PATHS ${LIBXLS_LIBRARY_DIR}
    NO_DEFAULT_PATH
)

# 创建 libxls 导入目标
if(LIBXLS_LIBRARY)
    message(STATUS "Found libxls: ${LIBXLS_LIBRARY}")
    
    add_library(libxls::libxls STATIC IMPORTED)
    set_target_properties(libxls::libxls PROPERTIES
        IMPORTED_LOCATION ${LIBXLS_LIBRARY}
        INTERFACE_INCLUDE_DIRECTORIES ${LIBXLS_INCLUDE_DIR}
    )
else()
    # 如果没找到预编译的 libxls，尝试从源码构建
    message(STATUS "libxls library not found")
endif()

# 创建可执行文件
add_executable(CellFuncTest test/CellfuncTest.cpp src/Exceptions.h)

# 链接库
target_link_libraries(CellFuncTest PRIVATE
    libxls::libxls
    gtest
    gmock
    gtest_main
)

# 包含目录
target_include_directories(CellFuncTest PRIVATE
    ${CMAKE_SOURCE_DIR}/src
    ${LIBXLS_INCLUDE_DIR}
    ${GTEST_DIR}/googletest/include
    ${GTEST_DIR}/googlemock/include
)

# 设置输出目录
set_target_properties(CellFuncTest PROPERTIES
    RUNTIME_OUTPUT_DIRECTORY ${CMAKE_BINARY_DIR}/bin
)
