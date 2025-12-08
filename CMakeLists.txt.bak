cmake_minimum_required(VERSION 3.10)
project(simpleTableReader)

set(CMAKE_CXX_STANDARD 17)
set(CMAKE_CXX_STANDARD_REQUIRED ON)
set(CMAKE_C_COMPILER "E:/msys64/mingw64/bin/gcc.exe")
set(CMAKE_CXX_COMPILER "E:/msys64/mingw64/bin/g++.exe")


# 方法1：优先使用系统安装的 gtest/gmock
find_package(GTest QUIET)

# 方法2：如果系统没有，使用 pkg-config
if(NOT GTest_FOUND)
    find_package(PkgConfig REQUIRED)
    pkg_check_modules(GTEST REQUIRED gtest)
    pkg_check_modules(GMOCK REQUIRED gmock)
    pkg_check_modules(LIBXLS REQUIRED libxls)
    
    # 创建导入目标
    add_library(GTest::gtest INTERFACE IMPORTED)
    set_target_properties(GTest::gtest PROPERTIES
        INTERFACE_INCLUDE_DIRECTORIES "${GTEST_INCLUDE_DIRS}"
        INTERFACE_LINK_LIBRARIES "${GTEST_LIBRARIES}"
    )
    
    add_library(GTest::gmock INTERFACE IMPORTED)
    set_target_properties(GTest::gmock PROPERTIES
        INTERFACE_INCLUDE_DIRECTORIES "${GMOCK_INCLUDE_DIRS}"
        INTERFACE_LINK_LIBRARIES "${GMOCK_LIBRARIES}"
    )
    
    add_library(GTest::gtest_main INTERFACE IMPORTED)
    set_target_properties(GTest::gtest_main PROPERTIES
        INTERFACE_LINK_LIBRARIES "${GTEST_LIBRARIES}"
    )

    add_library(libxls::libxls INTERFACE IMPORTED)
    set_target_properties(libxls::libxls PROPERTIES
      INTERFACE_LINK_LIBRARIES "${LIBXLS_LIBRARIES}")
endif()

# 使用 pkg-config 查找 libxls（第三方库）
find_package(PkgConfig REQUIRED)

message(STATUS "Found libxls:")
message(STATUS "  Include dirs: ${LIBXLS_INCLUDE_DIRS}")
message(STATUS "  Libraries: ${LIBXLS_LIBRARIES}")

# 创建 libxls 导入目标

add_library(libxls::libxls INTERFACE IMPORTED)
set_target_properties(libxls::libxls PROPERTIES
    INTERFACE_INCLUDE_DIRECTORIES "${LIBXLS_INCLUDE_DIRS}"
    INTERFACE_LINK_LIBRARIES "${LIBXLS_LIBRARIES}"
    INTERFACE_COMPILE_OPTIONS "${LIBXLS_CFLAGS_OTHER}"
)

add_executable(CellFuncTest test/CellfuncTest.cpp src/Exceptions.h)

# 链接库
target_link_libraries(CellFuncTest PRIVATE
    libxls::libxls
    GTest::gtest
    GTest::gmock
    GTest::gtest_main
)
target_include_directories(CellFuncTest SYSTEM PRIVATE
    /mingw64/include
    /mingw64/x86_64-w64-mingw32/include
)
