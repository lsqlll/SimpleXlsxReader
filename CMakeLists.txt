cmake_minimum_required(VERSION 3.10)
project(simpleTableReader)

set(CMAKE_CXX_STANDARD 17)
set(CMAKE_CXX_STANDARD_REQUIRED ON)

# 设置第三方库路径
set(THIRD_PARTY_DIR ${CMAKE_SOURCE_DIR}/third-party)
set(GTEST_DIR ${THIRD_PARTY_DIR}/googletest)
set(LIBXLS_INCLUDE_DIR ${THIRD_PARTY_DIR}/libxls/include)
set(LIBXLS_LIBRARY_DIR ${THIRD_PARTY_DIR}/libxls/lib/mingw64/lib)
set(LIBICONV_LIBRARY_DIR ${THIRD_PARTY_DIR}/libiconv/source/build/lib)

add_subdirectory(${GTEST_DIR} ${CMAKE_BINARY_DIR}/googletest EXCLUDE_FROM_ALL)
set_target_properties(gtest gmock gtest_main PROPERTIES
    EXCLUDE_FROM_ALL TRUE
    EXCLUDE_FROM_DEFAULT_BUILD TRUE)

# 查找 libiconv - 先修正这部分
message(STATUS "Looking for libiconv...")

# 方法1: 使用 find_package
find_library(ICONV_LIBRARY
    NAMES libiconv iconv libiconv.la
    PATHS ${LIBICONV_LIBRARY_DIR}
    NO_DEFAULT_PATH
)

if(ICONV_LIBRARY)
    # 创建正确的导入目标
    message(STATUS "Found libiconv: ${ICONV_LIBRARY}")
    add_library(iconv::iconv STATIC IMPORTED)
    set_target_properties(iconv::iconv PROPERTIES
        IMPORTED_LOCATION ${ICONV_LIBRARY}
    )
else()
    message(FATAL_ERROR "libiconv not found! Please install it with: pacman -S mingw-w64-x86_64-libiconv")
endif()

# 查找 libxls
find_library(LIBXLS_LIBRARY
    NAMES xlsreader libxlsreader
    PATHS ${LIBXLS_LIBRARY_DIR}
    NO_DEFAULT_PATH
)

if(LIBXLS_LIBRARY)
    message(STATUS "Found libxls: ${LIBXLS_LIBRARY}")

    add_library(libxls::libxls STATIC IMPORTED)
    set_target_properties(libxls::libxls PROPERTIES
        IMPORTED_LOCATION ${LIBXLS_LIBRARY}
        INTERFACE_INCLUDE_DIRECTORIES ${LIBXLS_INCLUDE_DIR}
    )
else()
    message(FATAL_ERROR "libxls library not found at ${LIBXLS_LIBRARY_DIR}")
endif()

# 添加测试程序
add_executable(CellFuncTest test/CellfuncTest.cpp src/Exceptions.h)

# 链接库 - 确保顺序正确
target_link_libraries(CellFuncTest PRIVATE
    libxls::libxls
    iconv::iconv # libxls 依赖 iconv
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

# testStringViewUtils.cpp
add_executable(testStringViewUtils test/testStringViewUtils.cpp src/Exceptions.h)

target_link_libraries(testStringViewUtils PRIVATE
    gtest
    gmock
    gtest_main
)

target_include_directories(testStringViewUtils PRIVATE
    ${CMAKE_SOURCE_DIR}/src
    ${GTEST_DIR}/googletest/include
    ${GTEST_DIR}/googlemock/include
)

set_target_properties(testStringViewUtils PROPERTIES
    RUNTIME_OUTPUT_DIRECTORY ${CMAKE_BINARY_DIR}/bin
)

add_custom_target(build_gtest
    COMMENT "Building Google Test libraries"
    DEPENDS gtest gmock gtest_main
)
add_dependencies(CellFuncTest build_gtest)
add_dependencies(testStringViewUtils build_gtest)