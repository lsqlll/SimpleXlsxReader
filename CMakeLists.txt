cmake_minimum_required(VERSION 3.10)
project(simpleTableReader)

set(CMAKE_CXX_STANDARD 17)
set(CMAKE_CXX_STANDARD_REQUIRED ON)

# 设置第三方库路径
set(THIRD_PARTY_DIR ${CMAKE_SOURCE_DIR}/third-party)
set(GTEST_DIR ${THIRD_PARTY_DIR}/googletest)
set(BOOST_DIR ${THIRD_PARTY_DIR}/Boost_1_89_0)

# 完全禁用系统的 Boost 配置
set(Boost_NO_BOOST_CMAKE ON)
set(Boost_NO_SYSTEM_PATHS ON)
set(BOOST_ROOT ${BOOST_DIR})
set(BOOST_INCLUDEDIR ${BOOST_DIR}/include)

add_subdirectory(${GTEST_DIR} ${CMAKE_BINARY_DIR}/googletest EXCLUDE_FROM_ALL)
set_target_properties(gtest gmock gtest_main PROPERTIES
    EXCLUDE_FROM_ALL TRUE
    EXCLUDE_FROM_DEFAULT_BUILD TRUE)

# 查找 libxls
set(LIBXLS_INCLUDE_DIR ${THIRD_PARTY_DIR}/libxls/include)
set(LIBXLS_LIBRARY_DIR ${THIRD_PARTY_DIR}/libxls/lib/mingw64/lib)

find_library(LIBXLS_LIBRARY
    NAMES xls libxls xlsreader
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
    message(STATUS "libxls library not found")
endif()

# 检查 Boost 目录结构
message(STATUS "Checking Boost directory: ${BOOST_DIR}")

if(EXISTS ${BOOST_DIR}/include/boost)
    message(STATUS "Boost headers found")
else()
    message(WARNING "Boost headers not found at ${BOOST_DIR}/include/boost")
endif()

add_executable(CellFuncTest test/CellfuncTest.cpp src/Exceptions.h)

# 链接库
target_link_libraries(CellFuncTest PRIVATE
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

# 链接库
target_link_libraries(testStringViewUtils PRIVATE
    gtest
    gmock
    gtest_main
)

# 包含目录
target_include_directories(testStringViewUtils PRIVATE
    ${CMAKE_SOURCE_DIR}/src
    ${GTEST_DIR}/googletest/include
    ${GTEST_DIR}/googlemock/include
)

# 设置输出目录
set_target_properties(testStringViewUtils PROPERTIES
    RUNTIME_OUTPUT_DIRECTORY ${CMAKE_BINARY_DIR}/bin
)

add_custom_target(build_gtest
    COMMENT "Building Google Test libraries"
    DEPENDS gtest gmock gtest_main
)
add_dependencies(CellFuncTest build_gtest)
add_dependencies(testStringViewUtils build_gtest)