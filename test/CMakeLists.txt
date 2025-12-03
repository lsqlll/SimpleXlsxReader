# CMakeLists.txt for unit tests
cmake_minimum_required(VERSION 3.14)
project(isEmptytest)

# 设置C++标准
set(CMAKE_CXX_STANDARD 17)
set(CMAKE_CXX_STANDARD_REQUIRED ON)

# 查找Google Test包
find_package(GTest REQUIRED)
find_package(Threads REQUIRED)

# 包含头文件目录
include_directories(
    ${PROJECT_SOURCE_DIR}/../src
    ${GTEST_INCLUDE_DIRS}
)

# 添加测试可执行文件
add_executable(isEmptytest
    isEmptyTest.cpp
)

# 链接必要的库
target_link_libraries(isEmptytest
    GTest::gtest
    GTest::gtest_main
    Threads::Threads
)

# 启用测试
enable_testing()

# 添加测试命令
add_test(
    NAME isEmptytest
    COMMAND isEmptytest
)