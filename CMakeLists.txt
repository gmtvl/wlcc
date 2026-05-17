cmake_minimum_required(VERSION 3.20)

project(ExcelCM LANGUAGES CXX)

set(CMAKE_CXX_STANDARD 20)

# Add OpenXLSX library
add_subdirectory(external/OpenXLSX)


add_executable(excel_cm
    src/main.cpp
    src/excel_cm.cpp
)

target_include_directories(excel_cm
    PRIVATE
        ${PROJECT_SOURCE_DIR}/include
)

target_link_libraries(excel_cm
    PRIVATE
        OpenXLSX::OpenXLSX
)