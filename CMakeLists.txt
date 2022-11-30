cmake_minimum_required(VERSION 3.16)

project(TeacherEnd VERSION 1.0.0 LANGUAGES CXX)

set(CMAKE_CXX_STANDARD 17)
set(CMAKE_CXX_STANDARD_REQUIRED ON)

find_package(Qt6 REQUIRED COMPONENTS Widgets AxContainer)
qt_standard_project_setup()

add_subdirectory(src/main)
add_subdirectory(src/pptsharingmanager)

# Resources:
set(resource_files
    "res/assets/next.jpg"
    "res/assets/play.jpg"
    "res/assets/play_40.jpg"
    "res/assets/prev.jpg"
)

qt6_add_resources(TeacherEnd "helloworld"
    PREFIX
        "/"
    FILES
        ${resource_files}
)


