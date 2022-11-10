cmake_minimum_required(VERSION 3.16)

project(TeacherEnd VERSION 1.0.0 LANGUAGES CXX)

set(CMAKE_CXX_STANDARD 17)
set(CMAKE_CXX_STANDARD_REQUIRED ON)


find_package(Qt6 REQUIRED COMPONENTS Widgets AxContainer)
qt_standard_project_setup()

add_subdirectory(src/main)
add_subdirectory(src/pptsharing)

# Resources:
#set(resource_files
#    "images/ThunderCliff.jpg"
#)

#qt6_add_resources(helloworld "helloworld"
#    PREFIX
#        "/"
#    FILES
#        ${resource_files}
#)
