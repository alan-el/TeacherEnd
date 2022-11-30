add_executable(TeacherEnd
    teacherend.ui
    teacherend.cpp
    teacherend.h
    main.cpp
)

set(OPENCV_WORLD_STATIC_LIB "D:/opencv/build/install/x64/mingw/lib/libopencv_world460.dll.a")

target_link_libraries(TeacherEnd PRIVATE
    ${OPENCV_WORLD_STATIC_LIB}
    PptSharingManager
    Qt6::Widgets
    Qt6::AxContainer)

set(OPENCV_WORLD_INC_PATH
    D:/opencv/build/install/include
    D:/opencv/build/install/include/opencv2)

target_include_directories(TeacherEnd
    PRIVATE ${OPENCV_WORLD_INC_PATH})

set_target_properties(TeacherEnd PROPERTIES
    WIN32_EXECUTABLE ON
    MACOSX_BUNDLE ON
)

