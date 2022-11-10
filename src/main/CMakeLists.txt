add_executable(teacherend
    TeacherEnd.ui
    TeacherEnd.cpp
    TeacherEnd.h
    main.cpp
)

target_link_libraries(teacherend PRIVATE
    pptsharing
    Qt6::Widgets
    Qt6::AxContainer)

set_target_properties(teacherend PROPERTIES
    WIN32_EXECUTABLE ON
    MACOSX_BUNDLE ON
)

