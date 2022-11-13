add_executable(TeacherEnd
    TeacherEnd.ui
    TeacherEnd.cpp
    TeacherEnd.h
    main.cpp
)

target_link_libraries(TeacherEnd PRIVATE
    PptSharingManager
    Qt6::Widgets
    Qt6::AxContainer)

set_target_properties(TeacherEnd PROPERTIES
    WIN32_EXECUTABLE ON
    MACOSX_BUNDLE ON
)

