add_library(pptsharing STATIC
    PptSharing.ui
    PptSharing.cpp
    PptSharing.h
    JniMethod.cpp
    JniMethod.h
)

set(JVM_STATIC_LIB "D:/jdk-17.0.3.1/lib/jvm.lib")
target_link_libraries(pptsharing PRIVATE
    Qt6::Widgets
    Qt6::AxContainer
    ${JVM_STATIC_LIB})

set(JNI_INC_PATH D:/jdk-17.0.3.1/include D:/jdk-17.0.3.1/include/win32)
set(UI_INC_PATH D:/Qt_workspace/build/build-TeacherEnd/src/pptsharing/pptsharing_autogen/include)
include_directories(${JNI_INC_PATH}
    ${UI_INC_PATH})
target_include_directories(pptsharing INTERFACE ${CMAKE_CURRENT_SOURCE_DIR})
