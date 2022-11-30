
add_library(PptSharingManager STATIC
    pptsharing.ui
    pptsharingmanager.cpp
    pptsharingmanager.h
    jnimethod.cpp
    jnimethod.h
    pptshape.cpp
    pptshape.h
    ppttextshape.cpp
    ppttextshape.h
    pptpictureshape.cpp
    pptpictureshape.h
    brailletranslator.cpp
    brailletranslator.h
    datasave.cpp
    datasave.h
)

set(JVM_STATIC_LIB "D:/jdk-17.0.3.1/lib/jvm.lib")
set(LOUIS_STATIC_LIB "D:/Qt_workspace/src/TeacherEnd/lib/liblouis/liblouis.dll")
target_link_libraries(PptSharingManager PRIVATE
    Qt6::Widgets
    Qt6::AxContainer
    ${JVM_STATIC_LIB}
    ${LOUIS_STATIC_LIB})

set(JNI_INC_PATH D:/jdk-17.0.3.1/include D:/jdk-17.0.3.1/include/win32)
set(LOUIS_INC_PATH ../../include/liblouis)
target_include_directories(PptSharingManager
    PRIVATE ${JNI_INC_PATH} ${LOUIS_INC_PATH}
    INTERFACE ${CMAKE_CURRENT_SOURCE_DIR})
