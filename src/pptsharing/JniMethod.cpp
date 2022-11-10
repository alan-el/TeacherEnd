#include "JniMethod.h"
#include <iostream>
using namespace std;

static JavaVM *jvm = nullptr;                      // Pointer to the JVM (Java Virtual Machine)
static JNIEnv *env = nullptr;                      // Pointer to native interface
void JniMethod::createJVM()
{
	if(jvm != nullptr)
		return;
	//================== prepare loading of Java VM ============================
	JavaVMInitArgs vm_args;                        // Initialization arguments
	JavaVMOption* options = new JavaVMOption[1];   // JVM invocation options
    options[0].optionString = (char *)"-Djava.class.path="
                                      "D:/eclipse_workspace/POIDemo/bin;"
                                      "D:/eclipse_workspace/poi-bin-5.2.2/lib/commons-codec-1.15.jar;"
                                      "D:/eclipse_workspace/poi-bin-5.2.2/lib/commons-collections4-4.4.jar;"
                                      "D:/eclipse_workspace/poi-bin-5.2.2/lib/commons-io-2.11.0.jar;"
                                      "D:/eclipse_workspace/poi-bin-5.2.2/lib/commons-math3-3.6.1.jar;"
                                      "D:/eclipse_workspace/poi-bin-5.2.2/lib/SparseBitSet-1.2.jar;"
                                      "D:/eclipse_workspace/apache-log4j-2.18.0-bin/log4j-api-2.18.0.jar;"
                                      "D:/eclipse_workspace/poi-bin-5.2.2/poi-5.2.2.jar;"
                                      "D:/eclipse_workspace/poi-bin-5.2.2/poi-scratchpad-5.2.2.jar;"
                                      "D:/eclipse_workspace/apache-log4j-2.18.0-bin/log4j-core-2.18.0.jar;"
                                      "D:/eclipse_workspace/poi-bin-5.2.2/poi-ooxml-5.2.2.jar;"
                                      "D:/eclipse_workspace/poi-bin-5.2.2/poi-ooxml-lite-5.2.2.jar;"
                                      "D:/eclipse_workspace/poi-bin-5.2.2/ooxml-lib/commons-compress-1.21.jar;"
                                      "D:/eclipse_workspace/poi-bin-5.2.2/ooxml-lib/xmlbeans-5.0.3.jar";   // where to find java .class
	vm_args.version = JNI_VERSION_1_8;             // minimum Java version
	vm_args.nOptions = 1;                          // number of options
	vm_args.options = options;
	vm_args.ignoreUnrecognized = false;     // invalid options make the JVM init fail
		//=============== load and initialize Java VM and JNI interface =============
	jint rc = JNI_CreateJavaVM(&jvm, (void**)&env, &vm_args);  // YES !!
    delete[] options;    // we then no longer need the initialisation options.
	if(rc != JNI_OK) {
		// TO DO: error processing... 
		cin.get();
		exit(EXIT_FAILURE);
	}
	//=============== Display JVM version =======================================
	cout << "JVM load succeeded: Version ";
	jint ver = env->GetVersion();
	cout << ((ver >> 16) & 0x0f) << "." << (ver & 0x0f) << endl;
}

void JniMethod::destroyJVM()
{
	jvm->DestroyJavaVM();
	jvm = nullptr;
}


int JniMethod::pptTextExtractor(const char *pathname, bool isPptx, int index)
{
	// TO DO: add the code that will use JVM <============  (see next steps)
	jint slides_num = 0;
	jclass cls = env->FindClass("com/alanel/pptparse/PptTextExtraction");  // try to find the class
	if(cls == nullptr) {
		cerr << "ERROR: class not found !";
	}
	else
	{   // if class found, continue
		jmethodID mid;
		if(!isPptx)
			mid = env->GetStaticMethodID(cls, "PptSingleSlideTextExtractor", "(Ljava/lang/String;I)I");  // find method
		else
			mid = env->GetStaticMethodID(cls, "PptxSingleSlideTextExtractor", "(Ljava/lang/String;I)I");

		if(mid == nullptr)
			cerr << "ERROR: method not found !" << endl;
		else
		{
			jstring str = env->NewStringUTF(pathname);
			slides_num = env->CallStaticIntMethod(cls, mid, str, index);   // call the method with the arr as argument.
			env->DeleteLocalRef(str);					// release the object
		}
	}
	return slides_num;
}

void JniMethod::pptPictExtractor(const char * pathname, bool isPptx, int index)
{
	// TO DO: add the code that will use JVM <============  (see next steps)
	jclass cls = env->FindClass("com/alanel/pptparse/PptPictExtraction");  // try to find the class
	if(cls == nullptr) {
		cerr << "ERROR: class not found !";
	}
	else
	{   // if class found, continue
		jmethodID mid;
		if(!isPptx)
			mid = env->GetStaticMethodID(cls, "PptSingleSlidePictExtractor", "(Ljava/lang/String;I)V");  // find method
		else
			mid = env->GetStaticMethodID(cls, "PptxSingleSlidePictExtractor", "(Ljava/lang/String;I)V");

		if(mid == nullptr)
			cerr << "ERROR: method not found !" << endl;
		else
		{
			jstring str = env->NewStringUTF(pathname);
			env->CallStaticVoidMethod(cls, mid, str, index);   // call the method with the arr as argument.
			env->DeleteLocalRef(str);					// release the object
		}
	}
}
