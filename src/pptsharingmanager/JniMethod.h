#pragma once

#include <jni.h>

class JniMethod
{
public:
	static void createJVM();
	static void destroyJVM();

	static int pptTextExtractor(const char *pathname, bool isPptx, int index);
	static void pptPictExtractor(const char *pathname, bool isPptx, int index);
};

