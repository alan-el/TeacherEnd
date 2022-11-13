#ifndef PPTPICTURESHAPE_H
#define PPTPICTURESHAPE_H
#include "PptShape.h"
#include <QString>

class PptPictureShape : public PptShape
{
public:
    PptPictureShape(QString filePathname);
};

#endif // PPTPICTURESHAPE_H
