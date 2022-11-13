#ifndef PPTTEXTSHAPE_H
#define PPTTEXTSHAPE_H
#include "PptShape.h"
#include <QString>

class PptTextShape : public PptShape
{
public:
    PptTextShape(QString filePathname);
};

#endif // PPTTEXTSHAPE_H
