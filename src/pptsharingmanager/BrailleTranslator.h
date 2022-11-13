#ifndef BRAILLETRANSLATOR_H
#define BRAILLETRANSLATOR_H
#include <QString>

class BrailleTranslator
{
public:
    BrailleTranslator();

    static QString brlTranslate(QString& plain);
};

#endif // BRAILLETRANSLATOR_H
