#include <liblouis.h>
#include <QDebug>
#include "BrailleTranslator.h"

BrailleTranslator::BrailleTranslator()
{

}

QString BrailleTranslator::brlTranslate(const QString plain)
{
    int len = plain.length();
    const wchar_t *pt16 = (const wchar_t *)plain.utf16();
    widechar *in = new widechar[len];

    for(int i = 0; i < len; i++)
        in[i] = pt16[i];

    widechar *out = new widechar[len * 3];
    int in_len = len;
    int out_len = len * 3;
    int ret = lou_translateString("D:/MyProject/BlinderReader/software/Chapter5/liblouis/liblouis-3.22.0-win64/share/liblouis/tables/zhcn-g1.ctb", in, &in_len, out, &out_len, NULL, NULL, noContractions);
    qDebug() << "Translate result: " << ret;

    char16_t *brl16 = new char16_t[out_len];
    for(int i = 0; i < out_len; i++)
        brl16[i] = out[i];

    QString brl = QString::fromUtf16(brl16, out_len);
    delete[] in;
    delete[] out;

    return brl;
}
