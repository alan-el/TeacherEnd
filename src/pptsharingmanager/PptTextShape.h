#ifndef PPTTEXTSHAPE_H
#define PPTTEXTSHAPE_H
#include "PptShape.h"

#include <QString>
#include <QJsonObject>

class PptTextShape : public PptShape
{
public:
    PptTextShape() = default;
    PptTextShape(QString filePathname);

    QString plain() const;
    void setPlain(const QString& plain);

    QString braille() const;
    void setBraille(const QString& braille);

    void read(const QJsonObject& json);
    void write(QJsonObject& json) const;

    void print() const;

private:
    QString mPlain;      // save
    QString mBraille;    // save
};

#endif // PPTTEXTSHAPE_H
