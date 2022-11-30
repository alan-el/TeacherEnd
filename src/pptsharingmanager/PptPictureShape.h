#ifndef PPTPICTURESHAPE_H
#define PPTPICTURESHAPE_H
#include "PptShape.h"

#include <QString>
#include <QImage>
#include <QJsonObject>

class PptPictureShape : public PptShape
{
public:
    PptPictureShape() = default;
    PptPictureShape(QString filePathname);

    QImage image() const;
    void setImage(const QImage& image);

    int width() const;
    void setWidth(const int width);

    int height() const;
    void setHeight(const int height);

    QString dotsPictureData() const;
    void setDotsPictureData(const QString& data);

    void read(const QJsonObject &json);
    void write(QJsonObject &json) const;

    void print() const;

private:
    QImage mImage;
    int mWidth;     // save
    int mHeight;     // save
    QString mDotsPictureData;    // save 十六进制数表示组成二维触点阵列的点方的触点凸起情况(根据分辨率从左到右从上到下的顺序)
};

#endif // PPTPICTURESHAPE_H
