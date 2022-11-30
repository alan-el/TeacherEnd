#include "PptPictureShape.h"

#include <QDebug>

PptPictureShape::PptPictureShape(QString filePathname) : PptShape(filePathname)
{
    mWidth = 0;
    mHeight = 0;
    mDotsPictureData.clear();
}

QImage PptPictureShape::image() const
{
    return mImage;
}

void PptPictureShape::setImage(const QImage &image)
{
    mImage = image.copy();
}

int PptPictureShape::width() const
{
    return mWidth;
}

void PptPictureShape::setWidth(const int width)
{
    mWidth = width;
}

int PptPictureShape::height() const
{
    return mHeight;
}

void PptPictureShape::setHeight(const int height)
{
    mHeight = height;
}

QString PptPictureShape::dotsPictureData() const
{
    return mDotsPictureData;
}

void PptPictureShape::setDotsPictureData(const QString &data)
{
    mDotsPictureData = data;
}

void PptPictureShape::read(const QJsonObject &json)
{
    if (json.contains("md5") && json["md5"].isString())
    {
        QString md5 = json["md5"].toString();
        for(int i = 0; i < md5.size(); i+=2)
        {
            bool ok;
            char temp = md5.sliced(i, 2).toInt(&ok, 16);
            mMd5Hash.append(temp);
        }
    }

    if (json.contains("isSelected") && json["isSelected"].isBool())
        mIsSelected = json["isSelected"].toBool();

    if (json.contains("dotsPicture") && json["dotsPicture"].isObject())
    {
        QJsonObject dp = json["dotsPicture"].toObject();

        if(dp.contains("width") && dp["width"].isDouble())
            mWidth = dp["width"].toInt();

        if(dp.contains("height") && dp["height"].isDouble())
            mHeight = dp["height"].toInt();

        if(dp.contains("data") && dp["data"].isString())
            mDotsPictureData = dp["data"].toString();
    }
}

void PptPictureShape::write(QJsonObject &json) const
{
    json["md5"] = md5Hash();
    json["isSelected"] = isSelected();

    QJsonObject dpObj;
    dpObj["width"] = mWidth;
    dpObj["height"] = mHeight;
    dpObj["data"] = mDotsPictureData;

    json["dotsPicture"] = dpObj;
}

void PptPictureShape::print() const
{
    qDebug() << "\t" << "md5: " << md5Hash();
    qDebug() << "\t" << "isSelected: " << isSelected();
    qDebug() << "\t" << "resolution: " << width() << "×" << height();
    qDebug() << "\t" << "data: " << mDotsPictureData;
}

