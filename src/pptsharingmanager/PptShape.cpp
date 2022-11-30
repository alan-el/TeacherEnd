#include <QCryptographicHash>
#include <QString>
#include <QFile>
#include "PptShape.h"

PptShape::PptShape(QString filePathname)
{
    calcuMd5Hash(filePathname);
    mShapeIndexInSlide = 0;
    mIsSelected = false;
}

bool PptShape::calcuMd5Hash(QString filePathname)
{
    QFile file{filePathname};

    if(file.open(QIODevice::ReadOnly))
    {
        QByteArray data = file.readAll();
        mMd5Hash = QCryptographicHash::hash(QByteArrayView{data}, QCryptographicHash::Md5);
        file.close();
        return true;
    }

    return false;
}

QString PptShape::md5Hash() const
{
    QString ret{(mMd5Hash.toHex(0).toUpper())};
    return ret;
}


int PptShape::shapeIndexInSlide() const
{
    return mShapeIndexInSlide;
}

void PptShape::setShapeIndexInSlide(const int index)
{
    mShapeIndexInSlide = index;
}

bool PptShape::isSelected() const
{
    return mIsSelected;
}

void PptShape::setIsSelected(const bool isSel)
{
    mIsSelected = isSel;
}

