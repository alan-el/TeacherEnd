#include <QCryptographicHash>
#include <QString>
#include <QFile>
#include "PptShape.h"

PptShape::PptShape(QString filePathname)
{
    calcuMd5Hash(filePathname);
    shapeIndexInSlide = 0;
    isSelected = false;
}

bool PptShape::calcuMd5Hash(QString filePathname)
{
    QFile file{filePathname};

    if(file.open(QIODevice::ReadOnly))
    {
        QByteArray data = file.readAll();
        md5Hash = QCryptographicHash::hash(QByteArrayView{data}, QCryptographicHash::Md5);
        file.close();
        return true;
    }

    return false;
}

QString PptShape::getMd5Hash()
{
    QString ret{(md5Hash.toHex(0).toUpper())};
    return ret;
}


int PptShape::getShapeIndexInSlide()
{
    return shapeIndexInSlide;
}

void PptShape::setShapeIndexInSlide(int index)
{
    shapeIndexInSlide = index;
}

bool PptShape::getIsSelected()
{
    return isSelected;
}

void PptShape::setIsSelected(bool isSel)
{
    isSelected = isSel;
}

