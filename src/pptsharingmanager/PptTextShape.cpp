#include "PptTextShape.h"

PptTextShape::PptTextShape(QString filePathname): PptShape(filePathname)
{
    mPlain.clear();
    mBraille.clear();
}

QString PptTextShape::plain() const
{
    return mPlain;
}

void PptTextShape::setPlain(const QString& plain)
{
    mPlain = plain;
}

QString PptTextShape::braille() const
{
    return mBraille;
}

void PptTextShape::setBraille(const QString& braille)
{
    mBraille = braille;
}

void PptTextShape::read(const QJsonObject& json)
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

    if (json.contains("plain") && json["plain"].isString())
        mPlain = json["plain"].toString();

    if (json.contains("braille") && json["braille"].isString())
        mBraille = json["braille"].toString();
}

void PptTextShape::write(QJsonObject& json) const
{
    json["md5"] = md5Hash();
    json["isSelected"] = mIsSelected;
    json["plain"] = mPlain;
    json["braille"] = mBraille;
}

void PptTextShape::print() const
{
    qDebug() << "\t" << "md5: " << md5Hash();
    qDebug() << "\t" << "isSelected: " << isSelected();
    qDebug() << "\t" << "plain: " << plain();
    qDebug() << "\t" << "braille: " << braille();
}
