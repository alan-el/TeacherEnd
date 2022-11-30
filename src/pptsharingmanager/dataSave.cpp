#include "dataSave.h"

#include <QCborMap>
#include <QCborValue>
#include <QFile>
#include <QJsonArray>
#include <QJsonDocument>
#include <QDebug>

DataSave::DataSave(SaveMode saveMode)
{
    if (saveMode == Test)
    {
        mFileName = "Test";

        PptTextShape txt{"D:/Qt_workspace/src/TeacherEnd/doc/贵俊涛-毕业论文答辩/texts/slide28_text1.txt"};
        txt.setPlain("4.Hi3519消息处理模块");
        txt.setBraille("⠄⠓⠩⠄⠱⠆⠅⠊⠆⠙⠀⠓⠊⠆⠞");
        txt.setIsSelected("true");
        mTexts.append(txt);

        PptPictureShape pic{"D:/Qt_workspace/src/TeacherEnd/doc/贵俊涛-毕业论文答辩/pictures/slide28_pict1.png"};
        pic.setWidth(48);
        pic.setHeight(32);
        pic.setIsSelected(false);
        pic.setDotsPictureData("XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX");
        mPictures.append(pic);
    }
}

bool DataSave::loadData(SaveFormat saveFormat, const QString& pathname)
{
    QFile loadFile(saveFormat == Json
                   ? (pathname + ".json")
                   : (pathname) + ".dat");

    if (!loadFile.open(QIODevice::ReadOnly))
    {
        qDebug() << "Couldn't open save file.";
        return false;
    }

    QByteArray dataSaved = loadFile.readAll();
    QJsonDocument loadDoc(saveFormat == Json
        ? QJsonDocument::fromJson(dataSaved)
        : QJsonDocument(QCborValue::fromCbor(dataSaved).toMap().toJsonObject()));

    read(loadDoc.object());

    qDebug() << "Loaded save for "
             << loadDoc["fileName"].toString()
             << " using "
             << (saveFormat != Json ? "CBOR" : "JSON") << "...\n";
    return true;
}

bool DataSave::saveData(SaveFormat saveFormat, const QString& pathname) const
{
    QFile saveFile(saveFormat == Json
                   ? (pathname + ".json")
                   : (pathname) + ".dat");

    if (!saveFile.open(QIODevice::WriteOnly))
    {
        qDebug() << "Couldn't open save file.";
        return false;
    }

    QJsonObject dataObject;
    write(dataObject);
    saveFile.write(saveFormat == Json
                   ? QJsonDocument(dataObject).toJson()
                   : QCborValue::fromJsonValue(dataObject).toCbor());

    return true;
}

void DataSave::read(const QJsonObject &json)
{
    if (json.contains("fileName") && json["fileName"].isString())
        mFileName = json["fileName"].toString();

    if(json.contains("texts") && json["texts"].isArray())
    {
        QJsonArray textsArray = json["texts"].toArray();
        mTexts.clear();
        mTexts.reserve(textsArray.size());
        for(int index = 0; index < textsArray.size(); index++)
        {
            QJsonObject textObj = textsArray[index].toObject();
            PptTextShape txtShape;
            txtShape.read(textObj);
            mTexts.append(txtShape);
        }
    }

    if(json.contains("pictures") && json["pictures"].isArray())
    {
        QJsonArray picsArray = json["pictures"].toArray();
        mPictures.clear();
        mPictures.reserve(picsArray.size());
        for(int index = 0; index < picsArray.size(); index++)
        {
            QJsonObject picObj = picsArray[index].toObject();
            PptPictureShape picShape;
            picShape.read(picObj);
            picShape.setShapeIndexInSlide(0);
            mPictures.append(picShape);
        }
    }
}

void DataSave::write(QJsonObject &json) const
{
    json["fileName"] = mFileName;

    QJsonArray textsArray;
    for(const PptTextShape txt : mTexts) {
        QJsonObject txtObj;
        txt.write(txtObj);
        textsArray.append(txtObj);
    }
    json["texts"] = textsArray;

    QJsonArray picturesArray;
    for(const PptPictureShape pic : mPictures) {
        QJsonObject picObj;
        pic.write(picObj);
        picturesArray.append(picObj);
    }
    json["pictures"] = picturesArray;
}

QString DataSave::fileName() const
{
    return mFileName;
}

void DataSave::setFileName(const QString &fn)
{
    mFileName = fn;
}

QList<PptTextShape> DataSave::texts() const
{
    return mTexts;
}

QList<PptPictureShape> DataSave::pictures() const
{
    return mPictures;
}

void DataSave::clear()
{
    mTexts.clear();
    mPictures.clear();
}

void DataSave::append(const PptTextShape &txtShape)
{
    mTexts.append(txtShape);
}

void DataSave::append(const PptPictureShape &picShape)
{
    mPictures.append(picShape);
}

void DataSave::print() const
{
    qDebug() << "File Name:\t" << mFileName;
    for(const PptTextShape& txt : mTexts)
    {
        qDebug() << "Text: ";
        txt.print();
    }
    for(const PptPictureShape& pic : mPictures)
    {
        qDebug() << "Picture: ";
        pic.print();
    }

}
