#ifndef DATASAVE_H
#define DATASAVE_H

#include "PptTextShape.h"
#include "PptPictureShape.h"

#include <QJsonObject>
#include <QList>

class DataSave
{
public:
    enum SaveMode {
        Test
    };

    enum SaveFormat {
        Json, Binary
    };


    DataSave() = default;

    DataSave(SaveMode saveMode);

    bool loadData(SaveFormat saveFormat, const QString& pathname);
    bool saveData(SaveFormat saveFormat, const QString& pathname) const;

    void read(const QJsonObject &json);
    void write(QJsonObject &json) const;

    QString fileName() const;
    void setFileName(const QString& fn);

    QList<PptTextShape> texts() const;
    QList<PptPictureShape> pictures() const;

    void clear();

    void append(const PptTextShape& txtShape);
    void append(const PptPictureShape& picShape);

    void print() const;

private:
    QString mFileName;
    QList<PptTextShape> mTexts;
    QList<PptPictureShape> mPictures;
};

#endif // DATASAVE_H
