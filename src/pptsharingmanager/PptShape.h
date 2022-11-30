#ifndef PPTSHAPE_H
#define PPTSHAPE_H
#include <QByteArray>

class PptShape
{
public:
    PptShape() = default;
    PptShape(QString filePathname);

    QString md5Hash() const;

    int shapeIndexInSlide()const;
    void setShapeIndexInSlide(const int index);

    bool isSelected()const;
    void setIsSelected(const bool isSel);

private:
    bool calcuMd5Hash(QString filePathname);

protected:
    QByteArray mMd5Hash;    // save
    int mShapeIndexInSlide;
    bool mIsSelected;       // save
};

#endif // PPTSHAPE_H
