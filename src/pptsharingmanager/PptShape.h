#ifndef PPTSHAPE_H
#define PPTSHAPE_H
#include <QByteArray>

class PptShape
{
public:
    PptShape(QString filePathname);

    QString getMd5Hash();
    int getShapeIndexInSlide();
    void setShapeIndexInSlide(int index);
    bool getIsSelected();
    void setIsSelected(bool isSel);

private:
    bool calcuMd5Hash(QString filePathname);
    // TODO Shape 的触点表示生成方法

protected:
    QByteArray md5Hash;
    int shapeIndexInSlide;
    bool isSelected;
    // TODO 需要属性: Shape 的触点表示 (盲文 or 触点图像)
};

#endif // PPTSHAPE_H
