#ifndef PPTSHARING_H
#define PPTSHARING_H

#include <QWidget>
#include <QAxObject>
#include <QDateTime>
#include <QPushButton>
#include <QLabel>

namespace Ui {
class PptSharing;
}

typedef struct SelectedData
{
    int type;	// 0 for texts, 1 for pics
    int slideNum;
    int shapeNum;
}SelectedData;

class PptSharing : public QWidget
{
    Q_OBJECT

public:
    explicit PptSharing(QWidget *parent = nullptr);
    ~PptSharing();

protected:
    bool eventFilter(QObject *obj, QEvent *ev) override;

private slots:
    void onButtonOpenClicked();
    void onButtonSavePPTClicked();
    void onButtonSaveAsBraillePPTClicked();

    void onButtonAllTxtsBtnClicked();
    void onButtonSlcdTxtsBtnClicked();
    void allPicsLblClkDetect();
    void slcdPicsLblClkDetect();
    void allTxtsBtnDblClkDetect();
    void slcdTxtsBtnDblClkDetect();


    void queryCurSlideIndex();
    void catchException(int, const QString&, const QString&, const QString&);
private:
    Ui::PptSharing *ui;

    void exportTextPicsInPpt(int index);
    void enableQueryCurSlideIndex();
    void updateAllPlainTextsInSlide();
    void updateAllPicturesInSlide();
    void updateSelectedTextsInSlide();
    void updateSelectedPicturesInSlide();
    void updateBigViewPict();
    void updateBigViewText();
    void updateUi();

    QAxObject* ppApp;
    QAxObject* presentations;
    QAxObject* presentation;
    QAxObject* activeWindow;
    QAxObject* view;
    QAxObject* slide;

    QString pathname;
    QString dirname;
    QString pathnameNoExtension;
    QDateTime lastMdf;
    QTimer *curSldIdxTmr;

    QList<QString> textBlocks;
    QList<QPushButton*> txtBtnList;
    QList<QPushButton*> slcdTxtBtnList;

    QList<QImage> imgs;
    QList<QLabel*> imgLabels;
    QList<QLabel*> slcdImgLabels;

    QList<QImage> slidesThumb;
    QList<QLabel*> slideLabels;

    QList<SelectedData> slcdData;

    int slidesNum;
    int curSlideIndex;
    int curPicsIndex;
    int allPicsLblClkdNum = 0;
    int slcdPicsLblClkdNum = 0;
    int allTxtsBtnClkdNum = 0;
    int slcdTxtsBtnClkdNum = 0;
};

#endif // PPTSHARING_H
