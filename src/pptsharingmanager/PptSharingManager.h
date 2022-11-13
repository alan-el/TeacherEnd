#ifndef PPTSHARINGMANAGER_H
#define PPTSHARINGMANAGER_H

#include <QWidget>
#include <QAxObject>
#include <QDateTime>
#include <QPushButton>
#include <QLabel>
#include "PptTextShape.h"
#include "PptPictureShape.h"

namespace Ui {
class PptSharing;
}

//typedef struct SelectedData
//{
//    int type;	// 0 for texts, 1 for pics
//    int slideNum;
//    int shapeNum;
//}SelectedData;

class PptSharingManager : public QWidget
{
    Q_OBJECT

public:
    explicit PptSharingManager(QWidget *parent = nullptr);
    ~PptSharingManager();

protected:
    bool eventFilter(QObject *obj, QEvent *ev) override;

private slots:
    void onButtonOpenClicked();
    void onButtonPlayClicked();
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
    QAxObject* slideShowSettings;
    QAxObject* slideShowWindow;
    QAxObject* slideShowView;

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

//    QList<SelectedData> slcdData;
    QList<PptPictureShape> allPptPictShapes;
    QList<PptTextShape> allPptTextShapes;
//    QList<PptShape> slcdPptShapes;

    int slidesNum;
    int curSlideIndex;
    int curPicsIndex;
    int curTxtsIndex;
    int allPicsLblClkdNum = 0;
    int slcdPicsLblClkdNum = 0;
    int allTxtsBtnClkdNum = 0;
    int slcdTxtsBtnClkdNum = 0;

    bool isSlideShowRunning;
};

#endif // PPTSHARINGMANAGER_H
