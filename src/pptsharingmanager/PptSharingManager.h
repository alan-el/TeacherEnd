#ifndef PPTSHARINGMANAGER_H
#define PPTSHARINGMANAGER_H

#include <QWidget>
#include <QAxObject>
#include <QDateTime>
#include <QPushButton>
#include <QLabel>
#include "PptTextShape.h"
#include "PptPictureShape.h"
#include "DataSave.h"

namespace Ui {
class PptSharing;
}

class PptSharingManager : public QWidget
{
    Q_OBJECT

public:
    enum SaveShapeType {
        Text, Picture
    };

    explicit PptSharingManager(QWidget *parent = nullptr);
    ~PptSharingManager();

protected:
    bool eventFilter(QObject *obj, QEvent *ev) override;

private slots:
    void onButtonOpenClicked();
    void onButtonPlayClicked();
    void onButtonSaveClicked();

    void onButtonAllTxtsBtnClicked();
    void onButtonSlcdTxtsBtnClicked();

    void allPicsLblClkDetect();
    void slcdPicsLblClkDetect();

    void allTxtsBtnDblClkDetect();
    void slcdTxtsBtnDblClkDetect();

    void queryCurSlideIndex();

    void autoSaveData();

    void saveData();

    void catchException(int, const QString&, const QString&, const QString&);
private:
    Ui::PptSharing *ui;

    void exportTextPicsInPpt(int index);

    void enableQueryCurSlideIndex();

    void enableAutoSaveData();

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
    QTimer *autoSaveDataTmr;

    QList<QPushButton*> txtBtnList;
    QList<QPushButton*> slcdTxtBtnList;

    QList<QLabel*> imgLabels;
    QList<QLabel*> slcdImgLabels;

    QList<PptPictureShape> allPptPictShapes;
    QList<PptTextShape> allPptTextShapes;

    DataSave savedData;

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
