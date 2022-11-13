#include <QImageReader>
#include <QColorSpace>
#include <QFileDialog>
#include <QFileInfo>
#include <QMessageBox>
#include <QStringDecoder>
#include <QEvent>
#include <QMouseEvent>
#include <QTimer>
#include <QCryptographicHash>
#include "JniMethod.h"
#include "BrailleTranslator.h"
#include "PptSharingManager.h"
#include "ui_PptSharing.h"

PptSharingManager::PptSharingManager(QWidget *parent) :
    QWidget(parent),
    ui(new Ui::PptSharing)
{
    ui->setupUi(this);
//    QIcon* icon = new QIcon(":res/assets/play_40.jpg");
//    ui->buttonOpen->setIcon(*icon);

//    PptShape a{PptShape::Pictures, "D:/Qt_workspace/src/TeacherEnd/doc/test/pictures/slide6_pict1.jpg"};
//    PptShape b{PptShape::Pictures, "D:/Qt_workspace/src/TeacherEnd/doc/test/pictures/slide6_pict1_small.jpg"};
//    qDebug() << "origin: " << a.getMd5Hash();
//    qDebug() << "alias: " << b.getMd5Hash();

    QString plain{"一二三四五"};
    QString brl{BrailleTranslator::brlTranslate(plain)};
    qDebug() << brl;

    connect(ui->buttonOpen, SIGNAL(clicked()), SLOT(onButtonOpenClicked()));
    connect(ui->buttonPlay, SIGNAL(clicked()), SLOT(onButtonPlayClicked()));
    connect(ui->buttonSavePPT, SIGNAL(clicked()), SLOT(onButtonSavePPTClicked()));
    connect(ui->buttonSaveAsBraillePPT, SIGNAL(clicked()), SLOT(onButtonSaveAsBraillePPTClicked()));

    ui->buttonPlay->setVisible(false);
    ui->buttonSavePPT->setEnabled(false);
    ui->buttonSaveAsBraillePPT->setEnabled(false);

    isSlideShowRunning = false;
}

PptSharingManager::~PptSharingManager()
{
    delete ui;
}

bool PptSharingManager::eventFilter(QObject * obj, QEvent * event)
{
    for(int i = 0; i < imgLabels.size(); i++)
    {
        if(obj == imgLabels.at(i))
        {
            if(event->type() == QEvent::MouseButtonPress)
            {

                QMouseEvent *mouseEvent = static_cast<QMouseEvent*>(event);
                if(mouseEvent->button() == Qt::LeftButton)
                {
                    allPicsLblClkdNum = 1;
                    curPicsIndex = i + 1;
                    QTimer::singleShot(500, this, SLOT(allPicsLblClkDetect()));

                    qDebug() << "Label " << (i + 1) << " clicked.";
                    return true;
                }
            }
        }
    }

    for(int i = 0; i < slcdImgLabels.size(); i++)
    {
        if(obj == slcdImgLabels.at(i))
        {
            if(event->type() == QEvent::MouseButtonPress)
            {
                QMouseEvent *mouseEvent = static_cast<QMouseEvent*>(event);
                if(mouseEvent->button() == Qt::LeftButton)
                {
                    int index = 0;
                    for(int j = 0; j < allPptPictShapes.size(); j++)
                    {
                        if(allPptPictShapes[j].getShapeIndexInSlide() > 0 && allPptPictShapes[j].getIsSelected())
                            index++;

                        if(index == i + 1)
                        {
                            slcdPicsLblClkdNum = 1;
                            curPicsIndex = allPptPictShapes[j].getShapeIndexInSlide();
                            QTimer::singleShot(500, this, SLOT(slcdPicsLblClkDetect()));
                            break;
                        }
                    }
                    qDebug() << "Label " << (i + 1) << "clicked.";
                    qDebug() << "Current Picture Index: " << curPicsIndex;
                    return true;
                }
            }
        }
    }

    for(int i = 0; i < imgLabels.size(); i++)
    {
        if(obj == imgLabels.at(i))
        {
            if(event->type() == QEvent::MouseButtonDblClick)
            {
                QMouseEvent *mouseEvent = static_cast<QMouseEvent*>(event);
                if(mouseEvent->button() == Qt::LeftButton)
                {
                    allPicsLblClkdNum = 2;
                    for(int j = 0; j < allPptPictShapes.size(); j++)
                    {
                        if(allPptPictShapes[j].getShapeIndexInSlide() == i + 1)
                        {
                            allPptPictShapes[j].setIsSelected(true);
                            break;
                        }
                    }

                    qDebug() << "Label " << (i + 1) << "double clicked.";
                    updateSelectedPicturesInSlide();
                    return true;
                }
            }
        }
    }

    for(int i = 0; i < slcdImgLabels.size(); i++)
    {
        if(obj == slcdImgLabels.at(i))
        {
            if(event->type() == QEvent::MouseButtonDblClick)
            {
                QMouseEvent *mouseEvent = static_cast<QMouseEvent*>(event);
                if(mouseEvent->button() == Qt::LeftButton)
                {
                    int index = 0;
                    for(int j = 0; j < allPptPictShapes.size(); j++)
                    {
                        if(allPptPictShapes[j].getShapeIndexInSlide() > 0 && allPptPictShapes[j].getIsSelected())
                            index++;

                        if(index == i + 1)
                        {
                            allPptPictShapes[j].setIsSelected(false);
                            break;
                        }
                    }

                    qDebug() << "Label " << (i + 1) << "double clicked.";
                    slcdPicsLblClkdNum = 2;
                    updateSelectedPicturesInSlide();
                    return true;
                }
            }
        }
    }

    // pass the event on to the parent class
    return QWidget::eventFilter(obj, event);
}

void PptSharingManager::onButtonOpenClicked()
{
    pathname = QFileDialog::getOpenFileName(0, tr("Open PPTX File"), "D:/Qt_workspace/src/TeacherEnd/doc", "Presentation files (*.pptx);;All files(*.*)");
    if(pathname.isEmpty())
        return;

    for(int i = pathname.length() - 1; i >= 0; i--)
    {
        if(pathname.at(i) == '.')
        {
            pathnameNoExtension = pathname.left(i);
        }

        if(pathname.at(i) == '/')
        {
            dirname = pathname.left(i);
            break;
        }
    }

    pathname.replace('/', '\\');

    ppApp = new QAxObject("PowerPoint.Application", nullptr);
    presentations = ppApp->querySubObject("Presentations");
    presentation = presentations->querySubObject("Open(QString)", pathname);

    connect(ppApp, SIGNAL(exception(int, QString, QString, QString)), this, SLOT(catchException(int, QString, QString, QString)));
    connect(presentations, SIGNAL(exception(int, QString, QString, QString)), this, SLOT(catchException(int, QString, QString, QString)));
    connect(presentation, SIGNAL(exception(int, QString, QString, QString)), this, SLOT(catchException(int, QString, QString, QString)));

    curSlideIndex = curPicsIndex = curTxtsIndex = 1;
    QFileInfo fi(pathname);
    lastMdf = fi.lastModified();
    updateUi();
    enableQueryCurSlideIndex();

    ui->buttonPlay->setVisible(true);
}

void PptSharingManager::onButtonPlayClicked()
{
    slideShowSettings = presentation->querySubObject("SlideShowSettings");
    connect(slideShowSettings, SIGNAL(exception(int, QString, QString, QString)), this, SLOT(catchException(int, QString, QString, QString)));
    slideShowSettings->dynamicCall("Run()");
    slideShowWindow = presentation->querySubObject("SlideShowWindow");
    connect(slideShowWindow, SIGNAL(exception(int, QString, QString, QString)), this, SLOT(catchException(int, QString, QString, QString)));
    slideShowView = slideShowWindow->querySubObject("View");
    connect(slideShowView, SIGNAL(exception(int, QString, QString, QString)), this, SLOT(catchException(int, QString, QString, QString)));

    isSlideShowRunning = true;
}

void PptSharingManager::exportTextPicsInPpt(int index)
{
    /* Call Java method to Extract Texts in PPT */
    std::string str = pathname.toStdString();
    int str_length = str.length();
    const char *ch = str.c_str();
    const char *tail = (ch + str_length - 1) - 4;

    if(strcmp(tail + 1, ".ppt") == 0)
    {
        JniMethod::createJVM();
        slidesNum = JniMethod::pptTextExtractor(ch, false, index);
        JniMethod::pptPictExtractor(ch, false, index);
    }
    else if(strcmp(tail, ".pptx") == 0)
    {
        JniMethod::createJVM();
        slidesNum = JniMethod::pptTextExtractor(ch, true, index);
        JniMethod::pptPictExtractor(ch, true, index);
    }
}

void PptSharingManager::enableQueryCurSlideIndex()
{
    curSldIdxTmr = new QTimer(this);
    connect(curSldIdxTmr, &QTimer::timeout, this, &PptSharingManager::queryCurSlideIndex);
    curSldIdxTmr->start(1000);
}

void PptSharingManager::onButtonSavePPTClicked()
{
    presentation->dynamicCall("Save()");
}

void PptSharingManager::onButtonSaveAsBraillePPTClicked()
{
    QString bpptfile = pathnameNoExtension + ".bppt";
    QString dir;
    QString fn;

    do
    {
        fn = QFileDialog::getSaveFileName(0, tr("Save BPPT File"), bpptfile, "Braille PPT files (*.bppt);;All files(*.*)");

        for(int i = fn.length() - 1; i >= 0; i--)
        {
            if(fn.at(i) == '/')
            {
                dir = fn.left(i);
                break;
            }
        }

        if(dir.compare(dirname) == 0 || fn.isEmpty())
            break;

        QMessageBox::warning(this, "提示", "请保存在原PPT文件相同目录下");
    } while(1);

}

void PptSharingManager::onButtonAllTxtsBtnClicked()
{
    QObject* obj = sender();
    // 无论单击 or 双击，更新当前选中的文本
    for(int i = 0; i < txtBtnList.size(); i++)
    {
        if(obj == txtBtnList.at(i))
        {
            curTxtsIndex = i + 1;
        }
    }

    if(allTxtsBtnClkdNum == 0)
    {
        allTxtsBtnClkdNum++;
        QTimer::singleShot(500, this, SLOT(allTxtsBtnDblClkDetect()));
        return;
    }
    else if(allTxtsBtnClkdNum == 1)
    {
        allTxtsBtnClkdNum++;
    }

    for(int i = 0; i < txtBtnList.size(); i++)
    {
        if(obj == txtBtnList.at(i))
        {
            for(int j = 0; j < allPptTextShapes.size(); j++)
            {
                if(allPptTextShapes[j].getShapeIndexInSlide() == i + 1)
                {
                    allPptTextShapes[j].setIsSelected(true);
                    break;
                }
            }
            qDebug() << "Button " << i + 1 << "clicked.";
            updateSelectedTextsInSlide();
            allTxtsBtnClkdNum = 0;
            break;
        }
    }
}

void PptSharingManager::onButtonSlcdTxtsBtnClicked()
{
    if(slcdTxtsBtnClkdNum == 0)
    {
        slcdTxtsBtnClkdNum++;
        QTimer::singleShot(500, this, SLOT(slcdTxtsBtnDblClkDetect()));
    }
    else if(slcdTxtsBtnClkdNum == 1)
    {
        slcdTxtsBtnClkdNum++;
    }

    QObject* obj = QObject::sender();
    for(int i = 0; i < slcdTxtBtnList.size(); i++)
    {
        if(obj == slcdTxtBtnList.at(i))
        {
            int index = 0;
            for(int j = 0; j < allPptTextShapes.size(); j++)
            {
                if(allPptTextShapes[j].getShapeIndexInSlide() > 0 && allPptTextShapes[j].getIsSelected())
                    index++;

                if(index == i + 1)
                {
                    curTxtsIndex = allPptTextShapes[j].getShapeIndexInSlide();
                    if(slcdTxtsBtnClkdNum == 2)
                        allPptTextShapes[j].setIsSelected(false);
                    break;
                }
            }
            if(slcdTxtsBtnClkdNum == 2)
            {
                updateSelectedTextsInSlide();
                slcdTxtsBtnClkdNum = 0;
            }
            break;
        }
    }
}

void PptSharingManager::allPicsLblClkDetect()
{
    if(allPicsLblClkdNum == 1)
    {
        updateBigViewPict();
    }
}

void PptSharingManager::slcdPicsLblClkDetect()
{
    if(slcdPicsLblClkdNum == 1)
    {
        updateBigViewPict();
    }
}

void PptSharingManager::allTxtsBtnDblClkDetect()
{
    if(allTxtsBtnClkdNum == 1)
    {
        updateBigViewText();
        allTxtsBtnClkdNum = 0;
    }

}

void PptSharingManager::slcdTxtsBtnDblClkDetect()
{
    if(slcdTxtsBtnClkdNum == 1)
    {
        updateBigViewText();
        slcdTxtsBtnClkdNum = 0;
    }

}

void PptSharingManager::queryCurSlideIndex()
{
    int slideIndex = 0;
    // 编辑模式
    if(isSlideShowRunning == false)
    {
        activeWindow = ppApp->querySubObject("ActiveWindow");

        presentation->dynamicCall("Save()");
        connect(activeWindow, SIGNAL(exception(int, QString, QString, QString)), this, SLOT(catchException(int, QString, QString, QString)));

        if(activeWindow != nullptr)
        {

            view = activeWindow->querySubObject("View");
            if(view != nullptr)
            {
                connect(view, SIGNAL(exception(int, QString, QString, QString)), this, SLOT(catchException(int, QString, QString, QString)));
                slide = view->querySubObject("slide");
                if(slide != nullptr)
                {
                    connect(slide, SIGNAL(exception(int, QString, QString, QString)), this, SLOT(catchException(int, QString, QString, QString)));

                    slideIndex = slide->property("SlideIndex").toInt();
                }
            }
        }
        else
        {
            //QMessageBox::warning(this, "提示", "PPT 软件已关闭!");
            //curSldIdxTmr->stop();
        }

        delete activeWindow;
    }
    // 幻灯片放映模式
    else
    {
        slideIndex = slideShowView->property("CurrentShowPosition").toLongLong();
    }


    // 当前幻灯片页改变了，需要更新 Qt ui。
    if(slideIndex > 0 && slideIndex != curSlideIndex)
    {
        qDebug() << "curSlideIndex Changed: " << slideIndex;
        curSlideIndex = slideIndex;
        curPicsIndex = curTxtsIndex = 1;
        updateUi();
    }
    else
    {
        QFileInfo fi(pathname);
        QDateTime lastMdfNew = fi.lastModified();
        // 当前幻灯片页未改变但是内容做出了修改，需要更新 Qt ui。
        if(lastMdf < lastMdfNew)
        {
            updateUi();
            lastMdf = lastMdfNew;
        }
    }
}

void PptSharingManager::catchException(int code, const QString &source, const QString &disc, const QString &help)
{
    qDebug() << "code: " << code;
    qDebug() << "source: " << source;
    qDebug() << "disc: " << disc;
    qDebug() << "help: " << help;

    if(disc.compare("SlideShowView.CurrentShowPosition : Object does not exist.") == 0)
    {
        qDebug() << "Slide Show has QUIT.";
        isSlideShowRunning = false;
    }
}

void PptSharingManager::updateAllPlainTextsInSlide()
{
    textBlocks.clear();
    qDeleteAll(txtBtnList);
    txtBtnList.clear();
    for(int i = 0; i < allPptTextShapes.size(); i++)
        allPptTextShapes[i].setShapeIndexInSlide(0);

    QString txtPtnPrf = pathnameNoExtension + "/texts/slide" + QString::number(curSlideIndex);
    QFile plainTextFile;

    for(int i = 1; ; i++)
    {
        plainTextFile.setFileName(txtPtnPrf + "_text" + QString::number(i) + ".txt");

        if(plainTextFile.exists())
        {
            // 更新全部 TextShape 对象
            PptTextShape psNew{plainTextFile.fileName()};
            bool hasContained = false;
            for(int j = 0; j < allPptTextShapes.size(); j++)
            {
                if(allPptTextShapes[j].getMd5Hash().compare(psNew.getMd5Hash()) == 0)
                {
                    hasContained = true;
                    allPptTextShapes[j].setShapeIndexInSlide(i);
                    break;
                }
            }
            if(!hasContained)
            {
                allPptTextShapes.append(psNew);
                allPptTextShapes.last().setShapeIndexInSlide(i);
            }

            // 更新全部 TextShape 对象UI
            bool ret = plainTextFile.open(QIODevice::ReadOnly);
            if(ret == true)
            {
                QByteArray array = plainTextFile.readAll();
                QStringDecoder toUtf16 = QStringDecoder(QStringDecoder::System);
                QString str = toUtf16.decode(array);
                textBlocks.append(str);

                QPushButton* pBtn = new QPushButton();
                txtBtnList.append(pBtn);

                plainTextFile.close();
            }
        }
        else
            break;
    }

    for(int i = 1; i <= txtBtnList.size(); i++)
    {
        int firstTxtNum = 3;

        if(textBlocks.at(i - 1).size() > firstTxtNum)
            txtBtnList.at(i - 1)->setText(textBlocks.at(i - 1).left(firstTxtNum) + "..");
        else
            txtBtnList.at(i - 1)->setText(textBlocks.at(i - 1));

        txtBtnList.at(i - 1)->setFixedSize(70, 28);

        ui->gridLayoutAllTexts->addWidget(txtBtnList.at(i - 1), (i - 1) / 7, (i - 1) % 7);

        connect(txtBtnList.at(i - 1), SIGNAL(clicked()), SLOT(onButtonAllTxtsBtnClicked()));
        txtBtnList.at(i - 1)->show();
    }
}

void PptSharingManager::updateAllPicturesInSlide()
{
    imgs.clear();
    qDeleteAll(imgLabels);
    imgLabels.clear();
    for(int i = 0; i < allPptPictShapes.size(); i++)
        allPptPictShapes[i].setShapeIndexInSlide(0);

    QString picPtnPrf = pathnameNoExtension + "/pictures/slide" + QString::number(curSlideIndex);

    for(int i = 1; ; i++)
    {
        QString imgFile = picPtnPrf + "_pict" + QString::number(i);
        // 更新全部 PictureShape 对象
        // TODO Encapsulation into method
        QString fullPathname{""};
        if(QFileInfo::exists(imgFile + ".jpg"))
        {
            fullPathname = imgFile + ".jpg";
        }
        else if(QFileInfo::exists(imgFile + ".png"))
        {
            fullPathname = imgFile + ".png";
        }

        if(!fullPathname.isEmpty())
        {
            PptPictureShape psNew{fullPathname};
            bool hasContained = false;
            for(int j = 0; j < allPptPictShapes.size(); j++)
            {
                if(allPptPictShapes[j].getMd5Hash().compare(psNew.getMd5Hash()) == 0)
                {
                    hasContained = true;
                    allPptPictShapes[j].setShapeIndexInSlide(i);
                    break;
                }
            }
            if(!hasContained)
            {
                allPptPictShapes.append(psNew);
                allPptPictShapes.last().setShapeIndexInSlide(i);
            }
        }

        // 更新全部 PictureShape 对象UI
        QImageReader reader(imgFile);
        reader.setScaledSize(QSize(80, 80));
        QImage img = reader.read();
        if(img.isNull())
        {
            break;
        }
        if(img.colorSpace().isValid())
            img.convertToColorSpace(QColorSpace::SRgb);

        // TODO Encapsulation into method
        imgs.append(img);
        QLabel* pLbl = new QLabel();
        imgLabels.append(pLbl);
        pLbl->setPixmap(QPixmap::fromImage(img));
        pLbl->adjustSize();

        ui->horizontalLayoutAllPics->addWidget(pLbl);
        imgLabels.last()->installEventFilter(this);
    }
}

void PptSharingManager::updateSelectedTextsInSlide()
{
    qDeleteAll(slcdTxtBtnList);
    slcdTxtBtnList.clear();

    for(PptTextShape textPsInAll: allPptTextShapes)
    {
        if(textPsInAll.getShapeIndexInSlide() > 0)
        {
            if(textPsInAll.getIsSelected())
            {
                QPushButton* pBtn = new QPushButton();
                slcdTxtBtnList.append(pBtn);

                QString brl = BrailleTranslator::brlTranslate(textBlocks[(textPsInAll.getShapeIndexInSlide() - 1)]);
                if(brl.size() > 3)
                    slcdTxtBtnList.last()->setText(brl.left(3) + "..");
                else
                    slcdTxtBtnList.last()->setText(brl);

                slcdTxtBtnList.last()->setFixedSize(70, 28);

                ui->gridLayoutSelectedTexts->addWidget(slcdTxtBtnList.last(), (slcdTxtBtnList.size() - 1) / 7, (slcdTxtBtnList.size() - 1) % 7);

                connect(slcdTxtBtnList.last(), SIGNAL(clicked()), SLOT(onButtonSlcdTxtsBtnClicked()));
                slcdTxtBtnList.last()->show();
            }
        }
    }
}

void PptSharingManager::updateSelectedPicturesInSlide()
{
    qDeleteAll(slcdImgLabels);
    slcdImgLabels.clear();

    for(PptPictureShape pictPsInAll: allPptPictShapes)
    {
        if(pictPsInAll.getShapeIndexInSlide() > 0)
        {
            if(pictPsInAll.getIsSelected())
            {
                QLabel* pLbl = new QLabel();
                slcdImgLabels.append(pLbl);
                pLbl->setPixmap(QPixmap::fromImage(imgs.at(pictPsInAll.getShapeIndexInSlide() - 1)));
                pLbl->adjustSize();

                ui->horizontalLayoutSelectedPics->addWidget(pLbl);
                slcdImgLabels.last()->installEventFilter(this);
            }
        }
    }
}

void PptSharingManager::updateBigViewPict()
{
    QString picPtnPrf = pathnameNoExtension + "/pictures/slide" + QString::number(curSlideIndex);
    QString imgFile = picPtnPrf + "_pict" + QString::number(curPicsIndex);

    QImageReader reader(imgFile);
    reader.setScaledSize(QSize(300, 180));
    QImage img = reader.read();
    if(img.isNull())
    {
        ui->imageLabel->setPixmap(QPixmap());
        return;
    }

    if(img.colorSpace().isValid())
        img.convertToColorSpace(QColorSpace::SRgb);

    ui->imageLabel->setPixmap(QPixmap::fromImage(img));
}

void PptSharingManager::updateBigViewText()
{
    if(textBlocks.size() > curTxtsIndex - 1)
    {
        QString brl = BrailleTranslator::brlTranslate(textBlocks[curTxtsIndex - 1]);
        ui->textEdit->setText(brl);
    }
}

void PptSharingManager::updateUi()
{
    if(curSlideIndex == 0)
        return;

    QDir dataDir(pathnameNoExtension);
    dataDir.removeRecursively();
    dataDir.removeRecursively();
    dataDir.removeRecursively();
    exportTextPicsInPpt(curSlideIndex);
    updateAllPicturesInSlide();
    updateBigViewPict();
    updateSelectedPicturesInSlide();
    updateAllPlainTextsInSlide();
    updateBigViewText();
    updateSelectedTextsInSlide();
}


