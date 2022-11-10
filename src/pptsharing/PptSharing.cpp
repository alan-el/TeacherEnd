#include <QImageReader>
#include <QColorSpace>
#include <QFileDialog>
#include <QFileInfo>
#include <QMessageBox>
#include <QStringDecoder>
#include <QEvent>
#include <QMouseEvent>
#include <QTimer>
#include "JniMethod.h"
#include "PptSharing.h"
#include "ui_PptSharing.h"

PptSharing::PptSharing(QWidget *parent) :
    QWidget(parent),
    ui(new Ui::PptSharing)
{
    ui->setupUi(this);
    connect(ui->buttonOpen, SIGNAL(clicked()), SLOT(onButtonOpenClicked()));
    connect(ui->buttonSavePPT, SIGNAL(clicked()), SLOT(onButtonSavePPTClicked()));
    connect(ui->buttonSaveAsBraillePPT, SIGNAL(clicked()), SLOT(onButtonSaveAsBraillePPTClicked()));

}

PptSharing::~PptSharing()
{
    delete ui;
}

bool PptSharing::eventFilter(QObject * obj, QEvent * event)
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
                    for(int j = 0; j < slcdData.size(); j++)
                    {
                        if(slcdData.at(j).type == 1 && slcdData.at(j).slideNum == curSlideIndex)
                            index++;

                        if(index == i + 1)
                        {
                            slcdPicsLblClkdNum = 1;
                            curPicsIndex = slcdData.at(j).shapeNum;
                            QTimer::singleShot(500, this, SLOT(slcdPicsLblClkDetect()));
                            break;
                        }
                    }
                    qDebug() << "Label " << (i + 1) << "clicked.";
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
                    SelectedData sd = {1, curSlideIndex, i + 1};
                    bool hasContained = false;
                    for(SelectedData s : slcdData)
                    {
                        if(sd.type == s.type && sd.slideNum == s.slideNum && sd.shapeNum == s.shapeNum)
                        {
                            hasContained = true;
                            break;
                        }
                    }
                    if(!hasContained)
                        slcdData.append(sd);

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
                    for(int j = 0; j < slcdData.size(); j++)
                    {
                        if(slcdData.at(j).type == 1 && slcdData.at(j).slideNum == curSlideIndex)
                            index++;

                        if(index == i + 1)
                        {
                            slcdData.removeAt(j);
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

void PptSharing::onButtonOpenClicked()
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

    connect(ppApp, SIGNAL(exception(int, const QString&, const QString&, const QString&)), this, SLOT(catchException(int, const QString&, const QString&, const QString&)));
    connect(presentations, SIGNAL(exception(int, const QString&, const QString&, const QString&)), this, SLOT(catchException(int, const QString&, const QString&, const QString&)));
    connect(presentation, SIGNAL(exception(int, const QString&, const QString&, const QString&)), this, SLOT(catchException(int, const QString&, const QString&, const QString&)));

    curSlideIndex = curPicsIndex = 1;
    QFileInfo fi(pathname);
    lastMdf = fi.lastModified();
    updateUi();
    enableQueryCurSlideIndex();
}

void PptSharing::exportTextPicsInPpt(int index)
{
    /* Call Java method to Extract Texts in PPT */
    int str_length = pathname.length();
    std::string str = pathname.toStdString();
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

void PptSharing::enableQueryCurSlideIndex()
{
    curSldIdxTmr = new QTimer(this);
    connect(curSldIdxTmr, &QTimer::timeout, this, &PptSharing::queryCurSlideIndex);
    curSldIdxTmr->start(1000);
}

void PptSharing::onButtonSavePPTClicked()
{
    presentation->dynamicCall("Save()");
}

void PptSharing::onButtonSaveAsBraillePPTClicked()
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

void PptSharing::onButtonAllTxtsBtnClicked()
{
    if(allTxtsBtnClkdNum == 0)
    {
        allTxtsBtnClkdNum++;
        QTimer::singleShot(500, this, SLOT(allTxtsBtnDblClkDetect()));
        return;
    }
    else if(allTxtsBtnClkdNum == 1)
    {
        allTxtsBtnClkdNum = 0;
    }

    QObject* obj = sender();
    for(int i = 0; i < txtBtnList.size(); i++)
    {
        if(obj == txtBtnList.at(i))
        {
            SelectedData sd = {0, curSlideIndex, i + 1};
            bool hasContained = false;
            for(SelectedData s : slcdData)
            {
                if(sd.type == s.type && sd.slideNum == s.slideNum && sd.shapeNum == s.shapeNum)
                {
                    hasContained = true;
                    break;
                }
            }
            if(!hasContained)
                slcdData.append(sd);

            qDebug() << "Button " << i + 1 << "clicked.";
            updateSelectedTextsInSlide();
            break;
        }
    }
}

void PptSharing::onButtonSlcdTxtsBtnClicked()
{
    QObject* obj = sender();

    for(int i = 0; i < slcdTxtBtnList.size(); i++)
    {
        if(obj == slcdTxtBtnList.at(i))
        {
            int index = 0;
            for(int j = 0; j < slcdData.size(); j++)
            {
                if(slcdData.at(j).type == 0 && slcdData.at(j).slideNum == curSlideIndex)
                    index++;

                if(index == i + 1)
                {
                    slcdData.removeAt(j);
                    break;
                }
            }
            updateSelectedTextsInSlide();
            break;
        }
    }
}

void PptSharing::allPicsLblClkDetect()
{
    if(allPicsLblClkdNum == 1)
    {
        updateBigViewPict();
    }
}

void PptSharing::slcdPicsLblClkDetect()
{
    if(slcdPicsLblClkdNum == 1)
    {
        updateBigViewPict();
    }
}

void PptSharing::allTxtsBtnDblClkDetect()
{
    allTxtsBtnClkdNum = 0;
}

void PptSharing::slcdTxtsBtnDblClkDetect()
{
}

void PptSharing::queryCurSlideIndex()
{
    presentation->dynamicCall("Save()");
    activeWindow = ppApp->querySubObject("ActiveWindow");
    connect(activeWindow, SIGNAL(exception(int, const QString&, const QString&, const QString&)), this, SLOT(catchException(int, const QString&, const QString&, const QString&)));

    if(activeWindow != nullptr)
    {
        view = activeWindow->querySubObject("View");
        connect(view, SIGNAL(exception(int, const QString&, const QString&, const QString&)), this, SLOT(catchException(int, const QString&, const QString&, const QString&)));
        slide = view->querySubObject("slide");
        connect(slide, SIGNAL(exception(int, const QString&, const QString&, const QString&)), this, SLOT(catchException(int, const QString&, const QString&, const QString&)));

        int slideIndex = slide->property("SlideIndex").toInt();

        // 当前幻灯片页改变了，需要更新 Qt ui。
        if(slideIndex != curSlideIndex)
        {
            qDebug() << "curSlideIndex Changed: " << slideIndex;
            curSlideIndex = slideIndex;
            curPicsIndex = 1;
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
    else
    {
        QMessageBox::warning(this, "提示", "PPT 软件已关闭!");
        curSldIdxTmr->stop();
    }


    delete activeWindow;
}

void PptSharing::catchException(int code, const QString &source, const QString &disc, const QString &help)
{
    qDebug() << "code: " << code;
    qDebug() << "source: " << source;
    qDebug() << "disc: " << disc;
    qDebug() << "help: " << help;
}

void PptSharing::updateAllPlainTextsInSlide()
{
    textBlocks.clear();
    qDeleteAll(txtBtnList);
    txtBtnList.clear();

    QString txtPtnPrf = pathnameNoExtension + "/texts/slide" + QString::number(curSlideIndex);
    QFile plainTextFile;

    for(int i = 1; ; i++)
    {
        plainTextFile.setFileName(txtPtnPrf + "_text" + QString::number(i) + ".txt");

        if(plainTextFile.exists())
        {
            bool ret = plainTextFile.open(QIODevice::ReadOnly);
            if(ret == true)
            {
                QByteArray array = plainTextFile.readAll();
                //QTextCodec* codec = QTextCodec::codecForName("GBK");
                QStringDecoder toUtf16 = QStringDecoder(QStringDecoder::System);
                //QString str = codec->toUnicode(array);
                QString str = toUtf16.decode(array);
                textBlocks.append(str);

                QPushButton* pBtn = new QPushButton();
                txtBtnList.append(pBtn);
            }
            plainTextFile.close();
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

void PptSharing::updateAllPicturesInSlide()
{
    imgs.clear();
    qDeleteAll(imgLabels);
    imgLabels.clear();

    QString picPtnPrf = pathnameNoExtension + "/pictures/slide" + QString::number(curSlideIndex);

    for(int i = 1; ; i++)
    {
        QString imgFile = picPtnPrf + "_pict" + QString::number(i);
        QImageReader reader(imgFile);
        reader.setScaledSize(QSize(80, 80));
        QImage img = reader.read();
        if(img.isNull())
        {
            break;
        }
        if(img.colorSpace().isValid())
            img.convertToColorSpace(QColorSpace::SRgb);

        imgs.append(img);
        QLabel* pLbl = new QLabel();
        imgLabels.append(pLbl);
        pLbl->setPixmap(QPixmap::fromImage(img));
        pLbl->adjustSize();

        ui->horizontalLayoutAllPics->addWidget(pLbl);
        imgLabels.last()->installEventFilter(this);
    }
}

void PptSharing::updateSelectedTextsInSlide()
{
    qDeleteAll(slcdTxtBtnList);
    slcdTxtBtnList.clear();

    for(SelectedData s : slcdData)
    {
        if((s.type == 0) && (s.slideNum == curSlideIndex))
        {
            QPushButton* pBtn = new QPushButton();
            slcdTxtBtnList.append(pBtn);

            if(textBlocks.at(s.shapeNum - 1).size() > 3)
                slcdTxtBtnList.last()->setText(textBlocks.at(s.shapeNum - 1).left(3) + "..");
            else
                slcdTxtBtnList.last()->setText(textBlocks.at(s.shapeNum - 1));

            slcdTxtBtnList.last()->setFixedSize(70, 28);

            ui->gridLayoutSelectedTexts->addWidget(slcdTxtBtnList.last(), (slcdTxtBtnList.size() - 1) / 7, (slcdTxtBtnList.size() - 1) % 7);

            connect(slcdTxtBtnList.last(), SIGNAL(clicked()), SLOT(onButtonSlcdTxtsBtnClicked()));
            slcdTxtBtnList.last()->show();
        }
    }
}

void PptSharing::updateSelectedPicturesInSlide()
{
    // TODO md5 标识
    qDeleteAll(slcdImgLabels);
    slcdImgLabels.clear();

    for(SelectedData s : slcdData)
    {
        if((s.type == 1) && (s.slideNum == curSlideIndex))
        {
            QLabel* pLbl = new QLabel();
            slcdImgLabels.append(pLbl);
            pLbl->setPixmap(QPixmap::fromImage(imgs.at(s.shapeNum - 1)));
            pLbl->adjustSize();

            ui->horizontalLayoutSelectedPics->addWidget(pLbl);
            slcdImgLabels.last()->installEventFilter(this);
        }
    }
}

void PptSharing::updateBigViewPict()
{
    QString picPtnPrf = pathnameNoExtension + "/pictures/slide" + QString::number(curSlideIndex);
    QString imgFile = picPtnPrf + "_pict" + QString::number(curPicsIndex);

    QImageReader reader(imgFile);
    reader.setScaledSize(QSize(300, 180));
    QImage img = reader.read();
    if(img.isNull())
    {
        QMessageBox::information(this, QGuiApplication::applicationDisplayName(),
            tr("Cannot load %1: %2")
            .arg(QDir::toNativeSeparators(imgFile), reader.errorString()));
    }
    if(img.colorSpace().isValid())
        img.convertToColorSpace(QColorSpace::SRgb);

    ui->imageLabel->setPixmap(QPixmap::fromImage(img));
}

void PptSharing::updateBigViewText()
{
}

void PptSharing::updateUi()
{
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


