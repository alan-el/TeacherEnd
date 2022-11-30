#include "TeacherEnd.h"
#include <QtWidgets/QApplication>
#include <QAxObject>

//#include <opencv2/opencv.hpp>
//#include <opencv2/aruco.hpp>

//using namespace cv;

int main(int argc, char *argv[])
{
    QApplication a(argc, argv);
    TeacherEnd w;
    w.show();
    /*QString pptFile = "D:/MyProject/BlinderReader/test.ppt";
	
	QAxObject pp("PowerPoint.Application", nullptr);
	auto presentations = pp.querySubObject("Presentations");
	auto presentation = presentations->querySubObject("Open(QString)", pptFile);
	delete presentation;
	delete presentations;
	pp.dynamicCall("Quit()");

    Mat markerImage;
    Ptr<aruco::Dictionary> dict = aruco::getPredefinedDictionary(aruco::DICT_7X7_250);
    aruco::drawMarker(dict, 33, 200, markerImage, 1);

    imshow("marker33", markerImage);*/
    return a.exec();


}
