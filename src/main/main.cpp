#include "TeacherEnd.h"
#include <QtWidgets/QApplication>
#include <QAxObject>

int main(int argc, char *argv[])
{
    QApplication a(argc, argv);
    TeacherEnd w;
    w.show();
/*	QString pptFile = "D:/MyProject/BlinderReader/test.ppt";
	
	QAxObject pp("PowerPoint.Application", nullptr);
	auto presentations = pp.querySubObject("Presentations");
	auto presentation = presentations->querySubObject("Open(QString)", pptFile);
	delete presentation;
	delete presentations;
	pp.dynamicCall("Quit()");
*/
    return a.exec();
}
