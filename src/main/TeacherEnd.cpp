#include "TeacherEnd.h"
#include "PptSharing.h"


TeacherEnd::TeacherEnd(QWidget *parent)
    : QMainWindow(parent)
{
    ui.setupUi(this);

	connect(ui.buttonPPTShare, SIGNAL(clicked()), SLOT(onButtonPPTShareClicked()));
}

TeacherEnd::~TeacherEnd()
{}

void TeacherEnd::onButtonPPTShareClicked()
{
    this->hide();
    PptSharing* pptShWindow = new PptSharing(nullptr);
    pptShWindow->show();
}
