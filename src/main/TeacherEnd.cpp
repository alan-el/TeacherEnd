#include "TeacherEnd.h"
#include "PptSharingManager.h"


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
    PptSharingManager* pptShWindow = new PptSharingManager(nullptr);
    pptShWindow->show();
}
