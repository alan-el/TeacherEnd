#pragma once

#include <QtWidgets/QMainWindow>
#include "ui_TeacherEnd.h"

class TeacherEnd : public QMainWindow
{
    Q_OBJECT

public:
    TeacherEnd(QWidget *parent = nullptr);
    ~TeacherEnd();

private slots:
	void onButtonPPTShareClicked();

private:
    Ui::TeacherEndClass ui;
};
