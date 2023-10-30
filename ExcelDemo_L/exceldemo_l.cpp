#include "exceldemo_l.h"
#include "ui_exceldemo_l.h"

ExcelDemo_L::ExcelDemo_L(QWidget *parent)
    : QMainWindow(parent)
    , ui(new Ui::ExcelDemo_L)
{
    ui->setupUi(this);
}

ExcelDemo_L::~ExcelDemo_L()
{
    delete ui;
}

