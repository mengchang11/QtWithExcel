#include "exceldemo_m.h"
#include "ui_exceldemo_m.h"

ExcelDemo_M::ExcelDemo_M(QWidget *parent)
    : QMainWindow(parent)
    , ui(new Ui::ExcelDemo_M)
{
    ui->setupUi(this);
}

ExcelDemo_M::~ExcelDemo_M()
{
    delete ui;
}

