#include "exceldemo_m.h"

#include <QApplication>

int main(int argc, char *argv[])
{
    QApplication a(argc, argv);
    ExcelDemo_M w;
    w.show();
    return a.exec();
}
