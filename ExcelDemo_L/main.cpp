#include "exceldemo_l.h"

#include <QApplication>

int main(int argc, char *argv[])
{
    QApplication a(argc, argv);
    ExcelDemo_L w;
    w.show();
    return a.exec();
}
