#ifndef EXCELDEMO_L_H
#define EXCELDEMO_L_H

#include <QMainWindow>

QT_BEGIN_NAMESPACE
namespace Ui { class ExcelDemo_L; }
QT_END_NAMESPACE

class ExcelDemo_L : public QMainWindow
{
    Q_OBJECT

public:
    ExcelDemo_L(QWidget *parent = nullptr);
    ~ExcelDemo_L();

private:
    Ui::ExcelDemo_L *ui;
};
#endif // EXCELDEMO_L_H
