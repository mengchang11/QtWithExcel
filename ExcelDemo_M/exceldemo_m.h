#ifndef EXCELDEMO_M_H
#define EXCELDEMO_M_H

#include <QMainWindow>

QT_BEGIN_NAMESPACE
namespace Ui { class ExcelDemo_M; }
QT_END_NAMESPACE

class ExcelDemo_M : public QMainWindow
{
    Q_OBJECT

public:
    ExcelDemo_M(QWidget *parent = nullptr);
    ~ExcelDemo_M();

private:
    Ui::ExcelDemo_M *ui;
};
#endif // EXCELDEMO_M_H
