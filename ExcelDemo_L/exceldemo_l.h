#ifndef EXCELDEMO_L_H
#define EXCELDEMO_L_H

#include <QMainWindow>

#include <QString>
#include <QFileDialog>
#include <QAxObject>
#include <QVariant>
#include <QVariantList>

#include <QtDebug>

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

protected:
    void Init();
    // 新增表
    void AppendSheet();
    // 数据写入
    //bool write(const QString& filename, const QList<QList<QVariant>>& datas);

private:
    QAxObject *excel;
    QAxObject *workBooks;  // 工作簿集合
    QAxObject *workBook;  // 工作簿
    QAxObject *workSheets;  // 工作表集合
    QAxObject* sheet;  // 工作表
};
#endif // EXCELDEMO_L_H
