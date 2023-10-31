#ifndef EXCELDEMO_M_H
#define EXCELDEMO_M_H

#include <QMainWindow>

#include <QVBoxLayout>
#include <QTabWidget>
#include <QAction>
#include <QFileDialog>
#include <QAxObject>
#include <QTableView>
#include <QStandardItemModel>

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

private:
    void InitExcelWidget();
    void Load();
    void InitResource();

private:

    // Ui界面布局
    QVBoxLayout *m_mainWidgetVBoxLayout;

    QWidget *m_topWidgetInMainWidget;
    QTabWidget *m_middleTabWidgetInMainWidget;
    QWidget *m_lowWidgetInMainWidget;

    // 菜单栏
    QAction *m_loadExcelFile;
    QAction *m_saveTableViewData;

    // 表格文件数据
    QString m_excelFilePath;
    QList<QTableView*> m_tableViewList;
    QList<QStandardItemModel*> m_excelDataModelList;
    QList<QStringList> m_sheetHeadStringTableList;

public slots:
    void LoadExcelFile();
};
#endif // EXCELDEMO_M_H
