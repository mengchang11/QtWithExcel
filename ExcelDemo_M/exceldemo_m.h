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
    // 初始化Excel界面
    void InitExcelWidget();

    // 加载Excel文件
    void Load();

    // 初始化资源
    void InitResource();

private:

    // Ui界面布局
    QVBoxLayout *m_mainWidgetVBoxLayout;

    QWidget *m_topWidgetInMainWidget;           // 用于放筛选条件
    QTabWidget *m_middleTabWidgetInMainWidget;  // 显示表格内容，以及筛选后的结构
    QWidget *m_lowWidgetInMainWidget;           // 图标显示

    // 菜单栏
    QAction *m_loadExcelFile;       // 加载Excel文件动作
    QAction *m_saveTableViewData;   // 保存数据到Excel文件内动作

    // 表格文件数据
    QString m_excelFilePath;                            // 表格的绝对路径
    QString m_excelFilePathBck;                         // 表格的绝对路径备份
    QList<QTableView*> m_tableViewList;                 // 堆上表格视图存储
    QList<QStandardItemModel*> m_excelDataModelList;    // 堆上表格数据模型存储
    QList<QStringList> m_sheetHeadStringTableList;      // 表头字符串存储

    // 缓存修改的单元格的值，在按save后，统一写入Excel文件
    QMap<int, QMap<QPair<int, int>, QVariant>> m_changedValueCache; // 修改的数据缓存

public slots:

    // 加载Excel文件数据
    void on_loadFileAction_clicked();

    // 保存数据到Excel文件
    void on_saveAction_clicked();

    // 单元格数据改变处理
    void tableCellDataChanged(const QModelIndex& topLeft, const QModelIndex& bottomRight, const QVector<int>& roles);
};
#endif // EXCELDEMO_M_H
