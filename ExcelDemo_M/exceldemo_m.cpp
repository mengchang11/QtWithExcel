#include "exceldemo_m.h"
#include "ui_exceldemo_m.h"

#include <QDebug>

ExcelDemo_M::ExcelDemo_M(QWidget *parent)
    : QMainWindow(parent)
    , ui(new Ui::ExcelDemo_M)
{
    ui->setupUi(this);

    InitExcelWidget();
}

ExcelDemo_M::~ExcelDemo_M()
{
    delete ui;
}

void ExcelDemo_M::InitExcelWidget()
{
    // 新增widget，显示和选择筛选条件，下拉框和复选框，下拉框名字为当前行标题，复选框内容为当前展示的表格对于列的不重复内容
    m_mainWidgetVBoxLayout = new QVBoxLayout();
    m_topWidgetInMainWidget = new QTabWidget();
    m_mainWidgetVBoxLayout->addWidget(m_topWidgetInMainWidget);

    // 新增tabwidget，用于显示加载的表格内容或者筛选的内容
    m_middleTabWidgetInMainWidget = new QTabWidget();
    m_mainWidgetVBoxLayout->addWidget(m_middleTabWidgetInMainWidget);

    // 新增widget，用于显示折线图
    m_lowWidgetInMainWidget = new QTabWidget();
    m_mainWidgetVBoxLayout->addWidget(m_lowWidgetInMainWidget);

    // 设置组件占用mainwindow比例分别为1:2:1
    m_mainWidgetVBoxLayout->setStretch(0, 1);
    m_mainWidgetVBoxLayout->setStretch(1, 2);
    m_mainWidgetVBoxLayout->setStretch(2, 1);
    this->centralWidget()->setLayout(m_mainWidgetVBoxLayout);

    // 菜单设置
    QMenu *fileMenu = menuBar()->addMenu("文件");
    m_loadExcelFile = new QAction("加载Excel文件", this);
    fileMenu->addAction(m_loadExcelFile);
    connect(m_loadExcelFile, &QAction::triggered, this, &ExcelDemo_M::on_loadFileAction_clicked);

    m_saveTableViewData = new QAction("保存数据", this);
    fileMenu->addAction(m_saveTableViewData);
    connect(m_saveTableViewData, &QAction::triggered, this, &ExcelDemo_M::on_saveAction_clicked);
}


// 加载Excel文件
void ExcelDemo_M::on_loadFileAction_clicked()
{
    m_excelFilePath = QFileDialog::getOpenFileName(this, "打开Excel文件", "", "Excel文件 (*.xlsx *xls)");
    if (!m_excelFilePath.isEmpty()) {
        m_excelFilePathBck = m_excelFilePath;
        Load();
    } else {

        // 防止Save的时候缓存的Excel文件路径为空导致工作表打开异常
        m_excelFilePath = m_excelFilePathBck;
    }
}

void ExcelDemo_M::Load()
{
    QAxObject excelApp("Excel.Application");
    QAxObject *workBooks = excelApp.querySubObject("Workbooks");
    if (workBooks == nullptr) {
        qDebug() << "Workbooks error";
        return;
    }
    QAxObject *workBook = workBooks->querySubObject("Open(const QString)", m_excelFilePath);
    if (workBook == nullptr) {
        qDebug() << "Open(const QString) error";
        return;
    }

    // 获取工作表数量
    QAxObject *workSheets = workBook->querySubObject("Worksheets");
    if (workSheets == nullptr) {
        qDebug() << "Worksheets error";
        return;
    }
    int sheetCount = workSheets->property("Count").toInt();

    // 初始化资源
    InitResource();

    // 遍历表格数据
    for (int sheetIndex = 0; sheetIndex < sheetCount; ++sheetIndex) {

        // 获取工作表的行数和列数
        QAxObject *workSheet = workSheets->querySubObject("Item(int)", sheetIndex + 1); // Excel表格从1开始
        QAxObject *usedRange = workSheet->querySubObject("UsedRange");
        QAxObject *rows = usedRange->querySubObject("Rows");
        QAxObject *columns = usedRange->querySubObject("Columns");
        int rowCount = rows->property("Count").toInt();
        int columnCount = columns->property("Count").toInt();

        // 获取表名称
        QString sheetName = workSheet->property("Name").toString();

        // 将单元格数据显示在TableView上
        QStringList columnHeadstringLabelList; // 存储标题
        QTableView *tableView = new QTableView(); // 表格视图
        m_tableViewList.append(tableView);
        QStandardItemModel *tableModel = new QStandardItemModel(); // 表格数据模型
        tableModel->setRowCount(rowCount - 1); // 设置显示行数 标题显示在TableView，不占用表格行数
        tableModel->setColumnCount(columnCount); // 设置显示列数
        m_excelDataModelList.append(tableModel);
        for (int row = 0; row < rowCount; ++row) {
            for (int column = 0; column < columnCount; ++column) {
                QAxObject *cell = workSheet->querySubObject("Cells(inty, int)", row + 1, column + 1); // Excel文件行列数从1开始
                QVariant value = cell->property("Value"); //  获取单元格的值
                if (row == 0) {
                    columnHeadstringLabelList.append(value.toString()); // 存储标题
                } else {
                    QModelIndex index = tableModel->index(row - 1, column);
                    tableModel->setData(index, value);
                }
            }
        }

        // 设置TableView水平标题
        tableModel->setHorizontalHeaderLabels(columnHeadstringLabelList);
        m_sheetHeadStringTableList.append(columnHeadstringLabelList);

        tableView->setModel(tableModel);
        connect(tableModel, &QStandardItemModel::dataChanged, this, &ExcelDemo_M::tableCellDataChanged);

        m_middleTabWidgetInMainWidget->addTab(tableView, sheetName);
    }

    // 关闭Excel
    workBook->dynamicCall("Close()");
    excelApp.dynamicCall("Quit()");
}

void ExcelDemo_M::InitResource()
{
    if (!m_tableViewList.empty()) {
        for (auto itor = m_tableViewList.begin(); itor != m_tableViewList.end(); ++itor) {
            delete *itor;
        }
    }
    m_tableViewList.clear();

    if (!m_excelDataModelList.isEmpty()) {
        for (auto itor = m_excelDataModelList.begin(); itor != m_excelDataModelList.end(); ++itor) {
            delete *itor;
        }
    }
    m_excelDataModelList.clear();

    m_sheetHeadStringTableList.clear();

    m_changedValueCache.clear();
}

// 数据存进Excel文件
void ExcelDemo_M::on_saveAction_clicked()
{
    // 没有打开的Excel文件，直接返回
    if (m_excelFilePath.isEmpty()) {
        qDebug() << "No open Excel file";
        return;
    }

    // 没有修改内容，直接返回
    if (m_changedValueCache.isEmpty()) {
        qDebug() << "No modifications were made";
        return;
    }

    // 获取工作表集合
    QAxObject excelApp("Excel.Application");
    QAxObject *workBooks = excelApp.querySubObject("Workbooks");
    if (workBooks == nullptr) {
        qDebug() << "Workbooks error";
        return;
    }
    QAxObject *workBook = workBooks->querySubObject("Open(const QString)", m_excelFilePath);
    if (workBook == nullptr) {
        qDebug() << "Open(const QString) error";
        return;
    }
    QAxObject *workSheets = workBook->querySubObject("Worksheets");
    if (workSheets == nullptr) {
        qDebug() << "Worksheets error";
        return;
    }

    // 遍历缓存，依次打开工作表存入需要修改的单元格的数据
    for (auto sheetIndexItor = m_changedValueCache.begin(); sheetIndexItor != m_changedValueCache.end(); ++sheetIndexItor) {

        // 拿到工作表索引
        int sheetIndex = sheetIndexItor.key();
        QAxObject *workSheet = workSheets->querySubObject("Item(int)", sheetIndex + 1); // Excel表格从1开始

        auto &sheetChangeData = sheetIndexItor.value();
        for (auto sheetChangeDataItor = sheetChangeData.begin(); sheetChangeDataItor != sheetChangeData.end(); ++sheetChangeDataItor) {
            QPair<int, int> cellPosition = sheetChangeDataItor.key();
            int row = cellPosition.first;
            int column = cellPosition.second;
            QVariant value = sheetChangeDataItor.value();

            // 拿到需要更新数据的单元格
            QAxObject *cell = workSheet->querySubObject("Cells(inty, int)", row + 2, column + 1); // Excel文件行列数从1开始, 还需要越过标题
            cell->setProperty("Value", value);
        }
    }
    m_changedValueCache.clear();

    // 保存并关闭工作簿
    workBook->dynamicCall("Save()");
    workBook->dynamicCall("Close()");
    excelApp.dynamicCall("Quit()");
}

// 单元格修改缓存
void ExcelDemo_M::tableCellDataChanged(const QModelIndex& topLeft, const QModelIndex& bottomRight, const QVector<int>& roles)
{
    // 目前仅支持存储修改单个单元格的值
    if (topLeft != bottomRight) {
        qDebug() << "Modifying values of multiple cells at once is not supported";
        return;
    }

    // 获取值改变的工作表，单元格位置以及改变后的内容
    if (roles.isEmpty()) {
        qDebug() << "Nothing changed";
        return;
    }
    int row = topLeft.row();
    int column = topLeft.column();
    int sheetIndex = m_middleTabWidgetInMainWidget->currentIndex();
    QStandardItemModel *currentDataChangedModel = m_excelDataModelList.at(sheetIndex);
    QModelIndex modelIndex = currentDataChangedModel->index(row, column);
    QVariant changedValue = currentDataChangedModel->data(modelIndex);

    // 将修改的值存进缓存
    QPair<int, int> cellPosition(row, column);
    auto sheetIndexItor = m_changedValueCache.find(sheetIndex);
    if (sheetIndexItor != m_changedValueCache.end()) {
        auto &changedSheetInfo = sheetIndexItor.value();
        auto changedCellItor = changedSheetInfo.find(cellPosition);
        if (changedCellItor != changedSheetInfo.end()) {
            changedCellItor.value() = changedValue;
        } else {

            // 单元格第一次修改
            changedSheetInfo.insert(cellPosition, changedValue);
        }
    } else {

        // 表格第一次修改
        QMap<QPair<int, int>, QVariant> sheetChangedInfo;
        sheetChangedInfo.insert(cellPosition, changedValue);
        m_changedValueCache.insert(sheetIndex, sheetChangedInfo);
    }
}
