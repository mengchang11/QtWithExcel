#include "exceldemo_m.h"
#include "ui_exceldemo_m.h"

ExcelDemo_M::ExcelDemo_M(QWidget *parent)
    : QMainWindow(parent)
    , ui(new Ui::ExcelDemo_M)
{
    ui->setupUi(this);

    InitExcelWidget();
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
}

ExcelDemo_M::~ExcelDemo_M()
{
    delete ui;
}

