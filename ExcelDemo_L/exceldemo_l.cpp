#include "exceldemo_l.h"
#include "ui_exceldemo_l.h"

ExcelDemo_L::ExcelDemo_L(QWidget *parent)
    : QMainWindow(parent)
    , ui(new Ui::ExcelDemo_L)
{
    ui->setupUi(this);
}

ExcelDemo_L::~ExcelDemo_L()
{
    delete ui;
}

/*
  querySubObject() 方法来查询工作表集合中的指定项
*/

void ExcelDemo_L::Init()
{
    // 打开文件夹
    QString strFilePathName = QFileDialog::getOpenFileName(this, QStringLiteral("选择Excel文件"),"", tr("Exel file(*.xls *.xlsx)"));
    if(strFilePathName.isNull()) {
        return ;
    }

    excel = new QAxObject(this);  // 初始化
    if (excel == nullptr) {
        qDebug("excel is null!");
        return;
    }

    excel->setControl("Excel.Application");  // 加载office excle控件
    if (excel->isNull()) {  // isNull():若对象没有加载COM对象则返回true，否则返回false
        excel->setControl("KET.Application");  // 如果office excle控件加载失败就加载wps控件
    }

    excel->setProperty("Visible", false);  // 不显示 Excel 窗体

    workBooks = excel->querySubObject("WorkSheets");  // 获取工作簿集合
    if (workBooks == nullptr) {
        excel->dynamicCall("Quit()");  // 如果获取工作表集合失败 退出 Excel 应用程序
        return;
    }

    workBooks->dynamicCall("Open(const QString&)", strFilePathName);  //打开打开已存在的工作簿

    //workbooks->dynamicCall("Add()");  // 新增一个工作簿

    workBook = workBooks->querySubObject("ActiveWorkBook");  // 获取当前活动工作簿
    workSheets = workBook->querySubObject("WorkSheets");  // 获取工作表集合
    sheet = workBook->querySubObject("WorkSheets(int)", 1);  //获取工作表集合的工作表1，即sheet1
}

void ExcelDemo_L::AppendSheet()
{
    int sheetCount = workSheets->property("Count").toInt();  // 获取工作表的数量
    QAxObject *lastSheet = workSheets->querySubObject("Item(int)", sheetCount);  // 获取最后一个工作表lastSheet
    workSheets->dynamicCall("Add(QVariant)", lastSheet->asVariant()); // 在lastSheet之前插入一个新工作表
    QAxObject* newSheet = workSheets->querySubObject("Item(int)", sheetCount);  // 获取新增的工作表newSheet
    lastSheet->dynamicCall("Move(QVariant)", newSheet->asVariant());  // 将lastSheet移动到newSheet之前
}
// 注意：
// Add(QVariant var)：新增一个对象插入到var对应的对象“之前”
// Move(QVariant var)：将调用对象移动到var对应对象“之前”
// 若worksheets直接调用Add()，入参为空，则是在活动工作表之前插入一个新的工作表

