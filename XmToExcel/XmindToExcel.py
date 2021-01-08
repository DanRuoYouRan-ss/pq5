'''
@Project:testcase_xmind
@Time:2020/12/9  16:35
@Author:曾令涛
'''

import sys
import re

from PyQt5.QtGui import QIcon
from PyQt5.QtWidgets import QWidget, QApplication, QMainWindow, QMessageBox, QDesktopWidget, QPushButton, QVBoxLayout, \
    QTableWidget, QAbstractItemView, QFileDialog, QTableWidgetItem, QLabel, QComboBox, QHeaderView

from XmToExcel.ExcelData import ExcelData
from XmToExcel.XmindData import XMindData


class Example(QWidget):

    def __init__(self):
        super().__init__()
        self.initUI()
        self.row = -1

    def initUI(self):
        # self.setGeometry(300, 300, 300, 220)
        self.resize(1144,400)
        self.setMinimumSize(624,400)
        # self.setMinimumSize(1144, 400)
        self.center()
        self.setWindowTitle('Xmind转Excel')
        self.setWindowIcon(QIcon(':/contacts.png'))
        # self.setWindowIcon(QIcon("../Icon/window.ico"))
        layout = QVBoxLayout()

        self.column=6

        # 使用水平布局
        self.tablewidget = QTableWidget()

        self.tablewidget.setColumnCount(self.column)
        self.tablewidget.setSelectionBehavior(QAbstractItemView.SelectRows)  # 设置表格的选取方式是行选取
        self.tablewidget.setSelectionMode(QAbstractItemView.SingleSelection)  # 设置选取方式为单个选取
        self.tablewidget.setHorizontalHeaderLabels(['#ID', '用例名称', '预期结果', '用例等级','需求ID','用例类型'])
        self.tablewidget.setAlternatingRowColors(True)  # 行是否自动变色
        self.tablewidget.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)  # 设置列宽的适应方式
        # self.tablewidget.setColumnWidth(1, 350)

        self.qc = QComboBox(self)
        self.qc.addItems(['请选择用例类型', '功能测试', '冒烟测试', '回归测试'])

        self.qc.activated[str].connect(self.onActivated)

        # self.tablewidget.itemChanged.connect(QCoreApplication.instance().quit)
        self.tablewidget.itemChanged.connect(self.table_update)

        self.button1 = QPushButton("选择XMind文件",self)
        self.button1.clicked.connect(self.read_XMind)

        self.button2 = QPushButton("转成Excel",self)
        self.button2.clicked.connect(self.to_Excel)

        layout.addWidget(self.tablewidget)
        layout.addWidget(self.qc)
        layout.addWidget(self.button1)
        layout.addWidget(self.button2)
        self.setLayout(layout)

        self.show()

    def read_XMind(self):
        # self.tablewidget.clearContents()  # 清除所有数据--不包括标题头
        for i in range(1,self.row+1):
            self.tablewidget.removeRow(0)
        self.tablewidget.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.qc.setCurrentIndex(0)

        fileName, _ = QFileDialog.getOpenFileName(self, "打开文件", ".", "XMind(*.xmind)")

        if fileName == '':
            return
        ds = XMindData.read_XMind_to_list(fileName)[0]["topic"]
        # 获取头节点title，就是需求名
        self.title = ds["title"]
        # 获取头节点下的其余所有节点内容
        nodes_data = ds["topics"]
        # 创建XMData格式化xMind读取的数据
        md = XMindData()
        # 清空缓存数据
        md.clear_init_list_data()
        # 调用
        data = md.get_lists_data(data=nodes_data)
        # 动态设置行
        self.row = len(data)
        # 设置表格的行
        self.tablewidget.setRowCount(self.row)
        for case in enumerate(data):
            # id
            item1 = QTableWidgetItem(str(case[0] + 1))
            # 需求名
            # item2 = QTableWidgetItem("_".join(case[1][:3]))

            item_1="_".join(case[1][:-2])
            item_2=item_1.replace('_','',1)

            # aaa=re.findall(r"\d+\.?\d*", item_2)
            # print(aaa[0])
            # item2 = QTableWidgetItem(item_2)
            # item2 = str(items).replace('_','',1)
            # 模块名
            item3 = QTableWidgetItem(case[1][-2])
            # 用例
            item4 = QTableWidgetItem(case[1][-1])
            item2 = None
            item5 = None

            # 需求ID
            bool1=False
            try:
                # item2 = QTableWidgetItem(re.sub(re.findall(r"\d+\.?\d*", item_2)[0], '', item_2))
                item2 = QTableWidgetItem(re.sub(re.findall(r"\d+\.?\d*", item_2)[0], '', item_2))
                item5 = QTableWidgetItem(re.findall(r"\d+\.?\d*", case[1][0])[0])
                bool1 = True
            except:
                item2 = QTableWidgetItem(item_2)
                pass

            # print(re.sub(re.findall(r"\d+\.?\d*", item_2)[0],'',item_2))

            self.tablewidget.setItem(case[0], 0, item1)
            self.tablewidget.setItem(case[0], 1, item2)
            self.tablewidget.setItem(case[0], 2, item3)
            self.tablewidget.setItem(case[0], 3, item4)

            if bool1:
                self.tablewidget.setItem(case[0], 4, item5)

        for i in range(0,self.row+1):
            item6 = QTableWidgetItem(None)
            self.tablewidget.setItem(i, 5, item6)
        self.qc.setCurrentIndex(0)  # 设置下拉框默认值
        self.tablewidget.horizontalHeader().setSectionResizeMode(QHeaderView.Fixed)
        self.tablewidget.setColumnWidth(1, 400)
        self.tablewidget.setColumnWidth(2, 400)

        for i in [0,3,4,5]:
            self.tablewidget.setColumnWidth(i, 70)
        fileName =''

    def table_update(self):
        """
        如果表格被编辑，则会保存编辑后的数据
        """
        self.tablewidget.selectedItems()

    def to_Excel(self):

        """
        保存到Excel中
        """

        if self.row == -1:
            QMessageBox.about(self, "请先选择", '请先选择Xmind文件...')
            self.qc.setCurrentIndex(0)
            return
        # 返回值是一个元祖，需要两个参数接收，第一个是文件名，第二个是文件类型
        fileName, fileType = QFileDialog.getSaveFileName(self, "保存Excel", self.title, "xlsx(*.xlsx)")

        if fileName == "":
            return
        else:
            try:
                if not fileName.endswith('.xlsx'):
                    # 如果后缀不是.xlsx，那么就加上后缀
                    fileName = fileName + '.xlsx'
                QMessageBox.about(self, "转换中...", '转换中，可能会有延迟，请不要关闭应用...')
                we = ExcelData(file_name=fileName)
                # 创建Excel
                we.creat_excel_and_set_title(titles=['#ID', '用例名称', '预期结果', '用例等级','需求ID','用例类型','前置条件','操作步骤','实际结果'])

                # 保存并写入(表格中更新后的数据)写入从第二行开始写入(行需要+1，列需要+1)
                for r in range(self.row):
                    for c in range(self.column):
                        # print(self.tablewidget.item(r, c))
                        if self.tablewidget.item(r, c) is None:
                        # if self.tablewidget.item(r, c).text() is None:
                            we.write_excel_data(row=r + 2, column=c + 1, value='')
                            continue
                        we.write_excel_data(row=r + 2, column=c + 1, value=self.tablewidget.item(r, c).text())
                QMessageBox.about(self, "写入Excel", "写入Excel成功！\n路径：{}".format(fileName))
            except Exception as e:
                QMessageBox.critical(self, "写入Excel", "出错了，请重试！！！\n错误信息：{}".format(e), QMessageBox.Yes, QMessageBox.Yes)

    def onActivated(self, text):  # 把下拉列表框中选中的列表项的文本显示在标签组件上
        if self.row ==-1:
            QMessageBox.about(self, "请先选择", '请先选择Xmind文件...')
            self.qc.setCurrentIndex(0)
        for i in range(0,self.row+1):
            if self.qc.currentText()=='请选择用例类型':
                item6 = QTableWidgetItem(None)
            else:
                item6 = QTableWidgetItem(self.qc.currentText())
            self.tablewidget.setItem(i, 5, item6)

        # 退出提醒
    def closeEvent(self, event):
        reply = QMessageBox.question(self, '退出提醒',"确认退出吗?", QMessageBox.Yes | QMessageBox.No, QMessageBox.No)

        if reply == QMessageBox.Yes:
            event.accept()
        else:
            event.ignore()

    # 居中显示
    def center(self):
        qr = self.frameGeometry()
        cp = QDesktopWidget().availableGeometry().center()
        qr.moveCenter(cp)
        self.move(qr.topLeft())

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = Example()
    sys.exit(app.exec_())