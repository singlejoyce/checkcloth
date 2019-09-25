# -*- coding:utf-8 -*-
"""
Basic Layout
"""
import sys
import time
import os
from PyQt5 import QtGui
from PyQt5.QtWidgets import *
import pandas as pd
import numpy as np

__author__ = "joyce"


class MainUi(QMainWindow):
    def __init__(self, parent=None):
        super(MainUi, self).__init__(parent)
        self.cloth_data_list = np.array(self.read_cloth_data_excel()).tolist()
        self.fileName = ""
        self.resize(1066, 784)
        self.setWindowTitle('服饰数值检查工具')
        self.sheetname = QTextEdit()
        self.xls_dir = QTextEdit()
        self.xls_path = QLineEdit(os.getcwd())
        self.startButton = QPushButton("开始")
        self.selectButton = QPushButton("选择xls文件")
        self.selectButton.clicked.connect(self.btn_chooseFile)
        self.startButton.clicked.connect(self.startWork)

        self.attribute_dict = {
            1: "华丽",
            2: "简约",
            3: "优雅",
            4: "活泼",
            5: "成熟",
            6: "可爱",
            7: "性感",
            8: "清纯",
            9: "保暖",
            10: "清凉",
        }

        self.effect_level_dict = {
            1: "普通",
            2: "星光",
            3: "多彩",
            4: "流光",
            5: "绚丽",
            6: "梦幻",
        }

        self.cloth_type_dict = {
            0: "发型",
            1: "脸型",
            2: "上衣",
            3: "手部配饰",
            4: "下装",
            5: "头部配饰",
            6: "脸部配饰",
            7: "肩部配饰",
            8: "背部装扮",
            9: "左手手持",
            10: "右手手持",
            11: "手腕配饰",
            12: "臀部配饰",
            13: "袜子",
            14: "鞋子",
            15: "肤色",
            16: "变身",
            17: "套装",
            18: "双手纹身",
            19: "双腿纹身",
            20: "身体纹身",
            21: "腿部挂饰",
        }

        self.cloth_color_dict = {
            1: "红色",
            2: "橙色",
            3: "黄色",
            4: "绿色",
            5: "青色",
            6: "蓝色",
            7: "紫色",
        }

        font = QtGui.QFont()
        font.setFamily("Arial")
        font.setPointSize(11)
        self.setFont(font)

        self.retranslateUi()

    def retranslateUi(self):
        pathGroupBox = QGroupBox("xls文件起始路径：")
        dirGroupBox = QGroupBox("xls文件地址：")
        sheetGroupBox = QGroupBox("检查结果：")
        btnGroupBox = QGroupBox("")

        layout = QVBoxLayout()
        layout.addWidget(self.xls_path)
        pathGroupBox.setLayout(layout)

        layout1 = QVBoxLayout()
        layout1.addWidget(self.xls_dir)
        dirGroupBox.setLayout(layout1)

        layout2 = QVBoxLayout()
        layout2.addWidget(self.sheetname)
        sheetGroupBox.setLayout(layout2)

        layout3 = QHBoxLayout()
        layout3.addWidget(self.selectButton)
        layout3.addWidget(self.startButton)
        btnGroupBox.setLayout(layout3)

        mainLayout = QVBoxLayout()
        mainLayout.addWidget(pathGroupBox)
        mainLayout.addWidget(dirGroupBox)
        mainLayout.addWidget(sheetGroupBox)
        mainLayout.addWidget(btnGroupBox)

        main_frame = QWidget()
        main_frame.setLayout(mainLayout)
        self.setCentralWidget(main_frame)

    def closeEvent(self, event):
        reply = QMessageBox.question(self, '提示', '确认退出吗？',
                                     QMessageBox.Ok | QMessageBox.Cancel, QMessageBox.Cancel)
        if reply == QMessageBox.Ok:
            event.accept()
        elif reply == QMessageBox.Cancel:
            event.ignore()

    def btn_chooseFile(self):
        self.xls_dir.setPlainText("")
        self.sheetname.setPlainText("")

        if not os.path.exists(self.xls_path.text()):
            reply = QMessageBox.critical(self, "提示", self.tr("xls文件起始路径不存在，请检查!"), QMessageBox.Ok)
            if reply == QMessageBox.Ok:
                return

        self.fileName, filetype = QFileDialog.getOpenFileName(self,
                                                              "选取文件",
                                                              self.xls_path.text(),  # 起始路径
                                                              "Excel Files(*.xls *.xlsx)")
        if self.fileName == '':
            reply = QMessageBox.critical(self, "提示", self.tr("未选择xls文件，请检查!"), QMessageBox.Ok)
            if reply == QMessageBox.Ok:
                return

        self.xls_dir.setPlainText(self.fileName)

    def startWork(self):
        # 如果xls未选择进行警告提示
        if self.fileName == '':
            reply = QMessageBox.critical(self, "提示", self.tr("未选择需检查的xls文件，请检查!"), QMessageBox.Ok)
            if reply == QMessageBox.Ok:
                return
        self.check_cloth_excel()

    def read_cloth_data_excel(self):
        xls = os.getcwd() + '\\服饰数值.xlsx'
        if os.path.exists(xls):
            clothdata = pd.read_excel(xls, sheet_name='服饰数值')
            clothdata['type&effectlevel'] = clothdata['三级子类'].map(str) + ',' + clothdata['光效等级'].map(str)
            clothdata = clothdata[['type&effectlevel', '时尚值', '分数']]
            clothdata.reset_index(drop=True, inplace=True)
            return clothdata
        else:
            reply = QMessageBox.critical(self, "提示", self.tr("服饰数值.xlsx缺失，请检查!"), QMessageBox.Ok)
            if reply == QMessageBox.Ok:
                return

    def check_cloth_excel(self):
        sheet_names = ['服饰', '旧特效服饰']
        file_save_path = os.getcwd() + '\\reslut\\'
        if not os.path.isdir(file_save_path):
            os.makedirs(file_save_path)
        rq = time.strftime('%Y%m%d%H%M%S', time.localtime(time.time()))
        write = pd.ExcelWriter(file_save_path + rq + ".xlsx")
        for d in range(2):
            df = pd.read_excel(self.fileName, sheet_name=sheet_names[d])
            df['三级子类new'] = df['三级子类'].apply(lambda x: self.cloth_type_dict.get(x))
            df['服饰颜色new'] = df['服饰颜色'].apply(lambda x: self.cloth_color_dict.get(x))
            df['属性new'] = df['属性'].apply(lambda x: self.attribute_dict.get(x))
            df['光效等级new'] = df['光效等级'].apply(lambda x: self.effect_level_dict.get(x))
            df['type&effectlevel'] = df['三级子类'].map(str) + ',' + df['光效等级'].map(str)
            df = df[['物品ID', '物品名称', '三级子类', '三级子类new', '服饰颜色', '服饰颜色new',
                     '属性', '属性new', '光效等级', '光效等级new', '时尚值', '分数', 'TIPS',
                     'type&effectlevel']]
            df.reset_index(drop=True, inplace=True)

            # 根据数据表查找对应光效和部位的分数和时尚值
            cloth_list = np.array(df).tolist()
            score_new = []
            fashion_new = []

            for i in cloth_list:
                for j in self.cloth_data_list:
                    if i[13] in j:  # 第13列为type&effectlevel，光效等级与部位的合并值
                        score_new.append(j[2])  # j[1]为时尚值， j[2]为分数值
                        fashion_new.append(j[1])

            df['时尚值new'] = pd.DataFrame(fashion_new, columns=['时尚值new'])
            df['分数new'] = pd.DataFrame(score_new, columns=['分数new'])
            df = df[['物品ID', '物品名称', '三级子类', '三级子类new', '服饰颜色', '服饰颜色new',
                     '属性', '属性new', '光效等级', '光效等级new',
                     '分数', '分数new', '时尚值', '时尚值new', 'TIPS']]
            df.reset_index(drop=True, inplace=True)
            df['分数是否正确'] = df[['分数', '分数new']].apply(lambda r: 1 if r['分数'] == r['分数new'] else 0, axis=1)
            df['时尚值是否正确'] = df[['时尚值', '时尚值new']].apply(lambda r: 1 if r['时尚值'] == r['时尚值new'] else 0, axis=1)
            df.reset_index(drop=True, inplace=True)
            # 保存
            df.to_excel(excel_writer=write, sheet_name=sheet_names[d], index=False)
            fashion_failed = df[df['时尚值是否正确'].isin([0])]['物品ID'].tolist()
            score_failed = df[df['分数是否正确'].isin([0])]['物品ID'].tolist()
            self.sheetname.append('%s: \n分数异常的物品id：%s' % (sheet_names[d], str(score_failed)))
            self.sheetname.append('时尚值异常的物品id：%s' % str(fashion_failed))

        write.save()
        write.close()
        QMessageBox.information(self, "提示", self.tr("xls文件检查完成!"), QMessageBox.Ok)


if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = MainUi()
    ex.show()
    sys.exit(app.exec_())
