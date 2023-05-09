import sys
import pandas as pd
from PyQt5.QtWidgets import QApplication, QMainWindow, QFileDialog, QPushButton, QLabel, QLineEdit, QMessageBox
from PyQt5.QtGui import QIcon
# -*- coding: utf-8 -*-

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.float_number = None
        self.initUI()

    def initUI(self):
        super().__init__()
        self.setWindowTitle("志愿汇审核")
        self.setGeometry(200, 200, 650, 500)
        self.setWindowIcon(QIcon("./Image/icon.png"))

        self.label = QLabel(self)
        self.label.setText('输入志愿时长阈值')
        self.label.adjustSize()
        self.label.move(10, 20)

        self.textbox = QLineEdit(self)
        self.textbox.move(10, 50)
        self.textbox.resize(200, 25)

        self.button = QPushButton('保存', self)
        self.button.move(240, 46)
        self.button.clicked.connect(self.saveFloat)

        self.duration_label = QLabel(self)
        self.duration_label.move(10, 100)
        self.duration_label.setText("选择志愿时长名单")
        self.duration_label.adjustSize()

        self.duration_path_label = QLabel(self)
        self.duration_path_label.move(120, 120)
        self.duration_path_label.setFixedWidth(500)

        self.duration_button = QPushButton('选择文件', self)
        self.duration_button.move(10, 120)
        self.duration_button.clicked.connect(self.get_duration_file)

        self.participant_label = QLabel(self)
        self.participant_label.move(10, 165)
        self.participant_label.setText("选择参与者名单")
        self.participant_label.adjustSize()

        self.participant_path_label = QLabel(self)
        self.participant_path_label.move(120, 185)
        self.participant_path_label.setFixedWidth(500)

        self.participant_button = QPushButton('选择文件', self)
        self.participant_button.move(10, 185)
        self.participant_button.clicked.connect(self.get_participant_file)

        self.folder_label = QLabel(self)
        self.folder_label.move(10, 225)
        self.folder_label.setText("选择输出路径")
        self.folder_label.adjustSize()

        self.folder_path_label = QLabel(self)
        self.folder_path_label.move(120, 245)
        self.folder_path_label.setFixedWidth(500)

        self.folder_button = QPushButton('选择输出路径', self)
        self.folder_button.move(10, 245)
        self.folder_button.clicked.connect(self.get_folder)

        self.output_Normal_label = QLabel(self)
        self.output_Normal_label.move(10, 290)
        self.output_Normal_label.setText("选择正常名单的文件名")
        self.output_Normal_label.adjustSize()

        self.output_Normal_textbox = QLineEdit(self)
        self.output_Normal_textbox.move(10, 310)
        self.output_Normal_textbox.resize(250, 25)

        self.output_Normal_button = QPushButton('保存', self)
        self.output_Normal_button.move(300, 305)
        self.output_Normal_button.clicked.connect(self.saveOutputNormalName)

        self.output_Error_label = QLabel(self)
        self.output_Error_label.move(10, 350)
        self.output_Error_label.setText("选择异常名单的文件名")
        self.output_Error_label.adjustSize()

        self.output_Error_textbox = QLineEdit(self)
        self.output_Error_textbox.move(10, 370)
        self.output_Error_textbox.resize(250, 25)

        self.output_Error_button = QPushButton('保存', self)
        self.output_Error_button.move(300, 375)
        self.output_Error_button.clicked.connect(self.saveOutputErrorName)

        self.generate_button = QPushButton('Start！', self)
        self.generate_button.move(250, 430)
        self.generate_button.clicked.connect(self.generate)

    # 保存阈值
    def saveFloat(self):
        try:
            float_number = float(self.textbox.text())
            # 在这里可以将该数字存储到一个变量或者数据库中
        except ValueError:
            QMessageBox.warning(self, "警告", "输入数据异常")

    # 选择志愿时长名单
    def get_duration_file(self):
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        filename, _ = QFileDialog.getOpenFileName(self, "选择志愿时长名单", "", "Excel Files (*.xlsx *.xls)", options=options)
        if filename:
            self.duration_path_label.setText(filename)

    # 选择参与者名单
    def get_participant_file(self):
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        filename, _ = QFileDialog.getOpenFileName(self, "选择参与者名单", "", "Excel Files (*.xlsx);;Excel Files (*.xls)",
                                                  options=options)
        if filename:
            self.participant_path_label.setText(filename)

    # 选择正常名单输出的文件名
    def saveOutputNormalName(self):
        # 获取输出文件名
        output_name = self.output_Normal_textbox.text()
        # 检查输出文件名是否为空
        if output_name == "":
            QMessageBox.warning(self, "警告", "输出文件名不能为空")
            return
        # 检查输出文件名是否合法
        '''ext = (".xlsx", ".xls")
        if not output_name.endswith(ext):
            QMessageBox.warning(self, "警告", "输出文件名必须以.xlsx结尾")
            return'''
        # 保存输出文件名为类属性
        self.output_Normal_name = output_name
        # 显示提示信息
        QMessageBox.information(self, "提示", f"已保存输出文件名为{output_name}")

    # 选择异常名单输出的文件名
    def saveOutputErrorName(self):
        # 获取输出文件名
        output_name = self.output_Error_textbox.text()
        # 检查输出文件名是否为空
        if output_name == "":
            QMessageBox.warning(self, "警告", "输出文件名不能为空")
            return
        '''# 检查输出文件名是否合法
        ext = (".xlsx", ".xls")
        if not output_name.endswith(ext):
            QMessageBox.warning(self, "警告", "输出文件名必须以.xlsx结尾")
            return'''
        # 保存输出文件名为类属性
        self.output_Error_name = output_name
        # 显示提示信息
        QMessageBox.information(self, "提示", f"已保存输出文件名为{output_name}")

    # 选择输出路径
    def get_folder(self):
        selected_directory = QFileDialog.getExistingDirectory()
        if selected_directory:
            # do something with the selected folder
            self.folder_path_label.setText(selected_directory)

    # 生成
    def generate(self):
        participant_file_path = self.participant_path_label.text()
        duration_file_path = self.duration_path_label.text()
        if participant_file_path and duration_file_path:
            participants_df = pd.read_excel(participant_file_path)
            signers_df = pd.read_excel(duration_file_path)
        else:
            return 0

        # 筛选参与者中未签到的人
        not_signed_df = participants_df[~participants_df['姓名'].isin(signers_df['姓名'])]

        # 筛选签到者中未参与活动的人
        not_participated_df = signers_df[~signers_df['姓名'].isin(participants_df['姓名'])]

        # 筛选签到者中未签退的人
        not_checked_out_df = signers_df[(signers_df['信用时数'] == 0) & (~signers_df['姓名'].isin(not_participated_df['姓名']))]

        # 筛选签到者中信用时数超过要求的人
        credit_hour = self.float_number
        exceeded_df = signers_df[
            (signers_df['信用时数'] > credit_hour) & (~signers_df['姓名'].isin(not_participated_df['姓名']))]

        # 输出异常签到名单
        with pd.ExcelWriter(f'{self.selected_directory}+{self.output_Error_name}.xlsx') as writer:
            not_signed_df.to_excel(writer, sheet_name='未签到')
            not_participated_df.to_excel(writer, sheet_name='未参与')
            not_checked_out_df.to_excel(writer, sheet_name='未签退')
            exceeded_df.to_excel(writer, sheet_name='信用时数超过要求')

        # 筛选正常签到人员
        normal_signers_df = signers_df[
            signers_df['姓名'].isin(participants_df['姓名']) & (signers_df['信用时数'] > 0) & (
                    signers_df['信用时数'] <= self.saveFloat())]

        # 输出正常签到名单
        with pd.ExcelWriter(f'{self.selected_directory}+{self.output_Normal_name}.xlsx') as writer:
            # 根据姓名在参与者表单中查找班级
            normal_signers_df['班级'] = normal_signers_df['姓名'].apply(
                lambda x: participants_df[participants_df['姓名'] == x]['班级'].values[0] if len(
                    participants_df[participants_df['姓名'] == x]) > 0 else 'null')
            # 根据姓名在签到者表单中查找信用时数
            normal_signers_df['信用时数'] = normal_signers_df['姓名'].apply(
                lambda x: signers_df[signers_df['姓名'] == x]['信用时数'].values[0] if len(
                    signers_df[signers_df['姓名'] == x]) > 0 else 'null')
            # 输出到名为“正常签到名单”的xlsx文件
            normal_signers_df.to_excel(writer, index=False)


if __name__ == '__main__':
    app = QApplication(sys.argv)
    main = MainWindow()
    main.show()
    sys.exit(app.exec_())
