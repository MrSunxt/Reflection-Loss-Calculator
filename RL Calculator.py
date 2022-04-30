import sys
import os
import time
import cmath
import math
from openpyxl import load_workbook, Workbook
from PyQt5.QtCore import QThread, pyqtSignal, Qt
from PyQt5.QtGui import QIcon, QColor, QFont
from PyQt5.QtWidgets import *

path = ''
thick = 0


class RlCalculate(QThread):
    display_sig = pyqtSignal(str)
    path_sig = pyqtSignal(str)
    flag_sig = pyqtSignal(bool)

    def __init__(self):
        super(RlCalculate, self).__init__()

    @staticmethod
    def check_time():
        t = time.perf_counter()
        return t

    def sort_data(self):
        with open(path, 'r') as file:
            data = file.readlines()
        f = []
        r_e = []
        i_e = []
        r_u = []
        i_u = []
        e = []
        u = []
        if path.split('.')[-1] == 'dat':
            self.display_sig.emit('已识别 思仪-3672C-S 的数据')
            for i in data[27:]:
                f.append(float(i.split()[0]) * 1000000000)
                r_e.append(float(i.split()[1]))
                i_e.append(float(i.split()[2]))
                r_u.append(float(i.split()[3]))
                i_u.append(float(i.split()[4]))
            for i in range(0, len(f)):
                e.append(complex(r_e[i], -i_e[i]))
                u.append(complex(r_u[i], -i_u[i]))
            return f, e, u
        elif path.split('.')[-1] == 'prn':
            if len(data) > 3000:
                self.display_sig.emit('已识别 安捷伦-E5071C 的数据')
                data = [i for i in data if i != '\n']
                for i in data[1:]:
                    f.append(float(i.split()[0]))
                    r_e.append(float(i.split()[1]))
                    i_e.append(float(i.split()[2]))
                    r_u.append(float(i.split()[3]))
                    i_u.append(float(i.split()[4]))
                for i in range(0, len(f)):
                    e.append(complex(r_e[i], -i_e[i]))
                    u.append(complex(r_u[i], -i_u[i]))
                return f, e, u
            else:
                self.display_sig.emit('已识别 安捷伦-PNA-N5244A 的数据')
                for i in data[2:]:
                    f.append(float(i.split()[0]))
                    r_e.append(float(i.split()[1]))
                    i_e.append(float(i.split()[2]))
                    r_u.append(float(i.split()[3]))
                    i_u.append(float(i.split()[4]))
                for i in range(0, len(f)):
                    e.append(complex(r_e[i], -i_e[i]))
                    u.append(complex(r_u[i], -i_u[i]))
                return f, e, u
        elif path.split('.')[-1] == 'csv':
            self.display_sig.emit('已识别 安捷伦(Agilent N5245A) 的数据')
            data = [i for i in data[14:] if i != '\n']
            for i in data:
                f.append(float((i.split(',')[0]).strip()) * 1000000000)
                r_e.append(float((i.split(',')[1]).strip()))
                i_e.append(float((i.split(',')[2]).strip()))
                r_u.append(float((i.split(',')[3]).strip()))
                i_u.append(float((i.split(',')[4]).strip()))
            for i in range(0, len(f)):
                e.append(complex(r_e[i], -i_e[i]))
                u.append(complex(r_u[i], -i_u[i]))
            return f, e, u

    def run(self):
        self.flag_sig.emit(False)
        t1 = self.check_time()
        data = self.sort_data()
        f = data[0]
        e = data[1]
        u = data[2]
        self.display_sig.emit('加载文件......')
        wb = load_workbook('RL-IM.xlsx')
        rl_ws = wb['RL']
        im_ws = wb['IM']
        wb.remove(rl_ws)
        wb.remove(im_ws)
        ws_rl = wb.create_sheet('RL')
        ws_rl.append(
            [str(round(i * (thick / 100), 2)) + 'mm' for i in range(101)])
        ws_im = wb.create_sheet('IM')
        ws_im.append(
            [str(round(i * (thick / 100), 2)) + 'mm' for i in range(101)])
        self.display_sig.emit('已选厚度：{}mm\n计算中......'.format(thick))
        for i in range(0, len(f)):
            list_rl = []
            list_im = []
            for j in range(0, 101):
                d = round((thick/100000) * j, 5)
                zin = cmath.sqrt(u[i] / e[i]) * cmath.tanh(complex(0, 1) * ((2 * cmath.pi * f[i] * d) / 300000000) * cmath.sqrt(u[i] * e[i]))
                im = abs(zin)
                rl = 20 * cmath.log10(abs((zin - 1) / (zin + 1)))
                list_rl.append(rl.real)
                list_im.append(im)
            ws_rl.append(list_rl)
            ws_im.append(list_im)
        self.display_sig.emit('保存中......')
        wb.save('./RL-IM.xlsx')
        wb.close()
        t2 = self.check_time()
        self.display_sig.emit('运行结束！\n耗时：{:.4f}s'.format(t2 - t1))
        self.path_sig.emit('RL、IM数据已保存在： RL-IM.xlsx')
        self.flag_sig.emit(True)


class AlphaCalculate(QThread):
    display_sig = pyqtSignal(str)
    path_sig = pyqtSignal(str)
    flag_sig = pyqtSignal(bool)

    def __init__(self):
        super(AlphaCalculate, self).__init__()

    @staticmethod
    def check_time():
        t = time.perf_counter()
        return t

    def sort_data(self):
        with open(path, 'r') as file:
            data = file.readlines()
        frequency = []
        real_e = []
        imag_e = []
        real_u = []
        imag_u = []
        if path.split('.')[-1] == 'dat':
            self.display_sig.emit('已识别 思仪-3672C-S 的数据')
            for i in data[27:]:
                frequency.append(float(i.split()[0]) * 1000000000)
                real_e.append(float(i.split()[1]))
                imag_e.append(float(i.split()[2]))
                real_u.append(float(i.split()[3]))
                imag_u.append(float(i.split()[4]))
            return frequency, real_e, imag_e, real_u, imag_u
        elif path.split('.')[-1] == 'prn':
            if len(data) > 3000:
                self.display_sig.emit('已识别 安捷伦-E5071C 的数据')
                data = [i for i in data if i != '\n']
                for i in data[1:]:
                    frequency.append(float(i.split()[0]))
                    real_e.append(float(i.split()[1]))
                    imag_e.append(float(i.split()[2]))
                    real_u.append(float(i.split()[3]))
                    imag_u.append(float(i.split()[4]))
                return frequency, real_e, imag_e, real_u, imag_u
            else:
                self.display_sig.emit('已识别 安捷伦-PNA-N5244A 的数据')
                for i in data[2:]:
                    frequency.append(float(i.split()[0]))
                    real_e.append(float(i.split()[1]))
                    imag_e.append(float(i.split()[2]))
                    real_u.append(float(i.split()[3]))
                    imag_u.append(float(i.split()[4]))
                return frequency, real_e, imag_e, real_u, imag_u
        elif path.split('.')[-1] == 'csv':
            self.display_sig.emit('已识别 安捷伦(Agilet N5245A) 的数据')
            data = [i for i in data[14:] if i != '\n']
            for i in data:
                frequency.append(float((i.split(',')[0]).strip()) * 1000000000)
                real_e.append(float((i.split(',')[1]).strip()))
                imag_e.append(float((i.split(',')[2]).strip()))
                real_u.append(float((i.split(',')[3]).strip()))
                imag_u.append(float((i.split(',')[4]).strip()))
            return frequency, real_e, imag_e, real_u, imag_u

    def run(self):
        self.flag_sig.emit(False)
        t3 = self.check_time()
        data = self.sort_data()
        f = data[0]
        r_e = data[1]
        i_e = data[2]
        r_u = data[3]
        i_u = data[4]
        self.display_sig.emit('加载文件......')
        wb = load_workbook('alpha.xlsx')
        ws = wb['alpha']
        wb.remove(ws)
        new_ws = wb.create_sheet('alpha')
        self.display_sig.emit('计算α中......')
        for i in range(0, len(f)):
            list_alpha = []
            part1 = (math.sqrt(2) * math.pi * f[i]) / 300000000
            part2 = i_u[i] * i_e[i] - r_u[i] * r_e[i]
            part3 = r_u[i] * i_e[i] + i_u[i] * r_e[i]
            alpha = part1 * math.sqrt(part2 + math.sqrt((part2 ** 2) + (part3 ** 2)))
            list_alpha.append(alpha)
            new_ws.append(list_alpha)
        self.display_sig.emit('保存中......')
        wb.save('./alpha.xlsx')
        wb.close()
        t4 = self.check_time()
        self.display_sig.emit('运行结束！\n耗时：{:.4f}s'.format(t4 - t3))
        self.path_sig.emit('α数据已保存在： alpha.xlsx')
        self.flag_sig.emit(True)


class MyWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.chose_button = QPushButton('选择文件')
        self.c_button1 = QPushButton('计算RL-IM')
        self.c_button2 = QPushButton('计算α')
        self.check_button = QPushButton('查看')
        self.shut_button = QPushButton('关闭')
        self.about_button = QPushButton('关于')
        self.text = QTextBrowser()
        self.display_text = QTextBrowser()
        self.path_text = QTextBrowser()
        self.a = None
        self.rl_im = None
        self.init_ui()

    def init_ui(self):
        # 主窗口设置
        self.setFont(QFont('Microsoft YaHei', 10))
        self.setFixedSize(500, 500)
        self.setWindowTitle('RL Calculator')
        self.setWindowIcon(QIcon('D:/Pycharm/Project/PyQtProject/pic.ico'))
        self.setStyleSheet('background-color:rgb(240, 248, 255)')
        # 主布局
        layout = QVBoxLayout()
        # 选择文件布局
        layout.addStretch()
        h_layout1 = QHBoxLayout()
        h_layout1.addWidget(self.text)
        h_layout1.addWidget(self.chose_button)
        layout.addLayout(h_layout1)
        # 计算按钮布局
        layout.addStretch()
        h_layout2 = QHBoxLayout()
        h_layout5 = QVBoxLayout()
        h_layout2.addWidget(self.display_text)
        h_layout5.addWidget(self.c_button1)
        h_layout5.addWidget(self.c_button2)
        h_layout2.addLayout(h_layout5)
        layout.addLayout(h_layout2)
        # 查看文件
        layout.addStretch()
        h_layout3 = QHBoxLayout()
        h_layout3.addWidget(self.path_text)
        h_layout3.addWidget(self.check_button)
        layout.addLayout(h_layout3)
        # 关闭
        layout.addStretch()
        h_layout4 = QHBoxLayout()
        h_layout4.addStretch()
        h_layout4.addWidget(self.about_button)
        h_layout4.addWidget(self.shut_button)
        h_layout4.addStretch()
        layout.addLayout(h_layout4)
        self.setLayout(layout)
        # 选择类样式
        self.chose_button.setFixedSize(100, 80)
        self.chose_button.setStyleSheet('background-color:rgb(230, 230, 250)')
        self.text.setFixedSize(350, 80)
        self.text.setStyleSheet('background-color:rgb(253, 245, 230)')
        self.text.setPlaceholderText('源文件路径\n'
                                     '(*.dat, *.prn, *.csv)')
        # 计算类样式
        self.c_button1.setFixedSize(100, 80)
        self.c_button2.setFixedSize(100, 80)
        self.c_button1.setStyleSheet('background-color:rgb(151, 255, 255)')
        self.c_button2.setStyleSheet('background-color:rgb(151, 255, 255)')
        # 展示类样式
        self.display_text.setFixedSize(350, 220)
        self.display_text.setTextColor(QColor(255, 0, 0))
        self.display_text.setStyleSheet('background-color:rgb(253, 245, 230)')
        self.display_text.setPlaceholderText('注意事项：\n'
                                             '# 先确定仪器型号是否支持\n'
                                             '# RL-IM默认计算101个数据点\n'
                                             '# 计算结束请及时另存数据\n'
                                             '仪器型号：\n'
                                             '# 思仪 3672C-S (*.dat)\n'
                                             '# 安捷伦 E5071C (*.prn)\n'
                                             '# 安捷伦 PNA-N5244A (*.prn)\n'
                                             '# 安捷伦(Agilent N5245A) (*.csv)')
        # 查看类样式
        self.check_button.setFixedSize(100, 35)
        self.check_button.setStyleSheet('background-color:rgb(230, 230, 250)')
        self.path_text.setFixedSize(350, 35)
        self.path_text.setStyleSheet('background-color:rgb(253, 245, 230)')
        self.path_text.setPlaceholderText('数据文件')
        # 按钮点击事件
        self.chose_button.clicked.connect(self.findfile)

        self.c_button1.clicked.connect(self.click_rl)
        self.c_button2.clicked.connect(self.click_alpha)

        self.check_button.clicked.connect(self.openfile)

        self.about_button.clicked.connect(self.click_about)
        self.shut_button.clicked.connect(QApplication.quit)

    def findfile(self):
        findfile_name = QFileDialog.getOpenFileName(self, '选择文件',
                                                    '',
                                                    'Data files(*.dat *.prn *.csv)')
        tex = str(findfile_name[0])
        self.text.setText(tex)
        self.path_text.clear()
        self.display_text.clear()
        global path
        path = tex

    def check_file1(self, filename):
        self.display_text.clear()
        self.path_text.clear()
        self.display_text.setText('检查文件......')
        if not os.path.exists(filename):
            wb = Workbook()
            wb.remove(wb['Sheet'])
            wb.create_sheet('RL')
            wb.create_sheet('IM')
            wb.save(filename)
            wb.close()
        else:
            pass

    def check_file2(self, filename):
        self.display_text.clear()
        self.path_text.clear()
        self.display_text.setText('检查文件......')
        if not os.path.exists(filename):
            wb = Workbook()
            wb.remove(wb['Sheet'])
            wb.create_sheet('alpha')
            wb.save(filename)
            wb.close()
        else:
            pass

    def openfile(self):
        str_path = self.path_text.toPlainText()
        os.startfile(str_path.split(' ')[-1])

    def click_about(self):
        msg = 'Version：1.2.5<br>' \
              'Author：Xuetao Sun<br>Download：' \
              '<a href="https://github.com/MrSunxt/Reflection-Loss-Calculator/releases">GitHub</a><br>' \
              'Feedback：18268219283@163.com'
        QMessageBox.about(self, 'About', msg)

    def click_rl(self):
        if path == '':
            self.display_text.setText('\n\n\n{}\n{}\n{}'
                                      .format('！！！ 请先选择文件 ！！！',
                                              '！！！ 请先选择文件 ！！！',
                                              '！！！ 请先选择文件 ！！！'))
        else:
            num, ok = QInputDialog.getInt(self, '请输入', '最大厚度(mm)：',
                                          value=10, min=1, max=10, step=5,
                                          flags=Qt.WindowCloseButtonHint)
            if ok and num:
                global thick
                thick = num
                self.check_file1('RL-IM.xlsx')
                self.rl_im = RlCalculate()
                self.rl_im.flag_sig.connect(self.enable_button)
                self.rl_im.flag_sig.connect(self.button_type1)
                self.rl_im.display_sig.connect(self.displaytext)
                self.rl_im.path_sig.connect(self.p_text)
                self.rl_im.start()

    def click_alpha(self):
        if path == '':
            self.display_text.setText('\n\n\n{}\n{}\n{}'
                                      .format('！！！ 请先选择文件 ！！！',
                                              '！！！ 请先选择文件 ！！！',
                                              '！！！ 请先选择文件 ！！！'))
        else:
            self.check_file2('alpha.xlsx')
            self.a = AlphaCalculate()
            self.a.flag_sig.connect(self.enable_button)
            self.a.flag_sig.connect(self.button_type2)
            self.a.display_sig.connect(self.displaytext)
            self.a.path_sig.connect(self.p_text)
            self.a.start()

    def displaytext(self, t):
        self.display_text.append(t)

    def p_text(self, p):
        self.path_text.setText(p)

    def enable_button(self, flag):
        self.chose_button.setEnabled(flag)
        self.c_button1.setEnabled(flag)
        self.c_button2.setEnabled(flag)
        self.shut_button.setEnabled(flag)

    def button_type1(self, flag):
        if not flag:
            self.c_button1.setStyleSheet('background-color:rgb(253, 245, 230)')
            self.c_button1.setFont(QFont('Microsoft YaHei', 10))
        else:
            self.c_button1.setStyleSheet('background-color:rgb(151, 255, 255)')

    def button_type2(self, flag):
        if not flag:
            self.c_button2.setStyleSheet('background-color:rgb(253, 245, 230)')
            self.c_button2.setFont(QFont('Microsoft YaHei', 10))
        else:
            self.c_button2.setStyleSheet('background-color:rgb(151, 255, 255)')


if __name__ == '__main__':
    app = QApplication(sys.argv)

    window = MyWindow()
    window.show()

    app.exec()
