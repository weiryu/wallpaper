# -*- coding: utf-8 -*-
"""
说明：
需要建立一个1.qrc文件，将所有的资源文件（图标）写入，在文件目录下执行pyrcc5 -o images_qr.py 1.qrc
最后在代码中import images_qr，并修改资源文件路径名前加：
打包程序 pyinstaller -F -w -i title.ico 20170330.py
"""

import sys
import os
import shutil
import urllib.request
import re
import time
import win32gui
import win32con
import win32api
import images_qr

from PyQt5 import QtCore, QtGui
from PyQt5.QtCore import QRect, QMetaObject, QCoreApplication
from PyQt5.QtGui import QIcon, QPalette, QBrush, QPixmap, QFont
from PyQt5.QtWidgets import QFileDialog, QMenuBar, qApp, QDesktopWidget, QPushButton, QMenu, QAction
from PyQt5.QtWidgets import QWidget, QApplication, QMainWindow, QMessageBox, QStatusBar, QToolTip

# 界面函数
class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        # 设置窗口大小
        MainWindow.resize(600, 500)

        # 设置主窗口背景
        palette = QPalette()
        # 背景图片设置
        palette.setBrush(QPalette.Background, QBrush(QPixmap(":/background.jpg")))
        self.setPalette(palette)
        self.setAutoFillBackground(True)

        # 窗口置中
        self.center()
        self.centralwidget = QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        MainWindow.setCentralWidget(self.centralwidget)

        # 窗口无边框，效果不大合适
        # self.setWindowFlags(QtCore.Qt.FramelessWindowHint|QtCore.Qt.SubWindow|QtCore.Qt.WindowStaysOnTopHint)

        # 菜单栏部分
        self.menubar = QMenuBar(MainWindow)
        # 第一个，第二个参数为左上角坐标，后两个参数为宽度和高度
        self.menubar.setGeometry(QRect(0, 0, 600, 30))
        self.menubar.setObjectName("menubar")
        self.menu = QMenu(self.menubar)
        self.menu.setObjectName("menu")
        self.menu_2 = QMenu(self.menubar)
        self.menu_2.setObjectName("menu_2")
        MainWindow.setMenuBar(self.menubar)

        # 状态栏部分
        self.statusbar = QStatusBar(MainWindow)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)

        # 往菜单栏添加动作
        self.action = QAction(MainWindow)
        self.action.setObjectName("action")
        self.action_2 = QAction(MainWindow)
        self.action_2.setObjectName("action_2")
        self.action_3 = QAction(MainWindow)
        self.action_3.setObjectName("action_3")
        self.action_4 = QAction(MainWindow)
        self.action_4.setObjectName("action_4")
        self.action_5 = QAction(MainWindow)
        self.action_5.setObjectName("action_5")
        self.action_6 = QAction(MainWindow)
        self.action_6.setObjectName("action_6")
        self.menu.addAction(self.action)
        self.menu.addAction(self.action_2)
        self.menu.addAction(self.action_3)
        self.menu.addAction(self.action_4)
        self.menu_2.addAction(self.action_5)
        self.menu_2.addAction(self.action_6)
        self.menubar.addAction(self.menu.menuAction())
        self.menubar.addAction(self.menu_2.menuAction())


        # 中心区主要功能部件
        # 可选带图标的pushbutton
        # self.NI_mainwindow = QPushButton(QIcon(":/NI.jpg"), u'是', self )
        self.NI_mainwindow = QPushButton(self)
        self.NI_mainwindow.setGeometry(QRect(100, 100, 200, 200))
        # 透明效果
        self.NI_mainwindow.setFlat(True)
        self.NI_mainwindow.setStyleSheet('''color: yellow;
                                ''')
        self.NI_mainwindow.setObjectName("NI_mainwindow")
        #self.NI_mainwindow.setToolTip('设置国家地理每日精选为桌面壁纸')  # 悬浮提示框,在当前透明效果下失效
        self.NI_mainwindow.setStatusTip('设置国家地理每日精选为桌面壁纸')
        self.Bing_mainwindow = QPushButton(self)
        self.Bing_mainwindow.setGeometry(QRect(225, 200, 200, 200))
        self.Bing_mainwindow.setObjectName("Bing_mainwindow")
        self.Bing_mainwindow.setFlat(True)
        self.Bing_mainwindow.setStyleSheet('''color: yellow;
                                ''')
        self.Bing_mainwindow.setStatusTip("设置bing每日精选为桌面壁纸")
        self.WallHaven_mainwindow = QPushButton(self)
        self.WallHaven_mainwindow.setGeometry(QRect(350, 100, 200, 200))
        self.WallHaven_mainwindow.setObjectName("WallHaven_mainwindow")
        self.WallHaven_mainwindow.setFlat(True)
        self.WallHaven_mainwindow.setStyleSheet('''color: yellow;
                                ''')
        self.WallHaven_mainwindow.setStatusTip("设置壁纸天堂小时精选为桌面壁纸")
        self.retranslateUi(MainWindow)
        QMetaObject.connectSlotsByName(MainWindow)

        # 工具栏
        toolbar = self.addToolBar('工具栏')
        toolbar.addAction(QAction(QIcon(":/file.ico"), "设置本地图片为壁纸", self, triggered= self.user_set_wallpaper))
        toolbar.addAction(QAction(QIcon(":/folder.ico"), "设置壁纸保存文件夹", self, triggered= self.set_save_folder))
        toolbar.addAction(QAction(QIcon(":/clean.ico"), "清空本地壁纸文件夹", self, triggered = self.delete_wallpaper))
        # 间隔符
        # toolbar.addSeparator()
        toolbar.addAction(QAction(QIcon(":/close.ico"), "关闭软件", self, triggered=qApp.quit))
        
        palette = QtGui.QPalette()
        brush = QtGui.QBrush(QtGui.QColor(255, 255, 255))
        brush.setStyle(QtCore.Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Active, QtGui.QPalette.Base, brush)

    # 窗口居中函数1
    """
    def center(self):
        # 计算显示器分辨率
        screen = QDesktopWidget().screenGeometry()
        size = self.geometry()
        # 将窗口左上角移动到该位置，相当于将窗口移动到中间的位置
        self.move((screen.width() - size.width())/2, (screen.height() - size.height())/2)
     """

    # 窗口居中函数2，效果较好
    def center(self):
        # 得到主窗体的矩形说明
        qr = self.frameGeometry()
        # 算出显示器分辨率，得到中心点
        cp = QDesktopWidget().availableGeometry().center()
        # 设置矩形的中心点到屏幕的中心
        qr.moveCenter(cp)
        # 移动程序窗口的左上角到qr矩形的左上角，将窗口放到屏幕中心
        self.move(qr.topLeft())

    # 菜单名
    def retranslateUi(self, MainWindow):
        _translate = QCoreApplication.translate
        # 界面右上角的程序名
        MainWindow.setWindowTitle(_translate("MainWindow", "壁纸大师"))
        # 程序右上角的图标
        MainWindow.setWindowIcon(QIcon(r":/title.ico"))
        # 菜单名
        self.menu.setTitle(_translate("MainWindow", "设置"))
        self.menu_2.setTitle(_translate("MainWindow", "帮助"))
        # 二级菜单名
        self.action.setText(_translate("MainWindow", "设置本地图片壁纸"))
        self.action_2.setText(_translate("MainWindow", "设置壁纸保存文件夹"))
        self.action_3.setText(_translate("MainWindow", "清空壁纸保存文件夹"))
        self.action_4.setText(_translate("MainWindow", "退出"))
        self.action_5.setText(_translate("MainWindow", "关于软件"))
        self.action_6.setText(_translate("MainWindow", "版权说明"))
        # 主函数部分
        self.NI_mainwindow.setText("国家地理每日精选")
        self.Bing_mainwindow.setText("bing每日精选")
        self.WallHaven_mainwindow.setText("壁纸天堂小时精选")


# 功能函数
class Window(QMainWindow, Ui_MainWindow):
    def __init__(self, *args, **kwargs):
        # 初始化
        super(Window, self).__init__(*args, **kwargs)
        self.setupUi(self)

        self.action.setIcon(QIcon(":/file.ico"))
        self.action.setShortcut("Ctrl+O")
        self.action.setStatusTip("设置本地图片为壁纸")
        self.action.triggered.connect(self.user_set_wallpaper)

        self.action_2.setIcon(QIcon(":/folder.ico"))
        self.action_2.setShortcut("Ctrl+F")
        self.action_2.setStatusTip("设置本地壁纸保存文件夹")
        self.action_2.triggered.connect(self.set_save_folder)

        self.action_3.setIcon(QIcon(":/clean.ico"))
        self.action_3.setShortcut("Ctrl+D")
        self.action_3.setStatusTip("清空本地壁纸保存文件夹")
        self.action_3.triggered.connect(self.delete_wallpaper)

        self.action_4.setIcon(QIcon(":/close.ico"))
        self.action_4.setShortcut("Ctrl+Q")
        self.action_4.setStatusTip("退出程序")
        self.action_4.triggered.connect(qApp.quit)
        # 可选确认后退出函数
        # self.action_3.triggered.connect(self.closeEvent)

        self.action_5.setShortcut("Ctrl+A")
        self.action_5.setStatusTip("关于软件")
        self.action_5.triggered.connect(self.about_software_information)

        self.action_6.setShortcut("Ctrl+U")
        self.action_6.setStatusTip("版权说明")
        self.action_6.triggered.connect(self.copyright_information)

        self.NI_mainwindow.clicked.connect(self.nationalgeographic_wallpaper)
        self.Bing_mainwindow.clicked.connect(self.bing_wallpaper)
        self.WallHaven_mainwindow.clicked.connect(self.wallheaven_wallpaper)


    # 通用网址解码函数
    def open_url(self, url):
        # 增加浏览器头，防止反爬虫屏蔽
        headers = {'User-Agent': 'Mozilla/5.0 (Windows; U; Windows NT 6.1; en-US; rv:1.9.1.6) Gecko/20091201 Firefox/3.5.6',
                   'Referer': 'https://alpha.wallhaven.cc/latest',
                   'Connection': 'keep-alive'}
        # 异常处理
        try:
            req = urllib.request.Request(url, headers=headers)
            page = urllib.request.urlopen(req)
            html = page.read().decode('UTF-8')
            return html
        except urllib.request.URLError as e:
            if hasattr(e, "reason"):
                print("Failed to reach the server")
                print("The reason:", e.reason)
                return None
            elif hasattr(e, "code"):
                print("The server couldn't fulfill the request")
                print("Error code:", e.code)
                print("Return content:", e.read())
                return None
        except:
            return None
        # 无异常执行
        else:
            print("OK")
        # 无论是否异常均执行
        finally:
            print("OK")

    def set_save_folder(self):
        local = self.dictionary_get()
        print(local)
        folder = QFileDialog.getExistingDirectory(self, "选择壁纸保存文件夹", local)
        # 没有得到返回的文件夹名就不进行操作
        if folder == '':
            pass
        else:
            # 针对目录选择为根文件夹的情况修复bug
            if str(folder)[-1] == "/":
                folder_new = str(folder).replace("/","")
            else:
                folder_new = folder
            with open("壁纸保存目录.txt", "w") as fp:
                fp.write(folder_new)
                fp.close()


     # 用户自定义设置壁纸

    def user_set_wallpaper(self):
        local = self.dictionary_get()
        filename, _ = QFileDialog.getOpenFileName(self, "选取壁纸文件", local, "All Files (*);;Images (*.jpg *.png *.jpeg)")
        print(filename)
        if filename == '':
            pass
        else:
            setWallpaper(filename)

    # 清空本地壁纸函数
    def delete_wallpaper(self):
        local = self.dictionary_get()
        if local == '':
            QMessageBox.about(self,"壁纸大师","本地壁纸文件夹为空")
        else:
            reply = QMessageBox.information(self, '壁纸大师 ', '您确定要清空本地壁纸文件夹吗?', QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
            if reply == QMessageBox.Yes:
                # 判断文件夹是否存在
                if os.path.exists(local):
                    # 清空文件夹
                    local_nationalgeographic = local + "\\" + "nationalgeographic"
                    local_bing = local + "\\" + "bing_wallpaper"
                    local_wallheaven = local + "\\" + "wallheaven"
                    if os.path.exists(local_nationalgeographic):
                        shutil.rmtree(local_nationalgeographic)
                    if os.path.exists(local_bing):
                        shutil.rmtree(local_bing)
                    if os.path.exists(local_wallheaven):
                        shutil.rmtree(local_wallheaven)
                    QMessageBox.about(self, "壁纸大师", "本地壁纸文件夹清空完成")
                else:
                    QMessageBox.about(self, "壁纸大师", "本地壁纸文件夹不存在")
            elif reply == QMessageBox.No:
                pass
            else:
                pass

    # 读取壁纸保存目录函数
    def dictionary_get(self):
        # 设定第一次使用时默认壁纸保存目录
        if not os.path.isfile("壁纸保存目录.txt"):
            with open("壁纸保存目录.txt", "w") as fp:
                fp.write("E:\壁纸")
                fp.close()
        # 读取壁纸保存目录
        with open("壁纸保存目录.txt", "r") as fp:
            location = fp.read()
            fp.close()
        # 判断文件夹是否存在
        if not os.path.exists(location):
            os.mkdir(location)
        return location

    # 退出时进行确认的函数
    def closeEvent(self, event):
        # 默认退出选项是YES
        reply = QMessageBox.question(self, '壁纸大师', '您确定要退出吗?',
                                     QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes)

        if reply == QMessageBox.Yes:
            event.accept()
        else:
            event.ignore()

    def about_software_information(self):
        QMessageBox.about(self, "关于软件", "Version 1.3.2\nCopyright@weir\nAll rights reserved")

    def copyright_information(self):
        QMessageBox.about(self, "版权说明", "本软件仅供内部技术交流使用\n相关图片版权归著作权人所有\n请勿用于商业用途")
        # QMessageBox.information(self, "使用说明", " 本软件仅供内部技术交流使用，请勿用于商业用途\n 相关图片版权归著作权人所有，请勿用于商业用途", QMessageBox.Yes)

    # bing壁纸主函数
    def bing_wallpaper(self):
        # 正则匹配获得图片,可以匹配jpg和jpeg格式
        p = '/az/.*?\.jpe*g'
        # 获得本地文件夹地址
        local = self.dictionary_get()
        # 在这里local指定文件保存的文件夹位置！！！！！
        local_bing = local + "\\" + "bing_wallpaper"
        print(local_bing)
        # bing中国区主页网址
        url_china = 'http://global.bing.com/?FORM=HPCNEN&setmkt=zh-cn&setlang=zh-cn'
        # bing美国区主页地址
        url1_america = 'http://global.bing.com/?FORM=HPCNEN&setmkt=en-us&setlang=en-us'
        html = self.open_url(url_china)
        # 网络连接不通异常处理
        # 双网址备份，优先选择中文网址，其次选择英文网址
        if html == None:
            html = self.open_url(url1_america)
            if html == None:
                QMessageBox.about(self,"壁纸大师",'网络错误，请检查网络连接后重试')
                imglist = []
            else:
                imglist = re.findall(p, html)
        else:
            imglist = re.findall(p, html)
        # 异常处理
        if imglist != []:
            print(imglist)
            # 获得系统时间作为文件名
            time.localtime(time.time())
            n = time.strftime('%Y-%m-%d', time.localtime(time.time()))
            # 设定文件名，获得文件名后缀图片类型
            filetype = imglist[0].split('.')[-1]
            filename = 'bing-' + str(n) + "." + str(filetype)
            # 修改当天更换过壁纸后无法再次更换的bug
            imagepath = local_bing + "\\" + filename
            # 判断当前文件夹是否存在
            if not os.path.exists(local_bing):
                os.mkdir(local_bing)
            # 判断当前壁纸是否存在
            if os.path.isfile(local_bing + "\\" + filename):
                print("您今天已经更换过壁纸了")
            else:
                # 保存壁纸文件，返回文件位置
                bing_final_url = r'http://cn.bing.com' + imglist[0]
                urllib.request.urlretrieve(bing_final_url, imagepath)
            setWallpaper(imagepath)
        # 错误处理
        else:
            pass

    # 国家地理壁纸主函数
    def nationalgeographic_wallpaper(self):
        new_urllist = []
        local = self.dictionary_get()
        local_nationalgeographic = local + "\\" + "nationalgeographic"
        p = r'/photography/.*?\.html'
        url = 'http://www.nationalgeographic.com.cn/photography/photo_of_the_day/'
        html = self.open_url(url)
        # 网络连接异常处理
        if html == None:
            QMessageBox.about(self, "壁纸大师", '网络错误，请检查网络连接后重试')
            urllist = []
        else:
            urllist = re.findall(p, html)
        if urllist!= []:
            for url in urllist:
                if urllist.count(url) > 3:
                    new_urllist.append(url)
            new_url = r"http://www.nationalgeographic.com.cn" + new_urllist[0]
            new_html = self.open_url(new_url)
            q = r"http://image.nationalgeographic.com.cn/.\d+/.\d+/.\d+.jpe*g"
            if new_html == None:
                QMessageBox.about(self, "壁纸大师", '网络错误，请检查网络连接后重试')
                imglist = []
            else:
                imglist = re.findall(q, new_html)
            print(imglist)
            if imglist != []:
                n = time.strftime('%Y-%m-%d-%H', time.localtime(time.time()))
                filetype = imglist[len(imglist) - 1].split(".")[-1]
                print(filetype)
                filename = 'national-geographic' + str(n) + "." + str(filetype)
                if not os.path.exists(local_nationalgeographic):
                    os.mkdir(local_nationalgeographic)
                if os.path.isfile(local_nationalgeographic + "\\" + filename):
                    print("您今天已经更换过壁纸了")
                else:
                    print(imglist[len(imglist)-1])
                    urllib.request.urlretrieve(imglist[len(imglist) - 1], local_nationalgeographic + "\\" + filename)
                # 修改当天更换过壁纸后无法再次更换的bug
                imagepath = local_nationalgeographic + "\\" + filename
                setWallpaper(imagepath)
            else:
                pass
        else:
            pass

    # wallheaven壁纸主函数
    def wallheaven_wallpaper(self):
        local = self.dictionary_get()
        local_wallheaven = local + "\\" + "wallheaven"
        # 正则匹配获得图片网址
        # .\d+ 匹配出数字id,\d表示数字，+表示多位
        print(local_wallheaven)
        p = 'data-wallpaper-id=.\d+'
        url = 'https://alpha.wallhaven.cc/latest'
        html = self.open_url(url)
        temp_result = []
        if html == None:
            QMessageBox.about(self, "壁纸大师", '网络错误，请检查网络连接后重试')
            imglist = []
        else:
            imglist = re.findall(p, html)
        print(imglist)
        if imglist != []:
            picture_id = imglist[0].replace("data-wallpaper-id=", "").replace("\"", "")
            # 获得中继网址
            temp_url = r"https://alpha.wallhaven.cc/wallpaper/" + str(picture_id)
            temp_html = self.open_url(temp_url)
            if temp_html == None:
                QMessageBox.about(self, "壁纸大师", '网络错误，请检查网络连接后重试')
                temp_result = []
            else:
                q = '<meta property="og:image" content="//wallpapers.wallhaven.cc/wallpapers/full/.*?" />'
                temp_result = re.findall(q, temp_html)
            if temp_result != []:
                temp_result[0] = temp_result[0].replace('<meta property="og:image" content=', '').replace("/>", "").replace(
                    "\"",
                    "")
                print(temp_result[0])
                # 得到图片类型
                filetype = temp_result[0].split('.')[-1]
                print(filetype)
                # 获得图片最终网址
                result = 'https:' + str(temp_result[0])
                print(result)
                # 获得系统时间作为文件名
                time.localtime(time.time())
                # 针对该网站每日多张的特殊情况修改文件命名方式
                n = time.strftime('%Y-%m-%d-%H', time.localtime(time.time()))
                # 设定文件名
                filename = 'wallhaven-' + str(n) + '.' + str(filetype)
                print(filename)
                # 判断当前文件夹是否存在
                if not os.path.exists(local_wallheaven):
                    os.mkdir(local_wallheaven)
                # 判断当前壁纸是否存在
                if os.path.isfile(local_wallheaven + "\\" + filename):
                    print("您今天已经更换过壁纸了")
                else:
                    # 保存壁纸文件，返回文件位置。直接使用urllib.request.urlretrieve函数会被封，必须增加opener
                    print("hhh")
                    opener = urllib.request.build_opener()
                    opener.addheaders = [('User-Agent',
                                          'Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/36.0.1941.0 Safari/537.36')]
                    urllib.request.install_opener(opener)
                    urllib.request.urlretrieve(result, local_wallheaven + "\\" + filename)
                imagepath = local_wallheaven + "\\" + filename
                print(imagepath)
                setWallpaper(imagepath)
            else:
                pass
        else:
            pass

# 设置壁纸函数（共用）
def setWallpaper(imagepath):
    k = win32api.RegOpenKeyEx(win32con.HKEY_CURRENT_USER, "Control Panel\\Desktop", 0, win32con.KEY_SET_VALUE)
    # 2拉伸适应桌面，0桌面居中
    win32api.RegSetValueEx(k, "WallpaperStyle", 0, win32con.REG_SZ, "2")
    win32api.RegSetValueEx(k, "TileWallpaper", 0, win32con.REG_SZ, "0")
    win32gui.SystemParametersInfo(win32con.SPI_SETDESKWALLPAPER, imagepath, 1+2)

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = Window()
    window.show()
    sys.exit(app.exec_())