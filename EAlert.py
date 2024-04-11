# 2022年10月29日开始. 功能是检测屏幕指定区域内是否有指定颜色,如果有就播放声音,闪光提醒.
import sys,threading,os,json,datetime,time
from PyQt5.QtWidgets import QMainWindow,QApplication,QMessageBox,QFileDialog
from PyQt5.QtCore import pyqtSignal,QRegExp,QUrl
from PyQt5.QtGui import QRegExpValidator
from PyQt5 import QtMultimedia
from ctypes import windll
import pyautogui as pag
from Ui_EAlertUI import Ui_Form
 
# 定义全局常量
ISRUN = 0   # 0表示停止, 1表示运行
PARAMDICT = {}

#定义全局函数

#获得鼠标位置的坐标和颜色   返回字典
def getMouseParam():
    posX, posY = pag.position()     #获取位置
    gdi32 = windll.gdi32
    user32 = windll.user32
    hdc = user32.GetDC(None)  # 获取颜色值
    pixel = gdi32.GetPixel(hdc, posX, posY)  # 提取RGB值
    r = pixel & 0x0000ff
    g = (pixel & 0x00ff00) >> 8
    b = pixel >> 16
    return {'x':posX, 'y':posY, 'r':r, 'g':g, 'b':b}


class EAlert( QMainWindow, Ui_Form): 
    #定义一个信号，用于子线程给主线程发信号
    signalThread = pyqtSignal(str, str)     #两个str参数，第一个接收信号类型，第二个接收信号内容

    def __init__(self,parent =None):
        super( EAlert,self).__init__(parent)
        self.setupUi(self)

        # 用正则表达式限定输入 0-9的数字 最大4位
        self.leX1.setValidator(QRegExpValidator(QRegExp("[0-9]{4}"),self))
        self.leY1.setValidator(QRegExpValidator(QRegExp("[0-9]{4}"),self))
        self.leX2.setValidator(QRegExpValidator(QRegExp("[0-9]{4}"),self))
        self.leY2.setValidator(QRegExpValidator(QRegExp("[0-9]{4}"),self))
        self.leR.setValidator(QRegExpValidator(QRegExp("[0-9]{3}"),self))
        self.leG.setValidator(QRegExpValidator(QRegExp("[0-9]{3}"),self))
        self.leB.setValidator(QRegExpValidator(QRegExp("[0-9]{3}"),self))
        self.leInterval.setValidator(QRegExpValidator(QRegExp("[0-9]{4}"),self))
        self.leColorOffset.setValidator(QRegExpValidator(QRegExp("[0-9]{3}"),self))

        #打开配置文件，初始化界面数据
        if os.path.exists( "./EAlert.ini"):
            try:
                iniFileDir = os.getcwd() + "\\"+ "EAlert.ini"
                with open( iniFileDir, 'r', encoding="utf-8") as iniFile:
                    iniDict = json.loads( iniFile.read())
                if iniDict:
                    self.leX1.setText( iniDict['X1'])
                    self.leY1.setText( iniDict['Y1'])
                    self.leX2.setText( iniDict['X2'])
                    self.leY2.setText( iniDict['Y2'])
                    self.leR.setText( iniDict['R'])
                    self.leG.setText( iniDict['G'])
                    self.leB.setText( iniDict['B'])
                    self.leInterval.setText( iniDict['interval'])
                    self.labelFlash.setText( iniDict['flashText'])
                    self.teRecorder.setText( iniDict['recorder'])
                    self.labelAlertDir.setText( iniDict['soundDir'])
                    self.leColorOffset.setText( iniDict['colorOffset'])

                    if iniDict['sound'] == True:
                        self.cbSound.setChecked( True)
                    else:
                        self.cbSound.setChecked( False)

                    if iniDict['flash'] == True:
                        self.cbFlash.setChecked( True)
                    else:
                        self.cbFlash.setChecked( False)

            except:
                QMessageBox.about( self, "提示", "打开初始化文件EAlert.ini异常, 软件关闭时会自动重新创建EAlert.ini文件")


        #绑定槽函数
        self.btnStart.clicked.connect( self.mfStart)
        self.btnStop.clicked.connect( self.mfStop)
        self.btnChooseAlert.clicked.connect( self.mfChooseAlert)
        self.btnNote.clicked.connect( self.mfNote)
        self.btnClear.clicked.connect( self.mfClear)
        self.btnAuto1.clicked.connect( self.mfAuto1)
        self.btnAuto2.clicked.connect( self.mfAuto2)
        self.btnAutoColor.clicked.connect( self.mfAutoColor)

        self.signalThread.connect( self.mfSignal)       # 处理子线程给主线程发的信号

    #自动获得左上角坐标
    def mfAuto1( self):
        time.sleep(3)
        mouseParamDict = getMouseParam()
        self.leX1.setText( str(mouseParamDict['x']))
        self.leY1.setText( str(mouseParamDict['y']))
        QMessageBox.about( self, '提示', '数据已更新, 坐标 ('+ str(mouseParamDict['x'])+', '+ 
                        str(mouseParamDict['y'])+')  RGB('+str(mouseParamDict['r'])+', '+str(mouseParamDict['g'])+', '
                        +str(mouseParamDict['b'])+')')

    #自动获得右下角坐标
    def mfAuto2( self):
        time.sleep(3)
        mouseParamDict = getMouseParam()
        self.leX2.setText( str(mouseParamDict['x']))
        self.leY2.setText( str(mouseParamDict['y']))
        QMessageBox.about( self, '提示', '数据已更新, 坐标 ('+ str(mouseParamDict['x'])+', '+ 
                        str(mouseParamDict['y'])+')  RGB('+str(mouseParamDict['r'])+', '+str(mouseParamDict['g'])+', '
                        +str(mouseParamDict['b'])+')')

    #自动获得鼠标位置颜色
    def mfAutoColor( self):
        time.sleep(3)
        mouseParamDict = getMouseParam()
        self.leR.setText( str(mouseParamDict['r']))
        self.leG.setText( str(mouseParamDict['g']))
        self.leB.setText( str(mouseParamDict['b']))
        QMessageBox.about( self, '提示', '数据已更新, 坐标 ('+ str(mouseParamDict['x'])+', '+ 
                        str(mouseParamDict['y'])+')  RGB('+str(mouseParamDict['r'])+', '+str(mouseParamDict['g'])+', '
                        +str(mouseParamDict['b'])+')')


    # 清空报警记录
    def mfClear( self):
        self.teRecorder.clear()

    # 选择报警的声音
    def mfChooseAlert( self):
        try:
            mFileName, mFileFilt = QFileDialog.getOpenFileName( self, "选择报警声音", os.getcwd(), "wav(*.wav)")
            self.labelAlertDir.setText( mFileName)
        except:
            #打开QFileDialog之后，不选择文件直接关闭，会抛出异常
            pass

    # 按钮  注意事项和免责声明
    def mfNote( self):
        QMessageBox.about( self, '注意事项', '软件思路:判断屏幕指定区域内是否存在指定颜色. \
        \n如果存在播放声音 闪屏报警, \
        \n检测区域坐标点和颜色RGB值可用Snipaste或微信截图工具来获得. \
        \n也可以点击自动获取按钮,3秒后会自动获取鼠标位置的坐标或颜色RGB值.\
        \n色差±值表示指定颜色的范围.例如R=200  ±30,则判断R是否在170--230之间\
        \n被检测的屏幕区域不要遮挡. 检测区域不要选的太大, 否则一次检测会用很长时间.\
        \n报警声音可以用自己的wav音频.\
        \n软件没有对游戏进行任何操作.账号或物品丢失不要找我. \
        \n我也不知道会不会封号.')

    # 处理子线程给主线程发的信号, 信号signalType是字符串'QMessageBox' 'Display' 'alertSound'
    def mfSignal( self, signalType, content):
        if signalType == 'QMessageBox':
            QMessageBox.about( self, "提示", content)

        elif signalType == 'Display':
            self.teRecorder.append( content)

        elif signalType == 'alertSound':
            # 播放警报声音
            url = QUrl.fromLocalFile( self.labelAlertDir.text())
            content = QtMultimedia.QMediaContent(url)
            # --------------------------------------------------------------------------------------------------
            # 注意这里的player前面要加self.   把player加入到这个类中
            # 因为QtMultimedia.QMediaPlayer() 必须在app = QApplication(sys.argv)之后才能播放声音
            # QtMultimedia会在主线程里加一个timer来播放声音, 没有app = QApplication(sys.argv)就没有主线程,没有信号队列
            # 单独写一个QtMultimedia.QMediaPlayer()是不能播放的
            #---------------------------------------------------------------------------------------------------
            self.player = QtMultimedia.QMediaPlayer()
            self.player.setMedia(content)
            self.player.setVolume(100)
            self.player.play()

        elif signalType == 'displayPos':
            self.labelDisplayPos.setText( content)



    # 检查用户输入的所有参数, 有错误返回False, 否则返回True
    def mfCheckParam( self):
        try:
            if int(self.leX1.text()) >= int(self.leX2.text()) or int(self.leY1.text()) >= int(self.leY2.text()):
                self.signalThread.emit( 'QMessageBox', '右下角坐标值应该比左上角大')
                return False
            if int(self.leR.text()) > 255 or int(self.leG.text()) > 255 or int(self.leB.text()) > 255 or int(self.leColorOffset.text()) > 255:
                self.signalThread.emit( 'QMessageBox', 'RGB颜色值和色差范围值应该在0~255之间')
                return False
            if not os.access(self.labelAlertDir.text(), os.R_OK):       #判断文件是否存在并可读
                self.signalThread.emit( 'QMessageBox', '报警声音不可用,请重新选择')
                return False
            if self.labelAlertDir.text()[-3:] != 'wav':     #判断是否wav格式
                self.signalThread.emit( 'QMessageBox', '请选择wav格式的报警声音')
                return False
            if self.leInterval.text() == '' or self.leInterval.text() == '0':
                self.signalThread.emit( 'QMessageBox', '检测间隔时间错误')
                return False

        except:
            self.signalThread.emit( 'QMessageBox', '坐标或RGB值错误')
            return False

        return True

    # 点击停止按钮，在执行完当前任务后停止
    def mfStop( self):
        global ISRUN
        ISRUN = 0
        self.signalThread.emit('Display', '当前区域检测结束后暂停')

    # 开始运行
    def mfStart( self):
        global ISRUN
        global PARAMDICT
        ISRUN = 1

        # 检查用户输入的所有参数, mfCheckParam()有错误返回False, 否则返回True
        if self.mfCheckParam() == False:
            return      # mfCheckParam检测出来错误,就直接退出函数,不执行后面的语句

        #保存界面上的参数到全局字典PARAMDICT, 在mfRun中要用
        PARAMDICT.clear()
        PARAMDICT['intX1'] = int(self.leX1.text())
        PARAMDICT['intY1'] = int(self.leY1.text())
        PARAMDICT['intX2'] = int(self.leX2.text())
        PARAMDICT['intY2'] = int(self.leY2.text())
        PARAMDICT['intR'] = int(self.leR.text())
        PARAMDICT['intG'] = int(self.leG.text())
        PARAMDICT['intB'] = int(self.leB.text())
        PARAMDICT['intInterval'] = int(self.leInterval.text())
        PARAMDICT['intColorOffset'] = int(self.leColorOffset.text())

        # 创建一个新线程
        inRunThreading = threading.Thread( target= self.mfRun)
        inRunThreading.start()
        # inRunThreading.join()     join()是将指定线程加入当前线程，将两个交替执行的线程转换成顺序执行。
        #                           把inRunThreading加入到主线程，会引起主线程阻塞，这里不能.join()


    # 核心代码**********************************
    def mfRun(self):
        global ISRUN
        global PARAMDICT

        if ISRUN == 0:
            return

        # 本来想用截图方案,后来用了直接遍历屏幕像素点的方法
        # 这里要 import pyautogui as pag          from PIL import Image
        # img = pag.screenshot( region = [ PARAMDICT['intX1'], PARAMDICT['intY1'], \
        #                         PARAMDICT['intX2'] - PARAMDICT['intX1'], PARAMDICT['intY2'] - PARAMDICT['intY1'] ])
        # #感觉截图没有必要保存下来, 直接判断颜色就可以了
        # img.save(os.getcwd() + "\\"+ "EAlert.png")

        hdc = windll.user32.GetDC(None)

        exitFlag = False
        for i in range( PARAMDICT['intX1'], PARAMDICT['intX2']):
            for j in range( PARAMDICT['intY1'], PARAMDICT['intY2']):
                #在界面上显示当前检测点的坐标
                posStr = str(i) + ', ' +str(j)
                self.signalThread.emit( 'displayPos', posStr)


                pixel = windll.gdi32.GetPixel( hdc, i, j)
                r = pixel & 0x0000ff
                g = (pixel & 0x00ff00) >> 8
                b = pixel >> 16

                # #如果找到指定的RGB颜色
                # if r == PARAMDICT['intR'] and g == PARAMDICT['intG'] and b == PARAMDICT['intB']:
                #这里把判断RGB的值是否等于指定值,改为是否在指定范围
                if r <= PARAMDICT['intR'] + PARAMDICT['intColorOffset'] and r >= PARAMDICT['intR'] - PARAMDICT['intColorOffset'] \
                    and g <= PARAMDICT['intG'] + PARAMDICT['intColorOffset'] and g >= PARAMDICT['intG'] - PARAMDICT['intColorOffset'] \
                    and b <= PARAMDICT['intB'] + PARAMDICT['intColorOffset'] and b >= PARAMDICT['intB'] - PARAMDICT['intColorOffset']:

                    #警报记录窗口添加一条警报
                    self.signalThread.emit( 'Display', datetime.datetime.now().strftime('%H:%M:%S') + '  警报')

                    #播放警报声音
                    if self.cbSound.isChecked():
                        self.signalThread.emit( 'alertSound', '第二个参数没有用')

                    #闪烁警报区域
                    if self.cbFlash.isChecked():
                        for flashCount in range( 0, 10):
                            self.labelFlash.setStyleSheet("background-color: rgb(255, 0, 0)")
                            time.sleep(0.3)
                            self.labelFlash.setStyleSheet("background-color: rgb(0, 0, 255)")
                            time.sleep(0.3)
                        #闪烁之后还改回原来的颜色
                        self.labelFlash.setStyleSheet("background-color: rgb(240, 240, 240)")

                    exitFlag = True
                    break

            if exitFlag == True:
                break

        tempTimer = threading.Timer( PARAMDICT['intInterval'], self.mfRun)
        tempTimer.start()

        
#主程序入口
if __name__ == "__main__":
    app = QApplication(sys.argv)
    myWin = EAlert()
    myWin.show()

    appExit = app.exec_()
    #退出程序之前，保存界面上的设置
    tempDict = { 'X1':myWin.leX1.text(), 'Y1':myWin.leY1.text(), 'X2':myWin.leX2.text(), 'Y2':myWin.leY2.text(),
                'R':myWin.leR.text(), 'G':myWin.leG.text(), 'B':myWin.leB.text(), 'soundDir':myWin.labelAlertDir.text(),
                'interval':myWin.leInterval.text(), 'flashText':myWin.labelFlash.text(), 
                'recorder':myWin.teRecorder.toHtml(), 'colorOffset':myWin.leColorOffset.text(),
                'sound':myWin.cbSound.isChecked(), 'flash':myWin.cbFlash.isChecked()
    }
    saveIniJson = json.dumps( tempDict, indent=4)
    try:
        saveIniFile = open( "./EAlert.ini", "w",  encoding="utf-8")
        saveIniFile.write( saveIniJson)
        saveIniFile.close()
    except:
        QMessageBox.about( myWin, "提示", "保存配置文件EAlert.ini失败")

    # 这一句特别重要, 程序是两个线程在运行, 关闭窗口只能结束主线程, 子线程还在运行. 
    # 创建子线程的标志ISRUN 一定要改成0, 子线程在检测ISRUN==0之后,就不再用Timer创建新的线程了
    ISRUN = 0

    sys.exit( appExit)
# sys.exit(app.exec_())  