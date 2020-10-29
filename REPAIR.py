#!-*-coding:utf-8 -*-
#!@Time   :2020-5-11  16:48
#!@Author : yuhao_liu
#!@File   : .py

import sys
import time
import serial.tools.list_ports
from PyQt5 import QtWidgets,QtGui
from PyQt5.QtWidgets import QMessageBox
from AUV_2020_UI import Ui_Form
from PyQt5.QtCore import *
import xlwt
import math
import struct
from mode import Mode

#根据经纬度计算路程函数
def haversine(lat1, long1, lat2, long2):
    from math import radians, sin, cos, atan2, sqrt
    lat1, long1, lat2, long2 = map(radians, [lat1, long1, lat2, long2])
    a = sin((lat1-lat2)/2)**2 + cos(lat1)*cos(lat2)*(sin((long1-long2)/2)**2)
    c = 2 * atan2(sqrt(a), sqrt(1-a))
    return 6371 * c

#转换浮点函数
def ReadFloat(*args,reverse=False):
    for n,m in args:
        n,m = '%04x'%n,'%04x'%m
    if reverse:
        v = n + m
    else:
        v = m + n
    y_bytes = bytes.fromhex(v)
    y = struct.unpack('!f',y_bytes)[0]
    y = round(y,6)
    return y


class auv2(QtWidgets.QWidget,Ui_Form):
    # 创建信号槽
    sig_1 = pyqtSignal()

    def __init__(self):

        #窗口提示
        self.message_box_state = 0

        #关闭按钮滑台x、y回归初始值，鱼尾停止摆动
        self.step_x_close = 95
        self.step_y_close = 60

        self.fish_stop = 0

        #控制数据初始化
        #打开状态
        self.open_state = 0
        #开始工作状态
        self.start_state = 0
        #运行状态
        self.run_state = 0

        #油门控制
        self.you = 0

        #滑台位置-X
        self.step_x = 0
        #滑台位置-Y
        self.step_y = 0

        #转弯快慢速油门控制
        self.tail_wan_max = 0
        self.tail_wan_min = 0


        #接收数据初始化
        #结构体数据长度
        self.uart_len = 184
        #接收总数据
        self.num = 0
        #惯导横滚角
        self.roll = 0
        #惯导俯仰角
        self.pitch = 0
        #惯导偏航角
        self.yaw = 0
        #惯导速度
        self.v = 0
        #惯导GPS经度
        self.gps_jd = 0
        #惯导GPS纬度
        self.gps_wd = 0

        #根据GPS经纬度计算路程，进而计算速度，需到水面
        #路程计算起点经纬度，时间
        self.start_gps_jd = 0
        self.start_gps_wd = 0
        self.start_gps_time = 0
        #路程计算终点经纬度，时间
        self.end_gps_jd = 0
        self.end_gps_wd = 0
        self.end_gps_time = 0
        #根据GPS经纬度算出的路程
        self.gps_x = 0
        #根据GPS计算起点终点的时间差
        self.gps_start_end_time = 0
        #根据GPS路程算出的实际前进速度
        self.gps_v = 0

        #驱动油门
        self.dr_you = 0
        #驱动转速
        self.dr_rpm = 0
        #驱动控制器温度
        self.dr_ctr_tem = 0
        #驱动电机温度
        self.dr_mot_tem = 0
        #驱动电压
        self.dr_vol = 0
        #驱动电流
        self.dr_cur = 0
        #驱动里程
        self.dr_licheng = 0
        #驱动故障码
        self.dr_errcode = 0

        #保护板电压
        self.bms_vol = 0
        #保护板电流
        self.bms_cur = 0
        #保护板功率
        self.bms_pw = 0
        #保护板温度
        self.bms_tem = 0
        #保护板电量百分比
        self.bms_percent = 0

        #滑台x
        self.step_x_disp = 0
        #滑台y
        self.step_y_disp = 0

        #鱼尾位置
        self.fish_pos = 0
        #鱼尾摆动频率
        self.fish_frequency = 0
        #鱼尾下潜深度
        self.fish_deep = 0
        #进水检测标志
        self.water_state_0 = 0
        self.water_state_1 = 0
        self.water_state_2 = 0

        #预留位用于查看空闲邮箱数
        self.tx_mail_num = 0

        #校验位
        self.cksum = 0

        #内部时间
        self.time_cnt = 0

        #ht控制标志
        self.ht_run_flag = 0

        self.you_ctr = 0

        #水深报警
        self.deep_error = 20

        #EXCEL文件保存
        #显示数据
        self.excel = xlwt.Workbook(encoding='utf-8')
        self.sheet = self.excel.add_sheet('Sheet1')

        self.sheet.write(0, 0, u'保护板-电量百分比-%')
        self.sheet.write(0, 1, u'保护板-电压-V')
        self.sheet.write(0, 2, u'保护板-电流-A')
        self.sheet.write(0, 3, u'保护板-功率-W')
        self.sheet.write(0, 4, u'保护板-温度-°C')

        self.sheet.write(0, 5, u'惯导-横滚角-°')
        self.sheet.write(0, 6, u'惯导-俯仰角-°')
        self.sheet.write(0, 7, u'惯导-偏航角-°')
        self.sheet.write(0, 8, u'惯导-速度-M/S')
        self.sheet.write(0, 9, u'惯导-经度')
        self.sheet.write(0, 10, u'惯导-纬度')

        self.sheet.write(0, 11, u'GPS-计算路程')
        self.sheet.write(0, 12, u'GPS-计算时间差')
        self.sheet.write(0, 13, u'GPS-计算速度')

        self.sheet.write(0, 14, u'滑台-X-MM')
        self.sheet.write(0, 15, u'滑台-Y-MM')

        self.sheet.write(0, 16, u'水深-M')

        self.sheet.write(0, 17, u'进水报警-1')
        self.sheet.write(0, 18, u'进水报警-2')
        self.sheet.write(0, 19, u'进水报警-3')

        self.sheet.write(0, 20, u'鱼尾控制状态')
        self.sheet.write(0, 21, u'驱动-油门')
        self.sheet.write(0, 22, u'驱动-转速-RPM')
        self.sheet.write(0, 23, u'驱动-控制器温度-°C')
        self.sheet.write(0, 24, u'驱动-电机温度-°C')
        self.sheet.write(0, 25, u'驱动-电压-V')
        self.sheet.write(0, 26, u'驱动-电流-A')
        self.sheet.write(0, 27, u'驱动-里程-RP')
        self.sheet.write(0, 28, u'驱动-故障码')

        self.sheet.write(0, 29, u'驱动-鱼尾位置-°')
        self.sheet.write(0, 30, u'驱动-摆尾频率-HZ')

        #控制数据
        self.sheet.write(0, 31, u'滑台控制-X-MM')
        self.sheet.write(0, 32, u'滑台控制-Y-MM')

        self.sheet.write(0, 33, u'系统时间-MS')

        self.sheet.write(0, 34, u'CAN可用的发送邮箱数-个')

        self.e_i = 1
        self.length = 35

        #文件保存名字
        self.excel_name = 'data_' + str(time.localtime().tm_year) + '-' + \
                        str(time.localtime().tm_mon) + '-' + str(time.localtime().tm_mday) + \
                        '-' + str(time.localtime().tm_hour) + '-' + str(time.localtime().tm_min) + \
                        '-' + str(time.localtime().tm_sec) + '.xls'

        #界面初始化
        super(auv2,self).__init__()
        self.setupUi(self)
        self.init()
        self.setWindowTitle('AUV-2020控制系统')

        self.label_25.setText('停止')

        #创建定时器接收和发送数据，并保存数据
        self.timer_receive = QTimer(self)
        self.timer_receive.timeout.connect(self.received)
        self.timer_receive.start(15)

        #创建串口
        self.ser = serial.Serial()

        #获取模式设定窗口参数
        self.Window = Mode()

    #功能初始化
    def init(self):
        #按键功能初始化
        self.pushButton.clicked.connect(self.opened)
        self.pushButton_2.clicked.connect(self.closed)
        self.pushButton_4.clicked.connect(self.moved)
        self.pushButton_3.clicked.connect(self.stoped)
        self.pushButton_5.clicked.connect(self.lefted)
        self.pushButton_6.clicked.connect(self.righted)
        self.pushButton_7.clicked.connect(self.mode)
        self.pushButton_8.clicked.connect(self.gps_x_start)
        self.pushButton_9.clicked.connect(self.gps_x_end)

        self.pushButton_11.clicked.connect(self.ctr_inited)

        self.sig_1.connect(self.sig_1_slot)

        #5个灯初始化
        pixmap_1 = QtGui.QPixmap("D:/green.jpg")
        self.label_10.setPixmap(pixmap_1)
        self.label_11.setPixmap(pixmap_1)
        self.label_12.setPixmap(pixmap_1)
        self.label_13.setPixmap(pixmap_1)
        self.label_14.setPixmap(pixmap_1)


    def opened(self):

        self.ser.port = 'COM8'
        self.ser.baudrate = 115200
        self.ser.bytesize = 8
        self.ser.parity = 'N'
        self.ser.stopbits = 1

        try :
            self.ser.open()

            self.pushButton.setEnabled(False)
            self.pushButton_2.setEnabled(True)
            self.pushButton_3.setEnabled(True)
            self.pushButton_4.setEnabled(True)
            self.pushButton_5.setEnabled(True)
            self.pushButton_6.setEnabled(True)
            self.pushButton_8.setEnabled(True)
            self.pushButton_9.setEnabled(True)
            self.pushButton_11.setEnabled(True)

            #发送控制数据状态
            self.open_state = 1

        except:

            #发送控制数据状态
            self.open_state = 0


    def closed(self):
        try :

            cksum = ( ( 85 + 170 + 84 ) + ( self.step_x_close & 0xff ) + ( self.step_x_close >> 8 ) +
                                        ( self.step_y_close & 0xff ) + ( self.step_y_close >> 8 ) ) & 0xff

            #发送鱼尾停止命令
            self.ser.write(b'\x55\xAA\x54\x00' +
                           b'\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00' +
                           b'\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00' +
                           #step_x_l and step_x_h
                           bytes([self.step_x_close & 0xff]) + bytes([self.step_x_close >> 8]) +
                           #step_y_l and step_y_h
                           bytes([self.step_y_close & 0xff]) + bytes([self.step_y_close >> 8]) +
                           b'\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00' +
                           b'\x00\x00\x00\x00\x00\x00\x00\x00\x00' +
                           bytes([cksum]))

            # 发送控制数据状态
            self.open_state = 0
            self.start_state = 0
            self.ser.close()

        except :
            pass

        self.pushButton.setEnabled(True)
        self.pushButton_2.setEnabled(False)
        self.pushButton_3.setEnabled(False)
        self.pushButton_4.setEnabled(False)
        self.pushButton_5.setEnabled(False)
        self.pushButton_6.setEnabled(False)
        self.pushButton_8.setEnabled(False)
        self.pushButton_9.setEnabled(False)

        self.pushButton_11.setEnabled(False)

    #设定鱼尾前进
    def moved(self):
        #往复摆动
        self.label_25.setText('前进')
        # 发送控制数据状态
        self.run_state = 1
        #self.start_state = 1

        #急停指示灯变红色
        pixmap_4 = QtGui.QPixmap("D:/green.jpg")
        self.label_13.setPixmap(pixmap_4)

        self.pushButton_4.setEnabled(False)
        self.pushButton_5.setEnabled(True)
        self.pushButton_6.setEnabled(True)


    #设定鱼尾停止
    def stoped(self):
        self.label_25.setText('停止')
        # 发送控制数据状态
        self.run_state = 0
        #self.start_state = 1

        self.you_ctr = 0
        self.verticalSlider.setValue(0)
        self.spinBox.setValue(0)

        #急停指示灯变红色
        pixmap_4 = QtGui.QPixmap("D:/red.jpg")
        self.label_13.setPixmap(pixmap_4)

        self.pushButton_4.setEnabled(True)
        self.pushButton_5.setEnabled(True)
        self.pushButton_6.setEnabled(True)

        self.Window.selfmode_flag = 0

    #设定鱼尾左转
    def lefted(self):
        self.label_25.setText('左转')

        # 发送控制数据状态
        self.run_state = 2      #左右转对调一下
        #self.start_state = 1

        #急停指示灯变绿色
        pixmap_4 = QtGui.QPixmap("D:/green.jpg")
        self.label_13.setPixmap(pixmap_4)

        self.pushButton_4.setEnabled(True)
        self.pushButton_5.setEnabled(False)
        self.pushButton_6.setEnabled(True)

    #设定鱼尾右转
    def righted(self):
        self.label_25.setText("右转")
        # 发送控制数据状态
        self.run_state = 3      #左右转对调一下
        #self.start_state = 1

        #急停指示灯变绿色
        pixmap_4 = QtGui.QPixmap("D:/green.jpg")
        self.label_13.setPixmap(pixmap_4)

        self.pushButton_4.setEnabled(True)
        self.pushButton_5.setEnabled(True)
        self.pushButton_6.setEnabled(False)

    #设定自主模式运行
    def mode(self):
        self.sig_1.emit()

    #设定自主模式中的线槽
    def sig_1_slot(self):
        #设定模式窗口显示
        self.Window.show()
        self.label_25.setText('自主模式')

        #急停指示灯变绿色
        pixmap_4 = QtGui.QPixmap("D:/green.jpg")
        self.label_13.setPixmap(pixmap_4)

    #GPS模式保存起点经纬度和时间
    def gps_x_start(self):

        self.start_gps_jd = self.gps_jd
        self.start_gps_wd = self.gps_wd

        self.start_gps_time = time.time()

        self.pushButton_8.setEnabled(False)
        self.pushButton_9.setEnabled(True)


    #GPS模式保存终点经纬度和时间，并计算路程和速度
    def gps_x_end(self):

        self.end_gps_jd = self.gps_jd
        self.end_gps_wd = self.gps_wd

        self.end_gps_time = time.time()
        #计算路程
        self.gps_x = round( haversine(self.start_gps_wd,self.start_gps_jd,self.end_gps_wd,self.end_gps_jd) * 1000 ,2 )
        self.lineEdit_4.setText(str(self.gps_x))

        #计算时间差
        self.gps_start_end_time = round( ( self.end_gps_time - self.start_gps_time ),2)
        #计算速度
        self.gps_v = round( self.gps_x / self.gps_start_end_time , 2 )
        self.lineEdit_5.setText(str(self.gps_v))

        self.pushButton_8.setEnabled(True)
        self.pushButton_9.setEnabled(False)

        #弹出GPS路程和速度计算完成窗口
        QMessageBox.warning(self, '完成', 'GPS计算路程和速度已完成')

    #控制参数初始化
    def ctr_inited(self):
        self.run_state = 0

        self.you_ctr = 0
        self.verticalSlider.setValue(0)
        self.spinBox.setValue(0)

        self.lineEdit_36.setText( str(self.step_x_close) )
        self.lineEdit_41.setText( str(self.step_y_close) )

        self.spinBox_2.setValue(int(0))
        self.spinBox_3.setValue(int(0))

        self.tail_wan_max = 0
        self.tail_wan_min = 0
        self.horizontalSlider.setValue(0)
        self.horizontalSlider_2.setValue(0)

    # 实时检测键盘按键
    def keyPressEvent(self, event):
        '''
        if event.key() == Qt.Key_W:
            # 控制数据设定
            self.label_25.setText('前进')
            self.run_state = 1
            self.you_ctr = int(self.verticalSlider.value()) + 2
            if self.you_ctr > 200 :
                self.you_ctr = 0

            if self.you_ctr > 0 :
                # 急停指示灯变绿色
                pixmap_4 = QtGui.QPixmap("D:/green.jpg")
                self.label_13.setPixmap(pixmap_4)

        elif event.key() == Qt.Key_S:
            # 控制数据设定
            self.label_25.setText('前进')
            self.run_state = 1
            self.you_ctr = int(self.verticalSlider.value())  - 2
            if self.you_ctr < 0 :
                self.you_ctr = 0

            if self.you_ctr > 0:
                # 急停指示灯变绿色
                pixmap_4 = QtGui.QPixmap("D:/green.jpg")
                self.label_13.setPixmap(pixmap_4)
    '''
        if event.key() == Qt.Key_Q:
            #急停键盘控制
            self.run_state = 0
            self.label_25.setText('停止')
            # 发送控制数据状态
            #self.run_state = 0
            #self.start_state = 1

            self.you_ctr = 0
            self.verticalSlider.setValue(0)
            self.spinBox.setValue(0)

            # 急停指示灯变红色
            pixmap_4 = QtGui.QPixmap("D:/red.jpg")
            self.label_13.setPixmap(pixmap_4)

            self.pushButton_4.setEnabled(True)
            self.pushButton_5.setEnabled(True)
            self.pushButton_6.setEnabled(True)

            self.Window.selfmode_flag = 0

        else:
            self.you_ctr = self.you_ctr
            pass

        self.verticalSlider.setValue(self.you_ctr)
        self.spinBox.setValue(self.you_ctr)

    #接收显示数据
    def received(self):
        try:
            #15ms定时，内部加15ms用作时间基准
            self.time_cnt = self.time_cnt + 15
            #15ms接收数据计数
            # 通电状态，接收数据显示
            if self.open_state == 1:
                self.num = self.ser.inWaiting()
                if self.num >= self.uart_len:
                    self.data = self.ser.read(self.num)
                    if len(self.data) >= self.uart_len :

                        for num_i in range(self.uart_len):
                            try:
                                if self.data[num_i+0] == 85 and self.data[num_i+1] == 170 and self.data[num_i+2] == 92:
                                    data_list = []
                                    ck_sum = 0
                                    for i in range(num_i,num_i+91):
                                        ck_sum = ck_sum + self.data[i]
                                        data_list.append(self.data[i])
                                    data_list.append(self.data[num_i + 91])
                                    ck_sum = ck_sum & 0xff

                                    #后续数据换成完整一帧，data_list，92个字节
                                    if ck_sum == data_list[91]:

                                        if self.time_cnt % 15000 == 0:
                                            # 弹出5次进水报警窗口，等待10s后，继续弹出窗口
                                            self.message_box_state = 0

                                        if( ( data_list[84] == 1 ) and ( self.message_box_state < 5 ) ):
                                            QMessageBox.critical(self, 'error', '头部进水，请及时检查')
                                            self.message_box_state = self.message_box_state + 1

                                            #进水指示灯变红色
                                            pixmap_1 = QtGui.QPixmap("D:/red.jpg")
                                            self.label_10.setPixmap(pixmap_1)

                                        elif( ( data_list[85] == 1 ) and ( self.message_box_state < 5 ) ):
                                            QMessageBox.critical(self, 'error', '中仓左部进水，请及时检查')
                                            self.message_box_state = self.message_box_state + 1

                                            # 进水指示灯变红色
                                            pixmap_1 = QtGui.QPixmap("D:/red.jpg")
                                            self.label_10.setPixmap(pixmap_1)

                                        elif( ( data_list[86] == 1 ) and ( self.message_box_state < 5 ) ):
                                            QMessageBox.critical(self, 'error', '中仓右部进水，请及时检查')
                                            self.message_box_state = self.message_box_state + 1

                                            # 进水指示灯变红色
                                            pixmap_1 = QtGui.QPixmap("D:/red.jpg")
                                            self.label_10.setPixmap(pixmap_1)

                                        else :
                                            # 进水指示灯变绿色
                                            pixmap_1 = QtGui.QPixmap("D:/green.jpg")
                                            self.label_10.setPixmap(pixmap_1)

                                        #惯导横滚角
                                        #AUV-2020
                                        roll_l = data_list[4] + (data_list[5] << 8)
                                        roll_h = data_list[6] + (data_list[7] << 8)
                                        data = (roll_l, roll_h)
                                        self.roll = ReadFloat(data)

                                        if self.roll < -90:
                                            self.roll = self.roll + 180
                                        elif self.roll > 90:
                                            self.roll = self.roll - 180
                                        else:
                                            pass

                                        self.lineEdit.setText(str(round(self.roll, 2)))
                                        #self.horizontalSlider.setValue(self.roll)

                                        # 惯导俯仰角
                                        #AUV-2020
                                        pitch_l = data_list[8] + (data_list[9] << 8)
                                        pitch_h = data_list[10] + (data_list[11] << 8)
                                        data = (pitch_l, pitch_h)
                                        self.pitch = ReadFloat(data)
                                        self.lineEdit_2.setText(str(round(self.pitch, 2)))
                                        #self.verticalSlider.setValue(self.pitch)

                                        # 惯导偏航角
                                        #AUV-2020
                                        yaw_l = data_list[12] + (data_list[13] << 8)
                                        yaw_h = data_list[14] + (data_list[15] << 8)
                                        data = (yaw_l, yaw_h)
                                        self.yaw = ReadFloat(data)
                                        self.lineEdit_3.setText(str(round(self.yaw, 2)))
                                        #self.dial.setValue(-self.yaw)

                                        # 惯导速度
                                        #AUV-2020
                                        v_north_l = data_list[16] + (data_list[17] << 8)
                                        v_north_h = data_list[18] + (data_list[19] << 8)
                                        data = (v_north_l, v_north_h)
                                        self.v_north = ReadFloat(data)

                                        v_east_l = data_list[24] + (data_list[25] << 8)
                                        v_east_h = data_list[26] + (data_list[27] << 8)
                                        data = (v_east_l, v_east_h)
                                        self.v_east = ReadFloat(data)

                                        self.v = math.sqrt(self.v_north * self.v_north + self.v_east * self.v_east)
                                        self.lineEdit_32.setText(str(round(self.v,3)))

                                        #姿态灯设置
                                        if self.roll < -40 or self.roll > 40 or self.pitch < -40 or self.pitch > 40 :
                                            #姿态指示灯变绿色
                                            pixmap_2 = QtGui.QPixmap("D:/yellow.jpg")
                                            self.label_11.setPixmap(pixmap_2)
                                        else:
                                            # 进水指示灯变绿色
                                            pixmap_2 = QtGui.QPixmap("D:/green.jpg")
                                            self.label_11.setPixmap(pixmap_2)

                                        #惯导GPS经度
                                        self.gps_jd =  ( (data_list[28] + (data_list[29] << 8) + (data_list[30] << 16) + (data_list[31] << 24) ) / 10000000 )
                                        #self.lineEdit_7.setText(str(round(self.gps_jd,2)))

                                        #惯导GPS纬度
                                        self.gps_wd =  ( (data_list[32] + (data_list[33] << 8) + (data_list[34] << 16) + (data_list[35] << 24) ) / 10000000 )
                                        #self.lineEdit_6.setText(str(round(self.gps_wd,2)))


                                        # 保护板电压
                                        self.bms_vol = str( (data_list[36] + (data_list[37] << 8) + (data_list[38] << 16) + (data_list[39] << 24) ) / 10 )
                                        self.lineEdit_28.setText(self.bms_vol)

                                        # 保护板电流
                                        self.bms_cur = (data_list[40] + (data_list[41] << 8) + (data_list[42] << 16) + (data_list[43] << 24 ))

                                        if self.bms_cur > 32767:
                                            self.bms_cur = (self.bms_cur - 65536) / 10
                                        else:
                                            self.bms_cur = self.bms_cur / 10
                                        self.lineEdit_29.setText(str(self.bms_cur))

                                        # 保护板功率
                                        self.bms_pw = (data_list[44] + (data_list[45] << 8) + (data_list[46] << 16) + (data_list[47] << 24))
                                        if self.bms_pw > 32767:
                                            self.bms_pw = (self.bms_pw - 65536)
                                        else:
                                            self.bms_pw = self.bms_pw
                                        self.lineEdit_33.setText(str(self.bms_pw))

                                        # 保护板电量百分比
                                        self.bms_percent = str(data_list[48])
                                        self.lineEdit_31.setText(self.bms_percent)

                                        # 保护板温度-short
                                        self.bms_tem = (data_list[50] + (data_list[51] << 8))
                                        if self.bms_tem > 32767:
                                            self.bms_tem = (self.bms_tem - 65536)
                                        else:
                                            self.bms_tem = self.bms_tem
                                        self.lineEdit_30.setText(str(self.bms_tem))

                                        # 滑台x-unsigned short
                                        self.step_x_disp = (data_list[52] + (data_list[53] << 8))
                                        #self.horizontalSlider_3.setValue(self.step_x_disp)
                                        self.lineEdit_43.setText(str(self.step_x_disp))

                                        # 滑台y-unsigned short
                                        self.step_y_disp = (data_list[54] + (data_list[55] << 8))
                                        #self.verticalSlider_2.setValue(self.step_y_disp)
                                        self.lineEdit_44.setText(str(self.step_y_disp))

                                        self.lineEdit_41.setText(str(self.step_y_disp))

                                        # 鱼尾下潜深度-short
                                        self.fish_deep = (data_list[60] + (data_list[61] << 8))
                                        if self.fish_deep > 32767:
                                            self.fish_deep = (self.fish_deep - 65536) / 1000
                                        else:
                                            self.fish_deep = self.fish_deep / 1000
                                        self.lineEdit_38.setText(str(self.fish_deep))

                                        # 驱动转速
                                        self.dr_rpm = data_list[70] + (data_list[71] << 8)
                                        self.lineEdit_45.setText(str(self.dr_rpm))

                                        #驱动电机温度
                                        self.dr_mot_tem = data_list[72] - 40
                                        if self.dr_mot_tem < 0 :
                                            self.dr_mot_tem = 0
                                        self.lineEdit_35.setText(str(self.dr_mot_tem))

                                        #控制器温度
                                        self.dr_ctr_tem = data_list[73] - 40
                                        if self.dr_ctr_tem < 0 :
                                            self.dr_ctr_tem = 0
                                        self.lineEdit_46.setText(str(self.dr_ctr_tem))

                                        #驱动电压
                                        #self.dr_vol = (data_list[74] + (data_list[75] << 8))
                                        #self.dr_vol = self.dr_vol / 10
                                        #self.lineEdit_47.setText(str(self.dr_vol))

                                        # 驱动电流
                                        #self.dr_cur = (data_list[76] + (data_list[77] << 8))
                                        #self.dr_cur = self.dr_cur / 10
                                        #self.lineEdit_50.setText(str(self.dr_cur))

                                        #驱动里程
                                        #self.dr_licheng = ( data_list[78] + (data_list[79] << 8) )
                                        #if self.dr_licheng < 0 :
                                        #    self.dr_licheng = 0
                                        #self.lineEdit_51.setText(str(self.dr_licheng))

                                        #驱动故障码
                                        #self.dr_errcode = data_list[80] + (data_list[81] << 8)
                                        #self.lineEdit_52.setText(str(self.dr_errcode))

                                        #auv内部时间
                                        self.auv_time = int(data_list[76] + (data_list[77] << 8) + (data_list[78] << 16) + (data_list[79] << 24))

                                        #驱动鱼尾位置
                                        self.fish_pos = (data_list[88] + (data_list[89] << 8))
                                        self.fish_pos = round( ( self.fish_pos - 21000 ) / 1000 , 2 )

                                        self.lineEdit_37.setText(str(self.fish_pos))
                                        #self.horizontalSlider_2.setValue(self.fish_pos)


                                        #空闲发送邮箱数
                                        self.tx_mail_num = data_list[90]
                                        if self.tx_mail_num < 3 :
                                            #正常指示灯变橙色
                                            pixmap_3 = QtGui.QPixmap("D:/orange.jpg")
                                            self.label_12.setPixmap(pixmap_3)
                                        else:
                                            #正常指示灯变绿色
                                            pixmap_3 = QtGui.QPixmap("D:/green.jpg")
                                            self.label_12.setPixmap(pixmap_3)
                                        print(self.tx_mail_num)

                                        #驱动鱼尾摆尾频率
                                        self.fish_frequency = round( self.dr_rpm / 60 / 10 , 2 )
                                        self.lineEdit_40.setText(str(self.fish_frequency))

                                        #接收到完整帧数据后，发送控制数据，进而保存数据
                                        #通电状态
                                        if self.open_state == 1:

                                            if self.run_state == 1 :
                                                self.you_ctr = int(self.verticalSlider.value()) * 2
                                                self.spinBox.setValue(int(self.verticalSlider.value()))

                                                self.tail_wan_max = 0

                                            elif self.run_state == 2 or self.run_state == 3 :
                                                self.you_ctr = 0 #int(self.verticalSlider.value()) * 2
                                                self.tail_wan_max = int(self.verticalSlider.value()) * 2
                                                self.spinBox.setValue(int(self.verticalSlider.value()))

                                            else:
                                                pass

                                            # 触底指示灯设置
                                            if self.fish_deep > self.deep_error:

                                                # 触底指示灯变蓝色
                                                pixmap_5 = QtGui.QPixmap("D:/blue.jpg")
                                                self.label_14.setPixmap(pixmap_5)

                                                # 急停控制
                                                self.run_state = 0
                                                self.label_25.setText('停止')

                                                self.you_ctr = 0
                                                self.verticalSlider.setValue(0)
                                                self.spinBox.setValue(0)

                                                # 急停指示灯变红色
                                                pixmap_4 = QtGui.QPixmap("D:/red.jpg")
                                                self.label_13.setPixmap(pixmap_4)

                                                self.pushButton_4.setEnabled(True)
                                                self.pushButton_5.setEnabled(True)
                                                self.pushButton_6.setEnabled(True)

                                            else:
                                                # 触底指示灯变绿色
                                                pixmap_5 = QtGui.QPixmap("D:/green.jpg")
                                                self.label_14.setPixmap(pixmap_5)


                                            # 滑台位置-X
                                            self.step_x = int((self.lineEdit_36.text()))
                                            # 滑台位置-Y
                                            self.step_y =  int(self.horizontalSlider.value()) #int((self.lineEdit_41.text()))

                                            #转弯快速
                                            #self.tail_wan_max = int(self.horizontalSlider.value())
                                            #self.spinBox_2.setValue(int(self.horizontalSlider.value()))

                                            #转弯慢速
                                            self.tail_wan_min = 0 #int(self.horizontalSlider_2.value())
                                            #self.spinBox_3.setValue(int(self.horizontalSlider_2.value()))

                                            #计算校验和
                                            self.cksum = ( 85 + 170 + 84 + self.run_state + ( self.step_x >> 8 ) + (self.step_x & 0xff) +
                                                         (self.step_y >> 8) + (self.step_y & 0xff) + self.you_ctr + self.tail_wan_max   +
                                                         self.tail_wan_min ) & 0x00ff

                                            #写数据到串口
                                            self.ser.write(
                                                b'\x55\xAA\x54\x00' +
                                                b'\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00' +
                                                b'\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00' +
                                                bytes([self.step_x & 0xff]) + bytes([self.step_x >> 8]) +
                                                bytes([self.step_y & 0xff]) + bytes([self.step_y >> 8]) +
                                                b'\x00\x00\x00\x00' +
                                                bytes([self.run_state & 0xff]) + bytes([self.run_state >> 8]) +
                                                bytes([self.you_ctr & 0xff]) + bytes([self.you_ctr >> 8]) +
                                                b'\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00'+ bytes([self.tail_wan_max]) + b'\x00' + bytes([self.tail_wan_min]) +
                                                b'\x00\x00\x00\x00\x00\x00\x00' +
                                                b'\x00' + bytes([self.cksum]))


                                        for i in range(self.length):

                                            # 保护板百分比
                                            if i == 0:
                                                self.sheet.write(self.e_i, i, self.bms_percent)
                                            # 保护板电压
                                            elif i == 1:
                                                self.sheet.write(self.e_i, i, self.bms_vol)
                                            # 保护板电流
                                            elif i == 2:
                                                self.sheet.write(self.e_i, i, self.bms_cur)
                                            # 保护板功率
                                            elif i == 3:
                                                self.sheet.write(self.e_i, i, self.bms_pw)
                                            # 保护板温度
                                            elif i == 4:
                                                self.sheet.write(self.e_i, i, self.bms_tem)


                                            # 惯导横滚角
                                            elif i == 5:
                                                self.sheet.write(self.e_i, i, self.roll)
                                            # 惯导俯仰角
                                            elif i == 6:
                                                self.sheet.write(self.e_i, i, self.pitch)
                                            # 惯导偏航角
                                            elif i == 7:
                                                self.sheet.write(self.e_i, i, self.yaw)
                                            # 惯导速度
                                            elif i == 8:
                                                self.sheet.write(self.e_i, i, self.v)
                                            #惯导经度
                                            elif i == 9:
                                                self.sheet.write(self.e_i, i, self.gps_jd)
                                            #惯导纬度
                                            elif i == 10:
                                                self.sheet.write(self.e_i, i, self.gps_wd)
                                            #惯导GPS计算路程
                                            elif i == 11:
                                                self.sheet.write(self.e_i, i, self.gps_x)
                                            #惯导GPS计算时间差
                                            elif i == 12:
                                                self.sheet.write(self.e_i, i, self.gps_start_end_time)
                                            #惯导GPS计算速度
                                            elif i == 13:
                                                self.sheet.write(self.e_i,i,self.gps_v)

                                            # 滑台x显示
                                            elif i == 14:
                                                self.sheet.write(self.e_i, i, self.step_x_disp)
                                            # 滑台y显示
                                            elif i == 15:
                                                self.sheet.write(self.e_i, i, self.step_y_disp)
                                            # 深度传感器
                                            elif i == 16:
                                                self.sheet.write(self.e_i, i, self.fish_deep)


                                            # 进水报警
                                            elif i == 17:
                                                if self.data[84] == 1:
                                                    self.sheet.write(self.e_i, i, '头部进水')
                                                else:
                                                    self.sheet.write(self.e_i, i, '头部正常')
                                            elif i == 18:
                                                if self.data[85] == 1:
                                                    self.sheet.write(self.e_i, i, '中仓左部进水')
                                                else:
                                                    self.sheet.write(self.e_i, i, '中仓左部正常')
                                            elif i == 19:
                                                if self.data[86] == 1:
                                                    self.sheet.write(self.e_i, i, '中仓右部进水')
                                                else:
                                                    self.sheet.write(self.e_i, i, '中仓右部正常')


                                            # 驱动控制状态
                                            elif i == 20:
                                                if self.run_state == 0:
                                                    self.sheet.write(self.e_i, i, '停止')
                                                elif self.run_state == 1:
                                                    self.sheet.write(self.e_i, i, '前进')
                                                elif self.run_state == 2:
                                                    self.sheet.write(self.e_i, i, '左转')
                                                elif self.run_state == 3:
                                                    self.sheet.write(self.e_i, i, '右转')
                                                else:
                                                    pass
                                            # 驱动油门
                                            elif i == 21:
                                                self.sheet.write(self.e_i, i, self.you_ctr)
                                            # 驱动转速
                                            elif i == 22:
                                                self.sheet.write(self.e_i, i, self.dr_rpm)
                                            # 驱动控制器温度
                                            elif i == 23:
                                                self.sheet.write(self.e_i, i, self.dr_ctr_tem)
                                            # 驱动电机温度
                                            elif i == 24:
                                                self.sheet.write(self.e_i, i, self.dr_mot_tem)
                                            # 驱动电压
                                            elif i == 25:
                                                self.sheet.write(self.e_i, i, self.dr_vol)
                                            # 驱动电流
                                            elif i == 26:
                                                self.sheet.write(self.e_i, i, self.dr_cur)
                                            # 驱动里程
                                            elif i == 27:
                                                self.sheet.write(self.e_i, i, self.dr_licheng)
                                            # 驱动故障码
                                            elif i == 28:
                                                self.sheet.write(self.e_i, i, self.dr_errcode)
                                            # 驱动鱼尾位置
                                            elif i == 29:
                                                self.sheet.write(self.e_i, i, self.fish_pos)
                                            # 摆尾频率
                                            elif i == 30:
                                                self.sheet.write(self.e_i, i, self.fish_frequency)


                                            # 滑台控制x
                                            elif i == 31:
                                                self.sheet.write(self.e_i, i, self.step_x)
                                            # 滑台控制y
                                            elif i == 32:
                                                self.sheet.write(self.e_i, i, self.step_y)
                                            elif i == 33:
                                                self.sheet.write(self.e_i, i, self.auv_time)
                                            elif i == 34:
                                                self.sheet.write(self.e_i, i, self.tx_mail_num)
                                                self.e_i = self.e_i + 1
                                            else:
                                                pass

                                        #保存数据到EXCEL
                                        self.excel.save(self.excel_name)

                            except:
                                pass

            #判断是否到达设定1分钟时间
            if self.Window.selfmode_flag == 1:
                print('self')
                #self.Window.selfmode_flag = 0
                #自主模式控制字
                self.run_state = 0xAA

                #外部设定的油门
                self.you_ctr = int(self.verticalSlider.value()) * 2
                self.spinBox.setValue(int(self.verticalSlider.value()))

                #校验和
                cksum = ((85 + 170 + 84) + (self.step_x_close & 0xff) + (self.step_x_close >> 8) + (self.run_state) +
                         (self.you_ctr) + (self.step_y_close & 0xff) + (self.step_y_close >> 8)) & 0xff

                # 写数据到串口
                self.ser.write(
                    b'\x55\xAA\x54\x00' +
                    b'\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00' +
                    b'\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00' +
                    bytes([self.step_x_close & 0xff]) + bytes([self.step_x_close >> 8]) +
                    bytes([self.step_y_close & 0xff]) + bytes([self.step_y_close >> 8]) +
                    b'\x00\x00\x00\x00' +
                    bytes([self.run_state & 0xff]) + bytes([self.run_state >> 8]) +
                    bytes([self.you_ctr & 0xff]) + bytes([self.you_ctr >> 8]) +
                    b'\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00\x00' +
                    b'\x00' + bytes([cksum]))

        except:
            pass

if __name__ == '__main__' :
    app = QtWidgets.QApplication( sys.argv )
    myshow = auv2()
    myshow.show()
    sys.exit( app.exec_() )
