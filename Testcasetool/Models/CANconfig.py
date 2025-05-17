"""
#!/usr/bin/env python3.9
#-*- coding:utf-8 -*-
@Project:Testcasetool
@File:CANconfig.py
@Author:XU AO
@Time:2025/5/8 09:49
"""
import cantools
from TSMasterAPI import *
from ctypes import *
import tkinter as tk
from tkinter import filedialog
import can
msg = TLIBCAN()
msg.FIdxChn = 0
msg.FIdentifier = 0x100
msg.FProperties = 5  # 表示为扩展帧
msg.FDLC = 8
FData = [0x10, 0x11, 0x12, 0x13, 0x14, 0x15, 0x16, 0x17]
for i in range(len(FData)):
    msg.FData[i] = FData[i]
msg1 = TLIBCAN(FIdentifier = 0x111,FData = [10,11,12,13,14,15,16,17])

FDmsg = TLIBCANFD()
FDmsg.FIdxChn = 0
FDmsg.FIdentifier = 0x101
FDmsg.FProperties = 5
FDmsg.FFDProperties = 0x1
FDmsg.FDLC = 9
FData0 = [0x10, 0x11, 0x12, 0x13, 0x14, 0x15, 0x16, 0x17, 0x18, 0x19, 0x1A, 0x20]
for i in range(len(FData0)):
    FDmsg.FData[i] = FData0[i]


def On_CAN_EVENT(OBJ, ACAN):
    if (ACAN.contents.FIdentifier == 0x111 and ACAN.contents.FIdxChn == 0):
        ACAN.contents.FData[0] +=1


OnCANevent = TCANQueueEvent_Win32(On_CAN_EVENT)
obj = c_int32(0)
id1 = c_int32(0)  # 加载dbc句柄

AppName = b'TSMaster'
def connect():
    # 初始化函数，所需所有函数调用的接口
    # initialize_lib_tsmaster(AppName)
    # 设置can通道数
    if (tsapp_set_can_channel_count(1) == 0):
        print("CAN通道设置成功")
    else:
        print("CAN通道设置失败", tsapp_set_can_channel_count(1))
    # 设置lin通道数
    if (tsapp_set_lin_channel_count(0) == 0):
        print("LIN通道设置成功")
    else:
        print("LIN通道设置失败")
    # 硬件通道映射至软件通道
    # tosun其他硬件只需修改第6个参数，找到对应型号即可
    if 0 == tsapp_set_mapping_verbose(AppName, TLIBApplicationChannelType.APP_CAN, CHANNEL_INDEX.CHN1,
                                      "TC1016".encode("UTF8"), TLIBBusToolDeviceType.TS_USB_DEVICE,
                                      TLIB_TS_Device_Sub_Type.TC1013,0, 0, True):
        print("1通道映射成功")
    else:
        print("1通道映射失败")

    # 设置canfd波特率
    if 0 == tsapp_configure_baudrate_canfd(CHANNEL_INDEX.CHN1, 500.0, 2000.0,
                                           TLIBCANFDControllerType.lfdtISOCAN,
                                           TLIBCANFDControllerMode.lfdmNormal, True):
        print("1通道canfd波特率成功")
    else:
        print("1通道canfd波特率失败")

    if 0 == tsapp_register_pretx_event_can(obj, OnCANevent):
        print("回调事件注册成功")
    else:
        print("回调事件注册失败")
    if 0 == tsapp_connect():  # 0点
        print("can工具连接成功")
        # 硬件开启成功后，开启fifo接收

        r = tsfifo_enable_receive_fifo()
        print("tsfifo_enable_receive_fifo() = ",r)
    else:
        print("can工具连接失败")


# 发送can canfd报文
def SendCANFD_CAN_Message():
    global msg, FDmsg
    can_success, canfd_success = True, True

    for i in range(10):
        print(f"发送第 {i + 1} 次 CAN 报文: {msg}")
        can_result = tsapp_transmit_can_async(msg)
        print(f"发送第 {i + 1} 次 CAN FD 报文: {FDmsg}")
        canfd_result = tsapp_transmit_canfd_async(FDmsg)

        if can_result != 0:
            print(f"第 {i + 1} 次 CAN 报文发送失败，错误码：{can_result}")
            can_success = False

        if canfd_result != 0:
            print(f"第 {i + 1} 次 CAN FD 报文发送失败，错误码：{canfd_result}")
            canfd_success = False

    if can_success and canfd_success:
        ret1 = tsapp_add_cyclic_msg_can(msg, 100)
        ret2 = tsapp_add_cyclic_msg_canfd(FDmsg, 100)

        if ret1 == 0 and ret2 == 0:
            print("CAN 周期发送成功 && CAN FD 周期发送成功")
        else:
            print(f"CAN 周期发送返回值：{ret1}, CAN FD 周期发送返回值：{ret2}")
    else:
        print("异步发送失败，未设置周期发送")



def stop_cyclic_msg_can():
    global msg, FDmsg

    ret1 = tsapp_delete_cyclic_msg_can(msg)
    ret2 = tsapp_delete_cyclic_msg_canfd(FDmsg)

    if ret1 == 0:
        print("CAN 周期发送停止成功")
    else:
        print(f"CAN 周期发送停止失败，错误码：{ret1}")

    if ret2 == 0:
        print("CAN FD 周期发送停止成功")
    else:
        print(f"CAN FD 周期发送停止失败，错误码：{ret2}")



def receive_can_message():
    listcanmsg = (TLIBCAN * 100)()

    listcanfdmsg = (TLIBCANFD * 100)()

    cansize = c_int32(100)

    canfdsize = c_int32(100)
    r = tsfifo_receive_can_msgs(listcanmsg, cansize, 0, READ_TX_RX_DEF.TX_RX_MESSAGES)

    r = tsfifo_receive_canfd_msgs(listcanfdmsg, canfdsize, 0, READ_TX_RX_DEF.TX_RX_MESSAGES)
    print("接收返回值=", r)
    for i in range(cansize.value):
        print("fifo接收canID=", listcanmsg[i].FIdentifier)
    for i in range(canfdsize.value):
        print("fifo接收canfdID=", listcanfdmsg[i].FIdentifier)


def get_enumerate_hw_devices():
    ACount = c_int32(0)
    r = tsapp_enumerate_hw_devices(ACount)
    return r, ACount



def load_dbc():
    global id1
    id1 = c_ulong(0)  # 使用无符号长整型
    root = tk.Tk()
    root.withdraw()
    filepath = filedialog.askopenfilename()

    if not filepath:
        print("未选择文件")
        return 0

    if ".dbc" not in filepath:
        print("文件不正确")
        return 0

    filepath_c = c_char_p(filepath.encode('utf-8'))
    nunmber_c = c_char_p("0,1".encode('utf-8'))

    r = tsdb_load_can_db(filepath_c, nunmber_c, byref(id1))  # 使用 byref

    if r == 0:
        print("id1 =", id1.value)
        print(filepath[filepath.rindex("/") + 1:] + "文件加载成功")
        print_dbc_info(filepath)
        return filepath
    else:
        print("DBC 加载失败，错误码:", r)
        print("错误描述:", tsapp_get_error_description(r).decode('utf-8'))
        return 0

def print_dbc_info(filepath):
    # 读取DBC文件
    db = cantools.database.load_file(filepath)

    # 统计变量
    total_messages = len(db.messages)
    total_signals = 0

    print(f"共加载 {total_messages} 条 CAN 消息：")

    # 打印CAN消息的ID、周期和信号，并统计信号数量
    for i, message in enumerate(db.messages):
        print(f"\nMessage {i + 1}:")
        print("  Message ID: ", hex(message.frame_id))  # 显示为十六进制更直观
        print("  Message Period (ms): ", message.cycle_time)

        signal_count = len(message.signals)
        total_signals += signal_count
        print(f"  共 {signal_count} 个信号：")

        for signal in message.signals:
            print("    Signal Name: ", signal.name)
            print("    Signal Length: ", signal.length)
            print("    Signal Scale: ", signal.scale)

    print(f"\n总计：{total_messages} 条消息，{total_signals} 个信号")


def unload_dbcs():
    if 0 == tsdb_unload_can_dbs():
        print("DBC文件全部卸载")


# 需要绝对路径
fileName = "E:\\sofewareIDE\\python\\py36_32\\tsmaster_test\\1.blf".encode("utf8")


def start_logging():
    tsapp_start_logging(fileName)


def stop_logging():
    tsapp_stop_logging()


# udsHandle = c_byte(0)
# 初始化 udsHandle 为 c_long 类型，并使其为指针类型
udsHandle = c_long(0)  # 使用 c_long 类型而不是 c_byte


def creat_uds_module():
    global udsHandle
    # 确保传递的是指向 c_long 的指针
    r = tsdiag_can_create(POINTER(c_long)(udsHandle), CHANNEL_INDEX.CHN1, 0, 8, 0X1, True, 0X2, True, 0X3, True)
    if r == 0:
        print("udsHandle = ", udsHandle)
    else:
        print(tsapp_get_error_description(r))


def req_and_res_can():
    global udsHandle

    # 创建请求数据数组，包含 100 个 c_uint8 类型元素
    AReqDataArray = (c_uint8 * 100)()  # 创建一个数组

    # 填充请求数据数组
    AReqDataArray[0] = c_uint8(0x22)
    AReqDataArray[1] = c_uint8(0xf1)
    AReqDataArray[2] = c_uint8(0x90)

    # 初始化响应数据大小和数组
    AResSize = c_int32(1000)  # 定义一个整数，初始化为 1000
    AResponseDataArray = (c_uint8 * 1000)()  # 创建响应数据数组

    # 调用发送请求并接收响应的函数
    r = tstp_can_request_and_get_response(udsHandle,
                                          AReqDataArray,  # 请求数据数组
                                          3,  # 请求数据大小（字节数）
                                          AResponseDataArray,  # 响应数据数组
                                          byref(AResSize))  # 通过引用传递 AResSize

    # 检查返回值并处理响应
    if r != 0:
        print(f"请求失败，返回值：{r}")
        return

    print(f"响应数据大小：{AResSize.value}")

    # 打印接收到的响应数据
    for i in range(AResSize.value):
        print(hex(AResponseDataArray[i]), end="  ")
        if i == AResSize.value - 1:
            print()  # 换行



blfID = c_int32(0)
count = c_ulong(0)


def read_blf():
    global blfID, count
    root = tk.Tk()
    root.withdraw()
    filepath = filedialog.askopenfilename()
    if str(filepath).find(".blf"):
        r = tslog_blf_read_start(filepath, blfID, count)
    if r == 0:
        print(filepath[filepath.rindex("/") + 1:] + "文件加载成功")


def read_blf_datas():
    global blfID, count
    realCount = c_ulong(0)
    messageType = TSupportedObjType.sotUnknown
    CANtemp = TLIBCAN()
    CANFDtemp = TLIBCANFD()
    LINtemp = TLIBLIN()
    for i in range(count.value):
        tslog_blf_read_object(blfID, realCount, messageType, CANtemp, LINtemp, CANFDtemp)
        if messageType.value == TSupportedObjType.sotCAN.value:
            print(CANtemp.FTimeUs / 1000000, CANtemp.FIdxChn, CANtemp.FIdentifier, CANtemp.FProperties, CANtemp.FDLC,
                  CANtemp.FData[0], CANtemp.FData[1], CANtemp.FData[2], CANtemp.FData[3], CANtemp.FData[4],
                  CANtemp.FData[5], CANtemp.FData[6], CANtemp.FData[7])
    tslog_blf_read_end(blfID)

_curr_path = os.path.dirname(__file__)
writefileName = (_curr_path+"\\2.blf").encode('utf8')
writeHandle = c_int32(0)


def write_blf_start():
    r = tslog_blf_write_start(writefileName, writeHandle)
    if r == 0:
        global blfID, count
        realCount = c_ulong(0)
        messageType = TSupportedObjType.sotUnknown
        CANtemp = TLIBCAN()
        CANFDtemp = TLIBCANFD()
        LINtemp = TLIBLIN()
        for i in range(count.value):
            tslog_blf_read_object(blfID, realCount, messageType, CANtemp, LINtemp, CANFDtemp)
            if messageType.value == TSupportedObjType.sotCAN.value:
                CANtemp.FIdxChn = 2
                tslog_blf_write_can(writeHandle, CANtemp)
        tslog_blf_read_end(blfID)
        tslog_blf_write_end(writeHandle)
        print("blf_write_successful")
    else:
        print(r)


if __name__ == '__main__':

    initialize_lib_tsmaster("TSMaster".encode("utf8"))
    # tsapp_connect()
    # print(tsfifo_enable_receive_fifo())
    # # print(tsfifo_enable_receive_fifo())
    ret, ACount = get_enumerate_hw_devices()
    print("ret = ",ret)
    print("在线硬件数量有%#d个" % (ACount.value - 1))
    PTLIBHWInfo = TLIBHWInfo()
    for i in range(ACount.value):
        tsapp_get_hw_info_by_index(i, PTLIBHWInfo)
        print(PTLIBHWInfo.FDeviceType, PTLIBHWInfo.FDeviceIndex, PTLIBHWInfo.FVendorName.decode("utf8"),
              PTLIBHWInfo.FDeviceName.decode("utf8"),
              PTLIBHWInfo.FSerialString.decode("utf8"))
    print("0: 连接硬件")
    print("1: 发送报文")
    print("2: 停止周期发送")
    print("3: 接受can_canfd报文")
    print("4: 载入DBC文件")
    print("5: 卸载DBC文件")
    print("6: 开录制报文")
    print("7: 停止制报文")
    print("8: 新建诊断模块")
    print("9: req_res")
    print("q: 退出程序")
    print("a: 读取blf")
    print("b: 获取a blf中的数据")
    print("c: 写blf,在此环境下需先读取blf")
    print("q: 结束程序")
    print("注意后续对硬件操作必须先连接硬件，但如果需要加载dbc文件需先加载dbc再开启硬件")
    while True:

        key = input("请输入")
        if key == '0':  # 连接硬件
            connect()

            # tsapp_connect()

        elif key == '1':  # 先异步单帧发送报文，然后周期发送can canfd报文
            SendCANFD_CAN_Message()
        elif key == '2':  # 停止周期发送报文
            stop_cyclic_msg_can()
        elif key == '3':  # 接受can_canfd报文
            receive_can_message()
        elif key == '4':  # 加载dbc文件
            filename = load_dbc()
        elif key == '5':  # 卸载dbc文件
            unload_dbcs()
        elif key == '6':
            start_logging()
        elif key == '7':  # 停止录制
            stop_logging()
        elif key == '8':  # 诊断相关，创建诊断模块需要在连接函数之前创建模块
            creat_uds_module()
        elif key == '9':  # 请求并获的回复
            req_and_res_can()
        elif key == 'a':
            read_blf()  # 读取blf
        elif key == 'b':  # 获取a blf中的数据
            read_blf_datas()
        elif key == 'c':  # 在此环境下需先读取blf
            write_blf_start()
        elif key == 'q':
            break
    finalize_lib_tsmaster()