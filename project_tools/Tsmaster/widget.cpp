#include "widget.h"
#include "ui_widget.h"
#include "TSMaster.h"
#include <QtDebug>
#include <string.h>

#include "tp\isotp_user.h"

Widget::Widget(QWidget *parent) :
    QWidget(parent),
    ui(new Ui::Widget)
{
    ui->setupUi(this);

    /*下面这些自己定义窗口，在对应放*/
    initialize_lib_tsmaster("Tsmaster_qt");/*必要，初始化下*/
    tsapp_set_vendor_detect_preferences(0,0,1,0,0,0);/*当你使用非同星或者CANOE就需要掉用这个函数，选取哪个can卡就哪个置1*/

    if(1)
    {
        /*这个函数运行就可以直接打开同星硬件的窗口。注意需要更新到最新软件*/
        tsapp_show_tsmaster_window("Hardware",1);
        if (0 != tsapp_connect()){ /* handle error */ };
    }
    else
    {
        /*下面函数为自己配置，参数自己看TLIBTSMapping结构体内容*/
        tsapp_set_can_channel_count(1);
        TLIBTSMapping m;
        // TSMaster CAN 通道 1 - PEAK 1 CAN 通道 82
        sprintf_s(m.FAppName, "%s", "TSMaster");
        sprintf_s(m.FHWDeviceName, "%s", "PEAK");
        m.FAppChannelIndex = 0;
        m.FAppChannelType = (TLIBApplicationChannelType)0;
        m.FHWDeviceType = (TLIBBusToolDeviceType)4;
        m.FHWDeviceSubType = -1;
        m.FHWIndex = 0;
        m.FHWChannelIndex = 81;
        if (0 != tsapp_set_mapping(&m)) { /* handle error */ };

        if (0 != tsapp_connect()){ /* handle error */ };

        /* do your work here */

    }
    init_for_can_iso_tp();
    mainTimer = new QTimer;
    countTimer = new QTimer;
    QObject::connect(mainTimer,SIGNAL(timeout()),this,SLOT(mainTaskSlot()));
    mainTimer->start(200);

    /*计数使用*/
    QObject::connect(countTimer,SIGNAL(timeout()),this,SLOT(time_cnt()));
    countTimer->start(50);

}

Widget::~Widget()
{
    /*函数停止运行时调用这个*/
    tsapp_disconnect();
    finalize_lib_tsmaster();
    delete ui;
}

void Widget::mainTaskSlot()
{
    /*初始化TP，目前TP相关没弄好，单纯只是放在此处，有兴趣自己完善*/
    time_cont_increase();
    UDS_tp_main();

    /*发送数据*/
    TCAN f0 = {0,0x1,8,0,0x7B,0,{0x00, 0x08, 0x00, 0x00, 0x23, 0x00, 0x00, 0x00}};
    TCAN f1 = {0,0x1,8,0,0x7C,0,{0x00, 0x08, 0x00, 0x00, 0x23, 0x00, 0x00, 0x00}};
    TCAN f2 = {0,0x1,8,0,0x7D,0,{0x00, 0x08, 0x00, 0x00, 0x23, 0x00, 0x00, 0x00}};
    tsapp_transmit_can_async(&f0);
    tsapp_transmit_can_async(&f1);
    tsapp_transmit_can_async(&f2);

}



void Widget::time_cnt()
{
   isotp_user_get_ms();
}




/*初始化TP*/
void Widget::init_for_can_iso_tp(void)
{
 isotp_init();
}
