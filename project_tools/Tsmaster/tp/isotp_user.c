#include "isotp.h"
#include "TSMaster.h"
#include "isotp_user.h"
#define ISOTP_BUFSIZE 1003           // ISOTP缓存大小

uint32_t uds_txid = 0x715; // UDS响应ID，g_link的结构是支持多ID运行的，但是这里不想这么做。
uint32_t uds_rxid = 0x795; // UDS发送ID ，比如 IsoTpLink、txid、rxid配置成结构体，
uint32_t time_cont;
/* Alloc send and receive buffer statically in RAM */
static uint8_t g_isotpRecvBuf[ISOTP_BUFSIZE]; // ISOTP接收缓存数组
static uint8_t g_isotpSendBuf[ISOTP_BUFSIZE]; // ISOTP发送缓存数组
IsoTpLink g_link;

uint8_t payload[ISOTP_BUFSIZE] = {0}; // 接收缓存，用来缓存完整的ISOTP帧，当ISOTP_RET_OK == rets表示连续帧已经接收完成

// @brief ISOtp的底层CAN接收函数回调
// @param args
static void isotp_receive_callback(void)
{
    /* 用于处理多帧传输的轮询链路 */
    isotp_poll(&g_link);
    static uint16_t out_size;
    int rets = isotp_receive(&g_link, payload + 3, ISOTP_BUFSIZE, &out_size); // 处理CAN帧
    if (ISOTP_RET_OK == rets)                                                 // 表示连续帧已经接收完成
    {
        // if (uds_sock > 0)
        // {
        //     uint32_t to_write = out_size;
        //     while (to_write > 0)
        //     {
        //         i//nt written = send(uds_sock, payload + (out_size - to_write)+3, to_write, 0); // 接收完成后，直接通过wifi发送出去
        //         if (written < 0)
        //         {

        //         }
        //         to_write -= written;
        //     }
        // }
   }
}

uint8_t isotp_init()
{
    uint8_t ret = 0;
    isotp_init_link(&g_link, uds_txid,
                    g_isotpSendBuf, sizeof(g_isotpSendBuf),
                    g_isotpRecvBuf, sizeof(g_isotpRecvBuf));

    // ret = esp_timer_create(&periodic_timer_args, &isotp_timer);
    // ret = esp_timer_start_periodic(isotp_timer, 50 * 1000); // 50ms回调一次
    return ret;
}
// @brief 重置缓存和定时器
void isotp_deinit()
{
    //esp_timer_stop(isotp_timer);
    memset(g_isotpSendBuf, 0, ISOTP_BUFSIZE);
    memset(g_isotpRecvBuf, 0, ISOTP_BUFSIZE);
}
/*
 * user implemented, send can message. should return ISOTP_RET_OK when success.
 * ISOTP要求的低层can发送函数接口
 */
int isotp_user_send_can(const uint32_t arbitration_id, const uint8_t *data, const uint8_t size)
{
	TCAN rx_can;
	rx_can.FIdxChn = 0;
	rx_can.FProperties = 0X1;
	rx_can.FDLC = size;
	rx_can.FReserved = 0;
	rx_can.FIdentifier = arbitration_id;
	rx_can.FTimeUs = 0;
	memcpy(&rx_can.FData, data, size);
    if (tsapp_transmit_can_async(&rx_can))
    {
       return ISOTP_RET_OK;
    }
    else
    {
        return ISOTP_RET_TIMEOUT;
    }
}

/*
 * user implemented, get millisecond
 * ISPTP运行的基础计数器，ms级别的计数器
 */
void time_cont_increase(void)
{
    time_cont++;
}

uint32_t isotp_user_get_ms(void)
{
    return time_cont;
}



///////////////////////////////////////////////////////////////////////////////
//说明：网络层的主函数模块，供上层应用层以1ms的周期调度
void UDS_tp_main(void)
{
    isotp_poll(&g_link);
    isotp_receive_callback();    
    time_cont_increase();
}
