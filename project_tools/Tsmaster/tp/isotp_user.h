#ifndef __ISOTP_USER_H__
#define __ISOTP_USER_H__
#include "stdint.h"
#include "isotp_defines.h"
/* user implemented, print debug message */
// extern void isotp_user_debug(const char *message, ...);
#define isotp_user_debug(format, ...) printf(format, ##__VA_ARGS__)
/* 用户实现，发送CAN消息。成功时应返回ISOTP_RET_OK。*/
extern int isotp_user_send_can(const uint32_t arbitration_id, const uint8_t *data, const uint8_t size);


#ifdef __cplusplus
extern "C"
{
#endif

extern uint8_t isotp_init();
extern void UDS_tp_main(void);
extern void time_cont_increase(void);
extern uint32_t isotp_user_get_ms(void);

#ifdef __cplusplus
}
#endif



#endif // __ISOTP_H__
