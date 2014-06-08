
#ifndef __VB6_H__
#define __VB6_H__ 1
#ifdef  __cplusplus
extern "C" {
#endif

typedef enum{ 
	cb_output=0, 
	cb_dbgout = 1,
    cb_dbgmsg = 2
} cb_type;

typedef void (__stdcall *vbCallback)(cb_type, char*, int);

extern vbCallback vbStdOut;

#ifdef __cplusplus
}
#endif
#endif
