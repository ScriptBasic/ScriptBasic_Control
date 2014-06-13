
#ifndef __VB6_H__
#define __VB6_H__ 1
#ifdef  __cplusplus
extern "C" {
#endif

typedef enum{ 
	cb_output=0, 
	cb_dbgout = 1,
    cb_debugger = 2,
	cb_engine = 3
} cb_type;

//Public Sub vb_stdout(ByVal t As cb_type, ByVal lpMsg As Long, ByVal sz As Long)
typedef void (__stdcall *vbCallback)(cb_type, char*, int);

//Public Function GetDebuggerCommand(ByVal buf As Long, ByVal sz As Long) As Long
typedef int (__stdcall *vbDbgCallback)(char*, int);

extern vbCallback vbStdOut;
extern vbDbgCallback vbDbgHandler;

#ifdef __cplusplus
}
#endif
#endif
