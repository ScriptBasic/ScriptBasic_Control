
#ifndef __VB6_H__
#define __VB6_H__ 1
#ifdef  __cplusplus
extern "C" {
#endif

typedef enum{ 
	cb_output=0, 
	cb_dbgout = 1,
    cb_debugger = 2,
	cb_engine = 3,
	cb_error = 4
} cb_type;

//Public Sub vb_stdout(ByVal t As cb_type, ByVal lpMsg As Long, ByVal sz As Long)
typedef void (__stdcall *vbCallback)(cb_type, char*, int);

//Public Function GetDebuggerCommand(ByVal buf As Long, ByVal sz As Long) As Long
typedef int (__stdcall *vbDbgCallback)(char*, int);

typedef int (__stdcall *vbHostResolverCallback)(char*, int, int);

extern vbCallback vbStdOut;
extern vbDbgCallback vbDbgHandler;
extern vbHostResolverCallback vbHostResolver;
extern vbDbgCallback vbLineInput; //fpStdinFunction is character based input, we need a full string retrieval for gui..

//we override these to force all outputs to the vb gui. not sure the internal hook thing
//catches all
#ifdef IS_SB_ENGINE
	#define	printf my_printf 
	#define	fprintf my_fprintf

	extern int my_printf(char* format, ...);
	extern int my_fprintf(FILE* fp, char* format, ...);
#endif

#ifdef __cplusplus
}
#endif
#endif
