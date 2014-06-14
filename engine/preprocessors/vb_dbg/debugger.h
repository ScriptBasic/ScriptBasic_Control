/*
dbg_comm.h
*/
#ifndef __vb_DBG_comm_H__
#define __vb_DBG_comm_H__ 1
#ifdef  __cplusplus
extern "C" {
#endif

#define EXPORT comment(linker, "/EXPORT:"__FUNCTION__"="__FUNCDNAME__)


// Debug information on user defined functions.
typedef struct _UserFunction_t {
  long cLocalVariables;
  char *pszFunctionName;
  char **ppszLocalVariables;
  long NodeId; // node id where the function starts
} UserFunction_t, *pUserFunction_t;


// Debug information for each byte-code node.
typedef struct _DebugNode_t {
  char *pszFileName; // the file name where the source for the node is
  long lLineNumber;  // the line number in the file where the node is 
  long lNodeId;      // the id of the node 
  long lSourceLine;  // the source line number as it is in the memory with included lines counted from 1 
                     // this field is zero and is set when the line is first searched to avoid further searches 
} DebugNode_t, *pDebugNode_t;


// struct for a source line to hold in memory while debugging 
typedef struct _SourceLine_t {
  char *line;
  long lLineNumber;
  char *szFileName;
  int BreakPoint;
  } SourceLine_t, *pSourceLine_t;


// to maintain a call stack to make it available for the user to see local variables and PC and so on
typedef struct _DebugCallStack_t {
  long Node;//where the execution came here to the function (where the function call is)
  pUserFunction_t pUF;
  pFixSizeMemoryObject LocalVariables;
  struct _DebugCallStack_t *up,*down;
} DebugCallStack_t, *pDebugCallStack_t;

typedef struct _DebuggerObject {
  pPrepext pEXT;
  pExecuteObject pEo;
  long cGlobalVariables;
  char **ppszGlobalVariables;
  long cUserFunctions;
  pUserFunction_t pUserFunctions;
  long cFileNames;
  char **ppszFileNames;
  long cNodes;
  pDebugNode_t Nodes;
  long cSourceLines;
  pSourceLine_t SourceLines;
  pDebugCallStack_t DbgStack;
  pDebugCallStack_t StackTop;
  pDebugCallStack_t StackListPointer;
  long CallStackDepth;
  long Run2CallStack;
  long Run2Line;
  int bLocalStart;
  long FunctionNode;
  long lPrevPC,lPC;
} DebuggerObject, *pDebuggerObject;

long __stdcall GetCurrentDebugLine(pDebuggerObject pDO);

int  SPrintVariable(pDebuggerObject pDO,VARIABLE v,char *pszBuffer,unsigned long *cbBuffer);
int  SPrintVarByName(pDebuggerObject pDO,pExecuteObject pEo,char *pszName,char *pszBuffer,unsigned long *cbBuffer);
long GetSourceLineNumber(pDebuggerObject pDO,long PC);

void scomm_Init(pDebuggerObject pDO);
void scomm_WeAreAt(pDebuggerObject pDO, long i);
void scomm_List(pDebuggerObject pDO, long lStart, long lEnd, long lThis);
void scomm_Message(pDebuggerObject pDO, char *pszMessage);

int MyExecBefore(pExecuteObject pEo);

#ifdef __cplusplus
}
#endif
#endif
