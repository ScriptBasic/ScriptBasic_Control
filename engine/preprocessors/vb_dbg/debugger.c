#pragma warning( disable : 4996)

/*
GNU LGPL
This library is free software; you can redistribute it and/or
modify it under the terms of the GNU Lesser General Public
License as published by the Free Software Foundation; either
version 2.1 of the License, or (at your option) any later version.

This library is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
Lesser General Public License for more details.

You should have received a copy of the GNU Lesser General Public
License along with this library; if not, write to the Free Software
Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA

This program implements a simple debugger "preprocessor" for ScriptBasic.
*/

#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <ctype.h>

#include "conftree.h"
#include "report.h"
#include "reader.h"
#include "basext.h"
#include "prepext.h"

#include "debugger.h"
#include "vb.h"

// Push the item on the debugger stack when entering the function starting at the node Node
static void PushStackItem(pDebuggerObject pDO, long Node)
{
  pDebugCallStack_t p;
  long i;

  p = pDO->pEXT->pST->Alloc(sizeof(DebugCallStack_t),pDO->pEXT->pMemorySegment);
  if( p == NULL )return;
  if( pDO->StackTop == NULL )pDO->StackTop = p;
  p->up = pDO->DbgStack;
  p->down = NULL;
  p->Node = pDO->lPC;
  if( pDO->DbgStack )pDO->DbgStack->down = p;
  pDO->DbgStack = p;
  p->pUF = NULL;
  for( i = 0 ; i < pDO->cUserFunctions ; i++ )
    if( pDO->pUserFunctions[i].NodeId == Node ){
      p->pUF = pDO->pUserFunctions+i;
      break;
      }
  p->LocalVariables = NULL;
  pDO->CallStackDepth++;
  return;
}

/* return from a function and pop off the item from the stack */
static void PopStackItem(pDebuggerObject pDO)
{
  pDebugCallStack_t p;

  if( pDO->DbgStack == NULL || pDO->CallStackDepth == 0 )return;
  p = pDO->DbgStack;
  pDO->DbgStack = pDO->DbgStack->up;
  if( pDO->DbgStack )pDO->DbgStack->down = NULL;
  pDO->pEXT->pST->Free(p,pDO->pEXT->pMemorySegment);
  pDO->CallStackDepth--;
  if( pDO->CallStackDepth == 0 )pDO->StackTop = NULL;
  return;
}

static char hexi(unsigned int x ){
  if( x < 10 )return x+'0';
  return x+'A'-10;
}

/* Print the value of a variable into a string
This function should be used to get the textual representation of a
ScriptBasic T<VARIABLE>.
*/
int SPrintVariable(pDebuggerObject pDO,
                   VARIABLE v,
                   char *pszBuffer,
                   unsigned long *cbBuffer
  ){
/*noverbatim

=itemize
=item T<pDO> is the debugger object
=item T<v> is the variable to print
=item T<pszBuffer> is pointer to the buffer that has to have at least
=item T<cbBuffer> number of bytes available
=noitemize

The function returns zero on success. 

The function returns 1 if the buffer is not large enough. In this case the
number returned in T<*cbBuffer> will be the size of the buffer needed. It may
happen in case the buffer is extremely short that even the returned size is not
enough. Choosing a buffer length of 80 bytes or so ensures that either the
result fits into the buffer or the returned number is large enough to hold
the result.

Note that the number can be extremely large in case the variable is a string. In
this case all the characters are copied into the result and non-printable characters
are converted to hex.

The buffer should be large enough to hold the "->->->->->...->" string representing the
references and the number or string.
CUT*/
  long refcount;
  unsigned char *s,*r;
  char buf[80];
  unsigned long slen,i;
  unsigned long _cbBuffer = *cbBuffer;

  if( v == NULL || TYPE(v) == VTYPE_UNDEF ){
    if( _cbBuffer < 6 )return 1;
    strcpy(pszBuffer,"undef");
    return 0;
    }

#define APPEND(X) slen = strlen(X);\
                  if( _cbBuffer < slen+1 ){\
                    *cbBuffer += 40;\
                    return 1;\
                    }\
                  strcpy(s,X);\
                  s += slen;\
                  _cbBuffer -= slen;

  *pszBuffer = (char)0;
  s = pszBuffer;
  if( TYPE(v) == VTYPE_REF ){
    refcount = 0;
    while( TYPE(v) == VTYPE_REF ){
      v = *(v->Value.aValue);
      if( refcount < 5 ){
        APPEND("->")
        }
      refcount++;
      if( refcount == 1000 ){
        APPEND("... infinit")
        return 0;
        }
      }
    if( refcount > 5 ){
      APPEND(" ... ->")
      }
    }

  if( TYPE(v) == VTYPE_UNDEF ){
    APPEND("undef")
    return 0;
    }

  if( TYPE(v) == VTYPE_LONG ){
    sprintf(buf,"%d",v->Value.lValue);
    slen = strlen(buf);
    if( _cbBuffer < slen+1 ){
      *cbBuffer += slen - _cbBuffer;
      return 1;
      }
    strcpy(s,buf);
    return 0;
    }

  if( TYPE(v) == VTYPE_DOUBLE ){
    sprintf(buf,"%lf",v->Value.dValue);
    slen = strlen(buf);
    if( _cbBuffer < slen+1 ){
      *cbBuffer += slen - _cbBuffer;
      return 1;
      }
    strcpy(s,buf);
    return 0;
    }

  if( TYPE(v) == VTYPE_ARRAY ){
    sprintf(buf,"ARRAY@#%08X",LONGVALUE(v));
    slen = strlen(buf);
    if( _cbBuffer < slen+1 ){
      *cbBuffer += slen - _cbBuffer;
      return 1;
      }
    strcpy(s,buf);
    return 0;
    }

  if( TYPE(v) == VTYPE_STRING ){
    /* calculate the printed size */
    r = v->Value.pValue;
    slen = 2; /* starting and ending " */
    i = 0;
    while( i < STRLEN(v) ){
      if( *r < 0x20 || *r > 0x7F ){
        slen += 4 ; /* \xXX */
        i++;
        r++;
        continue;
        }
      if( *r == '"' ){
        slen += 2 ; /* \" */
        i++;
        r++;
        continue;
        }
      slen ++;
      i++;
      r++;
      continue;
      }

    if( _cbBuffer < slen+1 ){
      *cbBuffer += slen - _cbBuffer;
      return 1;
      }

    r = v->Value.pValue;
    *s ++ = '"';
    i = 0;
    while( i < STRLEN(v) ){
      if( *r < 0x20 || *r > 0x7F ){
        *s ++ = '\\';
        *s ++ = 'x';
        *s ++ = hexi( (*r) / 16 );
        *s ++ = hexi( (*r) & 0xF);
        i++;
        r++;
        continue;
        }
      if( *r == '"' ){
        *s ++ = '\\';
        *s ++ = '"';
        i++;
        r++;
        continue;
        }
      *s ++ = *r;
      i++;
      r++;
      continue;
      }
    *s ++ = '"';
    *s = (char)0;
    return 0;
    }
  return 1;
  }

/*POD
=section SPrintVarByName
=H Print the value of a variable into a string

This fucntion prints a variable string representation into a buffer.
The name of the variable is given in the variable T<pszName>.

The fucntion first searches the variable and then calls the function
R<SPrintVariable> to print the value.

The fucntion first tries to locate the variable as local variable.
For this not the normal debug stack pointer is used, but rather the
T<StackListPointer>. This allows the client to print local
variables levels higher than the bottom of the stack.

If the function succeeds finding the variable it returns the return value of the
function R<SPrintVariable>. If the variable is not found it returns 2.

/*FUNCTION*/
int SPrintVarByName(pDebuggerObject pDO,
                    pExecuteObject pEo,
                    char *pszName,
                    char *pszBuffer,
                    unsigned long *cbBuffer
  ){
/*noverbatim
CUT*/
  pUserFunction_t pUF;
  long i;
  char *s;

  s = pszName;
  while( *s ){
    if( isupper(*s) )*s = tolower(*s);
	if( *s == '\n' || *s == '\r' ){
	  *s = (char)0;
	  break;
	  }
    s++;
    }
  while( isspace(*pszName) )pszName++;

  if( pDO->StackListPointer && pDO->StackListPointer->pUF ){
    pUF = pDO->StackListPointer->pUF;
    for( i=0 ; i < pUF->cLocalVariables ; i++ ){
      if( !strcmp(pUF->ppszLocalVariables[i],pszName) )
        return SPrintVariable(pDO,ARRAYVALUE(pDO->StackListPointer->LocalVariables,i+1),pszBuffer,cbBuffer);
      }
    }
  for( i=0 ; i < pDO->cGlobalVariables ; i++ ){
     if( pDO->ppszGlobalVariables[i] && !strcmp(pDO->ppszGlobalVariables[i],pszName) ){
       if( pEo->GlobalVariables )
         return SPrintVariable(pDO,ARRAYVALUE(pEo->GlobalVariables,i+1),pszBuffer,cbBuffer);
       }
     }

  if( pDO->StackListPointer && pDO->StackListPointer->pUF ){
    pUF = pDO->StackListPointer->pUF;
    for( i=0 ; i < pUF->cLocalVariables ; i++ ){
      if( !strncmp(pUF->ppszLocalVariables[i],"main::",6) && !strcmp(pUF->ppszLocalVariables[i]+6,pszName) )
        return SPrintVariable(pDO,ARRAYVALUE(pDO->StackListPointer->LocalVariables,i+1),pszBuffer,cbBuffer);
      }
    }
  for( i=0 ; i < pDO->cGlobalVariables ; i++ ){
     if( pDO->ppszGlobalVariables[i] && !strncmp(pDO->ppszGlobalVariables[i],"main::",6) && !strcmp(pDO->ppszGlobalVariables[i]+6,pszName) ){
       if( pEo->GlobalVariables )
         return SPrintVariable(pDO,ARRAYVALUE(pEo->GlobalVariables,i+1),pszBuffer,cbBuffer);
       }
     }
  return 2;
  }

/*POD
=section GetSourceLineNumber
=H Get the source line number for a given node

/*FUNCTION*/
long GetSourceLineNumber(pDebuggerObject pDO,
                         long PC
  ){
/*noverbatim
CUT*/
  long i,j;
  long lLineNumber;
  char *pszFileName;

  if( PC < 1 || PC > pDO->cNodes )return 0;

  if( pDO->Nodes[PC-1].lSourceLine )return pDO->Nodes[PC-1].lSourceLine-1;

  /* fill in the whole array */
  for( j=0 ; j < pDO->cNodes ; j++ ){
    lLineNumber = pDO->Nodes[j].lLineNumber;
    pszFileName = pDO->Nodes[j].pszFileName;

    for( i=0 ; i < pDO->cSourceLines ; i ++ )
      if( pDO->SourceLines[i].lLineNumber == lLineNumber && 
          pDO->SourceLines[i].szFileName                 &&
          pszFileName                                    &&
          !strcmp(pDO->SourceLines[i].szFileName,pszFileName) )break;
    pDO->Nodes[j].lSourceLine = i+1;
    }

  return pDO->Nodes[PC-1].lSourceLine-1;
  }

/*FUNCTION*/
long GetCurrentDebugLine(pDebuggerObject pDO
  ){

  if( pDO->StackListPointer == NULL && pDO->StackTop )
     return GetSourceLineNumber(pDO,pDO->StackTop->Node);

  if( pDO->StackListPointer == NULL || pDO->StackListPointer->down == NULL )
    return GetSourceLineNumber(pDO,pDO->pEo->ProgramCounter);

  return GetSourceLineNumber(pDO,pDO->StackListPointer->down->Node);
  }

int MyExecBefore(pExecuteObject pEo){
  long i,j,lThisLine;
  pPrepext pEXT;
  pDebuggerObject pDO;
  char lbuf[80];
  char cmd;

  pEXT = pEo->pHookers->hook_pointer;
  pDO  = pEXT->pPointer;
  pDO->pEo = pEo;

  pDO->lPrevPC = pDO->lPC;
  pDO->lPC = pEo->ProgramCounter;
  if( pDO->DbgStack )pDO->DbgStack->LocalVariables = pEo->LocalVariables;

  lThisLine = GetSourceLineNumber(pDO,pEo->ProgramCounter);

  if( pDO->SourceLines[lThisLine].BreakPoint == 0 ){
    /* if we are executing some step over function */
    if( pDO->Run2CallStack != -1 && pDO->Run2CallStack < pDO->CallStackDepth )return 0;
    if( pDO->Run2Line && pDO->Nodes[pDO->lPC-1].lSourceLine != pDO->Run2Line )return 0;
    }

  scomm_WeAreAt(pDO,lThisLine);

  pDO->StackListPointer = pDO->DbgStack;
  while(1){

    cmd = scomm_GetCommand(pDO,lbuf,80);  // ------------> gets the command from teh user

    switch( cmd ){

      case 'D':/* step the stack list pointer to the bottom */
         pDO->StackListPointer = pDO->DbgStack;
         scomm_Message(pDO,"done");
         continue;
      case 'u':/* step the stack list pointer up */
        if( pDO->StackListPointer ){
          pDO->StackListPointer = pDO->StackListPointer->up;
          scomm_Message(pDO,"done");
          }else scomm_Message(pDO,"No way up more");
        continue;
      case 'd':/* step the stack list pointer down */
        if( pDO->StackListPointer && pDO->StackListPointer->down )
          pDO->StackListPointer = pDO->StackListPointer->down;
        else
          pDO->StackListPointer = pDO->StackTop;
        if( pDO->StackListPointer )
          scomm_Message(pDO,"done");
        else
          scomm_Message(pDO,"No way down more");
        continue;
      case 'b': /* set break point at a line */
        if( ! *lbuf )/* set it at the current line */
          i = GetCurrentDebugLine(pDO)+1;
        else
          GetRange(lbuf,&i,&j);
        if( i < 1 || i > pDO->cSourceLines ){
          scomm_Message(pDO,"invalid line number");
          continue;
          }
        pDO->SourceLines[i-1].BreakPoint = 1;
        scomm_Message(pDO,"done");
        continue;
      case 'B':/* remove breakpoint from line(s) */
        if( ! *lbuf )/* remove all */
          i = 1, j = pDO->cSourceLines;
        else
          GetRange(lbuf,&i,&j);
        if( i < 1 || i >= pDO->cSourceLines ){
          scomm_Message(pDO,"invalid line number");
          continue;
          }
        if( j == 0 )j = i;
        if( j > pDO->cSourceLines )j = pDO->cSourceLines;
        while( i <= j ){
          pDO->SourceLines[i-1].BreakPoint = 0;
          i++;
          }
        scomm_Message(pDO,"done");
        continue;
      case 'q':/* quit the program execution */
        scomm_Message(pDO,"Ok... you have said that... quitting...");
        pEo->pszModuleError = "Debugger Operator Forced Exit.";
        return COMMAND_ERROR_PREPROCESSOR_ABORT;

      case 's':/*step a single line and step into functions */
        pDO->Run2CallStack = pDO->CallStackDepth+1;
        pDO->Run2Line = 0;
        return 0; /* step one step forward */
      case 'o':/* run program until it gets out of the current function */
        pDO->Run2CallStack = pDO->CallStackDepth ? pDO->CallStackDepth - 1 : 0 ;
        pDO->Run2Line = 0;
        return 0; /* step one step forward */
      case 'S':
        pDO->Run2CallStack = pDO->CallStackDepth;
        pDO->Run2Line = 0;
        return 0; /* step one step forward but remain on the same level */
      case 'r':
         pDO->Run2CallStack = -1;/* any level deep */
         if( ! *lbuf ){
           pDO->Run2Line = -1;/* a nonzero value that can not be a valid line number */
           return 0;
           }
         GetRange(lbuf,&i,&j);
         pDO->Run2Line = i;
         return 0;
      case 'R':
         pDO->Run2CallStack = pDO->CallStackDepth; /* on the current level */
         if( ! *lbuf ){
           pDO->Run2Line = -1;/* a nonzero value that can not be a valid line number */
           return 0;
           }
         GetRange(lbuf,&i,&j);
         pDO->Run2Line = i;
         return 0;
      }
    }
  return 0;
  }
int MyExecAfter(pExecuteObject pEo){
  pPrepext pEXT;
  pDebuggerObject pDO;

  pEXT = pEo->pHookers->hook_pointer;
  pDO  = pEXT->pPointer;
  pDO->pEo = pEo;

  return 0;
  }
int MyExecCall(pExecuteObject pEo){
  pPrepext pEXT;
  pDebuggerObject pDO;

  pEXT = pEo->pHookers->hook_pointer;
  pDO  = pEXT->pPointer;
  pDO->pEo = pEo;

  PushStackItem(pDO,pEo->ProgramCounter);
  
  return 0;
  }
int MyExecReturn(pExecuteObject pEo){
  pPrepext pEXT;
  pDebuggerObject pDO;

  pEXT = pEo->pHookers->hook_pointer;
  pDO  = pEXT->pPointer;
  pDO->pEo = pEo;

  PopStackItem(pDO);

  return 0;
  }

static pDebuggerObject new_DebuggerObject(pPrepext pEXT){
  pDebuggerObject pDO;

  pDO = pEXT->pST->Alloc(sizeof(DebuggerObject),pEXT->pMemorySegment);
  if( pDO == NULL )return NULL;

  pDO->pEXT = pEXT;
  pDO->cGlobalVariables = 0;
  pDO->ppszGlobalVariables = NULL;

  pDO->cUserFunctions = 0;
  pDO->pUserFunctions = NULL;

  pDO->cFileNames = 0;
  pDO->ppszFileNames = NULL;

  pDO->cNodes = 0;
  pDO->Nodes = NULL;

  pDO->cSourceLines = 0;
  pDO->SourceLines = NULL;

  pDO->Run2CallStack = 0;
  pDO->Run2Line = 0;

  pDO->pszBindIP = "127.0.0.1";
  pDO->iPort = 6647;
  return pDO;
  }

/*
This function allocates space for the file name in the
preprocessor memory segment.

In case the name was already used then returns the pointer to
the already allocated file name.
*/
static char *AllocFileName(pPrepext pEXT,
                           char *pszFileName
  ){
  long i;
  pDebuggerObject pDO = pEXT->pPointer;
  char **p;

  if( pszFileName == NULL )return NULL;
  for( i=0 ;  i < pDO->cFileNames ; i++ )
    if( !strcmp(pDO->ppszFileNames[i],pszFileName) )return pDO->ppszFileNames[i];
  pDO->cFileNames++;
  p = pEXT->pST->Alloc( sizeof(char *)*pDO->cFileNames,pEXT->pMemorySegment);
  if( p == NULL )return NULL;
  if( pDO->ppszFileNames ){
    memcpy(p,pDO->ppszFileNames,sizeof(char *)*pDO->cFileNames);
    pEXT->pST->Free(pDO->ppszFileNames,pEXT->pMemorySegment);
    }
  pDO->ppszFileNames = p;
  pDO->ppszFileNames[pDO->cFileNames-1] = pEXT->pST->Alloc( strlen(pszFileName)+1,pEXT->pMemorySegment);
  if( pDO->ppszFileNames[pDO->cFileNames-1] == NULL )return NULL;
  strcpy(pDO->ppszFileNames[pDO->cFileNames-1],pszFileName);
  return pDO->ppszFileNames[pDO->cFileNames-1];
  }

static pUserFunction_t AllocUserFunction(pPrepext pEXT,
                                         char *pszUserFunction
  ){
  pDebuggerObject pDO = pEXT->pPointer;
  pUserFunction_t p;

  pDO->cUserFunctions++;
  p = pEXT->pST->Alloc( sizeof(UserFunction_t)*pDO->cUserFunctions,pEXT->pMemorySegment);
  if( p == NULL )return NULL;
  if( pDO->pUserFunctions ){
    memcpy(p,pDO->pUserFunctions,sizeof(UserFunction_t)*pDO->cUserFunctions);
    pEXT->pST->Free(pDO->pUserFunctions,pEXT->pMemorySegment);
    }
  pDO->pUserFunctions = p;
  pDO->pUserFunctions[pDO->cUserFunctions-1].pszFunctionName = pEXT->pST->Alloc( strlen(pszUserFunction)+1,pEXT->pMemorySegment);
  if( pDO->pUserFunctions[pDO->cUserFunctions-1].pszFunctionName == NULL )return NULL;
  strcpy(pDO->pUserFunctions[pDO->cUserFunctions-1].pszFunctionName,pszUserFunction);
  pDO->pUserFunctions[pDO->cUserFunctions-1].ppszLocalVariables = NULL;
  pDO->pUserFunctions[pDO->cUserFunctions-1].cLocalVariables = 0;
  return &(pDO->pUserFunctions[pDO->cUserFunctions-1]);
  }

void CBF_ListLocalVars(char *pszName,
                       void *pSymbol,
                       void **pv){
  pSymbolVAR pVAR = pSymbol;
  pUserFunction_t pUF= pv[0];
  pPrepext pEXT = pv[1];

  pUF->ppszLocalVariables[pVAR->Serial-1] = pEXT->pST->Alloc(strlen(pszName)+1,pEXT->pMemorySegment);
  if( pUF->ppszLocalVariables[pVAR->Serial-1] == NULL )return;
  strcpy(pUF->ppszLocalVariables[pVAR->Serial-1],pszName);
  }

void CBF_ListGlobalVars(char *pszName,
                       void *pSymbol,
                       void *pv){
  pSymbolVAR pVAR = pSymbol;
  pDebuggerObject pDO = pv;

  pDO->ppszGlobalVariables[pVAR->Serial-1] = pDO->pEXT->pST->Alloc(strlen(pszName)+1,pDO->pEXT->pMemorySegment);
  if( pDO->ppszGlobalVariables[pVAR->Serial-1] == NULL )return;
  strcpy(pDO->ppszGlobalVariables[pVAR->Serial-1],pszName);
  }



int vb_dbg_preproc(pPrepext pEXT,long *pCmd, void *p)
{
  char *s;

  switch( *pCmd ){

    case PreprocessorReadDone3:{
      pReadObject pRo = p;
      pDebuggerObject pDO = pEXT->pPointer;
      pSourceLine Result;
      long i;

      Result = pRo->Result;

      /* count the number of source lines */
      i = 0;
      while( Result ){
        i++;
        Result = Result->next;
        }
      pDO->cSourceLines = i;
      pDO->SourceLines = pEXT->pST->Alloc(sizeof(SourceLine_t)*i,pEXT->pMemorySegment);
      *pCmd = PreprocessorUnload;
      if( pDO->SourceLines == NULL )return 1;
      Result = pRo->Result;
      i = 0;
      while( Result ){
        pDO->SourceLines[i].line = pEXT->pST->Alloc(strlen(Result->line)+1,pEXT->pMemorySegment);
        if( pDO->SourceLines[i].line == NULL )return 1;
        strcpy(pDO->SourceLines[i].line,Result->line);
		    s = pDO->SourceLines[i].line;
		    while( *s ){
		      if( *s == '\n' || *s == '\r' )*s = (char)0;
		      s++;
		      }
        /* check for the listen port definition */
        /* REM DBGLISTEN 127.0.0.1:6647         */
        s = pDO->SourceLines[i].line;
        do{
          while( *s && isspace(*s) )s++;/* ignore spaces at the start of the line */
          if( strnicmp(s,"rem",3) )break;
          s += 3;
          if( ! isspace(*s) )break;
          while( isspace(*s) )s++;
          if( strnicmp(s,"dbglisten",9) )break;
          s += 9;
          if( ! isspace(*s) )break;
          while( isspace(*s) )s++;
          
          pDO->pszBindIP = pEXT->pST->Alloc(strlen(s)+1,pEXT->pMemorySegment);
          if( pDO->pszBindIP == NULL )return 1;
          strcpy(pDO->pszBindIP,s);
          s =  pDO->pszBindIP;
          /* seek forward to the : separating the IP address and the port number */
          while( *s && *s != ':' )s++;
          if( *s == ':' ){
            *s = (char)0; /* terminate the ip address*/
            s++;
            pDO->iPort = atoi(s); /* convert the string to int */
            }
        }while(0);
        pDO->SourceLines[i].szFileName = AllocFileName(pEXT,Result->szFileName);
        pDO->SourceLines[i].lLineNumber = Result->lLineNumber;
        pDO->SourceLines[i].BreakPoint = 0;
        i++;
        Result = Result->next;
        }
      *pCmd = PreprocessorContinue;
      return 0;
      }

    case PreprocessorLoad:{
      pDebuggerObject pDO;

      if( pEXT->lVersion != IP_INTERFACE_VERSION ){
        *pCmd = PreprocessorUnload;
        return 0;
        }

      pDO = new_DebuggerObject(pEXT);
      *pCmd = PreprocessorUnload;
      if( pDO == NULL )return 1;

      pEXT->pPointer = pDO;
      *pCmd = PreprocessorContinue;
      return 0;
      }

  case PreprocessorExFinish:{
      peXobject pEx = p;
      pDebuggerObject pDO = pEXT->pPointer;
      peNODE_l Result = pEx->pCommandList;
      long i;

      pDO->cNodes = pEx->NodeCounter;
      pDO->Nodes = pEXT->pST->Alloc(sizeof(DebugNode_t)*pDO->cNodes,pEXT->pMemorySegment);
      if( pDO->Nodes == NULL ){
        *pCmd = PreprocessorUnload;
        return 1;
        }
      for( i=0 ; i < pDO->cNodes ; i++ ){
        pDO->Nodes[i].pszFileName = NULL;
        pDO->Nodes[i].lLineNumber = 0;
        pDO->Nodes[i].lSourceLine = 0;
        }
      while( Result ){
        pDO->Nodes[Result->NodeId-1].pszFileName = AllocFileName(pEXT,Result->szFileName);
        pDO->Nodes[Result->NodeId-1].lLineNumber = Result->lLineNumber;
        Result = Result->rest;
        }
      pDO->cGlobalVariables = pEx->cGlobalVariables;
      pDO->ppszGlobalVariables = pEXT->pST->Alloc( sizeof(char *)*pDO->cGlobalVariables,pEXT->pMemorySegment);
	  pDO->ppszGlobalVariables[0] = NULL;
      if( pDO->ppszGlobalVariables == NULL ){
        *pCmd = PreprocessorUnload;
        return 1;
        }
      pEXT->pST->TraverseSymbolTable(pEx->GlobalVariables,CBF_ListGlobalVars,pDO);
      *pCmd = PreprocessorContinue;
      return 0;
      }
  case PreprocessorExStartLocal:{
      peXobject pEx = p;
      pDebuggerObject pDO = pEXT->pPointer;

      pDO->bLocalStart = 1;
      *pCmd = PreprocessorContinue;
      return 0;
      }
  case PreprocessorExLineNode:{
      peXobject pEx = p;
      pDebuggerObject pDO = pEXT->pPointer;

      if( pDO->bLocalStart ){
        pDO->bLocalStart = 0;
        pDO->FunctionNode = pEx->NodeCounter;
        }
      *pCmd = PreprocessorContinue;
      return 0;
      }
  case PreprocessorExEndLocal:{
      peXobject pEx = p;
      pUserFunction_t pUF;
      pDebuggerObject pDO = pEXT->pPointer;
      void *pv[2];

      *pCmd = PreprocessorContinue;
      if( pEx->ThisFunction == NULL )return 0;/* may happen if syntax error in the BASIC program */
      pUF = AllocUserFunction(pEXT,pEx->ThisFunction->FunctionName);
      pUF->cLocalVariables = pEx->cLocalVariables;
      if( pUF->cLocalVariables )
        pUF->ppszLocalVariables = pEXT->pST->Alloc( sizeof(char *)*pUF->cLocalVariables,pEXT->pMemorySegment);
      else
        pUF->ppszLocalVariables = NULL;
      pUF->NodeId = pDO->FunctionNode;
      *pCmd = PreprocessorUnload;
      if( pUF->cLocalVariables && pUF->ppszLocalVariables == NULL )return 1;
      pv[0] = pUF;
      pv[1] = pEXT;
      pEXT->pST->TraverseSymbolTable(pEx->LocalVariables,(void *)CBF_ListLocalVars,pv);
      *pCmd = PreprocessorContinue;
      return 0;
      }

  case PreprocessorExeStart:
    { pExecuteObject pEo = p;
      pDebuggerObject pDO = pEXT->pPointer;
      pEo->pHookers->hook_pointer = pEXT;
      pDO->CallStackDepth = 0;
      pDO->DbgStack = NULL;
      pDO->StackTop = NULL;
      pEo->pHookers->HOOK_ExecBefore = MyExecBefore;
      pEo->pHookers->HOOK_ExecAfter = MyExecAfter;
      pEo->pHookers->HOOK_ExecCall = MyExecCall;
      pEo->pHookers->HOOK_ExecReturn = MyExecReturn;
      GetSourceLineNumber(pDO,1);/* to calculate all the node numbers for each lines (or the other way around?) */
      scomm_Init(pDO);
      *pCmd = PreprocessorContinue;
      return 0;
      }

    default: /* in any cases that are not handled by the preprocessor just go on */
      *pCmd = PreprocessorContinue;
      return 0;
    }

  }


void scomm_Init(pDebuggerObject pDO)
{
  char cBuffer[500];
  int i;

  sprintf(cBuffer,"DEBUGGER_INIT");
  vbStdOut(cb_debugger, cBuffer, strlen(cBuffer));

  for( i=0 ; i < pDO->cFileNames ; i++ ){
	sprintf(cBuffer,"Source-File: %s\r\n",pDO->ppszFileNames[i]);
    vbStdOut(cb_debugger, cBuffer, strlen(cBuffer));
  }

}

/*POD
=section WeAreAt
=H Send prompt to the debugger station

This function is called by the debugger when it stops before executing
a BASIC line. This function can be used to give some information to the
client, displaying lines around the actual one, values of variables and so on.

/*FUNCTION*/
void scomm_WeAreAt(pDebuggerObject pDO,long i)
{

  char cBuffer[100];
  int cbBuffer;

  sprintf(cBuffer,"Current-Line: %u\r\n",i+1);
  vbStdOut(cb_debugger, cBuffer, strlen(cBuffer));
}

/*POD
=section List
=H List code lines

List the source lines from T<lStart> to T<lEnd>.

The optional T<lThis> may show the caret where the actual execution
context is.

/*FUNCTION*/
void scomm_List(pDebuggerObject pDO, long lStart, long lEnd, long lThis)
{
  long j;
  char cBuffer[1024];
  int cbBuffer;

  if( lStart < 1 )lStart = 1;
  if( lEnd   < 1 )lEnd   = 1;

  for( j = lStart-1 ; j < lEnd ; j++ ){
		
	    if( j >= pDO->cSourceLines )break;

		sprintf(cBuffer,"Break-Point: %s\r\n", pDO->SourceLines[j].BreakPoint ? "1": "0");
		vbStdOut(cb_debugger, cBuffer, strlen(cBuffer));

		sprintf(cBuffer,"Line-Number: %u\r\n",j+1);
		vbStdOut(cb_debugger, cBuffer, strlen(cBuffer));

		sprintf(cBuffer,"Line: %s\r\n",pDO->SourceLines[j].line);
		vbStdOut(cb_debugger, cBuffer, strlen(cBuffer));
  }

}

/*POD
=section GetRange
=H get line number range from a string

This is an auxilliary function, which is used by the debugger.
This simply gets the two numbers from the debugger command and returns
them in the variables pointed by T<plStart> and T<plEnd>.

For example the command T<B 2-5> removes breakpoints from lines 2,3,4 and 5.
In this case this function will return the numbers 2 and 5.

If the first number is missing it is returned as 0. If there is first number
but the last one is missing it is returned 999999999.

If there is first number but it is not followed by '-' then the T<*plEnd> will
be set to zero.

Finally if there are no numbers on the command line then bot variables are set zero.
/*FUNCTION*/

void GetRange(char *pszBuffer, long *plStart, long *plEnd)
{
/*
Arguments:
	pszBuffer the debugger command argument string to get the numbers from
	plStart pointer to the long that will hold the value of the first number
	plEnd pointer to the long that will hold the value of the second number following the dash character
*/
  *plStart = *plEnd = 0;
  while( isspace(*pszBuffer) )pszBuffer++;
  if( !*pszBuffer )return;
  *plStart = atol(pszBuffer);
  while( isdigit(*pszBuffer))pszBuffer++;
  while( isspace(*pszBuffer) )pszBuffer++;
  if( *pszBuffer == '-' ){
    pszBuffer++;
    *plEnd = 999999999;/* something large, very large */
  }
  while( isspace(*pszBuffer) )pszBuffer++;
  if( !*pszBuffer )return;
  *plEnd = atol(pszBuffer);
  return;
  }


/*
static void print_help(){
  printf(
"h help\n"
"s step one line, or just press return on the line\n"
"S step one line, do not step into functions or subs\n"
"o step until getting out of the current function\n"
"  (if you stepped into but changed your mind)\n"
"? var  print the value of a variable\n"
"u step one level up in the stack\n"
"d step one level down in the stack (for variable printing)\n"
"D step down in the stack to current execution depth\n"
"G list all global variables\n"
"L list all local variables\n"
"l [n-m] list the source lines\n"
"r [n] run to line n\n"
"R [n] run to line n but do not stop in recursive function call\n"
"b [n] set breakpoint on the line n or the current line\n"
"B [n-m] remove breakpoints from lines\n"
"q quit the program\n"
);
  }
*/

/*POD
=section Message
=H Report success of some command

This function is called when a command that results no output is executed.
The message is an informal message to the client that either tells that the
command was executed successfully or that the command failed and why.

/*FUNCTION*/
void scomm_Message(pDebuggerObject pDO, char *pszMessage)
{
  char cBuffer[200];
  int cbBuffer;
  sprintf(cBuffer,"Message: %s\r\n",pszMessage);
  vbStdOut(cb_debugger, cBuffer, strlen(cBuffer));
}

/*POD
=section GetCommand
=H Prompt the debugger station

This function should send the prompt to the client and get the client
input. The function should return a single character that represents the
command what the debugger is supposed to do and the possible string argument
in T<pszBuffer>. The available space for the argument is given T<cbBuffer>.

/*FUNCTION*/
int scomm_GetCommand(pDebuggerObject pDO, char *pszBuffer, long dwBuffer)
{
  int i,j;
  int cmd;
  char pszPrintBuff[1024];
  long cbPrintBuff;
  pUserFunction_t pUF;
  pExecuteObject pEo;
  long lStart,lEnd,lThis;
  char cBuffer[1025];
  int cbBuffer;
  pDebugCallStack_t StackListPointer;

  pEo = pDO->pEo;
  while( 1 ){
		lThis = GetCurrentDebugLine(pDO);
		scomm_WeAreAt(pDO,lThis);
	                     
		//cbBuffer = recv(pDO->socket,cBuffer,1024,0); //--------> blocks while waiting to receive a command
		cbBuffer = vbDbgHandler(&cBuffer[0], 1024);

		cmd = *cBuffer;
		while( ('\r' == cBuffer[cbBuffer-1] || '\n' == cBuffer[cbBuffer-1]) && cbBuffer  ){
		  cBuffer[--cbBuffer] = (char)0;
		}

		strcpy(pszBuffer,cBuffer+1);

		switch( cmd ){
		  case 'l':/*list lines*/
					lThis = GetCurrentDebugLine(pDO);

					if( cbBuffer > 2 ){/*if there are arguments: 1 command char, 2 new line */
					  GetRange(cBuffer+1,&lStart,&lEnd);
					  scomm_List(pDO,lStart,lEnd,lThis);
					  }else scomm_WeAreAt(pDO,lThis);

					continue;

		  case '?':
					cbPrintBuff = 1024;
					i = SPrintVarByName(pDO,pDO->pEo,cBuffer+1,pszPrintBuff,&cbPrintBuff);
					switch( i ){
						  case 1:
							scomm_Message(pDO,"variable is too long to print");
							continue;
						  case 2:
							scomm_Message(pDO,"variable is non-existent");
							continue;
						  default:
							sprintf(cBuffer,"Value: %s\r\n",pszPrintBuff);
							vbStdOut(cb_debugger, cBuffer, strlen(cBuffer));
					}
				    continue;

		case 'L': /* list local variables */

				  if( pDO->StackListPointer == NULL || pDO->StackListPointer->pUF == NULL ){
					  /* pUF is NULL when the subroutine is external implemented in a DLL */
					scomm_Message(pDO,"program is not local");
					continue;
					}
				  StackListPointer = pDO->StackListPointer;
				  if( pDO->pEo->ProgramCounter == StackListPointer->pUF->NodeId )
				  {
					/* In this case the debug call stack was already created to handle the function,
					   but the LocalVariables still hold the value of the caller local variables.
					*/
					if( pDO->StackListPointer->up == NULL || pDO->StackListPointer->up->pUF == NULL ){
						  scomm_Message(pDO,"program is not local");
						  continue;
				    }
					StackListPointer = StackListPointer->up;
				  }
				  pUF = StackListPointer->pUF;

				  if( StackListPointer->LocalVariables )
				  for( i=StackListPointer->LocalVariables->ArrayLowLimit ; i <= StackListPointer->LocalVariables->ArrayHighLimit ; i++ )
				  {
						sprintf(cBuffer,"Local-Variable-Name: %s\r\n",pUF->ppszLocalVariables[i-1]);
						vbStdOut(cb_debugger, cBuffer, strlen(cBuffer));

						if( StackListPointer->LocalVariables )
						{
							  j = SPrintVariable(pDO,ARRAYVALUE(pDO->StackListPointer->LocalVariables,i),pszPrintBuff,&cbPrintBuff);
							  switch( j )
							  {
								case 1:
									scomm_Message(pDO,"variable is too long to print");
									continue;
								case 2:
									scomm_Message(pDO,"variable is non-existent");
									continue;
								default:
									sprintf(cBuffer,"Local-Variable-Value: %s\r\n",pszPrintBuff);
									vbStdOut(cb_debugger, cBuffer, strlen(cBuffer));
							  }
						 }
						 else{
						  sprintf(cBuffer,"undef\r\n");
						  vbStdOut(cb_debugger, cBuffer, strlen(cBuffer));
						 }
					}
					continue;

		case 'G':/* list global variables */
				  for( i=0 ; i < pDO->cGlobalVariables ; i++ )
				  {
						if( NULL == pDO->ppszGlobalVariables[i] )continue;

						sprintf(cBuffer,"Global-Variable-Name: %s\r\n",pDO->ppszGlobalVariables[i]);
						vbStdOut(cb_debugger, cBuffer, strlen(cBuffer));

						if( pEo->GlobalVariables )
						{
							  j = SPrintVariable(pDO,ARRAYVALUE(pEo->GlobalVariables,i+1),pszPrintBuff,&cbPrintBuff);
							  switch( j ){
									case 1:
										scomm_Message(pDO,"variable is too long to print");
										continue;
									case 2:
										scomm_Message(pDO,"variable is non-existent");
										continue;
									default:
										sprintf(cBuffer,"Global-Variable-Value: %s\r\n",pszPrintBuff);
										vbStdOut(cb_debugger, cBuffer, strlen(cBuffer));
								}
						  }else{
							sprintf(cBuffer,"undef\r\n");
							vbStdOut(cb_debugger, cBuffer, strlen(cBuffer));
						  }
					}

					continue;
		  }//end switch

		  break;
    }//END WHILE

  return cmd;
}


