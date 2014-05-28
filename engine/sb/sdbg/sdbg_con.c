/*
FILE: sdbg_con.c
HEADER: dbg_comm.h

--GNU LGPL
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

#include "sdbg.h"

#define SEND2CLIENT do{cbBuffer = strlen(cBuffer); send(pDO->socket,cBuffer,cbBuffer,0);}while(0)

/*POD
=H Debugger communication module

This file implements the functions that are used by the debugger module and,
which communicate with the debugger station. This sample implementation is
the possible simplest example implementation using T<getchar()> to get characters
and T<printf()> to send characters to the user.

Other implementations should implement the same functions but using more sophisticated
methods, like connecting to a socket where a graphical debugger client application is
accepting connection and wants to communincate with the debugger module.

CUT*/

/*POD
=section Init
=H Initiate communication with the debugger station

This function is called by the debugger when the execution of the program starts.
This function has to set up the debugger environment with the client. Connecting to
the listening socket, clearing screen and so on.

/*FUNCTION*/
void comm_Init(pDebuggerObject pDO
  ){
/*noverbatim
CUT*/
  long i;
  char cBuffer[100];
  int cbBuffer;
  struct sockaddr_in SA;
#ifdef WIN32
  WSADATA wsaData;
  WORD wVersionRequested;
  wVersionRequested = MAKEWORD( 2, 2 );
  i = WSAStartup( wVersionRequested, &wsaData );
#endif

  pDO->listen_socket = socket(AF_INET,SOCK_STREAM, 0);
  SA.sin_family = AF_INET;
  SA.sin_addr.s_addr = inet_addr(pDO->pszBindIP);
  SA.sin_port = htons(pDO->iPort);
  i = bind(pDO->listen_socket,(struct sockaddr *)&SA,sizeof(SA));
  if( i ) i = WSAGetLastError();
  listen(pDO->listen_socket,1);
  pDO->socket = accept(pDO->listen_socket,NULL,NULL);
  sprintf(cBuffer,"Application: sbdbg 1.0\r\nVersion: 1.0\r\n");
  SEND2CLIENT;
  sprintf(cBuffer,"Source-File-Count: %u\r\n",pDO->cFileNames);
  SEND2CLIENT;
  for( i=0 ; i < pDO->cFileNames ; i++ ){
	sprintf(cBuffer,"Source-File: %s\r\n",pDO->ppszFileNames[i]);
    SEND2CLIENT;
    }
  }

/*POD
=section WeAreAt
=H Send prompt to the debugger station

This function is called by the debugger when it stops before executing
a BASIC line. This function can be used to give some information to the
client, displaying lines around the actual one, values of variables and so on.

/*FUNCTION*/
void comm_WeAreAt(pDebuggerObject pDO,
                  long i
  ){
/*noverbatim
CUT*/
  char cBuffer[100];
  int cbBuffer;

  sprintf(cBuffer,"Current-Line: %u\r\n",i+1);
  SEND2CLIENT;
  }

/*POD
=section List
=H List code lines

List the source lines from T<lStart> to T<lEnd>.

The optional T<lThis> may show the caret where the actual execution
context is.

/*FUNCTION*/
void comm_List(pDebuggerObject pDO,
               long lStart,
               long lEnd,
               long lThis
  ){
/*noverbatim
CUT*/
  long j;
  char cBuffer[1024];
  int cbBuffer;

  if( lStart < 1 )lStart = 1;
  if( lEnd   < 1 )lEnd   = 1;

  for( j = lStart-1 ; j < lEnd ; j++ ){
    if( j >= pDO->cSourceLines )break;
	sprintf(cBuffer,"Break-Point: %s\r\n", pDO->SourceLines[j].BreakPoint ? "1": "0");
    SEND2CLIENT;
	sprintf(cBuffer,"Line-Number: %u\r\n",j+1);
    SEND2CLIENT;
	sprintf(cBuffer,"Line: %s\r\n",pDO->SourceLines[j].line);
    SEND2CLIENT;
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
void GetRange(char *pszBuffer,
              long *plStart,
              long *plEnd
  ){
/*noverbatim
Arguments:
=itemize
=item T<pszBuffer> the debugger command argument string to get the numbers from
=item T<plStart> pointer to the long that will hold the value of the first number
=item T<plEnd> pointer to the long that will hold the value of the second number following the dash character
=noitemize
CUT*/
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
void comm_Message(pDebuggerObject pDO,
                  char *pszMessage
  ){
/*noverbatim
CUT*/
  char cBuffer[200];
  int cbBuffer;
  sprintf(cBuffer,"Message: %s\r\n",pszMessage);
  SEND2CLIENT;
  }

/*POD
=section GetCommand
=H Prompt the debugger station

This function should send the prompt to the client and get the client
input. The function should return a single character that represents the
command what the debugger is supposed to do and the possible string argument
in T<pszBuffer>. The available space for the argument is given T<cbBuffer>.

/*FUNCTION*/
int comm_GetCommand(pDebuggerObject pDO,
                    char *pszBuffer,
                    long dwBuffer
  ){
/*noverbatim
The commands that the debugger accepts: (see help function printout above).

The function may also implement some printing commands itself, like printing
a help screen.
CUT*/
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
    comm_WeAreAt(pDO,lThis);
    send(pDO->socket,".\r\n",3,0);
    cbBuffer = recv(pDO->socket,cBuffer,1024,0);
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
          comm_List(pDO,lStart,lEnd,lThis);
          }else comm_WeAreAt(pDO,lThis);

        continue;
      case '?':
        cbPrintBuff = 1024;
        i = SPrintVarByName(pDO,pDO->pEo,cBuffer+1,pszPrintBuff,&cbPrintBuff);
        switch( i ){
          case 1:
            comm_Message(pDO,"variable is too long to print");
            continue;
          case 2:
            comm_Message(pDO,"variable is non-existent");
            continue;
          default:
			sprintf(cBuffer,"Value: %s\r\n",pszPrintBuff);
            SEND2CLIENT;
          }
      continue;
    case 'L': /* list local variables */
      if( pDO->StackListPointer == NULL || pDO->StackListPointer->pUF == NULL ){
		  /* pUF is NULL when the subroutine is external implemented in a DLL */
        comm_Message(pDO,"program is not local");
        continue;
        }
	  StackListPointer = pDO->StackListPointer;
	  if( pDO->pEo->ProgramCounter == StackListPointer->pUF->NodeId ){
		/* In this case the debug call stack was already created to handle the function,
		   but the LocalVariables still hold the value of the caller local variables.
		*/
		if( pDO->StackListPointer->up == NULL || pDO->StackListPointer->up->pUF == NULL ){
          comm_Message(pDO,"program is not local");
          continue;
		  }
		StackListPointer = StackListPointer->up;
	    }
      pUF = StackListPointer->pUF;

	  if( StackListPointer->LocalVariables )
      for( i=StackListPointer->LocalVariables->ArrayLowLimit ;
		   i <= StackListPointer->LocalVariables->ArrayHighLimit ; i++ ){
		sprintf(cBuffer,"Local-Variable-Name: %s\r\n",pUF->ppszLocalVariables[i-1]);
        SEND2CLIENT;
        if( StackListPointer->LocalVariables ){
          j = SPrintVariable(pDO,ARRAYVALUE(pDO->StackListPointer->LocalVariables,i),pszPrintBuff,&cbPrintBuff);
          switch( j ){
			case 1:
				comm_Message(pDO,"variable is too long to print");
				continue;
			case 2:
				comm_Message(pDO,"variable is non-existent");
				continue;
			default:
				sprintf(cBuffer,"Local-Variable-Value: %s\r\n",pszPrintBuff);
				SEND2CLIENT;
            }
          }else{
          sprintf(cBuffer,"undef\r\n");
		  SEND2CLIENT;
          }
        }
      continue;
    case 'G':/* list global variables */
      for( i=0 ; i < pDO->cGlobalVariables ; i++ ){
        if( NULL == pDO->ppszGlobalVariables[i] )continue;
		sprintf(cBuffer,"Global-Variable-Name: %s\r\n",pDO->ppszGlobalVariables[i]);
		SEND2CLIENT;
        if( pEo->GlobalVariables ){
          j = SPrintVariable(pDO,ARRAYVALUE(pEo->GlobalVariables,i+1),pszPrintBuff,&cbPrintBuff);
          switch( j ){
			case 1:
				comm_Message(pDO,"variable is too long to print");
				continue;
			case 2:
				comm_Message(pDO,"variable is non-existent");
				continue;
			default:
				sprintf(cBuffer,"Global-Variable-Value: %s\r\n",pszPrintBuff);
				SEND2CLIENT;
            }
		  }else{
            sprintf(cBuffer,"undef\r\n");
		    SEND2CLIENT;
		    }
        }
      continue;
      }

    break;
    }
  return cmd;
  }
