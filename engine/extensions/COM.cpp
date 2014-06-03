
#include <stdio.h>

//extern "C" {
	#include "basext.h"
//}

int com_dbg = 1;
int initilized=0;

LPWSTR __C2W(char *szString){
	DWORD n;
	char *sz = NULL;
	LPWSTR ws= NULL;
	if(*szString && szString){
		sz = strdup(szString);
		n = MultiByteToWideChar(CP_ACP, MB_PRECOMPOSED, sz, -1, NULL, 0);
		if(n){
			ws = (LPWSTR)malloc(n*2);
			MultiByteToWideChar(CP_ACP, MB_PRECOMPOSED, sz, -1, ws, n);
		}
	}
	free(sz);
	return ws;
}

/* 
   besVERSION_NEGOTIATE, besSUB_START, besSUB_FINISH are optional
   
   since they are called via exports..they will cause a problem if we include
   multiple extensions compiled directly into the embedded intrepreter instance (same name in each)
   commands themselves will each have a unique name so they should be safe..and because they are compiled
   in, version_negotiate should not be needed..lack of a sub_start/finish of module load/unload may need
   to be comprensated for depending on extension. (check load flag per module and init if 0 etc)

   compiling extensions directly into the intrepreter was not an original design consideration and is more of
   a hack.. -dz


besVERSION_NEGOTIATE

  printf("The function bootmodu was started and the requested version is %d\n",Version);
  printf("The variation is: %s\n",pszVariation);
  printf("We are returning accepted version %d\n",(int)INTERFACE_VERSION);
  return (int)INTERFACE_VERSION;

besEND

besSUB_START
  long *pL;

  besMODULEPOINTER = besALLOC(sizeof(long));
  if( besMODULEPOINTER == NULL )return 0;
  pL = (long *)besMODULEPOINTER;
  *pL = 0L;

  printf("The function bootmodu was started.\n");

besEND

besSUB_FINISH
  printf("The function finimodu was started.\n");
besEND
*/

//note the braces..required so if(x)RETURN0(msg) uses the whole blob 
#define RETURN0(msg) {if(com_dbg) printf("%s\n",msg); \
	                 LONGVALUE(besRETURNVALUE) = 0; \
					 return 0;}

besFUNCTION(CreateObject)
  int i;
  int slen;
  char *s;
  char* myCopy = NULL;
  VARIABLE Argument;
  besRETURNVALUE = besNEWMORTALLONG;
  CLSID     clsid;
  HRESULT	hr;
  IDispatch *IDisp;

  if(com_dbg) printf("The number of arguments is: %ld\n",besARGNR);
  
  if( besARGNR != 1) RETURN0("CreateObject takes one argument!") 

  Argument = besARGUMENT(1);
  besDEREFERENCE(Argument);
  
  if( TYPE(Argument) != VTYPE_STRING) RETURN0("CreateObject requires a string argument")

  s = STRINGVALUE(Argument);
  slen = STRLEN(Argument);
  
  if(slen==0) RETURN0("string can not be 0 length") 
   
  if(!initilized){
	  hr = CoInitialize(NULL);
	  if( hr != S_OK  ) RETURN0("CoInitialize failed")
	  initilized = 1;
  }

  myCopy = (char*)malloc(slen+1);
  if(myCopy==0) RETURN0("malloc failed low mem")

  memcpy(myCopy,s, slen);
  myCopy[slen]=0;

  if(com_dbg) printf("CreateObject(%s)\n", myCopy);
  
  LPWSTR wStr = __C2W(myCopy);
  free(myCopy);

  if(wStr==0) RETURN0("unicode conversion failed")

  hr = CLSIDFromProgID( wStr , &clsid);
  if( hr != S_OK  ) RETURN0("malloc failed low mem")
  free(wStr);

  hr =  CoCreateInstance( clsid, NULL, CLSCTX_INPROC_SERVER, IID_IDispatch,(void**)&IDisp);

  if ( hr != S_OK ) RETURN0("CoCreateInstance failed does object support IDispatch?")

  //todo: keep track of valid objects we create for release/sanity check latter?
  LONGVALUE(besRETURNVALUE) = (int)IDisp;    

besEND

/*
besFUNCTION(set1)
  VARIABLE Argument;
  LEFTVALUE Lval;
  int i;
  unsigned long __refcount_;

  for( i=1 ; i <= besARGNR ; i++ ){
    Argument = besARGUMENT(i);

    besLEFTVALUE(Argument,Lval);
    if( Lval ){
      besRELEASE(*Lval);
      *Lval = besNEWLONG;
      if( *Lval )
        LONGVALUE(*Lval) = 1;
      }
    }

besEND

besFUNCTION(arbdata)
  VARIABLE Argument;
  LEFTVALUE Lval;
  static char buffer[1024];
  char *p;
  unsigned long __refcount_;

  p = buffer;
  sprintf(buffer,"%s","hohohoho\n");
  Argument = besARGUMENT(1);

  besLEFTVALUE(Argument,Lval);
  if( Lval ){
    besRELEASE(*Lval);
    *Lval = besNEWSTRING(sizeof(char*));
    memcpy(STRINGVALUE(*Lval),&p,sizeof(p));
    }

besEND

besFUNCTION(pzchar)
  int i;
  VARIABLE Argument;
  char *p;

  for( i=1 ; i <= besARGNR ; i++ ){
    Argument = besARGUMENT(i);
    besDEREFERENCE(Argument);
    memcpy(&p,STRINGVALUE(Argument),sizeof(p));
    printf("%s\n",p);
    }
besEND


besFUNCTION(trial)

  printf("Function trial was started...\n");

  besRETURNVALUE = besNEWMORTALLONG;
  LONGVALUE(besRETURNVALUE) = g_modVal++;

/*printf("Module directory is %s\n",besCONFIG("module"));
  printf("dll extension is %s\n",besCONFIG("dll"));
  printf("include directory is %s\n",besCONFIG("include"));
* /

besEND

besFUNCTION(myicall)
  VARIABLE Argument;
  VARIABLE pArgument;
  VARIABLE FunctionResult;
  unsigned long ulEntryPoint;
  unsigned long i;

  Argument = besARGUMENT(1);
  besDEREFERENCE(Argument);
  ulEntryPoint = LONGVALUE(Argument);

  pArgument = besNEWARRAY(0,besARGNR-2);
  for( i=2 ; i <= (unsigned)besARGNR ; i++ ){
     pArgument->Value.aValue[i-2] = besARGUMENT(i);
     }

  besHOOK_CALLSCRIBAFUNCTION(ulEntryPoint,
                             pArgument->Value.aValue,
                             besARGNR-1,
                             &FunctionResult);

  for( i=2 ; i <= (unsigned)besARGNR ; i++ ){
     pArgument->Value.aValue[i-2] = NULL;
     }
  besRELEASE(pArgument);
  besRELEASE(FunctionResult);
besEND

besSUB_AUTO
  printf("autoloading %s\n",pszFunction);
  *ppFunction = (void *)trial;
besEND

besCOMMAND(iff)
  NODE nItem;
  VARIABLE Op1;
  long ConditionValue;

  /* this is an operator and not a command, therefore we do not have our own mortal list * /
  USE_CALLER_MORTALS;

  /* evaluate the parameter * /
  nItem = besPARAMETERLIST;
  if( ! nItem ){
    RESULT = NULL;
    RETURN;
    }
  Op1 = besEVALUATEEXPRESSION(CAR(nItem));
  ASSERTOKE;

  if( Op1 == NULL )ConditionValue = 0;
  else{
    Op1 = besCONVERT2LONG(Op1);
    ConditionValue = LONGVALUE(Op1);
    }

  if( ! ConditionValue )
    nItem = CDR(nItem);

  if( ! nItem ){
    RESULT = NULL;
    RETURN;
    }
  nItem = CDR(nItem);

  RESULT = besEVALUATEEXPRESSION(CAR(nItem));
  ASSERTOKE;
  
  RETURN;
besEND_COMMAND
*/
