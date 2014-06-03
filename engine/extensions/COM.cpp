
#include <stdio.h>

#include "basext.h"

int com_dbg = 1;
int initilized=0;

//vbCallType aligns with DISPATCH_XX values for Invoke
enum vbCallType{ VbGet = 2, VbLet = 4, VbMethod = 1, VbSet = 8 };
enum colors{ mwhite=15, mgreen=10, mred=12, myellow=14, mblue=9, mpurple=5, mgrey=7, mdkgrey=8 };

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

 // BSTR to C String conversion
char* __B2C(BSTR bString)
{
	int i;
	int n = (int)SysStringLen(bString);
	char *sz;
	sz = (char *)malloc(n + 1);

	for(i = 0; i < n; i++){
		sz[i] = (char)bString[i];
	}
	sz[i] = 0;
	return sz;
}


char* GetCString(VARIABLE v){
  
	int slen;
	char *s;
    char* myCopy = NULL;

	s = STRINGVALUE(v);
	slen = STRLEN(v);
  
	if(slen==0) return 0;

	myCopy = (char*)malloc(slen+1);
	if(myCopy==0) return 0;

	memcpy(myCopy,s, slen);
	myCopy[slen]=0;
	return myCopy;

}


void color_printf(colors c, const char *format, ...)
{
	DWORD dwErr = GetLastError();
	HANDLE hConOut = GetStdHandle( STD_OUTPUT_HANDLE );

	if(format){
		char buf[1024]; 
		va_list args; 
		va_start(args,format); 
		try{
 			 _vsnprintf(buf,1024,format,args);
			 SetConsoleTextAttribute(hConOut, c);
			 printf("%s",buf);
			 SetConsoleTextAttribute(hConOut,7); 
		}
		catch(...){}
	}

	SetLastError(dwErr);
}

//note the braces..required so if(x)RETURN0(msg) uses the whole blob 
#define RETURN0(msg) {if(com_dbg) color_printf(colors::mred, "%s\n", msg); \
	                 LONGVALUE(besRETURNVALUE) = 0; \
					 return 0;}


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

//ReleaseObject(obj)
besFUNCTION(ReleaseObject)

	VARIABLE Argument;
	besRETURNVALUE = besNEWMORTALLONG;

	if( besARGNR != 1) RETURN0("ReleaseObject takes one argument!") 

	Argument = besARGUMENT(1);
	besDEREFERENCE(Argument);

	if( TYPE(Argument) != VTYPE_LONG) RETURN0("ReleaseObject requires a long argument")

	if( LONGVALUE(Argument) == 0) RETURN0("ReleaseObject(NULL) called")

	IDispatch* IDisp = (IDispatch*)LONGVALUE(Argument);
	IDisp->Release();
	Argument->Value.lValue = 0;

	return 0;

besEND

//Object CreateObject("ProgID")
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

  if(com_dbg) color_printf(colors::myellow, "The number of arguments is: %ld\n",besARGNR);
  
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

  if(com_dbg) color_printf(colors::myellow,"CreateObject(%s)\n", myCopy);
  
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
	callbyname object, "procname", [vbcalltype = VbMethod], [args() as variant]
*/	



besFUNCTION(CallByName)

  int i;
  int slen;
  char *s;
  char* myCopy = NULL;
  vbCallType CallType = VbMethod;

  VARIABLE arg_obj;
  VARIABLE arg_procName;
  VARIABLE arg_CallType;

  besRETURNVALUE = besNEWMORTALLONG;

  if(com_dbg) color_printf(colors::myellow,"CallByName %ld args\n",besARGNR);
  
  if(besARGNR < 2) RETURN0("CallByName requires at least 2 args..") 

  arg_obj = besARGUMENT(1);
  besDEREFERENCE(arg_obj);
  
  if( TYPE(arg_obj) != VTYPE_LONG) RETURN0("CallByName first argument must be a long")

  arg_procName = besARGUMENT(2);
  besDEREFERENCE(arg_procName);

  if( TYPE(arg_procName) != VTYPE_STRING) RETURN0("CallByName second argument must be a string")

  if( besARGNR >= 3 ){
    arg_CallType = besARGUMENT(3);
    besDEREFERENCE(arg_CallType);
	CallType = (vbCallType)LONGVALUE(arg_CallType);
  }

  s = STRINGVALUE(arg_procName);
  slen = STRLEN(arg_procName);
  
  if(slen==0) RETURN0("string can not be 0 length") 
   
  myCopy = (char*)malloc(slen+1);
  if(myCopy==0) RETURN0("malloc failed low mem")

  memcpy(myCopy,s, slen);
  myCopy[slen]=0;

  if(com_dbg) color_printf(colors::myellow,"CallByName(%x, %s, %d)\n", LONGVALUE(arg_obj), myCopy, CallType);
  
  LPWSTR wMethodName = __C2W(myCopy);
  free(myCopy);

  if(wMethodName==0) RETURN0("unicode conversion failed")

  //todo: sanity check somehow?	  
  if( LONGVALUE(arg_obj) == 0) RETURN0("CallByName(NULL) called")

  IDispatch* IDisp = (IDispatch*)LONGVALUE(arg_obj);
  DISPID  dispid; // long integer containing the dispatch ID

  // Get the Dispatch ID for the method name
  HRESULT hr=IDisp->GetIDsOfNames(IID_NULL, &wMethodName, 1, LOCALE_USER_DEFAULT, &dispid);
  if( FAILED(hr) ) RETURN0("GetIDsOfNames failed")
	 
  VARIANT    retVal;
  VARIANTARG* pvarg = NULL;
  DISPPARAMS dispparams;

  int com_args = besARGNR - 3;
  if(com_args < 0) com_args = 0;
   
  memset(&dispparams, 0, sizeof(dispparams));

  // Allocate memory for all VARIANTARG parameters.
  if(com_args > 0){
	 pvarg = new VARIANTARG[com_args];
	 if(pvarg == NULL) RETURN0("failed to alloc VARIANTARGs")
  }

  dispparams.rgvarg = pvarg;
  if(com_args > 0) memset(pvarg, 0, sizeof(VARIANTARG) * com_args);
	 
  dispparams.cArgs = com_args;  // num of args function takes
  dispparams.cNamedArgs = 0;

  for(int i=0; i< com_args; i++){
	  /* map in argument values and types */
	  VARIABLE arg_x;		
	  arg_x = besARGUMENT(i + 4);
	  besDEREFERENCE(arg_x);

		switch( TYPE(arg_x) ){

			  case VTYPE_DOUBLE:
			  case VTYPE_ARRAY:
				RETURN0("Arguments of type double and array not currently supported as arguments")
				break;

			  case VTYPE_LONG:
				//printf("This is a long: %ld\n",LONGVALUE(Argument));
				pvarg[i].vt = VT_I4;
				pvarg[i].lVal = LONGVALUE(arg_x);
				break;
			  
			  case VTYPE_STRING:
				char* myStr = GetCString(arg_x);
				LPWSTR wStr = __C2W(myStr);
				BSTR bstr = SysAllocString(wStr); //track these to free after call to prevent leak ?
				free(myStr);
				free(wStr);
				pvarg[i].vt = VT_BSTR;
				pvarg[i].bstrVal = bstr;
				break;			  
				
	  }

  }
   
  // and invoke the method
  if(CallType == VbLet){
	    DISPID mydispid = DISPID_PROPERTYPUT;
        dispparams.rgdispidNamedArgs = &mydispid;
		dispparams.cNamedArgs = 1;
		hr=IDisp->Invoke( dispid, IID_NULL, LOCALE_USER_DEFAULT, CallType, &dispparams, NULL, NULL, NULL); //no return value arg
		if( FAILED(hr) ) RETURN0("Invoke failed")
		return 0;
  }

  hr=IDisp->Invoke( dispid, IID_NULL, LOCALE_USER_DEFAULT, CallType, &dispparams, &retVal, NULL, NULL);
   
  if( FAILED(hr) ) RETURN0("Invoke failed")

  //map in return value to scriptbasic return val
  if(retVal.vt == VT_BSTR){
	    char* cstr = __B2C(retVal.bstrVal);
		slen = strlen(cstr);
		if(com_dbg) color_printf(colors::myellow,"return value from COM function was string: %s\n", cstr);
		besALLOC_RETURN_STRING(slen);
		memcpy(STRINGVALUE(besRETURNVALUE),cstr,slen);
		free(cstr);

  }else if(retVal.vt == VT_I4){
		if(com_dbg) color_printf(colors::myellow,"return value from COM function was: %d\n", retVal.lVal);
        LONGVALUE(besRETURNVALUE) = retVal.lVal;

  }else{
		color_printf(colors::mred,"currently unsupported VT return type: %x\n", retVal.vt);
  }
 
  return 0;

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
