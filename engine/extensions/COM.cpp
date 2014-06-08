/* 
 Author:  David Zimmer <dzzie@yahoo.com>
 Site:    http://sandsprite.com

 Notes: Not all COM types are currently handled, but enough to be useful
        this is still a bit of work in progress. I will make additions as
        I use it and find it necessary.

 Script Basic Declarations to use this extension:

		declare sub CreateObject alias "CreateObject" lib "test.exe"
		declare sub CallByName alias "CallByName" lib "test.exe"
		declare sub ReleaseObject alias "ReleaseObject" lib "test.exe"

		const VbGet = 2
		const VbLet = 4
		const VbMethod = 1
		const VbSet = 8

 Example:

		'you can load objects either by ProgID or CLSID
		'obj = CreateObject("SAPI.SpVoice") 
		obj = CreateObject("{96749377-3391-11D2-9EE3-00C04F797396}")

		if obj = 0 then 
			print "CreateObject failed!\n"
		else
			CallByName(obj, "rate", VbLet, 2)
			CallByName(obj, "volume", VbLet, 60)
			CallByName(obj, "speak", VbMethod, "This is my test")
			ReleaseObject(obj)
		end if 

*/

#include <stdio.h>
#include <list>
#include <string>

#include "basext.h"
#include "vb.h"

int com_dbg = 0;
int initilized=0;

//vbCallType aligns with DISPATCH_XX values for Invoke
enum vbCallType{ VbGet = 2, VbLet = 4, VbMethod = 1, VbSet = 8 };
enum colors{ mwhite=15, mgreen=10, mred=12, myellow=14, mblue=9, mpurple=5, mgrey=7, mdkgrey=8 };

//char* to wide string
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

//script basic STRING type to char*
char* GetCString(VARIABLE v){
  
	int slen;
	char *s;
    char* myCopy = NULL;

	s = STRINGVALUE(v);
	slen = STRLEN(v);
	if(slen==0) return strdup("");

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
			 if(vbStdOut){
				 vbStdOut(cb_dbgout, buf, strlen(buf));
			 }else{
				 SetConsoleTextAttribute(hConOut, c);
				 printf("%s",buf);
				 SetConsoleTextAttribute(hConOut,7); 
			 }
		}
		catch(...){}
	}

	SetLastError(dwErr);
}

//note the braces..required so if(x)RETURN0(msg) uses the whole blob 
//should this be goto cleanup instead of return 0? 
#define RETURN0(msg) {if(com_dbg) color_printf(colors::mred, "%s\n", msg); \
	                 LONGVALUE(besRETURNVALUE) = 0; \
					 goto cleanup;}

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
	
	try{
		IDisp->Release();
	}catch(...){
		RETURN0("Invalid IDisp pointer?")
	}

	Argument->Value.lValue = 0;

cleanup:
	return 0;

besEND

//Object CreateObject("ProgID")
besFUNCTION(CreateObject)
  int i;
  int slen;
  char *s;
  char* myCopy = NULL;
  LPWSTR wStr = NULL;
  VARIABLE Argument;
  besRETURNVALUE = besNEWMORTALLONG;
  CLSID     clsid;
  HRESULT	hr;
  IDispatch *IDisp = NULL;

  if(com_dbg) color_printf(colors::myellow, "The number of arguments is: %ld\n",besARGNR);
  
  if( besARGNR != 1) RETURN0("CreateObject takes one argument!") 

  Argument = besARGUMENT(1);
  besDEREFERENCE(Argument);
  
  if( TYPE(Argument) != VTYPE_STRING) RETURN0("CreateObject requires a string argument")

  if(!initilized){
	  CoInitialize(NULL);
	  initilized = 1;
  }

  myCopy = GetCString(Argument);
  if(myCopy==0) RETURN0("malloc failed low mem")

  wStr = __C2W(myCopy);
  if(wStr==0) RETURN0("unicode conversion failed")

  if(com_dbg) color_printf(colors::myellow,"CreateObject(%s)\n", myCopy);

  if(myCopy[0] == '{'){ 
	hr = CLSIDFromString( wStr , &clsid); //its a string CLSID directly
  }else{
	hr = CLSIDFromProgID( wStr , &clsid); //its a progid
  }

  if( hr != S_OK  ) RETURN0("Failed to get clsid")
  
  hr =  CoCreateInstance( clsid, NULL, CLSCTX_INPROC_SERVER, IID_IDispatch,(void**)&IDisp);
  if ( hr != S_OK ) RETURN0("CoCreateInstance failed does object support IDispatch?")

  //todo: keep track of valid objects we create for release/call sanity check latter?
  //	  tracking would break operation though if an embedded host used setvariable to add an obj reference..
  //      unless it used an AddObject(name,pointer) method to add it to the tracker..
  //      how else can we know if a random number is a valid com object other than tracking?
  //      handled with a try/catch block in CallByName right now

cleanup:
	LONGVALUE(besRETURNVALUE) = (int)IDisp;    
	if(myCopy) free(myCopy);
	if(wStr)   free(wStr);
	return 0;

besEND


// the idea behind this one is that we can use a string to embed a type specifier
// to explicitly declare and cast a variable to the type we want such as "VT_I2:2"
//
// in testing with VB6 however, if we pass .vt = VT_I4 when vb6 expects a VT_I1 (char)
// it works as long as the value is < 255, also works with VT_BOOL
//
// do we really need this function ? I prefer less complexity if possible.
//
// Note: there are many COM types, I have no plans to cover them all

bool HandleSpecial(VARIANTARG* va, char* str){

	return false; //disabled for now see notes above..

	if(str==0) return false;

	std::string s = str;
	 
	if(s.length() < 3) return false;
	if(s.substr(0,3) != "VT_") return false;
	
	int pos = s.find(":",0);
	if(pos < 1) return false;

	std::string cmd = s.substr(0,pos);
	if(s.length() < pos+2) return false;

	s = s.substr(pos+1);

	//todo implement handling of these types (there are many more than this)
	if(cmd == "VT_I1"){
	}else if(cmd == "VT_I2"){
	}else if(cmd == "VT_I8"){
    }else if(cmd == "VT_BOOL"){
	}
	
	return true;
}

/*
    arguments in [] are optional, default calltype = method
	callbyname object, "procname", [vbcalltype = VbMethod], [arg0], [arg1] ...
*/	

besFUNCTION(CallByName)

  int i;
  int slen;
  char *s;
  int com_args = 0;
  char* myCopy = NULL;
  LPWSTR wMethodName = NULL;
  vbCallType CallType = VbMethod;
  std::list<BSTR> bstrs;
  VARIANTARG* pvarg = NULL;

  VARIABLE arg_obj;
  VARIABLE arg_procName;
  VARIABLE arg_CallType;

  besRETURNVALUE = besNEWMORTALLONG;
  LONGVALUE(besRETURNVALUE) = 0;

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

  myCopy = GetCString(arg_procName);
  if(myCopy==0) RETURN0("malloc failed low mem")

  wMethodName = __C2W(myCopy);
  if(wMethodName==0) RETURN0("unicode conversion failed")

  if( LONGVALUE(arg_obj) == 0) RETURN0("CallByName(NULL) called")
  IDispatch* IDisp = (IDispatch*)LONGVALUE(arg_obj);
  DISPID  dispid; // long integer containing the dispatch ID
  HRESULT hr;

  // Get the Dispatch ID for the method name, 
  // try block is in case client passed in an invalid pointer
  try{
	  hr = IDisp->GetIDsOfNames(IID_NULL, &wMethodName, 1, LOCALE_USER_DEFAULT, &dispid);
	  if( FAILED(hr) ) RETURN0("GetIDsOfNames failed")
  }
  catch(...){
	  RETURN0("Invalid IDisp pointer?")
  }
	 
  VARIANT    retVal;
  DISPPARAMS dispparams;
  memset(&dispparams, 0, sizeof(dispparams));

  com_args = besARGNR - 3;
  if(com_args < 0) com_args = 0;
   
  if(com_dbg) color_printf(colors::myellow,"CallByName(obj=%x, method='%s', calltype=%d , comArgs=%d)\n", LONGVALUE(arg_obj), myCopy, CallType, com_args);

  // Allocate memory for all VARIANTARG parameters.
  if(com_args > 0){
	 pvarg = new VARIANTARG[com_args];
	 if(pvarg == NULL) RETURN0("failed to alloc VARIANTARGs")
  }

  dispparams.rgvarg = pvarg;
  if(com_args > 0) memset(pvarg, 0, sizeof(VARIANTARG) * com_args);
	 
  dispparams.cArgs = com_args;  // num of args function takes
  dispparams.cNamedArgs = 0;

  /* map in argument values and types    ->[ IN REVERSE ORDER ]<-    */
  for(int i=0; i < com_args; i++){
	  VARIABLE arg_x;		
	  arg_x = besARGUMENT(3 + com_args - i);
	  besDEREFERENCE(arg_x);

		switch( TYPE(arg_x) ){ //script basic type to COM variant type

			  case VTYPE_DOUBLE:
			  case VTYPE_ARRAY:
			  case VTYPE_REF:
				RETURN0("Arguments of script basic types [double, ref, array] not supported")
				break;

			  case VTYPE_LONG:
				pvarg[i].vt = VT_I4;
				pvarg[i].lVal = LONGVALUE(arg_x);
				break;
			  
			  case VTYPE_STRING:
				char* myStr = GetCString(arg_x);
				
				//peek at data and see if an explicit VT_ type was specified.. scriptbasic only supports a few types
				if( !HandleSpecial(&pvarg[i], myStr) ){
					//nope its just a standard string type
					LPWSTR wStr = __C2W(myStr);
					BSTR bstr = SysAllocString(wStr); 
					bstrs.push_back(bstr); //track these to free after call to prevent leak
					pvarg[i].vt = VT_BSTR;
					pvarg[i].bstrVal = bstr;
					free(myStr);
					free(wStr);
				}

				break;			  
				
	  }

  }
   
  //invoke should not need a try catch block because IDisp is already known to be ok and COM should only return a hr result?

  //property put gets special handling..
  if(CallType == VbLet){
	    DISPID mydispid = DISPID_PROPERTYPUT;
        dispparams.rgdispidNamedArgs = &mydispid;
		dispparams.cNamedArgs = 1;
		hr=IDisp->Invoke( dispid, IID_NULL, LOCALE_USER_DEFAULT, CallType, &dispparams, NULL, NULL, NULL); //no return value arg
		if( FAILED(hr) ) RETURN0("Invoke failed")
		goto cleanup;
  }

  hr=IDisp->Invoke( dispid, IID_NULL, LOCALE_USER_DEFAULT, CallType, &dispparams, &retVal, NULL, NULL);
  if( FAILED(hr) ) RETURN0("Invoke failed")

  char* cstr = 0;
  //map in return value to scriptbasic return val
  switch(retVal.vt)
  {
	case VT_EMPTY: break;

	case VT_BSTR:

	    cstr = __B2C(retVal.bstrVal);
		slen = strlen(cstr);
		if(com_dbg) color_printf(colors::myellow,"return value from COM function was string: %s\n", cstr);
		besALLOC_RETURN_STRING(slen);
		memcpy(STRINGVALUE(besRETURNVALUE),cstr,slen);
		free(cstr);
		break;

	case VT_I4:  /* this might be being really lazy but at least with VB6 it works ok.. */
	case VT_I2: 
	case VT_I1: 
    case VT_BOOL:
	case VT_UI1:
	case VT_UI2:
	case VT_UI4:
	case VT_I8:
	case VT_UI8:
	case VT_INT:
	case VT_UINT:

		if(com_dbg) color_printf(colors::myellow,"return value from COM function was numeric: %d\n", retVal.lVal);
        LONGVALUE(besRETURNVALUE) = retVal.lVal;
		break;

	default:
		color_printf(colors::mred,"currently unsupported VT return type: %x\n", retVal.vt);
		break;
  }
 

cleanup:

  for (std::list<BSTR>::iterator it=bstrs.begin(); it != bstrs.end(); ++it) SysFreeString(*it);
  if(pvarg)       delete pvarg;
  if(wMethodName) free(wMethodName); //return0 maybe should goto cleanup cause these would leak 
  if(myCopy)      free(myCopy);
  return 0;

besEND




