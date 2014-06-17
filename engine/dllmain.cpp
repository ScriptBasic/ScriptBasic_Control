
#include <windows.h>
#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <stdarg.h>

#include "basext.h"
#include "scriba.h" 
#include "errcodes.h"

//this has to be done before including vb.h where fprint is redefined..
typedef int (*realfprintf)(FILE*, const char*, ...);
realfprintf real_fprintf = fprintf;

typedef int (*realprintf)(const char*, ...);
realprintf real_printf = printf;

#define IS_SB_ENGINE 1
#include "vb.h"



extern "C" void*  vb_dbg_preproc; 

#define EXPORT comment(linker, "/EXPORT:"__FUNCTION__"="__FUNCDNAME__)
extern LPWSTR __C2W(char *szString);

void __stdcall SetDefaultDirs(char* incDir, char* modDir){
#pragma EXPORT
	if(pszDefaultIncludeDir) free(pszDefaultIncludeDir);
	if(pszDefaultModuleDir) free(pszDefaultModuleDir);
	pszDefaultIncludeDir = strdup(incDir);
	pszDefaultModuleDir  = strdup(modDir);
}

void __stdcall SetCallBacks(void* lpfnMsgHandler, void* lpfnDbgHandler, void* lpfnHostResolver){
#pragma EXPORT
	vbStdOut     = (vbCallback)lpfnMsgHandler;
	vbDbgHandler = (vbDbgCallback)lpfnDbgHandler;
	vbHostResolver = (vbHostResolverCallback)lpfnHostResolver;
}

int __stdcall GetErrorString(unsigned int iErrorCode, char* buf, int bufSz){
#pragma EXPORT
  int sz = 0;
  if( iErrorCode >= MAX_ERROR_CODE ) iErrorCode = COMMAND_ERROR_EXTENSION_SPECIFIC;
  sz = strlen(en_error_messages[iErrorCode]);
  if(sz < bufSz) strcpy(buf,en_error_messages[iErrorCode]);
  return sz;
   
}

int __stdcall run_script(char* fPath, int use_debugger)
{
#pragma EXPORT

  int iError=0;
  char* cmdline = "";
  pSbProgram pProgram;
  char* buf[255]; 

  pProgram = scriba_new(NULL,NULL);
  scriba_SetFileName(pProgram, fPath);

  if(use_debugger){
		iError = scriba_LoadInternalPreprocessorByFunction(pProgram, "vb_dbg", &vb_dbg_preproc);
		if(iError != 0) goto cleanup;
  }
  
  //build will call report error if syntax fail, these goto vbStdOut if set..
  if(scriba_LoadSourceProgram(pProgram) ) goto cleanup; 

  if(vbStdOut){
	  sprintf((char*)buf, "ENGINE_PRECONFIG:%d", pProgram); 
	  vbStdOut(cb_engine, (char*)buf, strlen((char*)buf));
  }

  if( iError = scriba_Run(pProgram,cmdline) ){
	  if( iError > 0 ) report_report(stderr,"",0,iError,REPORT_ERROR,NULL,NULL,NULL);	  
	  goto cleanup;
	  
  }

cleanup:
  scriba_destroy(pProgram);

  if(vbStdOut){
	  sprintf((char*)buf, "ENGINE_DESTROY:%d", pProgram); 
	  vbStdOut(cb_engine, (char*)buf, strlen((char*)buf));
  }

  return 0;

}

int my_printf(char* format, ...){
	
	char *ret = 0;
	char buf[1024]; //avoid malloc/free for short strings if we can
	int alloced = 0;

	if(!format) return 0;

	va_list args; 
	va_start(args,format); 
	int size = _vscprintf(format, args); 
	if(size==0){va_end(args); return 0;}

	if(size < 1020){
		ret = &buf[0];
	}else{
		alloced = true;
		size++; //for null
		ret = (char*)malloc(size+2);
		if(ret==0){va_end(args); return 0;}
	}

	_vsnprintf(ret, size, format, args);
	ret[size]=0; //explicitly null terminate
	va_end(args);

	//here is where you could forward the char* to a UI handler..
	if(vbStdOut != NULL) vbStdOut(cb_output, ret, strlen(ret));
	else real_printf("%s",ret);
	
	if(alloced) free(ret);
	return 0;
}

int my_fprintf(FILE* fp, char* format, ...){
	
	char *ret = 0;
	char buf[1024]; //avoid malloc/free for short strings if we can
	int alloced = 0;

	if(!format) return 0;

	va_list args; 
	va_start(args,format); 
	int size = _vscprintf(format, args); 
	if(size==0){va_end(args); return 0;}

	if(size < 1020){
		ret = &buf[0];
	}else{
		alloced = true;
		size++; //for null
		ret = (char*)malloc(size+2);
		if(ret==0){va_end(args); return 0;}
	}

	_vsnprintf(ret, size, format, args);
	ret[size]=0; //explicitly null terminate
	va_end(args);

	if(vbStdOut != NULL && (fp == stdout || fp == stderr) ){
		vbStdOut( (fp == stdout ? cb_output : cb_error) , ret, strlen(ret));
	}
	else real_fprintf(fp, "%s", ret);
	
	if(alloced) free(ret);
	return 0;
}