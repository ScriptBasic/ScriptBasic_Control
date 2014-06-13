
#include <windows.h>
#include <stdio.h>
#include <stdlib.h>
#include <string.h>

#include "basext.h"
#include "scriba.h" 
#include "errcodes.h"
#include "vb.h"

//include dir used in reader.c
// for( cft_GetEx(pRo->pConfig,"include",&Node,&s,NULL,NULL,NULL);  

//module dir used in modumana and ipreproc 
//(493  )       if( ! cft_GetEx(pEo->pConfig,"module",&Node,&s,NULL,NULL,NULL) ){


extern void*  vb_dbg_preproc; 

#define EXPORT comment(linker, "/EXPORT:"__FUNCTION__"="__FUNCDNAME__)
extern LPWSTR __C2W(char *szString);

void __stdcall SetDefaultDirs(char* incDir, char* modDir){
#pragma EXPORT
	if(pszDefaultIncludeDir) free(pszDefaultIncludeDir);
	if(pszDefaultModuleDir) free(pszDefaultModuleDir);
	pszDefaultIncludeDir = strdup(incDir);
	pszDefaultModuleDir  = strdup(modDir);
}

void __stdcall SetCallBacks(void* lpfnMsgHandler, void* lpfnDbgHandler){
#pragma EXPORT
	vbStdOut     = (vbCallback)lpfnMsgHandler;
	vbDbgHandler = (vbDbgCallback)lpfnDbgHandler;
}

int __stdcall GetErrorString(int iErrorCode, char* buf, int bufSz){
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

  int iError;
  char* cmdline = "";
  pSbProgram pProgram;
  char* buf[255]; 

  pProgram = scriba_new(NULL,NULL);
  scriba_SetFileName(pProgram, fPath);

  if(use_debugger){
		iError = scriba_LoadInternalPreprocessorByFunction(pProgram, "vb_dbg", &vb_dbg_preproc);
		if(iError != 0) goto cleanup;
  }

  if( iError = scriba_LoadSourceProgram(pProgram) ) goto cleanup;

  if(vbStdOut){
	  sprintf(buf, "ENGINE_PRECONFIG:%d", pProgram); 
	  vbStdOut(cb_engine, buf, strlen(buf));
  }

  if( iError = scriba_Run(pProgram,cmdline) ) goto cleanup;
	  

cleanup:
  scriba_destroy(pProgram);

  if(vbStdOut){
	  sprintf(buf, "ENGINE_DESTROY:%d", pProgram); 
	  vbStdOut(cb_engine, buf, strlen(buf));
  }

  return iError;

}
