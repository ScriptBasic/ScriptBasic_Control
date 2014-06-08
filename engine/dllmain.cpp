
#define _CRT_SECURE_NO_WARNINGS

#include <windows.h>
#include <stdio.h>
#include <stdlib.h>
#include <string.h>

#include "basext.h"
#include "scriba.h" 
#include "errcodes.h"

extern void*  dbg_preproc;  //for console based debugger preprocessor
extern void*  sdbg_preproc; //for the socket based debugger preprocessor
extern LPWSTR __C2W(char *szString);

#define EXPORT comment(linker, "/EXPORT:"__FUNCTION__"="__FUNCDNAME__)

void* vb_stdout = 0;

/*
enum error_codes{
	ec_load_source_failed = 1,
	ec_run_failed = 2

};
*/

void __stdcall SetVBStdout(void* lpfnHandler){
#pragma EXPORT
	vb_stdout = lpfnHandler;
}

int __stdcall GetErrorString(int iErrorCode, char* buf, int bufSz){
#pragma EXPORT
 
  if( iErrorCode >= MAX_ERROR_CODE ) iErrorCode = COMMAND_ERROR_EXTENSION_SPECIFIC;
  int sz = strlen(en_error_messages[iErrorCode]);
  if(sz < bufSz) strcpy(buf,en_error_messages[iErrorCode]);
  return sz;
   
}

int __stdcall run_script(char* fPath, int use_debugger)
{
#pragma EXPORT

  int iError;
  int iErrorCounter;
  unsigned long fErrorFlags;
  char* cmdline = "";
  pSbProgram pProgram;

  pProgram = scriba_new(NULL,NULL);
  scriba_SetFileName(pProgram, fPath);

  if(vb_stdout != NULL) pProgram->fpVbStdOutFunction = vb_stdout;


  if(use_debugger){
		//iError = scriba_LoadInternalPreprocessorByFunction(pProgram, "sdbg", &sdbg_preproc);
  }

  if( iError = scriba_LoadSourceProgram(pProgram) ) goto cleanup;
  if( iError = scriba_Run(pProgram,cmdline) ) goto cleanup;
	  

cleanup:
  scriba_destroy(pProgram);
  return iError;

}
