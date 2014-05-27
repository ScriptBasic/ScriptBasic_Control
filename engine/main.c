
#define _CRT_SECURE_NO_WARNINGS

#include <stdio.h>
#include <stdlib.h>
#include <string.h>

#include "basext.h"
#include "scriba.h" 

void main(int argc, char *argv[]){
  
  int iError;
  int iErrorCounter;
  unsigned long fErrorFlags;

  char *ppszPreprocessorName[] = {"sdbg", NULL};
  char* CmdLinBuffer = "";

  pSbProgram pProgram;
  pProgram = scriba_new(malloc,free);
  scriba_LoadConfiguration(pProgram,"C:\\scriptbasic\\bin\\scriba.conf");
  scriba_SetFileName(pProgram,"C:\\scriptbasic\\examples\\ntexamples\\hello.sb");

  /*
  if( scriba_UseCacheFile(pProgram) == SCRIBA_ERROR_SUCCESS ){
    if( (iError = scriba_LoadBinaryProgram(pProgram)) != 0 ){
		  printf("failed to load binary program!");
		  exit(0);
     }
//	 iError=scriba_RunExternalPreprocessor(pProgram,pszEPreproc);
	printf("running from cached compile\n"); 
	scriba_Run(pProgram, CmdLinBuffer);
	scriba_destroy(pProgram);
    exit(0);
  }*/
 
  iError = scriba_LoadInternalPreprocessor(pProgram, &ppszPreprocessorName);

  if( scriba_LoadSourceProgram(pProgram) ){
	  printf("failed to load source program!");
	  exit(0);
  }

  //scriba_SaveCacheFile(pProgram);
   
  if( iError=scriba_Run(pProgram,CmdLinBuffer) ){
	report_report(stderr,"",0,iError,REPORT_ERROR,&iErrorCounter,NULL,&fErrorFlags);
    exit(0);
  }
 
  scriba_destroy(pProgram);
	
  printf("Press any key to continue...\n");
  getchar();

}
