/*
dbg_comm.h
*/
#ifndef __vb_DBG_comm_H__
#define __vb_DBG_comm_H__ 1
#ifdef  __cplusplus
extern "C" {
#endif
/*FUNDEF*/

void scomm_Init(pDebuggerObject pDO);
/*FEDNUF*/
/*FUNDEF*/

void scomm_WeAreAt(pDebuggerObject pDO,
                  long i);
/*FEDNUF*/
/*FUNDEF*/

void scomm_List(pDebuggerObject pDO,
               long lStart,
               long lEnd,
               long lThis);
/*FEDNUF*/
/*FUNDEF*/

void GetRange(char *pszBuffer,
              long *plStart,
              long *plEnd);
/*FEDNUF*/
/*FUNDEF*/

void scomm_Message(pDebuggerObject pDO,
                  char *pszMessage);
/*FEDNUF*/
/*FUNDEF*/

int scomm_GetCommand(pDebuggerObject pDO,
                    char *pszBuffer,
                    long cbBuffer);
/*FEDNUF*/
#ifdef __cplusplus
}
#endif
#endif
