/*print.c

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

*/

#include <stdlib.h>
#include <stdio.h>
#include <string.h>
#include "command.h"
#include "vb.h"


COMMAND(MSGBOX)
#if NOTIMP_MSGBOX
NOTIMPLEMENTED;
#else
  NODE nItem;
  VARIABLE ItemResult;
  char *s;
  char* tmp;
  unsigned long slen;
  void (*fpExtOut)(char, void *);
  char buffer[40];

  nItem = PARAMETERNODE;
  //while( nItem ){
		ItemResult = _EVALUATEEXPRESSION_A(CAR(nItem));
		ASSERTOKE;

		if( memory_IsUndef(ItemResult) ){
			strcpy(buffer,"undef");
		}
		else
		{
			switch( TYPE(ItemResult) ){
			  case VTYPE_LONG:
					sprintf(buffer,"%ld",LONGVALUE(ItemResult));
					break;
			  case VTYPE_DOUBLE:
					sprintf(buffer,"%le",DOUBLEVALUE(ItemResult));
					break;
			  case VTYPE_STRING:
					s = STRINGVALUE(ItemResult);
					slen = STRLEN(ItemResult);
				    tmp = (char*)malloc(slen+1);
					memcpy(tmp,s,slen);
					tmp[slen]=0;
					MessageBoxA(0,tmp,"Script Basic",0);
					free(tmp);
					*buffer = (char)0;/* do not print anything afterwards */
					break;
			  case VTYPE_ARRAY:
				sprintf(buffer,"ARRAY@#%08X",LONGVALUE(ItemResult));
				break;
			  }
		}
		
		if(*buffer)	MessageBoxA(0,buffer,"Script Basic",0);
		 

		//nItem = CDR(nItem);
	//}


#endif
END

