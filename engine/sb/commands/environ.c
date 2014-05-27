/* environ.c

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

#ifdef __DARWIN__
	#define _environ (*_NSGetEnviron())
#endif

extern char **_environ;

/**ENVIRON
=section misc
=title ENVIRON("envsymbol") or ENVIRON(n)
=display ENVIRON()

This function returns the value of an environment variable.
Environment variables are string values associated to names that are provided by
the executing environment for the programs. The executing environment is usually
the operating system, but it can also be the Web server in CGI programs that
alters the environment variables provided by the surrounding operating system
specifying extra values.

This function can be used to get the string of an environment variable in case
the program knows the name of the variable or to list all the environment
variables one by one.

If the environment variable name is known then the name as a string has to be
passed to this function as argument. In this case the return value is the value
of the environment variable. For example

=verbatim
  MyPath = ENVIRON("PATH")
=noverbatim

If the program wants to list all the environment variables the argument to the
function T<ENVIRON> should be an integer number T<n>. In this case the function
returns a string containing the name and the value of the
T<n>-th environment variable joined by a T<=> sign.
The numbering starts with T<n=0>.

If the argument value is integer and is out of the range of the possible
environment variable ordinal numbers (negative or larger or equal to the number
of the available environment variables) then the function returns T<undef>.

If the argument to the function is T<undef> then the function also returns the
T<undef> value.

Note that ScriptBasic provides an easy way for the embedding applications to
redefine the underlying function that returns the environment variable. Thus
an embedding application can "fool" a BASIC program providing its own
environment variable. For example the Eszter SB Application Engine provides an
alternative environment variable reading function and thus BASIC applications
can read the environment using the function T<ENVIRON> as if the program was
running in a real CGI environment.

=details

The following sample code prints all environment variables:

=verbatim
i=0
do
  e$ = environ(i)
  if IsDefined(e$) then
    print e$
    print
  endif
  i = i + 1
loop while IsDefined(e$)
=noverbatim


*/
COMMAND(ENVIRON)
#if NOTIMP_ENVIRON
NOTIMPLEMENTED;
#else

  NODE nItem;
  VARIABLE Op1;
#define MAX_ENVIRV_LEN 256
  char *s,buffer[MAX_ENVIRV_LEN];
  size_t sLen;
  char **EnvironmentPointer;
  long index;
  char *(*EnvirFunc)(void *, char *, long );
  /* this is an operator and not a command, therefore we do not have our own mortal list */
  USE_CALLER_MORTALS;

  /* evaluate the parameter */
  nItem = PARAMETERLIST;
  Op1 = _EVALUATEEXPRESSION(CAR(nItem));
  ASSERTOKE;

  if( memory_IsUndef(Op1) ){
    RESULT = NULL;
    RETURN;
    }

  EnvirFunc = pEo->fpEnvirFunction;
  if( TYPE(Op1) == VTYPE_LONG ){
    index = LONGVALUE(Op1);
    if( EnvirFunc ){
      s = EnvirFunc(pEo->pEmbedder,NULL,index);
      }else{
      EnvironmentPointer = _environ;
      if( index < 0 ){
        RESULT = NULL;
        RETURN;
        }
      /* we have to walk one by one, because index may be too large and _environ is terminated with NULL and
         we do not have the length */
      while( index && *EnvironmentPointer ){
        index--;
        EnvironmentPointer++;
        }
      s = *EnvironmentPointer;
      }
    }else{

    Op1 = CONVERT2STRING(Op1);
    if( (sLen=STRLEN(Op1)) > MAX_ENVIRV_LEN -1 )sLen = MAX_ENVIRV_LEN-1;
    memcpy(buffer,STRINGVALUE(Op1),sLen);
    buffer[sLen] = (char)0;
    if( EnvirFunc )
      s = EnvirFunc(pEo->pEmbedder,buffer,0);
    else
      s = getenv( buffer );
    }

  if( s == NULL ){
    RESULT = NULL;
    RETURN;
    }
  sLen=strlen(s);
  RESULT = NEWMORTALSTRING(sLen);
  ASSERTNULL(RESULT)
  memcpy(STRINGVALUE(RESULT),s,sLen);

#endif
END

/**COMMANDF
=section misc
=title COMMAND()
=display COMMAND()

This function returns the command line arguments of the program in a single
string. This does not include the name of the interpreter and the name of the
BASIC program, only the arguments that are to be passed to the BASIC program.
For example the program started as

=verbatim
# scriba test_command.sb arg1 arg2  arg3
=noverbatim

will see T<"arg1 arg2 arg3"> string as the return value of the function
T<COMMAND()>.

=details

You can start the ScriptBasic interpreter several ways. The most evident way is
to specify the interpreter on the command line:

=verbatim
# scriba test_command.sb arg1 arg2  arg3
=noverbatim

You can also start the program using the name of the program:

=verbatim
# test_command.sb arg1 arg2  arg3
=noverbatim

This assumes that the script is executable and that the first line is
T<#!/usr/bin/scriba> pointing to the interpreter (UNIX). On Windows the file
extension T<.sb> should be associated with the interpreter. This is done
by the ScriptBasic installer. (Also see the note below.)

Under Windows you can also start the code without specifying the extension:

=verbatim
# test_command arg1 arg2  arg3
=noverbatim

You can have the program compiled to executable and you can start the exe with
the command

=verbatim
# test_command arg1 arg2  arg3
=noverbatim

where T<test_command> now refers to the executable file. In any of these cases
the function T<COMMAND> returns the string containning the commal line
arguments.

NOTE:

Some versions of the installer has a bug that will prevent passing command-line
arguments to a BASIC program when it is started on the command by its name. To
check whether your have ScriptBasic corretcly installed try execute the
following program:

=verbatim
REM test_command.sb
print COMMAND()
=noverbatim
Save the program into the file T<test_command.sb> and type the command:

=verbatim
C:\> test_command arg1 arg2  arg3
=noverbatim

If the program prints the arguments, then you are not exposed to this bug. If
the program returns without printing anything but a new line you have installed
a ScriptBasic version that contains this installer bug.

To correct this bug you should start the registry editor and open the key

=verbatim
HKEY_CLASSES_ROOT\ScriptBasic\shell\command
=noverbatim

The default value should be

=verbatim
C:\ScriptBasic\BIN\scriba.exe %1
=noverbatim

The actual path to the executable may differ from the listed above. If that is
the case you can modify the value to be

=verbatim
C:\ScriptBasic\BIN\scriba.exe %1 %*
=noverbatim

You should NOT use the value from this document as listed above, but rather you
have to edit the value and append T< %*> after the original value, because
it is likely that you have a differet installation path of the interpreter.

*/
COMMAND(COMMANDF)
#if NOTIMP_COMMANDF
NOTIMPLEMENTED;
#else

  int q;

  /* this is an operator and not a command, therefore we do not have our own mortal list */
  USE_CALLER_MORTALS;

  if( pEo->CmdLineArgument == NULL ){
    RESULT = NULL;
    RETURN;
    }

  RESULT = NEWMORTALSTRING(q=strlen(pEo->CmdLineArgument));
  ASSERTNULL(RESULT)
  memcpy(STRINGVALUE(RESULT),pEo->CmdLineArgument,q);

#endif
END
