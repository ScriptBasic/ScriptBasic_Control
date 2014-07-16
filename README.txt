
This project aims to create a VB6 usable ScriptBasic Engine.
along with a an integrated IDE + debugger.

http://en.wikipedia.org/wiki/ScriptBasic

Features include:

 VB6 access class to ScriptBasic Engine 
   - AddObject
   - AddCode
   ? Eval

 IDE as VB6 ActiveX control
   - intellisense
   - syntax highlighting
   - integrated debugger
      - breakpoints
      - single step
      - step over
      - step out
      - variable inspection 
      - call stack
      - variable modification
      - run to line
 
Status:
   - standalone debugger and vb usable script engine is complete.
      switching over to dll/ocx control will be completed next time I 
      need this functionality embedded in another app. (hard part done)

Notes:

  - auto complete/intellisense has several scopes. hit ctrl+space to trigger.
    if there is a partial identifer already typed, with only one match, the
    string will be auto completed. If there are multiple matches, then the 
    filtered results will be show in intellisense list. If no matches are found
    then entire list will be shown. 

    The following scopes are supported:

      - import statements - lists *.bas in specified /include directory
      - external module functions - parses the *.bas headers to build func list.
      - built in script basic functions 
      - is not currently aware of script variable names
 
   - for module functions (ex curl::) to show up, the matching import must exist
      (include file name, must match embedded module name)

   - debugger variable inspection / modification - When debugging a list view
     of variable names, scopes, and values is kept. You can edit values by right
     clicking its list entry. Array values can be viewed by double clicking on 
     its variable name to bring up the array viewer form. 

     You can also display a variable value, by hovering the mouse over it in
     the IDE window. A call tip will popup showing its value. Click on the call tip
     to being up the edit value form. Longs and string values are supported. You can
     also prefix a string with 0x for hex numbers.

   - parse errors will show up in their own listview. Each error will get its own entry.
     where possible line numbers, files, and error descriptions are provided. Clicking 
     on the entry will jump to that line in the IDE (if one was given by SB engine)

   - changes to scripts are automatically saved each time they are executed.

   - special hot keys:

              ctrl-f - find/replace
              ctrl-g - goto line
              ctrl-z - undo
              ctrl-y - redo

              F2     - set breakpoint
              F5     - go
              F7     - single step
              F9     - step out
              F8     - step over
 