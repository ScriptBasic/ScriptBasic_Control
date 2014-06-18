
This project aims to create a VB6 usable ScriptBasic Engine.

http://en.wikipedia.org/wiki/ScriptBasic

Features to include:

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
      ? call stack
      ? variable modification
 
Directory structure:
  - top level directory is the vb6 debugger UI using sb_engine.dll
    - requires the /dependancy/scivb_lite.ocx to be registered

  - scripts - test scripts for the sb_engine build