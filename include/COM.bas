
const VbGet = 2
const VbLet = 4
const VbMethod = 1
const VbSet = 8

module com
  declare sub ::CreateObject  alias "CreateObject"  lib "sb_engine.dll"
  declare sub ::CallByName    alias "CallByName"    lib "sb_engine.dll"
  declare sub ::ReleaseObject alias "ReleaseObject" lib "sb_engine.dll"
  declare sub ::GetHostObject alias "GetHostObject" lib "sb_engine.dll"
  declare sub ::GetHostString alias "GetHostString" lib "sb_engine.dll"
  declare sub ::TypeName      alias "TypeName"      lib "sb_engine.dll"
end module

