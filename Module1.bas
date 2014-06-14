Attribute VB_Name = "Module1"
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, source As Any, ByVal Length As Long)
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'get the currently executing line number
Public Declare Function GetCurrentDebugLine Lib "sb_engine" (ByVal hDebug As Long) As Long

'get a variables value from its name
Private Declare Function dbg_getVarVal Lib "sb_engine" (ByVal hDebug As Long, ByRef varName As Byte, ByRef lpBuf As Byte, ByRef bufSz As Long) As Long

'enumerate global and local variable names (uses vbstdOut callback)
Private Declare Sub dbg_EnumVars Lib "sb_engine" (ByVal hDebug As Long)

Private Declare Sub dbg_EnumAryVarsByName Lib "sb_engine" (ByVal hDebug As Long, ByRef varName As Byte)
Private Declare Sub dbg_EnumAryVarsByPointer Lib "sb_engine" (ByVal hDebug As Long, ByVal pVar As Long)

Private Declare Function dbg_VarTypeFromName Lib "sb_engine" (ByVal hDebug As Long, ByRef varName As Byte) As Long

Public Declare Function dbg_LineCount Lib "sb_engine" (ByVal hDebug As Long) As Long

'we may not need this we should track internally...
Private Declare Function dbg_isBpSet Lib "sb_engine" (ByVal hDebug As Long, ByVal lineNo As Long) As Long

Private Declare Sub dbg_EnumCallStack Lib "sb_engine" (ByVal hDebug As Long)
Private Declare Sub dbg_RunToLine Lib "sb_engine" (ByVal hDebug As Long, ByVal lineNo As Long)

Private Declare Sub SetDefaultDirs Lib "sb_engine" (ByRef incDir As Byte, ByRef modDir As Byte)

'scripts that use import directive internally turn into one long flat script file. for debugging
'we must dump them to file and then load them for display so line numbers line up..
Private Declare Sub dbg_WriteFlatSourceFile Lib "sb_engine" (ByVal hDebug As Long, ByRef fPath As Byte)
Private Declare Function dbg_SourceLineCount Lib "sb_engine" (ByVal hDebug As Long) As Long
'

Public hProgram As Long
Public hDebugObject As Long     'handle to the current debug object - pDO
Public readyToReturn As Boolean
Public dbg_cmd As String
Public running As Boolean
Public variables As New Collection 'of CVariable
Public callStack As New Collection 'of CCallStack
Public fso As New CFileSystem2

Enum cb_type
    cb_output = 0
    cb_dbgout = 1
    cb_debugger = 2
    cb_engine = 3
End Enum

Enum sb_VarTypes
    VTYPE_UNKNOWN = -1
    VTYPE_LONG = 0
    VTYPE_DOUBLE = 1
    VTYPE_STRING = 2
    VTYPE_ARRAY = 3
    VTYPE_REF = 4
    VTYPE_UNDEF = 5
End Enum

Function LoadFlatFile() As Boolean
    
    Dim tmp As String
    Dim b() As Byte
    
    tmp = fso.GetFreeFileName(Environ("temp"))
    b() = StrConv(tmp & Chr(0), vbFromUnicode)
    dbg_WriteFlatSourceFile hDebugObject, b(0)
    
    If fso.FileExists(tmp) Then
        Form1.hasImports = True
        Form1.scivb.ReadOnly = False
        LoadFlatFile = Form1.scivb.LoadFile(tmp)
        Form1.scivb.ReadOnly = True
    End If
    
End Function

Public Sub RunToLine(lineNo As Long)
    dbg_RunToLine hDebugObject, lineNo
    DebuggerCmd "m"
End Sub

Public Function EnumCallStack() As Collection
    
    Set callStack = Nothing
    dbg_EnumCallStack hDebugObject 'this goes into syncronous set of callbacks
    Set EnumCallStack = callStack
    
End Function

Public Function VariableTypeToString(x As sb_VarTypes) As String

    types = Array("LONG", "DOUBLE", "STRING", "ARRAY", "REF", "UNDEF")
    
    If x < 0 Or x > 5 Then
        VariableTypeToString = "???"
    Else
        VariableTypeToString = LCase(types(x))
    End If
    
End Function

Public Function VariableType(varName As String) As String
    
    Dim x  As sb_VarTypes
    Dim v() As Byte
    
    v() = StrConv(varName & Chr(0), vbFromUnicode)
    x = dbg_VarTypeFromName(hDebugObject, v(0))
    VariableType = VariableTypeToString(x)
    
End Function

Public Sub DebuggerCmd(cmd As String)
    dbg_cmd = cmd
    readyToReturn = True
End Sub

Public Sub SetConfig(includeDir As String, moduleDir As String)
    
    Dim i() As Byte, m() As Byte
    
    i() = StrConv(includeDir & Chr(0), vbFromUnicode)
    m() = StrConv(moduleDir & Chr(0), vbFromUnicode)
     
    SetDefaultDirs i(0), m(0)
    
End Sub

Public Function EnumArrayVariables(varNameOrPointer As Variant) As Collection
    
    Dim v() As Byte
    Set variables = Nothing
    
    If TypeName(varNameOrPointer) = "String" Then
        v() = StrConv(varNameOrPointer & Chr(0), vbFromUnicode)
        dbg_EnumAryVarsByName hDebugObject, v(0) 'this goes into syncronous set of callbacks
    Else
        dbg_EnumAryVarsByPointer hDebugObject, CLng(varNameOrPointer) 'this goes into syncronous set of callbacks
    End If
    
    Set EnumArrayVariables = variables
    
End Function

Public Function EnumVariables() As Collection
   
    Set variables = Nothing
    dbg_EnumVars hDebugObject 'this goes into syncronous set of callbacks
    Set EnumVariables = variables
    
End Function

Public Function GetVariableValue(varName As String) As String
    
    Dim v() As Byte
    Dim buf() As Byte
    Dim sz As Long
    Dim ret As Long
    Dim i As Long
    
    sz = 1024
    v() = StrConv(varName & Chr(0), vbFromUnicode)
    ReDim buf(sz)
     
    ret = dbg_getVarVal(hDebugObject, v(0), buf(0), sz)
    
    If ret = 0 Then
        GetVariableValue = StrConv(buf, vbUnicode)
        i = InStr(GetVariableValue, Chr(0))
        If i > 1 Then
            GetVariableValue = Left(GetVariableValue, i - 1)
        End If
    ElseIf ret = 1 Then
        GetVariableValue = "[ > 1024 chars ]"
    ElseIf ret = 2 Then
        GetVariableValue = "[Variable not found]"
    Else
        GetVariableValue = "[Unknown return value: " & ret & "]"
    End If
        
    
End Function


Public Function GetDebuggerCommand(ByVal buf As Long, ByVal sz As Long) As Long
        
    Dim b() As Byte
    Dim source As String, curline As Long
    
    dbg_cmd = Empty
    
    'there are some lines we dont want to stop and show as execution to the user,
    'such as declares and function starts
    curline = GetCurrentDebugLine(hDebugObject)
    If Not BreakPointExists(curline) Then
        source = LCase(Form1.scivb.GetLineText(curline))
        source = Trim(Replace(source, vbTab, Empty))
        If InStr(source, " ") > 1 Then
            source = Left(source, InStr(source, " ") - 1)
            If source = "declare" Or source = "function" Then
                dbg_cmd = "s"
            End If
        End If
    End If
    
    If dbg_cmd = Empty Then
        Form1.SyncUI
        
        'we block here until the UI sets the readyToReturn = true
        'this is not a CPU hog, and form remains responsive to user actions..
        readyToReturn = False
        While Not readyToReturn
            DoEvents
            Sleep 20
        Wend
    End If
    
    If Len(dbg_cmd) < sz Then
        dbg_cmd = dbg_cmd & Chr(0)
        ReDim b(Len(dbg_cmd))
        b() = StrConv(dbg_cmd, vbFromUnicode)
        CopyMemory ByVal buf, b(0), Len(dbg_cmd)
        GetDebuggerCommand = Len(dbg_cmd)
    Else
        GetDebuggerCommand = 0
        MsgBox "Shouldnt happen!"
    End If
    
End Function

Public Sub HandleDebugMessage(msg As String)

    Dim cmd() As String
    Dim v As CVariable
    Dim c As New cCallStack
    Dim handled As Boolean
    
    If Left(msg, 10) = "Call-Stack" Then
        cmd = Split(msg, ":", 3)
    ElseIf Left(msg, 14) = "Array-Variable" Then
        cmd = Split(msg, ":", 4)
    Else
        cmd = Split(msg, ":", 2)
    End If
    
    
    Select Case cmd(0)
        Case "DEBUGGER_INIT" 'DEBUGGER_INIT: hDebugObj
            'reint structures here
            hDebugObject = CLng(cmd(1))
            If Form1.scivb.DirectSCI.GetLineCount <> dbg_SourceLineCount(hDebugObject) Then LoadFlatFile
            InitDebuggerBpx
            handled = True
            
        Case "Local-Variable-Name"
            Set v = New CVariable
            v.name = cmd(1)
            v.value = GetVariableValue(v.name)
            v.varType = VariableType(v.name)
            variables.Add v
            handled = True
            
        Case "Global-Variable-Name"
            Set v = New CVariable
            v.isGlobal = True
            v.name = cmd(1)
            v.value = GetVariableValue(v.name)
            v.varType = VariableType(v.name)
            variables.Add v
            handled = True
            
        Case "Call-Stack"
            Set c = New cCallStack
            c.index = callStack.Count
            c.lineNo = CLng(cmd(1))
            c.func = cmd(2)
            callStack.Add c
            handled = True
            
        Case "Array-Variable" 'Array-Variable:%d:%d:%s , index, varType, buf);
            Set v = New CVariable
            v.index = CLng(cmd(1))
            v.varType = VariableTypeToString(CLng(cmd(2)))
            v.value = cmd(3) 'if is array then aryPointer will be parsed from value..
            variables.Add v
            handled = True
        
        Case "Current-Line":
            handled = True 'we dont need these anymore..
    
    'Line: %s\r\n
    'Message: %s\r\n
    'Value: %s\r\n
      
    End Select
    
    If Not handled Then
        Form1.txtDebug = Form1.txtDebug & msg
    End If

End Sub

Public Sub vb_stdout(ByVal t As cb_type, ByVal lpMsg As Long, ByVal sz As Long)

    Dim b() As Byte
    Dim msg As String
    
    ReDim b(sz)
    CopyMemory b(0), ByVal lpMsg, sz
    msg = StrConv(b, vbUnicode)
    If Right(msg, 1) = Chr(0) Then msg = Left(msg, Len(msg) - 1)
    
    Select Case t
        Case cb_debugger: HandleDebugMessage msg
        Case cb_engine:   HandleEngineMessage msg
        Case Default:
                          If t = cb_dbgout Then msg = "DBG> " & msg
                          With Form1.txtOut
                               .Text = .Text & Replace(msg, vbLf, vbCrLf)
                               .Refresh
                               DoEvents
                          End With
    End Select
    
End Sub

Private Sub HandleEngineMessage(msg As String)
    On Error Resume Next
    Dim tmp() As String
    
    tmp = Split(msg, ":")
    If tmp(0) = "ENGINE_PRECONFIG" Then
        hProgram = CLng(tmp(1))
    ElseIf tmp(0) = "ENGINE_DESTROY" Then
        hProgram = 0
        hDebugObject = 0
    End If
    
End Sub





















