Attribute VB_Name = "modSBApi"
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, Source As Any, ByVal length As Long)
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

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
'int __stdcall SetVariable(pSbProgram pProgram, int isLong, BSTR* bvarName, BSTR* bbuf){
Private Declare Function sbSetVariable Lib "sb_engine" (ByVal hProgram As Long, ByVal isLong As Long, ByVal bstrVarName As Long, ByVal bstrValue As Long) As Long

Public hProgram As Long
Public hDebugObject As Long     'handle to the current debug object - pDO
Public readyToReturn As Boolean
Public dbg_cmd As Debug_Commands
Public running As Boolean
Public variables As New Collection 'of CVariable
Public callStack As New Collection 'of CCallStack
Public flatFile As String
Public hadError As Boolean
Public shuttingDown As Boolean

Public includeDir As String, moduleDir As String
Global dlg As New CCmnDlg
Global Const LANG_US = &H409

Enum cb_type
    cb_output = 0
    cb_dbgout = 1
    cb_debugger = 2
    cb_engine = 3
    cb_error = 4
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

Enum Debug_Commands
    dc_NotSet = 0
    dc_Run = 1
    dc_stepinto = 3
    dc_StepOut = 4
    dc_StepOver = 5
    dc_RunToLine = 6
    dc_Quit = 7
    dc_Manual = 8
End Enum

Function SetVariable(ByVal name As String, ByVal Value As String) As Boolean
    
    On Error Resume Next
    Dim v As Long
    Dim isNumeric As Long
    
    If InStr(name, "::") < 1 Then name = "main::" & name
    
    If Left(Value, 2) = "0x" Then
        v = CLng("&h" & Mid(Value, 3))
        If Err.Number = 0 Then
            Value = v
            isNumeric = 1
        End If
    Else
        v = CLng(Value)
        If Err.Number = 0 Then isNumeric = 1
    End If
    
    SetVariable = sbSetVariable(hProgram, isNumeric, VarPtr(name), VarPtr(Value))
    
End Function


Function LoadFlatFile() As Boolean
    
    Dim tmp As String
    Dim b() As Byte
    
    tmp = GetFreeFileName(Environ("temp"))
    b() = StrConv(tmp & Chr(0), vbFromUnicode)
    dbg_WriteFlatSourceFile hDebugObject, b(0)
    
    If FileExists(tmp) Then
        Form1.hasImports = True
        Form1.scivb.ReadOnly = False
        LoadFlatFile = Form1.scivb.LoadFile(tmp)
        Form1.scivb.ReadOnly = True
        flatFile = tmp
    End If
    
End Function

Public Sub RunToLine(lineNo As Long)
    dbg_RunToLine hDebugObject, lineNo
    DebuggerCmd dc_Manual
End Sub

Public Function EnumCallStack() As Collection
    
    Set callStack = Nothing
    dbg_EnumCallStack hDebugObject 'this goes into syncronous set of callbacks
    Set EnumCallStack = callStack
    
End Function

Public Function VariableTypeToString(X As sb_VarTypes) As String

    types = Array("LONG", "DOUBLE", "STRING", "ARRAY", "REF", "UNDEF")
    
    If X < 0 Or X > 5 Then
        VariableTypeToString = "???"
    Else
        VariableTypeToString = LCase(types(X))
    End If
    
End Function

Public Function VariableType(varName As String) As String
    
    Dim X  As sb_VarTypes
    Dim v() As Byte
    
    v() = StrConv(varName & Chr(0), vbFromUnicode)
    X = dbg_VarTypeFromName(hDebugObject, v(0))
    VariableType = VariableTypeToString(X)
    
End Function

Public Sub DebuggerCmd(cmd As Debug_Commands)
    
    With Form1
        .scivb.DeleteMarker .lastEIP, 1 'remove the yellow arrow
        .scivb.DeleteMarker .lastEIP, 3 'remove the yellow line backcolor
    End With
    
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

Public Function VbLineInput(ByVal buf As Long, ByVal sz As Long) As Long
    Dim b() As Byte
    Dim retVal As String
    VbLineInput = 0 'return value default..

    retVal = InputBox("Script is requesting input value:", "Script Basic Line Input")
    If Len(retVal) = 0 Then Exit Function
    
    If Len(retVal) < sz Then
        retVal = retVal & Chr(0)
        ReDim b(Len(retVal))
        b() = StrConv(retVal, vbFromUnicode)
        CopyMemory ByVal buf, b(0), Len(retVal)
        VbLineInput = Len(retVal) - 1
    Else
        MsgBox "Sorry VbLineInput is limited to " & sz & " characters", vbInformation
    End If
  
End Function


Public Function GetDebuggerCommand(ByVal buf As Long, ByVal sz As Long) As Long
        
    Dim b() As Byte
    Dim Source As String, curline As Long
    
    dbg_cmd = dc_NotSet
    
    'there are some lines we dont want to stop and show as execution to the user,
    'such as declares and function starts
    curline = GetCurrentDebugLine(hDebugObject)
    If Not BreakPointExists(curline) Then
        Source = LCase(Form1.scivb.GetLineText(curline))
        Source = Trim(Replace(Source, vbTab, Empty))
        If InStr(Source, " ") > 1 Then
            Source = Left(Source, InStr(Source, " ") - 1)
            If Source = "declare" Or Source = "function" Then
                dbg_cmd = dc_stepinto
            End If
        End If
    End If
    
    If dbg_cmd = dc_NotSet Then
        Form1.SyncUI
        
        'we block here until the UI sets the readyToReturn = true
        'this is not a CPU hog, and form remains responsive to user actions..
        readyToReturn = False
        While Not readyToReturn
            DoEvents
            Sleep 20
        Wend
    End If
    
    GetDebuggerCommand = dbg_cmd 'now were enum based..
    
'    If Len(dbg_cmd) < sz Then
'        dbg_cmd = dbg_cmd & Chr(0)
'        ReDim b(Len(dbg_cmd))
'        b() = StrConv(dbg_cmd, vbFromUnicode)
'        CopyMemory ByVal buf, b(0), Len(dbg_cmd)
'        GetDebuggerCommand = Len(dbg_cmd)
'    Else
'        GetDebuggerCommand = 0
'        MsgBox "Shouldnt happen!"
'    End If
    
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
            v.Value = GetVariableValue(v.name)
            v.varType = VariableType(v.name)
            variables.Add v
            handled = True
            
        Case "Global-Variable-Name"
            Set v = New CVariable
            v.isGlobal = True
            v.name = cmd(1)
            v.Value = GetVariableValue(v.name)
            v.varType = VariableType(v.name)
            variables.Add v
            handled = True
            
        Case "Call-Stack"
            Set c = New cCallStack
            c.Index = callStack.Count
            c.lineNo = CLng(cmd(1))
            c.func = cmd(2)
            callStack.Add c
            handled = True
            
        Case "Array-Variable" 'Array-Variable:%d:%d:%s , index, varType, buf);
            Set v = New CVariable
            v.Index = CLng(cmd(1))
            v.varType = VariableTypeToString(CLng(cmd(2)))
            v.Value = cmd(3) 'if is array then aryPointer will be parsed from value..
            variables.Add v
            handled = True
        
        Case "Current-Line":
            handled = True 'we dont need these anymore..
    
    'Line: %s\r\n
    'Message: %s\r\n
    'Value: %s\r\n
      
    End Select
    
    If Not handled Then
        'Form1.txtDebug = Form1.txtDebug & msg
    End If

End Sub

Public Sub vb_stdout(ByVal t As cb_type, ByVal lpMsg As Long, ByVal sz As Long)

    Dim b() As Byte
    Dim msg As String
    
    If shuttingDown Then Exit Sub
    
    ReDim b(sz)
    CopyMemory b(0), ByVal lpMsg, sz
    msg = StrConv(b, vbUnicode)
    If Right(msg, 1) = Chr(0) Then msg = Left(msg, Len(msg) - 1)
    
    Select Case t
        Case cb_debugger: HandleDebugMessage msg
        Case cb_engine:   HandleEngineMessage msg
        Case Else:
                          If t = cb_error Then
                                msg = "Error: " & IIf(Len(flatFile) > 0, flatFile, "") & vbCrLf & msg
                                hadError = True
                          End If
                          
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
        hadError = False
    ElseIf tmp(0) = "ENGINE_DESTROY" Then
        hProgram = 0
        hDebugObject = 0
        If Not hadError And FileExists(flatFile) Then
            Kill flatFile
            flatFile = Empty
        End If
    End If
    
End Sub





















