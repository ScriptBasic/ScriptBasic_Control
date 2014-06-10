Attribute VB_Name = "Module1"
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'get the currently executing line number
Public Declare Function GetCurrentDebugLine Lib "sb_engine" (ByVal hDebug As Long) As Long

'get a variables value from its name
Private Declare Function dbg_getVarVal Lib "sb_engine" (ByVal hDebug As Long, ByRef varName As Byte, ByRef lpBuf As Byte, ByRef bufSz As Long) As Long

'enumerate global and local variable names (uses vbstdOut callback)
Private Declare Sub dbg_EnumVars Lib "sb_engine" (ByVal hDebug As Long)

Private Declare Function dbg_VarTypeFromName Lib "sb_engine" (ByVal hDebug As Long, ByRef varName As Byte) As Long

'set or remove a breakpoint on a line
Private Declare Sub dbg_ModifyBreakpoint Lib "sb_engine" (ByVal hDebug As Long, ByVal lineNo As Long, ByVal value As Long)

Public Declare Function dbg_LineCount Lib "sb_engine" (ByVal hDebug As Long) As Long

'we may not need this we should track internally...
Private Declare Function dbg_isBpSet Lib "sb_engine" (ByVal hDebug As Long, ByVal lineNo As Long) As Long

Public hDebugObject As Long     'handle to the current debug object - pDO
Public readyToReturn As Boolean
Public dbg_cmd As String
Public running As Boolean
Public variables As New Collection 'of CVariable
Public breakpoints As New Collection 'of CBreakPoint

Enum cb_type
    cb_output = 0
    cb_dbgout = 1
    cb_debugger = 2
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

Public Function BreakPointExists(lineNo As Long) As Boolean

'    Dim b As CBreakpoint
'    For Each b In breakpoints
'        If b.lineNo = lineNo Then
'            BreakPointExists = True
'            Exit Function
'        End If
'    Next

    Dim b As CBreakpoint
    On Error Resume Next
    Set b = breakpoints("bp:" & lineNo)
    If Not b Is Nothing Then BreakPointExists = True
    
End Function

Public Sub ToggleBreakPoint(lineNo As Long)
    If BreakPointExists(lineNo) Then
        RemoveBreakpoint lineNo
    Else
        SetBreakpoint lineNo
    End If
End Sub

Public Sub SetBreakpoint(lineNo As Long)
    Dim b As CBreakpoint
    
    If BreakPointExists(lineNo) Then Exit Sub
    If running Then dbg_ModifyBreakpoint hDebugObject, lineNo + 1, 1
    
    Set b = New CBreakpoint
    b.lineNo = lineNo
    breakpoints.Add b, "bp:" & lineNo
    
    Form1.scivb.SetMarker lineNo
End Sub

Public Sub RemoveBreakpoint(lineNo As Long)
    If Not BreakPointExists(lineNo) Then Exit Sub
    If running Then dbg_ModifyBreakpoint hDebugObject, lineNo + 1, 0
    Form1.scivb.DeleteMarker lineNo
    breakpoints.Remove "bp:" & lineNo
End Sub

Private Sub InitDebuggerBpx()
    Dim b As CBreakpoint
    For Each b In breakpoints
        dbg_ModifyBreakpoint hDebugObject, b.lineNo + 1, 1
    Next
End Sub

Public Function VariableType(varName As String) As String
    
    Dim x  As sb_VarTypes
    Dim v() As Byte
    
    v() = StrConv(varName & Chr(0), vbFromUnicode)
    x = dbg_VarTypeFromName(hDebugObject, v(0))
    
    types = Array("LONG", "DOUBLE", "STRING", "ARRAY", "REF", "UNDEF")
    
    If x < 0 Or x > 5 Then
        VariableType = "???"
    Else
        VariableType = LCase(types(x))
    End If
    
End Function

Public Sub DebuggerCmd(cmd As String)
    dbg_cmd = cmd
    readyToReturn = True
End Sub

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
            GetVariableValue = Left(GetVariableValue, i)
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
    
    Form1.SyncUI
    
    readyToReturn = False
    While Not readyToReturn
        DoEvents
        Sleep 20
    Wend
    
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
    
    Form1.sbStatus.Panels(1).Text = "Running"
    
End Function

Public Sub HandleDebugMessage(msg As String)

    Dim cmd() As String
    Dim v As CVariable
    Dim handled As Boolean
    
    cmd = Split(msg, ":", 2)
    
    Select Case cmd(0)
        Case "DEBUGGER_INIT" 'DEBUGGER_INIT: hDebugObj
            'reint structures here
            hDebugObject = CLng(cmd(1))
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
            
    'Source-File: %s\r\n
    'Current-Line: %u\r\n
    'Break-Point: %s\r\n = 1/0
    'Line-Number: %u\r\n
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
    
    If t = cb_debugger Then
        HandleDebugMessage msg
    Else
        If t = cb_dbgout Then msg = "DBG> " & msg
        With Form1.txtOut
            .Text = .Text & Replace(msg, vbLf, vbCrLf)
            .Refresh
            DoEvents
        End With
    End If
    
End Sub
