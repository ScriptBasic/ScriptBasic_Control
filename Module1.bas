Attribute VB_Name = "Module1"
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function GetCurrentDebugLine Lib "sb_engine" (ByVal hDebug As Long) As Long

Public hDebugObject As Long
Public readyToReturn As Boolean
Public dbg_cmd As String

Enum cb_type
    cb_output = 0
    cb_dbgout = 1
    cb_debugger = 2
End Enum

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
    cmd = Split(msg, ":")
    
    Select Case cmd(0)
        Case "DEBUGGER_INIT" 'DEBUGGER_INIT:hDebug:%d
            'todo reint structures
            hDebugObject = CLng(cmd(2))
        
    'Source-File: %s\r\n
    'Current-Line: %u\r\n
    'Break-Point: %s\r\n = 1/0
    'Line-Number: %u\r\n
    'Line: %s\r\n
    'Message: %s\r\n
    'Value: %s\r\n
    'Local-Variable-Name: %s\r\n
    'Local-Variable-Value: %s\r\n
    'Global-Variable-Name: %s\r\n
    'Global-Variable-Value: %s\r\n
    
    End Select
    
    Form1.txtDebug = Form1.txtDebug & msg

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
