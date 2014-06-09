Attribute VB_Name = "Module1"
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
 
Public readyToReturn As Boolean

Enum cb_type
    cb_output = 0
    cb_dbgout = 1
    cb_debugger = 2
End Enum

Public Function GetDebuggerCommand(ByVal buf As Long, ByVal sz As Long) As Long
    
    MsgBox "Waiting on debugger command!"
    
    readyToReturn = False
    While Not readyToReturn
        DoEvents
        Sleep 20
    Wend
    
    
    
End Function

Public Sub HandleDebugMessage(msg As String)
    'DEBUGGER_INIT
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
    
    Form1.List1.AddItem msg

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
