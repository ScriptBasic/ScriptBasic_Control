Attribute VB_Name = "Module1"
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, Source As Any, ByVal Length As Long)

Enum cb_type
    cb_output = 0
    cb_dbgout = 1
    cb_dbgmsg = 2
End Enum


Public Sub vb_stdout(ByVal t As cb_type, ByVal lpMsg As Long, ByVal sz As Long)

    Dim b() As Byte
    Dim msg As String
    
    ReDim b(sz)
    CopyMemory b(0), ByVal lpMsg, sz
    msg = StrConv(b, vbUnicode)
    If Right(msg, 1) = Chr(0) Then msg = Left(msg, Len(msg) - 1)
    
    If t = cb_dbgmsg Then msg = "DBG> " & msg
    
    With Form1.txtOut
        .Text = .Text & Replace(msg, vbLf, vbCrLf)
        .Refresh
        DoEvents
    End With
    
End Sub
