Attribute VB_Name = "modBreakpoints"
'breakpoint strategy..so we want to save breakpoints across runs, but only allow
'them to be set at runtime debug mode because then line numbers are locked in
'and knowable (import statements pull entire files in)..
'so we need to set breakpoints at runtime, and restore them on debugger start
'if the source line text hasnt changed..we also need to clearUI breakpoints visually
'when the run ends. also we limit which lines breakpoints can be added on some if it
'doesnt look like executable code to make more sense.

Private Declare Sub dbg_ModifyBreakpoint Lib "sb_engine" (ByVal hDebug As Long, ByVal lineNo As Long, ByVal Value As Long)

Public breakpoints As New Collection 'of CBreakPoint



Function isExecutableLine(lineNo As Long) As Boolean
    Dim tmp As String
    On Error Resume Next
    
    tmp = LCase(frmMain.scivb.GetLineText(lineNo))
    tmp = Trim(Replace(tmp, vbTab, Empty))
    tmp = Replace(tmp, vbCr, Empty)
    tmp = Replace(tmp, vbLf, Empty)
    
    If Len(tmp) = 0 Then GoTo fail
    If Left(tmp, 1) = "'" Then GoTo fail 'is comment
    If Left(tmp, 5) = "local" Then GoTo fail
    If Left(tmp, 5) = "const" Then GoTo fail
    If Left(tmp, 8) = "function" Then GoTo fail  'functio/sub start lines are hit more than you expect, once as it skips over it, so we block it as bp cause confusing..
    If Left(tmp, 3) = "sub" Then GoTo fail
    If Left(tmp, 3) = "rem" Then GoTo fail
    
    
    isExecutableLine = True
Exit Function
fail: isExecutableLine = False
End Function

Public Function BreakPointExists(lineNo As Long) As Boolean

    Dim b As CBreakpoint
    On Error Resume Next
    Set b = breakpoints("bp:" & lineNo)
    If Not b Is Nothing Then BreakPointExists = True
    
End Function

Public Sub ToggleBreakPoint(Optional ByVal lineNo As Long)
    
    If lineNo = 0 Then lineNo = frmMain.scivb.CurrentLine
    
    If running Then
        If isExecutableLine(lineNo) Then
            If BreakPointExists(lineNo) Then
                RemoveBreakpoint lineNo
            Else
                SetBreakpoint lineNo
            End If
        Else
            frmMain.Caption = "Can not set breakpoint here not an executable line"
        End If
    Else
        'because import statements change line numbers, and lines can change during edits..
        frmMain.Caption = "Can only set breakpoint once debugging"
    End If
                        
    
End Sub

Public Sub SetBreakpoint(lineNo As Long)
    Dim b As CBreakpoint
    
    If BreakPointExists(lineNo) Then Exit Sub
    If running Then dbg_ModifyBreakpoint hDebugObject, lineNo + 1, 1
    
    Set b = New CBreakpoint
    b.lineNo = lineNo
    b.Source = frmMain.scivb.GetLineText(lineNo)
    breakpoints.Add b, "bp:" & lineNo
    
    frmMain.scivb.SetMarker lineNo
End Sub

Public Sub RemoveBreakpoint(lineNo As Long)
    If Not BreakPointExists(lineNo) Then Exit Sub
    If running Then dbg_ModifyBreakpoint hDebugObject, lineNo + 1, 0
    frmMain.scivb.DeleteMarker lineNo
    breakpoints.Remove "bp:" & lineNo
End Sub

Sub ClearUIBreakpoints()
    Dim b As CBreakpoint
    For Each b In breakpoints
        frmMain.scivb.DeleteMarker b.lineNo
    Next
End Sub

Sub RemoveAllBreakpoints()
    Dim b As CBreakpoint
    For Each b In breakpoints
        RemoveBreakpoint b.lineNo
    Next
End Sub

Sub InitDebuggerBpx()
    Dim b As CBreakpoint
    For Each b In breakpoints
        If b.Source = frmMain.scivb.GetLineText(b.lineNo) Then
            dbg_ModifyBreakpoint hDebugObject, b.lineNo + 1, 1
            frmMain.scivb.SetMarker b.lineNo
        End If
    Next
End Sub


