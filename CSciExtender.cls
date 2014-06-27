VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSciExtender"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'http://www.scintilla.org/aprilw/SciLexer.bas

Public Event MarginClick(lline As Long, Position As Long, margin As Long, modifiers As Long)
Public Event MouseDwellStart(lline As Long, Position As Long)
Public Event MouseDwellEnd(lline As Long, Position As Long)
'Public Event UpdateUI()

Private WithEvents sc As clsSubClass
Attribute sc.VB_VarHelpID = -1
Private ucHwnd As Long
Private owner As SciSimple
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal length As Long)

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Declare Function GetCursorPos Lib "User32" (lpPoint As POINTAPI) As Long
Private Declare Function WindowFromPoint Lib "User32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function GetClassName Lib "User32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
         
Private Const WM_NOTIFY = &H4E
Private Const SCI_LINEFROMPOSITION = 2166
Private Const SCN_MARGINCLICK = 2010
Private Const SCN_DWELLSTART = 2016
Private Const SCN_DWELLEND = 2017
Private Const SCN_UPDATEUI = 2007

Private Type NMHDR
    hwndFrom As Long
    idFrom As Long
    Code As Long
End Type

Private Type SCNotification
    NotifyHeader As NMHDR
    Position As Long
    ch As Long
    modifiers As Long
    modificationType As Long
    Text As Long
    length As Long
    linesAdded As Long
    message As Long
    wParam As Long
    lParam As Long
    line As Long
    foldLevelNow As Long
    foldLevelPrev As Long
    margin As Long
    listType As Long
    x As Long
    y As Long
End Type

Private Declare Function GetFocus Lib "User32" () As Long
Private Declare Function GetParent Lib "User32" (ByVal hwnd As Long) As Long

Private Function GetUCHwnd(ctrl As Control) As Long
    
    Dim a As Long, b As Long
    
    ctrl.SetFocus
    a = GetFocus

    Do
        b = a
        a = GetParent(a)
    Loop While a <> ctrl.Parent.hwnd
    
    GetUCHwnd = b
    
End Function

Function Init(scivb As SciSimple, Optional dwelltime As Long = 600) As Boolean
    Dim i As Long
     
    Set owner = scivb
    Set sc = New clsSubClass
     
    ucHwnd = GetUCHwnd(scivb)
    If sc.AttachMessage(ucHwnd, WM_NOTIFY) Then
        For i = 0 To 4
            owner.DirectSCI.SetMarginSensitiveN i, 1
        Next
        owner.DirectSCI.SetMouseDwellTime dwelltime
        Init = True
    End If
    
End Function

Private Sub Class_Terminate()
    sc.DetatchMessage ucHwnd, WM_NOTIFY
End Sub

Private Sub sc_MessageReceived(hwnd As Long, wMsg As Long, wParam As Long, lParam As Long, Cancel As Boolean)

    Dim scMsg As SCNotification
    Dim tHdr As NMHDR
    
    If wMsg = WM_NOTIFY Then
        CopyMemory scMsg, ByVal lParam, Len(scMsg)
        tHdr = scMsg.NotifyHeader
        If (tHdr.hwndFrom = owner.sciHWND) Then HandleSciMsg tHdr, scMsg
        Exit Sub
    End If
    
End Sub


Private Sub HandleSciMsg(tHdr As NMHDR, scMsg As SCNotification)
    Dim strTmp As String
    Dim zPos As Long
    Dim chl As String, strMatch As String
    Dim lPos As Long
    Dim pos As Long, pos2 As Long
    Dim lline As Long
    
    Select Case tHdr.Code
                                
            'Case SCN_UPDATEUI
            '                    RaiseEvent UpdateUI 'PosChanged was depreceated and removed...
                                
            Case SCN_MARGINCLICK
                                lline = owner.DirectSCI.SendEditor(SCI_LINEFROMPOSITION, scMsg.Position)
                                RaiseEvent MarginClick(lline, scMsg.Position, scMsg.margin, scMsg.modifiers)
                                
            Case SCN_DWELLSTART
                                lline = owner.DirectSCI.SendEditor(SCI_LINEFROMPOSITION, scMsg.Position)
                                RaiseEvent MouseDwellStart(lline, scMsg.Position)
            Case SCN_DWELLEND
                                lline = owner.DirectSCI.SendEditor(SCI_LINEFROMPOSITION, scMsg.Position)
                                RaiseEvent MouseDwellEnd(lline, scMsg.Position)

    End Select
    
End Sub

Public Function WordUnderMouse(pos As Long, Optional ignoreWhiteSpace As Boolean = False) As String
    Dim sWord As Long, eWord As Long
    
    On Error Resume Next
    'behavior warning.. space characters are counted as words we should count space chars
    'back from pos and pos -= spaceCount
    
    If ignoreWhiteSpace Then pos = pos - GetSpaceCountBack(pos)
    
    With owner
        sWord = .DirectSCI.WordStartPosition(pos, True) + 1
        eWord = .DirectSCI.WordEndPosition(pos, True) + 1
        WordUnderMouse = Mid(.Text, sWord, eWord - sWord)
    End With
    
End Function

'gets the number of spaces counting back to next non white space character
Private Function GetSpaceCountBack(pos As Long)
    On Error Resume Next
    Dim lline As Long, curpos As Long, lText As String, i As Long
    Dim lStart As Long, curCol As Long, b() As Byte, count As Long
    
    lline = owner.DirectSCI.LineFromPosition(pos)
    lText = owner.GetLineText(lline)
    lStart = owner.PositionFromLine(lline)
    curCol = pos - lStart
    
    lText = Left(lText, curCol)
    If Len(lText) = 0 Then Exit Function
    
    b() = StrConv(lText, vbFromUnicode)
    For i = UBound(b) To 0 Step -1
        If b(i) = &H20 Or b(i) = 9 Then
            count = count + 1
        Else
            Exit For
        End If
    Next
    
    GetSpaceCountBack = count
    
End Function

Sub LockEditor(Optional locked As Boolean = True)
    Dim i As Long
    
    With owner
        If locked Then
            .ReadOnly = True
            .DirectSCI.StyleSetBack 32, &HF0F0F0
            For i = 0 To 127
                .DirectSCI.StyleSetBack i, &HF0F0F0
            Next i
        Else
            .ReadOnly = False
            .SetHighlighter "VB" 'sets back to defaults..
        End If
    End With
 
End Sub

Function isMouseOverCallTip() As Boolean
    Dim p As POINTAPI
    Dim hWin As Long
    Dim sz As Long
    Dim sClassName As String * 100
    
    On Error Resume Next
    
    GetCursorPos p
    hWin = WindowFromPoint(p.x, p.y)
    sz = GetClassName(hWin, sClassName, 100)
    If Left(sClassName, sz) = "CallTip" Then isMouseOverCallTip = True
    
End Function

