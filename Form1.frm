VERSION 5.00
Object = "{FBE17B58-A1F0-4B91-BDBD-C9AB263AC8B0}#78.0#0"; "scivb_lite.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   12000
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   13905
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   12000
   ScaleWidth      =   13905
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   510
      Left            =   9810
      TabIndex        =   10
      Top             =   8775
      Width           =   1230
   End
   Begin MSComctlLib.ListView lvVars 
      Height          =   1815
      Left            =   90
      TabIndex        =   7
      Top             =   9630
      Width           =   7530
      _ExtentX        =   13282
      _ExtentY        =   3201
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "scope"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "type"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "value"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   510
      Left            =   12015
      TabIndex        =   6
      Top             =   45
      Width           =   1230
   End
   Begin VB.TextBox txtDebug 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1725
      Left            =   8595
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   5
      Top             =   6750
      Width           =   5145
   End
   Begin VB.TextBox txtOut 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2760
      Left            =   135
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   2
      Top             =   6750
      Width           =   8340
   End
   Begin SCIVB_LITE.SciSimple scivb 
      Height          =   5865
      Left            =   135
      TabIndex        =   0
      Top             =   585
      Width           =   13650
      _ExtentX        =   24077
      _ExtentY        =   10345
   End
   Begin MSComctlLib.ImageList ilToolbars 
      Left            =   11025
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   15
      ImageHeight     =   15
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   21
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0000
            Key             =   "Toggle Bookmark"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":010A
            Key             =   "Execute"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0216
            Key             =   "Next Bookmark"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0320
            Key             =   "Previous Bookmark"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":042A
            Key             =   "Clear All Bookmarks"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0534
            Key             =   "Comment Block"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":063E
            Key             =   "UnCommentBlock"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0748
            Key             =   "Tab Left"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0852
            Key             =   "Tab Right"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":095C
            Key             =   "Parse code"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0A66
            Key             =   "Run"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0B70
            Key             =   "Break"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0C7A
            Key             =   "Stop"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0D84
            Key             =   "Toggle Breakpoint"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0E8E
            Key             =   "Clear All Breakpoints"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0F98
            Key             =   "Step In"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":10A2
            Key             =   "Step Over"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":11AC
            Key             =   "Step Out"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":12B6
            Key             =   "Immediate"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":13C0
            Key             =   "Callstack"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":14CA
            Key             =   "Set Next"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbarDebug 
      Height          =   315
      Left            =   180
      TabIndex        =   3
      Top             =   225
      Width           =   3870
      _ExtentX        =   6826
      _ExtentY        =   556
      ButtonWidth     =   582
      ButtonHeight    =   556
      Style           =   1
      ImageList       =   "ilToolbars"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   15
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Execute"
            Object.ToolTipText     =   "Execute"
            ImageKey        =   "Execute"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Run"
            Object.ToolTipText     =   "Run"
            ImageKey        =   "Run"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Break"
            Object.ToolTipText     =   "Break"
            ImageKey        =   "Break"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Stop"
            Object.ToolTipText     =   "Stop"
            ImageKey        =   "Stop"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Breakpoint"
            Object.ToolTipText     =   "Toggle Breakpoint"
            ImageKey        =   "Toggle Breakpoint"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Clear Breakpoints"
            Object.ToolTipText     =   "Clear All Breakpoiunts"
            ImageKey        =   "Clear All Breakpoints"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Step In"
            Object.ToolTipText     =   "Step In"
            ImageKey        =   "Step In"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Step Over"
            Object.ToolTipText     =   "Step Over"
            ImageKey        =   "Step Over"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Step Out"
            Object.ToolTipText     =   "Step Out"
            ImageKey        =   "Step Out"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Run to Cursor"
            Object.ToolTipText     =   "Run to Cursor"
            ImageKey        =   "Set Next"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "Immediate"
            Object.ToolTipText     =   "Immediate Window"
            ImageKey        =   "Immediate"
            Style           =   1
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "Callstack"
            Object.ToolTipText     =   "Callstack"
            ImageKey        =   "Callstack"
            Style           =   1
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar sbStatus 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   11745
      Width           =   13905
      _ExtentX        =   24527
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Bevel           =   0
            Object.Width           =   21458
            Text            =   "Ready"
            TextSave        =   "Ready"
            Key             =   "Status"
            Object.ToolTipText     =   "Shows status of script"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Text            =   "1x1"
            TextSave        =   "1x1"
            Key             =   "SelStart"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvCallStack 
      Height          =   1815
      Left            =   7695
      TabIndex        =   8
      Top             =   9630
      Width           =   6090
      _ExtentX        =   10742
      _ExtentY        =   3201
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Line"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Function"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label lblStatus 
      Height          =   285
      Left            =   4185
      TabIndex        =   9
      Top             =   270
      Width           =   6540
   End
   Begin VB.Label Label1 
      Caption         =   "Output"
      Height          =   240
      Left            =   180
      TabIndex        =   1
      Top             =   6525
      Width           =   1860
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuOpen 
         Caption         =   "Open"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "Save"
      End
      Begin VB.Menu mnuSpacer 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOptions 
         Caption         =   "Options"
      End
   End
   Begin VB.Menu mnuCallStackPopup 
      Caption         =   "mnuCallStackPopup"
      Begin VB.Menu mnuExecuteTillReturn 
         Caption         =   "Execute Until Return"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long


Private Declare Function run_script Lib "sb_engine" (ByVal lpLibFileName As String, ByVal use_debugger As Long) As Long
Private Declare Sub GetErrorString Lib "sb_engine" (ByVal iErrorCode As Long, ByVal buf As String, ByVal sz As Long)
Private Declare Sub SetCallBacks Lib "sb_engine" (ByVal msgProc As Long, ByVal dbgCmdProc As Long, ByVal hostResolverProc As Long)


Dim loadedFile As String
Dim hsbLib As Long
Dim lastEIP As Long
Public hasImports As Boolean

Const SC_MARK_CIRCLE = 0
Const SC_MARK_ARROW = 2
'http://www.scintilla.org/aprilw/SciLexer.bas
'
'if we subclass scivb.sciHWND we can use Dwelltime for mouse over variable popups, and
'we can capture gutterclick events to set breakpoints..then if i like it enough, I can
'break compatability on the ocx to add them in permanent if need be..or just keep as external
'extension module..

Dim selCallStackItem As ListItem
Dim selVariable As ListItem

Private Sub cmdManual_Click()
    'txtDebug.Text = Empty
    'dbg_cmd = txtCmd.Text
    'readyToReturn = True
    
    MsgBox GetVariableValue(txtCmd.Text)
    
End Sub



Private Sub RefreshVariables()
    
    Dim li As ListItem
    Dim vars As Collection
    Dim v As CVariable
    
    lvVars.ListItems.Clear
    Set vars = EnumVariables()
    
    For Each v In vars
        Set li = lvVars.ListItems.Add(, , IIf(v.isGlobal, "Global", "Local"))
        li.SubItems(1) = v.name
        li.SubItems(2) = v.varType
        li.SubItems(3) = v.value
        If v.varType = "array" Then li.Tag = v.pAryElement
    Next
    
End Sub

Private Sub RefreshCallStack()
    Dim c As Collection
    Dim cs As cCallStack
    Dim li As ListItem
    
    lvCallStack.ListItems.Clear
    
    Set c = EnumCallStack()
    
    For Each cs In c
        Set li = lvCallStack.ListItems.Add(, , cs.lineNo)
        li.SubItems(1) = cs.func
    Next
    
End Sub

Sub LockEditor(Optional locked As Boolean = True)
    Dim i As Long
    
    With scivb
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

Private Sub lvCallStack_ItemClick(ByVal Item As MSComctlLib.ListItem)
    scivb.GotoLine CLng(Item.Text)
    Set selCallStackItem = Item
End Sub

Private Sub lvVars_DblClick()
    If selVariable Is Nothing Then Exit Sub
    If selVariable.SubItems(2) <> "array" Then Exit Sub
    
    Dim c As Collection
    Dim varName As String
    
    varName = selVariable.SubItems(1)
    Set c = EnumArrayVariables(varName)
    If c.Count > 0 Then
        frmAryDump.DumpArrayValues varName, c
    End If
    
End Sub

Private Sub lvVars_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Set selVariable = Item
End Sub

Private Sub mnuExecuteTillReturn_Click()
    MsgBox "todo: disable all breakpoints,run to line selCallStackItem.text + 1, reenable breakpoints", vbInformation
End Sub

 
Private Sub mnuOpen_Click()
    Dim f As String
    f = dlg.OpenDialog(AllFiles)
    If Len(f) = 0 Then Exit Sub
    LoadFile f
End Sub



Private Sub mnuOptions_Click()
    frmOptions.Show 1, Me
End Sub

Private Sub mnuSave_Click()
    scivb.SaveFile loadedFile
End Sub

Private Sub scivb_DoubleClick()
    Dim word As String
    word = scivb.CurrentWord
    If Len(word) < 20 Then
        Me.Caption = "  " & scivb.hilightWord(word, , vbTextCompare) & " instances of '" & word & " ' found"
    End If
End Sub

Private Sub scivb_MouseUp(Button As Integer, Shift As Integer, X As Long, Y As Long)
    If scivb.SelLength > 0 And scivb.SelLength < 20 Then
        Dim word As String
        word = scivb.SelText
        If Len(word) < 20 Then
            Me.Caption = "  " & scivb.hilightWord(word, , vbTextCompare) & " instances of '" & word & " ' found"
        End If
    End If
End Sub

Private Sub scivb_KeyDown(KeyCode As Long, Shift As Long)

    'Debug.Print KeyCode & " " & Shift
    Select Case KeyCode
        Case vbKeyF2: ToggleBreakPoint
        Case vbKeyF5: If running Then DebuggerCmd "R" Else ExecuteScript True
        Case vbKeyF7: DebuggerCmd "s" 'step into
        Case vbKeyF8: DebuggerCmd "S" 'step over
        Case vbKeyF9: DebuggerCmd "o" 'step out
    End Select
    
End Sub

Private Sub scivb_KeyUp(KeyCode As Long, Shift As Long)

    Dim curWord As String
    Dim txt As String
    Dim curPos As Long
    Dim prevChar As String
    
    If KeyCode = 186 Then
        curPos = scivb.GetCaretInLine()
        txt = scivb.GetLineText(scivb.CurrentLine)

        If curPos < 3 Then Exit Sub
        prevChar = Mid(txt, curPos - 1, 1)
        If prevChar <> ":" Then Exit Sub
        
        scivb.GotoCol curPos - 2
        curWord = scivb.CurrentWord
        scivb.GotoCol curPos
        
        If curWord = "nt" Then

            scivb.ShowAutoComplete ":RegRead :RegDel :RegWrite :MsgBox :ShutDown :ListProcesses :StartService :StopService :PauseService :ContinueService :HardLink"

        End If
    End If


End Sub



Private Sub tbarDebug_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.key
        Case "Execute":   If running Then DebuggerCmd "R" Else ExecuteScript
        Case "Run":       If running Then DebuggerCmd "R" Else ExecuteScript True
        Case "Stop":      DebuggerCmd "q"
        Case "Step In":   DebuggerCmd "s"
        Case "Step Over": DebuggerCmd "S"
        Case "Step Out":  DebuggerCmd "o"
        Case "Run to Cursor": RunToLine scivb.CurrentLine + 1
        Case "Breakpoint": ToggleBreakPoint
        Case "Clear Breakpoints": RemoveAllBreakpoints
        'Case "Immediate"
        'Case "Callstack"
    End Select
    
End Sub


Private Sub ExecuteScript(Optional withDebugger As Boolean)

    Dim rv As Long
    Dim buf As String
    
    
    txtOut.Text = Empty
    
    LockEditor
    If scivb.isDirty Then scivb.SaveFile loadedFile
    
    running = True
    lblStatus = IIf(withDebugger, "Debugging", "Running")
    
    rv = run_script(loadedFile, IIf(withDebugger, 1, 0))
    
'    If rv <> 0 Then
'        buf = String(255, " ")
'        Call GetErrorString(rv, buf, 255)
'        txtOut = txtOut & vbCrLf & "Error: " & buf
'    End If
    
    lblStatus = "Idle"
    running = False
    LockEditor False
    scivb.HighLightActiveLine = False
    scivb.DeleteMarker lastEIP, 1
    ClearUIBreakpoints
    If hasImports Then scivb.LoadFile loadedFile
    
End Sub

Private Sub Form_Load()
    
    mnuCallStackPopup.Visible = False
    hsbLib = LoadLibrary(App.path & "\engine\sb_engine.dll")
    
    If hsbLib = 0 Then
        MsgBox "Failed to load sb_engine.dll by explicit path?"
    End If
    
    includeDir = GetMySetting("includeDir", App.path & "\include\")
    moduleDir = GetMySetting("moduleDir", App.path & "\modules\")
    SetConfig includeDir, moduleDir

    SetCallBacks AddressOf vb_stdout, AddressOf GetDebuggerCommand, AddressOf HostResolver
    scivb.LoadHighlighter App.path & "\dependancies\vb.bin"
    
    scivb.DirectSCI.HideSelection False
    scivb.DirectSCI.MarkerDefine 2, SC_MARK_CIRCLE
    scivb.DirectSCI.MarkerSetFore 2, vbRed 'set breakpoint color
    scivb.DirectSCI.MarkerSetBack 2, vbRed
    
    scivb.DirectSCI.MarkerDefine 1, SC_MARK_ARROW
    scivb.DirectSCI.MarkerSetFore 1, vbBlack 'current eip
    scivb.DirectSCI.MarkerSetBack 1, vbYellow
  
    lvCallStack.ColumnHeaders(2).Width = lvCallStack.Width - lvCallStack.ColumnHeaders(2).Left - 100
    lvVars.ColumnHeaders(lvVars.ColumnHeaders.Count).Width = lvVars.Width - lvVars.ColumnHeaders(lvVars.ColumnHeaders.Count).Left - 100
    
    'App.Path & "\scripts\com_voice_test.sb"
    'LoadFile App.Path & "\scripts\functions.txt"
    
    'AddObject "Form1", Me
    'AddString "test", "this is my string from vb!"
    'LoadFile App.Path & "\scripts\GetHostObject.sb"
    
    scivb.DirectSCI.AutoCSetIgnoreCase True
    scivb.DisplayCallTips = True
    Call scivb.LoadCallTips(App.path & "\dependancies\calltips.txt")
    LoadFile App.path & "\scripts\importNT.sb"
    
End Sub

Sub LoadFile(fPath As String)
   
   loadedFile = fPath
   scivb.DeleteAllMarkers
   scivb.LoadFile loadedFile
   Set breakpoints = New Collection
   
End Sub


Private Sub Form_Resize()
    On Error Resume Next
    With scivb
        .Width = Me.Width - .Left - 200
        '.Height = Me.Height - .Top - 500
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveMySetting("includeDir", includeDir)
    Call SaveMySetting("moduleDir", moduleDir)
    FreeLibrary hsbLib
End Sub

Public Sub SyncUI()
    
    Dim curline As Long
    
    scivb.DeleteMarker lastEIP, 1
    
    curline = GetCurrentDebugLine(hDebugObject)
    scivb.SetMarker curline, 1
    lastEIP = curline
    
    scivb.GotoLine curline
    scivb.HighLightActiveLine = True
    scivb.SetFocus
    
    RefreshVariables
    RefreshCallStack
    
End Sub

Public Function Alert(msg As String)
    MsgBox msg
End Function
