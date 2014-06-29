VERSION 5.00
Object = "{FBE17B58-A1F0-4B91-BDBD-C9AB263AC8B0}#78.0#0"; "scivb_lite.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   Caption         =   "Script Basic IDE"
   ClientHeight    =   9870
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   13905
   BeginProperty Font 
      Name            =   "Courier"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   9870
   ScaleWidth      =   13905
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ListView lvErrors 
      Height          =   1050
      Left            =   4500
      TabIndex        =   6
      Top             =   5850
      Width           =   1860
      _ExtentX        =   3281
      _ExtentY        =   1852
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Line"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "File"
         Object.Width           =   3881
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Error"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Timer tmrHideCallTip 
      Enabled         =   0   'False
      Interval        =   600
      Left            =   9720
      Top             =   135
   End
   Begin MSComctlLib.ListView lvVars 
      Height          =   1050
      Left            =   11745
      TabIndex        =   5
      Top             =   5895
      Width           =   1860
      _ExtentX        =   3281
      _ExtentY        =   1852
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
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
   Begin MSComctlLib.ListView lvCallStack 
      Height          =   1185
      Left            =   9810
      TabIndex        =   3
      Top             =   5850
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   2090
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
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
   Begin VB.TextBox txtOut 
      Height          =   1185
      Left            =   6525
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   5760
      Width           =   3165
   End
   Begin SCIVB_LITE.SciSimple scivb 
      Height          =   5865
      Left            =   90
      TabIndex        =   0
      Top             =   630
      Width           =   13650
      _ExtentX        =   24077
      _ExtentY        =   10345
   End
   Begin MSComctlLib.ImageList ilToolbar 
      Left            =   10305
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   15
      ImageHeight     =   15
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0000
            Key             =   "Run"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":010C
            Key             =   "Start Debugger"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0216
            Key             =   "Break"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0320
            Key             =   "Stop"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":042A
            Key             =   "Toggle Breakpoint"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0534
            Key             =   "Clear All Breakpoints"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":063E
            Key             =   "Step In"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0748
            Key             =   "Step Over"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0852
            Key             =   "Step Out"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":095C
            Key             =   "Run to Cursor"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0A66
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbarDebug 
      Height          =   330
      Left            =   180
      TabIndex        =   1
      Top             =   225
      Width           =   3870
      _ExtentX        =   6826
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   13
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Run"
            Object.ToolTipText     =   "Run"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Start Debugger"
            Object.ToolTipText     =   "Start Debugger"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Break"
            Object.ToolTipText     =   "Break"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Stop"
            Object.ToolTipText     =   "Stop"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Toggle Breakpoint"
            Object.ToolTipText     =   "Toggle Breakpoint"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Clear All Breakpoints"
            Object.ToolTipText     =   "Clear All Breakpoiunts"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Step In"
            Object.ToolTipText     =   "Step In"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Step Over"
            Object.ToolTipText     =   "Step Over"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Step Out"
            Object.ToolTipText     =   "Step Out"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Run to Cursor"
            Object.ToolTipText     =   "Run to Cursor"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilToolbars_Disabled 
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
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0B72
            Key             =   "Run"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0C7E
            Key             =   "Break"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0D8A
            Key             =   "Clear All Breakpoints"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0E96
            Key             =   "Run to Cursor"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0FA2
            Key             =   "Step Over"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":10AE
            Key             =   "Step Out"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":11BA
            Key             =   "Step In"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":12C6
            Key             =   "Stop"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":13D2
            Key             =   "Start Debugger"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":14DC
            Key             =   "Toggle Breakpoint"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TabStrip ts 
      Height          =   3120
      Left            =   180
      TabIndex        =   4
      Top             =   6615
      Width           =   13650
      _ExtentX        =   24077
      _ExtentY        =   5503
      Placement       =   1
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   4
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Output"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Errors"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Variables"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "CallStack"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblStatus 
      Caption         =   "Status: Idle"
      Height          =   375
      Left            =   4185
      TabIndex        =   7
      Top             =   270
      Width           =   4560
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuNewFile 
         Caption         =   "New"
      End
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
   Begin VB.Menu mnuVarsPopup 
      Caption         =   "mnuVarsPopup"
      Begin VB.Menu mnuVarSetValue 
         Caption         =   "Modify Value"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents sciext As CSciExtender
Attribute sciext.VB_VarHelpID = -1

Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long

Private Declare Function run_script Lib "sb_engine" (ByVal lpLibFileName As String, ByVal use_debugger As Long) As Long
Private Declare Sub GetErrorString Lib "sb_engine" (ByVal iErrorCode As Long, ByVal buf As String, ByVal sz As Long)
Private Declare Sub SetCallBacks Lib "sb_engine" (ByVal msgProc As Long, ByVal dbgCmdProc As Long, ByVal hostResolverProc As Long, ByVal lineInputfunc As Long)

Dim loadedFile As String
Dim hsbLib As Long
Public lastEIP As Long
Public hasImports As Boolean

Const SC_MARK_CIRCLE = 0
Const SC_MARK_ARROW = 2
Const SC_MARK_BACKGROUND = 22
'http://www.scintilla.org/aprilw/SciLexer.bas


Dim selCallStackItem As ListItem
Dim selVariable As ListItem

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
        li.SubItems(3) = v.Value
        Set li.Tag = v
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


Private Sub lvCallStack_ItemClick(ByVal Item As MSComctlLib.ListItem)
    scivb.GotoLine CLng(Item.Text)
    Set selCallStackItem = Item
End Sub


Private Sub lvErrors_ItemClick(ByVal Item As MSComctlLib.ListItem)
    On Error Resume Next
    Dim lline As Long
    lline = CLng(Item.Text) - 1
    If lline > 0 Then
        scivb.GotoLine lline
        scivb.SelLength = Len(scivb.GetLineText(lline)) - 2
    End If
End Sub

Private Sub lvVars_DblClick()

    If selVariable Is Nothing Then Exit Sub
    If selVariable.SubItems(2) <> "array" Then Exit Sub
    
    Dim c As Collection
    Dim varName As String
    Dim v As CVariable
    
    On Error Resume Next
    Set v = selVariable.Tag

    varName = selVariable.SubItems(1)
    Set c = EnumArrayVariables(varName)
    If c.count > 0 Then
        frmAryDump.DumpArrayValues varName, c, v.pAryElement
    End If
    
End Sub

Private Sub lvVars_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Set selVariable = Item
End Sub

Private Sub lvVars_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then PopupMenu mnuVarsPopup
End Sub

Private Sub mnuExecuteTillReturn_Click()
    MsgBox "todo: disable all breakpoints,run to line selCallStackItem.text + 1, reenable breakpoints", vbInformation
End Sub

 
Private Sub mnuNewFile_Click()

    If scivb.isDirty Then
            If MsgBox("Editor has changed save contents?", vbInformation + vbYesNo) = vbYes Then
                If Len(loadedFile) = 0 Then
                    loadedFile = dlg.SaveDialog(AllFiles, "default.sb")
                    If Len(loadedFile) = 0 Then Exit Sub
                End If
                scivb.SaveFile loadedFile
            End If
    End If
        
    scivb.Text = Empty
    loadedFile = dlg.OpenDialog(AllFiles)
    If Len(loadedFile) = 0 Then Exit Sub
    scivb.LoadFile loadedFile
    
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

Private Sub mnuVarSetValue_Click()

    If Not running Then Exit Sub
    If selVariable Is Nothing Then Exit Sub
    
    Dim Value As String, newVal As String
    Dim v As CVariable
    
    On Error Resume Next
    
    Set v = selVariable.Tag
    
    If v Is Nothing Then
        MsgBox "Variable tag not set?"
        Exit Sub
    End If
    
    If v.varType = "array" Then
        lvVars_DblClick
        Exit Sub 'unless they want to change the var type here? todo?
    End If
    
    If v.varType = "ref" Then
        MsgBox "Can not edit ref variables edit the parent variable directly..", vbInformation
        Exit Sub
    End If
    
    Value = selVariable.SubItems(3)
    If Left(Value, 1) = """" Then Value = Mid(Value, 2)
    If Right(Value, 1) = """" Then Value = Mid(Value, 1, Len(Value) - 1)
    
    newVal = InputBox("Modify variable " & name, , Value)
    If Len(newVal) > 0 And newVal <> Value Then
        SetVariable v, newVal
        RefreshVariables
    End If
    
End Sub

Private Sub sciext_MarginClick(lline As Long, Position As Long, margin As Long, modifiers As Long)
    'Debug.Print "MarginClick: line,pos,margin,modifiers", lLine, Position, margin, modifiers
    ToggleBreakPoint lline
End Sub

Private Sub sciext_MouseDwellEnd(lline As Long, Position As Long)
   If running Then tmrHideCallTip.Enabled = True
End Sub

Private Sub sciext_MouseDwellStart(lline As Long, Position As Long)
    'Debug.Print "MouseDwell: " & lLine & " CurWord: " & sciext.WordUnderMouse(Position)
    
    Dim li As ListItem
    Dim curWord As String
    
    If running Then
         curWord = sciext.WordUnderMouse(Position)
         For Each li In lvVars.ListItems
            If LCase(li.SubItems(1)) = LCase(curWord) Then 'they have moused over a variable..
                Set selVariable = li
                scivb.SelStart = Position 'so call tip shows right under it..
                scivb.SelLength = 0
                scivb.ShowCallTip curWord & " = " & li.SubItems(3)
                Exit For
            End If
         Next
    End If
    
        
End Sub

Private Sub scivb_AutoCompleteEvent(className As String)
    'Debug.Print className
    Dim matches() As String
    Dim prevWord As String
    Dim orgPos As Long
    Dim curpos As Long
    
    'first lets see if this is an import/include statement
    prevWord = LCase(sciext.WordUnderMouse(scivb.SelStart - Len(className) - 1, True))
    If prevWord = "include" Or prevWord = "import" Then
        matches() = GetAutoCompleteStringForIncludes(className)
        If Not AryIsEmpty(matches) Then
            If UBound(matches) = 0 Then
                'only one match so just auto complete it..
                'scivb.SelStart = scivb.SelStart - Len(className)
                'scivb.SelLength = Len(className)
                scivb.SelStart = scivb.DirectSCI.WordStartPosition(scivb.SelStart, True)
                scivb.SelLength = scivb.DirectSCI.WordEndPosition(scivb.SelStart, True) - scivb.SelStart
                scivb.SelText = matches(0)
            Else
                'show all matches for partial string
                scivb.ShowAutoComplete Join(matches, " ")
            End If
        Else
               'show all include files
                scivb.ShowAutoComplete Join(IncludeFiles, " ")
        End If
        Exit Sub
    End If
    
    
    'now lets see if it scoped to a specific module
    curpos = scivb.DirectSCI.GetCurPos()
    curpos = curpos - Len(className) - 2
    If curpos > 4 Then
        'this check wont trigger for nt::Msg[ctrl+space], only nt::
        If Mid(scivb.Text, curpos + 1, 2) = "::" Then 'its an module lookup
            orgPos = scivb.DirectSCI.GetCurPos()
            scivb.SetCurrentPosition curpos
            prevWord = scivb.CurrentWord
            scivb.SetCurrentPosition orgPos
            If ShowAutoCompleteForModule(prevWord, className) Then Exit Sub
        End If
    End If
        
    'now search the built in api for partial matches to whats already typed..
    matches() = GetAutoCompleteString(className)
    If Not AryIsEmpty(matches) Then
        If UBound(matches) = 0 Then
            'only one match so just auto complete it..
            'scivb.SelStart = scivb.SelStart - Len(className)
            'scivb.SelLength = Len(className)
            scivb.SelStart = scivb.DirectSCI.WordStartPosition(scivb.SelStart, True)
            scivb.SelLength = scivb.DirectSCI.WordEndPosition(scivb.SelStart, True) - scivb.SelStart
            scivb.SelText = matches(0)
        Else
            'show all matches for partial string
            scivb.ShowAutoComplete Join(matches, " ")
        End If
    Else
        'ok no partial matches, lets just show entire api list..
        If Not AryIsEmpty(FunctionPrototypes) Then
            scivb.ShowAutoComplete Join(FunctionPrototypes, " ")
        End If
    End If
    
End Sub

Private Function ShowAutoCompleteForModule(modName As String, fragment As String) As Boolean
    On Error Resume Next
    Dim methods As String
    Dim matches() As String
     
    If Len(modName) = 0 Then Exit Function
    If Not isIncludeFile(modName) Then Exit Function
    If Not isFileIncluded(modName, scivb.Text) Then Exit Function
    
    methods = modules(modName)
    If Err.Number <> 0 Then Exit Function
    
    If Len(fragment) = 0 Then
        scivb.ShowAutoComplete methods 'ctrl-space after [module]::
        ShowAutoCompleteForModule = True
    Else
        matches() = GetAutoCompleteStringForModule(methods, fragment) 'example nt::msg[ctrl-space]
        If Not AryIsEmpty(matches) Then
            If UBound(matches) = 0 Then
                'scivb.SelStart = scivb.SelStart - Len(modName) - 1
                'scivb.SelLength = Len(modName) + 1
                scivb.SelStart = scivb.DirectSCI.WordStartPosition(scivb.SelStart, True)
                scivb.SelLength = scivb.DirectSCI.WordEndPosition(scivb.SelStart, True) - scivb.SelStart
                scivb.SelText = matches(0)
            Else
                scivb.ShowAutoComplete ":" & Join(matches, ":")
            End If
            ShowAutoCompleteForModule = True
        End If
    End If
        
End Function

Private Sub scivb_CallTipClick(Position As Long)
    If running Then mnuVarSetValue_Click
End Sub

Private Sub scivb_DoubleClick()
    Dim word As String
    word = scivb.CurrentWord
    If Len(word) < 20 Then
        Me.Caption = "  " & scivb.hilightWord(word, , vbTextCompare) & " instances of '" & word & " ' found"
    End If
End Sub

Private Sub scivb_MouseUp(Button As Integer, Shift As Integer, x As Long, y As Long)
    If scivb.SelLength > 0 And scivb.SelLength < 20 Then
        Dim word As String
        word = Trim(scivb.SelText)
        word = Replace(word, vbTab, "")
        If Len(word) < 20 Then
            Me.Caption = "  " & scivb.hilightWord(word, , vbTextCompare) & " instances of '" & word & " ' found"
        End If
    Else
        scivb.hilightClear
    End If
End Sub

Private Sub scivb_KeyDown(KeyCode As Long, Shift As Long)

    'Debug.Print KeyCode & " " & Shift
    Select Case KeyCode
        Case vbKeyF2: ToggleBreakPoint
        Case vbKeyF5: If running Then DebuggerCmd dc_Run Else ExecuteScript True
        Case vbKeyF7: DebuggerCmd dc_stepinto
        Case vbKeyF8: DebuggerCmd dc_StepOver
        Case vbKeyF9: DebuggerCmd dc_StepOut
    End Select

End Sub

Private Sub scivb_KeyUp(KeyCode As Long, Shift As Long)

    Dim curWord As String
    Dim txt As String
    Dim curpos As Long
    Dim prevChar As String
    Dim methods As String
    
    If KeyCode = 186 Then ': character
        curpos = scivb.GetCaretInLine()
        txt = scivb.GetLineText(scivb.CurrentLine)

        If curpos < 3 Then Exit Sub
        prevChar = Mid(txt, curpos - 1, 1)
        If prevChar <> ":" Then Exit Sub

        scivb.GotoCol curpos - 2
        curWord = scivb.CurrentWord
        scivb.GotoCol curpos

        On Error Resume Next
        If Len(curWord) > 0 Then
            If isIncludeFile(curWord) And isFileIncluded(curWord, scivb.Text) Then
                methods = modules(curWord)
                If Err.Number = 0 Then scivb.ShowAutoComplete methods
            End If
        End If

    End If


End Sub


Private Sub tbarDebug_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.key
        Case "Run":               If running Then DebuggerCmd dc_Run Else ExecuteScript
        Case "Start Debugger":    If running Then DebuggerCmd dc_Run Else ExecuteScript True
        Case "Stop":              DebuggerCmd dc_Quit
        Case "Step In":           DebuggerCmd dc_stepinto
        Case "Step Over":         DebuggerCmd dc_StepOver
        Case "Step Out":          DebuggerCmd dc_StepOut
        Case "Run to Cursor":     RunToLine scivb.CurrentLine + 1
        Case "Toggle Breakpoint": ToggleBreakPoint
        Case "Clear All Breakpoints": RemoveAllBreakpoints
    End Select
    
End Sub

Private Sub CheckError()
    On Error Resume Next
    Dim lline As Long
    
    If lvErrors.ListItems.count = 0 Then Exit Sub
    ts.Tabs(2).Selected = True
    
    lline = CLng(lvErrors.ListItems(1).Text) - 1
    If lline <> 0 Then
        scivb.GotoLine lline
        scivb.SelLength = Len(scivb.GetLineText(lline)) - 2
    End If
           
 
End Sub
Private Sub ExecuteScript(Optional withDebugger As Boolean)

    Dim rv As Long
    Dim buf As String
    
    If Len(Trim(scivb.Text)) = 0 Then Exit Sub
    
    If Len(loadedFile) = 0 Then
        loadedFile = dlg.SaveDialog(AllFiles, "default.sb")
        If Len(loadedFile) = 0 Then Exit Sub
    End If
    
    txtOut.Text = Empty
    lvErrors.ListItems.Clear
    ts.Tabs(1).Selected = True
    
    sciext.LockEditor
    If scivb.isDirty Then scivb.SaveFile loadedFile
    
    running = True
    SetToolBarIcons
    lblStatus = "Status: " & IIf(withDebugger, "Debugging...", "Running...")
    
    rv = run_script(loadedFile, IIf(withDebugger, 1, 0))
     
    'if user closed form while debugger running..we must exit now or form_load again hidden..
    If shuttingDown Then Exit Sub
    
    CheckError
    
    lblStatus = "Status: Idle"
    running = False
    SetToolBarIcons
    
    ClearUIBreakpoints
    sciext.LockEditor False
    scivb.DeleteMarker lastEIP, 1
    lvVars.ListItems.Clear
    lvCallStack.ListItems.Clear
    
    Set selVariable = Nothing
    Set selCallStackItem = Nothing
    
    If hasImports Then scivb.LoadFile loadedFile
    
End Sub

Private Sub SetToolBarIcons()
    Dim b As Button
    
    Set tbarDebug.ImageList = Nothing
    Set tbarDebug.ImageList = IIf(running, ilToolbar, ilToolbars_Disabled)
    
    For Each b In tbarDebug.Buttons
        If Len(b.key) > 0 Then
            b.Image = b.key
            b.ToolTipText = b.key
            If b.key <> "Run" And b.key <> "Start Debugger" Then
                b.Enabled = running
            End If
        End If
    Next
    
End Sub

Private Sub Form_Load()

    SetToolBarIcons
    FormPos Me, True
    
    lvVars.Visible = False
    lvCallStack.Visible = False
    lvErrors.Visible = False
    
    mnuCallStackPopup.Visible = False
    mnuVarsPopup.Visible = False
    
    hsbLib = LoadLibrary(App.path & "\engine\sb_engine.dll")

    If hsbLib = 0 Then
        hsbLib = LoadLibrary(App.path & "\sb_engine.dll")
        If hsbLib = 0 Then
            MsgBox "Failed to load sb_engine.dll by explicit path?", vbInformation
        End If
    End If

    includeDir = GetMySetting("includeDir", App.path & "\include\")
    moduleDir = GetMySetting("moduleDir", App.path & "\modules\")
    InitIntellisense includeDir
    SetDefaultDirs includeDir, moduleDir

    SetCallBacks AddressOf vb_stdout, AddressOf GetDebuggerCommand, AddressOf HostResolver, AddressOf VbLineInput
    
    LoadFunctionPrototypes App.path & "\dependancies\calltips.txt"
    scivb.LoadHighlighter App.path & "\dependancies\vb.bin"
    scivb.DirectSCI.HideSelection False
    scivb.DirectSCI.MarkerDefine 2, SC_MARK_CIRCLE
    scivb.DirectSCI.MarkerSetFore 2, vbRed 'set breakpoint color
    scivb.DirectSCI.MarkerSetBack 2, vbRed

    scivb.DirectSCI.MarkerDefine 1, SC_MARK_ARROW
    scivb.DirectSCI.MarkerSetFore 1, vbBlack 'current eip
    scivb.DirectSCI.MarkerSetBack 1, vbYellow

    scivb.DirectSCI.MarkerDefine 3, SC_MARK_BACKGROUND
    scivb.DirectSCI.MarkerSetFore 3, vbBlack 'current eip
    scivb.DirectSCI.MarkerSetBack 3, vbYellow

    If GetMySetting("firstrun", 1) = 1 Then
        LoadFile App.path & "\scripts\com_voice_test.sb"
        SaveMySetting "firstrun", 0
    End If
    
    'LoadFile App.path & "\scripts\functions.txt"

    'AddObject "frmMain", Me
    'AddString "test", "this is my string from vb!"
    'LoadFile App.path & "\scripts\GetHostObject.sb"

    scivb.DirectSCI.AutoCSetIgnoreCase True
    scivb.DisplayCallTips = True
    Call scivb.LoadCallTips(App.path & "\dependancies\calltips.txt")
    'LoadFile App.path & "\scripts\importNT.sb"
    scivb.ReadOnly = False

    Set sciext = New CSciExtender
    sciext.Init scivb

End Sub

Sub LoadFile(fpath As String)
   
   loadedFile = fpath
   scivb.DeleteAllMarkers
   scivb.LoadFile loadedFile
   Set breakpoints = New Collection
   lvErrors.ListItems.Clear
   txtOut.Text = Empty
   
End Sub


Private Sub Form_Resize()
    On Error Resume Next
    With scivb
        .Width = Me.Width - .Left - 200
        ts.Width = .Width
        txtOut.Width = .Width - 200
        ts.Top = Me.Height - ts.Height - 800
        .Height = Me.Height - .Top - ts.Height - 1000
        With txtOut
            .Move ts.Left + 100, ts.Top + 150, ts.Width - 200, ts.Height - 500
            lvVars.Move .Left, .Top, .Width, .Height
            lvCallStack.Move .Left, .Top, .Width, .Height
            lvErrors.Move .Left, .Top, .Width, .Height
        End With
        SetLastColumnWidth lvCallStack
        SetLastColumnWidth lvVars
        SetLastColumnWidth lvErrors
    End With
End Sub

Private Sub SetLastColumnWidth(lv As ListView)
    lv.ColumnHeaders(lv.ColumnHeaders.count).Width = lv.Width - lv.ColumnHeaders(lv.ColumnHeaders.count).Left - 100
End Sub

Private Sub Form_Unload(Cancel As Integer)
    shuttingDown = True
    If running Then DebuggerCmd dc_Quit
    Call SaveMySetting("includeDir", includeDir)
    Call SaveMySetting("moduleDir", moduleDir)
    FormPos Me, True, True
    SetCallBacks 0, 0, 0, 0
    FreeLibrary hsbLib
End Sub

Public Sub SyncUI()
    
    Dim curline As Long
    
    curline = GetCurrentDebugLine(hDebugObject)
    scivb.SetMarker curline, 1
    scivb.SetMarker curline, 3
    lastEIP = curline
    
    scivb.GotoLine curline
    scivb.SetFocus
    
    RefreshVariables
    RefreshCallStack
    
End Sub

Public Function Alert(msg As String)
    MsgBox msg
End Function

'we use a timer for this to give them a chance to click on the calltip to edit the variable..
Private Sub tmrHideCallTip_Timer()
    If sciext.isMouseOverCallTip() Then Exit Sub
    tmrHideCallTip.Enabled = False
    scivb.StopCallTip
    Set selVariable = Nothing
End Sub

Private Sub ts_Click()
    Dim i As Long
    i = ts.SelectedItem.Index
    txtOut.Visible = IIf(i = 1, True, False)
    lvErrors.Visible = IIf(i = 2, True, False)
    lvVars.Visible = IIf(i = 3, True, False)
    lvCallStack.Visible = IIf(i = 4, True, False)
End Sub
















