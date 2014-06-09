VERSION 5.00
Object = "{FBE17B58-A1F0-4B91-BDBD-C9AB263AC8B0}#78.0#0"; "scivb_lite.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   9855
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13905
   LinkTopic       =   "Form1"
   ScaleHeight     =   9855
   ScaleWidth      =   13905
   StartUpPosition =   2  'CenterScreen
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
      Height          =   1455
      Left            =   5670
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   11
      Top             =   8055
      Width           =   8205
   End
   Begin VB.CommandButton cmdManual 
      Caption         =   "Command1"
      Height          =   375
      Left            =   12825
      TabIndex        =   10
      Top             =   6660
      Width           =   960
   End
   Begin VB.TextBox txtCmd 
      Height          =   285
      Left            =   10620
      TabIndex        =   9
      Top             =   6705
      Width           =   1995
   End
   Begin VB.ListBox List1 
      Height          =   1035
      Left            =   5625
      TabIndex        =   5
      Top             =   6750
      Width           =   4065
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
      Height          =   2715
      Left            =   135
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   4
      Top             =   6750
      Width           =   5190
   End
   Begin VB.CheckBox chkDebug 
      Caption         =   "use debugger"
      Height          =   285
      Left            =   9360
      TabIndex        =   2
      Top             =   225
      Width           =   1455
   End
   Begin VB.CommandButton cmdRun 
      Caption         =   "Run Script"
      Height          =   330
      Left            =   7965
      TabIndex        =   1
      Top             =   135
      Width           =   1185
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
         NumListImages   =   20
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0000
            Key             =   "Toggle Bookmark"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":010A
            Key             =   "Next Bookmark"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0214
            Key             =   "Previous Bookmark"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":031E
            Key             =   "Clear All Bookmarks"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0428
            Key             =   "Comment Block"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0532
            Key             =   "UnCommentBlock"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":063C
            Key             =   "Tab Left"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0746
            Key             =   "Tab Right"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0850
            Key             =   "Parse code"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":095A
            Key             =   "Run"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0A64
            Key             =   "Break"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0B6E
            Key             =   "Stop"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0C78
            Key             =   "Toggle Breakpoint"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0D82
            Key             =   "Clear All Breakpoints"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0E8C
            Key             =   "Step In"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0F96
            Key             =   "Step Over"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":10A0
            Key             =   "Step Out"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":11AA
            Key             =   "Immediate"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":12B4
            Key             =   "Callstack"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":13BE
            Key             =   "Set Next"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbarDebug 
      Height          =   315
      Left            =   180
      TabIndex        =   6
      Top             =   225
      Width           =   9540
      _ExtentX        =   16828
      _ExtentY        =   556
      ButtonWidth     =   582
      ButtonHeight    =   556
      Style           =   1
      ImageList       =   "ilToolbars"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   14
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Run"
            Object.ToolTipText     =   "Run"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Break"
            Object.ToolTipText     =   "Break"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Stop"
            Object.ToolTipText     =   "Stop"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Breakpoint"
            Object.ToolTipText     =   "Toggle Breakpoint"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Clear Breakpoints"
            Object.ToolTipText     =   "Clear All Breakpoiunts"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Step In"
            Object.ToolTipText     =   "Step In"
            ImageIndex      =   15
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Step Over"
            Object.ToolTipText     =   "Step Over"
            ImageIndex      =   16
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Step Out"
            Object.ToolTipText     =   "Step Out"
            ImageIndex      =   17
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Set Next"
            Object.ToolTipText     =   "Set Next Statement"
            ImageIndex      =   20
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Immediate"
            Object.ToolTipText     =   "Immediate Window"
            ImageIndex      =   18
            Style           =   1
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Callstack"
            Object.ToolTipText     =   "Callstack"
            ImageIndex      =   19
            Style           =   1
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar sbStatus 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   9600
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
   Begin VB.Label Label2 
      Caption         =   "Cmd"
      Height          =   285
      Left            =   9900
      TabIndex        =   8
      Top             =   6750
      Width           =   555
   End
   Begin VB.Label Label1 
      Caption         =   "Output"
      Height          =   240
      Left            =   180
      TabIndex        =   3
      Top             =   6525
      Width           =   1860
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long


Private Declare Function run_script Lib "sb_engine" (ByVal lpLibFileName As String, ByVal use_debugger As Long) As Long
Private Declare Sub GetErrorString Lib "sb_engine" (ByVal iErrorCode As Long, ByVal buf As String, ByVal sz As Long)
Private Declare Sub SetCallBacks Lib "sb_engine" (ByVal msgProc As Long, ByVal dbgCmdProc As Long)


Dim loadedFile As String
Dim hsbLib As Long
Public running As Boolean



Private Sub cmdManual_Click()
    txtDebug.Text = Empty
    dbg_cmd = txtCmd.Text
    readyToReturn = True
End Sub

'h help
's step one line, or just press return on the line
'S step one line, do not step into functions or subs
'o step until getting out of the current function (if you stepped into but changed your mind)
'? var  print the value of a variable
'u step one level up in the stack
'd step one level down in the stack (for variable printing)
'D step down in the stack to current execution depth
'G list all global variables
'L list all local variables
'l [n-m] list the source lines
'r [n] run to line n
'R [n] run to line n but do not stop in recursive function call
'b [n] set breakpoint on the line n or the current line
'B [n-m] remove breakpoints from lines
'q quit the program

Private Sub tbarDebug_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.Key
        Case "Run":
        Case "Stop":
        Case "Step In":
        Case "Step Over":
        Case "Step Out":
        Case "Set Next":
        Case "Immediate"
        Case "Callstack"
        Case "Breakpoint":
        Case "Clear Breakpoints":
    End Select
    
End Sub


Private Sub cmdRun_Click()
    Dim rv As Long
    Dim buf As String
     
    txtOut.Text = Empty
    List1.Clear
    
    scivb.ReadOnly = True
    If scivb.isDirty Then scivb.SaveFile loadedFile
    
    running = True
    Form1.sbStatus.Panels(1).Text = "Running"
    
    rv = run_script(loadedFile, chkDebug.Value)
    
    If rv <> 0 Then
        buf = String(255, " ")
        Call GetErrorString(rv, buf, 255)
        sbStatus.Panels(1).Text = "Error: " & buf
    End If
    
    running = False
    scivb.ReadOnly = False
    Form1.sbStatus.Panels(1).Text = "Idle"
    
End Sub

Private Sub Form_Load()
    
    hsbLib = LoadLibrary(App.Path & "\engine\sb_engine.dll")
    
    If hsbLib = 0 Then
        MsgBox "Failed to load sb_engine.dll by explicit path?"
    End If
    
    SetCallBacks AddressOf vb_stdout, AddressOf GetDebuggerCommand
    
    loadedFile = App.Path & "\scripts\com_voice_test.sb"
    scivb.LoadHighlighter App.Path & "\dependancies\vb.bin"
    scivb.LoadFile loadedFile
    scivb.DirectSCI.HideSelection False
    
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    With scivb
        .Width = Me.Width - .Left - 200
        '.Height = Me.Height - .Top - 500
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    FreeLibrary hsbLib
End Sub

Public Sub SyncUI()
    
    Dim curLine As Long
    
    curLine = GetCurrentDebugLine(hDebugObject)
    scivb.GotoLine curLine
    scivb.HighLightActiveLine = True
    scivb.SetFocus
    
End Sub
