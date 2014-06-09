VERSION 5.00
Object = "{FBE17B58-A1F0-4B91-BDBD-C9AB263AC8B0}#78.0#0"; "scivb_lite.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   9825
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14100
   LinkTopic       =   "Form1"
   ScaleHeight     =   9825
   ScaleWidth      =   14100
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   2595
      Left            =   6570
      TabIndex        =   5
      Top             =   6975
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
      Top             =   6975
      Width           =   5190
   End
   Begin VB.CheckBox chkDebug 
      Caption         =   "use debugger"
      Height          =   285
      Left            =   9315
      TabIndex        =   2
      Top             =   135
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
   Begin VB.Label Label1 
      Caption         =   "Output"
      Height          =   240
      Left            =   90
      TabIndex        =   3
      Top             =   6615
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


Private Sub cmdRun_Click()
    Dim rv As Long
    Dim buf As String
     
    txtOut.Text = Empty
    SetCallBacks AddressOf vb_stdout, AddressOf GetDebuggerCommand
    rv = run_script(loadedFile, chkDebug.Value)
    
    Me.Caption = "run_script() = " & rv & "        -    " & Format(Now, "m:s:ss AM/PM")
    
    If rv <> 0 Then
        buf = String(255, " ")
        Call GetErrorString(rv, buf, 255)
        MsgBox buf
    End If
    
End Sub

Private Sub Form_Load()
    
    hsbLib = LoadLibrary(App.Path & "\engine\sb_engine.dll")
    
    If hsbLib = 0 Then
        MsgBox "Failed to load sb_engine.dll by explicit path?"
    End If
    
    loadedFile = App.Path & "\scripts\com_voice_test.sb"
    scivb.LoadHighlighter App.Path & "\dependancies\vb.bin"
    scivb.LoadFile loadedFile
    
    
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
