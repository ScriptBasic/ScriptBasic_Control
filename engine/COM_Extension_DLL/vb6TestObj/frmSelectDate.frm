VERSION 5.00
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Begin VB.Form frmSelectDate 
   Caption         =   "Select Date"
   ClientHeight    =   3405
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4725
   LinkTopic       =   "Form2"
   ScaleHeight     =   3405
   ScaleWidth      =   4725
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   330
      Left            =   2025
      TabIndex        =   2
      Top             =   2970
      Width           =   1050
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Height          =   330
      Left            =   3285
      TabIndex        =   1
      Top             =   2970
      Width           =   1050
   End
   Begin MSACAL.Calendar Calendar1 
      Height          =   2625
      Left            =   225
      TabIndex        =   0
      Top             =   180
      Width           =   4155
      _Version        =   524288
      _ExtentX        =   7329
      _ExtentY        =   4630
      _StockProps     =   1
      BackColor       =   -2147483633
      Year            =   2014
      Month           =   6
      Day             =   16
      DayLength       =   1
      MonthLength     =   2
      DayFontColor    =   0
      FirstDay        =   1
      GridCellEffect  =   1
      GridFontColor   =   10485760
      GridLinesColor  =   -2147483632
      ShowDateSelectors=   -1  'True
      ShowDays        =   -1  'True
      ShowHorizontalGrid=   -1  'True
      ShowTitle       =   -1  'True
      ShowVerticalGrid=   -1  'True
      TitleFontColor  =   10485760
      ValueIsNull     =   0   'False
      BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmSelectDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cancel As Boolean 'default is false

Private Sub cmdCancel_Click()
    cancel = True
    Me.Visible = False
End Sub

Private Sub cmdOk_Click()
    
    If Calendar1.Day = 0 Then
        MsgBox "Please select a day", vbInformation
        Exit Sub
    End If
         
    Me.Visible = False 'this breaks the me.show modal lock
    
End Sub

'defaults to empty string
Public Function SelectDate() As String
    
    Me.Show 1 'this will block until the form is unloaded or hidden..
              'form events such as button clicks and controls will
              'continue to process normally...
    
    If Not cancel Then
        With Calendar1
            SelectDate = .Day & "/" & .Month & "/" & .Year
        End With
    End If
    
    Unload Me
    
End Function
