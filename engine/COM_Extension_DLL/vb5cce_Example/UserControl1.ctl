VERSION 5.00
Begin VB.UserControl CCESample 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
End
Attribute VB_Name = "CCESample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public CallBackHandlers As New Collection

Public Function LaunchCallBackForm(ByVal defVal As Long)
    LaunchCallBackForm = frmCallBack.ShowCallBackForm(defVal, Me)
End Function
