Attribute VB_Name = "modGeneral"



Function GetMySetting(key, def)
    GetMySetting = GetSetting(App.EXEName, "Settings", key, def)
End Function

Sub SaveMySetting(key, value)
    SaveSetting App.EXEName, "Settings", key, value
End Sub
