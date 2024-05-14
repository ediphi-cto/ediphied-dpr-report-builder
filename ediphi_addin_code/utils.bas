Attribute VB_Name = "utils"
Option Explicit

Function userXLStartPath() As String
    
    'any excel file in this path will auto open when excel opens
    Dim appDataPath As String
    appDataPath = Environ("APPDATA")
    userXLStartPath = appDataPath & "\Microsoft\Excel\XLSTART\"
    
End Function

Function timestamper()
    
    timestamper = Format(Now, "YYYYMMDD_HHMMSS")

End Function

Public Function downloadsPath() As String
    
    Dim userProfilePath As String
    userProfilePath = Environ("USERPROFILE")
    If Len(userProfilePath) > 0 Then
        downloadsPath = userProfilePath & "\Downloads\"
    Else
        Err.Raise 404, "downoadsPath()", "download path not found"
    End If

End Function

Function getEnv(varName As String) As String
    
    On Error GoTo e1
    getEnv = ThisWorkbook.Worksheets("env").Range(varName).Value

Exit Function
e1:
    'returns blank string

End Function

Sub setEnv(varName As String, val, Optional temporary As Boolean)
    
    On Error GoTo e1
    Dim str As String
    str = CStr(val)
    ThisWorkbook.Worksheets("env").Range(varName).Value = str
    If Not temporary Then ThisWorkbook.Save

Exit Sub
e1:
    logError "failed to set env var """ & varName & """ to """ & str & """"

End Sub

Sub logError(msg As String)

    Debug.Print "===ediphi=== ERROR: " & msg

End Sub

Public Sub hideMe()

    Dim win As Window
    For Each win In ThisWorkbook.Windows
        win.Visible = False
    Next

End Sub
