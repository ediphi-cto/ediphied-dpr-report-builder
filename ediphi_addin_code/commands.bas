Attribute VB_Name = "commands"
Option Explicit

Public Sub ediphiUpdate()
    
    On Error GoTo e1
    Dim WB As Workbook
    Set WB = fetchReportBuilder
    WB.Saved = True 'ensures a silent close, doesnt save, but tells Excel that it is
    WB.Close
    MsgBox "Ediphi Report Builder SUCCESSFULLY UPDATED!", vbInformation

Exit Sub
e1:
    logError "ReportBuilder failed to update"

End Sub

Public Sub ediphiAutoUpdatesOn()

    setEnv "AUTO_UPDATE", 1
    MsgBox "Ediphi Report Builder AUTO UPDATING is now ON", vbInformation

End Sub

Public Sub ediphiAutoUpdatesOff()

    setEnv "AUTO_UPDATE", 0
    MsgBox "Ediphi Report Builder AUTO UPDATING is now OFF", vbInformation

End Sub

Public Sub ediphiSetApiKey()
    
    Dim apiKey As String
    apiKey = InputBox("Enter the API Key", "Ediphi Security")
    setEnv "API_KEY", apiKey
    MsgBox "ediphi API key has been SET", vbInformation

End Sub

Public Sub ediphiDebug()

    setEnv "DEBUG", 1, temporary:=True
    MsgBox "DEBUG mode is now ON", vbExclamation
    
End Sub


