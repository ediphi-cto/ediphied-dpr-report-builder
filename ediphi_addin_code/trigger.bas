Attribute VB_Name = "trigger"
Option Explicit
Public thisAddin As EdiphiAddin

Sub triggerReportBuilder(ediphiCSVfilename As String)

    Dim reportBuilderWB As Workbook
    On Error GoTo e1
    
    Dim checkIfOpen As Workbook
    On Error Resume Next
    Set checkIfOpen = Workbooks(REPORT_BUILDER_FILENAME)
    If Not checkIfOpen Is Nothing Then
        MsgBox "Only one DPR Report can be open at a time." & vbLf & vbLf & _
        "Close down your current DPR Report Builder and try again.", vbExclamation
        Workbooks(ediphiCSVfilename).Close SaveChanges:=False
        Exit Sub
    End If
    On Error GoTo e1
    
    On Error Resume Next
    Set reportBuilderWB = Workbooks.Open(reportBuilderFullname)
    On Error GoTo e1
    
    If reportBuilderWB Is Nothing Then
        Set reportBuilderWB = fetchReportBuilder(forceDownload:=True)
    Else
        If getEnv("AUTO_UPDATE") <> 0 And updateNeeded Then
            Set reportBuilderWB = fetchReportBuilder(wbIfUpdateDenied:=reportBuilderWB)
        End If
    End If

    If reportBuilderWB Is Nothing Then GoTo e1
    'this is the hand off, where the addin tells the report builder where the data is
    reportBuilderWB.Worksheets("trigger").[a1].Value = ediphiCSVfilename
    
Exit Sub
e1:
    logError "ReportBuilder failed to open"

Exit Sub
e2:
    Set reportBuilderWB = fetchReportBuilder()
    Resume Next
    
End Sub
