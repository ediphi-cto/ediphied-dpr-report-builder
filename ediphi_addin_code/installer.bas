Attribute VB_Name = "installer"
Option Explicit

Public Const REPORT_BUILDER_FILENAME As String = "DPR_reportBuilder.xlsm"
Public Const REPORT_BUILDER_DIR As String = "ediphi_cache"
Public Const MY_FILENAME As String = "ediphi_addin.xlam"

Sub installMe()
    
    'ask for an api key if thisworkbook does not have one
    Dim apiKey As String
    If getEnv("API_KEY") = "" Then
        apiKey = InputBox("Please enter an API key: ", "ediphi installer")
        If Len(apiKey) < 1 Then
            MsgBox "Invalid API key provided.  Installion is exiting before complete.", vbCritical
            Exit Sub
        Else
            'persist the key in thisworkbook
            setEnv "API_KEY", apiKey
        End If
    End If
    
    Dim fullFileName As String
    fullFileName = userXLStartPath & MY_FILENAME
    With ThisWorkbook
        Application.DisplayAlerts = False
        hideMe
        .SaveAs filename:=fullFileName, FileFormat:=xlOpenXMLAddIn
        Application.DisplayAlerts = True
        MsgBox "SUCCESS!!!" & vbLf & vbLf & "The Ediphi / DPR report builder is now installed"
    End With
    
End Sub

Function fetchReportBuilder() As Workbook
    
    'go get xlsm from S3, overwrite in user's excel start, this is a CICD hack
    Dim req As New MSXML2.XMLHTTP60
    Dim url As String
    Dim stream As New ADODB.stream
    
    On Error GoTo e1
    url = getEnv("S3_URL")
    
    With req
        .Open "GET", url, False
        .setRequestHeader "Cache-Control", "no-cache, no-store, must-revalidate"
        .setRequestHeader "Pragma", "no-cache"
        .setRequestHeader "Expires", "0"
        .send
        While req.readyState <> 4
            DoEvents
        Wend
        If .Status = "200" Then
                On Error Resume Next
                Workbooks(REPORT_BUILDER_FILENAME).Close
                On Error GoTo e1
                If Dir(reportBuilderFullname) <> "" Then
                    SetAttr reportBuilderFullname, vbNormal
                End If
                stream.Open
                stream.Type = 1
                stream.Write .responseBody
                stream.SaveToFile reportBuilderFullname, 2
                SetAttr reportBuilderFullname, vbReadOnly
                stream.Close
                Set stream = Nothing
                Set fetchReportBuilder = Workbooks.Open(reportBuilderFullname)
                GoTo finally
        Else
            GoTo e1
        End If
    End With
    
finally:
Set stream = Nothing
    
Exit Function
e1:
    'response can be nothing, failure will return nothing
    Resume finally

End Function

Function reportBuilderFullname() As String
    
    Dim dirPath As String, fullFileName As String
    dirPath = ThisWorkbook.Path & "\" & REPORT_BUILDER_DIR
    If Dir(dirPath, vbDirectory) = "" Then
        MkDir dirPath
    End If
    reportBuilderFullname = dirPath & "\" & REPORT_BUILDER_FILENAME

End Function

Function updateNeeded() As Boolean

    updateNeeded = getS3LastModifiedDate > getFileDateModified(reportBuilderFullname)
    
End Function

Function getS3LastModifiedDate() As Date
    On Error GoTo e1
    Dim req As New MSXML2.ServerXMLHTTP60
    Dim url As String
    Dim lastModified As String
    Dim arr
    
    url = getEnv("S3_URL")

    With req
        .Open "HEAD", url, False
        .send
        lastModified = .getResponseHeader("Last-Modified")
    End With
    Set req = Nothing
    arr = Split(lastModified, " ")
    getS3LastModifiedDate = CDate(arr(2) & " " & arr(1) & " " & arr(3))
    
Exit Function
e1:
    getS3LastModifiedDate = CDate("1/2/2000") 'so that an update won't process if this fails

End Function

Function getFileDateModified(filePath As String) As Date
    On Error GoTo e1
    Dim fso As FileSystemObject
    Dim file As file
    Set fso = New FileSystemObject
    
    If fso.FileExists(filePath) Then
        Set file = fso.GetFile(filePath)
        getFileDateModified = file.DateLastModified
    Else
        GoTo e1
    End If
    
    Set file = Nothing
    Set fso = Nothing
    
Exit Function
e1:
    getFileDateModified = CDate("1/1/2000") 'make it very old so if failure, will prompt an udpate
End Function

