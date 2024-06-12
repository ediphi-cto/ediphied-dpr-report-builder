Attribute VB_Name = "ediphiUtils"
Option Explicit
Public Const EDIPHI_ADDIN_FILENAME As String = "ediphi_addin.xlam"
Public thisReportBuilder As EdiphiReportBuilder
Public errors As Collection
Public Const TRY_UPDATING_MSG As String = "Try updating it by pressing ALT + F8, then typing 'ediphiUpdate', and hit Enter"

Sub eventsOn()

    With Application
        '.Calculation = xlCalculationAutomatic
        .EnableEvents = True
        .ScreenUpdating = True
    End With

End Sub

Sub eventsOff()

    With Application
        '.Calculation = xlCalculationManual
        .EnableEvents = False
        .ScreenUpdating = False
    End With

End Sub

Sub tryApi()

    Dim api As New EdiphiAPI
    Dim rawJson As String
    rawJson = api.getEstimateJSON("5347cddc-d3dd-46dd-b2de-d81f774ac3cc")
    Dim jsonDict As Dictionary
    Set jsonDict = ParseJson(rawJson)
    
    pp jsonDict
    
End Sub

Function GetUniqueSheetName(baseName As String) As String
    Dim ws As Worksheet
    Dim uniqueName As String
    Dim suffix As Integer
    Dim nameExists As Boolean
    
    ' Start with the base name
    uniqueName = baseName
    suffix = 1
    
    ' Loop to check if the name already exists
    Do
        nameExists = False
        For Each ws In ThisWorkbook.Sheets
            If ws.name = uniqueName Then
                ' If the name exists, generate a new name with a suffix
                uniqueName = baseName & " " & suffix
                suffix = suffix + 1
                nameExists = True
                Exit For
            End If
        Next ws
    Loop While nameExists
    
    ' Return the unique name
    GetUniqueSheetName = uniqueName
End Function


Function getSortFieldColl() As Collection

    Set getSortFieldColl = New Collection
    
    Dim pivotDataWS As Worksheet
    Set pivotDataWS = ThisWorkbook.Worksheets("pivot data")
    
    Dim tbl As ListObject
    Set tbl = pivotDataWS.ListObjects("tblEdiphiPivotData")
    
    Dim headerCell As Range
    Dim sortFieldDict As Dictionary
    Dim sortFieldName As String, sortFieldCode As String
    
    For Each headerCell In tbl.HeaderRowRange.Cells
        If InStr(headerCell.Value, "_code") > 0 Then
            Set sortFieldDict = New Dictionary
            sortFieldName = headerCell.Offset(0, 1).Value
            sortFieldCode = sortFieldName & "_code"
            With sortFieldDict
                .Add "code", sortFieldCode
                .Add "name", sortFieldName
            End With
            getSortFieldColl.Add sortFieldDict
        End If
    Next

End Function

Sub toggle_visible()

    Dim is_visible As Boolean
    is_visible = ThisWorkbook.Worksheets("trigger").visible = xlSheetVisible
    
    If is_visible Then
        setOpenState
    Else
        setEditState
    End If
    
End Sub

Sub setOpenState()
    
    Dim sht As Worksheet
    
    ThisWorkbook.Worksheets("Cover").visible = xlSheetVisible
    For Each sht In ThisWorkbook.Worksheets
        If sht.name <> "Cover" Then
            sht.visible = xlSheetVeryHidden
        End If
    Next
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("raw data")
    ws.UsedRange.EntireColumn.Delete
    
    ThisWorkbook.Windows(1).visible = True
    
End Sub

Sub setEditState()

    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        ws.visible = xlSheetVisible
    Next
    
End Sub

Sub logError(msg As String)
    
    'for now
    If errors Is Nothing Then Set errors = New Collection
    errors.Add msg
    Debug.Print "===ediphi=== ERROR: " & msg

End Sub

Function reportErrors() As String

    Dim e
    Dim msg As String
    Dim post As New UserEvents
    
    If errors Is Nothing Then Exit Function
    If errors.count = 0 Then Exit Function
    
    For Each e In errors
        msg = msg + e + vbLf
    Next
    
    post.slackPost "!!! ERRORS !!!" & vbLf & vbLf & msg, url:=myUrl
    MsgBox "The following errors occured: " & vbLf & vbLf & msg, vbCritical
    
    Set errors = Nothing
    
End Function

Function myUrl() As String
    On Error GoTo e1
    
    Dim projectId As String, estimateId As String
    projectId = ThisWorkbook.Worksheets("raw data").[b1]
    estimateId = ThisWorkbook.Worksheets("raw data").[c1]
    myUrl = "https://dpr.ediphi.com/projects/" & projectId & "/estimates/" & estimateId

Exit Function
e1:
    'pass

End Function

Sub closeMe(Optional hideErrors As Boolean)
    
    Dim post As New UserEvents
    
    If Not hideErrors Then
        reportErrors
    ElseIf Not thisReportBuilder.success Then
        post.slackPost "report process terminated early, user canceled"
    End If
    
    With Application
        .ScreenUpdating = False
        .EnableEvents = False
    End With
    
    With ThisWorkbook
        .Saved = True 'does not save, makes it close silently
        .Close
    End With
    
End Sub

Sub setEnv(varName As String, val)

    On Error GoTo e1
    Dim str As String
    str = CStr(val)
    Workbooks(EDIPHI_ADDIN_FILENAME).Worksheets("env").Range(varName).Value = str

Exit Sub
e1:
    logError "Failed to set env """ & varName & """ to """ & str & """"

End Sub

Function getEnv(varName As String) As String

    On Error GoTo e1
    getEnv = Workbooks(EDIPHI_ADDIN_FILENAME).Worksheets("env").Range(varName).Value

Exit Function
e1:
    'returns blank string
    
End Function

Function ediphiLink(tenant As String, project_id As String, estimate_id As String) As String
    
    ediphiLink = "https://" & tenant & ".ediphi.com/projects/" & _
                project_id & "/estimates/" & estimate_id

Exit Function
e1:
    ediphiLink = ""
    
End Function


Function validateShtName(str As String) As String

    Dim badChars As String
    badChars = "\/*?:[]'"
    Dim i As Integer
    
    For i = 1 To Len(badChars)
        str = Replace(str, Mid(badChars, i, 1), "_")
    Next

    If Len(str) > 30 Then str = Left(str, 27) & "..."

    validateShtName = str

End Function

Function validateRangeName(str As String) As String
    
    Const charsNotAllowed As String = "!""#$%&'()*+'`~/-:;<=>@[]^{}|~"
    Dim i As Integer
    
    str = Replace(str, " ", "_")
    For i = 1 To Len(charsNotAllowed)
        str = Replace(str, Mid(charsNotAllowed, i, 1), ".")
    Next
    
    If Len(str) > 249 Then str = Left(str, 249)
    validateRangeName = str

End Function

Function printDictList2Range(dictColl As Collection, startCell As Range, Optional noHeaders As Boolean, Optional asText As Boolean) As Range

    Dim arr
    arr = dictsTo2DArray(dictColl, noHeaders)
    Set printDictList2Range = printArr(startCell, arr, asText)

End Function

Function dictsTo2DArray(dictCollection As Collection, Optional noHeaders As Boolean) As Variant
    Dim allKeys As New Dictionary
    Dim dict As Dictionary
    Dim key As Variant
    Dim i As Integer, j As Integer
    Dim outputArray() As Variant
    
    If dictCollection.count = 0 Then Exit Function
    
    ' Gather all unique keys from each dictionary
    For Each dict In dictCollection
        For Each key In dict.Keys
            If Not allKeys.Exists(key) Then
                allKeys.Add key, Nothing
            End If
        Next key
    Next dict
    
    Dim startInt As Integer, endInt As Integer
    
    If noHeaders Then
        startInt = 0
        endInt = dictCollection.count - 1
    Else
        startInt = 1
        endInt = dictCollection.count
    End If
    
    ' Redim the output array to fit all keys and dictionaries
    If noHeaders Or dictCollection.count = 1 Then
        ReDim outputArray(0 To endInt, 0 To allKeys.count - 1)
    Else
        ReDim outputArray(0 To dictCollection.count, 0 To allKeys.count - 1)
    End If
    
    If Not noHeaders Then
        ' Set the first row to be the keys (headers)
        i = 0
        For Each key In allKeys.Keys
            outputArray(0, i) = key
            i = i + 1
        Next key
    End If
    
    ' Set the subsequent rows to be the values from each dictionary
    For i = startInt To endInt
        Set dict = dictCollection(i - startInt + 1) ' Adjusting index for collection
        For j = 0 To allKeys.count - 1
            key = allKeys.Keys(j)
            If dict.Exists(key) Then
                outputArray(i, j) = dict(key)
            Else
                outputArray(i, j) = ""
            End If
        Next j
    Next i
    
    ' Return the filled array
    dictsTo2DArray = outputArray
    
End Function


Function isFirstReport()

    isFirstReport = Not ThisWorkbook.Worksheets("EstData").[rngIsTemp]

End Function

Sub updateLocally()
    
    Dim deployPath As String
    
    gitExplode
    setOpenState
    With ThisWorkbook
        deployPath = Workbooks(EDIPHI_ADDIN_FILENAME).Path & "\ediphi_cache\" & .name
        On Error Resume Next
        SetAttr deployPath, vbNormal
        On Error GoTo 0
        .Save
        Application.DisplayAlerts = False
        .SaveAs deployPath
        Application.DisplayAlerts = True
        SetAttr deployPath, vbReadOnly
    End With
    
    MsgBox "Report Builder Updated Locally", vbInformation
    ThisWorkbook.Close Savechanges:=False
    
End Sub

Function deepCopy(source As Dictionary) As Dictionary
    Dim key As Variant
    Dim newDict As Dictionary
    Set newDict = New Dictionary
    
    For Each key In source.Keys
        If TypeName(source(key)) = "Dictionary" Then
            ' Recursive call to deepcopy if the item is a dictionary
            newDict.Add key, deepCopy(source(key))
        Else
            ' Directly copy the value if it is not a dictionary
            newDict.Add key, source(key)
        End If
    Next key
    
    Set deepCopy = newDict
End Function

