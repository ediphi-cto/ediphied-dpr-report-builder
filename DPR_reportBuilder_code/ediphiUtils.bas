Attribute VB_Name = "ediphiUtils"
Option Explicit
Public Const EDIPHI_ADDIN_FILENAME As String = "ediphi_addin.xlam"
Public thisReportBuilder As EdiphiReportBuilder
Public errors As Collection
Public Const TRY_UPDATING_MSG As String = "Try updating it by pressing ALT + F8, then typing 'ediphiUpdate', and hit Enter"

Sub tryApi()

    Dim api As New EdiphiAPI
    Dim rawJson As String
    rawJson = api.getEstimateJSON("5347cddc-d3dd-46dd-b2de-d81f774ac3cc")
    Dim jsonDict As Dictionary
    Set jsonDict = ParseJson(rawJson)
    
    pp jsonDict
    
    
End Sub


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
    
    ThisWorkbook.Worksheets("splash").visible = xlSheetVisible
    For Each sht In ThisWorkbook.Worksheets
        If sht.Name <> "splash" Then
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
    
    post.slackPost "!!! ERRORS !!!" & vbLf & vbLf & msg
    MsgBox "The following errors occured: " & vbLf & vbLf & msg, vbCritical
    
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

Sub setEnv(varName As String, Val)

    On Error GoTo e1
    Dim str As String
    str = CStr(Val)
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

Function printDictList2Range(dictColl As Collection, startCell As Range) As Range

    Dim arr
    arr = dictsTo2DArray(dictColl)
    Set printDictList2Range = printArr(startCell, arr)

End Function

Function dictsTo2DArray(dictCollection As Collection) As Variant
    Dim allKeys As New Dictionary
    Dim Dict As Dictionary
    Dim Key As Variant
    Dim i As Integer, j As Integer
    Dim outputArray() As Variant
    
    If dictCollection.count = 0 Then Exit Function
    
    ' Gather all unique keys from each dictionary
    For Each Dict In dictCollection
        For Each Key In Dict.Keys
            If Not allKeys.Exists(Key) Then
                allKeys.Add Key, Nothing
            End If
        Next Key
    Next Dict
    
    ' Redim the output array to fit all keys and dictionaries
    ReDim outputArray(0 To dictCollection.count, 0 To allKeys.count - 1)
    
    ' Set the first row to be the keys (headers)
    i = 0
    For Each Key In allKeys.Keys
        outputArray(0, i) = Key
        i = i + 1
    Next Key
    
    ' Set the subsequent rows to be the values from each dictionary
    For i = 1 To dictCollection.count
        Set Dict = dictCollection(i)
        For j = 0 To allKeys.count - 1
            Key = allKeys.Keys(j)
            If Dict.Exists(Key) Then
                outputArray(i, j) = Dict(Key)
            Else
                outputArray(i, j) = ""
            End If
        Next j
    Next i
    
    ' Return the filled array
    dictsTo2DArray = outputArray
    
End Function

Sub updateLocally()
    
    Dim deployPath As String
    
    gitExplode
    setOpenState
    With ThisWorkbook
        deployPath = Workbooks(EDIPHI_ADDIN_FILENAME).Path & "\ediphi_cache\" & .Name
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
    
End Sub
