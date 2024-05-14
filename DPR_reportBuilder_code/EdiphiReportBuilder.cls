VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EdiphiReportBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public WithEvents triggerWS As Worksheet
Attribute triggerWS.VB_VarHelpID = -1
Public rawDataWS As Worksheet
Public pivotDataWS As Worksheet
Public reportWB As Workbook
Public sortFieldColl As Collection
Public Cancel As Boolean
Public debugMode As Boolean
Public success As Boolean

Private Sub triggerWS_Change(ByVal Target As Range)
    
    If UCase(Target.Address(False, False)) = "A1" Then
            If Me.debugMode Then On Error GoTo 0 Else On Error Resume Next
            debugMode = (getEnv("DEBUG") = 1)
            If Me.debugMode Then
                DoEvents
                Stop
                On Error GoTo 0
            Else
                On Error GoTo e2
            End If
            migrateDataFrom csvFileName:=Target.Value
            If Not Me.Cancel Then frmReportLevel.Show Else closeMe
    End If

Exit Sub
e2:
    logError "Failed to parse data from ediphi" & vbLf & vbLf & TRY_UPDATING_MSG
    closeMe
    
End Sub

Sub migrateDataFrom(csvFileName As String)

    'first, get the column count
    Dim colCt As Integer
    Dim csvWS As Worksheet
    Dim csvWB As Workbook
    Dim csvFullname As String
    Set csvWB = Workbooks(csvFileName)
    csvFullname = csvWB.FullName
    colCt = csvWB.Worksheets(1).UsedRange.Columns.count
    
    'then close the csv silent
    csvWB.Close SaveChanges:=False
    
    'create an array of 2's the length of the column count
    Dim i As Integer
    Dim textDataTypes
    ReDim textDataTypes(1 To colCt)
    
    For i = 1 To colCt
        textDataTypes(i) = 2  ' 2 means text format
    Next i
    
    Dim ws As Worksheet
    Set ws = Me.rawDataWS
    With ws.QueryTables.Add(Connection:="TEXT;" & csvFullname, Destination:=ws.Range("A1"))
        .TextFileParseType = xlDelimited
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        .TextFileConsecutiveDelimiter = False
        .TextFileCommaDelimiter = True
        .TextFileColumnDataTypes = textDataTypes ' Set all columns to text
        .Refresh
    End With
    
    Me.setMetaData
    Me.transformDataForPivotTable
    
End Sub

Sub setMetaData()
    
    Dim metaDataWS As Worksheet
    Set metaDataWS = ThisWorkbook.Worksheets("EstData")
    With metaDataWS
        .[rngEstName].Value = thisReportBuilder.EstimateName
        .[rngEstDate].Value = Format(Now(), "m/dd/yyyy")
        .[rngProjectName].Value = thisReportBuilder.ProjectName
        .[rngEstimator].Value = Application.UserName
    End With

End Sub

Sub transformDataForPivotTable()

    Dim ran As Range
    Set ran = Me.rawDataWS.[a3]
    Dim colNameMap As Dictionary
    Set colNameMap = ediphiToDPRcolumnNameMap
    Dim neededCols As New Dictionary
    Set Me.sortFieldColl = New Collection
    Dim sortFieldDict As Dictionary
    Dim sortFieldName As String, sortFieldCode As String
    
    '[1] Find Columns of Interest
    'create a neededCols as dict where:
    '      key: column names, value: column ints
    
    Do Until ran.Value = 0
        If colNameMap.Exists(ran.Value) Then
            neededCols.Add colNameMap(ran.Value), ran.Column
        ElseIf InStr(ran.Value, "_code") > 0 Then
            'it's a sort code column, and it's description column is always one column to the right
            Set sortFieldDict = New Dictionary
            sortFieldName = ran.Offset(0, 1).Value
            sortFieldCode = sortFieldName & "_code"
            With sortFieldDict
                .Add "code", sortFieldCode
                .Add "name", sortFieldName
            End With
            Me.sortFieldColl.Add sortFieldDict
            neededCols.Add sortFieldCode, ran.Column
            neededCols.Add sortFieldName, ran.Offset(0, 1).Column
        End If
        Set ran = ran.Offset(0, 1)
    Loop
   neededCols.Add "ItemNote", 1000 'MN: TOD Once notes are included in csv
   
   '[2] Gather Data Rows into Collection of Dicts
   
    Dim rowDict As Dictionary
    Dim dataColl As New Collection
    Dim colOffset As Integer
    Dim k
    Dim grandTotal As Double
    
    Set ran = Me.rawDataWS.[a4]
    Do Until ran.Value = 0
        Set rowDict = New Dictionary
        For Each k In neededCols.Keys
            colOffset = neededCols(k) - 1
            rowDict.Add k, ran.Offset(0, colOffset).Value
        Next
        If Not isChildItem(rowDict) And isBaseBid(rowDict) Then
            'filter out alternates and assembly child line items
            grandTotal = cDbl_safe(rowDict("UnitPrice")) * cDbl_safe(rowDict("TakeoffQty"))
            grandTotal = Round(grandTotal, 2)
            rowDict.Add "GrandTotal", grandTotal
            calculateUseGroupValues rowDict
            dataColl.Add rowDict
        End If
        Set ran = ran.Offset(1, 0)
        If ran.row > ran.Parent.UsedRange.Rows.count Then Exit Do
    Loop
    
    If dataColl.count = 0 Then
        'no line items, can happen when the estimate is all alternates and no base bid
        Me.Cancel = True
        MsgBox "There aren't any Base Bid line items in your ediphi estimate", vbCritical
        Exit Sub
    End If
    
    Dim printRan As Range
    Set printRan = printDictList2Range(dataColl, Me.pivotDataWS.[a1])
   
    Dim tbl As ListObject
    Set tbl = Me.pivotDataWS.ListObjects.Add(xlSrcRange, printRan, , xlYes)
    tbl.Name = "tblEdiphiPivotData"
    
   'populate the markups table
   parseMarkups2Table

End Sub

Sub parseMarkups2Table()
    If Me.debugMode Then On Error GoTo 0 Else On Error GoTo e1
    Dim root As Dictionary
    Set root = ParseJson(Me.embeddedJSON)
    Dim markups As Collection
    Set markups = root("markups")
    Dim markup As Dictionary
    Dim markupRow As Dictionary
    Dim dataColl As New Collection
    
    For Each markup In markups
        Set markupRow = New Dictionary
        With markupRow
            .Add "Identity", markup("id")
            .Add "SortOrder", markup("order")
            .Add "Class", "GROSS"
            .Add "Name", markup("description")
            .Add "Percent", markup("percentage")
            .Add "Amount", markup("total")
            .Add "Allocate", "Upper"
            .Add "PrimarySort", markup("order")
            .Add "JobCostCategory", ""
            .Add "JobCost", ""
            .Add "JobCostName", ""
            .Add "Updated", ""
            .Add "est2Amount", ""
            .Add "Variance", ""
        End With
        dataColl.Add markupRow
    Next
    
    Dim printRan As Range
    Set printRan = printDictList2Range(dataColl, ThisWorkbook.Worksheets("XMLTables").[\markups])
    Dim tbl As ListObject
    Set tbl = printRan.Parent.ListObjects.Add(xlSrcRange, printRan, , xlYes)
    tbl.Name = "tblTotals"
    
    'MN TODO: Sort the table

Exit Sub
e1:
    logError "Markups did not parse correctly"
    'failsafe, but markups table will not exist

End Sub

Function isChildItem(Dict As Dictionary) As Boolean
    
    isChildItem = Dict("parent") <> ""
    
End Function

Function isBaseBid(Dict As Dictionary) As Boolean
    
    isBaseBid = Dict("alternate") = ""

End Function

Sub calculateUseGroupValues(Dict As Dictionary)

    Dim allocations As Dictionary
    Set allocations = ParseJson(Dict("use_groups"))
    
    Dim k
    For Each k In allocations.Keys
        Dict.Add k, Round(Dict("GrandTotal") * cDbl_safe(allocations(k)) / 100, 2)
    Next
    
    Dict.Remove "use_groups"

End Sub

Property Get EstimateName() As String
    EstimateName = Me.rawDataWS.Range("E1").Value
End Property

Property Get ProjectName() As String
    ProjectName = Me.rawDataWS.Range("D1").Value
End Property

Property Get EstimateID() As String
    EstimateID = Me.rawDataWS.Range("C1").Value
End Property

Property Get ProjectID() As String
    ProjectID = Me.rawDataWS.Range("B1").Value
End Property

Property Get embeddedJSON() As String
    embeddedJSON = Me.rawDataWS.Range("A2").Value
End Property

Property Get embeddedXML() As String
    embeddedXML = Me.rawDataWS.Range("B2").Value
End Property

Property Get metaDataAsString() As String

    With Me
        metaDataAsString = metaDataAsString & "PROJECT: " & .ProjectName & vbLf
        metaDataAsString = metaDataAsString & "ESTIMATE: " & .EstimateName & vbLf
    End With

End Property

Private Sub Class_Initialize()
    If Me.debugMode Then
        On Error GoTo 0
    Else
        On Error GoTo e1
    End If
    
    With Me
        Set .triggerWS = ThisWorkbook.Worksheets("trigger")
        'MN: I would instead typically make this a new WB as an xlsx to remove code from user files
        'the reportWB is the final result for the user
        Set .reportWB = ThisWorkbook
        Set .rawDataWS = reportWB.Worksheets("raw data")
        Set .pivotDataWS = reportWB.Worksheets("pivot data")
    End With

Exit Sub
e1:
    logError "The ReportBuilder seems to be corrupt." & vbLf & vbLf & TRY_UPDATING_MSG
    closeMe

End Sub

Function ediphiToDPRcolumnNameMap() As Dictionary

    Set ediphiToDPRcolumnNameMap = New Dictionary
    With ediphiToDPRcolumnNameMap
        'key: ediphi column name, value: DPR column name
        .Add "id", "id"
        .Add "order", "ItemCode"
        .Add "name", "Description"
        .Add "note", "ItemNote"
        .Add "quantity", "TakeoffQty"
        .Add "uom", "TakeoffUnit"
        .Add "total_uc", "UnitPrice"
        .Add "alternate", "alternate"
        .Add "use_groups", "use_groups"
    End With
    
End Function
