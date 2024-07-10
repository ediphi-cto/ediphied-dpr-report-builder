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
Public estimateJson As Dictionary
Public projectJson As Dictionary

Private Sub triggerWS_Change(ByVal Target As Range)
    
    If UCase(Target.address(False, False)) = "A1" Then
            If Me.debugMode Then On Error GoTo 0 Else On Error Resume Next
            debugMode = (getEnv("DEBUG") = 1)
            If Me.debugMode Then
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
    
    If Me.debugMode Then Stop
    'first, get the column count
    Dim colCt As Integer
    Dim csvWS As Worksheet
    Dim csvWB As Workbook
    Dim csvFullname As String
    Set csvWB = Workbooks(csvFileName)
    csvFullname = csvWB.FullName
    colCt = csvWB.Worksheets(1).UsedRange.Columns.count
    
    'then close the csv silent
    csvWB.Close Savechanges:=False
    
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
    
    With Me
        .fetchEstimateJson
        .transformDataForPivotTable
        .setMetaData
    End With
    
End Sub

Function ProperCase(val) As String
    
    Dim str As String
    str = cStr_safe(val)
    ProperCase = Replace(WorksheetFunction.Proper(str), "_", " ")

End Function

Sub setMetaData()
    
    Dim metaDataWS As Worksheet
    Set metaDataWS = ThisWorkbook.Worksheets("EstData")
    On Error Resume Next
    With metaDataWS
        .[rngEstName].Value = thisReportBuilder.EstimateName
        .[rngEstDate].Value = Format(Now(), "m/dd/yyyy")
        .[rngProjectName].Value = thisReportBuilder.ProjectName
        .[rngEstimator].Value = Application.UserName
        .[rngJobSize].Value = Me.estimateJson("estimate")("stats")("grossArea")
        .[rngProjectAddress].Value = Split(Me.projectJson("address"), ",")(0)
        .[rngProjectCityStateZip].Value = _
            Me.projectJson("city") & ", " & Me.projectJson("state") & " " & Me.projectJson("zip")
        .[rngProjectClient].Value = Me.projectJson("client_name")
        .[rngSubHeading2].Value = ProperCase(Me.estimateJson("estimate")("phase"))
        .[rngSubHeading3].Value = Me.projectJson("task_number")
        .[rngEstStatus].Value = ProperCase(Me.projectJson("phase"))
        .[rngEstType].Value = ProperCase(Me.projectJson("industry"))
    End With

End Sub

Sub transformDataForPivotTable()
    
    Dim neededCols As Dictionary
    Set neededCols = findNeededColumns()
    Dim dataColl As Collection
    
    Dim startRan As Range
    Set startRan = Me.pivotDataWS.[a1]

    Set dataColl = pivotDataRows(neededCols, splitByUse:=False)
    printPivotDataAsTable dataColl, startRan, "tblEdiphiPivotData"
    
    Dim tbl As ListObject
    Dim lastHeaderCol As Integer
    Set tbl = Me.pivotDataWS.ListObjects("tblEdiphiPivotData")
    lastHeaderCol = tbl.HeaderRowRange.Columns(tbl.HeaderRowRange.Columns.count).Column
    Set startRan = Me.pivotDataWS.Cells(tbl.HeaderRowRange.row, lastHeaderCol + 2)
    
    Set dataColl = pivotDataRows(neededCols, splitByUse:=True)
    printPivotDataAsTable dataColl, startRan, "tblEdiphiPivotDataUseSplit"
    
   'populate the markups table
   parseMarkups2Table
   
End Sub


Function findNeededColumns() As Dictionary
    Dim ran As Range
    Set ran = Me.rawDataWS.[a3]
    Dim colNameMap As Dictionary
    Set colNameMap = ediphiToDPRcolumnNameMap
    Dim neededCols As New Dictionary
    Set Me.sortFieldColl = New Collection
    Dim sortFieldDict As Dictionary
    Dim sortFieldName As String, sortFieldCode As String
    
    'Find Columns of Interest
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
            neededCols.Add sortFieldName & "_combined", ran.Column
        End If
        Set ran = ran.Offset(0, 1)
    Loop

    Set findNeededColumns = neededCols

End Function
   
Function pivotDataRows(neededCols As Dictionary, splitByUse As Boolean) As Collection
    'this function takes the raw data ws and creates dictionaries for each row (a line item)
    'it will create multiple dictionaries for line items that are allocated to use groups if splitByUse is true
    
    Dim ran As Range
    Dim rowDict As Dictionary
    Dim dataColl As New Collection
    Dim colOffset As Integer
    Dim k
    Dim grandTotal As Double, crewHours As Double, unitsPerHr As Double
    Dim li As Dictionary
    
    Set ran = Me.rawDataWS.[a4]
    Do Until ran.Value = 0
        Set rowDict = New Dictionary
        For Each k In neededCols.Keys
            colOffset = neededCols(k) - 1
            If InStr(k, "_combined") > 0 Then
                rowDict.Add k, ran.Offset(0, colOffset).Value & " - " & ran.Offset(0, colOffset + 1).Value
            Else
                rowDict.Add k, ran.Offset(0, colOffset).Value
            End If
        Next
        If Not isChildItem(rowDict) And isBaseBid(rowDict) Then
            'filter out alternates and assembly child line items
            
            'create a grand total column
            grandTotal = cDbl_safe(rowDict("UnitPrice")) * cDbl_safe(rowDict("TakeoffQty"))
            grandTotal = Round(grandTotal, 2)
            rowDict.Add "GrandTotal", grandTotal
            
            'create a manhours column
            unitsPerHr = cDbl_safe(rowDict("UnitsPerHour"))
            If unitsPerHr <> 0 Then
                crewHours = cDbl_safe(rowDict("TakeoffQty")) / unitsPerHr
            Else
                crewHours = 0
            End If
            rowDict.Add "Manhours", crewHours 'MN TODO this needs to change once ediphi allows crews applied to labor rates
            
            'create a notes column
            If rowDict.Exists("ItemNote") Then
                'pass
            Else
                Set li = Me.findOne(rowDict("id"), "line_items")
                If Not li Is Nothing And li.Exists("notes") Then rowDict.Add "ItemNote", li("notes")
            End If
            
            'create use group columns
            Dim splitUseRows As Collection
            Dim useRowDict As Dictionary
            Set splitUseRows = calculateUseGroupValues(rowDict, splitByUse:=splitByUse)
            For Each useRowDict In splitUseRows
                dataColl.Add useRowDict
            Next
            
        End If
        Set ran = ran.Offset(1, 0)
        If ran.row > ran.Parent.UsedRange.Rows.count Then Exit Do
    Loop
    
    If dataColl.count = 0 Then
        'no line items, can happen when the estimate is all alternates and no base bid
        Me.Cancel = True
        MsgBox "There aren't any Base Bid line items in your ediphi estimate", vbCritical
        Exit Function
    End If
    
    Set pivotDataRows = dataColl
    
End Function

Sub printPivotDataAsTable(dataColl As Collection, startCell As Range, tblName As String)
    Dim printRan As Range
    Set printRan = printDictList2Range(dictColl:=dataColl, startCell:=startCell, asText:=True)
   
    Dim tbl As ListObject
    Set tbl = Me.pivotDataWS.ListObjects.Add(xlSrcRange, printRan, , xlYes)
    tbl.name = tblName
    
    'numberformat columns
    Dim col As ListColumn
    Dim cell As Range
    For Each col In tbl.ListColumns
        If InStr("UnitPrice TakeoffQty", col.name) > 0 Then
            With col.DataBodyRange
                .NumberFormat = "0.00"
                For Each cell In .Cells
                    cell.Value = cell.Value 'required for formatting to take effect
                Next
            End With
        End If
    Next

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
    
    If markups.count = 0 Then
        'add a blank row so the table creates, many things ref the resulting tblTotals table
        markups.Add ParseJson( _
        "{""id"":"""",""order"":"""",""description"":"""",""percentage"":"""",""total"":""""}")
    End If
        
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
    Set printRan = printDictList2Range(dataColl, ThisWorkbook.Worksheets("tables").[\markups])
    Dim tbl As ListObject
    Set tbl = printRan.Parent.ListObjects.Add(xlSrcRange, printRan, , xlYes)
    
    With tbl
        .name = "tblTotals"
        .Sort.SortFields.Clear
        .Sort.SortFields.Add key:=Range("tblTotals[SortOrder]"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        With .Sort
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
    End With

Exit Sub
e1:
    logError "Markups did not parse correctly"
    'failsafe, but markups table will not exist

End Sub


Sub fetchEstimateJson()

    If Me.debugMode Then On Error GoTo 0 Else On Error GoTo e1
    
    Dim api As New EdiphiAPI
    Dim rawJson As String
    rawJson = api.getEstimateJSON(Me.estimateId)
    Set Me.estimateJson = ParseJson(rawJson)("sanitizedEstimate")
    
    If Me.debugMode Then On Error GoTo 0 Else On Error GoTo e2
    Set Me.projectJson = api.getProject(Me.estimateJson("estimate")("project"))
    
Exit Sub
e1:
    logError "failed to fetch estimate json from the api"


Exit Sub
e2:
    logError "failed to fetch project json from the api"

End Sub

Function findOne(thingPropVal As String, thingName As String, Optional thingProp As String = "id") As Dictionary
    If Me.debugMode Then On Error GoTo 0 Else On Error GoTo e1
    
    Dim things As Collection
    Set things = Me.estimateJson(thingName)
    Dim thing As Dictionary
    For Each thing In things
        If thing(thingProp) = thingPropVal Then
            Set findOne = thing
            Exit Function
        End If
    Next

Exit Function
e1:
    pp "ERROR === could not find {" & thingProp & ": " & thingPropVal & "} in '" & thingName & "' within estimate json"
    'do not warn user, could be noisy

End Function

Function isChildItem(dict As Dictionary) As Boolean
    
    On Error Resume Next
    isChildItem = dict("parent") <> ""
    
End Function

Function isBaseBid(dict As Dictionary) As Boolean
    Dim alt As Dictionary
    Dim altId As String
    If dict.Exists("alternate") Then
        altId = dict("alternate")
        If altId = "" Then
            isBaseBid = True
            Exit Function
        End If
    Else
        isBaseBid = True
        Exit Function
    End If
    
    Set alt = findOne(dict("alternate"), "alternates")
    If alt Is Nothing Then Exit Function
    
    isBaseBid = alt("accepted")

End Function

Function calculateUseGroupValues(dict As Dictionary, splitByUse As Boolean) As Collection 'MN TO: return as collection
    On Error GoTo e1
    Dim allocations As Dictionary
    Dim i As Integer, hasMultiple As Boolean
    Dim coll As New Collection
    Dim k
    Dim useAreasDict As Dictionary
    
    
    Set allocations = ParseJson(dict("use_groups"))
    dict.Remove "use_groups"

    If allocations.count = 0 Then GoTo e1
    
    If splitByUse And allocations.count > 1 Then
    
        Dim newDict As Dictionary
        For Each k In allocations.Keys
            Set newDict = New Dictionary
            Set newDict = deepCopy(dict)
            newDict("GrandTotal") = Round(cDbl_safe(dict("GrandTotal")) * cDbl_safe(allocations(k)) / 100, 2)
            newDict("TakeoffQty") = cDbl_safe(dict("TakeoffQty")) * cDbl_safe(allocations(k) / 100)
            newDict.Add "Use Group_code", k
            newDict.Add "Use Group", k
            newDict.Add "Use Group_area", getAreaByUseName(k)
            newDict.Add "Use Group_combined", k
            coll.Add newDict
        Next
        
    Else
    
        i = 0
        For Each k In allocations.Keys
            i = i + 1
            dict.Add k, Round(cDbl_safe(dict("GrandTotal")) * cDbl_safe(allocations(k)) / 100, 2)
            dict.Add k & "_qty", cDbl_safe(dict("TakeoffQty")) * cDbl_safe(allocations(k) / 100)
            If i = 1 Then
                dict.Add "Use Group_code", k
                dict.Add "Use Group", k
                dict.Add "Use Group_combined", k
            End If
        Next
        
        If i > 1 Then
            dict("Use Group_code") = "zzz"
            dict("Use Group") = "Multiple"
            dict("Use Group_combined") = "Multiple"
        End If
        coll.Add dict
        
    End If

finally:
    Set calculateUseGroupValues = coll

Exit Function
e1:
    dict.Add "Use Group_code", "zzzz"
    dict.Add "Use Group", "None"
    dict.Add "Use Group_combined", "None"
    Set coll = New Collection
    coll.Add dict
    Resume finally
    
End Function

Function getAreaByUseName(useName) As Double
    If Me.debugMode Then On Error GoTo 0 Else On Error GoTo e1
    Dim useGroup As Dictionary
    Dim useGroups As Collection
    Set useGroups = Me.estimateJson("estimate")("stats")("useGroups")
    
    For Each useGroup In useGroups
        If useGroup("label") = useName Then
            getAreaByUseName = cDbl_safe(useGroup("gsf"))
            Exit Function
        ElseIf useGroup("use_group") = useName Then
            getAreaByUseName = cDbl_safe(useGroup("gsf"))
            Exit Function
        End If
    Next

Exit Function
e1:
    'pass
    
End Function

Property Get EstimateName() As String
    EstimateName = Me.rawDataWS.Range("E1").Value
End Property

Property Get ProjectName() As String
    ProjectName = Me.rawDataWS.Range("D1").Value
End Property

Property Get estimateId() As String
    estimateId = Me.rawDataWS.Range("C1").Value
End Property

Property Get projectId() As String
    projectId = Me.rawDataWS.Range("B1").Value
End Property

Property Get embeddedJSON() As String
    embeddedJSON = Me.rawDataWS.Range("A2").Value
End Property

Property Get embeddedXML() As String
    embeddedXML = Me.rawDataWS.Range("B2").Value
End Property

Function getStat(statName As String, Optional useGroupID As String)
    'stats
    '     "grossArea": 306068,
    '     "processed": "2024-05-02T05:03:00.795Z",
    '     "updatedAt": "2024-05-02T05:03:00.794Z",
    '     "totalUnits": 0,
    '     "markupsTotal": 4761999.49,
    '     "totalLineItems": 40,
    '     "costOfTradesSubtotal": 46594157
    
    ' use group stats, for totals, just provide keyname like "area" for statName
    '          "id": "ea8ee16b-6b83-46f0-b7d5-79bdc25fc0d1",
    '          "gsf": 73889,
    '          "nsf": 0,
    '          "label": "Proton Therapy",
    '          "order": 1,
    '          "totals": {
    '            "area": 73889,
    '            "beds_total": 0,
    '            "unit_count": 0,
    '            "baths_total": 0,
    '            "unit_area_total": 0,
    '            "half_baths_total": 0,
    '            "terminal_hvacs_total": 0
    '          },
    '          "estimate": "5347cddc-d3dd-46dd-b2de-d81f774ac3cc",
    '          "use_group": "Cancer Care",
    '          "unit_count": 0,
    '          "exclude_area": false,
    '          "use_group_id": "cf570083-c783-4c31-bd4f-bed540426a79",
    '          "costOfTradesSubtotal": 41832086
    
    If Not Me.debugMode Then GoTo e1
    Dim val
    Dim useGroups As Collection
    Dim useGroup As Dictionary
    
    If useGroupID = "" Then
        val = Me.estimateJson("estimate")("stats")(statName)
    Else
        Set useGroups = Me.estimateJson("estimate")("stats")("UseGroups")
        For Each useGroup In useGroups
            If useGroup("id") = useGroupID Then
                If useGroup.Exists(statName) Then
                    val = useGroup(statName)
                Else
                    val = useGroup("totals")(statName)
                End If
                GoTo finally
            End If
        Next
    End If

finally:
    getStat = cStr_safe(val)

Exit Function
e1:
    pp "failed to find stat '" & statName & "' within estimate json " & useGroupID
    'pass

End Function

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

    ActiveWindow.WindowState = xlMaximized

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
        .Add "labor_uc", "Labor10"
        .Add "prod_rate", "UnitsPerHour"
        .Add "material_uc", "Material20"
        .Add "sub_uc", "Sub50"
        .Add "equip_uc", "Equipment30"
        .Add "other_uc", "Other40"
        .Add "alternate", "alternate"
        .Add "use_groups", "use_groups"
    End With
    
End Function
