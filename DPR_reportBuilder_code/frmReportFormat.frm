VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmReportFormat 
   Caption         =   "DPR Report Builder"
   ClientHeight    =   6012
   ClientLeft      =   72
   ClientTop       =   468
   ClientWidth     =   13500
   OleObjectBlob   =   "frmReportFormat.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmReportFormat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit
Dim sPF As String
Dim sType As String
Dim bLoading As Boolean
Dim X, iRow
Dim C As Integer

Private Sub CheckBox1_Click()
    If bLoading = True Then Exit Sub
    Col_Update (1)
End Sub

Private Sub CheckBox2_Click()
    If bLoading = True Then Exit Sub
    Col_Update (2)
End Sub

Private Sub CheckBox3_Click()
    If bLoading = True Then Exit Sub
    Col_Update (3)
End Sub

Private Sub CheckBox4_Click()
    If bLoading = True Then Exit Sub
    Col_Update (4)
End Sub

Private Sub CheckBox5_Click()
    If bLoading = True Then Exit Sub
    Col_Update (5)
End Sub

Private Sub CheckBox6_Click()
    If bLoading = True Then Exit Sub
    Col_Update (6)
End Sub

Private Sub CheckBox7_Click()
    If bLoading = True Then Exit Sub
    Col_Update (7)
End Sub

Private Sub CheckBox8_Click()
    If bLoading = True Then Exit Sub
    Col_Update (8)
End Sub

Private Sub CheckBox9_Click()
    If bLoading = True Then Exit Sub
    Col_Update (9)
End Sub

Private Sub CheckBox10_Click()
    If bLoading = True Then Exit Sub
    Col_Update (10)
End Sub

Private Sub CheckBox11_Click()
    If bLoading = True Then Exit Sub
    Col_Update (11)
End Sub

Private Sub CheckBox12_Click()
    If bLoading = True Then Exit Sub
    Col_Update (12)
End Sub

Private Sub CheckBox13_Click()
    If bLoading = True Then Exit Sub
    Col_Update (13)
End Sub

Private Sub CheckBox14_Click()
    If bLoading = True Then Exit Sub
    Col_Update (14)
End Sub

Private Sub CheckBox15_Click()
    If bLoading = True Then Exit Sub
    Col_Update (15)
End Sub

Private Sub CheckBox16_Click()
    If bLoading = True Then Exit Sub
    Col_Update (16)
End Sub

Private Sub CheckBox17_Click()
    If bLoading = True Then Exit Sub
    Col_Update (17)
End Sub

Private Sub CheckBox18_Click()
    If bLoading = True Then Exit Sub
    Col_Update (18)
End Sub

Private Sub chkVarAddon_Click()
    Call ckbAddon_Click
End Sub

Private Sub chkVarComments_Click()
Dim iCol As Integer
    If bLoading = True Then Exit Sub
    With ows.UsedRange
        iCol = .Columns(.Columns.Count).Column
    End With
    If chkVarComments.Value = True Then
        ows.Columns(iCol).Hidden = False
    Else
        ows.Columns(iCol).Hidden = True
    End If
    Call SheetFormatting
End Sub

Private Sub chkVarQty_Click()
    If bLoading = True Then Exit Sub
    On Error Resume Next
    For C = iCol - 4 To iCol - 10 Step -3
        If chkVarQty.Value = True Then
            ows.Columns(C).EntireColumn.Hidden = False
        Else
            ows.Columns(C).EntireColumn.Hidden = True
        End If
    Next C
 'Hide U/M columns
    C = 0
    For C = iCol - 6 To iCol - 9 Step -3
        If chkVarQty.Value = True Then
            ows.Columns(C).EntireColumn.Hidden = False
        Else
            ows.Columns(C).EntireColumn.Hidden = True
        End If
    Next C
    On Error GoTo 0
End Sub

Private Sub chkVarSlicers_Click()
Dim Sh As Shape
Dim oSlicer As Slicer
Dim oSlicercache As SlicerCache
Dim X As Integer, C
If bLoading = True Then Exit Sub

Application.ScreenUpdating = False
    If chkVarSlicers = True Then
        sRprt = ows.PivotTables(1).name
        Set lObj = Sheet0.ListObjects("tblRptTrack")
        Set C = lObj.ListColumns(1).DataBodyRange.Find(sRprt, LookIn:=xlValues)
        If Not C Is Nothing Then
            X = C.Offset(0, 1).Value
        End If
        For i = 1 To X
            Select Case i
                Case Is = 1
                    ActiveWorkbook.SlicerCaches.Add2(pt, sLvl1Item) _
                        .Slicers.Add ActiveSheet, , , sLvl1Item, 10, 900, 220, 400
                Case Is = 2
                    ActiveWorkbook.SlicerCaches.Add2(pt, sLvl2Item) _
                        .Slicers.Add ActiveSheet, , , sLvl2Item, 10, 1125, 220, 400
                Case Is = 3
                    ActiveWorkbook.SlicerCaches.Add2(pt, sLvl3Item) _
                    .Slicers.Add ActiveSheet, , , sLvl3Item, 120, 1000, 220, 400
                Case Is = 4
                    ActiveWorkbook.SlicerCaches.Add2(pt, sLvl4Item) _
                    .Slicers.Add ActiveSheet, , , sLvl4Item, 120, 1225, 220, 400
                Case Is = 5
                    ActiveWorkbook.SlicerCaches.Add2(pt, sLvl5Item) _
                    .Slicers.Add ActiveSheet, , , sLvl5Item, 190, 1100, 220, 400
                Case Else
            End Select
        Next i
    Else
        For Each Sh In ActiveSheet.Shapes
            If Sh.name <> "grpHeading" And Sh.name <> "grpHeadingVar" Then
                Sh.Delete
            End If
        Next
    End If
Application.ScreenUpdating = True
End Sub

Private Sub chkVarUnit_Click()
    If bLoading = True Then Exit Sub
    On Error Resume Next
    If bLoading = True Then Exit Sub
    On Error Resume Next
    For C = iCol - 5 To iCol - 8 Step -3
        If chkVarUnit.Value = True Then
            ows.Columns(C).EntireColumn.Hidden = False
            ows.Columns(iCol - 3).EntireColumn.Hidden = False
        Else
            ows.Columns(C).EntireColumn.Hidden = True
            ows.Columns(iCol - 3).EntireColumn.Hidden = True
        End If
    Next C
    On Error GoTo 0
End Sub

Private Sub ckbAddon_Click()
    If bLoading = True Then Exit Sub
    If ckbAddon.Value = True And sType = "Level" Then
        Call Addons
        SheetFormatting
    ElseIf ckbAddon.Value = True And sType = "XTab" Then
        Call XTabAddons
    ElseIf ckbAddon.Value = True And sType = "CEst" Then
        Call CEstAddons
    ElseIf chkVarAddon.Value = True And sType = "Var" Then
        Call Addons_VAR
    Else
        Call clearAddons
    End If
End Sub

Private Sub ckbBlankLine_Click()
    If bLoading = True Then Exit Sub
    On Error Resume Next
    If ckbAddon = True Then
        clearAddons
        If ckbBlankLine = True Then
            pt.PivotFields(sPF).LayoutBlankLine = True
        Else
            pt.PivotFields(sPF).LayoutBlankLine = False
        End If
        ckbAddon_Click
    Else
        If ckbBlankLine = True Then
            pt.PivotFields(sPF).LayoutBlankLine = True
        Else
            pt.PivotFields(sPF).LayoutBlankLine = False
        End If
    End If
    On Error GoTo 0
End Sub

Private Sub ckbComments_Click()
If bLoading = True Then Exit Sub
    Set pt = ows.PivotTables(1)
    Set pf = pt.PivotFields("ItemNote")
    Application.ScreenUpdating = False
    If ckbComments = True Then
        With pf
            .dataRange.EntireColumn.Hidden = True
            .dataRange.WrapText = False
        End With
        pt.PivotFields("Description").dataRange.ColumnWidth = 70
    Else
       With pf
            .dataRange.EntireColumn.Hidden = False
            .dataRange.WrapText = True
        End With
        pt.PivotFields("Description").dataRange.ColumnWidth = 40
    End If
    Application.ScreenUpdating = True
End Sub

Private Sub ckbPageBreak_Click()
If bLoading = True Then Exit Sub
If sPF = "" Then Exit Sub
    If ckbPageBreak = True Then
        pt.PivotFields(sPF).LayoutPageBreak = True
    Else
        pt.PivotFields(sPF).LayoutPageBreak = False
    End If
End Sub

Private Sub ckbSlicers_Click()
Dim Sh As Shape
Dim oSlicer As Slicer
Dim oSlicercache As SlicerCache
Dim X As Integer, C
If bLoading = True Then Exit Sub

Application.ScreenUpdating = False
    If ckbSlicers = True Then
        sRprt = ows.PivotTables(1).name
        Set lObj = Sheet0.ListObjects("tblRptTrack")
        Set C = lObj.ListColumns(1).DataBodyRange.Find(sRprt, LookIn:=xlValues)
        If Not C Is Nothing Then
            X = C.Offset(0, 1).Value
        End If
        For i = 1 To X
            Select Case i
                Case Is = 1
                    ActiveWorkbook.SlicerCaches.Add2(pt, sLvl1Item) _
                        .Slicers.Add ActiveSheet, , , sLvl1Item, 10, 900, 220, 400
                Case Is = 2
                    ActiveWorkbook.SlicerCaches.Add2(pt, sLvl2Item) _
                        .Slicers.Add ActiveSheet, , , sLvl2Item, 10, 1125, 220, 400
                Case Is = 3
                    ActiveWorkbook.SlicerCaches.Add2(pt, sLvl3Item) _
                    .Slicers.Add ActiveSheet, , , sLvl3Item, 120, 1000, 220, 400
                Case Is = 4
                    ActiveWorkbook.SlicerCaches.Add2(pt, sLvl4Item) _
                    .Slicers.Add ActiveSheet, , , sLvl4Item, 120, 1225, 220, 400
                Case Is = 5
                    ActiveWorkbook.SlicerCaches.Add2(pt, sLvl5Item) _
                    .Slicers.Add ActiveSheet, , , sLvl5Item, 190, 1100, 220, 400
                Case Else
            End Select
        Next i
    Else
        For Each Sh In ActiveSheet.Shapes
            If Sh.name <> "grpHeading" And Sh.name <> "grpHeadingVar" Then
                Sh.Delete
            End If
        Next
    End If
Application.ScreenUpdating = True
End Sub

Private Sub ckbSubtotals_Click()
If bLoading = True Then Exit Sub
If sPF = "" Then Exit Sub
    If ckbAddon = True Then
        clearAddons
        If ckbSubtotals = False Then
            pt.PivotFields(sPF).Subtotals = Array( _
            False, False, False, False, False, False, False, False, False, False, False, False)
        Else
            pt.PivotFields(sPF).Subtotals = Array( _
            True, False, False, False, False, False, False, False, False, False, False, False)
        End If
        ckbAddon_Click
    Else
        If ckbSubtotals = False Then
            pt.PivotFields(sPF).Subtotals = Array( _
            False, False, False, False, False, False, False, False, False, False, False, False)
        Else
            pt.PivotFields(sPF).Subtotals = Array( _
            True, False, False, False, False, False, False, False, False, False, False, False)
        End If
    End If
End Sub

Private Sub cmdCollapse_Click()
If bLoading = True Then Exit Sub
    On Error Resume Next
    If sPF = "" Then Exit Sub
    If ckbAddon = True Then
        clearAddons
        pt.PivotFields(sPF).ShowDetail = False
        ckbAddon_Click
    Else
        pt.PivotFields(sPF).ShowDetail = False
    End If
    On Error GoTo 0
End Sub

Private Sub cmdExpand_Click()
If bLoading = True Then Exit Sub
    On Error Resume Next
    If sPF = "" Then Exit Sub
    If ckbAddon = True Then
        clearAddons
        pt.PivotFields(sPF).ShowDetail = True
        ckbAddon_Click
    Else
        pt.PivotFields(sPF).ShowDetail = True
    End If
    On Error GoTo 0
End Sub



Private Sub ListBox1_Click()
    X = ListBox1.ListIndex
    Select Case X
        Case 0
            sPF = sLvl1Item
        Case 1
            sPF = sLvl2Item
        Case 2
            sPF = sLvl3Item
        Case 3
            sPF = sLvl4Item
        Case 4
            sPF = sLvl5Item
    End Select
    chkControls
End Sub

Private Sub txtHeading_Change()
    ows.Cells(1, 2).Value = StrConv(txtHeading.Value, vbUpperCase)
    txtHeading.Value = ows.Cells(1, 2).Value
End Sub

Private Sub UserForm_Activate()
    bLoading = True
    Set ows = ActiveSheet
    Set pt = ows.PivotTables(1)
    If InStr(1, pt.name, "XTab") > 0 Then
        sType = "XTab"
        ckbComments.Enabled = False
        Me.Width = 341
    ElseIf InStr(1, pt.name, "Control Estimate") > 0 Then
        sType = "CEst"
        ckbComments.Enabled = True
        Me.Width = 683
        Update_ChkBox
        
        
        
    ElseIf InStr(1, pt.name, "Variance") > 0 Then
        sType = "Var"
        ckbComments.Enabled = True
        Frame6.visible = True
        Frame3.visible = False
        Me.Width = 341
    Else
        sType = "Level"
        ckbComments.Enabled = True
        Me.Width = 341
    End If
    cmdExpand.Picture = Application.CommandBars.GetImageMso("RecordsExpandAllSubdatasheets", 22, 22)
    cmdCollapse.Picture = Application.CommandBars.GetImageMso("RecordsCollapseAllSubdatasheets", 22, 22)
    txtHeading.Value = ows.Cells(1, 2).Value
    Call loadForm
    bLoading = False
End Sub

Sub loadForm()
Dim oShape As Shape
Dim C, arr
    X = 0
    sRprt = ows.PivotTables(1).name
    Set lObj = Sheet0.ListObjects("tblRptTrack")
    Set C = lObj.ListColumns(1).DataBodyRange.Find(sRprt, LookIn:=xlValues)
    If Not C Is Nothing Then
        X = C.Offset(0, 1).Value
    End If
   For i = 1 To X
    Select Case i
        Case Is = 1
            If C.Offset(0, 10).Value = False Then
                arr = Split(C.Offset(0, 13).Value, "_")
                sLvl1Item = arr(0)
            Else
                sLvl1Item = C.Offset(0, 13).Value
            End If
            ListBox1.AddItem sLvl1Item
        Case Is = 2
            If C.Offset(0, 15).Value = False Then
                arr = Split(C.Offset(0, 18).Value, "_")
                sLvl2Item = arr(0)
            Else
                sLvl2Item = C.Offset(0, 18).Value
            End If
            ListBox1.AddItem sLvl2Item
        Case Is = 3
            If C.Offset(0, 20).Value = False Then
                arr = Split(C.Offset(0, 23).Value, "_")
                sLvl3Item = arr(0)
            Else
                sLvl3Item = C.Offset(0, 23).Value
            End If
            ListBox1.AddItem sLvl3Item
        Case Is = 4
            If C.Offset(0, 25).Value = False Then
                arr = Split(C.Offset(0, 28).Value, "_")
                sLvl4Item = arr(0)
            Else
                sLvl4Item = C.Offset(0, 28).Value
            End If
            ListBox1.AddItem sLvl4Item
        Case Is = 5
            If C.Offset(0, 30).Value = False Then
                arr = Split(C.Offset(0, 33).Value, "_")
                sLvl5Item = arr(0)
            Else
                sLvl5Item = C.Offset(0, 33).Value
            End If
            ListBox1.AddItem sLvl5Item
        Case Else
    End Select
    Next i
    
    With ows.PivotTables(1).TableRange1
        iCol = .Cells(.Cells.Count).Column
    End With
    If ows.Columns(iCol + 1).Hidden = False Then Me.chkVarComments.Value = True
    If ows.Columns(iCol - 3).Hidden = False Then Me.chkVarUnit.Value = True
    If ows.Columns(iCol - 4).Hidden = False Then Me.chkVarQty.Value = True
    For Each oShape In ows.Shapes
        If oShape.name <> "grpHeading" And oShape.name <> "grpHeadingVar" Then
             ckbSlicers.Value = True
             chkVarSlicers.Value = True
             Exit For
        End If
    Next oShape
    ckbAddon.Value = bAddon
    chkVarAddon.Value = bAddon
End Sub

Sub chkControls()
On Error Resume Next
    If pt.PivotFields(sPF).LayoutBlankLine = True Then
        ckbBlankLine = True
    Else
        ckbBlankLine = False
    End If
    
    If pt.PivotFields(sPF).LayoutPageBreak = True Then
        ckbPageBreak = True
    Else
        ckbPageBreak = False
    End If
    
    If pt.PivotFields(sPF).Subtotals(1) = True Then
        ckbSubtotals.Value = True
    Else
        ckbSubtotals.Value = False
    End If
    ckbAddon.Value = bAddon
    chkVarAddon.Value = bAddon
On Error GoTo 0
End Sub

Function bAddon() As Boolean
    With pt.TableRange1
        iGTRow = .Cells(.Cells.Count).row
    End With
    iRow = ActualUsedRange(ows).Rows.Count
    If iRow > iGTRow Then
        bAddon = True
    Else
        bAddon = False
    End If
End Function

Private Sub UserForm_Initialize()
    stopEvents = True
    Me.StartUpPosition = 0
    Me.Top = Application.Top + 115
    Me.Left = Application.Left + 25
End Sub

Sub Update_ChkBox()
    Dim dataRange As Range
    Dim grandTotalRow As Range
    Dim cell As Range
    Dim colIndex As Integer
    Dim sColName As String
    
    Set ows = ActiveSheet
    Set pt = ows.PivotTables(1)
    Set dataRange = pt.DataBodyRange
    
    If pt.TableRange2.Rows.Count > dataRange.Rows.Count Then
        Set grandTotalRow = dataRange.Rows(dataRange.Rows.Count)
        For colIndex = 1 To grandTotalRow.Columns.Count - 1
            Set cell = grandTotalRow.Cells(1, colIndex)
            On Error Resume Next
            sColName = pt.PivotFields(cell.Column).name
            On Error GoTo 0
            With Controls("CheckBox" & colIndex)
                If cell.EntireColumn.Hidden = True Then
                    .Value = False
                    .ForeColor = &H40C0&
                    .Font.Bold = False
                Else
                    .Value = True
                    .ForeColor = &H814901
                    .Font.Bold = True
                End If
            End With
        Next colIndex
    End If
End Sub

Function Col_Update(colIndex As Integer)
    Dim dataRange As Range
    Dim grandTotalRow As Range
    Dim cell As Range
    
    Set ows = ActiveSheet
    Set pt = ows.PivotTables(1)
    Set dataRange = pt.DataBodyRange
    
    If pt.TableRange2.Rows.Count > dataRange.Rows.Count Then
        Set grandTotalRow = dataRange.Rows(dataRange.Rows.Count)
        Set cell = grandTotalRow.Cells(1, colIndex)
        With Controls("CheckBox" & colIndex)
            If .Value = False Then
                .ForeColor = &H40C0&
                .Font.Bold = False
                cell.EntireColumn.Hidden = True
            Else
                .Value = True
                .ForeColor = &H814901
                .Font.Bold = True
                cell.EntireColumn.Hidden = False
            End If
        End With
    End If
End Function

Private Sub UserForm_Terminate()
    stopEvents = False
End Sub
