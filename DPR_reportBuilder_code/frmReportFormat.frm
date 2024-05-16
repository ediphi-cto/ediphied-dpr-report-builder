VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmReportFormat 
   Caption         =   "DPR Report Builder"
   ClientHeight    =   4800
   ClientLeft      =   48
   ClientTop       =   384
   ClientWidth     =   5268
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
Dim x, iRow
Dim C As Integer

Private Sub chkVarAddon_Click()
    Call ckbAddon_Click
End Sub

Private Sub chkVarComments_Click()
Dim iCol As Integer
    If bLoading = True Then Exit Sub
    With ows.UsedRange
        iCol = .Columns(.Columns.count).Column
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
Dim x As Integer, C
If bLoading = True Then Exit Sub

Application.ScreenUpdating = False
    If chkVarSlicers = True Then
        sRprt = ows.PivotTables(1).Name
        Set lObj = Sheet0.ListObjects("tblRptTrack")
        Set C = lObj.ListColumns(1).DataBodyRange.Find(sRprt, LookIn:=xlValues)
        If Not C Is Nothing Then
            x = C.Offset(0, 1).Value
        End If
        For i = 1 To x
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
            If Sh.Name <> "grpHeading" And Sh.Name <> "grpHeadingVar" Then
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
            .DataRange.EntireColumn.Hidden = True
            .DataRange.WrapText = False
        End With
        pt.PivotFields("Description").DataRange.ColumnWidth = 70
    Else
       With pf
            .DataRange.EntireColumn.Hidden = False
            .DataRange.WrapText = True
        End With
        pt.PivotFields("Description").DataRange.ColumnWidth = 40
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
Dim x As Integer, C
If bLoading = True Then Exit Sub

Application.ScreenUpdating = False
    If ckbSlicers = True Then
        sRprt = ows.PivotTables(1).Name
        Set lObj = Sheet0.ListObjects("tblRptTrack")
        Set C = lObj.ListColumns(1).DataBodyRange.Find(sRprt, LookIn:=xlValues)
        If Not C Is Nothing Then
            x = C.Offset(0, 1).Value
        End If
        For i = 1 To x
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
            If Sh.Name <> "grpHeading" And Sh.Name <> "grpHeadingVar" Then
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
    x = ListBox1.ListIndex
    Select Case x
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
    If InStr(1, pt.Name, "XTab") > 0 Then
        sType = "XTab"
        ckbComments.Enabled = False
    ElseIf InStr(1, pt.Name, "Control Estimate") > 0 Then
        sType = "CEst"
        ckbComments.Enabled = True
    ElseIf InStr(1, pt.Name, "Variance") > 0 Then
        sType = "Var"
        ckbComments.Enabled = True
        Frame6.visible = True
        Frame3.visible = False
    Else
        sType = "Level"
        ckbComments.Enabled = True
    End If
    cmdExpand.Picture = Application.CommandBars.GetImageMso("RecordsExpandAllSubdatasheets", 22, 22)
    cmdCollapse.Picture = Application.CommandBars.GetImageMso("RecordsCollapseAllSubdatasheets", 22, 22)
    txtHeading.Value = ows.Cells(1, 2).Value
    Call loadForm
    bLoading = False
End Sub

Sub loadForm()
Dim oShape As Shape
Dim C
    x = 0
    sRprt = ows.PivotTables(1).Name
    Set lObj = Sheet0.ListObjects("tblRptTrack")
    Set C = lObj.ListColumns(1).DataBodyRange.Find(sRprt, LookIn:=xlValues)
    If Not C Is Nothing Then
        x = C.Offset(0, 1).Value
    End If
    For i = 1 To x
    Select Case i
        Case Is = 1
            sLvl1Item = pt.PivotFields(2).Name
            ListBox1.AddItem pt.PivotFields(2).Caption
        Case Is = 2
            sLvl2Item = pt.PivotFields(4).Name
            ListBox1.AddItem pt.PivotFields(4).Caption
        Case Is = 3
            sLvl3Item = pt.PivotFields(6).Name
            ListBox1.AddItem pt.PivotFields(6).Caption
        Case Is = 4
            sLvl4Item = pt.PivotFields(8).Name
            ListBox1.AddItem pt.PivotFields(8).Caption
        Case Is = 5
            sLvl5Item = pt.PivotFields(10).Name
            ListBox1.AddItem pt.PivotFields(10).Caption
        Case Else
    End Select
    Next i
    
    With ows.PivotTables(1).TableRange1
        iCol = .Cells(.Cells.count).Column
    End With
    If ows.Columns(iCol + 1).Hidden = False Then Me.chkVarComments.Value = True
    If ows.Columns(iCol - 3).Hidden = False Then Me.chkVarUnit.Value = True
    If ows.Columns(iCol - 4).Hidden = False Then Me.chkVarQty.Value = True
    For Each oShape In ows.Shapes
        If oShape.Name <> "grpHeading" And oShape.Name <> "grpHeadingVar" Then
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
        iGTRow = .Cells(.Cells.count).row
    End With
    iRow = ActualUsedRange(ows).Rows.count
    If iRow > iGTRow Then
        bAddon = True
    Else
        bAddon = False
    End If
End Function

Private Sub UserForm_Initialize()
    Me.StartUpPosition = 0
    Me.Top = Application.Top + 115
    Me.Left = Application.Left + 25
End Sub

