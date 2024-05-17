VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmReportLevel 
   Caption         =   "DPR Report Builder"
   ClientHeight    =   6252
   ClientLeft      =   48
   ClientTop       =   360
   ClientWidth     =   9120
   OleObjectBlob   =   "frmReportLevel.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmReportLevel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public reportsRequested As Collection
'***************************
'*****Multipage1 Page 1*****
'***************************

Private Sub cmdOK_Click()
    On Error GoTo e1
'    sCurrency = Range("rngCurrency").Text
    eventsOff
    If isFirstReport() Then
        sSht = "Detailed Backup Report"
        sRprt = "DETAILED BACKUP"
    Else
        sSht = "Level Report"
        sRprt = "LEVEL REPORT"
    End If
    sGTLvl1 = cboBLvl1.List(cboBLvl1.ListIndex, 1)
    iLvl = numLevel.Value
    ReportTrack
    createSummary sGTLvl1
    ThisWorkbook.Worksheets("splash").visible = xlHidden
    ExecSummary
    Create_PivotTable_ODBC_MO
    ReApplyAddons
    
    Call clearStrings
    bCkb1 = False
    bCkb2 = False
    bCkb3 = False
    bCkb4 = False
    bCkb5 = False
    bCkbAll = False
    Range("rngIsTemp").Value = True
        
    Dim sortLabel As String
    If iLvl >= 1 Then sortLabel = sLvl1Name
    If iLvl >= 2 Then sortLabel = sortLabel & vbLf & "   -" & sLvl2Name
    If iLvl >= 3 Then sortLabel = sortLabel & vbLf & "      -" & sLvl3Name
    If iLvl >= 4 Then sortLabel = sortLabel & vbLf & "         -" & sLvl4Name
    If iLvl >= 5 Then sortLabel = sortLabel & vbLf & "            -" & sLvl5Name
    
    Dim post As New UserEvents
    post.slackPost "Initial Estimate Created" & vbLf & sortLabel
    
finally:
    eventsOn
    reportErrors
    Unload Me
    
Exit Sub
e1:
    logError "general failure creating primary report"
    Resume finally
    
End Sub

Private Sub cboBLvl1_Change()
    If cboBLvl1.ListIndex <> -1 Then
        cboBLvl1_Click
    End If
End Sub

Private Sub cboBLvl1_Click()
    With cboBLvl1
        If .Value = "WBS14 - Meta WBS2.1 (PRECON)" _
                Or .Value = "WBS13 - Meta WBS2.1 (OPS)" Then
            'sXpath1 = .List(.ListIndex, 0)
            sLvl1Name = .List(.ListIndex, 1)
            'sLvl1xNd = .List(.ListIndex, 2)
            sLvl1Item = .List(.ListIndex, 3)
            sLvl1Code = .List(.ListIndex, 4)
            numLevel.Value = 1
            cmdOK.Enabled = True
            ckbBLvl1.Value = False
            ckbBLvl1.Enabled = False
            Exit Sub
        End If
        If .Value <> "" Then
            If .Value = "Group" Or .Value = "Alternates" Then
                ckbBLvl1.Value = True
            Else
                ckbBLvl1.Value = False
            End If
            If InStr(.List(.ListIndex, 4), "code") = 0 Then
                ckbBLvl1.Enabled = False
            Else
                ckbBLvl1.Enabled = True
            End If
            'sXpath1 = .List(.ListIndex, 0)
            sLvl1Name = .List(.ListIndex, 1)
            'sLvl1xNd = .List(.ListIndex, 2)
            sLvl1Item = .List(.ListIndex, 3)
            sLvl1Code = .List(.ListIndex, 4)
            numLevel.Value = 1
            cmdOK.Enabled = True
            ckbBLvl2.Enabled = True
            ckbBLvl2.Value = False
            cboBLvl2.Enabled = True
            LoadCBO "cboBLvl2", "Page1"
        Else
            cmdOK.Enabled = False
            cboBLvl2.Enabled = False
            ckbBLvl2.Enabled = False
            ckbBLvl2.Value = False
        End If
    End With
End Sub

Private Sub cboBLvl2_Change()
    If cboBLvl2.ListIndex <> -1 Then
        cboBLvl2_Click
    End If
End Sub

Private Sub cboBLvl2_Click()
    With cboBLvl2
        If .Value = "WBS14 - Meta WBS2.1 (PRECON)" _
                Or .Value = "WBS13 - Meta WBS2.1 (OPS)" Then
            'sXpath2 = .List(.ListIndex, 0)
            sLvl2Name = .List(.ListIndex, 1)
            'sLvl2xNd = .List(.ListIndex, 2)
            sLvl2Item = .List(.ListIndex, 3)
            sLvl2Code = .List(.ListIndex, 4)
            numLevel.Value = 2
            cmdOK.Enabled = True
            ckbBLvl2.Value = False
            ckbBLvl2.Enabled = False
            Exit Sub
        End If
        If .Value <> "" Then
            If .Value = "Group" Or .Value = "Alternates" Then
                ckbBLvl2.Value = True
            Else
                ckbBLvl2.Value = False
            End If
            If InStr(.List(.ListIndex, 4), "code") = 0 Then
                ckbBLvl2.Enabled = False
            Else
                ckbBLvl2.Enabled = True
            End If
            'sXpath2 = .List(.ListIndex, 0)
            sLvl2Name = .List(.ListIndex, 1)
            'sLvl2xNd = .List(.ListIndex, 2)
            sLvl2Item = .List(.ListIndex, 3)
            sLvl2Code = .List(.ListIndex, 4)
            numLevel.Value = 2
            ckbBLvl3.Enabled = True
            ckbBLvl3.Value = False
            cboBLvl3.Enabled = True
            LoadCBO "cboBLvl3", "Page1"
        Else
            cboBLvl3.Enabled = False
            ckbBLvl3.Enabled = False
            ckbBLvl3.Value = False
        End If
    End With
End Sub

Private Sub cboBLvl3_Change()
    If cboBLvl3.ListIndex <> -1 Then
        cboBLvl3_Click
    End If
End Sub

Private Sub cboBLvl3_Click()
    With cboBLvl3
        If .Value = "WBS14 - Meta WBS2.1 (PRECON)" _
                Or .Value = "WBS13 - Meta WBS2.1 (OPS)" Then
            'sXpath3 = .List(.ListIndex, 0)
            sLvl3Name = .List(.ListIndex, 1)
            'sLvl3xNd = .List(.ListIndex, 2)
            sLvl3Item = .List(.ListIndex, 3)
            sLvl3Code = .List(.ListIndex, 4)
            numLevel.Value = 3
            cmdOK.Enabled = True
            ckbBLvl3.Value = False
            ckbBLvl3.Enabled = False
            Exit Sub
        End If
        If .Value <> "" Then
            If .Value = "Group" Or .Value = "Alternates" Then
                ckbBLvl3.Value = True
            Else
                ckbBLvl3.Value = False
            End If
            If InStr(.List(.ListIndex, 4), "code") = 0 Then
                ckbBLvl3.Enabled = False
            Else
                ckbBLvl3.Enabled = True
            End If
            'sXpath3 = .List(.ListIndex, 0)
            sLvl3Name = .List(.ListIndex, 1)
            'sLvl3xNd = .List(.ListIndex, 2)
            sLvl3Item = .List(.ListIndex, 3)
            sLvl3Code = .List(.ListIndex, 4)
            numLevel.Value = 3
            ckbBLvl4.Enabled = True
            ckbBLvl4.Value = False
            cboBLvl4.Enabled = True
            LoadCBO "cboBLvl4", "Page1"
        Else
            cboBLvl4.Enabled = False
            ckbBLvl4.Enabled = False
            ckbBLvl4.Value = False
        End If
    End With
End Sub

Private Sub cboBLvl4_Change()
    If Me.cboBLvl4.Value = "WBS14 - Meta WBS2.1 (PRECON)" _
                Or Me.cboBLvl4.Value = "WBS13 - Meta WBS2.1 (OPS)" Then
        MsgBox "This WBS Code cannot be used at this level", vbCritical, "META WBS 13-14"
        cboBLvl4.Value = ""
        Exit Sub
    End If
    If cboBLvl4.ListIndex <> -1 Then
        cboBLvl4_Click
    End If
End Sub

Private Sub cboBLvl4_Click()
    With cboBLvl4
        If .Value <> "" Then
            If .Value = "Group" Or .Value = "Alternates" Then
                ckbBLvl4.Value = True
            Else
                ckbBLvl4.Value = False
            End If
            If InStr(.List(.ListIndex, 4), "code") = 0 Then
                ckbBLvl4.Enabled = False
            Else
                ckbBLvl4.Enabled = True
            End If
            'sXpath4 = .List(.ListIndex, 0)
            sLvl4Name = .List(.ListIndex, 1)
            'sLvl4xNd = .List(.ListIndex, 2)
            sLvl4Item = .List(.ListIndex, 3)
            sLvl4Code = .List(.ListIndex, 4)
            numLevel.Value = 4
            ckbBLvl5.Enabled = True
            ckbBLvl5.Value = False
            cboBLvl5.Enabled = True
            LoadCBO "cboBLvl5", "Page1"
        Else
            cboBLvl5.Enabled = False
            ckbBLvl5.Enabled = False
            ckbBLvl5.Value = False
        End If
    End With
End Sub

Private Sub cboBLvl5_Change()
    If Me.cboBLvl5.Value = "WBS14 - Meta WBS2.1 (PRECON)" _
                Or Me.cboBLvl5.Value = "WBS13 - Meta WBS2.1 (OPS)" Then
        MsgBox "This WBS Code cannot be used at this level", vbCritical, "META WBS 13-14"
        cboBLvl5.Value = ""
        Exit Sub
    End If
    With cboBLvl5
        If .Value <> "" Then
            If .Value = "GroupPhase" Or .Value = "Alternates" Then
                ckbBLvl5.Value = True
            Else
                ckbBLvl5.Value = False
            End If
            If InStr(.List(.ListIndex, 4), "code") = 0 Then
                ckbBLvl5.Enabled = False
            Else
                ckbBLvl5.Enabled = True
            End If
            'sXpath5 = .List(.ListIndex, 0)
            sLvl5Name = .List(.ListIndex, 1)
            'sLvl5xNd = .List(.ListIndex, 2)
            sLvl5Item = .List(.ListIndex, 3)
            sLvl5Code = .List(.ListIndex, 4)
            numLevel.Value = 5
        End If
    End With
End Sub

Private Sub ckbBLvl1_Click()
    If ckbBLvl1.Value = True Then
        bCkb1 = True
    Else
        bCkb1 = False
    End If
End Sub

Private Sub ckbBLvl2_Click()
    If ckbBLvl2.Value = True Then
        bCkb2 = True
    Else
        bCkb2 = False
    End If
End Sub

Private Sub ckbBLvl3_Click()
    If ckbBLvl3.Value = True Then
        bCkb3 = True
    Else
        bCkb3 = False
    End If
End Sub

Private Sub ckbBLvl4_Click()
    If ckbBLvl4.Value = True Then
        bCkb4 = True
    Else
        bCkb4 = False
    End If
End Sub

Private Sub ckbBLvl5_Click()
    If ckbBLvl5.Value = True Then
        bCkb5 = True
    Else
        bCkb5 = False
    End If
End Sub

'Standard Level Report
'***************************
'*****Multipage1 Page 2*****
'***************************
Private Sub cmdOK1_Click()
    If thisReportBuilder Is Nothing Then
        On Error GoTo e1
    Else
        If thisReportBuilder.debugMode Then
            Stop
            On Error GoTo 0
        Else
            On Error GoTo e1
            eventsOff
        End If
    End If
    
'   sCurrency = Range("rngCurrency").Text
    'sSht = "Level Report - " & getPtCount
    'sRprt = sRprtName
    sGTLvl1 = cboLvl1.List(cboLvl1.ListIndex, 1)
    iLvl = numLevel.Value
    ReportTrack
    Create_PivotTable_ODBC_MO
    clearStrings
    bCkb1 = False
    bCkb2 = False
    bCkb3 = False
    bCkb4 = False
    bCkb5 = False
    bCkbAll = False
    
    
    Dim sortLabel As String
    If iLvl >= 1 Then sortLabel = sLvl1Name
    If iLvl >= 2 Then sortLabel = sortLabel & vbLf & "   -" & sLvl2Name
    If iLvl >= 3 Then sortLabel = sortLabel & vbLf & "      -" & sLvl3Name
    If iLvl >= 4 Then sortLabel = sortLabel & vbLf & "         -" & sLvl4Name
    If iLvl >= 5 Then sortLabel = sortLabel & vbLf & "            -" & sLvl5Name
    
    Dim post As New UserEvents
    post.slackPost "Level Report Successfully Created" & vbLf & sortLabel

finally:
    eventsOn
    reportErrors
    Unload Me
    
Exit Sub
e1:
    logError "Failed to create your report" '& vbLf & vbLf & TRY_UPDATING_MSG
    Resume finally
    
End Sub

Private Sub cboLvl1_Change()
    If cboLvl1.ListIndex <> -1 Then
        cboLvl1_Click
    End If
End Sub

Private Sub cboLvl1_Click()
    With cboLvl1
        If .Value = "WBS14 - Meta WBS2.1 (PRECON)" _
                Or .Value = "WBS13 - Meta WBS2.1 (OPS)" Then
            'sXpath1 = .List(.ListIndex, 0)
            sLvl1Name = .List(.ListIndex, 1)
            'sLvl1xNd = .List(.ListIndex, 2)
            sLvl1Item = .List(.ListIndex, 3)
            sLvl1Code = .List(.ListIndex, 4)
            numLevel.Value = 1
            cmdOK1.Enabled = True
            ckbLvl1.Value = False
            ckbLvl1.Enabled = False
            Exit Sub
        End If
        If .Value <> "" Then
            If cboLvl1.Value = "Group" Or cboLvl1.Value = "Alternates" Then
                ckbLvl1.Value = True
            Else
                ckbLvl1.Value = False
            End If
            If InStr(cboLvl1.List(.ListIndex, 4), "code") = 0 Then
                ckbLvl1.Enabled = False
            Else
                ckbLvl1.Enabled = True
            End If
            'sXpath1 = .List(.ListIndex, 0)
            sLvl1Name = .List(.ListIndex, 1)
            'sLvl1xNd = .List(.ListIndex, 2)
            sLvl1Item = .List(.ListIndex, 3)
            sLvl1Code = .List(.ListIndex, 4)
            numLevel.Value = 1
            cmdOK1.Enabled = True
            ckbLvl2.Enabled = True
            ckbLvl2.Value = False
            cboLvl2.Enabled = True
            LoadCBO "cboLvl2", "Page2"
        Else
            cmdOK1.Enabled = False
            cboLvl2.Enabled = False
            ckbLvl2.Enabled = False
            ckbLvl2.Value = False
        End If
    End With
End Sub

Private Sub cboLvl2_Change()
    If cboLvl2.ListIndex <> -1 Then
        cboLvl2_Click
    End If
End Sub

Private Sub cboLvl2_Click()
    With cboLvl2
        If .Value = "WBS14 - Meta WBS2.1 (PRECON)" _
                Or .Value = "WBS13 - Meta WBS2.1 (OPS)" Then
            'sXpath2 = .List(.ListIndex, 0)
            sLvl2Name = .List(.ListIndex, 1)
            'sLvl2xNd = .List(.ListIndex, 2)
            sLvl2Item = .List(.ListIndex, 3)
            sLvl2Code = .List(.ListIndex, 4)
            numLevel.Value = 2
            cmdOK1.Enabled = True
            ckbLvl2.Value = False
            ckbLvl2.Enabled = False
            Exit Sub
        End If
        If .Value <> "" Then
            If .Value = "Group" Or .Value = "Alternates" Then
                ckbLvl2.Value = True
            Else
                ckbLvl2.Value = False
            End If
            If InStr(cboLvl2.List(.ListIndex, 4), "code") = 0 Then
                ckbLvl2.Enabled = False
            Else
                ckbLvl2.Enabled = True
            End If
            'sXpath2 = .List(.ListIndex, 0)
            sLvl2Name = .List(.ListIndex, 1)
            'sLvl2xNd = .List(.ListIndex, 2)
            sLvl2Item = .List(.ListIndex, 3)
            sLvl2Code = .List(.ListIndex, 4)
            numLevel.Value = 2
            ckbLvl3.Enabled = True
            ckbLvl3.Value = False
            cboLvl3.Enabled = True
            LoadCBO "cboLvl3", "Page2"
        Else
            cboLvl3.Enabled = False
            ckbLvl3.Enabled = False
            ckbLvl3.Value = False
        End If
    End With
End Sub

Private Sub cboLvl3_Change()
    If cboLvl3.ListIndex <> -1 Then
        cboLvl3_Click
    End If
End Sub

Private Sub cboLvl3_Click()
    With cboLvl3
        If .Value = "WBS14 - Meta WBS2.1 (PRECON)" _
                Or .Value = "WBS13 - Meta WBS2.1 (OPS)" Then
            'sXpath3 = .List(.ListIndex, 0)
            sLvl3Name = .List(.ListIndex, 1)
            'sLvl3xNd = .List(.ListIndex, 2)
            sLvl3Item = .List(.ListIndex, 3)
            sLvl3Code = .List(.ListIndex, 4)
            numLevel.Value = 3
            cmdOK1.Enabled = True
            ckbLvl3.Value = False
            ckbLvl3.Enabled = False
            Exit Sub
        End If
        If .Value <> "" Then
            If .Value = "Group" Or .Value = "Alternates" Then
                ckbLvl3.Value = True
            Else
                ckbLvl3.Value = False
            End If
            If InStr(cboLvl3.List(.ListIndex, 4), "code") = 0 Then
                ckbLvl3.Enabled = False
            Else
                ckbLvl3.Enabled = True
            End If
            'sXpath3 = .List(.ListIndex, 0)
            sLvl3Name = .List(.ListIndex, 1)
            'sLvl3xNd = .List(.ListIndex, 2)
            sLvl3Item = .List(.ListIndex, 3)
            sLvl3Code = .List(.ListIndex, 4)
            numLevel.Value = 3
            ckbLvl4.Enabled = True
            ckbLvl4.Value = False
            cboLvl4.Enabled = True
            LoadCBO "cboLvl4", "Page2"
        Else
            cboLvl4.Enabled = False
            ckbLvl4.Enabled = False
            ckbLvl4.Value = False
        End If
    End With
End Sub

Private Sub cboLvl4_Change()
    If Me.cboLvl4.Value = "WBS14 - Meta WBS2.1 (PRECON)" _
                Or Me.cboLvl4.Value = "WBS13 - Meta WBS2.1 (OPS)" Then
        MsgBox "This WBS Code cannot be used at this level", vbCritical, "META WBS 13-14"
        cboLvl4.Value = ""
        Exit Sub
    End If
    If cboLvl4.ListIndex <> -1 Then
        cboLvl4_Click
    End If
End Sub

Private Sub cboLvl4_Click()
    With cboLvl4
        If .Value <> "" Then
            If .Value = "Group" Or .Value = "Alternates" Then
                ckbLvl4.Value = True
            Else
                ckbLvl4.Value = False
            End If
            If InStr(cboLvl4.List(.ListIndex, 4), "code") = 0 Then
                ckbLvl4.Enabled = False
            Else
                ckbLvl4.Enabled = True
            End If
            'sXpath4 = .List(.ListIndex, 0)
            sLvl4Name = .List(.ListIndex, 1)
            'sLvl4xNd = .List(.ListIndex, 2)
            sLvl4Item = .List(.ListIndex, 3)
            sLvl4Code = .List(.ListIndex, 4)
            numLevel.Value = 4
            ckbLvl5.Enabled = True
            ckbLvl5.Value = False
            cboLvl5.Enabled = True
            LoadCBO "cboLvl5", "Page2"
        Else
            cboLvl5.Enabled = False
            ckbLvl5.Enabled = False
            ckbLvl5.Value = False
        End If
    End With
End Sub

Private Sub cboLvl5_Change()
    With cboLvl5
        If cboLvl5.Value <> "" Then
            If Me.cboLvl5.Value = "WBS14 - Meta WBS2.1 (PRECON)" _
                Or Me.cboLvl5.Value = "WBS13 - Meta WBS2.1 (OPS)" Then
                MsgBox "This WBS Code cannot be used at this level", vbCritical, "META WBS 13-14"
                .Value = ""
                Exit Sub
            End If
            If .Value = "GroupPhase" Or .Value = "Alternates" Then
                ckbLvl5.Value = True
            Else
                ckbLvl5.Value = False
            End If
            If InStr(cboLvl5.List(.ListIndex, 4), "code") = 0 Then
                ckbLvl5.Enabled = False
            Else
                ckbLvl5.Enabled = True
            End If
            'sXpath5 = .List(.ListIndex, 0)
            sLvl5Name = .List(.ListIndex, 1)
            'sLvl5xNd = .List(.ListIndex, 2)
            sLvl5Item = .List(.ListIndex, 3)
            sLvl5Code = .List(.ListIndex, 4)
            numLevel.Value = 5
        End If
    End With
End Sub

Private Sub ckbLvl1_Click()
    If ckbLvl1.Value = True Then
        bCkb1 = True
    Else
        bCkb1 = False
    End If
End Sub

Private Sub ckbLvl2_Click()
    If ckbLvl2.Value = True Then
        bCkb2 = True
    Else
        bCkb2 = False
    End If
End Sub

Private Sub ckbLvl3_Click()
    If ckbLvl3.Value = True Then
        bCkb3 = True
    Else
        bCkb3 = False
    End If
End Sub

Private Sub ckbLvl4_Click()
    If ckbLvl4.Value = True Then
        bCkb4 = True
    Else
        bCkb4 = False
    End If
End Sub

Private Sub ckbLvl5_Click()
    If ckbLvl5.Value = True Then
        bCkb5 = True
    Else
        bCkb5 = False
    End If
End Sub

Private Sub chkGrouping_Click()
    If chkGrouping.Value = True Then
        bCkbAll = True
    Else
        bCkbAll = False
    End If
End Sub

Private Sub chkSubs_Click()
    If chkSubs.Value = True Then
        bCkbSub = True
    Else
        bCkbSub = False
    End If
End Sub

'CONTROL ESTIMATE
'**********************************************************************************************************************************************
'*****Multipage1 Page 3************************************************************************************************************************
'**********************************************************************************************************************************************

Private Sub cmdOK3_Click()
'sCurrency = Range("rngCurrency").Text

    sSht = "Control Estimate - " & getPtCount
    sRprt = sCRprtName
    sGTLvl1 = cboCLvl1.List(cboCLvl1.ListIndex, 1)
    iLvl = numLevel.Value
    
    'MN TODO:
    ReportTrack
    
'Build pivot report
    Call Create_PivotTable_ODBC_CntrlEst
    Call clearStrings
    
    bCkb1 = False
    bCkb2 = False
    bCkb3 = False
    bCkb4 = False
    bCkb5 = False
    bCkbAll = False
    Unload Me
    MsgBox "Report Complete", vbOKOnly, "DPR Report Builder"
End Sub

Private Sub cboCLvl1_Change()
    If cboCLvl1.ListIndex <> -1 Then
        cboCLvl1_Click
    End If
End Sub

Private Sub cboCLvl1_Click()
    With cboCLvl1
        If .Value = "WBS14 - Meta WBS2.1 (PRECON)" _
            Or .Value = "WBS13 - Meta WBS2.1 (OPS)" Then
            'sXpath1 = .List(.ListIndex, 0)
            sLvl1Name = .List(.ListIndex, 1)
            'sLvl1xNd = .List(.ListIndex, 2)
            sLvl1Item = .List(.ListIndex, 3)
            sLvl1Code = .List(.ListIndex, 4)
            numLevel.Value = 1
            cmdOK3.Enabled = True
            ckbCLvl1.Value = False
            ckbCLvl1.Enabled = False
            Exit Sub
        End If
        If .Value <> "" Then
            If cboCLvl1.Value = "Group" Then
                ckbCLvl1.Value = True
            Else
                ckbCLvl1.Value = False
            End If
            If InStr(cboCLvl1.List(.ListIndex, 4), "code") = 0 Then
                ckbCLvl1.Enabled = False
            Else
                ckbCLvl1.Enabled = True
            End If
            'sXpath1 = .List(.ListIndex, 0)
            sLvl1Name = .List(.ListIndex, 1)
            'sLvl1xNd = .List(.ListIndex, 2)
            sLvl1Item = .List(.ListIndex, 3)
            sLvl1Code = .List(.ListIndex, 4)
            numLevel.Value = 1
            cmdOK3.Enabled = True
            ckbCLvl2.Enabled = True
            ckbCLvl2.Value = False
            cboCLvl2.Enabled = True
            LoadCBO "cboCLvl2", "Page3"
        Else
            cmdOK3.Enabled = False
            cboCLvl2.Enabled = False
            ckbCLvl2.Enabled = False
            ckbCLvl2.Value = False
        End If
    End With
End Sub

Private Sub cboCLvl2_Change()
    If cboCLvl2.ListIndex <> -1 Then
        cboCLvl2_Click
    End If
End Sub

Private Sub cboCLvl2_Click()
    With cboCLvl2
        If .Value = "WBS14 - Meta WBS2.1 (PRECON)" _
            Or .Value = "WBS13 - Meta WBS2.1 (OPS)" Then
            'sXpath2 = .List(.ListIndex, 0)
            sLvl2Name = .List(.ListIndex, 1)
            'sLvl2xNd = .List(.ListIndex, 2)
            sLvl2Item = .List(.ListIndex, 3)
            sLvl2Code = .List(.ListIndex, 4)
            numLevel.Value = 2
            cmdOK3.Enabled = True
            ckbCLvl2.Value = False
            ckbCLvl2.Enabled = False
            Exit Sub
        End If
        If .Value <> "" Then
            If .Value = "Group" Then
                ckbCLvl2.Value = True
            Else
                ckbCLvl2.Value = False
            End If
            If InStr(cboCLvl2.List(.ListIndex, 4), "code") = 0 Then
                ckbCLvl2.Enabled = False
            Else
                ckbCLvl2.Enabled = True
            End If
            'sXpath2 = .List(.ListIndex, 0)
            sLvl2Name = .List(.ListIndex, 1)
            'sLvl2xNd = .List(.ListIndex, 2)
            sLvl2Item = .List(.ListIndex, 3)
            sLvl2Code = .List(.ListIndex, 4)
            numLevel.Value = 2
            ckbCLvl3.Enabled = True
            ckbCLvl3.Value = False
            cboCLvl3.Enabled = True
            LoadCBO "cboCLvl3", "Page3"
        Else
            cboCLvl3.Enabled = False
            ckbCLvl3.Enabled = False
            ckbCLvl3.Value = False
        End If
    End With
End Sub

Private Sub cboCLvl3_Change()
    If cboCLvl3.ListIndex <> -1 Then
        cboCLvl3_Click
    End If
End Sub

Private Sub cboCLvl3_Click()
    With cboCLvl3
        If .Value = "WBS14 - Meta WBS2.1 (PRECON)" _
            Or .Value = "WBS13 - Meta WBS2.1 (OPS)" Then
            'sXpath3 = .List(.ListIndex, 0)
            sLvl3Name = .List(.ListIndex, 1)
            'sLvl3xNd = .List(.ListIndex, 2)
            sLvl3Item = .List(.ListIndex, 3)
            sLvl3Code = .List(.ListIndex, 4)
            numLevel.Value = 3
            cmdOK3.Enabled = True
            ckbCLvl3.Value = False
            ckbCLvl3.Enabled = False
            Exit Sub
        End If
        If .Value <> "" Then
            If .Value = "Group" Then
                ckbCLvl3.Value = True
            Else
                ckbCLvl3.Value = False
            End If
            If InStr(cboCLvl3.List(.ListIndex, 4), "code") = 0 Then
                ckbCLvl3.Enabled = False
            Else
                ckbCLvl3.Enabled = True
            End If
            'sXpath3 = .List(.ListIndex, 0)
            sLvl3Name = .List(.ListIndex, 1)
            'sLvl3xNd = .List(.ListIndex, 2)
            sLvl3Item = .List(.ListIndex, 3)
            sLvl3Code = .List(.ListIndex, 4)
            numLevel.Value = 3
            ckbCLvl4.Enabled = True
            ckbCLvl4.Value = False
            cboCLvl4.Enabled = True
            LoadCBO "cboCLvl4", "Page3"
        Else
            cboCLvl4.Enabled = False
            ckbCLvl4.Enabled = False
            ckbCLvl4.Value = False
        End If
    End With
End Sub

Private Sub cboCLvl4_Change()
    If cboCLvl4.Value = "WBS14 - Meta WBS2.1 (PRECON)" _
            Or cboCLvl4.Value = "WBS13 - Meta WBS2.1 (OPS)" Then
        MsgBox "This WBS Code cannot be used at this level", vbCritical, "META WBS 13-14"
        cboCLvl4.Value = ""
        Exit Sub
    End If
    If cboCLvl4.ListIndex <> -1 Then
        cboCLvl4_Click
    End If
End Sub

Private Sub cboCLvl4_Click()
    With cboCLvl4
        If .Value <> "" Then
            If .Value = "Group" Then
                ckbCLvl4.Value = True
            Else
                ckbCLvl4.Value = False
            End If
            If InStr(cboCLvl4.List(.ListIndex, 4), "code") = 0 Then
                ckbCLvl4.Enabled = False
            Else
                ckbCLvl4.Enabled = True
            End If
            'sXpath4 = .List(.ListIndex, 0)
            sLvl4Name = .List(.ListIndex, 1)
            'sLvl4xNd = .List(.ListIndex, 2)
            sLvl4Item = .List(.ListIndex, 3)
            sLvl4Code = .List(.ListIndex, 4)
            numLevel.Value = 4
            ckbCLvl5.Enabled = True
            ckbCLvl5.Value = False
            cboCLvl5.Enabled = True
            LoadCBO "cboCLvl5", "Page3"
        Else
            cboCLvl5.Enabled = False
            ckbCLvl5.Enabled = False
            ckbCLvl5.Value = False
        End If
    End With
End Sub

Private Sub cboCLvl5_Change()
    With cboCLvl5
        If cboCLvl5.Value = "WBS14 - Meta WBS2.1 (PRECON)" _
            Or cboCLvl5.Value = "WBS13 - Meta WBS2.1 (OPS)" Then
            MsgBox "This WBS Code cannot be used at this level", vbCritical, "META WBS 13-14"
            .Value = ""
            Exit Sub
        End If
        If cboCLvl5.Value <> "" Then
            If .Value = "GroupPhase" Then
                ckbCLvl5.Value = True
            Else
                ckbCLvl5.Value = False
            End If
            If InStr(cboCLvl5.List(.ListIndex, 4), "code") = 0 Then
                ckbCLvl5.Enabled = False
            Else
                ckbCLvl5.Enabled = True
            End If
            'sXpath5 = .List(.ListIndex, 0)
            sLvl5Name = .List(.ListIndex, 1)
            'sLvl5xNd = .List(.ListIndex, 2)
            sLvl5Item = .List(.ListIndex, 3)
            sLvl5Code = .List(.ListIndex, 4)
            numLevel.Value = 5
        End If
    End With
End Sub
Private Sub ckbCLvl1_Click()
    If ckbCLvl1.Value = True Then
        bCkb1 = True
    Else
        bCkb1 = False
    End If
End Sub

Private Sub ckbCLvl2_Click()
    If ckbCLvl2.Value = True Then
        bCkb2 = True
    Else
        bCkb2 = False
    End If
End Sub

Private Sub ckbCLvl3_Click()
    If ckbCLvl3.Value = True Then
        bCkb3 = True
    Else
        bCkb3 = False
    End If
End Sub

Private Sub ckbCLvl4_Click()
    If ckbCLvl4.Value = True Then
        bCkb4 = True
    Else
        bCkb4 = False
    End If
End Sub

Private Sub ckbCLvl5_Click()
    If ckbCLvl5.Value = True Then
        bCkb5 = True
    Else
        bCkb5 = False
    End If
End Sub

Private Sub chkGrouping3_Click()
    If chkGrouping3.Value = True Then
        bCkbAll = True
    Else
        bCkbAll = False
    End If
End Sub

'CROSSTAB REPORT
'***************************
'*****Multipage1 Page 4*****
'***************************
Private Sub cmdOK2_Click()
'sCurrency = Range("rngCurrency").Text
    sSht = "XTab Report - " & getPtCount
    sRprt = sXRprtName
    sXTRow = sRowName
    sGTLvl1 = cboXLvl1.List(cboXLvl1.ListIndex, 1)
    iLvl = numXLevel.Value
    ReportTrack
    'MN TODO:
    
'Build pivot report
    Create_PivotTable_ODBC_XT
    Call clearStrings
    ckbXLvl1 = False
    ckbXLvl2 = False
    ckbXLvl3 = False
    ckbXLvl4 = False
    ckbXLvl5 = False
    Unload Me
    MsgBox "Report Complete", vbOKOnly, "DPR Report Builder"
End Sub

Private Sub cboLvl0_Change()
    If cboLvl0.Value = "WBS14 - Meta WBS2.1 (PRECON)" _
            Or cboLvl0.Value = "WBS13 - Meta WBS2.1 (OPS)" Then
        MsgBox "This WBS Code cannot be used at this level", vbCritical, "META WBS 13-14"
        cboLvl0.Value = ""
        Exit Sub
    End If
    If cboLvl0.ListIndex <> -1 Then
        cboLvl0_Click
    End If
End Sub

Private Sub cboLvl0_Click()
    With cboLvl0
        If .Value <> "" Then
            If cboLvl0.Value = "Group" Then
                ckbLvl0.Value = True
            Else
                ckbLvl0.Value = False
            End If
            If InStr(.List(.ListIndex, 4), "code") = 0 Then
                ckbLvl0.Enabled = False
            Else
                ckbLvl0.Enabled = True
            End If
            'sXpath0 = .List(.ListIndex, 0)
            sLvl0Name = .List(.ListIndex, 1)
            'sLvl0xNd = .List(.ListIndex, 2)
            'sLvl0Item = .List(.ListIndex, 3)
            sLvl0Code = .List(.ListIndex, 4)
            ckbXLvl1.Enabled = True
            ckbXLvl1.Value = False
            cboXLvl1.Enabled = True
            LoadCBO "cboXLvl1", "Page4"
        Else
            cmdOK2.Enabled = False
            cboXLvl1.Enabled = False
            ckbXLvl1.Enabled = False
            ckbXLvl1.Value = False
        End If
    End With
End Sub

Private Sub cboXLvl1_Change()
    If cboXLvl1.ListIndex <> -1 Then
        cboXLvl1_Click
    End If
End Sub

Private Sub cboXLvl1_Click()
    With cboXLvl1
        If .Value = "WBS14 - Meta WBS2.1 (PRECON)" _
            Or .Value = "WBS13 - Meta WBS2.1 (OPS)" Then
            'sXpath1 = .List(.ListIndex, 0)
            sLvl1Name = .List(.ListIndex, 1)
            'sLvl1xNd = .List(.ListIndex, 2)
            sLvl1Item = .List(.ListIndex, 3)
            sLvl1Code = .List(.ListIndex, 4)
            numXLevel.Value = 1
            cmdOK2.Enabled = True
            ckbXLvl1.Value = False
            ckbXLvl1.Enabled = False
            Exit Sub
        End If
        If .Value <> "" Then
            If .Value = "Group" Then
                ckbXLvl1.Value = True
            Else
                ckbXLvl1.Value = False
            End If
            If InStr(.List(.ListIndex, 4), "code") = 0 Then
                ckbXLvl1.Enabled = False
            Else
                ckbXLvl1.Enabled = True
            End If
            'sXpath1 = .List(.ListIndex, 0)
            sLvl1Name = .List(.ListIndex, 1)
            'sLvl1xNd = .List(.ListIndex, 2)
            sLvl1Item = .List(.ListIndex, 3)
            sLvl1Code = .List(.ListIndex, 4)
            numXLevel.Value = 1
            cmdOK2.Enabled = True
            ckbXLvl2.Enabled = True
            ckbXLvl2.Value = False
            cboXLvl2.Enabled = True
            LoadCBO "cboXLvl2", "Page4"
        Else
            cmdOK2.Enabled = False
            cboXLvl2.Enabled = False
            ckbXLvl2.Enabled = False
            ckbXLvl2.Value = False
        End If
    End With
End Sub

Private Sub cboXLvl2_Change()
    If cboXLvl2.ListIndex <> -1 Then
        cboXLvl2_Click
    End If
End Sub

Private Sub cboXLvl2_Click()
    With cboXLvl2
        If .Value = "WBS14 - Meta WBS2.1 (PRECON)" _
            Or .Value = "WBS13 - Meta WBS2.1 (OPS)" Then
            'sXpath2 = .List(.ListIndex, 0)
            sLvl2Name = .List(.ListIndex, 1)
            'sLvl2xNd = .List(.ListIndex, 2)
            sLvl2Item = .List(.ListIndex, 3)
            sLvl2Code = .List(.ListIndex, 4)
            numXLevel.Value = 2
            cmdOK2.Enabled = True
            ckbXLvl2.Value = False
            ckbXLvl2.Enabled = False
            Exit Sub
        End If
        If .Value <> "" Then
            If .Value = "Group" Then
                ckbXLvl2.Value = True
            Else
                ckbXLvl2.Value = False
            End If
            If InStr(.List(.ListIndex, 4), "code") = 0 Then
                ckbXLvl2.Enabled = False
            Else
                ckbXLvl2.Enabled = True
            End If
            'sXpath2 = .List(.ListIndex, 0)
            sLvl2Name = .List(.ListIndex, 1)
            'sLvl2xNd = .List(.ListIndex, 2)
            sLvl2Item = .List(.ListIndex, 3)
            sLvl2Code = .List(.ListIndex, 4)
            numXLevel.Value = 2
            ckbXLvl3.Enabled = True
            ckbXLvl3.Value = False
            cboXLvl3.Enabled = True
            LoadCBO "cboXLvl3", "Page4"
        Else
            cboXLvl3.Enabled = False
            ckbXLvl3.Enabled = False
            ckbXLvl3.Value = False
        End If
    End With
End Sub

Private Sub cboXLvl3_Change()
    If cboXLvl3.ListIndex <> -1 Then
        cboXLvl3_Click
    End If
End Sub

Private Sub cboXLvl3_Click()
    With cboXLvl3
        If .Value = "WBS14 - Meta WBS2.1 (PRECON)" _
            Or .Value = "WBS13 - Meta WBS2.1 (OPS)" Then
            'sXpath3 = .List(.ListIndex, 0)
            sLvl3Name = .List(.ListIndex, 1)
            'sLvl3xNd = .List(.ListIndex, 2)
            sLvl3Item = .List(.ListIndex, 3)
            sLvl3Code = .List(.ListIndex, 4)
            numXLevel.Value = 3
            cmdOK2.Enabled = True
            ckbXLvl3.Value = False
            ckbXLvl3.Enabled = False
            Exit Sub
        End If
        If .Value <> "" Then
            If .Value = "Group" Then
                ckbXLvl3.Value = True
            Else
                ckbXLvl3.Value = False
            End If
            If InStr(.List(.ListIndex, 4), "code") = 0 Then
                ckbXLvl3.Enabled = False
            Else
                ckbXLvl3.Enabled = True
            End If
            'sXpath3 = .List(.ListIndex, 0)
            sLvl3Name = .List(.ListIndex, 1)
            'sLvl3xNd = .List(.ListIndex, 2)
            sLvl3Item = .List(.ListIndex, 3)
            sLvl3Code = .List(.ListIndex, 4)
            numXLevel.Value = 3
            ckbXLvl4.Enabled = True
            ckbXLvl4.Value = False
            cboXLvl4.Enabled = True
            LoadCBO "cboXLvl4", "Page4"
        Else
            cboXLvl4.Enabled = False
            ckbXLvl4.Enabled = False
            ckbXLvl4.Value = False
        End If
    End With
End Sub

Private Sub cboXLvl4_Change()
    If cboXLvl4.Value = "WBS14 - Meta WBS2.1 (PRECON)" _
            Or cboXLvl4.Value = "WBS13 - Meta WBS2.1 (OPS)" Then
        MsgBox "This WBS Code cannot be used at this level", vbCritical, "META WBS 13-14"
        cboXLvl4.Value = ""
        Exit Sub
    End If
    If cboXLvl4.ListIndex <> -1 Then
        cboXLvl4_Click
    End If
End Sub

Private Sub cboXLvl4_Click()
    With cboXLvl4
        If .Value <> "" Then
            If .Value = "Group" Then
                ckbXLvl4.Value = True
            Else
                ckbXLvl4.Value = False
            End If
            If InStr(.List(.ListIndex, 4), "code") = 0 Then
                ckbXLvl4.Enabled = False
            Else
                ckbXLvl4.Enabled = True
            End If
            'sXpath4 = .List(.ListIndex, 0)
            sLvl4Name = .List(.ListIndex, 1)
            'sLvl4xNd = .List(.ListIndex, 2)
            sLvl4Item = .List(.ListIndex, 3)
            sLvl4Code = .List(.ListIndex, 4)
            numXLevel.Value = 4
            ckbXLvl5.Enabled = True
            ckbXLvl5.Value = False
            cboXLvl5.Enabled = True
            LoadCBO "cboXLvl5", "Page4"
        Else
            cboXLvl5.Enabled = False
            ckbXLvl5.Enabled = False
            ckbXLvl5.Value = False
        End If
    End With
End Sub

Private Sub cboXLvl5_Change()
    If cboXLvl5.Value = "WBS14 - Meta WBS2.1 (PRECON)" _
            Or cboXLvl5.Value = "WBS13 - Meta WBS2.1 (OPS)" Then
        MsgBox "This WBS Code cannot be used at this level", vbCritical, "META WBS 13-14"
        cboXLvl5.Value = ""
        Exit Sub
    End If
    With cboXLvl5
        If .Value <> "" Then
            If .Value = "Group" Then
                ckbXLvl5.Value = True
            Else
                ckbXLvl5.Value = False
            End If
            If InStr(.List(.ListIndex, 4), "code") = 0 Then
                ckbXLvl5.Enabled = False
            Else
                ckbXLvl5.Enabled = True
            End If
            'sXpath5 = .List(.ListIndex, 0)
            sLvl5Name = .List(.ListIndex, 1)
            'sLvl5xNd = .List(.ListIndex, 2)
            sLvl5Item = .List(.ListIndex, 3)
            sLvl5Code = .List(.ListIndex, 4)
            numXLevel.Value = 5
        End If
    End With
End Sub

Private Sub ckbLvl0_Click()
    If ckbLvl0.Value = True Then
        bCkb0 = True
    Else
        bCkb0 = False
    End If
End Sub

Private Sub ckbXLvl1_Click()
    If ckbXLvl1.Value = True Then
        bCkb1 = True
    Else
        bCkb1 = False
    End If
End Sub

Private Sub ckbXLvl2_Click()
    If ckbXLvl1.Value = True Then
        bCkb2 = True
    Else
        bCkb2 = False
    End If
End Sub

Private Sub ckbXLvl3_Click()
    If ckbXLvl3.Value = True Then
        bCkb3 = True
    Else
        bCkb3 = False
    End If
End Sub

Private Sub ckbXLvl4_Click()
    If ckbXLvl4.Value = True Then
        bCkb4 = True
    Else
        bCkb4 = False
    End If
End Sub

Private Sub ckbXLvl5_Click()
    If ckbXLvl5.Value = True Then
        bCkb5 = True
    Else
        bCkb5 = False
    End If
End Sub

'Variance Report
'***************************
'*****Multipage1 Page 5*****
'***************************

Private Sub cmdOK4_Click()
'sCurrency = Range("rngCurrency").Text
'    xmlVarPath = Application.ThisWorkbook.Path & "\ReportData" & Range("rngVarReport").value
    sSht = "Variance Report - " & getPtCount
    sRprt = sRprtNameVar
    sGTLvl1 = cboLvl1.List(cboVLvl1.ListIndex, 1)
    iLvl = numLevel.Value
    sTotal = "GrandTotal"
    bMarkups = True
    'MN TODO:
    
    ReportTrack
    
    Call clearStrings
    bCkb1 = False
    bCkb2 = False
    bCkb3 = False
    bCkb4 = False
    bCkb5 = False
    bCkbAll = False
    Unload Me
   
End Sub

Private Sub cboVLvl1_Change()
    If cboVLvl1.ListIndex <> -1 Then
        cboVLvl1_Click
    End If
End Sub

Private Sub cboVLvl1_Click()
    With cboVLvl1
        If .Value <> "" Then
            If cboVLvl1.Value = "Group" Or cboVLvl1.Value = "Alternates" Then
                ckbVLvl1.Value = True
            Else
                ckbVLvl1.Value = False
            End If
            If InStr(cboVLvl1.List(.ListIndex, 4), "code") = 0 Then
                ckbVLvl1.Enabled = False
            Else
                ckbVLvl1.Enabled = True
            End If
            'sXpath1 = .List(.ListIndex, 0)
            sLvl1Name = .List(.ListIndex, 1)
            'sLvl1xNd = .List(.ListIndex, 2)
            sLvl1Item = .List(.ListIndex, 3)
            sLvl1Code = .List(.ListIndex, 4)
            numLevel.Value = 1
            cmdOK4.Enabled = True
            ckbVLvl2.Enabled = True
            ckbVLvl2.Value = False
            cboVLvl2.Enabled = True
            LoadCBO "cboVLvl2", "Page5"
        Else
            cmdOK4.Enabled = False
            cboVLvl2.Enabled = False
            ckbVLvl2.Enabled = False
            ckbVLvl2.Value = False
        End If
    End With
End Sub

Private Sub cboVLvl2_Change()
    If cboVLvl2.ListIndex <> -1 Then
        cboVLvl2_Click
    End If
End Sub

Private Sub cboVLvl2_Click()
    With cboVLvl2
        If .Value <> "" Then
            If .Value = "Group" Or .Value = "Alternates" Then
                ckbVLvl2.Value = True
            Else
                ckbVLvl2.Value = False
            End If
            If InStr(cboVLvl2.List(.ListIndex, 4), "code") = 0 Then
                ckbVLvl2.Enabled = False
            Else
                ckbVLvl2.Enabled = True
            End If
            'sXpath2 = .List(.ListIndex, 0)
            sLvl2Name = .List(.ListIndex, 1)
            'sLvl2xNd = .List(.ListIndex, 2)
            sLvl2Item = .List(.ListIndex, 3)
            sLvl2Code = .List(.ListIndex, 4)
            numLevel.Value = 2
            ckbVLvl3.Enabled = True
            ckbVLvl3.Value = False
            cboVLvl3.Enabled = True
            LoadCBO "cboVLvl3", "Page5"
        Else
            cboVLvl3.Enabled = False
            ckbVLvl3.Enabled = False
            ckbVLvl3.Value = False
        End If
    End With
End Sub

Private Sub cboVLvl3_Change()
    If cboVLvl3.ListIndex <> -1 Then
        cboVLvl3_Click
    End If
End Sub

Private Sub cboVLvl3_Click()
    With cboVLvl3
        If .Value <> "" Then
            If .Value = "Group" Or .Value = "Alternates" Then
                ckbVLvl3.Value = True
            Else
                ckbVLvl3.Value = False
            End If
            If InStr(cboVLvl3.List(.ListIndex, 4), "code") = 0 Then
                ckbVLvl3.Enabled = False
            Else
                ckbVLvl3.Enabled = True
            End If
            'sXpath3 = .List(.ListIndex, 0)
            sLvl3Name = .List(.ListIndex, 1)
            'sLvl3xNd = .List(.ListIndex, 2)
            sLvl3Item = .List(.ListIndex, 3)
            sLvl3Code = .List(.ListIndex, 4)
            numLevel.Value = 3
            ckbVLvl4.Enabled = True
            ckbVLvl4.Value = False
            cboVLvl4.Enabled = True
            LoadCBO "cboVLvl4", "Page5"
        Else
            cboVLvl4.Enabled = False
            ckbVLvl4.Enabled = False
            ckbVLvl4.Value = False
        End If
    End With
End Sub

Private Sub cboVLvl4_Change()
    If cboVLvl4.ListIndex <> -1 Then
        cboVLvl4_Click
    End If
End Sub

Private Sub cboVLvl4_Click()
    With cboVLvl4
        If .Value <> "" Then
            If .Value = "Group" Or .Value = "Alternates" Then
                ckbVLvl4.Value = True
            Else
                ckbVLvl4.Value = False
            End If
            If InStr(cboVLvl4.List(.ListIndex, 4), "code") = 0 Then
                ckbVLvl4.Enabled = False
            Else
                ckbVLvl4.Enabled = True
            End If
            'sXpath4 = .List(.ListIndex, 0)
            sLvl4Name = .List(.ListIndex, 1)
            'sLvl4xNd = .List(.ListIndex, 2)
            sLvl4Item = .List(.ListIndex, 3)
            sLvl4Code = .List(.ListIndex, 4)
            numLevel.Value = 4
            ckbVLvl5.Enabled = True
            ckbVLvl5.Value = False
            cboVLvl5.Enabled = True
            LoadCBO "cboVLvl5", "Page5"
        Else
            cboVLvl5.Enabled = False
            ckbVLvl5.Enabled = False
            ckbVLvl5.Value = False
        End If
    End With
End Sub

Private Sub cboVLvl5_Change()
    With cboVLvl5
        If cboVLvl5.Value <> "" Then
            If .Value = "GroupPhase" Or .Value = "Alternates" Then
                ckbVLvl5.Value = True
            Else
                ckbVLvl5.Value = False
            End If
            If InStr(cboVLvl5.List(.ListIndex, 4), "code") = 0 Then
                ckbVLvl5.Enabled = False
            Else
                ckbVLvl5.Enabled = True
            End If
            'sXpath5 = .List(.ListIndex, 0)
            sLvl5Name = .List(.ListIndex, 1)
            'sLvl5xNd = .List(.ListIndex, 2)
            sLvl5Item = .List(.ListIndex, 3)
            sLvl5Code = .List(.ListIndex, 4)
            numLevel.Value = 5
        End If
    End With
End Sub

Private Sub ckbVLvl1_Click()
    If ckbVLvl1.Value = True Then
        bCkb1 = True
    Else
        bCkb1 = False
    End If
End Sub

Private Sub ckbVLvl2_Click()
    If ckbVLvl2.Value = True Then
        bCkb2 = True
    Else
        bCkb2 = False
    End If
End Sub

Private Sub ckbVLvl3_Click()
    If ckbVLvl3.Value = True Then
        bCkb3 = True
    Else
        bCkb3 = False
    End If
End Sub

Private Sub ckbVLvl4_Click()
    If ckbVLvl4.Value = True Then
        bCkb4 = True
    Else
        bCkb4 = False
    End If
End Sub

Private Sub ckbVLvl5_Click()
    If ckbVLvl5.Value = True Then
        bCkb5 = True
    Else
        bCkb5 = False
    End If
End Sub

Private Sub txtVarXML_Change()
    If txtVarXML.Value <> "" Then
        LoadCBO "cboVLvl1", "Page5"
    End If
End Sub

Private Sub UserForm_Initialize()
        
    Dim firstRun As Boolean
    firstRun = Not ThisWorkbook.Worksheets("EstData").[rngIsTemp]
    Set reportsRequested = New Collection
    
    With Me
        .StartUpPosition = 0
        .Top = Application.Top + 115
        .Left = Application.Left + 25
    End With
    
    If firstRun Then
        With MultiPage1
            .Pages(0).Enabled = True
            .Pages(1).Enabled = False
            .Pages(2).Enabled = False
            .Pages(3).Enabled = False
            .Pages(4).Enabled = False
        End With
    Else
        With MultiPage1
            .Pages(0).Enabled = False
            .Pages(1).Enabled = True
            .Pages(2).Enabled = False
            .Pages(3).Enabled = True
            .Pages(4).Enabled = False
        End With
    End If
    
    LoadCBO "cboBLvl1", "Page1"
    LoadCBO "cboLvl1", "Page2"
    'LoadCBO "cboCLvl1", "Page3"
    LoadCBO "cboLvl0", "Page4"
    
End Sub

Public Function LoadCBO(cntrl As String, pg As String)
    
    Dim sortFieldDict As Dictionary
    Dim i As Integer
    i = 1
    With Controls(cntrl)
        .Clear
        .AddItem ""
        For Each sortFieldDict In getSortFieldColl()
                If CheckValue(pg, sortFieldDict("name")) = False Then
                    .AddItem
                    '.List(i, 0) = lObj.DataBodyRange.Cells(i, 5) 'sXpath1 'not needed
                    .List(i, 1) = sortFieldDict("name")
                    '.List(i, 2) = lObj.DataBodyRange.Cells(i, 4) 'sLvl1xNd 'not needed
                    .List(i, 3) = sortFieldDict("name") & "_combined"
                    .List(i, 4) = sortFieldDict("code")
                    i = i + 1
                End If
        Next
    End With
    
End Function

Public Function CheckValue(page As String, Value As String) As Boolean
Dim C As MSForms.control
    For Each C In MultiPage1.Pages(page).Controls
        If TypeName(C) = "ComboBox" Then
            If C.Value = Value Then
                CheckValue = True
                Exit Function
            End If
        End If
    Next
End Function

Function sRprtName() As String
    For i = 1 To 5
        If Controls("cboLvl" & i).Value <> "" Then
            sRprtName = sRprtName & Controls("cboLvl" & i).List(Controls("cboLvl" & i).ListIndex, 1)
            If i = 5 Then
                Exit Function
            Else
                If Controls("cboLvl" & i + 1).Value <> "" Then sRprtName = sRprtName & " | "
            End If
        End If
    Next
    sRprtName = sRprtName & " Level Sort Report"
End Function

Function sRprtNameVar() As String
    For i = 1 To 5
        If Controls("cboVLvl" & i).Value <> "" Then
            sRprtNameVar = sRprtNameVar & Controls("cboVLvl" & i).List(Controls("cboVLvl" & i).ListIndex, 1)
            If i = 5 Then
                Exit Function
            Else
                If Controls("cboVLvl" & i + 1).Value <> "" Then sRprtNameVar = sRprtNameVar & " | "
            End If
        End If
    Next
    sRprtNameVar = sRprtNameVar & " Level Variance Report"
End Function


Function tblExists(tblName As String, modName As String)
Set ows = Sheet0
    If TableExists(ows, tblName) Then
        MsgBox "Table1 Exists on sheet " & ActiveSheet.name
    Else
         MsgBox "Table1 Does Not Exist on sheet " & ActiveSheet.name
    End If
End Function

Function sXRprtName() As String
    sXRprtName = Controls("cboLvl0").List(Controls("cboLvl0").ListIndex, 1) & " by "
    For i = 1 To 5
        If Controls("cboXLvl" & i).Value <> "" Then
            sXRprtName = sXRprtName & Controls("cboXLvl" & i).List(Controls("cboXLvl" & i).ListIndex, 1)
            If i = 5 Then
                Exit Function
            Else
                If Controls("cboXLvl" & i + 1).Value <> "" Then sXRprtName = sXRprtName & " | "
            End If
        End If
    Next
    sXRprtName = sXRprtName & " CrossTab Report"
End Function

Function sCRprtName() As String
    For i = 1 To 5
        If Controls("cboCLvl" & i).Value <> "" Then
            sCRprtName = sCRprtName & Controls("cboCLvl" & i).List(Controls("cboCLvl" & i).ListIndex, 1)
            If i = 5 Then
                Exit Function
            Else
                If Controls("cboCLvl" & i + 1).Value <> "" Then sCRprtName = sCRprtName & " | "
            End If
        End If
    Next
    sCRprtName = sCRprtName & " Control Estimate"
End Function

Private Sub UserForm_Terminate()

    'If Not thisReportBuilder.success Then closeMe
    
End Sub
