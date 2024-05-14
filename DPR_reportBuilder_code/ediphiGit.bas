Attribute VB_Name = "ediphiGit"
Option Explicit

Sub gitExplode()
    
    'NOTE:  Call this method before commiting changes
    extractInPlace ThisWorkbook

End Sub

Sub clearRepo(repoPath As String)
    
    Dim file
    Dim fso As New FileSystemObject
    
    file = Dir(repoPath)
    
    While file <> ""
        fso.DeleteFile repoPath & file
        file = Dir
    Wend
    
End Sub

Function myRepoPath() As String

    myRepoPath = ThisWorkbook.Path & "\" & Split(ThisWorkbook.Name, ".")(0) & "_code\"

End Function

Sub extractInPlace(WB As Workbook)
     
    Dim vbaPROJ As VBProject
    Dim vbaMODULE As VBComponent
    Dim repoPath As String
    Dim fso As New FileSystemObject
    
    Set vbaPROJ = WB.VBProject
    repoPath = myRepoPath()
    
    If Not fso.FolderExists(repoPath) Then
        Call fso.CreateFolder(repoPath)
    Else
        clearRepo repoPath
    End If
    Set fso = Nothing
    
    For Each vbaMODULE In vbaPROJ.VBComponents
        With vbaMODULE
            Select Case .Type
                Case 1
                    .Export repoPath & .Name & ".bas"
                Case 2
                    .Export repoPath & .Name & ".cls"
                Case 3
                    .Export repoPath & .Name & ".frm"
            End Select
        End With
    Next
    
End Sub

Sub cloneMeInto(WB As Workbook)
    
    gitExplode
    clearVBAfrom WB
    ImportVBComponentsInto WB, myRepoPath()
    
End Sub

Function ImportVBComponentsInto(targetWb As Workbook, directoryPath As String)
    Dim fso As FileSystemObject
    Dim folder As folder
    Dim file As file
    Dim ext As String
    
    Set fso = New FileSystemObject
    Set folder = fso.GetFolder(directoryPath)

    For Each file In folder.Files
        ext = fso.GetExtensionName(file.Path)
        If ext = "bas" Or ext = "cls" Or ext = "frm" Then
            targetWb.VBProject.VBComponents.Import file.Path
        End If
    Next file
    
End Function

Sub clearVBAfrom(WB As Workbook)
    Dim vbComp As VBIDE.VBComponent
    Dim vbProj As VBIDE.VBProject
    Set vbProj = WB.VBProject
    
    For Each vbComp In vbProj.VBComponents
        If vbComp.Type = vbext_ct_StdModule Or vbComp.Type = vbext_ct_ClassModule Or vbComp.Type = vbext_ct_MSForm Then
            vbProj.VBComponents.Remove vbComp
        End If
    Next vbComp

End Sub


