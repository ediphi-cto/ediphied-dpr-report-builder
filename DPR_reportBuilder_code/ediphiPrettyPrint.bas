Attribute VB_Name = "ediphiPrettyPrint"
Option Explicit

Public LOG_LINES As String
Private Const DEFAULT_LINES_TO_PRINT As Integer = 190

Sub pp(var, Optional abbrev As Boolean = True, Optional lines_to_print As Integer = DEFAULT_LINES_TO_PRINT)
    
    Dim coll As Collection
    Dim dict As Dictionary
    Dim ans As String
    Dim line_char_limit As Integer
    
    If abbrev Then
        line_char_limit = 100
    Else
        line_char_limit = -1
    End If
    
    If Right(TypeName(var), 2) = "()" Then
        pp_lines array_as_str(var, line_char_limit), lines_to_print
    Else
        Select Case TypeName(var)
            Case "Dictionary"
                Set dict = var
                pp_lines dict_as_str(dict, 0, abbrev), lines_to_print
            Case "Collection"
                Set coll = var
                pp_lines coll_as_str(coll, 0, abbrev), lines_to_print
            Case "Range"
                Debug.Print range_as_str(var)
            Case Else
                ans = cStr_safe(var)
                If ans = "" Then ans = TypeName(var)
                Debug.Print ans
        End Select
    End If

End Sub

Sub pp_lines(many_lines As String, Optional lines_to_print As Integer = DEFAULT_LINES_TO_PRINT)
    
    Dim arr
    Dim arr_len As Integer
    Dim i As Integer
    Dim remaining_lines As String
    Dim j As Integer
    
    arr = Split(many_lines, vbLf)
    arr_len = arrLength(arr)
    
    For i = LBound(arr) To UBound(arr)
        If i < lines_to_print Then
            Debug.Print arr(i)
        Else
            j = j + 1
            remaining_lines = remaining_lines & vbLf & arr(i)
        End If
    Next
    If j > 0 Then
        Debug.Print "+---------------------------------------------------------"
        Debug.Print " | " & CStr(j) & " lines remaining"
        Debug.Print " | TYPE ppnext OR ppn TO PRINT MORE"
        Debug.Print "+---------------------------------------------------------"
    End If
    LOG_LINES = remaining_lines

End Sub

Sub ppn(Optional lines_to_print As Integer = DEFAULT_LINES_TO_PRINT)

    ppNext lines_to_print
    
End Sub

Sub ppNext(Optional lines_to_print As Integer = DEFAULT_LINES_TO_PRINT)
    
    pp_lines LOG_LINES, lines_to_print
    
End Sub

Function coll_as_str(coll As Collection, Optional indent_lvl As Integer = 0, Optional abbrev As Boolean = True) As String
    
    Dim ans As String
    Dim var
    Dim dict As Dictionary
    Dim coll2 As Collection
    Dim abbrev_coll As Collection
    Dim coll_size As Integer
    Dim i As Integer
    
    coll_size = coll.Count
    If coll_size = 0 Then
        coll_as_str = "EMPTY COLLECTION"
        Exit Function
    End If
    
    If abbrev Then
        Set abbrev_coll = New Collection
        abbrev_coll.Add coll(1)
        Set coll = abbrev_coll
    End If
    
    ans = ans & indents(indent_lvl) & "["
    For Each var In coll
            ans = ans & vbLf '& indents(indent_lvl)
            If Right(TypeName(var), 2) = "()" Then
                ans = ans & indents(indent_lvl) & array_stats(var)
            Else
                Select Case TypeName(var)
                    Case "Dictionary"
                        Set dict = var
                        ans = ans & dict_as_str(dict, indent_lvl + 2, abbrev)
                    Case "Collection"
                        Set coll2 = var
                        ans = ans & coll_as_str(coll2, indent_lvl + 2, abbrev)
                    Case Else
                        ans = ans & indents(indent_lvl) & cStr_safe(var)
                End Select
            End If
    Next
    If abbrev Then
        ans = ans & vbLf & indents(indent_lvl) & " ..." & CStr(coll_size) & " ITEMS IN COLLECTION"
    End If
    ans = ans & vbLf & indents(indent_lvl) & "]"
    coll_as_str = ans
    
End Function

Function dict_as_str(dict As Dictionary, Optional indent_lvl As Integer = 0, Optional abbrev As Boolean = True) As String

    Dim k
    Dim ans As String
    Dim lvl As Integer
    Dim var
    Dim i As Integer
    lvl = 0
    ans = indents(indent_lvl) & "{"
    For Each k In dict.Keys()
            i = i + 1
            ans = ans & vbLf & indents(indent_lvl + 1) & k & ": "
            If Right(TypeName(dict(k)), 2) = "()" Then
                ans = ans & indents(indent_lvl) & array_stats(var)
            Else
                Select Case TypeName(dict(k))
                    Case "DIctionary"
                        ans = ans & dict_as_str(dict(k), indent_lvl + 2, abbrev)
                    Case "Collection"
                        ans = ans & coll_as_str(dict(k), indent_lvl + 2, abbrev)
                    Case Else
                        ans = ans & cStr_safe(dict(k))
                End Select
            End If
    Next
    ans = ans & vbLf & indents(indent_lvl) & "}"

    dict_as_str = ans

End Function

Function indents(lvl As Integer, Optional indent_size As Integer = 2) As String
    
    Dim space_ct As Integer
    space_ct = lvl * indent_size
    Dim i As Integer
    For i = 1 To space_ct
        indents = indents & " "
    Next

End Function

Function range_as_str(ran) As String
    
    If TypeName(ran) <> "Range" Then Exit Function
    range_as_str = "Range Address: " & ran.Address & " Sheet: " & ran.Parent.name

End Function

Function array_as_str(arr, Optional line_char_limit As Integer = 100) As String

    Dim r As Integer, C As Integer
    Dim dimensions As Integer
    Dim row_str As String
    Dim ans As String
    Dim val As String
    
    dimensions = dimension_count(arr)
    ans = array_stats(arr) & vbLf
    Select Case dimensions
        Case 0
            ans = ans & "[EMPTY]"
        Case 1
            ans = ans + "[ "
            For r = LBound(arr, 1) To UBound(arr, 1)
                ans = ans & arr(r)
                If r <> UBound(arr, 1) Then
                    ans = ans + ", "
                Else
                    ans = ans + " ]"
                End If
            Next
        Case 2
            For r = LBound(arr, 1) To UBound(arr, 1)
                row_str = "["
                For C = LBound(arr, 2) To UBound(arr, 2)
                    row_str = row_str & arr(r, C)
                    If C <> UBound(arr, 2) Then
                        row_str = row_str + ", "
                    Else
                        row_str = row_str + " ]"
                    End If
                Next
                If Len(row_str) > line_char_limit And line_char_limit <> -1 Then
                    row_str = Left(row_str, line_char_limit - 5) & "... ]"
                End If
                ans = ans & vbLf & row_str
            Next
        Case Else
            'pass
    End Select
        
    array_as_str = ans

End Function

Function array_stats(arr) As String

    If IsEmpty(arr) Then
        array_stats = "[EMPTY]"
        Exit Function
    End If
    
    Dim ct As Integer
    ct = dimension_count(arr)
    Dim txt As String
    Dim i As Integer, j As Integer
    txt = "("
    For i = 1 To ct
        txt = txt & LBound(arr, i) & " to " & UBound(arr, i)
        If i <> ct Then
            txt = txt & ", "
        Else
            txt = txt + ")"
        End If
    Next
    
    array_stats = "array: " & txt

End Function

Private Function dimension_count(arr) As Integer

    Dim i As Integer
    Dim has_error As Boolean
    Dim ans As Integer
    
    On Error GoTo e1
    
    Do Until has_error
        i = i + 1
        ans = UBound(arr, i)
    Loop
    
    dimension_count = i - 1

Exit Function
e1:
    has_error = True
    Resume Next

End Function


Function cStr_safe(val) As String
    On Error GoTo e1
    cStr_safe = CStr(val)
    If LCase(cStr_safe) = "null" Then cStr_safe = ""
    
Exit Function
e1:
    cStr_safe = ""
    
End Function

Function cDbl_safe(val) As String
    On Error GoTo e1
    cDbl_safe = CDbl(val)
Exit Function
e1:
    cDbl_safe = 0
    
End Function

Function cInt_safe(val) As String
    On Error GoTo e1
    cInt_safe = CInt(val)
Exit Function
e1:
    cInt_safe = 0
    
End Function

Sub appendArray(ByRef arrayTable, ByVal arrayRecord)
    
    Dim i As Long
    Dim new_i As Long
    
    new_i = UBound(arrayTable, 2) + 1
    
    If (UBound(arrayRecord) - LBound(arrayRecord)) <> (UBound(arrayTable, 1) - LBound(arrayTable, 1)) Then GoTo e0
    
    ReDim Preserve arrayTable(LBound(arrayTable, 1) To UBound(arrayTable, 1), LBound(arrayTable, 2) To new_i)
    
    For i = LBound(arrayRecord) To UBound(arrayRecord)
        arrayTable(i, new_i) = arrayRecord(i)
    Next
    
Exit Sub
e0:
    Debug.Print "appendArray Failed, record width did not equal table width"

End Sub

Function str2array(str As String, Optional splitLen As Long = 30000) As String()

    Dim ans() As String
    Dim i As Long
    Dim j As Long
    
    ReDim ans(0)
    j = 0
    i = 1
    Do Until i >= Len(str)
        ReDim Preserve ans(j)
        ans(j) = Mid(str, i, splitLen)
        i = i + splitLen
        j = j + 1
    Loop
    
    str2array = ans

End Function

Function printArr(ByRef ran As Range, arr, Optional asText As Boolean) As Range

    Set ran = ran.Cells(1, 1).Resize(arrLength(arr, 1), arrLength(arr, 2))
    If asText Then ran.NumberFormat = "@"
    ran.Value = arr
    Set printArr = ran
    
End Function

Function arrLength(arr, Optional dimension As Integer = 1) As Long

    arrLength = UBound(arr, dimension) - LBound(arr, dimension) + 1

End Function

Function quote(str As String) As String
    'wraps text in quotes, useful for json
    
    str = Replace(str, "\", "\\")
    str = Replace(str, Chr(34), "\""")
    str = Replace(str, vbLf, "   \n ")
    
    quote = """" & str & """"
    
End Function



