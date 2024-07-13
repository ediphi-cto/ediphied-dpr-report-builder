VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EdiphiAPI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public apiKey As String
Const BASE_URL As String = "https://api.ediphi.com/api/external/"
Const TENANT_LABEL As String = "dpr"

Function getEstimateJSON(estimate_id As String) As String

    Dim req As New MSXML2.XMLHTTP60
    Dim url As String
    
    url = "https://api.ediphi.com/api/external/estimates/" & estimate_id & "?timestamp=" & timestamper()
    
    With req
        .Open "GET", url, True
        .setRequestHeader "Content-Type", "application/json"
        .setRequestHeader "api-token", apiKey
        .setRequestHeader "timestamp", timestamper()
        .setRequestHeader "api-tenant", TENANT_LABEL
        .send
        While req.readyState <> 4
            DoEvents
        Wend
        If .Status = "200" Then
            getEstimateJSON = .responseText
        Else
            GoTo e1
        End If
    End With

Exit Function
e1:
    'pass
    
End Function

Function getProject(project_id) As Dictionary
    '{
    '  id: 85c2c84e-cf97-4f16-870a-469a5329e842
    '  name:   test -Ketan
    '  owner: 98bee4f0-80d1-4010-af12-462dfa348184
    '  created_at: 2024-05-14 19:24:49.971+00
    '  updated_at: 2024-05-15 19:50:51.885+00
    '  phase:   WAITING
    '  deleted_at:
    '  industry:   LIFE_SCIENCES
    '  region: ae556d40-7b27-4162-892c-c98c3fbd456f
    '  client_name:   xyz
    '  address: 59 Moore Drive, Durham, NC, USA
    '  task_number:
    '  opportunity:
    '  test: False
    '  markup_budget: 0
    '  public: False
    '  color:
    '  archived:
    '  lng_lat: -78.86869279999999,35.9384034
    '  city:   Durham
    '  state:   NC
    '  zip: 27709
    '}

    Dim req As New MSXML2.XMLHTTP60
    Dim url As String
    Dim payload As String
    payload = "{""#project#"":{""table"":""projects"",""operation"":{""method"":""load"",""properties"":{""id"":""" & project_id & """}}}}"
    url = BASE_URL & "data/pipeline"
    With req
        .Open "POST", url, True
        .setRequestHeader "Content-Type", "application/json"
        .setRequestHeader "api-token", apiKey
        .setRequestHeader "api-tenant", TENANT_LABEL
        .send payload
        While req.readyState <> 4
            DoEvents
        Wend
        If .Status = "200" Then
            Set getProject = ParseJson(.responseText)("data")("load")("projects")(1)
        Else
            GoTo e1
        End If
    End With

Exit Function
e1:
    logError "failed to fetch projct " & project_id

End Function


Function get_sort_codes(sort_field_id) As Dictionary
    Dim req As New MSXML2.XMLHTTP60
    Dim url As String
    Dim payload As String
    payload = "{""#sort_code_{index.count()}#"":{""table"":""sort_codes"",""operation"":{""method"":""load.multiple()"",""properties"":{""sort_field"":""" & sort_field_id & """}}}}"
    url = BASE_URL & "data/pipeline"
    With req
        .Open "POST", url, True
        .setRequestHeader "Content-Type", "application/json"
        .setRequestHeader "api-token", apiKey
        .setRequestHeader "api-tenant", TENANT_LABEL
        .send payload
        While req.readyState <> 4
            DoEvents
        Wend
        If .Status = "200" Then
            Set get_sort_codes = parse_sort_codes_response(.responseText)
        Else
            GoTo e1
        End If
    End With

Exit Function
e1:
    'pass

End Function

Function get_sort_codes_by_field_key(project_id As String, sort_field_name As String) As Dictionary
    
    Dim req As New MSXML2.XMLHTTP60
    Dim url As String
    Dim payload As String
    
    payload = "{""#sort_field#"":{""table"":""sort_fields"",""operation"":{""method"":""load"",""properties"":{""project"":""" & _
        project_id & """,""name"":""" & sort_field_name & _
        """}}},""#sort_code_{index.count()}#"":{""table"":""sort_codes"",""operation"":{""method"":""load.multiple()"",""properties"":{""sort_field"":""#sort_field#""}}}}"
    
    url = BASE_URL & "data/pipeline"
    With req
        .Open "POST", url, True
        .setRequestHeader "Content-Type", "application/json"
        .setRequestHeader "api-token", apiKey
        .setRequestHeader "api-tenant", TENANT_LABEL
        .send payload
        While req.readyState <> 4
            DoEvents
        Wend
        If .Status = "200" Then
            Set get_sort_codes_by_field_key = parse_sort_codes_response(.responseText)
        Else
            GoTo e1
        End If
    End With

Exit Function
e1:
    'pass

End Function

Function parse_sort_codes_response(response_txt As String) As Dictionary
    
    Dim Data As Object
    Dim data_dict As Dictionary
    Dim sort_codes As Object, sort_code As Object, sort_field As Object
    Dim ans As New Dictionary
    Dim code_dict As Dictionary
    Dim desc As String
    Dim code As String
    Dim qty_var
    Dim dict As New Dictionary
    
    Set Data = ParseJson(response_txt)
    Set data_dict = Data("load")
    Set sort_codes = data_dict("sort_codes")
    Set sort_field = data_dict("sort_fields")(1)
    
    ans.Add "quantifier_uom", sort_field("quantifier_uom")
    For Each sort_code In sort_codes
        code = sort_code("code")
        desc = sort_code("description")
        qty_var = sort_code("quantifier")
        If desc <> "" And code <> "" Then
            Set code_dict = New Dictionary
            code_dict.Add "description", desc
            code_dict.Add "code", code
            If IsNumeric(CStr(qty_var)) Then
                If CDbl(qty_var) <> 0 Then
                    code_dict.Add "quantifier", CDbl(qty_var)
                End If
            End If
            dict.Add code, code_dict
        End If
    Next
    
    ans.Add "sort_codes", dict
    Set parse_sort_codes_response = ans

End Function

Function timestamper() As String
    
    timestamper = Format(Now, "yyyyMMddHHmmss")

End Function

Private Sub Class_Initialize()

    Me.apiKey = getEnv("API_KEY")

End Sub
