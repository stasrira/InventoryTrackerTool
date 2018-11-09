VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public ID As Integer
Public Name As String

Public SQLCommand As String

Public Cond_format_columns As String
Public Cond_format_rules As String
Public Action_columns As String
Public Action_functions As String

Public dictActions As New Dictionary
Public dictCondFormats As New Dictionary

Public Function GetActionProcedureName(Action_Column_Name As String)
    
    'GetActionProcedureName = "GetNumberOfSamples_ReloadReport" 'hardcoded value for test only
    
    Dim out As String
    If Me.dictActions.Exists(Action_Column_Name) Then
        out = Me.dictActions(Action_Column_Name)
    End If
    
    GetActionProcedureName = out
    
End Function

Private Function CreateDictionaryFromArrays(ByRef targetDict As Dictionary, strKeys As String, strValues As String, Optional delim As String = ",")
    Dim arr1() As String, arr2() As String
    Dim i As Integer
    
    arr1 = Split(strKeys, ",")
    arr2 = Split(strValues, ",")
    
    If (Not Not arr1) > 0 And (Not Not arr2) > 0 Then 'check if the arrays were identified
        For i = 0 To UBound(arr1)
            If i <= UBound(arr2) Then
                targetDict.Add arr1(i), arr2(i)
            End If
        Next
    End If
End Function

Public Sub SetInitialValues()
    CreateDictionaryFromArrays Me.dictActions, Me.Action_columns, Me.Action_functions
    CreateDictionaryFromArrays Me.dictCondFormats, Me.Cond_format_columns, Me.Cond_format_rules
End Sub
