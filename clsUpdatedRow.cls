VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsUpdatedRow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private iReport As clsReport
Private idictUpdFields As New Dictionary
Private iKeyValue As String

Public Property Get KeyValue() As String
     KeyValue = iKeyValue
End Property

Public Property Let KeyValue(val As String)
    iKeyValue = val
End Property

Public Property Get Report() As clsReport
    Set Report = iReport
End Property

Public Property Set Report(ByRef rep As clsReport)
    Set iReport = rep
End Property

Public Property Get dictUpdatedFields() As Dictionary
    Set dictUpdatedFields = idictUpdFields
End Property

Public Function CountOfUpdatedFields() As Integer
    CountOfUpdatedFields = idictUpdFields.Count
End Function

Public Sub AddUpdatedField(fieldName As String, _
                            OldValue As String, _
                            NewValue As String)
    
    Dim cUpdField As clsUpdatedField
    
    If NewValue <> OldValue Then
    
        If Not idictUpdFields.Exists(fieldName) Then
            cUpdField = New clsUpdatedField
            idictUpdFields.Add fieldName, cUpdField
        Else
            Set cUpdField = idictUpdFields(fieldName)
        End If
        
        With cUpdField
            .fieldName = fieldName
            .NewFieldValue = NewValue
            .OldFieldValue = OldValue
            .DataRow = Me
        End With
    End If
    
End Sub

Public Sub RemoveUpdatedField(fieldName As String)
    If idictUpdFields.Exists(fieldName) Then
        idictUpdFields.Remove (fieldName)
    End If
End Sub

Public Sub RemoveAllUpdatedField()
    idictUpdFields.RemoveAll
End Sub

