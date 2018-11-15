VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsUpdatedField"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private iOldFieldValue As String
Private iNewFieldValue As String
Private iFieldName As String
Private bOldValueUpdated As Boolean
Private clsRow As clsUpdatedRow

Public Property Get OldFieldValue() As String
    OldFieldValue = iOldFieldValue
End Property

Public Property Let OldFieldValue(val As String)
    If Not bOldValueUpdated Then
        iOldFieldValue = val
        bOldValueUpdated = True 'old value will be updated only once
    End If
End Property

Public Property Get NewFieldValue() As String
     NewFieldValue = iNewFieldValue
End Property

Public Property Let NewFieldValue(val As String)
    iNewFieldValue = val
End Property

Public Property Get fieldName() As String
     fieldName = iFieldName
End Property

Public Property Let fieldName(fldName As String)
    iFieldName = fldName
End Property

Public Property Get DataRow() As clsUpdatedRow
    Set DataRow = clsRow
End Property

Public Property Set DataRow(ByRef Row As clsUpdatedRow)
    Set clsRow = Row
End Property
