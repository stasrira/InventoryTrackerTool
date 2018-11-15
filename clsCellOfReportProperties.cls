VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCellOfReportProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private iRng As Range
Private iFieldHeader As String
Private iPKValue As String
Private iReport As clsReport

Public Property Get Range() As Range
     Set Range = iRng
End Property

Public Property Set Range(ByRef Range As Range)
    Set iRng = Range
End Property

Public Property Get Report() As clsReport
    Set Report = iReport
End Property

Public Property Set Report(ByRef rep As clsReport)
    Set iReport = rep
End Property

Public Property Get FieldHeader() As String
    FieldHeader = iFieldHeader
End Property

Public Property Let FieldHeader(val As String)
    iFieldHeader = val
End Property

Public Property Get PKValue() As String
    PKValue = iPKValue
End Property

Public Property Let PKValue(val As String)
    iPKValue = val
End Property
