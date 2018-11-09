VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_Open()
     
    'set comparision mode of the global dictionary - dictValidationResults
    dictValidationResults.CompareMode = TextCompare
    
    'set automatic calculation for the whole workbook - this will make sure that Automatic calculation is set as default
    Application.Calculation = xlCalculationAutomatic
    
    'set dropdown functionality based on the config settings
    bVoidDropDownFunctionality = CBool(GetConfigValue("dflt_bVoidDropDownFunctionality")) 'True
    
    'register local events
    'RegisterCustomEvents
    
    'create custom menus
    LoadCustomMenus
    
    'load list of reports into memory
    GetInventoryReportsInfo
     
End Sub

'Testing only ---->  test sub for testing loading of workbook
Public Sub TestLoad()
    Workbook_Open
End Sub
'Testing only ---->  test sub for testing closing of workbook
Public Sub TestBeforeClose()
    Workbook_BeforeClose False
End Sub
 
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    On Error Resume Next 'in case the menu item has already been deleted
    Application.CommandBars("Worksheet Menu Bar").Controls(cCustomMenuName).Delete 'delete the menu item
    
    UnRegisterCustomEvents 'remove registrations for custom events
    
    'Clean up memory - kill dictionary objects
    Set dictValidationResults = Nothing
    Set dictFieldSettings = Nothing
        
End Sub

Private Sub Workbook_SheetBeforeDoubleClick(ByVal Sh As Object, ByVal Target As Range, Cancel As Boolean)
    Dim colHeader As String
    Dim r As clsReport
    Dim rep As Variant
    
    'check if Report's dictionary was loaded from DB
    GetInventoryReportsInfo
    
    colHeader = Target.Offset(-Target.Row + 1, 0).Value2
    
    For Each rep In dictReports 'todo - think how to avoid looping through all dictionary items on each double click
        If dictReports(rep).Name = Sh.Name Then
            Set r = dictReports(rep)
        End If
    Next
    
    If Not r Is Nothing Then
        
        'Set r = dictReports(Sh.Name)
        
        'check if the current column (that the cell belongs to) is listed as an action column
        'If Target.Offset(-Target.Row + 1, 0).Value2 = cSpecialColumn_SampleQtyEstimated Then '"Sample Qty Estimated"
        If InStr(1, r.Action_columns, colHeader) Then
            
            Select Case r.GetActionProcedureName(colHeader)
                Case "GetNumberOfSamples_ReloadReport"
                    GetNumberOfSamples_ReloadReport Sh.Name, r.ID, Target.Value2
            End Select
            
            Cancel = True 'cancel the dbl-click event from proceeding further
        End If
    End If
End Sub
