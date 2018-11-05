VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)

    Dim numSamples As Integer
    
    If Target.Offset(-Target.Row + 1, 0).Value2 = cSpecialColumn_SampleQtyEstimated Then '"Sample Qty Estimated"
        
        GetNumberOfSamples_ReloadReport cInvWorkflowCapacityWorksheetName, InventoryWorkflowCapacityCheck, numSamples, Target.Value2
        
'        numSamples = GetInput_NumberOfSamlpes()
'        If numSamples > 0 Then 'if a positive number was returned, reload report and pass the received number of samples
'            LoadDataSheet cInvItemCapacityWorksheetName, InventoryItemsCapacityCheck, numSamples
'        End If
    End If
    

End Sub

