Attribute VB_Name = "mdlButton"
Option Explicit

Public Sub CreateButtons()

    Dim i As Long
    Dim shp As Object
    Dim dblLeft As Double
    Dim dblTop As Double
    Dim dblWidth As Double
    Dim dblHeight As Double
    Dim r As Range
    
    Const ButtonColumn = "A:A"
    Const btnCaption = "Name of the Button"
        
    With Sheets("ButtonsTest")
        .Buttons.Delete
        .Cells.ClearContents
        
        'adjust column width to fit the button's title
        Set r = .Range(ButtonColumn).Cells(1)
        r.value = btnCaption
        r.EntireColumn.AutoFit
        r.value = ""
        
        dblLeft = .Columns(ButtonColumn).Left      'All buttons have same Left position
        dblWidth = .Columns(ButtonColumn).Width    'All buttons have same Width
        For i = 1 To 20                     'Starts on row 2 and finishes row 20
            dblHeight = .Rows(i).Height     'Set Height to height of row
            dblTop = .Rows(i).Top           'Set Top top of row
            Set shp = .Buttons.Add(dblLeft, dblTop, dblWidth, dblHeight)
            shp.OnAction = "'IdentifySelected " & """" & "Value" & CStr(i) & """'"  '"IdentifySelected"
            shp.Characters.Text = btnCaption
            'r.Cells(i).value = "Value" & CStr(i)
        Next i
        .Cells.EntireColumn.AutoFit
    End With
   
End Sub


Public Sub IdentifySelected(val)
    'NOTE: The button will always be on the active sheet
    Dim strButtonName
    Dim lngRow As Long
    
    strButtonName = ActiveSheet.Shapes(Application.Caller).Name
    lngRow = ActiveSheet.Shapes(strButtonName).TopLeftCell.Row
    
    MsgBox "Button is on row " & lngRow & vbCrLf & "Passed value is " & val
End Sub
