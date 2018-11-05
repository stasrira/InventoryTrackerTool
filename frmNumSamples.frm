VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmNumSamples 
   Caption         =   "Number of Samples"
   ClientHeight    =   2130
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3780
   OleObjectBlob   =   "frmNumSamples.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmNumSamples"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const msg_MsgTitle = "Entering Number of Samples"
Const msg_NumericOnly = "Only positive numeric values are allowed!" & vbCrLf & "Please reenter the value."

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdCancel_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
'    If KeyCode = 27 And Shift = 0 Then 'if Escape was pressed, close the form
'        cmdCancel_Click
'    End If
    checkPressedButtons KeyCode, Shift
End Sub

Private Sub cmdSubmit_Click()
    Dim val As String
    Dim showErrMsg As Boolean
    
    'validate provided value
    val = Me.txtNumSamples.value
    If Not (IsNumeric(Trim(val))) Then
        showErrMsg = True
    Else
        If val < 0 Then
            showErrMsg = True
        End If
    End If
    
    If showErrMsg Then
        'if there is an error, display a msg box
        MsgBox msg_NumericOnly, vbCritical, msg_MsgTitle
        
        HighLightTextBoxValue
        
'        With Me.txtNumSamples
'            .SetFocus
'            .SelStart = 0
'            .SelLength = Len(.value)
'        End With
    Else
        'if no errors found, return entered value
        popUpFormResponse_SampleNum = val 'pass entered number back to caller
        Unload Me
    End If

End Sub

Private Sub cmdSubmit_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
'    If KeyCode = 27 And Shift = 0 Then 'if Escape was pressed, close the form
'        cmdCancel_Click
'    End If
    checkPressedButtons KeyCode, Shift
End Sub


Private Sub checkPressedButtons(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 And Shift = 0 Then 'if enter was pressed
        cmdSubmit_Click
        KeyCode = 0
    ElseIf KeyCode = 27 And Shift = 0 Then 'if Escape was pressed, close the form
        cmdCancel_Click
    End If
End Sub

Private Sub txtNumSamples_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    checkPressedButtons KeyCode, Shift
End Sub

Private Sub UserForm_Initialize()
    
    If IsNumeric(popUpFormResponse_SampleNum) Then
        If CInt(popUpFormResponse_SampleNum) > 0 Then
            Me.txtNumSamples.Text = popUpFormResponse_SampleNum
        End If
    End If
    
    popUpFormResponse_SampleNum = "-1" 'assign default value
    
    HighLightTextBoxValue
    'Me.txtNumSamples.SetFocus
End Sub

Private Sub HighLightTextBoxValue()
    With Me.txtNumSamples
        .SetFocus
        .SelStart = 0
        .SelLength = Len(.value)
    End With
End Sub
