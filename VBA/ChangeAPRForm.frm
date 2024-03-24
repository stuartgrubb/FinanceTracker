VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ChangeAPRForm 
   Caption         =   "Change ""Form"" APR"
   ClientHeight    =   1815
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5760
   OleObjectBlob   =   "ChangeAPRForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ChangeAPRForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub UserForm_Initialize()

    ' Apply theme
    UserFormTheme Me

    ' Update captions
    Me.Caption = "Change " & Form & " APR"
    Label1.Caption = "Select " & Form
    
    ' Identify entry names for combobox
    Dim Table As ListObject
    Dim TableData As Variant
    
    Set Table = ThisWorkbook.Sheets("Budget Tracker").ListObjects(Form)
    TableData = Table.ListColumns(1).DataBodyRange.Value2

    ' Add entry names to the combobox
    If IsArray(TableData) Then
        For Each i In TableData
            ComboBox1.AddItem i
        Next i
    Else
        ComboBox1.AddItem TableData
    End If
    
End Sub
Private Sub OkButton_Click()

    ' Check that a selection was made
    If ComboBox1.ListIndex = -1 Then
        MsgBox "No " & Form & " was selected.", vbInformation, "Input Required"
        Exit Sub
    End If

    ' Get the chosen selection from the combobox
    Dim Name As String
    Name = ComboBox1.Value

    ' Get APR from text box and run the 'ValidateAPR' function
    Dim APRstring As String
    Dim APR As Double
    
    APRstring = Trim(Me.TextBox1.Value)
    If ValidateAPR(APRstring, APR) <> 0 Then
        Exit Sub
    End If
    
    ' Get old APR from Keystone table
    Dim oldAPR As String
    Dim Table As ListObject
    Set Table = ThisWorkbook.Sheets("Keystone").ListObjects("Keystone")
    
    For Each Row In Table.ListRows
        If Row.Range.Cells(1, 1).Value2 = Name Then
            oldAPR = Row.Range.Cells(1, 3).Value2
            Exit For
        End If
    Next Row
    
    ' Confirm whether to proceed with APR change
    Dim response As VbMsgBoxResult
    response = MsgBox("Are you sure you want to change the APR on '" & Name & "' from " & oldAPR & "% to " & APR & "%?", vbYesNo + vbQuestion, "Confirm")
       
    ' Check response
    If response = vbNo Then
        Exit Sub
    End If
     
    ' Change APR using the 'ChangeAPR' subroutine
    ChangeAPR Name, APR
    
    MsgBox "The APR for '" & Name & "' has been changed to " & APR & "%.", vbInformation, "APR Changed"
    Unload Me
    
End Sub

'Cancel Button
Private Sub CancelButton_Click()
    Unload Me
End Sub
