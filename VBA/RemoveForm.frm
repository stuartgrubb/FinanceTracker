VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} RemoveForm 
   Caption         =   "Delete ""Form"""
   ClientHeight    =   3735
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5760
   OleObjectBlob   =   "RemoveForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "RemoveForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub UserForm_Initialize()

    ' Apply theme
    UserFormTheme Me
    
    ' Set warning label color
    Me.WarningLabel.BackColor = RGB(255, 255, 225)
    Me.WarningLabel.ForeColor = RGB(0, 0, 0)

    ' Update captions
    Me.Caption = "Delete " & Form
    Label2.Caption = "Select " & Form
    
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

Private Sub Delete_Click()

    ' Check that a selection was made
    If ComboBox1.ListIndex = -1 Then
        MsgBox "No " & Form & " was selected.", vbInformation, "Input Required"
        Exit Sub
    End If
    
    ' Get the chosen selection from the combobox
    Dim selectedValue As String
    selectedValue = ComboBox1.Value
    
    ' Confirm whether to proceed with deletion
    Dim response As VbMsgBoxResult
    response = MsgBox("Are you sure you want to delete '" & selectedValue & "'?", vbYesNo + vbQuestion, "Confirm")
       
    ' Check response
    If response = vbNo Then
        Exit Sub
    End If
    
    ' Delete the entry selected from each sheet using the Remove function.
    Remove (selectedValue)
    
    MsgBox Form & " '" & selectedValue & "' has been deleted.", vbInformation, "Item Deleted"
    Unload Me

End Sub
'Cancel Button
Private Sub CancelButton_Click()
    Unload Me
End Sub

