VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} RemoveYearForm 
   Caption         =   "Delete Year"
   ClientHeight    =   3720
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   2865
   OleObjectBlob   =   "RemoveYearForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "RemoveYearForm"
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
    
    ' Populate the listbox
    For Each Value In YearCollection
        ListBox1.AddItem Value
    Next Value

End Sub

Private Sub Delete_Click()
    
    ' Check that a selection was made
    If ListBox1.ListIndex = -1 Then
        MsgBox "No year selected.", vbInformation, "Input Required"
        Exit Sub
    End If
    
    ' Get the chosen selection from the combobox
    Dim selectedValue As String
    selectedValue = ListBox1.Value
    
    ' Confirm whether to proceed with deletion
    Dim response As VbMsgBoxResult
    response = MsgBox("Are you sure you want to delete '" & selectedValue & "' from the spreadsheet?", vbYesNo + vbQuestion, "Confirm")
       
    ' Check response
    If response = vbNo Then
        ' Reset Listbox selection
        ListBox1.ListIndex = -1
        Exit Sub
    End If
    
    ' Delete the entry from Data table using the Remove function.
    RemoveYear (selectedValue)
    
    MsgBox "'" & selectedValue & "' has been deleted.", vbInformation, "Item Deleted"
    Unload Me

End Sub

' Menu Cancel Button
Private Sub CancelButton_Click()
    Unload Me
End Sub
