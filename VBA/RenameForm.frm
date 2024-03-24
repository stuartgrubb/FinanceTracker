VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} RenameForm 
   Caption         =   "Rename ""Form"""
   ClientHeight    =   2415
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4335
   OleObjectBlob   =   "RenameForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "RenameForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()

    ' Apply theme
    UserFormTheme Me

    ' Update captions
    Me.Caption = "Rename " & Form
    Label1.Caption = "Select the " & Form & " you would like to rename:"
    
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
    Dim oldName As String
    oldName = ComboBox1.Value
    
    ' Get the new Name from text box and run the 'ValidateName' function
    Dim Name As String
    Name = Trim(Me.TextBox.Value)
    If ValidateName(Name) Then
        Exit Sub
    End If
    
    ' Check for duplicate names using the 'IsDuplicate' function
    If IsDuplicate(Name) Then
        Exit Sub
    End If
    
    ' Confirm whether to proceed with renaming
    Dim response As VbMsgBoxResult
    response = MsgBox("Are you sure you want to rename '" & oldName & "' to '" & Name & "'?", vbYesNo + vbQuestion, "Confirm")
       
    ' Check response
    If response = vbNo Then
        Exit Sub
    End If
    
    ' Rename the entry on each sheet
    Rename oldName, Name
    
    MsgBox "'" & oldName & "' has been renamed to '" & Name & "'.", vbInformation, "Item Renamed"
    Unload Me
    
End Sub

'Cancel Button
Private Sub CancelButton_Click()
    Unload Me
End Sub
