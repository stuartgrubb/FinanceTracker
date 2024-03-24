VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AddForm1 
   Caption         =   "Add ""Form"""
   ClientHeight    =   1110
   ClientLeft      =   90
   ClientTop       =   405
   ClientWidth     =   5760
   OleObjectBlob   =   "AddForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "AddForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()

    ' Apply theme
    UserFormTheme Me

    ' Update captions
    Me.Caption = "Add " & Form
    Label1.Caption = "Enter " & Form & " name"
        
        
End Sub

Private Sub OkButton_Click()

    ' Get Name from text box and run the 'ValidateName' function
    Dim Name As String
    Name = Trim(Me.TextBox.Value)
    If ValidateName(Name) Then
        Exit Sub
    End If
    
    ' Check for duplicate names using the 'IsDuplicate' function
    If IsDuplicate(Name) Then
        Exit Sub
    End If
    
    ' Add new entry to each sheet using the 'Add' subroutine
    Add Name
    
    Unload Me
    
End Sub
'Cancel Button
Private Sub CancelButton_Click()
    Unload Me
End Sub

