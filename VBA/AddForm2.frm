VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AddForm2 
   Caption         =   "Add APR ""Form"""
   ClientHeight    =   1815
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5760
   OleObjectBlob   =   "AddForm2.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "AddForm2"
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
    
    ' Get APR from text box and run the 'ValidateAPR' function
    Dim APRstring As String
    Dim APR As Double
    
    APRstring = Trim(Me.TextBox1.Value)
    If ValidateAPR(APRstring, APR) <> 0 Then
        Exit Sub
    End If
     
    ' Add new entry to each sheet using the 'Add' subroutine
    AddAPR Name, APR
      
    Unload Me
    
End Sub
'Cancel Button
Private Sub CancelButton_Click()
    Unload Me
End Sub

