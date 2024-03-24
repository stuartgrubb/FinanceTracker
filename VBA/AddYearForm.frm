VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AddYearForm 
   Caption         =   "Add Year"
   ClientHeight    =   1575
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   2520
   OleObjectBlob   =   "AddYearForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "AddYearForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()

    ' Apply theme
    UserFormTheme Me

End Sub
    
Private Sub OkButton_Click()

    ' Validate input and check for duplicate, must be a 4 digit year (YYYY)
    Dim Year As String ' Using string for invalid input handling
    Dim YearInt As Integer
    Year = Trim(Me.TextBox1.Value)
    If ValidateYear(Year, YearInt) <> 0 Then
        Exit Sub
    End If
    
    ' Add the entry selected from each sheet using the AddYear Function
    AddYear (YearInt)
    
    MsgBox "'" & Year & "' has been added.", vbInformation, "Item Added"
    Unload Me

End Sub

'Cancel Button
Private Sub CancelButton_Click()
    Unload Me
End Sub

