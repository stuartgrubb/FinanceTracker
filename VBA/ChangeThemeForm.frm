VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ChangeThemeForm 
   Caption         =   "Select Theme"
   ClientHeight    =   3810
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3735
   OleObjectBlob   =   "ChangeThemeForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ChangeThemeForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub UserForm_Initialize()
    
    ' Apply theme
    UserFormTheme Me
    
End Sub

Private Sub Light_Click()

    ' Store the theme
    ThisWorkbook.Sheets("Monthly Figures").Range("B2").Value2 = "Light"
    ChangeSheetTheme
    Unload Me
    ChangeThemeForm.Show

End Sub

Private Sub Dark_Click()

    ' Store the theme
    ThisWorkbook.Sheets("Monthly Figures").Range("B2").Value2 = "Dark"
    ChangeSheetTheme
    Unload Me
    ChangeThemeForm.Show
    
End Sub

Private Sub Blue_Click()

    ' Store the theme
    ThisWorkbook.Sheets("Monthly Figures").Range("B2").Value2 = "Blue"
    ChangeSheetTheme
    Unload Me
    ChangeThemeForm.Show
    
End Sub

Private Sub Green_Click()

    ' Store the theme
    ThisWorkbook.Sheets("Monthly Figures").Range("B2").Value2 = "Green"
    ChangeSheetTheme
    Unload Me
    ChangeThemeForm.Show
    

End Sub

Private Sub Purple_Click()

    ' Store the theme
    ThisWorkbook.Sheets("Monthly Figures").Range("B2").Value2 = "Purple"
    ChangeSheetTheme
    Unload Me
    ChangeThemeForm.Show

End Sub



