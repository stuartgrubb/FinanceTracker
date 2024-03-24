VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} OptionsForm 
   Caption         =   "Select Option"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5760
   OleObjectBlob   =   "OptionsForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "OptionsForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub UserForm_Initialize()
    
    ' Apply theme
    UserFormTheme Me
    
    ' Update captions
    Me.Caption = Form & " Options"
    Me.Add.Caption = "Add " & Form
    Me.Remove.Caption = "Remove " & Form
    Me.Rename.Caption = "Rename " & Form
      
End Sub

Private Sub Add_Click()

    ' Open the AddForm. If APR is required present Form2, otherwise present Form1.
    Unload OptionsForm
    
    If Form = "Mortgage" Or Form = "CreditCard" Or Form = "Loan" Then
        AddForm2.Show
    Else
        AddForm1.Show
    End If

End Sub

Private Sub ChangeAPR_Click()
    
    ' Open the ChangeAPRForm if entries are found in the table. Check the form type as only forms with APR should grant this option.
    If Form = "Mortgage" Or Form = "CreditCard" Or Form = "Loan" Then
        Dim Table As ListObject
        Set Table = ThisWorkbook.Sheets("Budget Tracker").ListObjects(Form)
    
        If Table.ListRows.Count = 0 Then
            MsgBox "No " & Form & "s found." & vbNewLine & "If you want to change a hidden entry, you must unhide it first.", vbInformation, "Error"
        Else
            Unload OptionsForm
            ChangeAPRForm.Show
        End If
    Else
        MsgBox "This function is only available for Mortgages, Credit Cards, or Loans", vbInformation, "Error"
    End If

End Sub

Private Sub ChangeOrder_Click()

    ' Open the ChangeOrderForm if entries are found in the form table.
    Dim Table As ListObject
    Set Table = ThisWorkbook.Sheets("Budget Tracker").ListObjects(Form)

    If Table.ListRows.Count = 0 Then
        MsgBox "No " & Form & "s found.", vbInformation, "Error"
    ElseIf Table.ListRows.Count = 1 Then
        MsgBox "More " & Form & "s are required to use this function.", vbInformation, "Error"
    Else
        Unload OptionsForm
        ChangeOrderForm.Show
    End If

End Sub

Private Sub HideUnhide_Click()

    ' Hide\Unhide option is only available when a month/year has not been selected
    If ThisWorkbook.Sheets("Monthly Figures").Range("B1").Value2 <> "" Then
        MsgBox "This function is not available when a month/year is selected. Please save the month/year first.", vbInformation, "Error"
        Exit Sub
    End If

    ' Check for entries in form table and Keystone table.
    Dim foundMatch As Boolean
    Dim Table As ListObject
    Set Table = ThisWorkbook.Sheets("Budget Tracker").ListObjects(Form)

    ' Check form table
    If Table.ListRows.Count = 0 Then
    
        ' Check Keystone table if no entries were found in the form table
        Set Table = ThisWorkbook.Sheets("Keystone").ListObjects("Keystone")
        Dim TableData As Variant
        
        If Table.ListRows.Count <> 0 Then
            TableData = Table.ListColumns(2).DataBodyRange.Value2
        
            If IsArray(TableData) Then
                For Each cell In TableData
                    If cell = Form Then
                        foundMatch = True ' Match found
                        Exit For
                    End If
                Next cell
            ElseIf Table.ListColumns(2).DataBodyRange.Value2 = Form Then
                foundMatch = True ' Match found
            End If
        End If
    Else
        foundMatch = True ' Entries found in form table
    End If

    ' Check result
    If foundMatch = True Then
        Unload OptionsForm
        HideUnhideForm.Show
    Else
        MsgBox "No " & Form & "s found.", vbInformation, "Error"
    End If

End Sub

Private Sub Remove_Click()
    
    ' Open the RemoveForm if entries are found in the form table.
    Dim Table As ListObject
    Set Table = ThisWorkbook.Sheets("Budget Tracker").ListObjects(Form)

    If Table.ListRows.Count = 0 Then
        MsgBox "No " & Form & "s found." & vbNewLine & "If you want to delete a hidden entry, you must unhide it first.", vbInformation, "Error"
    Else
        Unload OptionsForm
        RemoveForm.Show
    End If

End Sub

Private Sub Rename_Click()

    ' Open the RenameForm if entries are found in the form table.
    Dim Table As ListObject
    Set Table = ThisWorkbook.Sheets("Budget Tracker").ListObjects(Form)

    If Table.ListRows.Count = 0 Then
        MsgBox "No " & Form & "s found." & vbNewLine & "If you want to rename a hidden entry, you must unhide it first.", vbInformation, "Error"
    Else
        Unload OptionsForm
        RenameForm.Show
    End If

End Sub
