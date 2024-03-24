VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} HideUnhideForm 
   Caption         =   "Hide or Unhide ""Form""s"
   ClientHeight    =   2910
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4365
   OleObjectBlob   =   "HideUnhideForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "HideUnhideForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub UserForm_Initialize()

    ' Apply theme
    UserFormTheme Me

    ' Update captions
    Me.Caption = "Hide or Unhide " & Form & "s"
    Label1.Caption = "Visible " & Form & "s"
    Label2.Caption = "Hidden " & Form & "s"
    
    ' Populate ListBox1
    Dim Table As ListObject
    Set Table = ThisWorkbook.Sheets("Keystone").ListObjects("Keystone")
    
    For Each Row In Table.ListRows
        If Row.Range.Cells(1, 2).Value2 = Form Then
            If Row.Range.Cells(1, 4).Value2 = "Visible" Then
                ListBox1.AddItem Row.Range.Cells(1, 1).Value2
            End If
        End If
    Next Row
        
    ' Populate ListBox2
    For Each Row In Table.ListRows
        If Row.Range.Cells(1, 2).Value2 = Form Then
            If Row.Range.Cells(1, 4).Value2 = "Hidden" Then
                ListBox2.AddItem Row.Range.Cells(1, 1).Value2
            End If
        End If
    Next Row
    
End Sub

Private Sub Hide_Click()

    ' Check that a selection was made
    If ListBox1.ListIndex = -1 Then
        MsgBox "No " & Form & " was selected.", vbInformation, "Input Required"
        Exit Sub
    End If

    ' Get the chosen selection from the combobox
    Dim selectedName As String
    selectedName = ListBox1.Value
    
    ' Remove entry from Budget Tracker sheet
    Dim RowMatch As ListRow
    Set RowMatch = Nothing
    Set Table = ThisWorkbook.Sheets("Budget Tracker").ListObjects(Form)
    
    For Each Row In Table.ListRows
        If Row.Range.Cells(1, 1).Value2 = selectedName Then
            Set RowMatch = Row
            Row.Delete
            Exit For
        End If
    Next Row
    
    ' Update the Keystone visibility record
    Set Table = ThisWorkbook.Sheets("Keystone").ListObjects("Keystone")
    
    For Each Row In Table.ListRows
        If Row.Range.Cells(1, 1).Value2 = selectedName Then
            Row.Range.Cells(1, 4).Value2 = "Hidden"
            Exit For
        End If
    Next Row
        
    ' Update ListBoxes
    ListBox1.Clear
    ListBox2.Clear
    ' Populate ListBox1
    Set Table = ThisWorkbook.Sheets("Keystone").ListObjects("Keystone")
    
    For Each Row In Table.ListRows
        If Row.Range.Cells(1, 2).Value2 = Form Then
            If Row.Range.Cells(1, 4).Value2 = "Visible" Then
                ListBox1.AddItem Row.Range.Cells(1, 1).Value2
            End If
        End If
    Next Row
        
    ' Populate ListBox2
    For Each Row In Table.ListRows
        If Row.Range.Cells(1, 2).Value2 = Form Then
            If Row.Range.Cells(1, 4).Value2 = "Hidden" Then
                ListBox2.AddItem Row.Range.Cells(1, 1).Value2
            End If
        End If
    Next Row
    
    
End Sub

Private Sub Unhide_Click()

    ' Check that a selection was made
    If ListBox2.ListIndex = -1 Then
        MsgBox "No " & Form & " was selected.", vbInformation, "Input Required"
        Exit Sub
    End If
    
    ' Get the selection from the combobox
    Dim selectedName As String
    selectedName = ListBox2.Value
    Dim Table As ListObject
    
    ' Check if APR data is required, if so capture the APR from the keystone table.
    If Form = "Mortgage" Or Form = "CreditCard" Or Form = "Loan" Then

        Dim APR As String ' String is used instead of Double as Doubles are initialised as "0". This would cause the APR <> ""' check to not behave properly.
        Set Table = ThisWorkbook.Sheets("Keystone").ListObjects("Keystone")

        For Each Row In Table.ListRows
            If Row.Range.Cells(1, 1).Value2 = selectedName Then
            APR = Row.Range.Cells(1, 3).Value2
            Exit For
            End If
        Next Row
    End If
    
    
    
    ' Add data to the Budget Tracker sheet
    Dim newRowIndex As Long
    
    ' Add non-APR item to the Budget Tracker sheet
    If Form = "Income" Or Form = "Bill" Or Form = "SavingsAccount" Or Form = "Investment" Then
          
        Set Table = ThisWorkbook.Sheets("Budget Tracker").ListObjects(Form)
        newRowIndex = Table.ListRows.Count + 1
        Table.ListRows.Add newRowIndex
        Table.ListRows(newRowIndex).Range.Cells(1, 1).Value2 = selectedName
        Table.ListRows(newRowIndex).Range.Cells(1, 2).Value2 = 0
    
    ' Add APR item to the Budget Tracker sheet
    ElseIf Form = "Mortgage" Or Form = "CreditCard" Or Form = "Loan" Then

        Set Table = ThisWorkbook.Sheets("Budget Tracker").ListObjects(Form)
        newRowIndex = Table.ListRows.Count + 1
        Table.ListRows.Add newRowIndex
        Table.ListRows(newRowIndex).Range.Cells(1, 1).Value2 = selectedName
        Table.ListRows(newRowIndex).Range.Cells(1, 2).Value2 = APR
        Table.ListRows(newRowIndex).Range.Cells(1, 3).Value2 = 0
    
    End If
    
       
    ' Update the Keystone visibility record and move the entry to the bottom of the table (maintains order of entries on the input and keystone tables).
    Set Table = ThisWorkbook.Sheets("Keystone").ListObjects("Keystone")
    Dim selectedIndex As Long
    
    For Each Row In Table.ListRows
        If Row.Range.Cells(1, 1).Value2 = selectedName Then
            Row.Range.Cells(1, 4).Value2 = "Visible"
            selectedIndex = Row.Index
            newRowIndex = Table.ListRows.Count + 1
            Table.ListRows.Add newRowIndex
            Table.ListRows(newRowIndex).Range.Value = Table.ListRows(selectedIndex).Range.Value
            Row.Delete
            Exit For
        End If
    Next Row
    


    ' Update ListBoxes
    ListBox1.Clear
    ListBox2.Clear
    ' Populate ListBox1
    Set Table = ThisWorkbook.Sheets("Keystone").ListObjects("Keystone")
    
    For Each Row In Table.ListRows
        If Row.Range.Cells(1, 2).Value2 = Form Then
            If Row.Range.Cells(1, 4).Value2 = "Visible" Then
                ListBox1.AddItem Row.Range.Cells(1, 1).Value2
            End If
        End If
    Next Row
        
    ' Populate ListBox2
    For Each Row In Table.ListRows
        If Row.Range.Cells(1, 2).Value2 = Form Then
            If Row.Range.Cells(1, 4).Value2 = "Hidden" Then
                ListBox2.AddItem Row.Range.Cells(1, 1).Value2
            End If
        End If
    Next Row
    
End Sub
