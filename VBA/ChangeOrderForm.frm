VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ChangeOrderForm 
   Caption         =   "Change Order"
   ClientHeight    =   2655
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3015
   OleObjectBlob   =   "ChangeOrderForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ChangeOrderForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub UserForm_Initialize()

    ' Apply theme
    UserFormTheme Me
    
    ' Identify entry names for listbox
    Dim Table As ListObject
    Dim TableData As Variant
    
    Set Table = ThisWorkbook.Sheets("Budget Tracker").ListObjects(Form)
    TableData = Table.ListColumns(1).DataBodyRange.Value2

    ' Add entry names to the combobox
    For Each i In TableData
        ListBox1.AddItem i
    Next i
    
End Sub

Private Sub SpinButton1_SpinUp()

    ' Check that a selection was made
    If ListBox1.ListIndex = -1 Then
        MsgBox "Please select a " & Form & " to move.", vbInformation, "Input Required"
        Exit Sub
    End If
    
    ' Get selection from Listbox
    Dim selectedName As String
    selectedName = ListBox1.Value
    Dim selectedIndex As Long
    Dim tempItem As Variant
    
    ' Move the item on the Budget Tracker sheet
    ' Check that the item is not the first in the list (It can't move any higher)
    If ListBox1.ListIndex >= 1 Then
        selectedIndex = ListBox1.ListIndex
        tempItem = ListBox1.List(selectedIndex - 1)

        ' Swap the selected item with the one above it
        ListBox1.List(selectedIndex - 1) = ListBox1.List(selectedIndex)
        ListBox1.List(selectedIndex) = tempItem

        ' Update the selection to the moved item
        ListBox1.ListIndex = selectedIndex - 1
        
        ' Move the entry in the Budget Tracker sheet
        Dim Table As ListObject
        Set Table = ThisWorkbook.Sheets("Budget Tracker").ListObjects(Form)
       
        For Each Row In Table.ListRows
            If Row.Range.Cells(1, 1).Value2 = selectedName Then
                selectedIndex = Row.Index
                tempItem = Table.ListRows(selectedIndex - 1).Range.Value
                
                ' Swap the items around
                Table.ListRows(selectedIndex - 1).Range.Value = Table.ListRows(selectedIndex).Range.Value
                Table.ListRows(selectedIndex).Range.Value = tempItem
                
                Exit For
            End If
        Next Row
        
        ' Apply Accounting format (seems to set itself to Currency otherwise)
        ' Loop through each column in the table
        For Each Column In Table.ListColumns
            If Column.Range.Cells(1, 1).Value2 = "APR%" Then
                ' Set General and right aligned format for the APR data body range
                Column.DataBodyRange.NumberFormat = "General"
                Column.DataBodyRange.HorizontalAlignment = xlRight
            Else
                ' Apply the accounting format to the data body range
                Column.DataBodyRange.NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
            End If
        Next Column

    End If
    
    
    ' Move the entry in the Keystone sheet
    Set Table = ThisWorkbook.Sheets("Keystone").ListObjects("Keystone")
    
    ' Identify the row index in the keystone table
    For Each Row In Table.ListRows
        If Row.Range.Cells(1, 1).Value2 = selectedName Then
            selectedIndex = Row.Index
            Exit For
        End If
    Next Row

    ' Loop up from the selectedIndex to find the next entry of the same type.
    For RowIndex = selectedIndex - 1 To 1 Step -1
        If Table.ListRows(RowIndex).Range.Cells(1, 2).Value2 = Form Then
            tempItem = Table.ListRows(RowIndex).Range.Value
                
            ' Swap the items around
            Table.ListRows(RowIndex).Range.Value = Table.ListRows(selectedIndex).Range.Value
            Table.ListRows(selectedIndex).Range.Value = tempItem
            
            Exit For
        End If
    Next RowIndex
    
    
      
End Sub

Private Sub SpinButton1_SpinDown()

    ' Check that a selection was made
    If ListBox1.ListIndex = -1 Then
        MsgBox "Please select a " & Form & " to move.", vbInformation, "Input Required"
        Exit Sub
    End If
    
    ' Get selection from Listbox
    Dim selectedName As String
    selectedName = ListBox1.Value
    Dim selectedIndex As Long
    Dim tempItem As Variant

    ' Move the item on the Budget Tracker sheet.
    ' Check that the item is not the last entry in the list (It can't move any lower)
    If ListBox1.ListIndex < ListBox1.ListCount - 1 Then
        selectedIndex = ListBox1.ListIndex
        tempItem = ListBox1.List(selectedIndex + 1)

        ' Swap the selected item with the one above it
        ListBox1.List(selectedIndex + 1) = ListBox1.List(selectedIndex)
        ListBox1.List(selectedIndex) = tempItem

        ' Update the selection to the moved item
        ListBox1.ListIndex = selectedIndex + 1

        ' Move the entry in the Budget Tracker sheet
        Dim Table As ListObject
        Set Table = ThisWorkbook.Sheets("Budget Tracker").ListObjects(Form)
       
        For Each Row In Table.ListRows
            If Row.Range.Cells(1, 1).Value2 = selectedName Then
                selectedIndex = Row.Index
                tempItem = Table.ListRows(selectedIndex + 1).Range.Value
                
                ' Swap the items around
                Table.ListRows(selectedIndex + 1).Range.Value = Table.ListRows(selectedIndex).Range.Value
                Table.ListRows(selectedIndex).Range.Value = tempItem
                
                Exit For
            End If
        Next Row
        
        ' Apply Accounting format (seems to apply Currency format otherwise)
        ' Loop through each column in the table
        For Each Column In Table.ListColumns
            If Column.Range.Cells(1, 1).Value2 = "APR%" Then
                ' Set General and right aligned format for the APR data body range
                Column.DataBodyRange.NumberFormat = "General"
                Column.DataBodyRange.HorizontalAlignment = xlRight
            Else
                ' Apply the accounting format to the data body range
                Column.DataBodyRange.NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
            End If
        Next Column

    End If
    
    
    
    ' Move the entry in the Keystone sheet
    Set Table = ThisWorkbook.Sheets("Keystone").ListObjects("Keystone")
    
    ' Identify the row index in the keystone table
    For Each Row In Table.ListRows
        If Row.Range.Cells(1, 1).Value2 = selectedName Then
            selectedIndex = Row.Index
            Exit For
        End If
    Next Row

    ' Loop down from the selectedIndex to find the next entry of the same type.
    For RowIndex = selectedIndex + 1 To Table.ListRows.Count
        If Table.ListRows(RowIndex).Range.Cells(1, 2).Value2 = Form Then
            tempItem = Table.ListRows(RowIndex).Range.Value
                
            ' Swap the items around
            Table.ListRows(RowIndex).Range.Value = Table.ListRows(selectedIndex).Range.Value
            Table.ListRows(selectedIndex).Range.Value = tempItem
            
            Exit For
        End If
    Next RowIndex
    
    
End Sub
