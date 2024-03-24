Attribute VB_Name = "ElementHandler"
Public Sub Add(ByVal Name As String)
    
    ' Add new entry to each table e.g., add a new income, bill, savings account etc.
    Dim Table As ListObject
    Dim newRowIndex As Long
    Dim newColumn As ListColumn
    
    ' Budget Tracker Table
    Set Table = ThisWorkbook.Sheets("Budget Tracker").ListObjects(Form)
    newRowIndex = Table.ListRows.Count + 1
    Table.ListRows.Add newRowIndex
    Table.ListRows(newRowIndex).Range.Cells(1, 1).Value2 = Name
    Table.ListRows(newRowIndex).Range.Cells(1, 2).Value2 = 0

    ' Keystone Table
    Set Table = ThisWorkbook.Sheets("Keystone").ListObjects("Keystone")
    newRowIndex = Table.ListRows.Count + 1
    Table.ListRows.Add newRowIndex
    Table.ListRows(newRowIndex).Range.Cells(1, 1).Value2 = Name
    Table.ListRows(newRowIndex).Range.Cells(1, 2).Value2 = Form
    Table.ListRows(newRowIndex).Range.Cells(1, 4).Value2 = "Visible"

    ' Data Table
    Set Table = ThisWorkbook.Sheets("Data").ListObjects("Data")
    Set newColumn = Table.ListColumns.Add
    newColumn.Name = Name
    newColumn.DataBodyRange.NumberFormat = "General"
    newColumn.DataBodyRange.Value2 = 0 ' Set all values in the new column to "0"
    
End Sub

Public Sub AddAPR(ByVal Name As String, ByVal APR As Double)
    
    ' Add new entry to each table e.g., add a new mortgage, credit card, or loan.
    Dim Table As ListObject
    Dim newRowIndex As Long
    Dim newColumn As ListColumn
    
    ' Budget Tracker Table
    Set Table = ThisWorkbook.Sheets("Budget Tracker").ListObjects(Form)
    newRowIndex = Table.ListRows.Count + 1
    Table.ListRows.Add newRowIndex
    Table.ListRows(newRowIndex).Range.Cells(1, 1).Value2 = Name
    Table.ListRows(newRowIndex).Range.Cells(1, 2).Value2 = APR
    Table.ListRows(newRowIndex).Range.Cells(1, 3).Value2 = 0

    ' Keystone Table
    Set Table = ThisWorkbook.Sheets("Keystone").ListObjects("Keystone")
    newRowIndex = Table.ListRows.Count + 1
    Table.ListRows.Add newRowIndex
    Table.ListRows(newRowIndex).Range.Cells(1, 1).Value2 = Name
    Table.ListRows(newRowIndex).Range.Cells(1, 2).Value2 = Form
    Table.ListRows(newRowIndex).Range.Cells(1, 3).Value2 = APR
    Table.ListRows(newRowIndex).Range.Cells(1, 4).Value2 = "Visible"

    ' Data Table
    Set Table = ThisWorkbook.Sheets("Data").ListObjects("Data")
    Set newColumn = Table.ListColumns.Add
    newColumn.Name = Name
    newColumn.DataBodyRange.NumberFormat = "General"
    newColumn.DataBodyRange.Value2 = 0 ' Set all values in the new column to "0"
    
End Sub

Public Sub AddYear(ByVal YearInt As Integer)

    ' Find the right place in the table and then add all 12 new entries.
    Dim Table As ListObject
    Set Table = ThisWorkbook.Sheets("Data").ListObjects("Data")

    ' Generate the dates from January - December
    Dim dateArray() As Variant
    ReDim dateArray(1 To 12, 1 To 1)

    ' Populate the array with formatted dates
    For i = 1 To 12
        newDate = DateSerial(YearInt, i, 1)
        dateArray(i, 1) = Format(newDate, "mm/dd/yyyy")
    Next i

    ' Find the correct starting row to insert
    startRow = 0
    For startRow = 1 To Table.ListRows.Count
        If Table.ListRows(startRow).Range.Cells(1, 1).Value >= DateValue(dateArray(1, 1)) Then
            Exit For
        End If
    Next startRow
    
    ' Insert the dates at consecutive rows starting from the specified row
    For i = LBound(dateArray) To UBound(dateArray)
        ' Add a new Row at the specified start row
        Table.ListRows.Add startRow
        ' Apply the new date from dateArray to the new row
        Table.ListColumns(1).DataBodyRange.Cells(startRow, 1).Value2 = dateArray(i, 1)
        
        ' Set all values in the new row to 0
        For Each Column In Table.ListColumns
            If Column.Index > 1 Then
                Column.DataBodyRange.Cells(startRow, 1).Value2 = 0
            End If
        Next Column

        startRow = startRow + 1
    Next i

End Sub

Public Sub ChangeAPR(ByVal Name As String, ByVal APR As Double)

    ' Change APR in tables e.g., change the APR for Mortgage, CreditCard or Loan.
    Dim newAPR As Double
    newAPR = APR
    Dim Table As ListObject

    ' Budget Tracker Table
    Set Table = ThisWorkbook.Sheets("Budget Tracker").ListObjects(Form)

    For Each Row In Table.ListRows
        If Row.Range.Cells(1, 1).Value2 = Name Then
            Row.Range.Cells(1, 2).Value2 = newAPR
            Exit For
        End If
    Next Row

    ' Keystone Table
    Set Table = ThisWorkbook.Sheets("Keystone").ListObjects("Keystone")
    
    For Each Row In Table.ListRows
        If Row.Range.Cells(1, 1).Value2 = Name Then
            Row.Range.Cells(1, 3).Value2 = newAPR
            Exit For
        End If
    Next Row

End Sub

Public Sub ClearTables()
 
    ' Set values in all tables to 0 within the Budget Tracker sheet and remove any hidden entries.
    
    ' Identify all tables in the Budget Tracker sheet
    Dim InputSheet As Worksheet
    Set InputSheet = ThisWorkbook.Sheets("Budget Tracker")
    
    Dim ColumnToClear As Range
    For Each Table In InputSheet.ListObjects
    
        ' Set values to 0.
        If Table.ListRows.Count > 0 Then

             ' Set the reference to the range of the last column
             Set ColumnToClear = Table.ListColumns(Table.ListColumns.Count).DataBodyRange
             ' Set all values to 0
             ColumnToClear.Value = 0

        End If
    Next Table
    
    
    ' Remove hidden entries
    Dim i As Long
    Dim Keystone As ListObject
    Set Keystone = ThisWorkbook.Sheets("Keystone").ListObjects("Keystone")
    
    ' For each entry in the table check for "Hidden" status and remove if found.
    For Each Table In InputSheet.ListObjects
            
        ' Iterate backwards through each row. Looping backwards to prevent shifting issues causing row indices to change as an item is deleted.
        For i = Table.ListRows.Count To 1 Step -1
            ' Check for match in Keystone where visibility = "Hidden"
            For Each Row In Keystone.ListRows
                If Table.ListRows(i).Range.Cells(1, 1).Value2 = Row.Range.Cells(1, 1).Value2 And Row.Range.Cells(1, 4) = "Hidden" Then
                    Table.ListRows(i).Delete
                    Exit For
                End If
            Next Row
        Next i
    Next Table
    
End Sub

Public Sub GetYears()

    ' Gather years from the Data table and update the public YearCollection
    Dim Table As ListObject
    Dim dateArray() As Variant
    Dim i As Long
    
    ' Iterate through the collection and remove all items
    For i = YearCollection.Count To 1 Step -1
        YearCollection.Remove i
    Next i
    
    ' Capture dates from Data table
    Set Table = ThisWorkbook.Sheets("Data").ListObjects("Data")
    
    ' Check for rows in the Data table. If no rows are found we can assume there are no years added to the spreadsheet.
    If Table.ListRows.Count = 0 Then
        Exit Sub ' No years are found.
    Else
        dateArray = Table.ListColumns(1).DataBodyRange.Value
    End If
        
    
    ' Loop through the dates to extract the year. Only save unique years.
    On Error Resume Next
    For i = LBound(dateArray, 1) To UBound(dateArray, 1)
        ' Use CInt to convert the result of the Year function to Integer
        YearCollection.Add CInt(Year(dateArray(i, 1))), CStr(CInt(Year(dateArray(i, 1))))
    Next i
    On Error GoTo 0
    
End Sub

Public Sub Remove(ByVal selectedValue As String)
    
    ' Remove an entry to each table e.g., remove an income, bill, savings account etc.
    Dim Table As ListObject

    ' Budget Tracker Table
    Set Table = ThisWorkbook.Sheets("Budget Tracker").ListObjects(Form)

    For Each Row In Table.ListRows
        If Row.Range.Cells(1, 1).Value2 = selectedValue Then
            Row.Delete
            Exit For
        End If
    Next Row
    
    ' Keystone Table
    Set Table = ThisWorkbook.Sheets("Keystone").ListObjects("Keystone")
    
    For Each Row In Table.ListRows
        If Row.Range.Cells(1, 1).Value2 = selectedValue Then
            Row.Delete
            Exit For
        End If
    Next Row
    
    ' Data Table
    Set Table = ThisWorkbook.Sheets("Data").ListObjects("Data")
    
    For Each Column In Table.ListColumns
        If Column.Range.Cells(1, 1).Value2 = selectedValue Then
            Column.Delete
            Exit For
        End If
    Next Column

End Sub

Public Sub RemoveYear(ByVal selectedValue As String)

    ' Remove year from the Data table.
    Dim Table As ListObject
    Dim Column As ListColumn
    Dim RowMatch As Range
    Dim i As Long
    
    ' Select table and column
    Set Table = ThisWorkbook.Sheets("Data").ListObjects("Data")
    Set Column = Table.ListColumns(1)
    
    ' Iterate backward through each cell in the column's data body range
    ' Looping backwards helps prevent shifting issues causing row indices changing as they are deleted.
    For i = Column.DataBodyRange.Rows.Count To 1 Step -1
        Set Row = Column.DataBodyRange.Cells(i, 1)
            
        ' Check if the selectedValue is part of the cell's value
        If InStr(Row.Cells(1, 1).Value, selectedValue) > 0 Then
            Set RowMatch = Column.DataBodyRange.Rows(i)
            ' Delete only the row without affecting the entire column
            Table.ListRows(RowMatch.Row - Table.DataBodyRange.Row + 1).Delete
        End If
    Next i
    
End Sub

Public Sub Rename(ByVal oldName As String, ByVal Name As String)

    ' Rename the chosen entry in each table
    Dim newName As String
    newName = Name
    Dim Table As ListObject

    ' Budget Tracker Table
    Set Table = ThisWorkbook.Sheets("Budget Tracker").ListObjects(Form)
    
    For Each Row In Table.ListRows
        If Row.Range.Cells(1, 1).Value2 = oldName Then
            Row.Range.Cells(1, 1).Value2 = newName
            Exit For
        End If
    Next Row
    
    ' Keystone Table
    Set Table = ThisWorkbook.Sheets("Keystone").ListObjects("Keystone")
    
    For Each Row In Table.ListRows
        If Row.Range.Cells(1, 1).Value2 = oldName Then
            Row.Range.Cells(1, 1).Value2 = newName
            Exit For
        End If
    Next Row
    
    ' Data Table
    Set Table = ThisWorkbook.Sheets("Data").ListObjects("Data")
    
    For Each Column In Table.ListColumns
        If Column.Range.Cells(1, 1).Value2 = oldName Then
            Column.Range.Cells(1, 1).Value2 = newName
            Exit For
        End If
    Next Column

End Sub

