Attribute VB_Name = "PullModule"
Public Sub PullData(ByVal DateToPull As Date)

    ' Ensure all tables are clear prior to pulling data
    ClearTables

    Dim Table As ListObject
    Dim Keystone As ListObject
    Dim dateColumn As ListColumn
    Dim RowMatch As Integer

    ' Data Table
    Set Table = ThisWorkbook.Sheets("Data").ListObjects("Data")
    Set dateColumn = Table.ListColumns("Date")

    ' Identify the row index number that matches the DateToPull in the Data table
    For Each Row In Table.ListRows
        If Row.Range.Cells(1, dateColumn.Index).Value = DateToPull Then
            RowMatch = Row.Index
            Exit For
        End If
    Next Row
    


    Dim dataArray(1 To 5) As Variant
    Dim DataCollection As Collection
    Set DataCollection = New Collection
    
    ' Set Keystone Table
    Set Keystone = ThisWorkbook.Sheets("Keystone").ListObjects("Keystone")
    
    
    ' Get data for each column found in the Data Table
    For Each Column In Table.ListColumns
        If Column.Index > 1 Then ' Ignore the first column as it doesn't contain any financial values
            ' Reset variables for each iteration
            Dim DataType As String
            Dim Name As String
            Dim Value As Double
            Dim APR As Double
            Dim Visibility As String

            ' Loop through each row in the keystone table to find the name match
            For Each Row In Keystone.ListRows
                If Row.Range.Cells(1, 1).Value2 = Column.Range.Cells(1, 1).Value2 Then
                
                    ' Match found, assign values to the variables
                    DataType = Row.Range.Cells(1, 2).Value2
                    Name = Column.Range.Cells(1, 1).Value
                    Value = Column.DataBodyRange.Cells(RowMatch, 1).Value2
                    APR = Row.Range.Cells(1, 3).Value2
                    Visibility = Row.Range.Cells(1, 4).Value2
                    
                    ' Sort the data
                    dataArray(1) = DataType ' Keystone type
                    dataArray(2) = Name
                    dataArray(3) = Value
                    dataArray(4) = APR
                    dataArray(5) = Visibility
                    
                    ' Add to the DataCollection
                    DataCollection.Add dataArray

                    Exit For
                End If
            Next Row
        End If
    Next Column
        

    ' Pull the data to the Budget Tracker sheet
    Dim newRowIndex As Long
    Dim matchFound As Boolean
    
    For Each Item In DataCollection
        matchFound = False
        If Item(3) <> 0 Then ' Here we ensure that any item with a value is pulled regardless of if the visibility of the entry is "Hidden".

            ' Check if the item name is already in the Budget Tracker sheet
            Set Table = ThisWorkbook.Sheets("Budget Tracker").ListObjects(Item(1))
            
            For Each Row In Table.ListRows
                If Row.Range.Cells(1, 1).Value2 = Item(2) Then
                    matchFound = True
                    RowMatch = Row.Index
                    Exit For
                End If
            Next Row
            
            If matchFound = True Then ' Item was found in the form table
                ' Apply non-APR items
                If Item(1) = "Income" Or Item(1) = "Bill" Or Item(1) = "SavingsAccount" Or Item(1) = "Investment" Then
                    Table.DataBodyRange.Cells(RowMatch, 2).Value2 = Item(3)
                        
                ' Apply APR items
                ElseIf Item(1) = "Mortgage" Or Item(1) = "CreditCard" Or Item(1) = "Loan" Then
                    Table.DataBodyRange.Cells(RowMatch, 2).Value2 = Item(4)
                    Table.DataBodyRange.Cells(RowMatch, 3).Value2 = Item(3)
                        
                End If
                
                
            Else ' No match was found. Add new entries to the respective tables.
            
                ' Apply non-APR items
                If Item(1) = "Income" Or Item(1) = "Bill" Or Item(1) = "SavingsAccount" Or Item(1) = "Investment" Then
                
                    newRowIndex = Table.ListRows.Count + 1
                    Table.ListRows.Add newRowIndex
                    Table.ListRows(newRowIndex).Range.Cells(1, 1).Value2 = Item(2)
                    Table.ListRows(newRowIndex).Range.Cells(1, 2).Value2 = Item(3)

                ' Apply APR items
                ElseIf Item(1) = "Mortgage" Or Item(1) = "CreditCard" Or Item(1) = "Loan" Then
                    newRowIndex = Table.ListRows.Count + 1
                    Table.ListRows.Add newRowIndex
                    Table.ListRows(newRowIndex).Range.Cells(1, 1).Value2 = Item(2)
                    Table.ListRows(newRowIndex).Range.Cells(1, 2).Value2 = Item(4)
                    Table.ListRows(newRowIndex).Range.Cells(1, 3).Value2 = Item(3)

                End If
            End If
       End If
    Next Item
    
End Sub
