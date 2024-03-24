Attribute VB_Name = "PushModule"
Public Sub PushData()
    
    ' Get the DateSelected from the Monthly Figures sheet.
    Dim DateToPush As Date
    If ThisWorkbook.Sheets("Monthly Figures").Range("B1").Value2 = "" Then
        MsgBox "Please select a month & year.", vbInformation, "Select Month/Year"
        Exit Sub
    Else
        DateToPush = ThisWorkbook.Sheets("Monthly Figures").Range("B1").Value2
    End If
    
    
    ' Identify the row index number that matches the DateToPush in the Data table
    Dim DataTable As ListObject
    Dim dateColumn As ListColumn
    Dim RowMatch As Integer
    
    Set DataTable = ThisWorkbook.Sheets("Data").ListObjects("Data")
    Set dateColumn = DataTable.ListColumns("Date")
    For Each Row In DataTable.ListRows
        If Row.Range.Cells(1, dateColumn.Index).Value = DateToPush Then
            RowMatch = Row.Index
            Exit For
        End If
    Next Row


    ' Identify all tables in the Budget Tracker sheet
    Dim InputSheet As Worksheet
    Set InputSheet = ThisWorkbook.Sheets("Budget Tracker")

    ' Iterate through each table and store the values in the DataCollection
    Dim dataArray(1 To 2) As Variant
    Dim DataCollection As Collection
    Set DataCollection = New Collection

    For Each Table In InputSheet.ListObjects
        ' Reset variables for each iteration
        Dim Name As String
        Dim Value As Double

        If Table.Name = "Income" Or Table.Name = "Bill" Or Table.Name = "SavingsAccount" Or Table.Name = "Investment" Then
            
            For Each Row In Table.ListRows
                
                If Row.Range.Cells(1, 1) = "" Then ' Check if the table row exists but is blank, if so delete the row. This can happen when data was entered into a row and then cleared without properly deleting the row.
                   
                   Row.Delete
                   
                ElseIf Not IsNumeric(Row.Range.Cells(1, 2)) Then
                
                    MsgBox "Table: " & Table.Name & vbNewLine & "Invalid Entry: " & Row.Range.Cells(1, 2), vbInformation, "Invalid Entry"
                    Exit Sub
                    
                Else
            
                   Name = Row.Range.Cells(1, 1).Value2
                   Value = Row.Range.Cells(1, 2).Value2
                   
                   dataArray(1) = Name
                   dataArray(2) = Value
                   
                   DataCollection.Add dataArray
                   
                End If
            Next Row
            
        ElseIf Table.Name = "Mortgage" Or Table.Name = "CreditCard" Or Table.Name = "Loan" Then
        
            For Each Row In Table.ListRows
            
                If Row.Range.Cells(1, 2) = "" Then ' Check if the table row exists but is blank, if so delete the row. This can happen when data was entered into a row and then cleared without properly deleting the row.
                   
                   Row.Delete
                
                
                ElseIf Not IsNumeric(Row.Range.Cells(1, 2)) Then
                        
                    MsgBox "Table: " & Table.Name & vbNewLine & "Invalid Entry: " & Row.Range.Cells(1, 2), vbInformation, "Invalid Entry"
                    Exit Sub
                    
                Else
            
                    Name = Row.Range.Cells(1, 1).Value2
                    Value = Row.Range.Cells(1, 3).Value2
            
                    dataArray(1) = Name
                    dataArray(2) = Value
                    
                    DataCollection.Add dataArray
                    
                End If
            Next Row
        End If
 
    Next Table
    
    
    ' Push the values to the Data Table
    For Each Item In DataCollection
    
        DataTable.ListColumns(Item(1)).DataBodyRange.Cells(RowMatch, 1).Value2 = Item(2)
        
    Next Item
    
    
    ' Clear the DateSelected from the Budget Tracker and Monthly Figures sheet
    ThisWorkbook.Sheets("Budget Tracker").Range("N1").ClearContents
    ThisWorkbook.Sheets("Monthly Figures").Range("B1").ClearContents
    
    'Hide the Remaining Balance shape, Category % shape, and MMM savings rate image.
    ThisWorkbook.Sheets("Budget Tracker").Shapes("RemainingBalanceGroup").Visible = False
    ThisWorkbook.Sheets("Budget Tracker").Shapes("CategoryShape").Visible = False
    ThisWorkbook.Sheets("Budget Tracker").Shapes("Savings Rate to Retirement").Visible = False
    
    ' Hide the Save button
    ThisWorkbook.Sheets("Budget Tracker").Shapes("SaveBtn").Visible = False
    
    ' Clear the Budget Tracker sheet
    ClearTables
    

End Sub
