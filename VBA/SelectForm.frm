VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SelectForm 
   Caption         =   "Select Month/Year"
   ClientHeight    =   4020
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3825
   OleObjectBlob   =   "SelectForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SelectForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub UserForm_Initialize()
    
    ' Apply theme
    UserFormTheme Me
    
    ' Set info label color
    Me.InfoLabel.BackColor = RGB(255, 255, 225)
    Me.InfoLabel.ForeColor = RGB(0, 0, 0)

    ' Set the initial active page of the MultiPage control
    MultiPage1.Value = 0 ' Set to the desired page index (zero-based index)

    ' Run the GetYears subroutine to capture the years from the Data table
    GetYears

    ' Populate the Comboboxes
    MonthBox.List = Array("January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December")
    ' Add each year from the Collection
    For Each Value In YearCollection
        YearBox.AddItem Value
    Next Value
     
End Sub

Private Sub OkButton_Click()

    ' Check that a month was selected
    If MonthBox.ListIndex = -1 Then
        MsgBox "Please select a month", vbInformation, "Input Required"
        Exit Sub
    End If
 
    ' Check that a year was selected
    If YearBox.ListIndex = -1 Then
        MsgBox "Please select a year", vbInformation, "Input Required"
        Exit Sub
    End If
    
    
    ' Store the selected Month and Year Value.
    Dim DateSelected As Date
    DateSelected = DateSerial(YearBox.Value, MonthBox.ListIndex + 1, 1) ' Add 1 to Month since the ListBox starts at 0.
    
    
    ' Get Month & Year Value to Pull.
    Dim DateToPull As Date
    
    ' If AutoFill was checked, select the previous month.
    If AutoFillCheckBox = True Then
    
        ' If January is selected and there is NO prior year to flip to, notify user.
        If MonthBox.ListIndex = 0 And YearBox.ListIndex = 0 Then
            MsgBox "Unable to AutoFill as there is no data prior to " & YearBox.Value, vbInformation, "AutoFill Error"
            Exit Sub
                
        ' If January is selected and there IS a prior year to flip to, set Month to December and decrement the YearBox selection by 1.
        ElseIf MonthBox.ListIndex = 0 Then
            DateToPull = DateSerial(YearBox.Value - 1, 12, 1)
        
        ' Else Select Previous Month (We do not need to decrement as the listbox starts at 0.
        Else
            DateToPull = DateSerial(YearBox.Value, MonthBox.ListIndex, 1) '
        End If
        
    Else ' AutoFill was not selected, therefore DateToPull is the same as DateSelected.
        DateToPull = DateSelected
    End If
    
    
    ' Save the DateSelected to the Budget Tracker and Monthly Figures sheet
    ThisWorkbook.Sheets("Monthly Figures").Range("B1").Value2 = DateSelected
    ThisWorkbook.Sheets("Budget Tracker").Range("N1").Value2 = DateSelected
    
    ' Pull the data to the Budget Tracker sheet using the PullData subroutine
    PullData (DateToPull)
    
    ' Present the Remaining Balance shape, Category % shape, and MMM savings rate image.
    ThisWorkbook.Sheets("Budget Tracker").Shapes("RemainingBalanceGroup").Visible = True
    ThisWorkbook.Sheets("Budget Tracker").Shapes("CategoryShape").Visible = True
    ThisWorkbook.Sheets("Budget Tracker").Shapes("Savings Rate to Retirement").Visible = True
       
    ' Present the Save button
    ThisWorkbook.Sheets("Budget Tracker").Shapes("SaveBtn").Visible = True
    
    ' Close form
    Unload Me

End Sub

'Cancel Button
Private Sub CancelButton_Click()
    Unload Me
End Sub

Private Sub AddYear_Click()

    ' Open AddYearForm
    AddYearForm.Show
    
    ' Update the YearBox from the public YearCollection
    YearBox.Clear
    GetYears
    For Each Value In YearCollection
        YearBox.AddItem Value
    Next Value
    
End Sub

Private Sub RemoveYear_Click()
    
    ' Remove Year option is only available when a month/year has not been selected
    If ThisWorkbook.Sheets("Monthly Figures").Range("B1").Value2 <> "" Then
        MsgBox "This function is not available when a month/year is selected. Please save the month/year first.", vbInformation, "Error"
        Exit Sub
    End If
        
    ' Open RemoveYearForm
    RemoveYearForm.Show
    
    ' Update the YearBox from the public YearCollection
    YearBox.Clear
    GetYears
    For Each Value In YearCollection
        YearBox.AddItem Value
    Next Value
    
End Sub

Private Sub ChangeTheme_Click()

    ' Open ChangeThemeForm
    ChangeThemeForm.Show
    Unload Me
    SelectForm.Show

End Sub

Private Sub ExportData_Click()

    ' Ensure a month/year has not been selected. Prompt user to save before exporting.
    If ThisWorkbook.Sheets("Monthly Figures").Range("B1").Value2 <> "" Then
        MsgBox "Unable to export. Please save the month/year first.", vbInformation, "Error"
        Exit Sub
    End If

    ' Export specific sheets to xlsx
    Dim ExportFilePath As String
    Dim NewWorkbook As Workbook
    Dim Keystone As Worksheet
    Dim Data As Worksheet
    
    Set Keystone = ThisWorkbook.Sheets("Keystone")
    Set Data = ThisWorkbook.Sheets("Data")
    
    ' Create a new workbook
    Set NewWorkbook = Workbooks.Add

    ' Unhide the sheets (if the sheets are not made visible, deleting the default sheet from the newworkbook fails)
    Keystone.Visible = xlSheetVisible
    Data.Visible = xlSheetVisible

    ' Copy sheets
    Keystone.Copy Before:=NewWorkbook.Sheets(NewWorkbook.Sheets.Count)
    Data.Copy After:=NewWorkbook.Sheets(NewWorkbook.Sheets.Count)
    
    ' Delete the default "Sheet1" from the new workbook.
    Application.DisplayAlerts = False ' Turn off alerts to avoid confirmation prompt
    NewWorkbook.Sheets("Sheet1").Delete
    Application.DisplayAlerts = True ' Turn on alerts

    ' Hide the sheets
    Keystone.Visible = xlSheetHidden
    Data.Visible = xlSheetHidden


    ' Get the current date for the filename
    Dim currentDate As String
    currentDate = Format(Date, "dd-mm-yyyy")

    ' Prompt the user for the file path to save

    ExportFilePath = Application.GetSaveAsFilename(InitialFileName:="Finance Tracker Backup " & currentDate, FileFilter:="Excel Workbook (*.xlsx), *.xlsx")

    ' Check if the user canceled the operation
    If ExportFilePath = "False" Then
        ' Close the new workbook without saving.
        NewWorkbook.Close SaveChanges:=False
        Exit Sub
    End If
    
    ' Save the new workbook to the specified path
    NewWorkbook.SaveAs ExportFilePath
    NewWorkbook.Close
    
    ' Close the SelectForm
    Unload SelectForm
    
End Sub

Private Sub ImportData_Click()

    ' Ensure a month/year has not been selected. Prompt user to save before importing
    If ThisWorkbook.Sheets("Monthly Figures").Range("B1").Value2 <> "" Then
        MsgBox "Unable to import. Please save the month/year first.", vbInformation, "Error"
        Exit Sub
    End If

    ' Prompt user to select the spreadsheet to import
    Dim ImportFilePath As String
    ImportFilePath = Application.GetOpenFilename("Excel Files (*.xlsx; *.xlsm), *.xlsx; *.xlsm")
    
    ' Check if user canceled the operation
    If ImportFilePath = "False" Then
        Exit Sub
    End If

    ' Open the import file
    Dim ImportFile As Workbook
    Set ImportFile = Workbooks.Open(ImportFilePath)
    
    ' Check that the Keystone and Data sheets exist in the import file using the SheetExists function.
    Dim KeystoneExists As Boolean
    Dim DataExists As Boolean
    KeystoneExists = SheetExists(ImportFile, "Keystone")
    DataExists = SheetExists(ImportFile, "Data")



    ' Import the data if both sheets exist in the importfile (Keystone and Data)
    If KeystoneExists And DataExists Then
        
        Dim sourceSheet As Worksheet
        Dim targetSheet As Worksheet
        
        For Each Sheet In ImportFile.Sheets ' Copy all sheets from the importfile (in case more sheets are added in future updates).
        
            Set sourceSheet = ImportFile.Sheets(Sheet.Name)
        
            ' Delete the sheet in this workbook to prevent duplicate sheets
            Application.DisplayAlerts = False ' Turn off alerts to avoid confirmation prompt
            Set targetSheet = ThisWorkbook.Sheets(Sheet.Name)
            targetSheet.Delete
            Application.DisplayAlerts = True ' Turn on alerts
            
            ' Copy the source sheet to the this workbook.
            sourceSheet.Copy Before:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
                           
            ' Hide the sheet
            ThisWorkbook.Sheets(Sheet.Name).Visible = xlSheetHidden
        
        Next Sheet
    
    Else ' Display error message. Specify which sheet(s) were missing.
    
        Dim ErrorMessage As String
        
        If Not KeystoneExists And Not DataExists Then
            ErrorMessage = "Import failed. The following sheets were not found: " & vbNewLine & "- Keystone" & vbNewLine & "- Data" & vbNewLine & vbNewLine & "Please check the correct import file was selected."
        ElseIf Not KeystoneExists Then
            ErrorMessage = "Import failed. The following sheet was not found: " & vbNewLine & "- Keystone"
        ElseIf Not DataExists Then
            ErrorMessage = "Import failed. The following sheet was not found: " & vbNewLine & "- Data"
        Else
            ErrorMessage = "Import failed. Please check the import file and try again."
        End If
    
        MsgBox ErrorMessage, vbInformation, "Error"
        ImportFile.Close SaveChanges:=False
        
        ' Close the SelectForm
        Unload SelectForm
        Exit Sub
        
    End If
       
       
       
    ' Run the checktable subroutine. This checks to see if a table exists within the imported sheet and that it matches the correct name.
    CheckTable "Keystone", "Keystone"
    CheckTable "Data", "Data"

    ' Close the source file without saving.
    ImportFile.Close SaveChanges:=False
    
    

    ' Clear all entries from the Budget Tracker sheet to prevent save issues. If an item was added to the Budget Tracker sheet that wasn't included in the import file, users will be unable to save.
    Dim i As Long
    Dim InputSheet As Worksheet
    Set InputSheet = ThisWorkbook.Sheets("Budget Tracker")
    
    For Each Table In InputSheet.ListObjects
        ' Iterate backwards through each row. Looping backwards to prevent shifting issues causing row indices to change as an item is deleted.
        For i = Table.ListRows.Count To 1 Step -1
            Table.ListRows(i).Delete
        Next i
    Next Table
        
        
        
    ' Identify visible entries from the keystone table
    Dim dataArray(1 To 4) As Variant
    Dim DataCollection As Collection
    Set DataCollection = New Collection

    Dim Keystone As ListObject
    Set Keystone = ThisWorkbook.Sheets("Keystone").ListObjects("Keystone")
    
    For Each Row In Keystone.ListRows
    
        ' Reset variables for each iteration
        Dim DataType As String
        Dim Name As String
        Dim Value As Double
        Dim APR As Double
        Dim Visibility As String
    
        If Row.Range.Cells(1, 4).Value2 = "Visible" Then
        
            ' Assign values to the variables
            DataType = Row.Range.Cells(1, 2).Value2
            Name = Row.Range.Cells(1, 1).Value
            Value = 0
            APR = Row.Range.Cells(1, 3).Value2
            
            ' Sort data
            dataArray(1) = DataType ' Keystone type
            dataArray(2) = Name
            dataArray(3) = Value
            dataArray(4) = APR
            
            ' Add to the DataCollection
            DataCollection.Add dataArray
        
        End If
    Next Row
        
    ' Pull visible keystone entries to the Budget Tracker sheet
    For Each Item In DataCollection
        ' Set the correct table to add the entries to
        Set Table = ThisWorkbook.Sheets("Budget Tracker").ListObjects(Item(1))
                
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
        
    Next Item
    
    
    MsgBox "Import successful", , "Success"
    ' Close the SelectForm
    Unload Me

End Sub
