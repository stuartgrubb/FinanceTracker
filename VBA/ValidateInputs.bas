Attribute VB_Name = "ValidateInputs"

Public Function IsDuplicate(ByVal Name As String)

    Dim SheetNames As Variant
    SheetNames = Array("Budget Tracker", "Keystone", "Data")
    Dim Table As ListObject
    Dim TableData As Variant

    On Error Resume Next

    For Each Sheet In SheetNames
    
        ' Budget Tracker Table
        If Sheet = "Budget Tracker" Then '
            Set Table = ThisWorkbook.Sheets("Budget Tracker").ListObjects(Form)
            TableData = Table.ListColumns(1).DataBodyRange.Value2
            
        ' Keystone Table
        ElseIf Sheet = "Keystone" Then
            Set Table = ThisWorkbook.Sheets("Keystone").ListObjects("Keystone")
            TableData = Table.ListColumns(1).DataBodyRange.Value2
            
        ' Data Table
        ElseIf Sheet = "Data" Then
            Set Table = ThisWorkbook.Sheets("Data").ListObjects("Data")
            TableData = Table.HeaderRowRange
        End If
        
                
        Dim i As Variant
        
        ' If multiple entries are found (check each one)
        If IsArray(TableData) Then
            For Each i In TableData
                ' Perform a case-insensitive match
                If StrComp(i, Name, vbTextCompare) = 0 Then
                    MsgBox "Name already in use: " & i, vbInformation, "Duplicate Value Found"
                    IsDuplicate = 1 ' Match found
                    Exit Function
                End If
            Next i
        Else '(compare the entry in TableData to Name)
            ' Perform a case-insensitive match
            If StrComp(TableData, Name, vbTextCompare) = 0 Then
                MsgBox "Name already in use: " & TableData, vbInformation, "Duplicate Value Found"
                IsDuplicate = 1 ' Match found
                Exit Function
            End If
        End If
    Next Sheet
    
    On Error GoTo 0

End Function

Public Function ValidateAPR(ByVal APRstring As String, ByRef APR As Double) As Double

    ' Check for an input
    If APRstring <> "" Then
        ' Remove % sign if found
        If InStr(APRstring, "%") > 0 Then
            APRstring = Replace(APRstring, "%", "")  ' Remove %
        End If
    
        ' Check the input is numeric
        If IsNumeric(APRstring) Then
            ' Convert input to double
            APR = CDbl(APRstring) ' This APR is passed back to the sub through ByRef.
            ValidateAPR = 0 ' APR is Valid
        Else
            MsgBox "Invalid input. Please enter a valid numeric APR e.g., 4.99%", vbInformation, "Invalid Input"
            ValidateAPR = 1 ' APR is invalid
            Exit Function
        End If
    Else
        MsgBox "Please enter the APR%.", vbInformation, "Enter APR"
        ValidateAPR = 1 ' APR is invalid
        Exit Function
    End If
    
End Function

Public Function ValidateName(ByVal Name As String)

    ' Check for entry
    If Name = "" Then
        MsgBox "Please enter a name", vbInformation, "Enter Name"
        ValidateName = 1 ' No entry found
        Exit Function
    End If

    ' Check if entry is numerical only.
    If IsNumeric(Name) Then
        ValidateName = 1 ' Name is numerical only.
        MsgBox "Name should include alphabetical characters.", vbInformation, "Invalid Input"
        Exit Function
    End If
    
    ' Check for special chars
    Dim specialChars As String
    specialChars = "!@#$%^*()_+={}[]|\;:'"",.<>?/`¬£-~"
    
    Dim i As Integer
    ' Iterate through each character in the name
    For i = 1 To Len(Name)
        ' Check if the character is a special character
        If InStr(specialChars, Mid(Name, i, 1)) > 0 Then
            MsgBox "Special characters are not allowed." & vbNewLine & "Please remove the following character: " & Mid(Name, i, 1), vbInformation, "Invalid Input"
            ValidateName = 1 ' Special character found
            Exit Function
        End If
    Next i
    
End Function

Public Function ValidateYear(ByVal Year As String, ByRef YearInt As Integer) As Integer

    ' Check for entry
    If Year = "" Then
        MsgBox "Please enter a year", vbInformation, "Enter Year"
        ValidateYear = 1 ' No entry found
        Exit Function
    End If
    
    ' Check the input is numeric
    If Not IsNumeric(Year) Then
        MsgBox "Invalid input. Please enter a valid year, e.g. '2020'", vbInformation, "Invalid Input"
        ValidateYear = 1 ' YearInt is invalid
        Exit Function
    Else
        YearInt = CInt(Year) ' This YearInt is passed back to the sub through ByRef.
        
        ' Check year is within the accepted range:
        Const MinYear As Integer = 1000
        Const MaxYear As Integer = 9999
        If YearInt <= MinYear Or YearInt >= MaxYear Then
            MsgBox "Invalid input. Please enter a valid year, e.g. '2020'", vbInformation, "Invalid Input"
            ValidateYear = 1 ' YearInt is invalid
        End If
    End If
    
    ' Check that the year is not already in the spreadsheet by checking YearCollection
    For Each Item In YearCollection
        If Item = YearInt Then
            MsgBox "'" & YearInt & "' is already in this spreadsheet.", vbInformation, "Duplicate Value Found"
            ValidateYear = 1 ' Duplicate found
            Exit For
        End If
    Next Item
    
End Function

Public Sub CheckTable(ByVal tableName As String, ByVal sheetName As String)

    ' This checks to see if a table exists within the imported sheet and that it matches the correct name.

    Dim targetSheet As Worksheet
    Set targetSheet = ThisWorkbook.Sheets(sheetName)

    Dim tbl As ListObject
    Dim tblFound As Boolean
    tblFound = False

    ' Check for any tables in the worksheet.
    For Each tbl In targetSheet.ListObjects
        tblFound = True
        Exit For
    Next tbl

    ' Table was found
    If tblFound Then
    
        ' Check if the table name matches the correct table name. Rename it if not.
        If tbl.Name <> tableName Then
            tbl.Name = tableName
            tbl.DisplayName = tableName
        End If
        
    Else
    
        ' Table was not found, create it.
        Dim newTable As ListObject
        Set newTable = targetSheet.ListObjects.Add(xlSrcRange, targetSheet.UsedRange, , xlYes)
        newTable.Name = tableName
        
    End If
    
End Sub

Function SheetExists(wb As Workbook, sheetName As String) As Boolean
    
    Dim Sheet As Worksheet
    
    ' Check if the sheet matches the name provided.
    For Each Sheet In wb.Sheets
        If Sheet.Name = sheetName Then
            SheetExists = True
            Exit Function
        End If
    Next Sheet
    
    SheetExists = False
    
End Function
