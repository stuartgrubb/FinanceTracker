Attribute VB_Name = "ThemeModule"

Public Sub ChangeSheetTheme()

    GetTheme

    Dim InputSheet As Worksheet
    Dim shp As Shape
    Set InputSheet = ThisWorkbook.Sheets("Budget Tracker")
    
    ' Loop through all shapes in the worksheet except those included in the IF statement.
    For Each shp In InputSheet.Shapes
        If shp.Name <> "CategoryShape" And shp.Name <> "RemainingBalanceGroup" And shp.Name <> "Savings Rate to Retirement" Then
            ' Apply theme
            shp.Fill.ForeColor.RGB = ShapeBackColor
            shp.TextFrame.Characters.Font.Color = ShapeFontColor
        End If
    Next shp
    
    ' Change the underline color
    Dim cell As Range
    
    ' Date selected
    Set cell = InputSheet.Range("K1:O1")
    With cell.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Color = LineColor
    End With
    
    ' Income
    Set cell = InputSheet.Range("B4:C4")
    With cell.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Color = LineColor
    End With
    
    ' Bills
    Set cell = InputSheet.Range("E4:F4")
    With cell.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Color = LineColor
    End With
    
    ' Mortgage
    Set cell = InputSheet.Range("H4:J4")
    With cell.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Color = LineColor
    End With

    ' Credit Cards
    Set cell = InputSheet.Range("L4:N4")
    With cell.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Color = LineColor
    End With

    ' Loans
    Set cell = InputSheet.Range("P4:R4")
    With cell.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Color = LineColor
    End With
    
    ' Savings Accounts
    Set cell = InputSheet.Range("T4:U4")
    With cell.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Color = LineColor
    End With
    
    ' Investments
    Set cell = InputSheet.Range("W4:X4")
    With cell.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Color = LineColor
    End With
    
End Sub

Public Sub GetTheme()

    ' Get theme
    Dim Theme As String
    Theme = ThisWorkbook.Sheets("Monthly Figures").Range("B2").Value2
    
    If Theme = "Light" Then
        FormBackColor = RGB(240, 240, 240)
        ButtonBackColor = RGB(240, 240, 240)
        BoxBackColor = RGB(255, 255, 255)
        LabelFontColor = RGB(0, 0, 0)
        FontColor = RGB(0, 0, 0)
        ShapeBackColor = RGB(91, 155, 213)
        ShapeFontColor = RGB(255, 255, 255)
        LineColor = RGB(91, 155, 213)
        
    ElseIf Theme = "Dark" Then
        FormBackColor = RGB(50, 50, 50)
        ButtonBackColor = RGB(100, 100, 100)
        BoxBackColor = RGB(50, 50, 50)
        LabelFontColor = RGB(255, 255, 255)
        FontColor = RGB(255, 255, 255)
        ShapeBackColor = RGB(50, 50, 50)
        ShapeFontColor = RGB(255, 255, 255)
        LineColor = RGB(50, 50, 50)
        
    ElseIf Theme = "Blue" Then
        FormBackColor = RGB(77, 177, 255)
        ButtonBackColor = RGB(255, 255, 255)
        BoxBackColor = RGB(255, 255, 255)
        LabelFontColor = RGB(0, 0, 0)
        FontColor = RGB(0, 0, 0)
        ShapeBackColor = RGB(77, 177, 255)
        ShapeFontColor = RGB(255, 255, 255)
        LineColor = RGB(77, 177, 255)
        
    ElseIf Theme = "Green" Then
        FormBackColor = RGB(85, 197, 149)
        ButtonBackColor = RGB(255, 255, 255)
        BoxBackColor = RGB(255, 255, 255)
        LabelFontColor = RGB(0, 0, 0)
        FontColor = RGB(0, 0, 0)
        ShapeBackColor = RGB(85, 197, 149)
        ShapeFontColor = RGB(255, 255, 255)
        LineColor = RGB(85, 197, 149)
        
    ElseIf Theme = "Purple" Then
        FormBackColor = RGB(159, 74, 238)
        ButtonBackColor = RGB(255, 255, 255)
        BoxBackColor = RGB(255, 255, 255)
        LabelFontColor = RGB(255, 255, 255)
        FontColor = RGB(0, 0, 0)
        ShapeBackColor = RGB(159, 74, 238)
        ShapeFontColor = RGB(255, 255, 255)
        LineColor = RGB(159, 74, 238)
            
    Else ' No theme found, apply light theme.
        ThisWorkbook.Sheets("Monthly Figures").Range("B2").Value2 = "Light"
        FormBackColor = RGB(240, 240, 240)
        ButtonBackColor = RGB(240, 240, 240)
        BoxBackColor = RGB(255, 255, 255)
        LabelFontColor = RGB(0, 0, 0)
        FontColor = RGB(0, 0, 0)
        ShapeBackColor = RGB(91, 155, 213)
        ShapeFontColor = RGB(255, 255, 255)
        LineColor = RGB(91, 155, 213)
        
    
    End If
    
End Sub

Public Sub UserFormTheme(UserForm As Object)

    GetTheme
    UserForm.BackColor = FormBackColor
    
    For Each Item In UserForm.Controls
    
        ' Apply button theme
        If TypeOf Item Is MSForms.CommandButton Then
            Item.BackColor = ButtonBackColor
            Item.ForeColor = FontColor
        
        ' Apply label theme
        ElseIf TypeOf Item Is MSForms.Label Then
            Item.BackColor = FormBackColor
            Item.ForeColor = LabelFontColor
        
        ' Apply textbox theme
        ElseIf TypeOf Item Is MSForms.TextBox Then
            Item.BackColor = BoxBackColor
            Item.ForeColor = FontColor
        
        ' Apply listbox theme
        ElseIf TypeOf Item Is MSForms.ListBox Then
            Item.BackColor = BoxBackColor
            Item.ForeColor = FontColor
               
        ' Apply combobox theme
        ElseIf TypeOf Item Is MSForms.ComboBox Then
            Item.BackColor = BoxBackColor
            Item.ForeColor = FontColor
        
        ' Apply checkbox theme
        ElseIf TypeOf Item Is MSForms.CheckBox Then
            Item.BackColor = FormBackColor
            Item.ForeColor = FontColor
        
        ' Apply spinbutton theme
        ElseIf TypeOf Item Is MSForms.SpinButton Then
            Item.BackColor = ButtonBackColor
            Item.ForeColor = FontColor
        
        ' Apply multipage theme (only affects the top bar next to the page buttons)
        ElseIf TypeOf Item Is MSForms.MultiPage Then
            Item.BackColor = FormBackColor
        
        ' Apply frame theme (frames are used on the select form as multipage color cannot be changed)
        ElseIf TypeOf Item Is MSForms.Frame Then
            Item.BackColor = FormBackColor
            Item.ForeColor = FontColor
        
        End If
        
    Next Item

End Sub

