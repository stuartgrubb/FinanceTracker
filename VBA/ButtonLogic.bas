Attribute VB_Name = "ButtonLogic"

Public Sub IncomeButton()

    ' Set the value of public variable "Form"
    Form = "Income"
    ' Open the Options Form
    OptionsForm.Show
    
End Sub

Public Sub BillButton()

    ' Set the value of public variable "Form"
    Form = "Bill"
    ' Open the Options Form
    OptionsForm.Show

End Sub

Public Sub MortgageButton()

    ' Set the value of public variable "Form"
    Form = "Mortgage"
    ' Open the Options Form
    OptionsForm.Show
    
End Sub

Public Sub CreditCardButton()

    ' Set the value of public variable "Form"
    Form = "CreditCard"
    ' Open the Options Form
    OptionsForm.Show
    
End Sub

Public Sub LoanButton()

    ' Set the value of public variable "Form"
    Form = "Loan"
    ' Open the Options Form
    OptionsForm.Show
    
End Sub

Public Sub SavingAccButton()

    ' Set the value of public variable "Form"
    Form = "SavingsAccount"
    ' Open the Options Form
    OptionsForm.Show
    
End Sub

Public Sub InvestmentButton()

    ' Set the value of public variable "Form"
    Form = "Investment"
    ' Open the Options Form
    OptionsForm.Show
    
End Sub

Public Sub SelectButton()

    ' Open the SelectForm()
    SelectForm.Show
     
End Sub

Public Sub SaveButton()

    ' Run the PushData subroutine
    PushData
    
End Sub
