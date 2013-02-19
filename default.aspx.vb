
Partial Class _default
    Inherits System.Web.UI.Page

    Protected Sub PerformCalcButton_Click(sender As Object, e As System.EventArgs) Handles PerformCalcButton.Click

        'specify constant values
        Const INTEREST_CALCS_PER_YEAR As Integer = 12
        Const PAYMENTS_PER_YEAR As Integer = 12

        'create variables to hold values enterd by the user
        Dim p As Decimal = LoanAmount.Text
        Dim r As Decimal = Rate.Text / 100
        Dim t As Decimal = MortgageLength.Text

        Dim ratePerPeriod As Decimal
        ratePerPeriod = r / INTEREST_CALCS_PER_YEAR

        Dim payPeriods As Integer = t * PAYMENTS_PER_YEAR

        Dim annualRate As Decimal = Math.Exp(INTEREST_CALCS_PER_YEAR * Math.Log(1 + ratePerPeriod)) - 1

        Dim intPerPayment As Decimal = (Math.Exp(Math.Log(annualRate + 1) / payPeriods) - 1) * payPeriods

        'now compute the total cost of the loan
        Dim intPerMonth As Decimal = intPerPayment / PAYMENTS_PER_YEAR
        Dim costPerMonth = p * intPerMonth / (1 - Math.Pow(intPerMonth + 1, -payPeriods))

        'display the results
        Results.Text = "Your mortgage payment per month is " & String.Format("{0:C}", costPerMonth)
    End Sub
End Class
