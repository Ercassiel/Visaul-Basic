'Chapter 4:          Automobile Loan Calculator
    'Progarmmer:         Seth Reinhardt
    'Date:               12/1/15
    'Purpose:            To calculate the monthly payment for a car loan.
    Dim globalMonths As Double = 60.0
    Const maxLoanAmountAllowed As Integer = 25000
    Const maxLowCreditAllowed As Double = 5000


    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        'Set the maximum value of the loan amount NumericUpDown control
        nudLoanAmount.Maximum = maxLoanAmountAllowed

    End Sub

    Private Sub btnReset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnReset.Click
        nudLoanAmount.Value = 0.0
        nudRate.Value = 5.0
        fiveYears.Checked = True
        txtMonthlyPayment.Text = "$0.00"
        cmbCreditRating.SelectedIndex = -1
    End Sub

    Private Sub threeYears_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles threeYears.CheckedChanged
        globalMonths = 36
    End Sub

    Private Sub fiveYears_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles fiveYears.CheckedChanged
        globalMonths = 60
    End Sub

    Private Sub sevenYears_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles sevenYears.CheckedChanged
        globalMonths = 84
    End Sub

    Private Sub btnComputePayment_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnComputePayment.Click
        Dim dblRate, dblMonths, dblPayment As Double
        Dim dblLoanAmount As Double, strPayment As String
        Dim dblAdjustedRate As Double, strErrorMessage As String

        dblLoanAmount = Convert.ToDouble(nudLoanAmount.Value)

        If cmbCreditRating.Text = " " Then
            MessageBox.Show("Please enter the customer's credit rating in the credit rating box.", "No Credit Rating")
            cmbCreditRating.Focus()
            Exit Sub
        End If

        dblRate = (Convert.ToDouble(nudRate.Value) / 100) / 12

        Select Case cmbCreditRating.SelectedIndex
            Case 0 'credit rating of A
                dblAdjustedRate = dblRate
            Case 1 'credit rating of B
                dblAdjustedRate = dblRate * 1.1
            Case 2 'credit rating of C
                dblAdjustedRate = dblRate * 1.15
            Case 3 'credit rating of D
                dblAdjustedRate = dblRate * 1.17
            Case 4 'credit rating of E
                dblAdjustedRate = dblRate * 1.25

        End Select

        dblPayment = Pmt(dblAdjustedRate, globalMonths, -dblLoanAmount)
        strPayment = Format$(dblPayment, "Currency")
        txtMonthlyPayment.Text = strPayment

    End Sub
End Class
