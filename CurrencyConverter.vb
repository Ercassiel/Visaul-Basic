Dim dollarValue, euroValue, poundsValue As Decimal
    
    Private Sub btnReset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnReset.Click

        usDollars.Text = " "
        eurosOutput.Text = " "
        poundsOutput.Text = " "

    End Sub

    Private Sub btnConvert_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnConvert.Click

        dollarValue = usDollars.Text
        euroValue = euroRate.Value
        poundsValue = poundRate.Value

        eurosOutput.Text = dollarValue * euroValue
        poundsOutput.Text = dollarValue * poundsValue

    End Sub
