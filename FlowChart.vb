Dim score As Integer

    Private Sub initializeVariables() 'for each new customer

        score = 0
        
    End Sub

    Private Sub buttonReset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles buttonReset.Click

        'clear the customers score in the Score label box, reset score = 0
        labelScore.Text = "Score"
        textBoxAge.Text = " "
        score = 0

        'clear radio buttons
        RadioButtonYesSmoker.Checked = False
        RadioButtonNoSmoker.Checked = False
        RadioButtonYesHBP.Checked = False
        RadioButtonNoHBP.Checked = False
        radioButtonYesHFD.Checked = False
        RadioButtonNoHFD.Checked = False



    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        initializeVariables()
    End Sub


    Private Sub buttonEnter_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles buttonEnter.Click

        If textBoxAge.Text < 20 Then
            'If the user is less than 20, they gain 1 point
            score = score + 1

        ElseIf textBoxAge.Text > 19 And textBoxAge.Text < 30 Then
            'If the user is greater than 20 but less than 30, they gain 2 points
            score = score + 2

        Else
            'If the user is greater than 30, they gain 3 points
            score = score + 3
        End If

        If RadioButtonYesSmoker.Checked Then
            score = score + 4
        End If

        If RadioButtonYesHBP.Checked Then
            score = score + 2
        End If

        If radioButtonYesHFD.Checked Then
            score = score + 1
        End If

        labelScore.Text = CStr(score)

    End Sub
