Dim a, b, c, d, f, g As Integer

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        For a = 0 To 100 Step 1
            listBoxOne.Items.Add(CStr(a))
        Next

        For b = 40 To 0 Step -2
            listBoxTwo.Items.Add(CStr(b))
        Next

        For c = -15 To 15 Step 1
            listBoxThree.Items.Add(CStr(c))
        Next

        For d = 50 To -50 Step -5
            listBoxFour.Items.Add(CStr(d))
        Next

        For f = 75 To 0 Step -5
            listBoxFive.Items.Add(CStr(f))
        Next

        For g = 0 To 30 Step 3
            listBoxSix.Items.Add(CStr(g))
        Next
    End Sub
