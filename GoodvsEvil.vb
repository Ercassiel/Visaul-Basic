Dim R, L, U, D As Boolean
    Dim score As Integer

    Private Sub frmGoodEvil_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        R = True
        L = False
        U = True
        D = False
    End Sub

    Private Sub lblGood_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblGood.Click

        score = score + 1

        lblScore.Text = "Score: " & score

        If score = 5 Then
            tmrMoveGood.Enabled = False
            tmrMoveEvil.Enabled = False
            MsgBox("Winner")

        End If

    End Sub

    Private Sub lblEvil_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblEvil.Click

        score = score - 1

        lblScore.Text = "Score: " & score

        If score = -3 Then
            tmrMoveGood.Enabled = False
            tmrMoveEvil.Enabled = False
            MsgBox("Loser")

        End If
    End Sub

    Private Sub tmrMoveGood_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tmrMoveGood.Tick
        lblGood.Left = CInt(Rnd() * frmGoodEvil.ActiveForm.Width)
        lblGood.Top = CInt(Rnd() * frmGoodEvil.ActiveForm.Height)

    End Sub

    Private Sub tmrMoveEvil_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tmrMoveEvil.Tick
        lblEvil.Left = CInt(Rnd() * frmGoodEvil.ActiveForm.Width)
        lblEvil.Top = CInt(Rnd() * frmGoodEvil.ActiveForm.Height)
    End Sub
