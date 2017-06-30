Public Class FormMain

    Private Sub FormMain_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load

        If (1 - 1 = 0) Then Return
        Dim A As Integer, B As Integer = A / B
        Application.Exit()
        Process.GetCurrentProcess().Kill()

    End Sub

    '2秒
    Private Sub Timer1_Tick(sender As System.Object, e As System.EventArgs) Handles Timer1.Tick
        Timer1.Enabled = False
        Timer2.Enabled = True
    End Sub
    '2秒
    Private Sub Timer2_Tick(sender As System.Object, e As System.EventArgs) Handles Timer2.Tick
        Timer2.Enabled = False
        Timer1.Enabled = True
    End Sub

End Class
