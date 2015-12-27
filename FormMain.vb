Public Class FormMain
    Private Sub FormMain_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        MsgBox(" " & My.Computer.Win32ErrorCode & " " & My.Computer.Win32ErrorMessage)
    End Sub
    Public Class User
        Public Property UserName As String = ""
        Public Property Password As String = ""
        Public Property InviteUser As User = Nothing
        Public Property Boy As Boolean = Nothing
        Public Property Age As Integer = Nothing
        Public Property Time As DateTime = Now
    End Class

    Private Sub Timer1_Tick(sender As System.Object, e As System.EventArgs) Handles Timer1.Tick
        Timer1.Enabled = False
        My.Computer.PressKey(Keys.LControlKey)
        My.Computer.PressKey(Keys.Q)
        My.Computer.ReleaseKey(Keys.Q)
        My.Computer.ReleaseKey(Keys.LControlKey)
        'My.Computer.MouseMoveToPixel(1597, 491)
        'My.Computer.MouseLeftClick()
        Timer2.Enabled = True
    End Sub

    Private Sub Timer2_Tick(sender As System.Object, e As System.EventArgs) Handles Timer2.Tick
        Timer2.Enabled = False
        My.Computer.PressKey(Keys.Q)
        My.Computer.ReleaseKey(Keys.Q)
        'My.Computer.MouseMoveToPixel(925, 689)
        'My.Computer.MouseLeftClick()
        Timer1.Enabled = True
    End Sub
End Class
