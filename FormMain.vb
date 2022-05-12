Public Class FormMain

    Private Sub FormMain_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        Dim Thread As New Threading.Thread(New Threading.ThreadStart(AddressOf My.Example.MemoryTest))
        Thread.Start()
        Application.Exit()

        If (1 - 1 = 0) Then Return

        Dim A As Integer, B As Integer = A / B
        Application.Exit()
        Process.GetCurrentProcess().Kill()

    End Sub

    '2秒
    Private Sub Timer1_Tick(sender As System.Object, e As System.EventArgs) Handles Timer1.Tick
        Timer1.Enabled = False

        Dim 阴阳师 As IntPtr = My.Window.FindByTitle("阴阳师-网易游戏")
        If 阴阳师 <> IntPtr.Zero Then
            My.Window.SetCenterScreen(阴阳师)
            My.Window.SetFocus(阴阳师)
            My.Mouse.MoveToPixel(950, 500)
            My.Mouse.LeftClick()
            My.Mouse.LeftClick()
            My.Mouse.MoveToPixel(1380, 750)
            My.Mouse.LeftClick()
        End If

        Timer2.Enabled = True
    End Sub

    '2秒
    Private Sub Timer2_Tick(sender As System.Object, e As System.EventArgs) Handles Timer2.Tick
        Timer2.Enabled = False
        Timer1.Enabled = True
    End Sub

End Class
