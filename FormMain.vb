Public Class FormMain

    Private Sub FormMain_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        'CreateObject("Wscript.Shell").SendKeys("^+{ESC}")
        'keybd_event(Keys.ControlKey, 0, 0, 0)
        'keybd_event(Keys.ShiftKey, 0, 0, 0)
        'keybd_event(Keys.Escape, 0, 0, 0)
        'keybd_event(Keys.Escape, 0, KEYEVENTF_KEYUP, 0)
        'keybd_event(Keys.ShiftKey, 0, KEYEVENTF_KEYUP, 0)
        'keybd_event(Keys.ControlKey, 0, KEYEVENTF_KEYUP, 0)
        'My.Computer.SendKeys("^+{ESC}")
    End Sub

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        'My.Computer.MoveClick(100 / 1920 * 65536, 100 / 1080 * 65536)
        My.Computer.MouseMoveToPercent(0.3, 0.3)
        My.Computer.MouseLeftDoubleClick()
        My.Computer.MouseWheelDown(1000)
    End Sub
End Class
