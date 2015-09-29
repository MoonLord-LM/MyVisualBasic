Public Class FormMain
    Private Sub FormMain_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        Dim A As New User
        A.UserName = "ABC" & vbCrLf
        A.Password = "123"
        Dim B As New User
        B.UserName = "DEF"
        B.Password = "456"
        A.InviteUser = B
        Dim C As New User
        B.InviteUser = C
        Dim D As New List(Of User)
        D.Add(C)
        Clipboard.Clear()
        'Clipboard.SetText(My.StringData.ChangeObjectToJson(A))
        'MsgBox(My.StringData.ChangeObjectToJson(A))
        'My.MyApplication.TestException()
        Throw (New Exception(System.Web.HttpUtility.UrlEncodeUnicode(" 《》《》》<><><''''''你妹啊啊啊啊      啊2323你妹啊啊啊啊啊2323你妹啊啊啊啊啊2323你妹啊啊啊啊啊2323你妹啊啊啊啊啊2323你妹啊啊啊啊啊2323你妹啊啊啊啊啊2323你妹啊啊啊啊啊2323你妹啊啊啊啊啊2323你妹啊啊啊啊啊2323你妹啊啊啊啊啊2323你妹啊啊啊啊啊2323你妹啊啊啊啊啊2323你妹啊啊啊啊啊2323你妹啊啊啊啊啊2323你妹啊啊啊啊啊2323你妹啊啊啊啊啊2323你妹啊啊啊啊啊2323你妹啊啊啊啊啊2323你妹啊啊啊啊啊2323你妹啊啊啊啊啊2323你妹啊啊啊啊啊2323你妹啊啊啊啊啊2323你妹啊啊啊啊啊2323你妹啊啊啊啊啊2323你妹啊啊啊啊啊2323你妹啊啊啊啊啊2323")))
        'MsgBox(Json.ChangeObjectToString(D))
        'MsgBox(Json.ChangeObjectToString(New User() {C, C}))
        'MsgBox(My.StringData.ChangeObjectToJson(New String() {"123", "456", "789"}))
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
        Dim W As System.Drawing.Rectangle = My.Computer.FindWindow("内存清理.bat - 记事本")
        W = My.Computer.FindFocusWindow
        Console.Write("左上角坐标(" & W.Left & "," & W.Top & ")" & vbCrLf & "右下角坐标(" & W.Right & "," & W.Bottom & ")" & vbCrLf & "窗口高" & W.Height & "窗口宽" & W.Width)
        My.Computer.MoveWindow("【2015】", -10, -10)
        My.Computer.ResizeWindow("【2015】", 100, 100)
        My.Computer.MouseMove(10, 10)
        Me.BackColor = My.Computer.MousePositionColor
        Me.Text = "A:" & My.Computer.MousePositionColor.A & "R:" & My.Computer.MousePositionColor.R & "G:" & My.Computer.MousePositionColor.G & "B:" & My.Computer.MousePositionColor.B
    End Sub

End Class
