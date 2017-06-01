Public Class FormMain

    Private Sub FormMain_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load

        My.IO.WriteString(My.Http.GetString("http://www.baidu.com"), "1.txt")



        MsgBox(" " & My.Computer.Win32ErrorCode & " " & My.Computer.Win32ErrorMessage)

        Dim TPM As ULong = My.Computer.Info.TotalPhysicalMemory / 1024 / 1024
        Dim APM As ULong = My.Computer.Info.AvailablePhysicalMemory / 1024 / 1024

        MsgBox("总内存" & TPM & "MB " & "可用" & APM & "MB 循环次数：" & Int(APM * 0.9 / 1024 - 1))


        Dim REF As New ArrayList
        For N = 0 To Int(APM * 0.9 / 1024 - 1)
            Dim T As Integer = 1024 * 1024 * 1024 / 4 '1GB内存
            Dim Array(T) As Int32
            For I = 0 To Array.Length - 1
                If Array(I) <> 0 Then
                    MsgBox("0异常！")
                End If
            Next
            For I = 0 To Array.Length - 1
                Array(I) = Int32.MaxValue
            Next
            For I = 0 To Array.Length - 1
                If Array(I) <> Int32.MaxValue Then
                    MsgBox("Int32.MaxValue异常！")
                End If
            Next
            For I = 0 To Array.Length - 1
                Array(I) = Int32.MinValue
            Next
            For I = 0 To Array.Length - 1
                If Array(I) <> Int32.MinValue Then
                    MsgBox("Int32.MinValue异常！")
                End If
            Next
            REF.Add(Array)
        Next

        Dim TPM2 As ULong = My.Computer.Info.TotalPhysicalMemory / 1024 / 1024
        Dim APM2 As ULong = My.Computer.Info.AvailablePhysicalMemory / 1024 / 1024

        MsgBox("可用" & APM & "MB " & "可用" & APM2 & "MB")



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
        'Timer2.Enabled = True
    End Sub

    Private Sub Timer2_Tick(sender As System.Object, e As System.EventArgs) Handles Timer2.Tick
        Timer2.Enabled = False
        My.Computer.PressKey(Keys.Q)
        My.Computer.ReleaseKey(Keys.Q)
        'My.Computer.MouseMoveToPixel(925, 689)
        'My.Computer.MouseLeftClick()
        'Timer1.Enabled = True
    End Sub

End Class
