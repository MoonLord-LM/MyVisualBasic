Public Class FormMain

    Private Sub FormMain_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load

        My.Window.SetOpacity(My.Window.FindByTitle("MyVisualBasic"), 0.222)
        MsgBox(My.Window.GetOpacity(My.Window.FindByTitle("MyVisualBasic")))
        MsgBox(0.222 * 255 & My.Window.GetOpacity(My.Window.FindByTitle("MyVisualBasic")) * 255)
        Me.Opacity = 0.222
        MsgBox(Me.Opacity)

        If (1 - 1 = 0) Then Return
        Application.Exit()


        My.IO.WriteStringArray(My.Task.ListFile(), "1.txt")
        My.IO.WriteStringArray(My.Task.ListName(), "2.txt")
        My.IO.WriteStringArray(My.Task.ListTitle(), "3.txt")


        MsgBox(" " & My.MyComputer.Win32ErrorCode & " " & My.MyComputer.Win32ErrorMessage)
        Process.GetCurrentProcess().Kill()
        MemoryTest()

        'MsgBox(My.Window.Close(My.Window.FindByTitle("无标题 - 记事本")) & My.MyComputer.Win32ErrorCode & My.MyComputer.Win32ErrorMessage)
        'MsgBox(My.Window.Close(My.Window.FindByTitle("QQ")) & My.MyComputer.Win32ErrorCode & My.MyComputer.Win32ErrorMessage)
        'MsgBox(My.Window.Close(My.Window.FindByTitle("GitHub")) & My.MyComputer.Win32ErrorCode & My.MyComputer.Win32ErrorMessage)
        'MsgBox(My.Window.Close(My.Window.FindByTitle("计算器")) & My.MyComputer.Win32ErrorCode & My.MyComputer.Win32ErrorMessage)
        'MsgBox(My.Window.Destory(My.Window.FindByTitle("MyVisualBasic")) & My.MyComputer.Win32ErrorCode & My.MyComputer.Win32ErrorMessage)
        'MsgBox(My.Window.Destory(My.Window.FindByTitle("新标签页 - 360极速浏览器")) & My.MyComputer.Win32ErrorCode & My.MyComputer.Win32ErrorMessage)
        'MsgBox(My.Window.Destory(My.Window.FindByTitle("国软2013级本科生群")) & My.MyComputer.Win32ErrorCode & My.MyComputer.Win32ErrorMessage)
        'MsgBox(My.Window.Close(My.Window.FindByTitle("无标题 - 画图")) & My.MyComputer.Win32ErrorCode & My.MyComputer.Win32ErrorMessage)
        'MsgBox(My.Window.Close(My.Window.FindByTitle("MyVisualBasic")) & My.MyComputer.Win32ErrorCode & My.MyComputer.Win32ErrorMessage)
         

        'MsgBox(My.MyComputer.Win32ErrorCode & My.MyComputer.Win32ErrorMessage)

        Process.GetCurrentProcess().Kill()
        If (1 - 1 = 0) Then Return


        My.Window.ToggleDesktop()
        My.Time.Wait(1)
        My.Window.ToggleDesktop()
        My.Time.Wait(1)
        CreateObject("Shell.Application").WindowSwitcher() '3D窗口切换（Windows Vista及更新的系统）
        My.Time.Wait(1)
        My.Screen.Image().Save("F:\Desktop\1.png")
        My.Keyboard.Click(Keys.Escape)
        My.Time.Wait(1)
        CreateObject("Shell.Application").Open("F:\Desktop") '打开指定文件夹
        My.Time.Wait(1)
        CreateObject("Shell.Application").Open("F:\Desktop\1.png") '打开指定文件
        My.Time.Wait(1)

        System.GC.Collect()
        Me.FormBorderStyle = Windows.Forms.FormBorderStyle.None
        Me.TopMost = True
        Me.SetStyle(ControlStyles.OptimizedDoubleBuffer, True) '首先在缓冲区中绘制，而不是直接绘制到屏幕上，这样可以减少闪烁
        Me.SetStyle(ControlStyles.AllPaintingInWmPaint, True) '忽略擦除背景的窗口消息，不擦除之前的背景，以减少闪烁

        Me.DoubleBuffered = True
        Timer1.Interval = 16
        Timer2.Interval = 16
        'MsgBox("")

        Process.GetCurrentProcess().Kill()
        If (1 - 1 = 0) Then Return

        Dim W As IntPtr = My.Window.FindByTitle("GitHub")
        'MsgBox(My.Window.GetRectangle(W).Left & " " & My.Window.GetRectangle(W).Top & " " & My.Window.GetRectangle(W).Width & " " & My.Window.GetRectangle(W).Height)
        My.Window.SetLocation(W, New Point(100, 100))
        My.Window.ShowNormal(W)

        'MsgBox(My.Window.GetRectangle(W).Left & " " & My.Window.GetRectangle(W).Top & " " & My.Window.GetRectangle(W).Width & " " & My.Window.GetRectangle(W).Height)

        My.Window.Image(New IntPtr(0)).Save("F:\Desktop\1.png")

        'If (1 - 1 = 0) Then Return
        'MsgBox("")
        Application.Exit()
        Process.GetCurrentProcess().Kill()

        My.Window.SetSize(W, New Size(1280, 720))
        My.Window.SetCenterScreen(W)
        'My.Window.Hide(W)
        'My.Window.SetFocus(W)
        'My.Window.SetCanRedraw(W, True)
        'My.Window.SetSize(W, New Size(1280, 600))
        'My.Window.Refresh(W)
        My.Window.Image(W).Save("F:\Desktop\1.png")



        'Dim Focus As IntPtr = My.Window.FindFocus()
        'MsgBox(Focus.ToString() & "  " & My.Window.GetTitle(Focus) & "  " & My.Window.GetTitle(Focus).Length)
        'MsgBox(My.Window.Win32ErrorCode & My.Window.Win32ErrorMessage)
        'If (1 - 1 = 0) Then Return
        Application.Exit()





        Dim A As Integer, B As Integer = A / B




    End Sub

    Public Function MemoryTest() As Boolean
        Dim MemoryTested As New ArrayList()
        While My.Computer.Info.AvailablePhysicalMemory > 128 * 1024 * 1024
            Dim Test(128 * 1024 * 1024) As UInt32
            For Each T In Test
                If T <> 0 Then
                    Throw New Exception("发现内存错误，初始0值异常，测试通过区块：" & MemoryTested.Count & " * " & Int(Test.Length / 1024 / 1024) & " MB。")
                End If
            Next
            For I = 0 To Test.Length - 1
                Test(I) = UInt32.MaxValue
            Next
            For Each T In Test
                If T <> UInt32.MaxValue Then
                    Throw New Exception("发现内存错误，重置1值异常，测试通过区块：" & MemoryTested.Count & " * " & Int(Test.Length / 1024 / 1024) & " MB。")
                End If
            Next
            For I = 0 To Test.Length - 1
                Test(I) = UInt32.MinValue
            Next
            For Each T In Test
                If T <> UInt32.MinValue Then
                    Throw New Exception("发现内存错误，重置0值异常，测试通过区块：" & MemoryTested.Count & " * " & Int(Test.Length / 1024 / 1024) & " MB。")
                End If
            Next
            MemoryTested.Add(Test)
        End While
        Return True
    End Function

    Public Class User
        Public Property UserName As String = ""
        Public Property Password As String = ""
        Public Property InviteUser As User = Nothing
        Public Property Boy As Boolean = Nothing
        Public Property Age As Integer = Nothing
        Public Property Time As DateTime = Now
    End Class

    '2秒
    Private Sub Timer1_Tick(sender As System.Object, e As System.EventArgs) Handles Timer1.Tick
        Timer1.Enabled = False

        If Not BackgroundImage Is Nothing Then BackgroundImage.Dispose()
        Me.BackgroundImage = My.Screen.ImageThumbnail(0.6)
        Me.Size = Me.BackgroundImage.Size
        Me.Location = Point.Add(My.Mouse.Position(), New Point(10, 10))
        'MsgBox(My.Window.GetTitle(My.Window.FindFocus))
        'Dim Window As IntPtr = My.Window.FindByTitle("GitHub")
        'MsgBox(My.Window.SetSize(Window, New Size(1440, 900)))
        'My.Window.Hide(Window)
        'My.Window.SetRectangle(Window, My.Computer.Screen.Bounds)

        Timer2.Enabled = True
    End Sub
    '2秒
    Private Sub Timer2_Tick(sender As System.Object, e As System.EventArgs) Handles Timer2.Tick
        Timer2.Enabled = False

        If Not Me.BackgroundImage Is Nothing Then
            Me.BackgroundImage.Dispose()
        End If
        Me.BackgroundImage = My.Screen.ImageThumbnail(0.6)
        Me.Size = Me.BackgroundImage.Size
        Me.Location = Point.Add(My.Mouse.Position(), New Point(10, 10))
        'Dim Window As IntPtr = My.Window.FindByTitle("GitHub")
        'My.Window.Show(Window)
        'MsgBox(My.Window.Show(Window))
        'MsgBox(Focus.ToString() & "  " & My.Window.GetTitle(Focus) & "  " & My.Window.GetTitle(Focus).Length & "  " & My.Window.GetFocus(My.Window.FindByTitle("FormMain")))

        Timer1.Enabled = True
    End Sub

End Class
