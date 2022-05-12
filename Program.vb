Class Program

    ''' <summary>
    ''' 应用程序的主入口点。
    ''' </summary>
    ''' <remarks></remarks>
    <STAThread()> _
    Public Shared Sub Main()

        Try
            '始终捕获异常（UI线程不会因为异常而直接结束）
            Application.SetUnhandledExceptionMode(UnhandledExceptionMode.CatchException)
        Catch ex As InvalidOperationException
            '多次执行Main函数无效（只启动一次窗体）
            Return
        End Try

        AddHandler Application.ThreadException, AddressOf Application_ThreadException
        AddHandler AppDomain.CurrentDomain.UnhandledException, AddressOf CurrentDomain_UnhandledException

        Application.EnableVisualStyles()
        Application.SetCompatibleTextRenderingDefault(False)

        '设置启动窗体
        Dim FormMain = New FormMain()
        FormMain.Opacity = 0
        FormMain.Hide()
        Application.Run(FormMain)

    End Sub



    ''' <summary>
    ''' 在这里处理UI线程异常，注意：函数执行完成后，应用程序仍会继续运行
    ''' </summary>
    ''' <remarks></remarks>
    Private Shared Sub Application_ThreadException(ByVal sender As Object, ByVal e As Threading.ThreadExceptionEventArgs)
        MessageBox.Show("窗体UI线程发生异常，" & DateTime.Now.ToString() & "：" & vbCrLf & e.Exception.ToString())
    End Sub
    ''' <summary>
    ''' 在这里处理子线程异常，注意：函数执行完成后，应用程序就会被终止
    ''' </summary>
    ''' <remarks></remarks>
    Private Shared Sub CurrentDomain_UnhandledException(ByVal sender As Object, ByVal e As UnhandledExceptionEventArgs)
        MessageBox.Show("子线程发生异常，" & DateTime.Now.ToString() & "：" & vbCrLf & e.ExceptionObject.ToString())
    End Sub



    ''' <summary>
    ''' 测试异常捕获效果（触发一个UI线程异常和一个子线程异常）
    ''' </summary>
    ''' <remarks></remarks>
    Public Shared Sub TestException()
        Dim Temp As New Threading.Thread(New Threading.ThreadStart(AddressOf TestExceptionThread))
        Temp.Start()
        Throw New Exception("测试用的UI线程异常")
    End Sub
    Private Shared Sub TestExceptionThread()
        Throw New Exception("测试用的子线程异常")
    End Sub

End Class