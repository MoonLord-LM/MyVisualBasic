Namespace My
    
    ''' <summary>
    ''' 电源管理计划相关函数
    ''' </summary>
    ''' <remarks></remarks>
    Public NotInheritable Class Power

        ''' <summary>
        ''' 设定关机计划（同步阻塞，注意会覆盖之前设定的关机计划）
        ''' </summary>
        ''' <param name="DelaySecond">延时时间（单位秒，最大值315360000，10年，默认值为0，立即执行）</param>
        ''' <returns>是否执行成功</returns>
        ''' <remarks></remarks>
        Public Shared Function ShutDown(Optional ByVal DelaySecond As Integer = 0) As Boolean
            Try
                Dim p As New Process
                p.StartInfo.FileName = "cmd.exe"
                p.StartInfo.UseShellExecute = False
                p.StartInfo.RedirectStandardInput = True
                p.StartInfo.RedirectStandardOutput = True
                p.StartInfo.RedirectStandardError = True
                p.StartInfo.CreateNoWindow = True
                p.Start()
                p.StandardOutput.ReadToEnd()
                p.StandardInput.WriteLine("shutdown -a >nul 2>nul")
                p.StandardInput.WriteLine("shutdown -s -t " & DelaySecond & " >nul 2>nul")
                p.StandardInput.WriteLine("exit")
                p.StandardOutput.ReadToEnd()
                p.StandardOutput.Close()
                p.Dispose()
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function
        ''' <summary>
        ''' 设定关机计划（异步执行，注意会覆盖之前设定的关机计划）
        ''' </summary>
        ''' <param name="DelaySecond">延时时间（单位秒，最大值315360000，10年，默认值为0，立即执行）</param>
        ''' <returns>是否执行成功</returns>
        ''' <remarks></remarks>
        Public Shared Function ShutDownAsync(Optional ByVal DelaySecond As Integer = 0) As Boolean
            Try
                Shell("shutdown -a", AppWinStyle.Hide, False)
                Shell("shutdown -s -t " & DelaySecond, AppWinStyle.Hide, False)
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function



        ''' <summary>
        ''' 设定重启计划（同步阻塞，注意会覆盖之前设定的关机计划）
        ''' </summary>
        ''' <param name="DelaySecond">延时时间（单位秒，最大值315360000，10年，默认值为0，立即执行）</param>
        ''' <returns>是否执行成功</returns>
        ''' <remarks></remarks>
        Public Shared Function Reboot(Optional ByVal DelaySecond As Integer = 0) As Boolean
            Try
                Dim p As New Process
                p.StartInfo.FileName = "cmd.exe"
                p.StartInfo.UseShellExecute = False
                p.StartInfo.RedirectStandardInput = True
                p.StartInfo.RedirectStandardOutput = True
                p.StartInfo.RedirectStandardError = True
                p.StartInfo.CreateNoWindow = True
                p.Start()
                p.StandardOutput.ReadToEnd()
                p.StandardInput.WriteLine("shutdown -a >nul 2>nul")
                p.StandardInput.WriteLine("shutdown -r -t " & DelaySecond & " >nul 2>nul")
                p.StandardInput.WriteLine("exit")
                p.StandardOutput.ReadToEnd()
                p.StandardOutput.Close()
                p.Dispose()
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function
        ''' <summary>
        ''' 设定重启计划（异步执行，注意会覆盖之前设定的关机计划）
        ''' </summary>
        ''' <param name="DelaySecond">延时时间（单位秒，最大值315360000，10年，默认值为0，立即执行）</param>
        ''' <returns>是否执行成功</returns>
        ''' <remarks></remarks>
        Public Shared Function RebootAsync(Optional ByVal DelaySecond As Integer = 0) As Boolean
            Try
                Shell("shutdown -a", AppWinStyle.Hide, False)
                Shell("shutdown -r -t " & DelaySecond, AppWinStyle.Hide, False)
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function



        ''' <summary>
        ''' 取消所有计划（同步阻塞，没有关机或重启等计划时则无效果）
        ''' </summary>
        ''' <returns>是否执行成功</returns>
        ''' <remarks></remarks>
        Public Shared Function AbortPlan() As Boolean
            Try
                Dim p As New Process
                p.StartInfo.FileName = "cmd.exe"
                p.StartInfo.UseShellExecute = False
                p.StartInfo.RedirectStandardInput = True
                p.StartInfo.RedirectStandardOutput = True
                p.StartInfo.RedirectStandardError = True
                p.StartInfo.CreateNoWindow = True
                p.Start()
                p.StandardInput.WriteLine("shutdown -a >nul 2>nul")
                p.StandardInput.WriteLine("exit")
                p.StandardOutput.ReadToEnd()
                p.StandardOutput.Close()
                p.Dispose()
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function
        ''' <summary>
        ''' 取消所有计划（异步执行，没有关机或重启等计划时则无效果）
        ''' </summary>
        ''' <returns>是否执行成功</returns>
        ''' <remarks></remarks>
        Public Shared Function AbortPlanAsync() As Boolean
            Try
                Shell("shutdown -a", AppWinStyle.Hide, False)
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function

    End Class

End Namespace