Namespace My
    
    ''' <summary>
    ''' 进程管理相关函数
    ''' </summary>
    ''' <remarks></remarks>
    Partial Public NotInheritable Class Task

        ''' <summary>
        ''' 运行程序（同步阻塞，TaskName程序会获得鼠标焦点，直到TaskName程序结束才继续向下执行）
        ''' </summary>
        ''' <param name="TaskName">程序名称（例如"notepad"或"notepad.exe"）</param>
        ''' <returns>是否执行成功</returns>
        ''' <remarks></remarks>
        Public Shared Function Run(ByVal TaskName As String) As Boolean
            Try
                If TaskName.ToLower.EndsWith(".exe") = False Then
                    TaskName = TaskName & ".exe"
                End If
                Dim p As New Process
                p.StartInfo.FileName = "cmd.exe"
                p.StartInfo.UseShellExecute = False
                p.StartInfo.RedirectStandardInput = True
                p.StartInfo.RedirectStandardOutput = True
                p.StartInfo.RedirectStandardError = True
                p.StartInfo.CreateNoWindow = True
                p.Start()
                p.StandardInput.WriteLine("start """""""" " & """" & TaskName & """" & " >nul 2>nul")
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
        ''' 运行程序（异步执行，多次调用本函数会打开多个TaskName程序）
        ''' </summary>
        ''' <param name="TaskName">程序名称（例如"notepad"或"notepad.exe"）</param>
        ''' <param name="MouseFocus">TaskName程序是否获得鼠标焦点（程序运行后，有些可能会遮挡焦点窗体，如“记事本”、“画图”，有些甚至会强制抢占焦点，如“计算器”）</param>
        ''' <returns>是否执行成功</returns>
        ''' <remarks></remarks>
        Public Shared Function RunAsync(ByVal TaskName As String, Optional ByVal MouseFocus As Boolean = True) As Boolean
            Try
                If TaskName.ToLower.EndsWith(".exe") = False Then
                    TaskName = TaskName & ".exe"
                End If
                If MouseFocus Then
                    Shell(TaskName, AppWinStyle.NormalFocus, False)
                Else
                    Shell(TaskName, AppWinStyle.NormalNoFocus, False)
                End If
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function



        ''' <summary>
        ''' 关闭程序（同步阻塞，强制结束所有的TaskName程序的进程树）
        ''' </summary>
        ''' <param name="TaskName">程序名称（例如"notepad"或"notepad.exe"）</param>
        ''' <returns>是否执行成功</returns>
        ''' <remarks></remarks>
        Public Shared Function Kill(ByVal TaskName As String) As Boolean
            Try
                If TaskName.ToLower.EndsWith(".exe") = False Then
                    TaskName = TaskName & ".exe"
                End If
                Dim p As New Process
                p.StartInfo.FileName = "cmd.exe"
                p.StartInfo.UseShellExecute = False
                p.StartInfo.RedirectStandardInput = True
                p.StartInfo.RedirectStandardOutput = True
                p.StartInfo.RedirectStandardError = True
                p.StartInfo.CreateNoWindow = True
                p.Start()
                p.StandardInput.WriteLine("taskkill /f /t /im " & """" & TaskName & """" & " >nul 2>nul")
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
        ''' 关闭程序（异步执行，强制结束所有的TaskName程序的进程树）
        ''' </summary>
        ''' <param name="TaskName">程序名称（例如"notepad"或"notepad.exe"）</param>
        ''' <returns>是否执行成功</returns>
        ''' <remarks></remarks>
        Public Shared Function KillAsync(ByVal TaskName As String) As Boolean
            Try
                If TaskName.ToLower.EndsWith(".exe") = False Then
                    TaskName = TaskName & ".exe"
                End If
                Shell("taskkill /f /t /im " & """" & TaskName & """", AppWinStyle.Hide, False)
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function



        ''' <summary>
        ''' 获取进程名称列表（全部进程，包括"Idle"进程，返回结果的字符串中不含".exe"后缀）
        ''' </summary>
        ''' <returns>结果字符串数组（失败返回空String数组）</returns>
        ''' <remarks></remarks>
        Public Shared Function ListName() As String()
            Dim Processes As Process() = Process.GetProcesses()
            Dim TaskList As List(Of String) = New List(Of String)(Processes.Length)
            For Each P In Processes
                TaskList.Add(P.ProcessName)
            Next
            Return TaskList.ToArray()
        End Function
        ''' <summary>
        ''' 获取窗口标题列表（有窗体的进程，只能获取到进程的主窗体的标题）
        ''' </summary>
        ''' <returns>结果字符串数组（失败返回空String数组）</returns>
        ''' <remarks></remarks>
        Public Shared Function ListTitle() As String()
            Dim Processes As Process() = Process.GetProcesses()
            Dim TaskList As List(Of String) = New List(Of String)(Processes.Length)
            For Each P In Processes
                If P.MainWindowTitle <> "" Then
                    TaskList.Add(P.MainWindowTitle)
                End If
            Next
            Return TaskList.ToArray()
        End Function
        ''' <summary>
        ''' 获取进程文件路径列表（部分进程获取不到）
        ''' </summary>
        ''' <returns>结果字符串数组（失败返回空String数组）</returns>
        ''' <remarks></remarks>
        Public Shared Function ListFilePath() As String()
            Dim Processes As Process() = Process.GetProcesses()
            Dim TaskList As List(Of String) = New List(Of String)(Processes.Length)
            For Each P In Processes
                Try
                    TaskList.Add(P.MainModule.FileName)
                Catch ex As Exception
                    'MsgBox(P.ProcessName & ex.ToString())
                End Try
            Next
            Return TaskList.ToArray()
        End Function



        ''' <summary>
        ''' 根据进程名称，获取进程
        ''' </summary>
        ''' <param name="TaskName">进程名称（不含".exe"后缀）</param>
        ''' <returns>结果进程（Process）</returns>
        ''' <remarks></remarks>
        Public Shared Function FindByName(ByVal TaskName As String) As Process
            Dim Processes As Process() = Process.GetProcesses()
            For Each P In Processes
                If P.ProcessName.ToLower() = TaskName.ToLower() Then
                    Return P
                End If
            Next
            Return Nothing
        End Function
        ''' <summary>
        ''' 根据进程文件路径，获取进程
        ''' </summary>
        ''' <param name="FilePath">文件路径（可以是相对路径）</param>
        ''' <returns>结果进程（Process）</returns>
        ''' <remarks></remarks>
        Public Shared Function FindByFilePath(ByVal FilePath As String) As Process
            Dim Processes As Process() = Process.GetProcesses()
            If FilePath.Contains(":") = False Then
                FilePath = System.IO.Directory.GetCurrentDirectory + "\" + FilePath
            End If
            For Each P In Processes
                Try
                    If P.MainModule.FileName.ToLower() = FilePath.ToLower() Then
                        Return P
                    End If
                Catch ex As Exception
                    'MsgBox(P.ProcessName & ex.ToString())
                End Try
            Next
            Return Nothing
        End Function
        ''' <summary>
        ''' 搜索窗口标题，获取进程（优先搜索字符串完全相同的，然后搜索包含有指定字符串的）
        ''' </summary>
        ''' <param name="Title">窗口标题（字符串不必完全相同）</param>
        ''' <returns>结果进程（Process）</returns>
        ''' <remarks></remarks>
        Public Shared Function SearchByTitle(ByVal Title As String) As Process
            Dim Processes As Process() = Process.GetProcesses()
            For Each P In Processes
                If P.MainWindowTitle.ToLower() = Title.ToLower() Then
                    Return P
                End If
            Next
            For Each P In Processes
                If P.MainWindowTitle.ToLower().Contains(Title.ToLower()) Then
                    Return P
                End If
            Next
            Return Nothing
        End Function



        Private Declare Function SuspendThread Lib "kernel32.dll" Alias "SuspendThread" (ByVal hThread As IntPtr) As Int32
        Private Declare Function ResumeThread Lib "kernel32.dll" Alias "ResumeThread" (ByVal hThread As IntPtr) As Int32
        Private Declare Function OpenThread Lib "kernel32.dll" Alias "OpenThread" (ByVal dwDesiredAccess As UInt32, ByVal bInheritHandle As Boolean, ByVal dwThreadId As UInt32) As IntPtr
        <Flags()> _
        Private Enum ThreadAccess As UInt32
            Standard = &HF0000UI
            Synchronize = &H100000UI
            All = &H1F0FFFUI
        End Enum

        ''' <summary>
        ''' 挂起进程的所有线程（Suspend Count加1）
        ''' </summary>
        ''' <param name="Process">进程（Process）</param>
        ''' <returns>是否执行成功</returns>
        ''' <remarks></remarks>
        Public Shared Function ThreadSuspend(ByVal Process As Process) As Boolean
            If Process Is Nothing OrElse Process.HasExited Then
                Return False
            End If
            Dim Threads As ProcessThreadCollection = Process.Threads
            Dim Result As Boolean = True
            For I = 0 To Threads.Count - 1
                Dim ThreadId As Int32 = Threads(I).Id
                Dim ThreadhWnd As IntPtr = OpenThread(ThreadAccess.All, False, ThreadId)
                Result = Result And SuspendThread(ThreadhWnd) <> -1
            Next
            Return Result
        End Function
        ''' <summary>
        ''' 恢复进程的所有线程（Suspend Count减1）
        ''' </summary>
        ''' <param name="Process">进程（Process）</param>
        ''' <returns>是否执行成功</returns>
        ''' <remarks></remarks>
        Public Shared Function ThreadResume(ByVal Process As Process) As Boolean
            If Process Is Nothing OrElse Process.HasExited Then
                Return False
            End If
            Dim Threads As ProcessThreadCollection = Process.Threads
            Dim Result As Boolean = True
            For I = 0 To Threads.Count - 1
                Dim ThreadId As Int32 = Threads(I).Id
                Dim ThreadhWnd As IntPtr = OpenThread(ThreadAccess.All, False, ThreadId)
                Result = Result And ResumeThread(ThreadhWnd) <> -1
            Next
            Return Result
        End Function
        ''' <summary>
        ''' 强制恢复进程的所有线程（Suspend Count减到0）
        ''' </summary>
        ''' <param name="Process">进程（Process）</param>
        ''' <returns>是否执行成功</returns>
        ''' <remarks></remarks>
        Public Shared Function ThreadFree(ByVal Process As Process) As Boolean
            If Process Is Nothing OrElse Process.HasExited Then
                Return False
            End If
            Dim Threads As ProcessThreadCollection = Process.Threads
            Dim Result As Boolean = True
            For I = 0 To Threads.Count - 1
                Dim ThreadId As Int32 = Threads(I).Id
                Dim ThreadhWnd As IntPtr = OpenThread(ThreadAccess.All, False, ThreadId)
                Dim Temp As Int32 = ResumeThread(ThreadhWnd)
                Result = Result And Temp <> -1
                While Temp > 0
                    Temp = ResumeThread(ThreadhWnd)
                    Result = Result And Temp <> -1
                End While
            Next
            Return Result
        End Function

        ''' <summary>
        ''' 限制进程的CPU占用（通过不断挂起和恢复线程）
        ''' </summary>
        ''' <param name="Process">进程（Process）</param>
        ''' <param name="SleepMillisecond">挂起等待的持续时间（默认50毫秒，最少1毫秒）</param>
        ''' <param name="IntervalMillisecond">挂起恢复的间隔时间（默认50毫秒，最少1毫秒）</param>
        ''' <returns>是否执行成功</returns>
        ''' <remarks></remarks>
        Public Shared Function ThreadLimit(ByVal Process As Process, Optional ByVal SleepMillisecond As UInt32 = 50, Optional ByVal IntervalMillisecond As UInt32 = 50) As Boolean
            If Process Is Nothing OrElse Process.HasExited Then
                Return False
            End If
            If SleepMillisecond <= 0 Then
                SleepMillisecond = 1
            End If
            If IntervalMillisecond <= 0 Then
                IntervalMillisecond = 1
            End If
            Dim Temp As New ThreadLimitTask(Process, SleepMillisecond, IntervalMillisecond)
            Return True
        End Function
        Private Class ThreadLimitTask
            Private Process As Process
            Private SleepMillisecond As UInt32
            Private IntervalMillisecond As UInt32
            Private Thread As New Threading.Thread(New Threading.ThreadStart(AddressOf Run))
            Public Sub New(ByVal TaskProcess As Process, ByVal TaskSleepMillisecond As UInt32, ByVal TaskIntervalMillisecond As UInt32)
                Process = TaskProcess
                SleepMillisecond = TaskSleepMillisecond
                IntervalMillisecond = TaskIntervalMillisecond
                Thread.Start()
            End Sub
            Private Sub Run()
                While Thread.IsAlive
                    If Process.HasExited Then
                        Thread.Abort()
                    End If
                    Dim Threads As ProcessThreadCollection = Process.Threads
                    For I = 0 To Threads.Count - 1
                        Dim ThreadId As Int32 = Threads(I).Id
                        Dim ThreadhWnd As IntPtr = OpenThread(ThreadAccess.All, False, ThreadId)
                        SuspendThread(ThreadhWnd)
                    Next
                    Try
                        System.Threading.Thread.Sleep(SleepMillisecond)
                    Catch ex As Exception
                    End Try
                    For I = 0 To Threads.Count - 1
                        Dim ThreadId As Int32 = Threads(I).Id
                        Dim ThreadhWnd As IntPtr = OpenThread(ThreadAccess.All, False, ThreadId)
                        ResumeThread(ThreadhWnd)
                    Next
                    Try
                        System.Threading.Thread.Sleep(IntervalMillisecond)
                    Catch ex As Exception
                    End Try
                End While
            End Sub
        End Class

    End Class

End Namespace