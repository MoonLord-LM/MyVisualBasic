﻿Namespace My
    
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
                p.StandardInput.WriteLine("start " & """" & TaskName & """" & " >nul 2>nul")
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
        ''' <param name="MouseFocus">TaskName程序是否获得鼠标焦点</param>
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
            Dim TaskList As List(Of String) = New List(Of String)
            Dim Processes As Process() = Process.GetProcesses()
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
            Dim TaskList As List(Of String) = New List(Of String)
            Dim Processes As Process() = Process.GetProcesses()
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
        Public Shared Function ListFile() As String()
            Dim TaskList As List(Of String) = New List(Of String)
            Dim Processes As Process() = Process.GetProcesses()
            For Each P In Processes
                Try
                    TaskList.Add(P.MainModule.FileName)
                Catch ex As Exception
                    'MsgBox(P.ProcessName & ex.ToString)
                End Try
            Next
            Return TaskList.ToArray()
        End Function

    End Class

End Namespace