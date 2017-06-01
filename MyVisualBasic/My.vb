Namespace My

    ''' <summary>
    ''' 访问主机及其资源、服务和数据
    ''' </summary>
    ''' <remarks></remarks>
    Partial Class MyComputer
        ''' <summary>
        ''' 延时关闭计算机（注意会取代之前可能存在的关机计划）
        ''' </summary>
        ''' <param name="DelaySecond">延时时间（单位秒，最多可以延时10年）</param>
        ''' <returns>是否执行成功</returns>
        ''' <remarks></remarks>
        Public Function ShutDown(Optional ByVal DelaySecond As Integer = 0) As Boolean
            Try
                Dim p As New Process
                p.StartInfo.FileName = "cmd.exe"
                p.StartInfo.UseShellExecute = False
                p.StartInfo.RedirectStandardInput = True
                p.StartInfo.RedirectStandardOutput = True
                p.StartInfo.RedirectStandardError = True
                p.StartInfo.CreateNoWindow = True
                p.Start()
                p.StandardOutput.ReadLine() 'MsgBox(p.StandardOutput.ReadLine()) 'Microsoft Windows
                p.StandardOutput.ReadLine() 'MsgBox(p.StandardOutput.ReadLine()) '版权所有
                p.StandardOutput.ReadLine() 'MsgBox(p.StandardOutput.ReadLine()) '空行
                p.StandardInput.WriteLine("shutdown -a >nul 2>nul")
                p.StandardInput.WriteLine("shutdown -s -t " & DelaySecond & " >nul 2>nul")
                p.StandardInput.WriteLine("exit")
                p.StandardOutput.ReadToEnd()
                p.StandardOutput.Close()
                p.Dispose()
                '因为已有关机计划，这一句不会起作用（仅供参考写法）
                'Shell("shutdown -s -t " & DelaySecond, AppWinStyle.Hide)
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function
        ''' <summary>
        ''' 延时重启计算机（注意会取代之前可能存在的关机计划）
        ''' </summary>
        ''' <param name="DelaySecond">延时时间（单位秒，最多可以延时10年）</param>
        ''' <returns>是否执行成功</returns>
        ''' <remarks></remarks>
        Public Function ShutDownReboot(Optional ByVal DelaySecond As Integer = 0) As Boolean
            Try
                Dim p As New Process
                p.StartInfo.FileName = "cmd.exe"
                p.StartInfo.UseShellExecute = False
                p.StartInfo.RedirectStandardInput = True
                p.StartInfo.RedirectStandardOutput = True
                p.StartInfo.RedirectStandardError = True
                p.StartInfo.CreateNoWindow = True
                p.Start()
                p.StandardOutput.ReadLine() 'MsgBox(p.StandardOutput.ReadLine()) 'Microsoft Windows
                p.StandardOutput.ReadLine() 'MsgBox(p.StandardOutput.ReadLine()) '版权所有
                p.StandardOutput.ReadLine() 'MsgBox(p.StandardOutput.ReadLine()) '空行
                p.StandardInput.WriteLine("shutdown -a >nul 2>nul")
                p.StandardInput.WriteLine("shutdown -r -t " & DelaySecond & " >nul 2>nul")
                p.StandardInput.WriteLine("exit")
                p.StandardOutput.ReadToEnd()
                p.StandardOutput.Close()
                p.Dispose()
                '因为已有关机计划，这一句不会起作用（仅供参考写法）
                'Shell("shutdown -r -t " & DelaySecond, AppWinStyle.Hide)
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function
        ''' <summary>
        ''' 取消关机计划（没有关机计划时则无效果）
        ''' </summary>
        ''' <returns>是否执行成功</returns>
        ''' <remarks></remarks>
        Public Function ShutDownAbort() As Boolean
            Try
                Dim p As New Process
                p.StartInfo.FileName = "cmd.exe"
                p.StartInfo.UseShellExecute = False
                p.StartInfo.RedirectStandardInput = True
                p.StartInfo.RedirectStandardOutput = True
                p.StartInfo.RedirectStandardError = True
                p.StartInfo.CreateNoWindow = True
                p.Start()
                p.StandardOutput.ReadLine() 'MsgBox(p.StandardOutput.ReadLine()) 'Microsoft Windows
                p.StandardOutput.ReadLine() 'MsgBox(p.StandardOutput.ReadLine()) '版权所有
                p.StandardOutput.ReadLine() 'MsgBox(p.StandardOutput.ReadLine()) '空行
                p.StandardInput.WriteLine("shutdown -a >nul 2>nul")
                p.StandardInput.WriteLine("exit")
                p.StandardOutput.ReadToEnd()
                p.StandardOutput.Close()
                p.Dispose()
                '因为没有关机计划，这一句不会起作用（仅供参考写法）
                'Shell("shutdown -a >nul 2>nul", AppWinStyle.Hide)
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function
        ''' <summary>
        ''' 打开指定的程序（多次调用本函数会打开程序的多个实例，新打开的程序会夺取鼠标焦点）
        ''' </summary>
        ''' <param name="TaskName">程序名称（例如"notepad"或"notepad.exe"）</param>
        ''' <returns>是否执行成功</returns>
        ''' <remarks></remarks>
        Public Function TaskRun(ByVal TaskName As String) As Boolean
            Try
                If TaskName.ToLower.EndsWith(".exe") = False Then TaskName = TaskName & ".exe"
                Dim p As New Process
                p.StartInfo.FileName = "cmd.exe"
                p.StartInfo.UseShellExecute = False
                p.StartInfo.RedirectStandardInput = True
                p.StartInfo.RedirectStandardOutput = True
                p.StartInfo.RedirectStandardError = True
                p.StartInfo.CreateNoWindow = True
                p.Start()
                p.StandardOutput.ReadLine() 'MsgBox(p.StandardOutput.ReadLine()) 'Microsoft Windows
                p.StandardOutput.ReadLine() 'MsgBox(p.StandardOutput.ReadLine()) '版权所有
                p.StandardOutput.ReadLine() 'MsgBox(p.StandardOutput.ReadLine()) '空行
                p.StandardInput.WriteLine(TaskName)
                p.StandardOutput.ReadLine()
                '以下这种写法会导致线程一直阻塞，直到TaskName程序结束才继续执行（CMD认为执行的程序可能会无限输出数据，所以一直等待？？？）：
                'p.StandardInput.WriteLine("exit")
                'p.StandardOutput.ReadToEnd()
                p.StandardOutput.Close()
                p.Dispose()
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function
        ''' <summary>
        ''' 打开指定的程序（多次调用本函数会打开程序的多个实例，新打开的程序不会夺取鼠标焦点）
        ''' </summary>
        ''' <param name="TaskName">程序名称（例如"notepad"或"notepad.exe"）</param>
        ''' <returns>是否执行成功</returns>
        ''' <remarks></remarks>
        Public Function ShellRun(ByVal TaskName As String) As Boolean
            Try
                If TaskName.ToLower.EndsWith(".exe") = False Then TaskName = TaskName & ".exe"
                Interaction.Shell(TaskName, AppWinStyle.NormalNoFocus)
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function
        ''' <summary>
        ''' 打开指定的命令行程序（多次调用本函数会打开程序的多个实例，新打开的程序不会夺取鼠标焦点）
        ''' </summary>
        ''' <param name="TaskName">完整的命令行语句（例如，应当使用"notepad.exe"而不是"notepad"）</param>
        ''' <returns>是否执行成功</returns>
        ''' <remarks></remarks>
        Public Function Shell(ByVal TaskName As String) As Boolean
            Try
                Interaction.Shell(TaskName, AppWinStyle.NormalNoFocus)
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function
        ''' <summary>
        ''' 通过进程的名称，获取进程的文件的完整路径（注意，该程序必须在运行中）
        ''' </summary>
        ''' <param name="ProcessName">程序名称（例如"notepad"或"notepad.exe"）</param>
        ''' <returns>成功返回完整的文件路径，失败返回进程名称（例如"notepad.exe"）</returns>
        ''' <remarks></remarks>
        Public Function GetProcessFilePath(ByVal ProcessName As String) As String
            If ProcessName.ToLower.EndsWith(".exe") Then
                ProcessName = ProcessName.Substring(0, ProcessName.Length - 4)
            End If
            Dim processesByName As Process() = Process.GetProcessesByName(ProcessName)
            If (processesByName.Length > 0) Then
                Return processesByName(0).MainModule.FileName
            End If
            Return ProcessName
        End Function
        ''' <summary>
        ''' 关闭指定的程序（程序如果有多个实例，会一并结束，多次调用本函数无特别效果）
        ''' </summary>
        ''' <param name="TaskName">程序名称（例如"notepad"或"notepad.exe"）</param>
        ''' <returns>是否执行成功</returns>
        ''' <remarks></remarks>
        Public Function TaskKill(ByVal TaskName As String) As Boolean
            Try
                If TaskName.ToLower.EndsWith(".exe") = False Then TaskName = TaskName & ".exe"
                TaskName = """" + TaskName + """"
                Dim p As New Process
                p.StartInfo.FileName = "cmd.exe"
                p.StartInfo.UseShellExecute = False
                p.StartInfo.RedirectStandardInput = True
                p.StartInfo.RedirectStandardOutput = True
                p.StartInfo.RedirectStandardError = True
                p.StartInfo.CreateNoWindow = True
                p.Start()
                p.StandardOutput.ReadLine() 'MsgBox(p.StandardOutput.ReadLine()) 'Microsoft Windows
                p.StandardOutput.ReadLine() 'MsgBox(p.StandardOutput.ReadLine()) '版权所有
                p.StandardOutput.ReadLine() 'MsgBox(p.StandardOutput.ReadLine()) '空行
                p.StandardInput.WriteLine("taskkill /f /t /im " & TaskName)
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
        ''' 关闭指定的程序（程序如果有多个实例，会一并结束，多次调用本函数无特别效果）
        ''' </summary>
        ''' <param name="TaskName">程序名称（例如"notepad"或"notepad.exe"）</param>
        ''' <returns>是否执行成功</returns>
        ''' <remarks></remarks>
        Public Function ShellKill(ByVal TaskName As String) As Boolean
            Try
                If TaskName.ToLower.EndsWith(".exe") = False Then TaskName = TaskName & ".exe"
                TaskName = """" + TaskName + """"
                Interaction.Shell("taskkill /f /t /im " & TaskName, AppWinStyle.Hide)
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function
        ''' <summary>
        ''' 获取屏幕截图（System.Drawing.Bitmap类型）
        ''' </summary>
        ''' <returns>成功返回屏幕截图的Bitmap，失败返回新建的0*0像素的Bitmap</returns>
        ''' <remarks></remarks>
        Public Function SaveScreen() As System.Drawing.Bitmap
            Try
                Dim MyRectangle As Rectangle = My.Computer.Screen.Bounds
                Dim MyBmp As New Bitmap(MyRectangle.Width, MyRectangle.Height)
                Using MyGraphics As Graphics = Graphics.FromImage(MyBmp)
                    MyGraphics.CopyFromScreen(0, 0, 0, 0, MyRectangle.Size)
                End Using
                Return MyBmp
            Catch ex As Exception
                Return New Bitmap(0, 0)
            End Try
        End Function
        ''' <summary>
        ''' 获取指定区域的屏幕截图（System.Drawing.Bitmap类型）
        ''' </summary>
        ''' <param name="Position">指定区域</param>
        ''' <returns>成功返回指定区域的屏幕截图的Bitmap，失败返回新建的Position.Width*Position.Height像素的Bitmap</returns>
        ''' <remarks></remarks>
        Public Function SaveScreen(ByVal Position As System.Drawing.Rectangle) As System.Drawing.Bitmap
            Try
                Dim MyBmp As New Bitmap(Position.Width, Position.Height)
                Using MyGraphics As Graphics = Graphics.FromImage(MyBmp)
                    MyGraphics.CopyFromScreen(Position.Left, Position.Top, 0, 0, Position.Size)
                End Using
                Return MyBmp
            Catch ex As Exception
                Return New Bitmap(Position.Width, Position.Height)
            End Try
        End Function
        ''' <summary>
        ''' 获取屏幕截图（默认保存格式，实测1920*1080分辨率的截图文件大小为293K，文件大小中等）
        ''' </summary>
        ''' <param name="ScreenFilePath">截图文件保存路径</param>
        ''' <returns>是否执行成功</returns>
        ''' <remarks></remarks>
        Public Function SaveScreen(ByVal ScreenFilePath As String) As Boolean
            Try
                Dim MyRectangle As Rectangle = My.Computer.Screen.Bounds
                Dim MyBmp As New Bitmap(MyRectangle.Width, MyRectangle.Height)
                Using MyGraphics As Graphics = Graphics.FromImage(MyBmp)
                    MyGraphics.CopyFromScreen(0, 0, 0, 0, MyRectangle.Size)
                End Using
                MyBmp.Save(ScreenFilePath)
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function
        ''' <summary>
        ''' 获取屏幕截图（bmp格式，实测1920*1080分辨率的截图文件大小为7.91M，文件大小最大）
        ''' </summary>
        ''' <param name="ScreenFilePath">截图文件保存路径</param>
        ''' <returns>是否执行成功</returns>
        ''' <remarks></remarks>
        Public Function SaveScreenBmp(ByVal ScreenFilePath As String) As Boolean
            Try
                Dim MyRectangle As Rectangle = My.Computer.Screen.Bounds
                Dim MyBmp As New Bitmap(MyRectangle.Width, MyRectangle.Height)
                Using MyGraphics As Graphics = Graphics.FromImage(MyBmp)
                    MyGraphics.CopyFromScreen(0, 0, 0, 0, MyRectangle.Size)
                End Using
                MyBmp.Save(ScreenFilePath, System.Drawing.Imaging.ImageFormat.Bmp)
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function
        ''' <summary>
        ''' 获取屏幕截图（png格式，实测1920*1080分辨率的截图文件大小为293K，文件大小中等）
        ''' </summary>
        ''' <param name="ScreenFilePath">截图文件保存路径</param>
        ''' <returns>是否执行成功</returns>
        ''' <remarks></remarks>
        Public Function SaveScreenPng(ByVal ScreenFilePath As String) As Boolean
            Try
                Dim MyRectangle As Rectangle = My.Computer.Screen.Bounds
                Dim MyBmp As New Bitmap(MyRectangle.Width, MyRectangle.Height)
                Using MyGraphics As Graphics = Graphics.FromImage(MyBmp)
                    MyGraphics.CopyFromScreen(0, 0, 0, 0, MyRectangle.Size)
                End Using
                MyBmp.Save(ScreenFilePath, System.Drawing.Imaging.ImageFormat.Png)
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function
        ''' <summary>
        ''' 获取屏幕截图（jpeg格式，实测1920*1080分辨率的截图文件大小为212K，文件大小最小）
        ''' </summary>
        ''' <param name="ScreenFilePath">截图文件保存路径</param>
        ''' <returns>是否执行成功</returns>
        ''' <remarks></remarks>
        Public Function SaveScreenJpeg(ByVal ScreenFilePath As String) As Boolean
            Try
                Dim MyRectangle As Rectangle = My.Computer.Screen.Bounds
                Dim MyBmp As New Bitmap(MyRectangle.Width, MyRectangle.Height)
                Using MyGraphics As Graphics = Graphics.FromImage(MyBmp)
                    MyGraphics.CopyFromScreen(0, 0, 0, 0, MyRectangle.Size)
                End Using
                MyBmp.Save(ScreenFilePath, System.Drawing.Imaging.ImageFormat.Jpeg)
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function
        ''' <summary>
        ''' 获取屏幕截图（gif格式，实测1920*1080分辨率的截图文件大小为232K，文件大小较小）
        ''' </summary>
        ''' <param name="ScreenFilePath">截图文件保存路径</param>
        ''' <returns>是否执行成功</returns>
        ''' <remarks></remarks>
        Public Function SaveScreenGif(ByVal ScreenFilePath As String) As Boolean
            Try
                Dim MyRectangle As Rectangle = My.Computer.Screen.Bounds
                Dim MyBmp As New Bitmap(MyRectangle.Width, MyRectangle.Height)
                Using MyGraphics As Graphics = Graphics.FromImage(MyBmp)
                    MyGraphics.CopyFromScreen(0, 0, 0, 0, MyRectangle.Size)
                End Using
                MyBmp.Save(ScreenFilePath, System.Drawing.Imaging.ImageFormat.Gif)
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function
        ''' <summary>
        ''' 获取屏幕截图（ico格式，实测1920*1080分辨率的截图文件大小为294K，文件大小中等）
        ''' </summary>
        ''' <param name="ScreenFilePath">截图文件保存路径</param>
        ''' <returns>是否执行成功</returns>
        ''' <remarks></remarks>
        Public Function SaveScreenIcon(ByVal ScreenFilePath As String) As Boolean
            Try
                Dim MyRectangle As Rectangle = My.Computer.Screen.Bounds
                Dim MyBmp As New Bitmap(MyRectangle.Width, MyRectangle.Height)
                Using MyGraphics As Graphics = Graphics.FromImage(MyBmp)
                    MyGraphics.CopyFromScreen(0, 0, 0, 0, MyRectangle.Size)
                End Using
                MyBmp.Save(ScreenFilePath, System.Drawing.Imaging.ImageFormat.Icon)
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function
        ''' <summary>
        ''' 获取屏幕截图（tiff格式，实测1920*1080分辨率的截图文件大小为388K，文件大小较大）
        ''' </summary>
        ''' <param name="ScreenFilePath">截图文件保存路径</param>
        ''' <returns>是否执行成功</returns>
        ''' <remarks></remarks>
        Public Function SaveScreenTiff(ByVal ScreenFilePath As String) As Boolean
            Try
                Dim MyRectangle As Rectangle = My.Computer.Screen.Bounds
                Dim MyBmp As New Bitmap(MyRectangle.Width, MyRectangle.Height)
                Using MyGraphics As Graphics = Graphics.FromImage(MyBmp)
                    MyGraphics.CopyFromScreen(0, 0, 0, 0, MyRectangle.Size)
                End Using
                MyBmp.Save(ScreenFilePath, System.Drawing.Imaging.ImageFormat.Tiff)
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function
        ''' <summary>
        ''' 获取屏幕截图（exif格式，实测1920*1080分辨率的截图文件大小为294K，文件大小中等）
        ''' </summary>
        ''' <param name="ScreenFilePath">截图文件保存路径</param>
        ''' <returns>是否执行成功</returns>
        ''' <remarks></remarks>
        Public Function SaveScreenExif(ByVal ScreenFilePath As String) As Boolean
            Try
                Dim MyRectangle As Rectangle = My.Computer.Screen.Bounds
                Dim MyBmp As New Bitmap(MyRectangle.Width, MyRectangle.Height)
                Using MyGraphics As Graphics = Graphics.FromImage(MyBmp)
                    MyGraphics.CopyFromScreen(0, 0, 0, 0, MyRectangle.Size)
                End Using
                MyBmp.Save(ScreenFilePath, System.Drawing.Imaging.ImageFormat.Exif)
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function
        ''' <summary>
        ''' 获取屏幕截图（memorybmp格式，实测1920*1080分辨率的截图文件大小为294K，文件大小中等）
        ''' </summary>
        ''' <param name="ScreenFilePath">截图文件保存路径</param>
        ''' <returns>是否执行成功</returns>
        ''' <remarks></remarks>
        Public Function SaveScreenMemoryBmp(ByVal ScreenFilePath As String) As Boolean
            Try
                Dim MyRectangle As Rectangle = My.Computer.Screen.Bounds
                Dim MyBmp As New Bitmap(MyRectangle.Width, MyRectangle.Height)
                Using MyGraphics As Graphics = Graphics.FromImage(MyBmp)
                    MyGraphics.CopyFromScreen(0, 0, 0, 0, MyRectangle.Size)
                End Using
                MyBmp.Save(ScreenFilePath, System.Drawing.Imaging.ImageFormat.MemoryBmp)
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function
        ''' <summary>
        ''' 获取屏幕截图（emf格式，实测1920*1080分辨率的截图文件大小为293K，文件大小中等）
        ''' </summary>
        ''' <param name="ScreenFilePath">截图文件保存路径</param>
        ''' <returns>是否执行成功</returns>
        ''' <remarks></remarks>
        Public Function SaveScreenEmf(ByVal ScreenFilePath As String) As Boolean
            Try
                Dim MyRectangle As Rectangle = My.Computer.Screen.Bounds
                Dim MyBmp As New Bitmap(MyRectangle.Width, MyRectangle.Height)
                Using MyGraphics As Graphics = Graphics.FromImage(MyBmp)
                    MyGraphics.CopyFromScreen(0, 0, 0, 0, MyRectangle.Size)
                End Using
                MyBmp.Save(ScreenFilePath, System.Drawing.Imaging.ImageFormat.Emf)
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function
        ''' <summary>
        ''' 获取屏幕截图（wmf格式，实测1920*1080分辨率的截图文件大小为293K，文件大小中等）
        ''' </summary>
        ''' <param name="ScreenFilePath">截图文件保存路径</param>
        ''' <returns>是否执行成功</returns>
        ''' <remarks></remarks>
        Public Function SaveScreenWmf(ByVal ScreenFilePath As String) As Boolean
            Try
                Dim MyRectangle As Rectangle = My.Computer.Screen.Bounds
                Dim MyBmp As New Bitmap(MyRectangle.Width, MyRectangle.Height)
                Using MyGraphics As Graphics = Graphics.FromImage(MyBmp)
                    MyGraphics.CopyFromScreen(0, 0, 0, 0, MyRectangle.Size)
                End Using
                MyBmp.Save(ScreenFilePath, System.Drawing.Imaging.ImageFormat.Wmf)
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function
        ''' <summary>
        ''' 获取屏幕缩略图（默认保存格式，实测1920*1080【长宽都只保留50%】分辨率，缩略图文件大小为157K，文件大小中等）
        ''' </summary>
        ''' <param name="ScreenFilePath">截图文件保存路径</param>
        ''' <returns>是否执行成功</returns>
        ''' <remarks></remarks>
        Public Function SaveScreenThumbnail(ByVal ScreenFilePath As String) As Boolean
            Try
                Dim MyRectangle As Rectangle = My.Computer.Screen.Bounds
                Dim MyBmp As New Bitmap(MyRectangle.Width, MyRectangle.Height)
                Using MyGraphics As Graphics = Graphics.FromImage(MyBmp)
                    MyGraphics.CopyFromScreen(0, 0, 0, 0, MyRectangle.Size)
                End Using
                MyBmp = MyBmp.GetThumbnailImage(MyRectangle.Width / 2, MyRectangle.Height / 2, Nothing, New System.IntPtr(0))
                MyBmp.Save(ScreenFilePath)
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function
        ''' <summary>
        ''' 获取屏幕缩略图（png格式，实测1920*1080【长宽都只保留50%】分辨率，缩略图文件大小为157K，文件大小中等）
        ''' </summary>
        ''' <param name="ScreenFilePath">截图文件保存路径</param>
        ''' <returns>是否执行成功</returns>
        ''' <remarks></remarks>
        Public Function SaveScreenPngThumbnail(ByVal ScreenFilePath As String) As Boolean
            Try
                Dim MyRectangle As Rectangle = My.Computer.Screen.Bounds
                Dim MyBmp As New Bitmap(MyRectangle.Width, MyRectangle.Height)
                Using MyGraphics As Graphics = Graphics.FromImage(MyBmp)
                    MyGraphics.CopyFromScreen(0, 0, 0, 0, MyRectangle.Size)
                End Using
                MyBmp = MyBmp.GetThumbnailImage(MyRectangle.Width / 2, MyRectangle.Height / 2, Nothing, New System.IntPtr(0))
                MyBmp.Save(ScreenFilePath, System.Drawing.Imaging.ImageFormat.Png)
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function
        ''' <summary>
        ''' 获取屏幕缩略图（jpeg格式，实测1920*1080【长宽都只保留50%】分辨率，缩略图文件大小为64K，文件大小最小）
        ''' </summary>
        ''' <param name="ScreenFilePath">截图文件保存路径</param>
        ''' <returns>是否执行成功</returns>
        ''' <remarks></remarks>
        Public Function SaveScreenJpegThumbnail(ByVal ScreenFilePath As String) As Boolean
            Try
                Dim MyRectangle As Rectangle = My.Computer.Screen.Bounds
                Dim MyBmp As New Bitmap(MyRectangle.Width, MyRectangle.Height)
                Using MyGraphics As Graphics = Graphics.FromImage(MyBmp)
                    MyGraphics.CopyFromScreen(0, 0, 0, 0, MyRectangle.Size)
                End Using
                MyBmp = MyBmp.GetThumbnailImage(MyRectangle.Width / 2, MyRectangle.Height / 2, Nothing, New System.IntPtr(0))
                MyBmp.Save(ScreenFilePath, System.Drawing.Imaging.ImageFormat.Jpeg)
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function
        ''' <summary>
        ''' 发送按键（使用Microsoft.VisualBasic.Devices.Computer类的Keyboard.SendKeys方式）
        ''' </summary>
        ''' <param name="Keys">按键字符串（参照SendKeys的规则）</param>
        ''' <returns>是否执行成功</returns>
        ''' <remarks></remarks>
        Public Function SendKeys(ByVal Keys As String) As Boolean
            Try
                Dim A As New Microsoft.VisualBasic.Devices.Computer()
                A.Keyboard.SendKeys(Keys)
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function
        ''' <summary>
        ''' 发送按键（使用Wscript.Shell的SendKeys方式）
        ''' </summary>
        ''' <param name="Keys">按键字符串（参照SendKeys的规则）</param>
        ''' <returns>是否执行成功</returns>
        ''' <remarks></remarks>
        Public Function SendKeysWshShell(ByVal Keys As String) As Boolean
            Try
                CreateObject("Wscript.Shell").SendKeys(Keys)
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function
        ''' <summary>
        ''' 发送按键（使用System.Windows.Forms.SendKeys.Send方式）
        ''' </summary>
        ''' <param name="Keys">按键字符串（参照SendKeys的规则）</param>
        ''' <returns>是否执行成功</returns>
        ''' <remarks></remarks>
        Public Function SendKeysWinForm(ByVal Keys As String) As Boolean
            Try
                System.Windows.Forms.SendKeys.Send(Keys)
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function
        Private Declare Sub keybd_event Lib "user32.dll" Alias "keybd_event" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
        Private Declare Function MapVirtualKey Lib "user32.dll" Alias "MapVirtualKeyA" (ByVal wCode As Long, ByVal wMapType As Long) As Long
        ''' <summary>
        ''' 按下单个按键（并保持按下状态直到下次按同一个键，连续调用本函数，可执行组合键）
        ''' </summary>
        ''' <param name="Key">键位</param>
        ''' <returns>是否执行成功</returns>
        ''' <remarks></remarks>
        Public Function PressKey(ByVal Key As Keys) As Boolean
            Try
                keybd_event(Key, MapVirtualKey(Key, 0), 0, 0)
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function
        ''' <summary>
        ''' 按下单个按键（并保持按下状态直到下次按同一个键，连续调用本函数，可执行组合键）
        ''' </summary>
        ''' <param name="KeyChar">键位</param>
        ''' <returns>是否执行成功</returns>
        ''' <remarks></remarks>
        Public Function PressKey(ByVal KeyChar As Char) As Boolean
            Try
                keybd_event(Asc(KeyChar), MapVirtualKey(Asc(KeyChar), 0), 0, 0)
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function
        ''' <summary>
        ''' 按下多个按键（并保持按下状态直到下次按同一个键，连续调用本函数，可执行组合键）
        ''' </summary>
        ''' <param name="KeyString">键位（只允许字母和数字）</param>
        ''' <returns>是否执行成功</returns>
        ''' <remarks></remarks>
        Public Function PressKey(ByVal KeyString As String) As Boolean
            Try
                Dim Temp As Char() = KeyString.ToCharArray
                For Each TempChar As Char In Temp
                    keybd_event(Asc(TempChar), MapVirtualKey(Asc(TempChar), 0), 0, 0)
                Next
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function
        ''' <summary>
        ''' 释放单个按键（取消按下状态）
        ''' </summary>
        ''' <param name="Key">键位</param>
        ''' <returns>是否执行成功</returns>
        ''' <remarks></remarks>
        Public Function ReleaseKey(ByVal Key As Keys) As Boolean
            Try
                keybd_event(Key, MapVirtualKey(Key, 0), 2, 0)
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function
        ''' <summary>
        ''' 释放单个按键（取消按下状态）
        ''' </summary>
        ''' <param name="KeyChar">键位</param>
        ''' <returns>是否执行成功</returns>
        ''' <remarks></remarks>
        Public Function ReleaseKey(ByVal KeyChar As Char) As Boolean
            Try
                keybd_event(Asc(KeyChar), MapVirtualKey(Asc(KeyChar), 0), 2, 0)
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function
        ''' <summary>
        ''' 释放多个按键（取消按下状态）
        ''' </summary>
        ''' <param name="KeyString">键位（只允许字母和数字）</param>
        ''' <returns>是否执行成功</returns>
        ''' <remarks></remarks>
        Public Function ReleaseKey(ByVal KeyString As String) As Boolean
            Try
                Dim Temp As Char() = KeyString.ToCharArray
                For Each TempChar As Char In Temp
                    keybd_event(Asc(TempChar), MapVirtualKey(Asc(TempChar), 0), 2, 0)
                Next
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function
        Private Declare Function mouse_event Lib "user32.dll" Alias "mouse_event" (ByVal dwFlags As MouseEvent, ByVal dX As Int32, ByVal dY As Int32, ByVal dwData As Int32, ByVal dwExtraInfo As Int32) As Boolean
        <Flags()> _
        Private Enum MouseEvent
            Move = &H1
            AbsoluteLocation = &H8000
            LeftButtonDown = &H2
            LeftButtonUp = &H4
            MiddleButtonDown = &H20
            MiddleButtonUp = &H40
            RightButtonDown = &H8
            RightButtonUp = &H10
            Wheel = &H800
            AbsoluteScale = 65535
        End Enum
        ''' <summary>
        ''' 将鼠标位置移动一段距离（移动距离单位为像素，使用Win32 API）
        ''' </summary>
        ''' <param name="x">横向距离</param>
        ''' <param name="y">纵向距离</param>
        ''' <returns>是否执行成功</returns>
        ''' <remarks></remarks>
        Public Function MouseMoveByPixel(Optional ByVal x As Integer = 0, Optional ByVal y As Integer = 0) As Boolean
            Try
                mouse_event(MouseEvent.Move, x, y, 0, 0)
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function
        ''' <summary>
        ''' 将鼠标位置移动一段距离（移动距离单位为屏幕百分比）
        ''' </summary>
        ''' <param name="x">横向距离</param>
        ''' <param name="y">纵向距离</param>
        ''' <returns>是否执行成功</returns>
        ''' <remarks></remarks>
        Public Function MouseMoveByPercent(Optional ByVal x As Double = 0, Optional ByVal y As Double = 0) As Boolean
            Try
                mouse_event(MouseEvent.Move, x * My.Computer.Screen.Bounds.Width, y * My.Computer.Screen.Bounds.Height, 0, 0)
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function
        ''' <summary>
        ''' 移动鼠标到指定位置（定位单位为像素）
        ''' </summary>
        ''' <param name="x">横坐标</param>
        ''' <param name="y">纵坐标</param>
        ''' <returns>是否执行成功</returns>
        ''' <remarks></remarks>
        Public Function MouseMoveToPixel(Optional ByVal x As Integer = 0, Optional ByVal y As Integer = 0) As Boolean
            Try
                mouse_event(MouseEvent.Move Or MouseEvent.AbsoluteLocation, x / My.Computer.Screen.Bounds.Width * MouseEvent.AbsoluteScale, y / My.Computer.Screen.Bounds.Height * MouseEvent.AbsoluteScale, 0, 0)
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function
        ''' <summary>
        ''' 移动鼠标到指定位置（定位单位为屏幕百分比）
        ''' </summary>
        ''' <param name="x">横坐标</param>
        ''' <param name="y">纵坐标</param>
        ''' <returns>是否执行成功</returns>
        ''' <remarks></remarks>
        Public Function MouseMoveToPercent(Optional ByVal x As Double = 0, Optional ByVal y As Double = 0) As Boolean
            Try
                mouse_event(MouseEvent.Move Or MouseEvent.AbsoluteLocation, x * MouseEvent.AbsoluteScale, y * MouseEvent.AbsoluteScale, 0, 0)
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function
        ''' <summary>
        ''' 按下鼠标左键（保持按下状态）
        ''' </summary>
        ''' <returns>是否执行成功</returns>
        ''' <remarks></remarks>
        Public Function MouseLeftDown() As Boolean
            Try
                mouse_event(MouseEvent.LeftButtonDown, 0, 0, 0, 0)
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function
        ''' <summary>
        ''' 释放鼠标左键（取消按下状态）
        ''' </summary>
        ''' <returns>是否执行成功</returns>
        ''' <remarks></remarks>
        Public Function MouseLeftUp() As Boolean
            Try
                mouse_event(MouseEvent.LeftButtonUp, 0, 0, 0, 0)
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function
        ''' <summary>
        ''' 鼠标左键单击
        ''' </summary>
        ''' <returns>是否执行成功</returns>
        ''' <remarks></remarks>
        Public Function MouseLeftClick() As Boolean
            Try
                mouse_event(MouseEvent.LeftButtonDown Or MouseEvent.LeftButtonUp, 0, 0, 0, 0)
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function
        ''' <summary>
        ''' 鼠标左键双击
        ''' </summary>
        ''' <returns>是否执行成功</returns>
        ''' <remarks></remarks>
        Public Function MouseLeftDoubleClick() As Boolean
            Try
                mouse_event(MouseEvent.LeftButtonDown Or MouseEvent.LeftButtonUp, 0, 0, 0, 0)
                mouse_event(MouseEvent.LeftButtonDown Or MouseEvent.LeftButtonUp, 0, 0, 0, 0)
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function
        ''' <summary>
        ''' 按下鼠标中键（保持按下状态）
        ''' </summary>
        ''' <returns>是否执行成功</returns>
        ''' <remarks></remarks>
        Public Function MouseMiddleDown() As Boolean
            Try
                mouse_event(MouseEvent.MiddleButtonDown, 0, 0, 0, 0)
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function
        ''' <summary>
        ''' 释放鼠标中键（取消按下状态）
        ''' </summary>
        ''' <returns>是否执行成功</returns>
        ''' <remarks></remarks>
        Public Function MouseMiddleUp() As Boolean
            Try
                mouse_event(MouseEvent.MiddleButtonUp, 0, 0, 0, 0)
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function
        ''' <summary>
        ''' 鼠标中键单击
        ''' </summary>
        ''' <returns>是否执行成功</returns>
        ''' <remarks></remarks>
        Public Function MouseMiddleClick() As Boolean
            Try
                mouse_event(MouseEvent.MiddleButtonDown Or MouseEvent.MiddleButtonUp, 0, 0, 0, 0)
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function
        ''' <summary>
        ''' 鼠标中键双击
        ''' </summary>
        ''' <returns>是否执行成功</returns>
        ''' <remarks></remarks>
        Public Function MouseMiddleDoubleClick() As Boolean
            Try
                mouse_event(MouseEvent.MiddleButtonDown Or MouseEvent.MiddleButtonUp, 0, 0, 0, 0)
                mouse_event(MouseEvent.MiddleButtonDown Or MouseEvent.MiddleButtonUp, 0, 0, 0, 0)
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function
        ''' <summary>
        ''' 按下鼠标右键（保持按下状态）
        ''' </summary>
        ''' <returns>是否执行成功</returns>
        ''' <remarks></remarks>
        Public Function MouseRightDown() As Boolean
            Try
                mouse_event(MouseEvent.RightButtonDown, 0, 0, 0, 0)
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function
        ''' <summary>
        ''' 释放鼠标右键（取消按下状态）
        ''' </summary>
        ''' <returns>是否执行成功</returns>
        ''' <remarks></remarks>
        Public Function MouseRightUp() As Boolean
            Try
                mouse_event(MouseEvent.RightButtonUp, 0, 0, 0, 0)
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function
        ''' <summary>
        ''' 鼠标右键单击
        ''' </summary>
        ''' <returns>是否执行成功</returns>
        ''' <remarks></remarks>
        Public Function MouseRightClick() As Boolean
            Try
                mouse_event(MouseEvent.RightButtonDown Or MouseEvent.RightButtonUp, 0, 0, 0, 0)
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function
        ''' <summary>
        ''' 鼠标右键双击
        ''' </summary>
        ''' <returns>是否执行成功</returns>
        ''' <remarks></remarks>
        Public Function MouseRightDoubleClick() As Boolean
            Try
                mouse_event(MouseEvent.RightButtonDown Or MouseEvent.RightButtonUp, 0, 0, 0, 0)
                mouse_event(MouseEvent.RightButtonDown Or MouseEvent.RightButtonUp, 0, 0, 0, 0)
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function
        ''' <summary>
        ''' 鼠标滚轮向上滚动（滚动距离单位为像素）
        ''' </summary>
        ''' <param name="ScrollValue">滚动距离</param>
        ''' <returns>是否执行成功</returns>
        ''' <remarks></remarks>
        Public Function MouseWheelUp(ByVal ScrollValue As Integer) As Boolean
            Try
                mouse_event(MouseEvent.Wheel, 0, 0, ScrollValue, 0)
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function
        ''' <summary>
        ''' 鼠标滚轮向下滚动（滚动距离单位为像素）
        ''' </summary>
        ''' <param name="ScrollValue">滚动距离</param>
        ''' <returns>是否执行成功</returns>
        ''' <remarks></remarks>
        Public Function MouseWheelDown(ByVal ScrollValue As Integer) As Boolean
            Try
                mouse_event(MouseEvent.Wheel, 0, 0, -ScrollValue, 0)
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function
        Private Declare Function GetAsyncKeyState Lib "user32.dll" Alias "GetAsyncKeyState" (ByVal vKey As Int32) As Int16
        ''' <summary>
        ''' 判断物理键盘设备上的单个键位是否正处于被按下的状态（侦测键盘的硬件中断）
        ''' </summary>
        ''' <param name="Key">键位</param>
        ''' <returns>是否按下</returns>
        ''' <remarks></remarks>
        Public Function KeyBeingPressed(ByVal Key As Keys) As Boolean
            Dim Temp As Integer = GetAsyncKeyState(Key)
            If Temp = -32768 Or Temp = -32767 Then Return True
            Return False
        End Function
        ''' <summary>
        ''' 判断物理键盘设备上的单个键位是否正处于被按下的状态（侦测键盘的硬件中断）
        ''' </summary>
        ''' <param name="KeyChar">键位</param>
        ''' <returns>是否按下</returns>
        ''' <remarks></remarks>
        Public Function KeyBeingPressed(ByVal KeyChar As Char) As Boolean
            Dim Temp As Integer = GetAsyncKeyState(Asc(KeyChar))
            If Temp = -32768 Or Temp = -32767 Then Return True
            Return False
        End Function
        ''' <summary>
        ''' 判断物理键盘设备上的多个键位是否都处于被按下的状态（侦测键盘的硬件中断）
        ''' </summary>
        ''' <param name="KeyString">键位（只允许字母和数字）</param>
        ''' <returns>是否按下</returns>
        ''' <remarks></remarks>
        Public Function KeyBeingPressed(ByVal KeyString As String) As Boolean
            Dim Temp As Char() = KeyString.ToCharArray
            For Each TempChar As Char In Temp
                Dim TempInteger As Integer = GetAsyncKeyState(Asc(TempChar))
                If TempInteger <> -32768 And TempInteger <> -32767 Then Return False
            Next
            Return True
        End Function
        Private Declare Function FindWindow Lib "user32.dll" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As IntPtr
        Private Declare Function GetWindowRect Lib "user32.dll" Alias "GetWindowRect" (ByVal hwnd As IntPtr, ByRef lpRect As RECT) As IntPtr
        Private Structure RECT
            Dim Left As Integer
            Dim Top As Integer
            Dim Right As Integer
            Dim Bottom As Integer
        End Structure
        ''' <summary>
        ''' 根据窗口标题获取窗口的大小和位置（当多个标题相同的窗体存在时，默认获取上一个活动的窗体）
        ''' </summary>
        ''' <param name="WindowTitle">窗口标题（字符串不能有任何差别）</param>
        ''' <returns>窗口的大小和位置（System.Drawing.Rectangle）</returns>
        ''' <remarks></remarks>
        Public Function FindWindow(ByVal WindowTitle As String) As System.Drawing.Rectangle
            Dim hWnd As IntPtr
            Dim Rect As RECT
            hWnd = FindWindow(vbNullString, WindowTitle)
            GetWindowRect(hWnd, Rect)
            Return New System.Drawing.Rectangle(Rect.Left, Rect.Top, Rect.Right - Rect.Left, Rect.Bottom - Rect.Top)
        End Function
        Private Declare Function GetForegroundWindow Lib "user32.dll" () As IntPtr
        ''' <summary>
        ''' 获取当前的焦点窗口的大小和位置
        ''' </summary>
        ''' <returns>窗口的大小和位置（System.Drawing.Rectangle）</returns>
        ''' <remarks></remarks>
        Public Function FindFocusWindow() As System.Drawing.Rectangle
            Dim Rect As RECT
            GetWindowRect(GetForegroundWindow(), Rect)
            Return New System.Drawing.Rectangle(Rect.Left, Rect.Top, Rect.Right - Rect.Left, Rect.Bottom - Rect.Top)
        End Function
        Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hwnd As IntPtr, ByVal wMsg As Integer, ByVal wParam As Integer, ByVal lParam As System.Text.StringBuilder) As Integer
        ''' <summary>
        ''' 获取当前的焦点窗口的窗口标题文字
        ''' </summary>
        ''' <returns>窗口标题文字（失败则返回空字符串""）</returns>
        ''' <remarks></remarks>
        Public Function FindFocusWindowTitle() As String
            Dim StringBuilder As New System.Text.StringBuilder(4048)
            Dim WM_GETTEXT As Integer = 13
            SendMessage(GetForegroundWindow(), WM_GETTEXT, 4048, StringBuilder)
            Return StringBuilder.ToString()
        End Function
        Private Declare Function SetForegroundWindow Lib "user32.dll" (ByVal hwnd As IntPtr) As Integer
        ''' <summary>
        ''' 根据窗口标题获取窗口，并将其设置为当前的焦点窗口（实测：窗体处于最小化状态时，不会弹出到最前，只会在任务栏出现白色闪烁效果；窗体处于普通状态时，不一定会弹出到最前，可能只在任务栏出现黄色闪烁效果）
        ''' </summary>
        ''' <param name="WindowTitle">窗口标题（字符串不能有任何差别）</param>
        ''' <returns>设置成功返回非0值，失败返回0</returns>
        ''' <remarks></remarks>
        Public Function SetForegroundWindow(ByVal WindowTitle As String) As Integer
            Return SetForegroundWindow(FindWindow(vbNullString, WindowTitle))
        End Function
        Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hwnd As IntPtr, ByVal wMsg As Integer, ByVal wParam As Integer, ByVal lParam As Integer) As Integer
        ''' <summary>
        ''' 根据窗口标题获取窗口，并允许窗口重绘
        ''' </summary>
        ''' <param name="WindowTitle">窗口标题（字符串不能有任何差别）</param>
        ''' <returns>设置成功返回非0值，失败返回0</returns>
        ''' <remarks></remarks>
        Public Function SetWindowCanRedraw(ByVal WindowTitle As String) As Integer
            Dim wMsg As Integer = 11
            Return SendMessage(FindWindow(vbNullString, WindowTitle), wMsg, 1, 0)
        End Function
        ''' <summary>
        ''' 根据窗口标题获取窗口，并禁止窗口重绘
        ''' </summary>
        ''' <param name="WindowTitle">窗口标题（字符串不能有任何差别）</param>
        ''' <returns>设置成功返回非0值，失败返回0</returns>
        ''' <remarks></remarks>
        Public Function SetWindowCanNotRedraw(ByVal WindowTitle As String) As Integer
            Dim wMsg As Integer = 11
            Return SendMessage(FindWindow(vbNullString, WindowTitle), wMsg, 0, 0)
        End Function
        Private Declare Function ShowWindow Lib "user32.dll" (ByVal hWnd As IntPtr, ByVal nCmdShow As UInt32) As Boolean
        ''' <summary>
        ''' 根据窗口标题获取窗口，并将其设置为普通样式（取消最大化、最小化效果）
        ''' </summary>
        ''' <param name="WindowTitle">窗口标题（字符串不能有任何差别）</param>
        ''' <returns>是否设置成功</returns>
        ''' <remarks></remarks>
        Public Function ShowWindowNormal(ByVal WindowTitle As String) As Boolean
            Dim nCmdShow As UInt32 = 1
            Return ShowWindow(FindWindow(vbNullString, WindowTitle), nCmdShow)
        End Function
        ''' <summary>
        ''' 根据窗口标题获取窗口，并将其设置为最小化样式
        ''' </summary>
        ''' <param name="WindowTitle">窗口标题（字符串不能有任何差别）</param>
        ''' <returns>是否设置成功</returns>
        ''' <remarks></remarks>
        Public Function ShowWindowMinimize(ByVal WindowTitle As String) As Boolean
            Dim nCmdShow As UInt32 = 2
            Return ShowWindow(FindWindow(vbNullString, WindowTitle), nCmdShow)
        End Function
        ''' <summary>
        ''' 根据窗口标题获取窗口，并将其设置为最大化样式
        ''' </summary>
        ''' <param name="WindowTitle">窗口标题（字符串不能有任何差别）</param>
        ''' <returns>是否设置成功</returns>
        ''' <remarks></remarks>
        Public Function ShowWindowMaximize(ByVal WindowTitle As String) As Boolean
            Dim nCmdShow As UInt32 = 3
            Return ShowWindow(FindWindow(vbNullString, WindowTitle), nCmdShow)
        End Function
        Private Declare Function MoveWindow Lib "user32.dll" Alias "MoveWindow" (ByVal hwnd As IntPtr, ByVal X As Integer, ByVal Y As Integer, ByVal nWidth As Integer, ByVal nHeight As Integer, ByVal bRepaint As Boolean) As Boolean
        ''' <summary>
        ''' 根据窗口标题修改窗口的位置和大小（当多个标题相同的窗体存在时，默认修改上一个活动的窗体；注意可能会把窗口移动到用户鼠标无法触及的位置）
        ''' </summary>
        ''' <param name="WindowTitle">窗口标题（字符串不能有任何差别）</param>
        ''' <param name="Left">屏幕左边距</param>
        ''' <param name="Top">屏幕上边距</param>
        ''' <param name="Width">宽度</param>
        ''' <param name="Height">高度</param>
        ''' <returns>是否修改成功</returns>
        ''' <remarks></remarks>
        Public Function MoveWindow(ByVal WindowTitle As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer) As Boolean
            Dim hWnd As IntPtr
            Dim Rect As RECT
            hWnd = FindWindow(vbNullString, WindowTitle)
            GetWindowRect(hWnd, Rect)
            If hWnd = 0 Then Return False
            Return MoveWindow(hWnd, Left, Top, Width, Height, True)
        End Function
        ''' <summary>
        ''' 根据窗口标题修改窗口的位置和大小（当多个标题相同的窗体存在时，默认修改上一个活动的窗体；注意可能会把窗口移动到用户鼠标无法触及的位置）
        ''' </summary>
        ''' <param name="WindowTitle">窗口标题（字符串不能有任何差别）</param>
        ''' <param name="Position">指定区域</param>
        ''' <returns></returns>
        ''' <remarks>是否修改成功</remarks>
        Public Function MoveWindow(ByVal WindowTitle As String, ByVal Position As System.Drawing.Rectangle) As Boolean
            Dim hWnd As IntPtr
            Dim Rect As RECT
            hWnd = FindWindow(vbNullString, WindowTitle)
            GetWindowRect(hWnd, Rect)
            If hWnd = 0 Then Return False
            Return MoveWindow(hWnd, Position.Left, Position.Top, Position.Width, Position.Height, True)
        End Function
        ''' <summary>
        ''' 根据窗口标题拖动窗口（当多个标题相同的窗体存在时，默认修改上一个活动的窗体；注意可能会把窗口移动到用户鼠标无法触及的位置）
        ''' </summary>
        ''' <param name="WindowTitle">窗口标题（字符串不能有任何差别）</param>
        ''' <param name="MoveX">向右拖动的距离（为负值则向左拖动）</param>
        ''' <param name="MoveY">向下拖动的距离（为负值则向上拖动）</param>
        ''' <returns>是否拖动成功</returns>
        ''' <remarks></remarks>
        Public Function DragWindow(ByVal WindowTitle As String, Optional ByVal MoveX As Integer = 0, Optional ByVal MoveY As Integer = 0) As Boolean
            Dim hWnd As IntPtr
            Dim Rect As RECT
            hWnd = FindWindow(vbNullString, WindowTitle)
            GetWindowRect(hWnd, Rect)
            If hWnd = 0 Then Return False
            Return MoveWindow(hWnd, Rect.Left + MoveX, Rect.Top + MoveY, Rect.Right - Rect.Left, Rect.Bottom - Rect.Top, True)
        End Function
        ''' <summary>
        ''' 根据窗口标题修改窗口的大小（当多个标题相同的窗体存在时，默认修改上一个活动的窗体；注意某些程序的窗口大小，实际上不能被修改的太小）
        ''' </summary>
        ''' <param name="WindowTitle">窗口标题（字符串不能有任何差别）</param>
        ''' <param name="Width">宽度</param>
        ''' <param name="Height">高度</param>
        ''' <returns>是否修改成功</returns>
        ''' <remarks></remarks>
        Public Function ResizeWindow(ByVal WindowTitle As String, ByVal Width As Integer, ByVal Height As Integer) As Boolean
            Dim hWnd As IntPtr
            Dim Rect As RECT
            hWnd = FindWindow(vbNullString, WindowTitle)
            GetWindowRect(hWnd, Rect)
            If hWnd = 0 Then Return False
            Return MoveWindow(hWnd, Rect.Left, Rect.Top, Width, Height, True)
        End Function
        ''' <summary>
        ''' 将鼠标位置移动一段距离（移动距离单位为像素，使用System.Windows.Forms.Cursor）
        ''' </summary>
        ''' <param name="x">横向距离</param>
        ''' <param name="y">纵向距离</param>
        ''' <returns>是否执行成功</returns>
        ''' <remarks></remarks>
        Public Function MouseMove(ByVal x As Integer, ByVal y As Integer) As Boolean
            Try
                Dim P As System.Drawing.Point = System.Windows.Forms.Cursor.Position
                P.Offset(x, y)
                System.Windows.Forms.Cursor.Position = P
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function
        ''' <summary>
        ''' 获取鼠标的当前位置（System.Windows.Forms.Cursor.Position）
        ''' </summary>
        ''' <returns>位置坐标（System.Drawing.Point）</returns>
        ''' <remarks></remarks>
        Public Function MousePosition() As System.Drawing.Point
            Return System.Windows.Forms.Cursor.Position
        End Function
        ''' <summary>
        ''' 获取鼠标的当前位置的屏幕颜色
        ''' </summary>
        ''' <returns>颜色值（System.Drawing.Color）</returns>
        ''' <remarks></remarks>
        Public Function MousePositionColor() As System.Drawing.Color
            Dim MyRectangle As Rectangle = My.Computer.Screen.Bounds
            Dim MyBmp As New Bitmap(MyRectangle.Width, MyRectangle.Height)
            Using MyGraphics As Graphics = Graphics.FromImage(MyBmp)
                MyGraphics.CopyFromScreen(0, 0, 0, 0, MyRectangle.Size)
            End Using
            Return MyBmp.GetPixel(System.Windows.Forms.Cursor.Position.X, System.Windows.Forms.Cursor.Position.Y)
        End Function
        ''' <summary>
        ''' 强制把所有运行的程序窗口最小化，显示桌面（效果类似Win7系统鼠标点击屏幕右下角）
        ''' </summary>
        ''' <returns>是否执行成功</returns>
        ''' <remarks></remarks>
        Public Function ShowDesktop() As Boolean
            Try
                CreateObject("Shell.Application").ToggleDesktop()
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function
        ''' <summary>
        ''' 获取上一个Win32 API调用产生的错误代码（实测：出现错误后，错误信息会一直保留，直到被下一个错误信息替换）
        ''' </summary>
        ''' <returns>错误代码（默认为0）</returns>
        ''' <remarks></remarks>
        Public Function Win32ErrorCode() As Integer
            Return Runtime.InteropServices.Marshal.GetLastWin32Error
        End Function
        ''' <summary>
        ''' 获取上一个Win32 API调用产生的错误说明（实测：出现错误后，错误信息会一直保留，直到被下一个错误信息替换）
        ''' </summary>
        ''' <returns>错误说明（默认为"操作成功完成。"）</returns>
        ''' <remarks></remarks>
        Public Function Win32ErrorMessage() As String
            Return New System.ComponentModel.Win32Exception(Runtime.InteropServices.Marshal.GetLastWin32Error).Message
        End Function
    End Class

End Namespace