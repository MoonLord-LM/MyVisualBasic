Namespace My
    
    ''' <summary>
    ''' 进程管理相关函数
    ''' </summary>
    ''' <remarks></remarks>
    Public NotInheritable Class Screen

        ''' <summary>
        ''' 获取屏幕截图（全屏区域的截图）
        ''' </summary>
        ''' <returns>结果图片（失败返回空图片）</returns>
        ''' <remarks></remarks>
        Public Shared Function Image() As Bitmap
            Try
                Dim ScreenArea As Rectangle = My.Computer.Screen.Bounds
                Dim Temp As New Bitmap(ScreenArea.Width, ScreenArea.Height)
                Graphics.FromImage(Temp).CopyFromScreen(0, 0, 0, 0, ScreenArea.Size)
                Return Temp
            Catch ex As Exception
                Return New Bitmap(0, 0)
            End Try
        End Function
        ''' <summary>
        ''' 获取屏幕截图（指定区域的截图）
        ''' </summary>
        ''' <param name="Area">指定区域</param>
        ''' <returns>结果图片（失败返回空图片）</returns>
        ''' <remarks></remarks>
        Public Shared Function Image(ByVal Area As Rectangle) As Bitmap
            Try
                Dim Temp As New Bitmap(Area.Width, Area.Height)
                Graphics.FromImage(Temp).CopyFromScreen(Area.Left, Area.Top, 0, 0, Area.Size)
                Return Temp
            Catch ex As Exception
                Return New Bitmap(0, 0)
            End Try
        End Function
        ''' <summary>
        ''' 获取屏幕截图（全屏区域缩略图）
        ''' </summary>
        ''' <param name="Scale">缩略比例（应大于0，且小于等于1）</param>
        ''' <returns>结果图片（失败返回空图片）</returns>
        ''' <remarks></remarks>
        Public Shared Function Thumbnail(ByVal Scale As Double) As Bitmap
            If Scale <= 0 Or Scale > 1 Then
                Return New Bitmap(0, 0)
            End If
            Try
                Dim ScreenArea As Rectangle = My.Computer.Screen.Bounds
                Dim Temp As New Bitmap(ScreenArea.Width, ScreenArea.Height)
                Graphics.FromImage(Temp).CopyFromScreen(0, 0, 0, 0, ScreenArea.Size)
                Return Temp.GetThumbnailImage(ScreenArea.Width * Scale, ScreenArea.Height * Scale, Nothing, New System.IntPtr(0))
            Catch ex As Exception
                Return New Bitmap(0, 0)
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

    End Class

End Namespace