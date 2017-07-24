Namespace My
    
    ''' <summary>
    ''' 屏幕截图相关函数
    ''' </summary>
    ''' <remarks></remarks>
    Partial Public NotInheritable Class Screen

        ''' <summary>
        ''' 获取屏幕截图（全屏区域的截图）
        ''' </summary>
        ''' <returns>结果图片（失败返回1*1个像素，#00000000透明色的图片）</returns>
        ''' <remarks></remarks>
        Public Shared Function Image() As Bitmap
            Try
                Dim ScreenArea As Rectangle = My.Computer.Screen.Bounds
                Dim Temp As New Bitmap(ScreenArea.Width, ScreenArea.Height)
                Dim Graphics As Graphics = Graphics.FromImage(Temp)
                Graphics.CopyFromScreen(0, 0, 0, 0, ScreenArea.Size)
                Graphics.Dispose()
                Return Temp
            Catch ex As Exception
                Return New Bitmap(1, 1)
            End Try
        End Function

        ''' <summary>
        ''' 获取屏幕截图（指定区域的截图）
        ''' </summary>
        ''' <param name="Area">指定区域</param>
        ''' <returns>结果图片（失败返回1*1个像素，#00000000透明色的图片）</returns>
        ''' <remarks></remarks>
        Public Shared Function Image(ByVal Area As Rectangle) As Bitmap
            Try
                Dim Temp As New Bitmap(Area.Width, Area.Height)
                Dim Graphics As Graphics = Graphics.FromImage(Temp)
                Graphics.CopyFromScreen(Area.Left, Area.Top, 0, 0, Area.Size)
                Graphics.Dispose()
                Return Temp
            Catch ex As Exception
                Return New Bitmap(1, 1)
            End Try
        End Function

        ''' <summary>
        ''' 获取屏幕截图（全屏区域的缩略图）
        ''' </summary>
        ''' <param name="Scale">缩略比例（应大于0，且小于等于1）</param>
        ''' <returns>结果图片（失败返回1*1个像素，#00000000透明色的图片）</returns>
        ''' <remarks></remarks>
        Public Shared Function ImageThumbnail(ByVal Scale As Double) As Bitmap
            If Scale <= 0 Or Scale > 1 Then
                Return New Bitmap(1, 1)
            End If
            Try
                Dim Thumbnail As Bitmap
                Dim ScreenArea As Rectangle = My.Computer.Screen.Bounds
                Dim Temp As New Bitmap(ScreenArea.Width, ScreenArea.Height)
                Dim Graphics As Graphics = Graphics.FromImage(Temp)
                Graphics.CopyFromScreen(0, 0, 0, 0, ScreenArea.Size)
                Graphics.Dispose()
                Thumbnail = Temp.GetThumbnailImage(ScreenArea.Width * Scale, ScreenArea.Height * Scale, Nothing, New System.IntPtr(0))
                Temp.Dispose()
                Return Thumbnail
            Catch ex As Exception
                Return New Bitmap(1, 1)
            End Try
        End Function

        ''' <summary>
        ''' 获取屏幕截图（指定区域的缩略图）
        ''' </summary>
        ''' <param name="Area">指定区域</param>
        ''' <param name="Scale">缩略比例（应大于0，且小于等于1）</param>
        ''' <returns>结果图片（失败返回1*1个像素，#00000000透明色的图片）</returns>
        ''' <remarks></remarks>
        Public Shared Function ImageThumbnail(ByVal Area As Rectangle, ByVal Scale As Double) As Bitmap
            If Scale <= 0 Or Scale > 1 Then
                Return New Bitmap(1, 1)
            End If
            Try
                Dim Thumbnail As Bitmap
                Dim Temp As New Bitmap(Area.Width, Area.Height)
                Dim Graphics As Graphics = Graphics.FromImage(Temp)
                Graphics.CopyFromScreen(Area.Left, Area.Top, 0, 0, Area.Size)
                Graphics.Dispose()
                Thumbnail = Temp.GetThumbnailImage(Area.Width * Scale, Area.Height * Scale, Nothing, New System.IntPtr(0))
                Temp.Dispose()
                Return Thumbnail
            Catch ex As Exception
                Return New Bitmap(1, 1)
            End Try
        End Function



        ''' <summary>
        ''' 获取屏幕区域
        ''' </summary>
        ''' <returns>结果区域值（System.Drawing.Rectangle）</returns>
        ''' <remarks></remarks>
        Public Shared Function Area() As System.Drawing.Rectangle
            Return My.Computer.Screen.Bounds
        End Function
        ''' <summary>
        ''' 获取屏幕区域大小
        ''' </summary>
        ''' <returns>结果大小值（System.Drawing.Size）</returns>
        ''' <remarks></remarks>
        Public Shared Function Size() As Size
            Return My.Computer.Screen.Bounds.Size
        End Function
        ''' <summary>
        ''' 获取屏幕中心点坐标
        ''' </summary>
        ''' <returns>结果坐标值（System.Drawing.Point）</returns>
        ''' <remarks></remarks>
        Public Shared Function CenterPoint() As Point
            Return New Point(My.Computer.Screen.Bounds.Left / 2, My.Computer.Screen.Bounds.Top / 2)
        End Function

        ''' <summary>
        ''' 获取屏幕工作区域
        ''' </summary>
        ''' <returns>结果区域值（System.Drawing.Rectangle）</returns>
        ''' <remarks></remarks>
        Public Shared Function WorkingArea() As System.Drawing.Rectangle
            Return My.Computer.Screen.WorkingArea
        End Function
        ''' <summary>
        ''' 获取屏幕工作区域大小
        ''' </summary>
        ''' <returns>结果大小值（System.Drawing.Size）</returns>
        ''' <remarks></remarks>
        Public Shared Function WorkingAreaSize() As Size
            Return My.Computer.Screen.WorkingArea.Size
        End Function
        ''' <summary>
        ''' 获取屏幕工作区域中心点坐标
        ''' </summary>
        ''' <returns>结果坐标值（System.Drawing.Point）</returns>
        ''' <remarks></remarks>
        Public Shared Function WorkingAreaCenterPoint() As Point
            Return New Point(My.Computer.Screen.WorkingArea.Left / 2, My.Computer.Screen.WorkingArea.Top / 2)
        End Function

    End Class

End Namespace