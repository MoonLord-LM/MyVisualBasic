Namespace My

    ''' <summary>
    ''' 模拟鼠标操作相关函数
    ''' </summary>
    ''' <remarks></remarks>
    Partial Public NotInheritable Class Mouse

        Private Declare Sub mouse_event Lib "user32.dll" Alias "mouse_event" (ByVal dwFlags As Int32, ByVal dX As Int32, ByVal dY As Int32, ByVal dwData As Int32, ByVal dwExtraInfo As UInt32)
        <Flags()> _
        Private Enum MouseEvent As Int32
            Move = 1
            LeftButtonDown = 2
            LeftButtonUp = 4
            RightButtonDown = 8
            RightButtonUp = 16
            MiddleButtonDown = 32
            MiddleButtonUp = 64
            Wheel = 2048
            AbsoluteLocation = 32768
            AbsoluteScale = 65535
        End Enum

        ''' <summary>
        ''' 移动鼠标位置，位移一定的像素点距离
        ''' </summary>
        ''' <param name="x">横向位移（向右为正方向）</param>
        ''' <param name="y">纵向位移（向下为正方向）</param>
        ''' <returns>是否执行成功</returns>
        ''' <remarks></remarks>
        Public Shared Function MoveByPixel(ByVal x As Integer, ByVal y As Integer) As Boolean
            Try
                mouse_event(MouseEvent.Move, x, y, 0, 0)
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function
        ''' <summary>
        ''' 移动鼠标位置，位移一定的屏幕百分比
        ''' </summary>
        ''' <param name="x">横向位移（百分比，向右为正方向）</param>
        ''' <param name="y">纵向位移（百分比，向下为正方向）</param>
        ''' <returns>是否执行成功</returns>
        ''' <remarks></remarks>
        Public Shared Function MoveByPercent(ByVal x As Double, ByVal y As Double) As Boolean
            Try
                mouse_event(MouseEvent.Move, x * My.Computer.Screen.Bounds.Width, y * My.Computer.Screen.Bounds.Height, 0, 0)
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function
        ''' <summary>
        ''' 移动鼠标位置，到指定的像素点坐标
        ''' </summary>
        ''' <param name="x">横坐标（原点为屏幕左上角，向右为正方向）</param>
        ''' <param name="y">纵坐标（原点为屏幕左上角，向下为正方向）</param>
        ''' <returns>是否执行成功</returns>
        ''' <remarks></remarks>
        Public Shared Function MoveToPixel(ByVal x As Integer, ByVal y As Integer) As Boolean
            Try
                mouse_event(MouseEvent.Move Or MouseEvent.AbsoluteLocation, x / My.Computer.Screen.Bounds.Width * MouseEvent.AbsoluteScale, y / My.Computer.Screen.Bounds.Height * MouseEvent.AbsoluteScale, 0, 0)
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function
        ''' <summary>
        ''' 移动鼠标位置，到指定的屏幕百分比坐标
        ''' </summary>
        ''' <param name="x">横坐标（百分比，原点为屏幕左上角，向右为正方向）</param>
        ''' <param name="y">纵坐标（百分比，原点为屏幕左上角，向下为正方向）</param>
        ''' <returns>是否执行成功</returns>
        ''' <remarks></remarks>
        Public Shared Function MoveToPercent(ByVal x As Double, ByVal y As Double) As Boolean
            Try
                mouse_event(MouseEvent.Move Or MouseEvent.AbsoluteLocation, x * MouseEvent.AbsoluteScale, y * MouseEvent.AbsoluteScale, 0, 0)
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function



        ''' <summary>
        ''' 获取鼠标位置，像素点坐标值
        ''' </summary>
        ''' <returns>结果坐标值（原点为屏幕左上角，X向右为正方向，Y向下为正方向）</returns>
        ''' <remarks></remarks>
        Public Shared Function Position() As Point
            Return Windows.Forms.Control.MousePosition
        End Function
        ''' <summary>
        ''' 获取鼠标位置的显示颜色，像素点Color值
        ''' </summary>
        ''' <returns>结果颜色值（System.Drawing.Color）</returns>
        ''' <remarks></remarks>
        Public Shared Function PositionColor() As Color
            Dim ScreenArea As Rectangle = My.Computer.Screen.Bounds
            Dim Temp As New Bitmap(ScreenArea.Width, ScreenArea.Height)
            Dim Graphics As Graphics = Graphics.FromImage(Temp)
            Dim Result As Color
            Graphics.CopyFromScreen(0, 0, 0, 0, ScreenArea.Size)
            Graphics.Dispose()
            Result = Temp.GetPixel(Windows.Forms.Cursor.Position.X, Windows.Forms.Cursor.Position.Y)
            Temp.Dispose()
            Return Result
        End Function
        ''' <summary>
        ''' 移动鼠标位置，到指定的像素点坐标
        ''' </summary>
        ''' <param name="Position">位置坐标（原点为屏幕左上角，X向右为正方向，Y向下为正方向）</param>
        ''' <returns>是否执行成功</returns>
        ''' <remarks></remarks>
        Public Shared Function MoveToPosition(ByVal Position As Point) As Boolean
            Try
                mouse_event(MouseEvent.Move Or MouseEvent.AbsoluteLocation, Position.X / My.Computer.Screen.Bounds.Width * MouseEvent.AbsoluteScale, Position.Y / My.Computer.Screen.Bounds.Height * MouseEvent.AbsoluteScale, 0, 0)
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
        Public Shared Function LeftDown() As Boolean
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
        Public Shared Function LeftUp() As Boolean
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
        Public Shared Function LeftClick() As Boolean
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
        Public Shared Function LeftDoubleClick() As Boolean
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
        Public Shared Function MiddleDown() As Boolean
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
        Public Shared Function MiddleUp() As Boolean
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
        Public Shared Function MiddleClick() As Boolean
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
        Public Shared Function MiddleDoubleClick() As Boolean
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
        Public Shared Function RightDown() As Boolean
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
        Public Shared Function RightUp() As Boolean
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
        Public Shared Function RightClick() As Boolean
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
        Public Shared Function RightDoubleClick() As Boolean
            Try
                mouse_event(MouseEvent.RightButtonDown Or MouseEvent.RightButtonUp, 0, 0, 0, 0)
                mouse_event(MouseEvent.RightButtonDown Or MouseEvent.RightButtonUp, 0, 0, 0, 0)
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function



        ''' <summary>
        ''' 鼠标滚轮向下滚动
        ''' </summary>
        ''' <param name="ScrollValue">滚动距离（单位为像素）</param>
        ''' <returns>是否执行成功</returns>
        ''' <remarks></remarks>
        Public Shared Function WheelDown(ByVal ScrollValue As Integer) As Boolean
            Try
                mouse_event(MouseEvent.Wheel, 0, 0, -ScrollValue, 0)
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function
        ''' <summary>
        ''' 鼠标滚轮向上滚动
        ''' </summary>
        ''' <param name="ScrollValue">滚动距离（单位为像素）</param>
        ''' <returns>是否执行成功</returns>
        ''' <remarks></remarks>
        Public Shared Function WheelUp(ByVal ScrollValue As Integer) As Boolean
            Try
                mouse_event(MouseEvent.Wheel, 0, 0, ScrollValue, 0)
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function

    End Class

End Namespace