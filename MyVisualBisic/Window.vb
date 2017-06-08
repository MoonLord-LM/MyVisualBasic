Namespace My

    ''' <summary>
    ''' 窗口管理、控制相关函数
    ''' </summary>
    ''' <remarks></remarks>
    Partial Public NotInheritable Class Window

        Private Declare Function FindWindow Lib "user32.dll" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As IntPtr
        Private Declare Function GetWindowRect Lib "user32.dll" Alias "GetWindowRect" (ByVal hwnd As IntPtr, ByRef lpRect As Rect) As IntPtr
        Private Structure Rect
            Dim Left As Int32
            Dim Top As Int32
            Dim Right As Int32
            Dim Bottom As Int32
        End Structure

        ''' <summary>
        ''' 根据窗口标题，获取窗口的大小和位置（当多个标题相同的窗体存在时，默认获取上一个活动的窗体）
        ''' </summary>
        ''' <param name="WindowTitle">窗口标题（字符串必须完全相同）</param>
        ''' <returns>结果区域值（System.Drawing.Rectangle）</returns>
        ''' <remarks></remarks>
        Public Shared Function FindByTitle(ByVal WindowTitle As String) As System.Drawing.Rectangle
            Dim hWnd As IntPtr
            Dim Rect As Rect
            hWnd = FindWindow(Nothing, WindowTitle)
            GetWindowRect(hWnd, Rect)
            Return New System.Drawing.Rectangle(Rect.Left, Rect.Top, Rect.Right - Rect.Left, Rect.Bottom - Rect.Top)
        End Function

        Private Declare Function GetForegroundWindow Lib "user32.dll" () As IntPtr

        ''' <summary>
        ''' 获取当前的焦点窗口的大小和位置
        ''' </summary>
        ''' <returns>窗口的大小和位置（System.Drawing.Rectangle）</returns>
        ''' <remarks></remarks>
        Public Shared Function FindFocusWindow() As System.Drawing.Rectangle
            Dim Rect As Rect
            GetWindowRect(GetForegroundWindow(), Rect)
            Return New System.Drawing.Rectangle(Rect.Left, Rect.Top, Rect.Right - Rect.Left, Rect.Bottom - Rect.Top)
        End Function

        Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hwnd As IntPtr, ByVal wMsg As Integer, ByVal wParam As Integer, ByVal lParam As System.Text.StringBuilder) As Integer

        ''' <summary>
        ''' 获取当前的焦点窗口的窗口标题文字
        ''' </summary>
        ''' <returns>窗口标题文字（失败则返回空字符串""）</returns>
        ''' <remarks></remarks>
        Public Shared Function FindFocusWindowTitle() As String
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
        Public Shared Function SetForegroundWindow(ByVal WindowTitle As String) As Integer
            Return SetForegroundWindow(FindWindow(Nothing, WindowTitle))
        End Function
        Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hwnd As IntPtr, ByVal wMsg As Integer, ByVal wParam As Integer, ByVal lParam As Integer) As Integer
        ''' <summary>
        ''' 根据窗口标题获取窗口，并允许窗口重绘
        ''' </summary>
        ''' <param name="WindowTitle">窗口标题（字符串不能有任何差别）</param>
        ''' <returns>设置成功返回非0值，失败返回0</returns>
        ''' <remarks></remarks>
        Public Shared Function SetWindowCanRedraw(ByVal WindowTitle As String) As Integer
            Dim wMsg As Integer = 11
            Return SendMessage(FindWindow(Nothing, WindowTitle), wMsg, 1, 0)
        End Function
        ''' <summary>
        ''' 根据窗口标题获取窗口，并禁止窗口重绘
        ''' </summary>
        ''' <param name="WindowTitle">窗口标题（字符串不能有任何差别）</param>
        ''' <returns>设置成功返回非0值，失败返回0</returns>
        ''' <remarks></remarks>
        Public Shared Function SetWindowCanNotRedraw(ByVal WindowTitle As String) As Integer
            Dim wMsg As Integer = 11
            Return SendMessage(FindWindow(Nothing, WindowTitle), wMsg, 0, 0)
        End Function
        Private Declare Function ShowWindow Lib "user32.dll" (ByVal hWnd As IntPtr, ByVal nCmdShow As UInt32) As Boolean
        ''' <summary>
        ''' 根据窗口标题获取窗口，并将其设置为普通样式（取消最大化、最小化效果）
        ''' </summary>
        ''' <param name="WindowTitle">窗口标题（字符串不能有任何差别）</param>
        ''' <returns>是否设置成功</returns>
        ''' <remarks></remarks>
        Public Shared Function ShowWindowNormal(ByVal WindowTitle As String) As Boolean
            Dim nCmdShow As UInt32 = 1
            Return ShowWindow(FindWindow(Nothing, WindowTitle), nCmdShow)
        End Function
        ''' <summary>
        ''' 根据窗口标题获取窗口，并将其设置为最小化样式
        ''' </summary>
        ''' <param name="WindowTitle">窗口标题（字符串不能有任何差别）</param>
        ''' <returns>是否设置成功</returns>
        ''' <remarks></remarks>
        Public Shared Function ShowWindowMinimize(ByVal WindowTitle As String) As Boolean
            Dim nCmdShow As UInt32 = 2
            Return ShowWindow(FindWindow(Nothing, WindowTitle), nCmdShow)
        End Function
        ''' <summary>
        ''' 根据窗口标题获取窗口，并将其设置为最大化样式
        ''' </summary>
        ''' <param name="WindowTitle">窗口标题（字符串不能有任何差别）</param>
        ''' <returns>是否设置成功</returns>
        ''' <remarks></remarks>
        Public Shared Function ShowWindowMaximize(ByVal WindowTitle As String) As Boolean
            Dim nCmdShow As UInt32 = 3
            Return ShowWindow(FindWindow(Nothing, WindowTitle), nCmdShow)
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
        Public Shared Function MoveWindow(ByVal WindowTitle As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer) As Boolean
            Dim hWnd As IntPtr
            Dim Rect As Rect
            hWnd = FindWindow(Nothing, WindowTitle)
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
        Public Shared Function MoveWindow(ByVal WindowTitle As String, ByVal Position As System.Drawing.Rectangle) As Boolean
            Dim hWnd As IntPtr
            Dim Rect As Rect
            hWnd = FindWindow(Nothing, WindowTitle)
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
        Public Shared Function DragWindow(ByVal WindowTitle As String, Optional ByVal MoveX As Integer = 0, Optional ByVal MoveY As Integer = 0) As Boolean
            Dim hWnd As IntPtr
            Dim Rect As Rect
            hWnd = FindWindow(Nothing, WindowTitle)
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
        Public Shared Function ResizeWindow(ByVal WindowTitle As String, ByVal Width As Integer, ByVal Height As Integer) As Boolean
            Dim hWnd As IntPtr
            Dim Rect As Rect
            hWnd = FindWindow(Nothing, WindowTitle)
            GetWindowRect(hWnd, Rect)
            If hWnd = 0 Then Return False
            Return MoveWindow(hWnd, Rect.Left, Rect.Top, Width, Height, True)
        End Function

        ''' <summary>
        ''' 获取鼠标的当前位置的屏幕颜色
        ''' </summary>
        ''' <returns>颜色值（System.Drawing.Color）</returns>
        ''' <remarks></remarks>
        Public Shared Function MousePositionColor() As System.Drawing.Color
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
        Public Shared Function ShowDesktop() As Boolean
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
        Public Shared Function Win32ErrorCode() As Integer
            Return Runtime.InteropServices.Marshal.GetLastWin32Error
        End Function
        ''' <summary>
        ''' 获取上一个Win32 API调用产生的错误说明（实测：出现错误后，错误信息会一直保留，直到被下一个错误信息替换）
        ''' </summary>
        ''' <returns>错误说明（默认为"操作成功完成。"）</returns>
        ''' <remarks></remarks>
        Public Shared Function Win32ErrorMessage() As String
            Return New System.ComponentModel.Win32Exception(Runtime.InteropServices.Marshal.GetLastWin32Error).Message
        End Function
    End Class

End Namespace