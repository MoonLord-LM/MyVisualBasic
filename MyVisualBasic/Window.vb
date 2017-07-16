Namespace My

    ''' <summary>
    ''' 窗口管理、控制相关函数
    ''' </summary>
    ''' <remarks></remarks>
    Partial Public NotInheritable Class Window

        Private Declare Function FindWindow Lib "user32.dll" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As IntPtr
        Private Declare Function GetForegroundWindow Lib "user32.dll" Alias "GetForegroundWindow" () As IntPtr
        Private Declare Function GetParent Lib "user32.dll" Alias "GetParent" (ByVal hWnd As IntPtr) As IntPtr
        Private Declare Function WindowFromPoint Lib "user32.dll" Alias "WindowFromPoint" (ByVal Point As Point) As IntPtr

        ''' <summary>
        ''' 获取系统焦点窗口的窗口句柄
        ''' </summary>
        ''' <returns>结果窗口句柄（IntPtr）</returns>
        ''' <remarks></remarks>
        Public Shared Function FindFocus() As IntPtr
            Return GetForegroundWindow()
        End Function
        ''' <summary>
        ''' 获取父窗口的窗口句柄
        ''' </summary>
        ''' <param name="hWnd">窗口句柄（IntPtr）</param>
        ''' <returns>结果窗口句柄（IntPtr）</returns>
        ''' <remarks></remarks>
        Public Shared Function FindParent(ByVal hWnd As IntPtr) As IntPtr
            Return GetParent(hWnd)
        End Function
        ''' <summary>
        ''' 获取鼠标位置的窗口句柄
        ''' </summary>
        ''' <returns>结果窗口句柄（IntPtr）</returns>
        ''' <remarks></remarks>
        Public Shared Function FindByMouse() As IntPtr
            Return WindowFromPoint(Windows.Forms.Control.MousePosition)
        End Function
        ''' <summary>
        ''' 获取指定位置的窗口句柄
        ''' </summary>
        ''' <param name="Position">相对于屏幕的位置（Point）</param>
        ''' <returns>结果窗口句柄（IntPtr）</returns>
        ''' <remarks></remarks>
        Public Shared Function FindByPoint(ByVal Position As Point) As IntPtr
            Return WindowFromPoint(Position)
        End Function
        ''' <summary>
        ''' 根据窗口标题，获取窗口句柄（当有多个标题相同的窗体存在时，默认获取上一个活动的窗体）
        ''' </summary>
        ''' <param name="Title">窗口标题（字符串必须完全相同）</param>
        ''' <returns>结果窗口句柄（IntPtr）</returns>
        ''' <remarks></remarks>
        Public Shared Function FindByTitle(ByVal Title As String) As IntPtr
            Return FindWindow(Nothing, Title)
        End Function
        ''' <summary>
        ''' 根据窗口类名，获取窗口句柄（当有多个类名相同的窗体存在时，默认获取上一个活动的窗体）
        ''' </summary>
        ''' <param name="ClassName">窗口类名（字符串必须完全相同）</param>
        ''' <returns>结果窗口句柄（IntPtr）</returns>
        ''' <remarks></remarks>
        Public Shared Function FindByClassName(ByVal ClassName As String) As IntPtr
            Return FindWindow(ClassName, Nothing)
        End Function
        ''' <summary>
        ''' 根据进程名称，获取窗口句柄
        ''' </summary>
        ''' <param name="TaskName">进程名称（不含".exe"后缀）</param>
        ''' <returns>结果窗口句柄（IntPtr）</returns>
        ''' <remarks></remarks>
        Public Shared Function FindByTaskName(ByVal TaskName As String) As IntPtr
            Dim Processes As Process() = Process.GetProcesses()
            For Each P In Processes
                If P.ProcessName.ToLower() = TaskName.ToLower() Then
                    Return P.MainWindowHandle
                End If
            Next
            Return New IntPtr(0)
        End Function
        ''' <summary>
        ''' 根据进程文件路径，获取窗口句柄
        ''' </summary>
        ''' <param name="FilePath">文件路径（可以是相对路径）</param>
        ''' <returns>结果窗口句柄（IntPtr）</returns>
        ''' <remarks></remarks>
        Public Shared Function FindByFilePath(ByVal FilePath As String) As IntPtr
            Dim Processes As Process() = Process.GetProcesses()
            If FilePath.Contains(":") = False Then
                FilePath = System.IO.Directory.GetCurrentDirectory + "\" + FilePath
            End If
            For Each P In Processes
                Try
                    If P.MainModule.FileName.ToLower() = FilePath.ToLower() Then
                        Return P.MainWindowHandle
                    End If
                Catch ex As Exception
                    'MsgBox(P.ProcessName & ex.ToString())
                End Try
            Next
            Return New IntPtr(0)
        End Function
        ''' <summary>
        ''' 搜索窗口标题，获取窗口句柄（优先搜索字符串完全相同的，然后搜索包含有指定字符串的）
        ''' </summary>
        ''' <param name="Title">窗口标题（字符串不必完全相同）</param>
        ''' <returns>结果窗口句柄（IntPtr）</returns>
        ''' <remarks></remarks>
        Public Shared Function SearchByTitle(ByVal Title As String) As IntPtr
            Dim Processes As Process() = Process.GetProcesses()
            For Each P In Processes
                If P.MainWindowTitle.ToLower() = Title.ToLower() Then
                    Return P.MainWindowHandle
                End If
            Next
            For Each P In Processes
                If P.MainWindowTitle.ToLower().Contains(Title.ToLower()) Then
                    Return P.MainWindowHandle
                End If
            Next
            Return New IntPtr(0)
        End Function

        Private Declare Function EnumWindows Lib "user32.dll" Alias "EnumWindows" (ByVal lpEnumFunc As EnumWindowsProc, ByVal lParam As Object) As Boolean
        Private Declare Function EnumChildWindows Lib "user32.dll" Alias "EnumChildWindows" (ByVal hWndParent As IntPtr, ByVal lpEnumFunc As EnumWindowsProc, ByVal lParam As Object) As Boolean
        Private Delegate Function EnumWindowsProc(ByVal hWnd As IntPtr, ByVal lParam As Object) As Boolean
        Private Shared Function EnumWindows(ByVal hWnd As IntPtr, ByVal lParam As List(Of IntPtr)) As Boolean
            lParam.Add(hWnd)
            Return True
        End Function

        ''' <summary>
        ''' 获取所有屏幕上的顶层窗口的句柄列表
        ''' </summary>
        ''' <returns>结果IntPtr数组（失败返回空IntPtr数组）</returns>
        ''' <remarks></remarks>
        Public Shared Function List() As IntPtr()
            Dim TempList As List(Of IntPtr) = New List(Of IntPtr)()
            EnumWindows(AddressOf EnumWindows, TempList)
            Return TempList.ToArray()
        End Function
        ''' <summary>
        ''' 获取指定窗口的子窗口的句柄列表
        ''' </summary>
        ''' <param name="hWnd">窗口句柄（IntPtr）</param>
        ''' <returns>结果IntPtr数组（失败返回空IntPtr数组）</returns>
        ''' <remarks></remarks>
        Public Shared Function ListChildren(ByVal hWnd As IntPtr) As IntPtr()
            Dim TempList As List(Of IntPtr) = New List(Of IntPtr)()
            EnumChildWindows(hWnd, AddressOf EnumWindows, TempList)
            Return TempList.ToArray()
        End Function

        Private Declare Function IsWindow Lib "user32.dll" Alias "IsWindow" (ByVal hWnd As IntPtr) As Boolean
        Private Declare Function IsIconic Lib "user32.dll" Alias "IsIconic" (ByVal hWnd As IntPtr) As Boolean
        Private Declare Function IsZoomed Lib "user32.dll" Alias "IsZoomed" (ByVal hWnd As IntPtr) As Boolean
        Private Declare Function IsWindowVisible Lib "user32.dll" Alias "IsWindowVisible" (ByVal hWnd As IntPtr) As Boolean
        Private Declare Function IsWindowEnabled Lib "user32.dll" Alias "IsWindowEnabled" (ByVal hWnd As IntPtr) As Boolean

        ''' <summary>
        ''' 判断窗口是否获得了系统焦点
        ''' </summary>
        ''' <param name="hWnd">窗口句柄（IntPtr）</param>
        ''' <returns>是否获得焦点</returns>
        ''' <remarks></remarks>
        Public Shared Function CheckFocus(ByVal hWnd As IntPtr) As Boolean
            Return hWnd = GetForegroundWindow()
        End Function
        ''' <summary>
        ''' 判断窗口是否处于最小化状态
        ''' </summary>
        ''' <param name="hWnd">窗口句柄（IntPtr）</param>
        ''' <returns>是否最小化</returns>
        ''' <remarks></remarks>
        Public Shared Function CheckMinimized(ByVal hWnd As IntPtr) As Boolean
            Return IsIconic(hWnd)
        End Function
        ''' <summary>
        ''' 判断窗口是否处于最大化状态
        ''' </summary>
        ''' <param name="hWnd">窗口句柄（IntPtr）</param>
        ''' <returns>是否最大化</returns>
        ''' <remarks></remarks>
        Public Shared Function CheckMaximized(ByVal hWnd As IntPtr) As Boolean
            Return IsZoomed(hWnd)
        End Function
        ''' <summary>
        ''' 判断窗口是否处于可见状态
        ''' </summary>
        ''' <param name="hWnd">窗口句柄（IntPtr）</param>
        ''' <returns>是否可见</returns>
        ''' <remarks></remarks>
        Public Shared Function CheckVisible(ByVal hWnd As IntPtr) As Boolean
            Return IsWindowVisible(hWnd)
        End Function
        ''' <summary>
        ''' 判断窗口是否处于允许接受键鼠输入的状态
        ''' </summary>
        ''' <param name="hWnd">窗口句柄（IntPtr）</param>
        ''' <returns>是否接受输入</returns>
        ''' <remarks></remarks>
        Public Shared Function CheckEnabled(ByVal hWnd As IntPtr) As Boolean
            Return IsWindowEnabled(hWnd)
        End Function

        Private Declare Function GetWindowRect Lib "user32.dll" Alias "GetWindowRect" (ByVal hWnd As IntPtr, ByRef lpRect As Rect) As Boolean
        Private Declare Function GetWindowPlacement Lib "user32.dll" Alias "GetWindowPlacement" (ByVal hWnd As IntPtr, ByRef lpwndpl As Placement) As Boolean
        Private Declare Function GetClassName Lib "user32.dll" Alias "GetClassNameA" (ByVal hWnd As IntPtr, ByVal lpClassName As System.Text.StringBuilder, ByVal nMaxCount As UInt32) As UInt32
        Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hWnd As IntPtr, ByVal wMsg As UInt32, ByVal wParam As Int32, ByVal lParam As System.Text.StringBuilder) As Int32
        Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hWnd As IntPtr, ByVal wMsg As UInt32, ByVal wParam As Int32, Optional ByVal lParam As Object = Nothing) As Int32
        Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hWnd As IntPtr, ByVal wMsg As UInt32, Optional ByVal wParam As Object = Nothing, Optional ByVal lParam As Object = Nothing) As Int32
        Private Structure Rect
            Dim Left As Int32
            Dim Top As Int32
            Dim Right As Int32
            Dim Bottom As Int32
        End Structure
        Private Structure Placement
            Dim length As UInt32
            Dim flags As UInt32
            Dim showCmd As UInt32
            Dim ptMinPosition As Point
            Dim ptMaxPosition As Point
            Dim rcNormalPosition As Rect
        End Structure
        <Flags()> _
        Private Enum WindowsMessage As UInt32
            Destory = 2
            Quit = 10
            SetRedraw = 11
            SetText = 12
            GetText = 13
            GetTextLength = 14
            Close = 16
            KeyDown = 256
            KeyUp = 257
            KeyChar = 258
            SystemKeyDown = 260
            SystemKeyUp = 261
            SystemKeyChar = 262
            SystemCommand = 274
        End Enum

        Private Declare Function MapVirtualKey Lib "user32.dll" Alias "MapVirtualKeyA" (ByVal wCode As UInt32, ByVal wMapType As UInt32) As UInt32
        Private Declare Function PostMessage Lib "user32.dll" Alias "PostMessageA" (ByVal hWnd As IntPtr, ByVal wMsg As UInt32, ByVal wParam As UInt32, ByVal lParam As UInt32) As Boolean
        Private Declare Sub keybd_event Lib "user32.dll" Alias "keybd_event" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Int32, ByVal dwExtraInfo As UInt32)
        <Flags()> _
        Private Enum KeyEvent As Int32
            Down = 0
            Up = 2
        End Enum

        ''' <summary>
        ''' 向窗口发送按键消息（按键可能不被正常处理）
        ''' </summary>
        ''' <param name="hWnd">窗口句柄（IntPtr）</param>
        ''' <param name="Key">键位（Windows.Forms.Keys）</param>
        ''' <returns>是否执行成功</returns>
        ''' <remarks></remarks>
        Public Shared Function SendKey(ByVal hWnd As IntPtr, ByVal Key As Keys) As Boolean
            Dim KeyCode As UInt32 = Key
            Dim ScanCode As UInt32 = MapVirtualKey(KeyCode, 0)
            Dim LParam As UInt32 = 1 Or ScanCode << &H10UI
            Dim ExtendedKeys As New List(Of Keys)(New Keys() {Keys.RShiftKey, Keys.RControlKey, Keys.RMenu})
            If ExtendedKeys.Contains(Key) Then
                LParam = LParam Or &H1000000UI
            End If
            Dim Result As Boolean = True
            Result = Result And PostMessage(hWnd, WindowsMessage.KeyDown, KeyCode, LParam)
            LParam = LParam Or &HC0000000UI
            Result = Result And PostMessage(hWnd, WindowsMessage.KeyUp, KeyCode, LParam)
            Return Result
        End Function
        ''' <summary>
        ''' 向窗口发送按键消息（Alt组合键，实测：可发送Alt+F4组合键）
        ''' </summary>
        ''' <param name="hWnd">窗口句柄（IntPtr）</param>
        ''' <param name="Key">键位（Windows.Forms.Keys）</param>
        ''' <returns>是否执行成功</returns>
        ''' <remarks></remarks>
        Public Shared Function SendAltKey(ByVal hWnd As IntPtr, ByVal Key As Keys) As Boolean
            Dim AltKeyCode As UInt32 = Keys.Menu
            Dim AltScanCode As UInt32 = MapVirtualKey(AltKeyCode, 0)
            Dim AltLParam As UInt32 = 1 Or AltScanCode << &H10UI
            Dim KeyCode As UInt32 = Key
            Dim ScanCode As UInt32 = MapVirtualKey(KeyCode, 0)
            Dim LParam As UInt32 = 1 Or ScanCode << &H10UI
            Dim ExtendedKeys As New List(Of Keys)(New Keys() {Keys.RShiftKey, Keys.RControlKey, Keys.RMenu})
            If ExtendedKeys.Contains(Key) Then
                LParam = LParam Or &H1000000UI
            End If
            Dim Result As Boolean = True
            LParam = LParam Or &H20000000UI
            Result = Result And PostMessage(hWnd, WindowsMessage.SystemKeyDown, AltKeyCode, AltLParam Or &H20000000UI)
            Result = Result And PostMessage(hWnd, WindowsMessage.SystemKeyDown, KeyCode, LParam)
            LParam = LParam Or &HC0000000UI
            Result = Result And PostMessage(hWnd, WindowsMessage.SystemKeyUp, KeyCode, LParam)
            Result = Result And PostMessage(hWnd, WindowsMessage.KeyUp, AltKeyCode, AltLParam Or &HC0000000UI)
            Return Result
        End Function
        ''' <summary>
        ''' 向窗口发送按键消息（Ctrl组合键，实测：可发送Ctrl+S组合键）
        ''' </summary>
        ''' <param name="hWnd">窗口句柄（IntPtr）</param>
        ''' <param name="Key">键位（Windows.Forms.Keys）</param>
        ''' <returns>是否执行成功</returns>
        ''' <remarks></remarks>
        Public Shared Function SendCtrlKey(ByVal hWnd As IntPtr, ByVal Key As Keys) As Boolean
            Dim KeyCode As UInt32 = Key
            Dim ScanCode As UInt32 = MapVirtualKey(KeyCode, 0)
            Dim LParam As UInt32 = 1 Or ScanCode << &H10UI
            Dim ExtendedKeys As New List(Of Keys)(New Keys() {Keys.RShiftKey, Keys.RControlKey, Keys.RMenu})
            If ExtendedKeys.Contains(Key) Then
                LParam = LParam Or &H1000000UI
            End If
            Dim Result As Boolean = True
            keybd_event(Keys.ControlKey, MapVirtualKey(Keys.ControlKey, 0), KeyEvent.Down, 0)
            Try
                System.Threading.Thread.Sleep(10)
            Catch ex As Exception
            End Try
            Result = Result And PostMessage(hWnd, WindowsMessage.KeyDown, KeyCode, LParam)
            LParam = LParam Or &HC0000000UI
            Result = Result And PostMessage(hWnd, WindowsMessage.SystemKeyUp, KeyCode, LParam)
            Try
                System.Threading.Thread.Sleep(10)
            Catch ex As Exception
            End Try
            keybd_event(Keys.ControlKey, MapVirtualKey(Keys.ControlKey, 0), KeyEvent.Up, 0)
            Return Result
        End Function
        ''' <summary>
        ''' 向窗口发送按键消息（Shift组合键，实测：可发送Shift+1组合键）
        ''' </summary>
        ''' <param name="hWnd">窗口句柄（IntPtr）</param>
        ''' <param name="Key">键位（Windows.Forms.Keys）</param>
        ''' <returns>是否执行成功</returns>
        ''' <remarks></remarks>
        Public Shared Function SendShiftKey(ByVal hWnd As IntPtr, ByVal Key As Keys) As Boolean
            Dim KeyCode As UInt32 = Key
            Dim ScanCode As UInt32 = MapVirtualKey(KeyCode, 0)
            Dim LParam As UInt32 = 1 Or ScanCode << &H10UI
            Dim ExtendedKeys As New List(Of Keys)(New Keys() {Keys.RShiftKey, Keys.RControlKey, Keys.RMenu})
            If ExtendedKeys.Contains(Key) Then
                LParam = LParam Or &H1000000UI
            End If
            Dim Result As Boolean = True
            keybd_event(Keys.ShiftKey, MapVirtualKey(Keys.ShiftKey, 0), KeyEvent.Down, 0)
            Try
                System.Threading.Thread.Sleep(10)
            Catch ex As Exception
            End Try
            Result = Result And PostMessage(hWnd, WindowsMessage.KeyDown, KeyCode, LParam)
            LParam = LParam Or &HC0000000UI
            Result = Result And PostMessage(hWnd, WindowsMessage.SystemKeyUp, KeyCode, LParam)
            Try
                System.Threading.Thread.Sleep(10)
            Catch ex As Exception
            End Try
            keybd_event(Keys.ShiftKey, MapVirtualKey(Keys.ShiftKey, 0), KeyEvent.Up, 0)
            Return Result
        End Function

        ''' <summary>
        ''' 获取窗口区域（大小和位置，即使窗口处于隐藏、最小化、最大化状态也能获取到）
        ''' </summary>
        ''' <param name="hWnd">窗口句柄（IntPtr）</param>
        ''' <returns>结果区域值（System.Drawing.Rectangle）</returns>
        ''' <remarks></remarks>
        Public Shared Function GetRectangle(ByVal hWnd As IntPtr) As System.Drawing.Rectangle
            Dim WindowRect As Rect
            If IsIconic(hWnd) Or IsZoomed(hWnd) Then
                Dim WindowPlacement As Placement
                GetWindowPlacement(hWnd, WindowPlacement)
                WindowRect = WindowPlacement.rcNormalPosition
            Else
                GetWindowRect(hWnd, WindowRect)
            End If
            Return New System.Drawing.Rectangle(WindowRect.Left, WindowRect.Top, WindowRect.Right - WindowRect.Left, WindowRect.Bottom - WindowRect.Top)
        End Function
        ''' <summary>
        ''' 获取窗口位置（即使窗口处于隐藏、最小化、最大化状态也能获取到）
        ''' </summary>
        ''' <param name="hWnd">窗口句柄（IntPtr）</param>
        ''' <returns>结果窗口位置值（System.Drawing.Point）</returns>
        ''' <remarks></remarks>
        Public Shared Function GetLocation(ByVal hWnd As IntPtr) As Point
            Dim WindowRect As Rect
            If IsIconic(hWnd) Or IsZoomed(hWnd) Then
                Dim WindowPlacement As Placement
                GetWindowPlacement(hWnd, WindowPlacement)
                WindowRect = WindowPlacement.rcNormalPosition
            Else
                GetWindowRect(hWnd, WindowRect)
            End If
            Return New System.Drawing.Point(WindowRect.Left, WindowRect.Top)
        End Function
        ''' <summary>
        ''' 获取窗口大小（即使窗口处于隐藏、最小化、最大化状态也能获取到）
        ''' </summary>
        ''' <param name="hWnd">窗口句柄（IntPtr）</param>
        ''' <returns>结果窗口大小值（System.Drawing.Size）</returns>
        ''' <remarks></remarks>
        Public Shared Function GetSize(ByVal hWnd As IntPtr) As Size
            Dim WindowRect As Rect
            If IsIconic(hWnd) Or IsZoomed(hWnd) Then
                Dim WindowPlacement As Placement
                GetWindowPlacement(hWnd, WindowPlacement)
                WindowRect = WindowPlacement.rcNormalPosition
            Else
                GetWindowRect(hWnd, WindowRect)
            End If
            Return New System.Drawing.Size(WindowRect.Right - WindowRect.Left, WindowRect.Bottom - WindowRect.Top)
        End Function
        ''' <summary>
        ''' 获取窗口标题
        ''' </summary>
        ''' <param name="hWnd">窗口句柄（IntPtr）</param>
        ''' <returns>结果字符串（失败返回空字符串）</returns>
        ''' <remarks></remarks>
        Public Shared Function GetTitle(ByVal hWnd As IntPtr) As String
            Dim Length As Int32 = SendMessage(hWnd, WindowsMessage.GetTextLength, 0) + 1
            Dim StringBuilder As New System.Text.StringBuilder(Length)
            SendMessage(hWnd, WindowsMessage.GetText, Length, StringBuilder)
            Return StringBuilder.ToString()
        End Function
        ''' <summary>
        ''' 获取窗口类名（结果字符串最大长度限制为1024）
        ''' </summary>
        ''' <param name="hWnd">窗口句柄（IntPtr）</param>
        ''' <returns>结果字符串（失败返回空字符串）</returns>
        ''' <remarks></remarks>
        Public Shared Function GetClassName(ByVal hWnd As IntPtr) As String
            Dim Length As Int32 = 1024 + 1
            Dim StringBuilder As New System.Text.StringBuilder(Length)
            GetClassName(hWnd, StringBuilder, Length)
            Return StringBuilder.ToString()
        End Function

        Private Declare Function SetWindowPos Lib "user32.dll" Alias "SetWindowPos" (ByVal hWnd As IntPtr, ByVal hWndInsertAfter As IntPtr, ByVal X As Int32, ByVal Y As Int32, ByVal cx As Int32, ByVal cy As Int32, ByVal uFlags As UInt32) As Boolean
        Private Declare Function SetForegroundWindow Lib "user32.dll" Alias "SetForegroundWindow" (ByVal hWnd As IntPtr) As Boolean
        <Flags()> _
        Private Enum WindowPos As Int32
            Top = 0
            Bottom = 1
            TopMost = -1
            NoTopMost = -2
        End Enum
        <Flags()> _
        Private Enum SetPos As UInt32
            NoSize = 1 '忽略 cx、cy, 保持大小
            NoMove = 2 '忽略 X、Y, 保持位置
            NoZOrder = 4 '忽略 hWndInsertAfter, 保持窗口排列Z顺序
            NoRedraw = 8 '不重绘
            NoActivate = 16 '不激活，不改变窗口排列Z顺序
            FrameChanged = 32 '给窗口发送WM_NCCALCSIZE消息，即使窗口尺寸没有改变也会发送该消息。如果未指定这个标志，只有在改变了窗口尺寸时才发送WM_NCCALCSIZE。
            ShowWindow = 64 '显示窗口
            HideWindow = 128 '隐藏窗口
            AsyncWindowPos = 16384 '异步请求，不阻塞调用线程
        End Enum

        ''' <summary>
        ''' 设置窗口区域（即使窗口处于隐藏、最小化、最大化状态也能设置）
        ''' </summary>
        ''' <param name="hWnd">窗口句柄（IntPtr）</param>
        ''' <param name="Rectangle">窗口区域（System.Drawing.Rectangle）</param>
        ''' <returns>是否执行成功</returns>
        ''' <remarks></remarks>
        Public Shared Function SetRectangle(ByVal hWnd As IntPtr, ByVal Rectangle As System.Drawing.Rectangle) As Boolean
            Dim Result As Boolean
            If IsIconic(hWnd) Then
                ShowWindow(hWnd, ShowState.Hide)
                ShowWindow(hWnd, ShowState.ShowNormal)
                Result = SetWindowPos(hWnd, Nothing, Rectangle.Left, Rectangle.Top, Rectangle.Width, Rectangle.Height, SetPos.NoActivate)
                ShowWindow(hWnd, ShowState.ShowMinimized)
            ElseIf IsZoomed(hWnd) Then
                ShowWindow(hWnd, ShowState.Hide)
                ShowWindow(hWnd, ShowState.ShowNormal)
                Result = SetWindowPos(hWnd, Nothing, Rectangle.Left, Rectangle.Top, Rectangle.Width, Rectangle.Height, SetPos.NoActivate)
                ShowWindow(hWnd, ShowState.ShowMaximized)
            Else
                Result = SetWindowPos(hWnd, Nothing, Rectangle.Left, Rectangle.Top, Rectangle.Width, Rectangle.Height, SetPos.NoActivate)
            End If
            Return Result
        End Function
        ''' <summary>
        ''' 设置窗口位置（即使窗口处于隐藏、最小化、最大化状态也能设置）
        ''' </summary>
        ''' <param name="hWnd">窗口句柄（IntPtr）</param>
        ''' <param name="Point">窗口相对于屏幕左上角的位置，Left和Top（System.Drawing.Point）</param>
        ''' <returns>是否执行成功</returns>
        ''' <remarks></remarks>
        Public Shared Function SetLocation(ByVal hWnd As IntPtr, ByVal Point As Point) As Boolean
            Dim Result As Boolean
            If IsIconic(hWnd) Then
                ShowWindow(hWnd, ShowState.Hide)
                ShowWindow(hWnd, ShowState.ShowNormal)
                Result = SetWindowPos(hWnd, Nothing, Point.X, Point.Y, 0, 0, SetPos.NoSize Or SetPos.NoActivate)
                ShowWindow(hWnd, ShowState.ShowMinimized)
            ElseIf IsZoomed(hWnd) Then
                ShowWindow(hWnd, ShowState.Hide)
                ShowWindow(hWnd, ShowState.ShowNormal)
                Result = SetWindowPos(hWnd, Nothing, Point.X, Point.Y, 0, 0, SetPos.NoSize Or SetPos.NoActivate)
                ShowWindow(hWnd, ShowState.ShowMaximized)
            Else
                Result = SetWindowPos(hWnd, Nothing, Point.X, Point.Y, 0, 0, SetPos.NoSize Or SetPos.NoActivate)
            End If
            Return Result
        End Function
        ''' <summary>
        ''' 设置窗口居中（即使窗口处于隐藏、最小化、最大化状态也能设置，效果类似于StartPosition = FormStartPosition.CenterScreen）
        ''' </summary>
        ''' <param name="hWnd">窗口句柄（IntPtr）</param>
        ''' <returns>是否执行成功</returns>
        ''' <remarks></remarks>
        Public Shared Function SetCenterScreen(ByVal hWnd As IntPtr) As Boolean
            Dim Screen As Rectangle = My.Computer.Screen.Bounds
            Dim WindowPlacement As Placement
            GetWindowPlacement(hWnd, WindowPlacement)
            Dim WindowRect As Rect = WindowPlacement.rcNormalPosition
            Dim Result As Boolean
            If IsIconic(hWnd) Then
                ShowWindow(hWnd, ShowState.Hide)
                ShowWindow(hWnd, ShowState.ShowNormal)
                Result = SetWindowPos(hWnd, Nothing, (Screen.Width - (WindowRect.Right - WindowRect.Left)) / 2, (Screen.Height - (WindowRect.Bottom - WindowRect.Top)) / 2, 0, 0, SetPos.NoSize Or SetPos.NoActivate)
                ShowWindow(hWnd, ShowState.ShowMinimized)
            ElseIf IsZoomed(hWnd) Then
                ShowWindow(hWnd, ShowState.Hide)
                ShowWindow(hWnd, ShowState.ShowNormal)
                Result = SetWindowPos(hWnd, Nothing, (Screen.Width - (WindowRect.Right - WindowRect.Left)) / 2, (Screen.Height - (WindowRect.Bottom - WindowRect.Top)) / 2, 0, 0, SetPos.NoSize Or SetPos.NoActivate)
                ShowWindow(hWnd, ShowState.ShowMaximized)
            Else
                Result = SetWindowPos(hWnd, Nothing, (Screen.Width - (WindowRect.Right - WindowRect.Left)) / 2, (Screen.Height - (WindowRect.Bottom - WindowRect.Top)) / 2, 0, 0, SetPos.NoSize Or SetPos.NoActivate)
            End If
            Return Result
        End Function
        ''' <summary>
        ''' 设置窗口大小（即使窗口处于隐藏、最小化、最大化状态也能设置）
        ''' </summary>
        ''' <param name="hWnd">窗口句柄（IntPtr）</param>
        ''' <param name="Size">窗口大小，Width和Height（System.Drawing.Size）</param>
        ''' <returns>是否执行成功</returns>
        ''' <remarks></remarks>
        Public Shared Function SetSize(ByVal hWnd As IntPtr, ByVal Size As Size) As Boolean
            Dim Result As Boolean
            If IsIconic(hWnd) Then
                ShowWindow(hWnd, ShowState.Hide)
                ShowWindow(hWnd, ShowState.ShowNormal)
                Result = SetWindowPos(hWnd, Nothing, 0, 0, Size.Width, Size.Height, SetPos.NoMove Or SetPos.NoActivate)
                ShowWindow(hWnd, ShowState.ShowMinimized)
            ElseIf IsZoomed(hWnd) Then
                ShowWindow(hWnd, ShowState.Hide)
                ShowWindow(hWnd, ShowState.ShowNormal)
                Result = SetWindowPos(hWnd, Nothing, 0, 0, Size.Width, Size.Height, SetPos.NoMove Or SetPos.NoActivate)
                ShowWindow(hWnd, ShowState.ShowMaximized)
            Else
                Result = SetWindowPos(hWnd, Nothing, 0, 0, Size.Width, Size.Height, SetPos.NoMove Or SetPos.NoActivate)
            End If
            Return Result
        End Function
        ''' <summary>
        ''' 显示窗口（不获得系统焦点）
        ''' </summary>
        ''' <param name="hWnd">窗口句柄（IntPtr）</param>
        ''' <returns>是否执行成功</returns>
        ''' <remarks></remarks>
        Public Shared Function Show(ByVal hWnd As IntPtr) As Boolean
            Return SetWindowPos(hWnd, Nothing, 0, 0, 0, 0, SetPos.ShowWindow Or SetPos.NoMove Or SetPos.NoSize Or SetPos.NoActivate)
        End Function
        ''' <summary>
        ''' 隐藏窗口（不获得系统焦点）
        ''' </summary>
        ''' <param name="hWnd">窗口句柄（IntPtr）</param>
        ''' <returns>是否执行成功</returns>
        ''' <remarks></remarks>
        Public Shared Function Hide(ByVal hWnd As IntPtr) As Boolean
            Return SetWindowPos(hWnd, Nothing, 0, 0, 0, 0, SetPos.HideWindow Or SetPos.NoMove Or SetPos.NoSize Or SetPos.NoActivate)
        End Function
        ''' <summary>
        ''' 设置窗口是否置顶（置顶窗口，显示在其它所有非置顶窗口之上）
        ''' </summary>
        ''' <param name="hWnd">窗口句柄（IntPtr）</param>
        ''' <param name="TopMost">是否置顶</param>
        ''' <returns>是否执行成功</returns>
        ''' <remarks></remarks>
        Public Shared Function SetTopMost(ByVal hWnd As IntPtr, Optional ByVal TopMost As Boolean = True) As Boolean
            If TopMost Then
                Return SetWindowPos(hWnd, WindowPos.TopMost, 0, 0, 0, 0, SetPos.NoMove Or SetPos.NoSize Or SetPos.NoActivate)
            Else
                Return SetWindowPos(hWnd, WindowPos.NoTopMost, 0, 0, 0, 0, SetPos.NoMove Or SetPos.NoSize Or SetPos.NoActivate)
            End If
        End Function
        ''' <summary>
        ''' 设置窗口标题
        ''' </summary>
        ''' <param name="hWnd">窗口句柄（IntPtr）</param>
        ''' <param name="Title">窗口标题</param>
        ''' <returns>是否执行成功</returns>
        ''' <remarks></remarks>
        Public Shared Function SetTitle(ByVal hWnd As IntPtr, ByVal Title As String) As Boolean
            Dim StringBuilder As New System.Text.StringBuilder(Title)
            Return SendMessage(hWnd, WindowsMessage.SetText, StringBuilder.Length + 1, StringBuilder)
        End Function

        Private Declare Function ShowWindow Lib "user32.dll" Alias "ShowWindow" (ByVal hWnd As IntPtr, ByVal nCmdShow As UInt32) As Boolean
        <Flags()> _
        Private Enum ShowState As UInt32
            Hide = 0
            ShowNormal = 1
            ShowMinimized = 2
            ShowMaximized = 3
            ShowNoActivate = 4
            Restore = 9
        End Enum

        ''' <summary>
        ''' 还原显示窗口（效果类似于WindowState = FormWindowState.Normal，隐藏的窗口会显示出来，不获得系统焦点）
        ''' </summary>
        ''' <param name="hWnd">窗口句柄（IntPtr）</param>
        ''' <returns>是否执行成功</returns>
        ''' <remarks></remarks>
        Public Shared Function ShowNormal(ByVal hWnd As IntPtr) As Boolean
            ShowWindow(hWnd, ShowState.Hide)
            Return ShowWindow(hWnd, ShowState.ShowNormal)
        End Function
        ''' <summary>
        ''' 最小化显示窗口（效果类似于WindowState = FormWindowState.Minimized，隐藏的窗口会显示出来，不获得系统焦点）
        ''' </summary>
        ''' <param name="hWnd">窗口句柄（IntPtr）</param>
        ''' <returns>是否执行成功</returns>
        ''' <remarks></remarks>
        Public Shared Function ShowMinimized(ByVal hWnd As IntPtr) As Boolean
            ShowWindow(hWnd, ShowState.Hide)
            Return ShowWindow(hWnd, ShowState.ShowMinimized)
        End Function
        ''' <summary>
        ''' 最大化显示窗口（效果类似于WindowState = FormWindowState.Maximized，隐藏的窗口会显示出来，不获得系统焦点）
        ''' </summary>
        ''' <param name="hWnd">窗口句柄（IntPtr）</param>
        ''' <returns>是否执行成功</returns>
        ''' <remarks></remarks>
        Public Shared Function ShowMaximized(ByVal hWnd As IntPtr) As Boolean
            ShowWindow(hWnd, ShowState.Hide)
            Return ShowWindow(hWnd, ShowState.ShowMaximized)
        End Function

        Private Declare Function GetWindowThreadProcessId Lib "user32.dll" Alias "GetWindowThreadProcessId" (ByVal hWnd As IntPtr, ByRef lpdwProcessId As Int32) As Int32
        Private Declare Function AttachThreadInput Lib "user32.dll" Alias "AttachThreadInput" (ByVal idAttach As Int32, ByVal idAttachTo As Int32, ByVal fAttach As Boolean) As Boolean
        Private Declare Function EnableWindow Lib "user32.dll" Alias "EnableWindow" (ByVal hWnd As IntPtr, ByVal bEnable As Boolean) As Boolean

        ''' <summary>
        ''' 使窗口获得系统焦点（隐藏的窗口会显示出来，最小化/最大化的窗口会还原，禁止重绘的窗口会允许重绘）
        ''' </summary>
        ''' <param name="hWnd">窗口句柄（IntPtr）</param>
        ''' <returns>是否执行成功</returns>
        ''' <remarks></remarks>
        Public Shared Function SetFocus(ByVal hWnd As IntPtr) As Boolean
            Dim ForegroundThreadId As Int32
            Dim HandleThreadId As Int32
            Dim Result As Int32
            ForegroundThreadId = GetWindowThreadProcessId(GetForegroundWindow(), Nothing)
            HandleThreadId = GetWindowThreadProcessId(hWnd, Nothing)
            AttachThreadInput(HandleThreadId, ForegroundThreadId, True)
            ShowWindow(hWnd, ShowState.ShowNormal)
            SetWindowPos(hWnd, WindowPos.TopMost, 0, 0, 0, 0, SetPos.NoMove Or SetPos.NoSize)
            SetWindowPos(hWnd, WindowPos.NoTopMost, 0, 0, 0, 0, SetPos.NoMove Or SetPos.NoSize)
            Result = SetForegroundWindow(hWnd)
            AttachThreadInput(HandleThreadId, ForegroundThreadId, False)
            Return Result
        End Function

        ''' <summary>
        ''' 设置窗口是否允许重绘（禁止重绘后，窗口画面静止不变，可以减轻负荷，但是无法接收用户输入，此时用鼠标点击窗口区域，会点击到后面的窗口）
        ''' </summary>
        ''' <param name="hWnd">窗口句柄（IntPtr）</param>
        ''' <param name="CanRedraw">是否允许重绘</param>
        ''' <returns>是否执行成功</returns>
        ''' <remarks></remarks>
        Public Shared Function SetCanRedraw(ByVal hWnd As IntPtr, Optional ByVal CanRedraw As Boolean = True) As Boolean
            If CanRedraw Then
                Return SendMessage(hWnd, WindowsMessage.SetRedraw, 1)
            Else
                Return SendMessage(hWnd, WindowsMessage.SetRedraw, 0)
            End If
        End Function
        ''' <summary>
        ''' 设置窗口是否允许接受键鼠输入（禁止接收用户输入时，用鼠标点击窗口区域，不会点击到后面的窗口，但会有鼠标无法正常点击的警告声）
        ''' </summary>
        ''' <param name="hWnd">窗口句柄（IntPtr）</param>
        ''' <param name="Enable">是否允许接受输入</param>
        ''' <returns>是否执行成功</returns>
        ''' <remarks></remarks>
        Public Shared Function SetEnable(ByVal hWnd As IntPtr, Optional ByVal Enable As Boolean = True) As Boolean
            Return EnableWindow(hWnd, Enable)
        End Function

        Private Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal hWnd As IntPtr, ByVal nIndex As Int32) As Int32
        Private Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" (ByVal hWnd As IntPtr, ByVal nIndex As Int32, ByVal dwNewLong As Int32) As Int32
        Private Declare Function SetLayeredWindowAttributes Lib "user32.dll" Alias "SetLayeredWindowAttributes" (ByVal hWnd As IntPtr, ByVal crKey As IntPtr, ByVal bAlpha As Byte, ByVal dwFlags As Int32) As Boolean
        Private Declare Function GetLayeredWindowAttributes Lib "user32.dll" Alias "GetLayeredWindowAttributes" (ByVal hWnd As IntPtr, ByVal crKey As IntPtr, ByRef bAlpha As Byte, ByVal dwFlags As Int32) As Boolean
        <Flags()> _
        Private Enum WindowStyleEx As Int32
            Layered = 524288
        End Enum
        <Flags()> _
        Private Enum WindowLong As Int32
            ExStyle = -20
        End Enum
        <Flags()> _
        Private Enum LayeredWindowAttribute As Int32
            ColorKey = 1
            Alpha = 2
        End Enum

        ''' <summary>
        ''' 设置窗体的不透明度级别
        ''' </summary>
        ''' <param name="hWnd">窗口句柄（IntPtr）</param>
        ''' <param name="Opacity">不透明程度（介于0到1之间，0为完全透明，1为完全不透明）</param>
        ''' <returns></returns>
        ''' <remarks>是否执行成功</remarks>
        Public Shared Function SetOpacity(ByVal hWnd As IntPtr, Optional ByVal Opacity As Double = 1) As Boolean
            If Opacity < 0 Then
                Opacity = 0
            End If
            If Opacity > 1 Then
                Opacity = 1
            End If
            Dim Temp As Int32 = GetWindowLong(hWnd, WindowLong.ExStyle)
            Temp = Temp Or WindowStyleEx.Layered
            SetWindowLong(hWnd, WindowLong.ExStyle, Temp)
            Return SetLayeredWindowAttributes(hWnd, 0, Opacity * 255, LayeredWindowAttribute.Alpha)
        End Function
        ''' <summary>
        ''' 获取窗体的不透明度级别
        ''' </summary>
        ''' <param name="hWnd">窗口句柄（IntPtr）</param>
        ''' <returns></returns>
        ''' <remarks>不透明程度（介于0到1之间，0为完全透明，1为完全不透明）</remarks>
        Public Shared Function GetOpacity(ByVal hWnd As IntPtr) As Double
            Dim Temp As Byte
            GetLayeredWindowAttributes(hWnd, 0, Temp, LayeredWindowAttribute.Alpha)
            Return Temp / 255
        End Function

        Private Declare Function RedrawWindow Lib "user32.dll" Alias "RedrawWindow" (ByVal hWnd As IntPtr, lprcUpdate As Rect, ByVal hrgnUpdate As IntPtr, ByVal fuRedraw As UInt32) As Boolean
        <Flags()> _
        Private Enum Redraw As UInt32
            Invalidate = 1
            InternalPaint = 2
            DoErase = 4
            Validate = 8
            NoInternalPaint = 16
            NoErase = 32
            NoChildren = 64
            AllChildren = 128
            UpdateNow = 256
            EraseNow = 512
            Frame = 1024
            NoFrame = 2048
        End Enum

        ''' <summary>
        ''' 刷新窗口（更新一帧画面，禁止重绘的窗口仍然禁止重绘）
        ''' </summary>
        ''' <param name="hWnd">窗口句柄（IntPtr）</param>
        ''' <returns>是否执行成功</returns>
        ''' <remarks></remarks>
        Public Shared Function Refresh(ByVal hWnd As IntPtr) As Boolean
            Return RedrawWindow(hWnd, Nothing, Nothing, Redraw.Invalidate Or Redraw.DoErase Or Redraw.UpdateNow)
        End Function

        Private Declare Function GetWindowDC Lib "user32.dll" Alias "GetWindowDC" (ByVal hWnd As IntPtr) As IntPtr
        Private Declare Function ReleaseDC Lib "user32.dll" Alias "ReleaseDC" (ByVal hWnd As IntPtr, ByVal hDC As IntPtr) As Boolean
        Private Declare Function CreateCompatibleDC Lib "gdi32.dll" Alias "CreateCompatibleDC" (ByVal hDC As IntPtr) As IntPtr
        Private Declare Function DeleteDC Lib "gdi32.dll" Alias "DeleteDC" (ByVal hDC As IntPtr) As Boolean
        Private Declare Function CreateCompatibleBitmap Lib "gdi32.dll" Alias "CreateCompatibleBitmap" (ByVal hDC As IntPtr, ByVal nWidth As Int32, ByVal nHeight As Int32) As IntPtr
        Private Declare Function DeleteObject Lib "gdi32.dll" Alias "DeleteObject" (ByVal hObject As IntPtr) As Boolean
        Private Declare Function SelectObject Lib "gdi32.dll" Alias "SelectObject" (ByVal hDC As IntPtr, ByVal hgdiobj As IntPtr) As IntPtr
        Private Declare Function PrintWindow Lib "user32.dll" Alias "PrintWindow" (ByVal hWnd As IntPtr, ByVal hDCBlt As IntPtr, ByVal nFlags As UInt32) As Boolean

        ''' <summary>
        ''' 获取窗口截图（即使窗口处于屏幕外、被遮挡、最小化等状态，也能获取到）
        ''' </summary>
        ''' <param name="hWnd">窗口句柄（IntPtr）</param>
        ''' <returns>结果图片（失败返回1*1个像素，#00000000透明色的图片）</returns>
        ''' <remarks></remarks>
        Public Shared Function Image(ByVal hWnd As IntPtr) As Bitmap
            Try
                Dim WindowPlacement As Placement
                GetWindowPlacement(hWnd, WindowPlacement)
                Dim WindowRect As Rect = WindowPlacement.rcNormalPosition
                Dim SourceDC As IntPtr = GetWindowDC(hWnd)
                Dim SourceMemoryDC As IntPtr = CreateCompatibleDC(SourceDC)
                Dim TempBitmap As IntPtr = CreateCompatibleBitmap(SourceDC, WindowRect.Right - WindowRect.Left, WindowRect.Bottom - WindowRect.Top)
                SelectObject(SourceMemoryDC, TempBitmap)
                PrintWindow(hWnd, SourceMemoryDC, 0)
                Dim Result As Bitmap = Bitmap.FromHbitmap(TempBitmap)
                DeleteObject(TempBitmap)
                DeleteDC(SourceMemoryDC)
                ReleaseDC(hWnd, SourceDC)
                Return Result
            Catch ex As Exception
                Return New Bitmap(1, 1)
            End Try
        End Function

        Private Declare Function PostMessage Lib "user32.dll" Alias "PostMessageA" (ByVal hWnd As IntPtr, ByVal wMsg As UInt32, Optional ByVal wParam As Object = Nothing, Optional ByVal lParam As Object = Nothing) As Boolean

        ''' <summary>
        ''' 关闭窗口（同步阻塞，等待程序处理和用户确认，例如弹出确认关闭的对话框）
        ''' </summary>
        ''' <param name="hWnd">窗口句柄（IntPtr）</param>
        ''' <returns>是否执行成功</returns>
        ''' <remarks></remarks>
        Public Shared Function Close(ByVal hWnd As IntPtr) As Boolean
            Return SendMessage(hWnd, WindowsMessage.Close) = 0 And Runtime.InteropServices.Marshal.GetLastWin32Error() <> 1400 '无效的窗口句柄
        End Function
        ''' <summary>
        ''' 关闭窗口（异步执行，等待程序处理和用户确认，例如弹出确认关闭的对话框）
        ''' </summary>
        ''' <param name="hWnd">窗口句柄（IntPtr）</param>
        ''' <returns>是否执行成功</returns>
        ''' <remarks></remarks>
        Public Shared Function CloseAsync(ByVal hWnd As IntPtr) As Boolean
            Return PostMessage(hWnd, WindowsMessage.Close) And Runtime.InteropServices.Marshal.GetLastWin32Error() <> 1400 '无效的窗口句柄
        End Function

        Private Declare Function FlashWindow Lib "user32.dll" Alias "FlashWindow" (ByVal hWnd As IntPtr, ByVal bInvert As Boolean) As Boolean

        ''' <summary>
        ''' 闪烁窗口（同步阻塞，包括窗体和任务栏按钮的闪烁效果）
        ''' </summary>
        ''' <param name="hWnd">窗口句柄（IntPtr）</param>
        ''' <param name="Times">闪烁的次数（默认1次）</param>
        ''' <param name="IntervalMillisecond">闪烁的时间间隔（默认1秒）</param>
        ''' <returns>是否执行成功</returns>
        ''' <remarks></remarks>
        Public Shared Function Flash(ByVal hWnd As IntPtr, Optional ByVal Times As UInt32 = 1, Optional ByVal IntervalMillisecond As UInt32 = 1000) As Boolean
            If IntervalMillisecond <= 0 Then
                Times = 1
            End If
            If Times <= 0 Then
                Return False
            End If
            For I = 0 To Times - 1
                FlashWindow(hWnd, True)
                FlashWindow(hWnd, False)
                Try
                    System.Threading.Thread.Sleep(1000)
                Catch ex As Exception
                End Try
            Next
            Return True
        End Function
        ''' <summary>
        ''' 闪烁窗口（异步执行，包括窗体和任务栏按钮的闪烁效果）
        ''' </summary>
        ''' <param name="hWnd">窗口句柄（IntPtr）</param>
        ''' <param name="Times">闪烁的次数（默认1次）</param>
        ''' <param name="IntervalMillisecond">闪烁的时间间隔（默认1秒）</param>
        ''' <returns>是否执行成功</returns>
        ''' <remarks></remarks>
        Public Shared Function FlashAsync(ByVal hWnd As IntPtr, Optional ByVal Times As UInt32 = 1, Optional ByVal IntervalMillisecond As UInt32 = 1000) As Boolean
            If hWnd = 0 OrElse IsWindow(hWnd) = False Then
                Return False
            End If
            If IntervalMillisecond <= 0 Then
                Times = 1
            End If
            If Times <= 0 Then
                Return False
            End If
            Dim Temp As New FlashAsyncTask(hWnd, Times, IntervalMillisecond)
            Return True
        End Function
        Private Class FlashAsyncTask
            Private hWnd As IntPtr
            Private Times As UInt32
            Private IntervalMillisecond As UInt32
            Private Thread As New Threading.Thread(New Threading.ThreadStart(AddressOf Run))
            Public Sub New(ByVal TaskhWnd As IntPtr, Optional ByVal TaskTimes As UInt32 = 1, Optional ByVal TaskIntervalMillisecond As UInt32 = 1000)
                hWnd = TaskhWnd
                Times = TaskTimes
                IntervalMillisecond = TaskIntervalMillisecond
                Thread.Start()
            End Sub
            Private Sub Run()
                For I = 0 To Times - 1
                    If IsWindow(hWnd) = False Then
                        Thread.Abort()
                    End If
                    FlashWindow(hWnd, True)
                    FlashWindow(hWnd, False)
                    Try
                        System.Threading.Thread.Sleep(1000)
                    Catch ex As Exception
                    End Try
                Next
            End Sub
        End Class

        ''' <summary>
        ''' 显示或隐藏桌面（效果类似于在Win7系统，用鼠标点击一次屏幕右下角的“显示桌面”）
        ''' </summary>
        ''' <returns>是否执行成功</returns>
        ''' <remarks></remarks>
        Public Shared Function ToggleDesktop() As Boolean
            Try
                CreateObject("Shell.Application").ToggleDesktop()
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function

        ''' <summary>
        ''' 显示或隐藏窗口的3D切换效果（需要Windows Vista及更高版本的系统）
        ''' </summary>
        ''' <returns>是否执行成功</returns>
        ''' <remarks></remarks>
        Public Shared Function Toggle3DSwitcher() As Boolean
            Try
                CreateObject("Shell.Application").WindowSwitcher()
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function

        ''' <summary>
        ''' 显示关闭计算机界面
        ''' </summary>
        ''' <returns>是否执行成功</returns>
        ''' <remarks></remarks>
        Public Shared Function ShowShutDown() As Boolean
            Try
                CreateObject("Shell.Application").ShutdownWindows()
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function

        ''' <summary>
        ''' 显示文件搜索界面
        ''' </summary>
        ''' <returns>是否执行成功</returns>
        ''' <remarks></remarks>
        Public Shared Function ShowSearchFile() As Boolean
            Try
                CreateObject("Shell.Application").FindFiles()
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function

        ''' <summary>
        ''' 显示时间与日期界面
        ''' </summary>
        ''' <returns>是否执行成功</returns>
        ''' <remarks></remarks>
        Public Shared Function ShowSetTime() As Boolean
            Try
                CreateObject("Shell.Application").SetTime()
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function

        ''' <summary>
        ''' 显示任务栏属性界面
        ''' </summary>
        ''' <returns>是否执行成功</returns>
        ''' <remarks></remarks>
        Public Shared Function ShowTrayProperties() As Boolean
            Try
                CreateObject("Shell.Application").TrayProperties()
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function

        ''' <summary>
        ''' 显示运行界面
        ''' </summary>
        ''' <returns>是否执行成功</returns>
        ''' <remarks></remarks>
        Public Shared Function ShowFileRun() As Boolean
            Try
                CreateObject("Shell.Application").FileRun()
                Return True
            Catch ex As Exception
                Return False
            End Try
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
        ''' 挂起窗口线程（Suspend Count加1）
        ''' </summary>
        ''' <param name="hWnd">窗口句柄（IntPtr）</param>
        ''' <returns>是否执行成功</returns>
        ''' <remarks></remarks>
        Public Shared Function ThreadSuspend(ByVal hWnd As IntPtr) As Boolean
            If hWnd = 0 OrElse IsWindow(hWnd) = False Then
                Return False
            End If
            Dim ThreadId As Int32 = GetWindowThreadProcessId(hWnd, Nothing)
            Dim ThreadhWnd As IntPtr = OpenThread(ThreadAccess.All, False, ThreadId)
            Dim Result As Int32 = SuspendThread(ThreadhWnd)
            Return Result <> -1
        End Function
        ''' <summary>
        ''' 恢复窗口线程（Suspend Count减1）
        ''' </summary>
        ''' <param name="hWnd">窗口句柄（IntPtr）</param>
        ''' <returns>是否执行成功</returns>
        ''' <remarks></remarks>
        Public Shared Function ThreadResume(ByVal hWnd As IntPtr) As Boolean
            If hWnd = 0 OrElse IsWindow(hWnd) = False Then
                Return False
            End If
            Dim ThreadId As Int32 = GetWindowThreadProcessId(hWnd, Nothing)
            Dim ThreadhWnd As IntPtr = OpenThread(ThreadAccess.All, False, ThreadId)
            Dim Result As Int32 = ResumeThread(ThreadhWnd)
            Return Result <> -1
        End Function
        ''' <summary>
        ''' 强制恢复窗口线程（Suspend Count减到0）
        ''' </summary>
        ''' <param name="hWnd">窗口句柄（IntPtr）</param>
        ''' <returns>是否执行成功</returns>
        ''' <remarks></remarks>
        Public Shared Function ThreadFree(ByVal hWnd As IntPtr) As Boolean
            If hWnd = 0 OrElse IsWindow(hWnd) = False Then
                Return False
            End If
            Dim ThreadId As Int32 = GetWindowThreadProcessId(hWnd, Nothing)
            Dim ThreadhWnd As IntPtr = OpenThread(ThreadAccess.All, False, ThreadId)
            Dim Temp As Int32 = ResumeThread(ThreadhWnd)
            Dim Result As Boolean = Temp <> -1
            While Temp > 0
                Temp = ResumeThread(ThreadhWnd)
                Result = Result And Temp <> -1
            End While
            Return Result
        End Function

        ''' <summary>
        ''' 限制窗口的CPU占用（通过不断挂起和恢复线程）
        ''' </summary>
        ''' <param name="hWnd">窗口句柄（IntPtr）</param>
        ''' <param name="SleepMillisecond">挂起等待的持续时间（默认50毫秒，最少1毫秒）</param>
        ''' <param name="IntervalMillisecond">挂起恢复的间隔时间（默认50毫秒，最少1毫秒）</param>
        ''' <returns>是否执行成功</returns>
        ''' <remarks></remarks>
        Public Shared Function ThreadLimit(ByVal hWnd As IntPtr, Optional ByVal SleepMillisecond As UInt32 = 50, Optional ByVal IntervalMillisecond As UInt32 = 50) As Boolean
            If hWnd = 0 OrElse IsWindow(hWnd) = False Then
                Return False
            End If
            If SleepMillisecond <= 0 Then
                SleepMillisecond = 1
            End If
            If IntervalMillisecond <= 0 Then
                IntervalMillisecond = 1
            End If
            Dim Temp As New ThreadLimitTask(hWnd, SleepMillisecond, IntervalMillisecond)
            Return True
        End Function
        Private Class ThreadLimitTask
            Private WindowhWnd As IntPtr
            Private ThreadhWnd As IntPtr
            Private SleepMillisecond As UInt32
            Private IntervalMillisecond As UInt32
            Private Thread As New Threading.Thread(New Threading.ThreadStart(AddressOf Run))
            Public Sub New(ByVal TaskhWnd As IntPtr, ByVal TaskSleepMillisecond As UInt32, ByVal TaskIntervalMillisecond As UInt32)
                WindowhWnd = TaskhWnd
                SleepMillisecond = TaskSleepMillisecond
                IntervalMillisecond = TaskIntervalMillisecond
                Dim ThreadId As Int32 = GetWindowThreadProcessId(WindowhWnd, Nothing)
                ThreadhWnd = OpenThread(ThreadAccess.All, False, ThreadId)
                Thread.Start()
            End Sub
            Private Sub Run()
                While Thread.IsAlive
                    If IsWindow(WindowhWnd) = False Then
                        Thread.Abort()
                    End If
                    SuspendThread(ThreadhWnd)
                    Try
                        System.Threading.Thread.Sleep(SleepMillisecond)
                    Catch ex As Exception
                    End Try
                    ResumeThread(ThreadhWnd)
                    Try
                        System.Threading.Thread.Sleep(IntervalMillisecond)
                    Catch ex As Exception
                    End Try
                End While
            End Sub
        End Class

        ''' <summary>
        ''' 获取窗口线程ID
        ''' </summary>
        ''' <param name="hWnd">窗口句柄（IntPtr）</param>
        ''' <returns>结果线程ID</returns>
        ''' <remarks></remarks>
        Public Shared Function GetThreadId(ByVal hWnd As IntPtr) As Int32
            Return GetWindowThreadProcessId(hWnd, Nothing)
        End Function
        ''' <summary>
        ''' 获取窗口进程ID
        ''' </summary>
        ''' <param name="hWnd">窗口句柄（IntPtr）</param>
        ''' <returns>结果进程ID</returns>
        ''' <remarks></remarks>
        Public Shared Function GetProcessId(ByVal hWnd As IntPtr) As Int32
            Dim ProcessId As Int32
            GetWindowThreadProcessId(hWnd, ProcessId)
            Return ProcessId
        End Function
        ''' <summary>
        ''' 获取窗口进程
        ''' </summary>
        ''' <param name="hWnd">窗口句柄（IntPtr）</param>
        ''' <returns>结果进程（Process）</returns>
        ''' <remarks></remarks>
        Public Shared Function GetProcess(ByVal hWnd As IntPtr) As Process
            Dim ProcessId As Int32
            GetWindowThreadProcessId(hWnd, ProcessId)
            If ProcessId <> 0 Then
                Return System.Diagnostics.Process.GetProcessById(ProcessId)
            End If
            Return Nothing
        End Function

    End Class

End Namespace