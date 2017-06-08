Namespace My
    
    ''' <summary>
    ''' 模拟键盘操作相关函数
    ''' </summary>
    ''' <remarks></remarks>
    Partial Public NotInheritable Class Keyboard
        Private Declare Sub keybd_event Lib "user32.dll" Alias "keybd_event" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
        Private Declare Function MapVirtualKey Lib "user32.dll" Alias "MapVirtualKeyA" (ByVal wCode As Long, ByVal wMapType As Long) As Long
        <Flags()> _
        Private Enum KeyEvent As Long
            Down = 0
            Up = 2
        End Enum

        ''' <summary>
        ''' 按下单个键位（保持按下状态，要注意，按下+释放才是一次完整的按键过程）
        ''' </summary>
        ''' <param name="Key">键位（Windows.Forms.Keys）</param>
        ''' <returns>是否执行成功</returns>
        ''' <remarks></remarks>
        Public Shared Function Down(ByVal Key As Keys) As Boolean
            Try
                keybd_event(Key, MapVirtualKey(Key, 0), KeyEvent.Down, 0)
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function
        ''' <summary>
        ''' 释放单个键位（取消按下状态，要注意，按下+释放才是一次完整的按键过程）
        ''' </summary>
        ''' <param name="Key">键位（Windows.Forms.Keys）</param>
        ''' <returns>是否执行成功</returns>
        ''' <remarks></remarks>
        Public Shared Function Up(ByVal Key As Keys) As Boolean
            Try
                keybd_event(Key, MapVirtualKey(Key, 0), KeyEvent.Up, 0)
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function

        ''' <summary>
        ''' 点击单个键位（包括按下+释放过程）
        ''' </summary>
        ''' <param name="Key">键位（Windows.Forms.Keys）</param>
        ''' <returns>是否执行成功</returns>
        ''' <remarks></remarks>
        Public Shared Function Click(ByVal Key As Keys) As Boolean
            Try
                keybd_event(Key, MapVirtualKey(Key, 0), KeyEvent.Down, 0)
                keybd_event(Key, MapVirtualKey(Key, 0), KeyEvent.Up, 0)
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function
        ''' <summary>
        ''' 点击多个键位，完成组合键（包括按下+释放过程）
        ''' </summary>
        ''' <param name="Keys">键位数组（Windows.Forms.Keys）</param>
        ''' <returns>是否执行成功</returns>
        ''' <remarks></remarks>
        Public Shared Function Click(ByVal Keys As Keys()) As Boolean
            Try
                For Each Key As Keys In Keys
                    keybd_event(Key, MapVirtualKey(Key, 0), KeyEvent.Down, 0)
                Next
                For Each Key As Keys In Keys
                    keybd_event(Key, MapVirtualKey(Key, 0), KeyEvent.Up, 0)
                Next
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function
        ''' <summary>
        ''' 点击多个键位，完成组合键（包括按下+释放过程）
        ''' </summary>
        ''' <param name="Keys1">键位1（Windows.Forms.Keys）</param>
        ''' <param name="Keys2">键位2（Windows.Forms.Keys）</param>
        ''' <param name="Keys3">键位3（Windows.Forms.Keys）</param>
        ''' <param name="Keys4">键位4（Windows.Forms.Keys）</param>
        ''' <returns>是否执行成功</returns>
        ''' <remarks></remarks>
        Public Shared Function Click(Keys1 As Keys, ByVal Keys2 As Keys, Optional ByVal Keys3 As Keys = Nothing, Optional ByVal Keys4 As Keys = Nothing) As Boolean
            Try
                Dim Temp As New List(Of Keys)
                Temp.Add(Keys1)
                Temp.Add(Keys2)
                If Keys3 <> Nothing Then
                    Temp.Add(Keys3)
                End If
                If Keys4 <> Nothing Then
                    Temp.Add(Keys4)
                End If
                For Each Key As Keys In Temp
                    keybd_event(Key, MapVirtualKey(Key, 0), KeyEvent.Down, 0)
                Next
                For Each Key As Keys In Temp
                    keybd_event(Key, MapVirtualKey(Key, 0), KeyEvent.Up, 0)
                Next
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function



        Private Declare Function GetAsyncKeyState Lib "user32.dll" Alias "GetAsyncKeyState" (ByVal vKey As Int32) As Int16
        Private Shared ReadOnly STATE_DOWN_1 As Int16 = -32767
        Private Shared ReadOnly STATE_DOWN_2 As Int16 = -32768

        ''' <summary>
        ''' 判断单个键位是否处于按下状态
        ''' </summary>
        ''' <param name="Key">键位（Windows.Forms.Keys）</param>
        ''' <returns>是否处于按下状态</returns>
        ''' <remarks></remarks>
        Public Shared Function Check(ByVal Key As Keys) As Boolean
            Dim Temp As Int16 = GetAsyncKeyState(Key)
            If Temp = STATE_DOWN_1 Or Temp = STATE_DOWN_2 Then
                Return True
            Else
                Return False
            End If
        End Function

        ''' <summary>
        ''' 设置CapsLock的状态
        ''' </summary>
        ''' <param name="State">是否打开CapsLock</param>
        ''' <returns>是否执行成功</returns>
        ''' <remarks></remarks>
        Public Shared Function SetCapsLock(ByVal State As Boolean) As Boolean
            Try
                If Not My.Computer.Keyboard.CapsLock = State Then
                    keybd_event(Keys.CapsLock, MapVirtualKey(Keys.CapsLock, 0), KeyEvent.Down, 0)
                    keybd_event(Keys.CapsLock, MapVirtualKey(Keys.CapsLock, 0), KeyEvent.Up, 0)
                End If
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function
        ''' <summary>
        ''' 设置ScrollLock的状态
        ''' </summary>
        ''' <param name="State">是否打开ScrollLock</param>
        ''' <returns>是否执行成功</returns>
        ''' <remarks></remarks>
        Public Shared Function SetScrollLock(ByVal State As Boolean) As Boolean
            Try
                If Not My.Computer.Keyboard.ScrollLock = State Then
                    keybd_event(Keys.Scroll, MapVirtualKey(Keys.Scroll, 0), KeyEvent.Down, 0)
                    keybd_event(Keys.Scroll, MapVirtualKey(Keys.Scroll, 0), KeyEvent.Up, 0)
                End If
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function
        ''' <summary>
        ''' 设置NumLock的状态
        ''' </summary>
        ''' <param name="State">是否打开NumLock</param>
        ''' <returns>是否执行成功</returns>
        ''' <remarks></remarks>
        Public Shared Function SetNumLock(ByVal State As Boolean) As Boolean
            Try
                If Not My.Computer.Keyboard.NumLock = State Then
                    keybd_event(Keys.NumLock, MapVirtualKey(Keys.NumLock, 0), KeyEvent.Down, 0)
                    keybd_event(Keys.NumLock, MapVirtualKey(Keys.NumLock, 0), KeyEvent.Up, 0)
                End If
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function



        ''' <summary>
        ''' 点击多个键位，输入一段字符串（包括按下+释放过程，限定字符串内容）
        ''' </summary>
        ''' <param name="KeyString">键位字符串（只支持输入字母、数字、空格、换行、键盘上有的英文特殊符号，其它字符会被忽略）</param>
        ''' <param name="MillisecondsInterval">输入每个字符的时间间隔（单位毫秒，默认值为0，无时间间隔）</param>
        ''' <returns>是否执行成功</returns> 
        ''' <remarks></remarks>
        Public Shared Function Input(ByVal KeyString As String, Optional ByVal MillisecondsInterval As Integer = 0) As Boolean
            Dim AscString As String = "QWERTYUIOP" & "ASDFGHJKL" & "ZXCVBNM" & "1234567890" & " " & vbCrLf
            Dim LowerString As String = "qwertyuiop" & "asdfghjkl" & "zxcvbnm"
            Dim OemString As String = ";=,-./`[\]'"
            Dim ShiftOemString As String = ":+<_>?~{|}"""
            Dim OemKeys As Keys() = New Keys() {Keys.Oem1, Keys.Oemplus, Keys.Oemcomma, Keys.OemMinus, Keys.OemPeriod, Keys.Oem2, Keys.Oem3, Keys.Oem4, Keys.Oem5, Keys.Oem6, Keys.Oem7}
            Dim ShiftNumString As String = "!@#$%^&*()"
            Dim NumKeys As Keys() = New Keys() {Keys.D1, Keys.D2, Keys.D3, Keys.D4, Keys.D5, Keys.D6, Keys.D7, Keys.D8, Keys.D9, Keys.D0}
            Try
                Dim KeyArray As Char() = KeyString.ToCharArray()
                For N = 0 To KeyArray.Length - 1
                    Dim Key As Char = KeyArray(N)
                    If N > 0 And MillisecondsInterval <> 0 Then
                        System.Threading.Thread.Sleep(MillisecondsInterval)
                    End If
                    If AscString.Contains(Key) Then
                        '大写字母、数字、空格（虚拟键码VK值，与字符ASCII值相同）
                        If Not My.Computer.Keyboard.CapsLock = True Then
                            keybd_event(Keys.CapsLock, MapVirtualKey(Keys.CapsLock, 0), KeyEvent.Down, 0)
                            keybd_event(Keys.CapsLock, MapVirtualKey(Keys.CapsLock, 0), KeyEvent.Up, 0)
                        End If
                        keybd_event(Asc(Key), MapVirtualKey(Asc(Key), 0), KeyEvent.Down, 0)
                        keybd_event(Asc(Key), MapVirtualKey(Asc(Key), 0), KeyEvent.Up, 0)
                    ElseIf LowerString.Contains(Key) Then
                        '小写字母
                        If Not My.Computer.Keyboard.CapsLock = False Then
                            keybd_event(Keys.CapsLock, MapVirtualKey(Keys.CapsLock, 0), KeyEvent.Down, 0)
                            keybd_event(Keys.CapsLock, MapVirtualKey(Keys.CapsLock, 0), KeyEvent.Up, 0)
                        End If
                        Key = Key.ToString.ToUpper()
                        keybd_event(Asc(Key), MapVirtualKey(Asc(Key), 0), KeyEvent.Down, 0)
                        keybd_event(Asc(Key), MapVirtualKey(Asc(Key), 0), KeyEvent.Up, 0)
                    ElseIf OemString.Contains(Key) Then
                        'OEM键特殊符号
                        Dim I As Integer = OemString.IndexOf(Key)
                        keybd_event(OemKeys(I), MapVirtualKey(OemKeys(I), 0), KeyEvent.Down, 0)
                        keybd_event(OemKeys(I), MapVirtualKey(OemKeys(I), 0), KeyEvent.Up, 0)
                    ElseIf ShiftOemString.Contains(Key) Then
                        'Shift+OEM键特殊符号
                        Dim I As Integer = ShiftOemString.IndexOf(Key)
                        keybd_event(Keys.ShiftKey, MapVirtualKey(Keys.ShiftKey, 0), KeyEvent.Down, 0)
                        keybd_event(OemKeys(I), MapVirtualKey(OemKeys(I), 0), KeyEvent.Down, 0)
                        keybd_event(OemKeys(I), MapVirtualKey(OemKeys(I), 0), KeyEvent.Up, 0)
                        keybd_event(Keys.ShiftKey, MapVirtualKey(Keys.ShiftKey, 0), KeyEvent.Up, 0)
                    ElseIf ShiftNumString.Contains(Key) Then
                        'Shift+数字键特殊符号
                        Dim I As Integer = ShiftNumString.IndexOf(Key)
                        keybd_event(Keys.ShiftKey, MapVirtualKey(Keys.ShiftKey, 0), KeyEvent.Down, 0)
                        keybd_event(NumKeys(I), MapVirtualKey(NumKeys(I), 0), KeyEvent.Down, 0)
                        keybd_event(NumKeys(I), MapVirtualKey(NumKeys(I), 0), KeyEvent.Up, 0)
                        keybd_event(Keys.ShiftKey, MapVirtualKey(Keys.ShiftKey, 0), KeyEvent.Up, 0)
                    End If
                Next
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function

        ''' <summary>
        ''' 连续复制粘贴字符，输入一段字符串（使用Ctrl+V组合键，速度太快时可能出错）
        ''' </summary>
        ''' <param name="Source">要输入的字符串</param>
        ''' <param name="MillisecondsInterval">输入每个字符的时间间隔（单位毫秒，默认值为100，实测，在值较小、系统卡顿时，可能会发生字符混乱，建议设置为20以上）</param>
        ''' <returns>是否执行成功</returns>
        ''' <remarks></remarks>
        Public Shared Function PasteDelay(ByVal Source As String, Optional ByVal MillisecondsInterval As Integer = 100) As Boolean
            Try
                Dim UpperString As String = "QWERTYUIOP" & "ASDFGHJKL" & "ZXCVBNM"
                For I = 0 To Source.Length - 1
                    If I > 0 And MillisecondsInterval <> 0 Then
                        System.Threading.Thread.Sleep(MillisecondsInterval)
                    End If
                    If UpperString.Contains(Source(I)) Then
                        '大写字母
                        If Not My.Computer.Keyboard.CapsLock = True Then
                            keybd_event(Keys.CapsLock, MapVirtualKey(Keys.CapsLock, 0), KeyEvent.Down, 0)
                            keybd_event(Keys.CapsLock, MapVirtualKey(Keys.CapsLock, 0), KeyEvent.Up, 0)
                        End If
                    End If
                    System.Windows.Forms.Clipboard.SetText(Source(I))
                    keybd_event(Keys.ControlKey, MapVirtualKey(Keys.ControlKey, 0), KeyEvent.Down, 0)
                    keybd_event(Keys.V, MapVirtualKey(Keys.V, 0), KeyEvent.Down, 0)
                    keybd_event(Keys.V, MapVirtualKey(Keys.V, 0), KeyEvent.Up, 0)
                    keybd_event(Keys.ControlKey, MapVirtualKey(Keys.ControlKey, 0), KeyEvent.Up, 0)
                Next
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function



        ''' <summary>
        ''' 复制粘贴，输入一段字符串（使用Ctrl+V组合键）
        ''' </summary>
        ''' <param name="Source">要输入的字符串</param>
        ''' <returns>是否执行成功</returns>
        ''' <remarks></remarks>
        Public Shared Function Paste(ByVal Source As String) As Boolean
            Try
                System.Windows.Forms.Clipboard.SetText(Source)
                keybd_event(Keys.ControlKey, MapVirtualKey(Keys.ControlKey, 0), KeyEvent.Down, 0)
                keybd_event(Keys.V, MapVirtualKey(Keys.V, 0), KeyEvent.Down, 0)
                keybd_event(Keys.V, MapVirtualKey(Keys.V, 0), KeyEvent.Up, 0)
                keybd_event(Keys.ControlKey, MapVirtualKey(Keys.ControlKey, 0), KeyEvent.Up, 0)
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function
        ''' <summary>
        ''' 复制粘贴，输入一个图片（使用Ctrl+V组合键）
        ''' </summary>
        ''' <param name="Source">要输入的图片</param>
        ''' <returns>是否执行成功</returns>
        ''' <remarks></remarks>
        Public Shared Function Paste(ByVal Source As Bitmap) As Boolean
            Try
                System.Windows.Forms.Clipboard.SetImage(Source)
                keybd_event(Keys.ControlKey, MapVirtualKey(Keys.ControlKey, 0), KeyEvent.Down, 0)
                keybd_event(Keys.V, MapVirtualKey(Keys.V, 0), KeyEvent.Down, 0)
                keybd_event(Keys.V, MapVirtualKey(Keys.V, 0), KeyEvent.Up, 0)
                keybd_event(Keys.ControlKey, MapVirtualKey(Keys.ControlKey, 0), KeyEvent.Up, 0)
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function

    End Class

End Namespace