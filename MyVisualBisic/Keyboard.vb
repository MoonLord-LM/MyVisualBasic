Namespace My
    
    ''' <summary>
    ''' 模拟键盘操作相关函数
    ''' </summary>
    ''' <remarks></remarks>
    Public NotInheritable Class Keyboard

        Private Declare Sub keybd_event Lib "user32.dll" Alias "keybd_event" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
        Private Declare Function MapVirtualKey Lib "user32.dll" Alias "MapVirtualKeyA" (ByVal wCode As Long, ByVal wMapType As Long) As Long
        Private Shared ReadOnly KEY_DOWN As Long = 0
        Private Shared ReadOnly KEY_UP As Long = 2

        ''' <summary>
        ''' 按下单个键位（保持按下状态，要注意，按下Press+释放Release才是一次完整的按键过程）
        ''' </summary>
        ''' <param name="Key">键位（Windows.Forms.Keys）</param>
        ''' <returns>是否执行成功</returns>
        ''' <remarks></remarks>
        Public Shared Function Down(ByVal Key As Keys) As Boolean
            Try
                keybd_event(Key, MapVirtualKey(Key, 0), KEY_DOWN, 0)
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function
        ''' <summary>
        ''' 释放单个键位（取消按下状态，要注意，按下Press+释放Release才是一次完整的按键过程）
        ''' </summary>
        ''' <param name="Key">键位（Windows.Forms.Keys）</param>
        ''' <returns>是否执行成功</returns>
        ''' <remarks></remarks>
        Public Shared Function Up(ByVal Key As Keys) As Boolean
            Try
                keybd_event(Key, MapVirtualKey(Key, 0), KEY_UP, 0)
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function
        ''' <summary>
        ''' 点击单个键位（包括按下Press+释放Release过程）
        ''' </summary>
        ''' <param name="Key">键位（Windows.Forms.Keys）</param>
        ''' <returns>是否执行成功</returns>
        ''' <remarks></remarks>
        Public Shared Function Click(ByVal Key As Keys) As Boolean
            Try
                keybd_event(Key, MapVirtualKey(Key, 0), KEY_DOWN, 0)
                keybd_event(Key, MapVirtualKey(Key, 0), KEY_UP, 0)
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
                    keybd_event(Keys.CapsLock, MapVirtualKey(Keys.CapsLock, 0), KEY_DOWN, 0)
                    keybd_event(Keys.CapsLock, MapVirtualKey(Keys.CapsLock, 0), KEY_UP, 0)
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
                    keybd_event(Keys.Scroll, MapVirtualKey(Keys.Scroll, 0), KEY_DOWN, 0)
                    keybd_event(Keys.Scroll, MapVirtualKey(Keys.Scroll, 0), KEY_UP, 0)
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
                    keybd_event(Keys.NumLock, MapVirtualKey(Keys.NumLock, 0), KEY_DOWN, 0)
                    keybd_event(Keys.NumLock, MapVirtualKey(Keys.NumLock, 0), KEY_UP, 0)
                End If
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function



        ''' <summary>
        ''' 点击多个键位（包括按下Press+释放Release过程）
        ''' </summary>
        ''' <param name="KeyString">键位（只允许字母、数字、空格、换行组成的字符串）</param>
        ''' <returns>是否执行成功</returns>
        ''' <remarks></remarks>
        Public Shared Function Click(ByVal KeyString As String) As Boolean
            '大写字母、数字、空格、换行，其Windows.Forms.Keys值，和ASCII值一致
            Dim AscString As String = "QWERTYUIOP" & "ASDFGHJKL" & "ZXCVBNM" & "1234567890" & " " & vbCrLf
            Dim LowerString As String = "qwertyuiop" & "asdfghjkl" & "zxcvbnm"
            Dim SpecialString As String() = New String() {}
            Dim SpecialKeys As Keys() = New Keys() {}
            Try
                For Each Key As Char In KeyString.ToCharArray()
                    If AscString.Contains(Key) Then
                        If Not My.Computer.Keyboard.CapsLock = True Then
                            keybd_event(Keys.CapsLock, MapVirtualKey(Keys.CapsLock, 0), KEY_DOWN, 0)
                            keybd_event(Keys.CapsLock, MapVirtualKey(Keys.CapsLock, 0), KEY_UP, 0)
                        End If
                        keybd_event(Asc(Key), MapVirtualKey(Asc(Key), 0), KEY_DOWN, 0)
                        keybd_event(Asc(Key), MapVirtualKey(Asc(Key), 0), KEY_UP, 0)
                    ElseIf LowerString.Contains(Key) Then
                        If Not My.Computer.Keyboard.CapsLock = False Then
                            keybd_event(Keys.CapsLock, MapVirtualKey(Keys.CapsLock, 0), KEY_DOWN, 0)
                            keybd_event(Keys.CapsLock, MapVirtualKey(Keys.CapsLock, 0), KEY_UP, 0)
                        End If
                        Key = Key.ToString.ToUpper()
                        keybd_event(Asc(Key), MapVirtualKey(Asc(Key), 0), KEY_DOWN, 0)
                        keybd_event(Asc(Key), MapVirtualKey(Asc(Key), 0), KEY_UP, 0)
                    Else

                    End If
                Next
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function

    End Class

End Namespace