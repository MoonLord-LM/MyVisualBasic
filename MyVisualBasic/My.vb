Namespace My

    ''' <summary>
    ''' 待整理
    ''' </summary>
    ''' <remarks></remarks>
    Partial Class MyComputer

        '
        'SendMessage WM_KEYUP, WM_CHAR, and WM_KEYDOWN

        Public Declare Function SetWindowRgn Lib "user32.dll" Alias "SetWindowRgn" (ByVal hWnd As IntPtr, ByVal hRgn As IntPtr, ByVal bRedraw As Boolean) As Int32

        ''' <summary>
        ''' 获取上一个Win32 API调用产生的错误代码（实测：出现错误后，错误信息会一直保留，直到被下一个错误信息替换）
        ''' </summary>
        ''' <returns>错误代码（默认为0）</returns>
        ''' <remarks></remarks>
        Public Shared Function Win32ErrorCode() As Integer
            Return Runtime.InteropServices.Marshal.GetLastWin32Error()
        End Function
        ''' <summary>
        ''' 获取上一个Win32 API调用产生的错误说明（实测：出现错误后，错误信息会一直保留，直到被下一个错误信息替换）
        ''' </summary>
        ''' <returns>错误说明（默认为"操作成功完成。"）</returns>
        ''' <remarks></remarks>
        Public Shared Function Win32ErrorMessage() As String
            Return New System.ComponentModel.Win32Exception(Runtime.InteropServices.Marshal.GetLastWin32Error()).Message
        End Function

    End Class

End Namespace