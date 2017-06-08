Namespace My
    
    ''' <summary>
    ''' HTTP网络请求相关函数
    ''' </summary>
    ''' <remarks></remarks>
    Partial Public NotInheritable Class Http

        ''' <summary>
        ''' 获取字符串
        ''' </summary>
        ''' <param name="URL">网址链接</param>
        ''' <returns>结果字符串（失败返回空字符串）</returns>
        ''' <remarks></remarks>
        Public Shared Function GetString(ByVal URL As String) As String
            Dim Client As System.Net.WebClient = New System.Net.WebClient()
            Client.Encoding = System.Text.Encoding.UTF8
            Try
                Return Client.DownloadString(URL)
            Catch ex As Exception
                Return ""
            End Try
        End Function
        ''' <summary>
        ''' 获取字符串
        ''' </summary>
        ''' <param name="URL">网址链接</param>
        ''' <param name="Encoding">使用特定的字符编码（默认UTF-8）</param>
        ''' <returns>结果字符串（失败返回空字符串）</returns>
        ''' <remarks></remarks>
        Public Shared Function GetString(ByVal URL As String, ByVal Encoding As System.Text.Encoding) As String
            Dim Client As System.Net.WebClient = New System.Net.WebClient()
            Client.Encoding = Encoding
            Try
                Return Client.DownloadString(URL)
            Catch ex As Exception
                Return ""
            End Try
        End Function



        ''' <summary>
        ''' 获取字节数组
        ''' </summary>
        ''' <param name="URL">网址链接</param>
        ''' <returns>结果Byte数组（失败返回空Byte数组）</returns>
        ''' <remarks></remarks>
        Public Shared Function GetByte(ByVal URL As String) As Byte()
            Dim Client As System.Net.WebClient = New System.Net.WebClient()
            Try
                Return Client.DownloadData(URL)
            Catch ex As Exception
                Return New Byte() {}
            End Try
        End Function



        ''' <summary>
        ''' 下载文件
        ''' </summary>
        ''' <param name="URL">文件链接</param>
        ''' <param name="FilePath">保存到的文件路径（可以是相对路径）</param>
        ''' <returns>是否下载成功</returns>
        ''' <remarks></remarks>
        Public Shared Function DownloadFile(ByVal URL As String, ByVal FilePath As String) As Boolean
            Dim Client As System.Net.WebClient = New System.Net.WebClient()
            Try
                Client.DownloadFile(New System.Uri(URL), FilePath)
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function

    End Class

End Namespace