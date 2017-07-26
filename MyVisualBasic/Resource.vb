Namespace My

    ''' <summary>
    ''' 程序资源文件相关函数
    ''' </summary>
    ''' <remarks></remarks>
    Partial Public NotInheritable Class Resource

        ''' <summary>
        ''' 读取程序嵌入的资源文件（注意必须在解决方案资源管理器中，将资源文件的"属性"-"生成操作"设置为"嵌入的资源"）
        ''' </summary>
        ''' <param name="ResourceName">资源文件名称（注意这个参数的值为资源文件在工程内的相对路径，例如A.zip文件在Resources文件夹内，则此处的参数应为"Resources.A.zip"）</param>
        ''' <returns>结果Byte数组（失败返回空Byte数组）</returns>
        ''' <remarks></remarks>
        Public Shared Function ReadByte(ByVal ResourceName As String) As Byte()
            Dim Reader As System.IO.Stream
            Try
                Dim MyAssembly As System.Reflection.Assembly = System.Reflection.Assembly.GetExecutingAssembly()
                Reader = MyAssembly.GetManifestResourceStream(MyAssembly.GetName().Name & "." & ResourceName)
                Dim Temp(Reader.Length - 1) As Byte
                Reader.Read(Temp, 0, Temp.Length)
                Reader.Dispose()
                Return Temp
            Catch ex As Exception
                Return New Byte() {}
            End Try
        End Function

        ''' <summary>
        ''' 读取程序嵌入的图片类型的资源文件（注意必须在解决方案资源管理器中，将资源文件的"属性"-"生成操作"设置为"嵌入的资源"）
        ''' </summary>
        ''' <param name="ResourceName">资源文件名称（注意这个参数的值为资源文件在工程内的相对路径，例如A.jpg文件在Resources文件夹内，则此处的参数应为"Resources.A.jpg"）</param>
        ''' <returns>结果图片（失败返回1*1个像素，#00000000透明色的图片）</returns>
        ''' <remarks></remarks>
        Public Shared Function ReadPicture(ByVal ResourceName As String) As Bitmap
            Try
                Dim MyAssembly As System.Reflection.Assembly = System.Reflection.Assembly.GetExecutingAssembly()
                Return New Bitmap(MyAssembly.GetManifestResourceStream(MyAssembly.GetName().Name & "." & ResourceName))
            Catch ex As Exception
                Return New Bitmap(1, 1)
            End Try
        End Function

        ''' <summary>
        ''' 读取程序嵌入的文本类型的资源文件（注意必须在解决方案资源管理器中，将资源文件的"属性"-"生成操作"设置为"嵌入的资源"）
        ''' </summary>
        ''' <param name="ResourceName">资源文件名称（注意这个参数的值为资源文件在工程内的相对路径，例如A.txt文件在Resources文件夹内，则此处的参数应为"Resources.A.txt"）</param>
        ''' <returns>结果字符串（失败返回空字符串）</returns>
        ''' <remarks></remarks>
        Public Shared Function ReadString(ByVal ResourceName As String) As String
            Dim Reader As System.IO.Stream
            Try
                Dim MyAssembly As System.Reflection.Assembly = System.Reflection.Assembly.GetExecutingAssembly()
                Reader = MyAssembly.GetManifestResourceStream(MyAssembly.GetName().Name & "." & ResourceName)
                Dim Temp(Reader.Length - 1) As Byte
                Reader.Read(Temp, 0, Temp.Length)
                Reader.Dispose()
                Return System.Text.Encoding.UTF8.GetString(Temp)
            Catch ex As Exception
                Return ""
            End Try
        End Function
        ''' <summary>
        ''' 读取程序嵌入的文本类型的资源文件（注意必须在解决方案资源管理器中，将资源文件的"属性"-"生成操作"设置为"嵌入的资源"）
        ''' </summary>
        ''' <param name="ResourceName">资源文件名称（注意这个参数的值为资源文件在工程内的相对路径，例如A.txt文件在Resources文件夹内，则此处的参数应为"Resources.A.txt"）</param>
        ''' <param name="Encoding">使用特定的字符编码（默认UTF-8）</param>
        ''' <returns>结果字符串（失败返回空字符串）</returns>
        ''' <remarks></remarks>
        Public Shared Function ReadString(ByVal ResourceName As String, ByVal Encoding As System.Text.Encoding) As String
            Dim Reader As System.IO.Stream
            Try
                Dim MyAssembly As System.Reflection.Assembly = System.Reflection.Assembly.GetExecutingAssembly()
                Reader = MyAssembly.GetManifestResourceStream(MyAssembly.GetName().Name & "." & ResourceName)
                Dim Temp(Reader.Length - 1) As Byte
                Reader.Read(Temp, 0, Temp.Length)
                Reader.Dispose()
                Return System.Text.Encoding.UTF8.GetString(Temp)
            Catch ex As Exception
                Return ""
            End Try
        End Function

        ''' <summary>
        ''' 读取程序嵌入的字符串数组类型的资源文件（注意必须在解决方案资源管理器中，将资源文件的"属性"-"生成操作"设置为"嵌入的资源"）
        ''' </summary>
        ''' <param name="ResourceName">资源文件名称（注意这个参数的值为资源文件在工程内的相对路径，例如A.txt文件在Resources文件夹内，则此处的参数应为"Resources.A.txt"）</param>
        ''' <returns>结果字符串数组（失败返回空String数组）</returns>
        ''' <remarks></remarks>
        Public Shared Function ReadStringArray(ByVal ResourceName As String) As String()
            Dim Reader As System.IO.Stream
            Try
                Dim MyAssembly As System.Reflection.Assembly = System.Reflection.Assembly.GetExecutingAssembly()
                Reader = MyAssembly.GetManifestResourceStream(MyAssembly.GetName().Name & "." & ResourceName)
                Dim I(CInt(Reader.Length - 1)) As Byte
                Reader.Read(I, 0, I.Length)
                Reader.Dispose()
                Return System.Text.Encoding.UTF8.GetString(I).Replace(Chr(13) + Chr(10), Chr(10)).Split(New Char() {Chr(10)})
            Catch ex As Exception
                Return New String() {}
            End Try
        End Function
        ''' <summary>
        ''' 读取程序嵌入的字符串数组类型的资源文件（注意必须在解决方案资源管理器中，将资源文件的"属性"-"生成操作"设置为"嵌入的资源"）
        ''' </summary>
        ''' <param name="ResourceName">资源文件名称（注意这个参数的值为资源文件在工程内的相对路径，例如A.txt文件在Resources文件夹内，则此处的参数应为"Resources.A.txt"）</param>
        ''' <param name="Encoding">使用特定的字符编码（默认UTF-8）</param>
        ''' <returns>结果字符串数组（失败返回空String数组）</returns>
        ''' <remarks></remarks>
        Public Shared Function ReadStringArray(ByVal ResourceName As String, ByVal Encoding As System.Text.Encoding) As String()
            Dim Reader As System.IO.Stream
            Try
                Dim MyAssembly As System.Reflection.Assembly = System.Reflection.Assembly.GetExecutingAssembly()
                Reader = MyAssembly.GetManifestResourceStream(MyAssembly.GetName().Name & "." & ResourceName)
                Dim I(CInt(Reader.Length - 1)) As Byte
                Reader.Read(I, 0, I.Length)
                Reader.Dispose()
                Return Encoding.GetString(I).Replace(Chr(13) + Chr(10), Chr(10)).Split(New Char() {Chr(10)})
            Catch ex As Exception
                Return New String() {}
            End Try
        End Function

    End Class

End Namespace