Namespace My
    '2015/5/10
    'My命名空间的扩展

    ''' <summary>
    ''' 编码与解码函数
    ''' </summary>
    ''' <remarks></remarks>
    Public NotInheritable Class Security
        ''' <summary>
        ''' Base64加密算法
        ''' </summary>
        ''' <param name="Source">要加密的字符串</param>
        ''' <returns>加密后的结果字符串</returns>
        ''' <remarks></remarks>
        Public Shared Function Base64_Encode(ByVal Source As String) As String
            Return Convert.ToBase64String(System.Text.Encoding.UTF8.GetBytes(Source))
        End Function
        ''' <summary>
        ''' Base64加密算法
        ''' </summary>
        ''' <param name="Source">要加密的字符串</param>
        ''' <param name="Encoding">使用的编码格式（默认UTF-8）</param>
        ''' <returns>加密后的结果字符串</returns>
        ''' <remarks></remarks>
        Public Shared Function Base64_Encode(ByVal Source As String, ByVal Encoding As System.Text.Encoding) As String
            Return Convert.ToBase64String(Encoding.GetBytes(Source))
        End Function
        ''' <summary>
        ''' Base64解密算法
        ''' </summary>
        ''' <param name="Source">要解密的字符串</param>
        ''' <returns>解密后的结果字符串</returns>
        ''' <remarks></remarks>
        Public Shared Function Base64_Decode(ByVal Source As String) As String
            Return System.Text.Encoding.UTF8.GetString(Convert.FromBase64String(Source))
        End Function
        ''' <summary>
        ''' Base64解密算法
        ''' </summary>
        ''' <param name="Source">要解密的字符串</param>
        ''' <param name="Encoding">使用的编码格式（默认UTF-8）</param>
        ''' <returns>解密后的结果字符串</returns>
        ''' <remarks></remarks>
        Public Shared Function Base64_Decode(ByVal Source As String, ByVal Encoding As System.Text.Encoding) As String
            Return Encoding.GetString(Convert.FromBase64String(Source))
        End Function
        ''' <summary>
        ''' MD5加密算法（返回16位小写结果）
        ''' </summary>
        ''' <param name="Source">要加密的字符串</param>
        ''' <returns>加密后的结果字符串</returns>
        ''' <remarks></remarks>
        Public Shared Function MD5_Lower16_Encode(ByVal Source As String) As String
            Return System.Web.Security.FormsAuthentication.HashPasswordForStoringInConfigFile(Source, "MD5").Substring(8, 16).ToLower()
        End Function
        ''' <summary>
        ''' MD5加密算法（返回16位大写结果）
        ''' </summary>
        ''' <param name="Source">要加密的字符串</param>
        ''' <returns>加密后的结果字符串</returns>
        ''' <remarks></remarks>
        Public Shared Function MD5_Upper16_Encode(ByVal Source As String) As String
            Return System.Web.Security.FormsAuthentication.HashPasswordForStoringInConfigFile(Source, "MD5").Substring(8, 16)
        End Function
        ''' <summary>
        ''' MD5加密算法（返回32位小写结果）
        ''' </summary>
        ''' <param name="Source">要加密的字符串</param>
        ''' <returns>加密后的结果字符串</returns>
        ''' <remarks></remarks>
        Public Shared Function MD5_Lower32_Encode(ByVal Source As String) As String
            Return System.Web.Security.FormsAuthentication.HashPasswordForStoringInConfigFile(Source, "MD5").ToLower()
        End Function
        ''' <summary>
        ''' MD5加密算法（返回32位大写结果）
        ''' </summary>
        ''' <param name="Source">要加密的字符串</param>
        ''' <returns>加密后的结果字符串</returns>
        ''' <remarks></remarks>
        Public Shared Function MD5_Upper32_Encode(ByVal Source As String) As String
            Return System.Web.Security.FormsAuthentication.HashPasswordForStoringInConfigFile(Source, "MD5")
        End Function
        ''' <summary>
        ''' SHA1加密算法（返回40位小写结果）
        ''' </summary>
        ''' <param name="Source">要加密的字符串</param>
        ''' <returns>加密后的结果字符串</returns>
        ''' <remarks></remarks>
        Public Shared Function SHA1_Lower40_Encode(ByVal Source As String) As String
            Return BitConverter.ToString((New System.Security.Cryptography.SHA1CryptoServiceProvider).ComputeHash(System.Text.Encoding.UTF8.GetBytes(Source))).Replace("-", "").ToLower()
        End Function
        ''' <summary>
        ''' SHA1加密算法（返回40位大写结果）
        ''' </summary>
        ''' <param name="Source">要加密的字符串</param>
        ''' <returns>加密后的结果字符串</returns>
        ''' <remarks></remarks>
        Public Shared Function SHA1_Upper40_Encode(ByVal Source As String) As String
            Return BitConverter.ToString((New System.Security.Cryptography.SHA1CryptoServiceProvider).ComputeHash(System.Text.Encoding.UTF8.GetBytes(Source))).Replace("-", "")
        End Function
        ''' <summary>
        ''' SHA256加密算法（返回64位小写结果）
        ''' </summary>
        ''' <param name="Source">要加密的字符串</param>
        ''' <returns>加密后的结果字符串</returns>
        ''' <remarks></remarks>
        Public Shared Function SHA256_Lower64_Encode(ByVal Source As String) As String
            Return BitConverter.ToString((New System.Security.Cryptography.SHA256Managed).ComputeHash(System.Text.Encoding.UTF8.GetBytes(Source))).Replace("-", "").ToLower()
        End Function
        ''' <summary>
        ''' SHA256加密算法（返回64位大写结果）
        ''' </summary>
        ''' <param name="Source">要加密的字符串</param>
        ''' <returns>加密后的结果字符串</returns>
        ''' <remarks></remarks>
        Public Shared Function SHA256_Upper64_Encode(ByVal Source As String) As String
            Return BitConverter.ToString((New System.Security.Cryptography.SHA256Managed).ComputeHash(System.Text.Encoding.UTF8.GetBytes(Source))).Replace("-", "")
        End Function
        ''' <summary>
        ''' SHA384加密算法（返回96位小写结果）
        ''' </summary>
        ''' <param name="Source">要加密的字符串</param>
        ''' <returns>加密后的结果字符串</returns>
        ''' <remarks></remarks>
        Public Shared Function SHA384_Lower96_Encode(ByVal Source As String) As String
            Return BitConverter.ToString((New System.Security.Cryptography.SHA384Managed).ComputeHash(System.Text.Encoding.UTF8.GetBytes(Source))).Replace("-", "").ToLower()
        End Function
        ''' <summary>
        ''' SHA384加密算法（返回96位大写结果）
        ''' </summary>
        ''' <param name="Source">要加密的字符串</param>
        ''' <returns>加密后的结果字符串</returns>
        ''' <remarks></remarks>
        Public Shared Function SHA384_Upper96_Encode(ByVal Source As String) As String
            Return BitConverter.ToString((New System.Security.Cryptography.SHA384Managed).ComputeHash(System.Text.Encoding.UTF8.GetBytes(Source))).Replace("-", "")
        End Function
        ''' <summary>
        ''' SHA512加密算法（返回128位小写结果）
        ''' </summary>
        ''' <param name="Source">要加密的字符串</param>
        ''' <returns>加密后的结果字符串</returns>
        ''' <remarks></remarks>
        Public Shared Function SHA512_Lower128_Encode(ByVal Source As String) As String
            Return BitConverter.ToString((New System.Security.Cryptography.SHA512Managed).ComputeHash(System.Text.Encoding.UTF8.GetBytes(Source))).Replace("-", "").ToLower()
        End Function
        ''' <summary>
        ''' SHA512加密算法（返回128位大写结果）
        ''' </summary>
        ''' <param name="Source">要加密的字符串</param>
        ''' <returns>加密后的结果字符串</returns>
        ''' <remarks></remarks>
        Public Shared Function SHA512_Upper128_Encode(ByVal Source As String) As String
            Return BitConverter.ToString((New System.Security.Cryptography.SHA512Managed).ComputeHash(System.Text.Encoding.UTF8.GetBytes(Source))).Replace("-", "")
        End Function
        ''' <summary>
        ''' DES加密算法
        ''' </summary>
        ''' <param name="Source">要加密的字符串</param>
        ''' <param name="SecretKey">加密密钥（8的整数倍字节数的字符串）</param>
        ''' <returns>加密后的结果字符串</returns>
        ''' <remarks></remarks>
        Public Shared Function DES_Encode(ByVal Source As String, ByVal SecretKey As String) As Byte()
            Dim DES As New System.Security.Cryptography.DESCryptoServiceProvider()
            Try
                DES.Key = System.Text.Encoding.UTF8.GetBytes(SecretKey)
            Catch ex As System.ArgumentException '加密失败（通常是由于SecretKey的字节数不是8的倍数）
                Return New Byte() {}
            End Try
            DES.IV = System.Text.Encoding.UTF8.GetBytes(SecretKey)
            Dim MS As New System.IO.MemoryStream()
            Dim CS As New System.Security.Cryptography.CryptoStream(MS, DES.CreateEncryptor, System.Security.Cryptography.CryptoStreamMode.Write)
            CS.Write(System.Text.Encoding.UTF8.GetBytes(Source), 0, System.Text.Encoding.UTF8.GetBytes(Source).Length)
            CS.FlushFinalBlock()
            Return MS.ToArray()
        End Function
        ''' <summary>
        ''' DES解密算法
        ''' </summary>
        ''' <param name="Source">要解密的字符串</param>
        ''' <param name="SecretKey">解密密钥（8的整数倍字节数的字符串）</param>
        ''' <returns>解密后的结果字符串</returns>
        ''' <remarks></remarks>
        Public Shared Function DES_Decode(ByVal Source As Byte(), ByVal SecretKey As String) As String
            Dim DES As New System.Security.Cryptography.DESCryptoServiceProvider()
            DES.Key = System.Text.Encoding.UTF8.GetBytes(SecretKey)
            DES.IV = System.Text.Encoding.UTF8.GetBytes(SecretKey)
            Dim MS As New System.IO.MemoryStream()
            Dim CS As New System.Security.Cryptography.CryptoStream(MS, DES.CreateDecryptor, System.Security.Cryptography.CryptoStreamMode.Write)
            CS.Write(Source, 0, Source.Length)
            Try
                CS.FlushFinalBlock()
            Catch ex As System.Security.Cryptography.CryptographicException '解密失败（通常是由于SecretKey错误）
                Return ""
            End Try
            Return System.Text.Encoding.UTF8.GetString(MS.ToArray())
        End Function
        ''' <summary>
        ''' RSA加密算法
        ''' </summary>
        ''' <param name="Source">要加密的字符串</param>
        ''' <returns>加密后的结果字符串</returns>
        ''' <remarks></remarks>
        Public Shared Function RSA_Encode(ByVal Source As String) As Byte()
            Dim RSA As New System.Security.Cryptography.RSACryptoServiceProvider
            RSA.FromXmlString("<RSAKeyValue><Modulus>pZGIiC3CxVYpTJ4dLylSy2TLXW+R9EyRZ39ekSosvRKf7iPuz4oPlHqjssh4Glbj/vTUIMFzHFC/9zC56GggNLfZBjh6fc3adq5cXGKlU74kAyM2z7gdYlUHtLT/GwDp4YcQKeSb9GjcvsXbUp0mrzI/axzueLIqK+R07rnv3yc=</Modulus><Exponent>AQAB</Exponent></RSAKeyValue>")
            Return RSA.Encrypt(System.Text.Encoding.UTF8.GetBytes(Source), True)
        End Function
        ''' <summary>
        ''' RSA解密算法
        ''' </summary>
        ''' <param name="Source">要解密的字符串</param>
        ''' <returns>解密后的结果字符串</returns>
        ''' <remarks></remarks>
        Public Shared Function RSA_Decode(ByVal Source As Byte()) As String
            Dim RSA As New System.Security.Cryptography.RSACryptoServiceProvider
            RSA.FromXmlString("<RSAKeyValue><Modulus>pZGIiC3CxVYpTJ4dLylSy2TLXW+R9EyRZ39ekSosvRKf7iPuz4oPlHqjssh4Glbj/vTUIMFzHFC/9zC56GggNLfZBjh6fc3adq5cXGKlU74kAyM2z7gdYlUHtLT/GwDp4YcQKeSb9GjcvsXbUp0mrzI/axzueLIqK+R07rnv3yc=</Modulus><Exponent>AQAB</Exponent><P>0wCnxVUMgu+Uqp3UJ18bp9Ahdad36wDMwa0tmHxZJUvBZEfcYpsxmSHLpTUBCcAIg2eJL5g/iK9LrIwDBvUZ+w==</P><Q>yOB6ZwG9TuXMRPCA9cFTKCoHEsreDZluptHEfG3HvnS1lp5xwRCHXVuh7VWOM0G2gnZ/JWwWIfcqf30UTWvTxQ==</Q><DP>BTc67nHPwVzSu/TyzZZYRKmsahAdsr1uUktJmT9ZpMZenW/5Tqavby2arxbEU81faIAir/5/c42BvV4opP9iCQ==</DP><DQ>QETR5LMBxoRvXn80Q2yfFnKb4L9XXDKC3IywuL7G8YCVuKLo8kQ/ivcOT8jXvj6ADi2rcGWsjyFtT2zNWhftoQ==</DQ><InverseQ>jwpY6fpkzwtLOABZQncXMC4h7VbYrx+sZeSrBFXAgw1WMSs9YsT6EQcDRjpGt7JAkP14nSTSIVJNd23jZURCLw==</InverseQ><D>cw6SqcfbLVV198d9EnQOFEgkRvcsn2/CMAFET27WjkHuIAiagWE4+H7NWYWUaQFvCZNMAsNMYiX/cSFMYCRUFBBgkPqaqQ3+3qCs/kKiWpDjRwX8eXrMAnWniFDEoxc229Mxl4QZrcYKVRxrCIq8wKamuoWgwN0M+3CAiLwLvNk=</D></RSAKeyValue>")
            Try
                Return System.Text.Encoding.UTF8.GetString(RSA.Decrypt(Source, True))
            Catch ex As Exception
                Return ""
            End Try
        End Function
    End Class

    ''' <summary>
    ''' 网络访问函数
    ''' </summary>
    ''' <remarks></remarks>
    Public NotInheritable Class Http
        ''' <summary>
        ''' 获取网页源码
        ''' </summary>
        ''' <param name="URL">网页链接</param>
        ''' <returns>网页源码字符串</returns>
        ''' <remarks></remarks>
        Public Shared Function GetString(ByVal URL As String) As String
            Dim Client As System.Net.WebClient = New System.Net.WebClient
            Client.Encoding = System.Text.Encoding.UTF8
            Try
                Return Client.DownloadString(URL)
            Catch ex As Exception
                Return ""
            End Try
        End Function
        ''' <summary>
        ''' 获取网页源码
        ''' </summary>
        ''' <param name="URL">网页链接</param>
        ''' <param name="Encoding">使用的编码格式（默认UTF-8）</param>
        ''' <returns>网页源码字符串</returns>
        ''' <remarks></remarks>
        Public Shared Function GetString(ByVal URL As String, ByVal Encoding As System.Text.Encoding) As String
            Dim Client As System.Net.WebClient = New System.Net.WebClient
            Client.Encoding = Encoding
            Try
                Return Client.DownloadString(URL)
            Catch ex As Exception
                Return ""
            End Try
        End Function
        ''' <summary>
        ''' 下载文件到磁盘
        ''' </summary>
        ''' <param name="URL">文件链接</param>
        ''' <param name="FilePath">文件保存路径（可以是相对路径）</param>
        ''' <returns>是否下载成功</returns>
        ''' <remarks></remarks>
        Public Shared Function DownloadFile(ByVal URL As String, ByVal FilePath As String) As Boolean
            Dim Client As System.Net.WebClient = New System.Net.WebClient
            Try
                Client.DownloadFile(New System.Uri(URL), FilePath)
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function
    End Class

    ''' <summary>
    ''' 磁盘读写函数
    ''' </summary>
    ''' <remarks></remarks>
    Public NotInheritable Class IO
        ''' <summary>
        ''' 将字符串数组写入文件
        ''' </summary>
        ''' <param name="StringArray">字符串数组</param>
        ''' <param name="FilePath">文件路径（可以是相对路径）</param>
        ''' <returns>是否写入成功</returns>
        ''' <remarks></remarks>
        Public Shared Function SaveStringArray(ByRef StringArray As ArrayList, ByVal FilePath As String) As Boolean
            Dim Builder As System.Text.StringBuilder
            Dim Writer As System.IO.StreamWriter
            Try
                Builder = New System.Text.StringBuilder()
                For I = 0 To StringArray.Count - 2
                    Builder.Append(StringArray(I).ToString() & vbCrLf)
                Next
                If StringArray.Count > 0 Then Builder.Append(StringArray(StringArray.Count - 1))
                Writer = New System.IO.StreamWriter(FilePath, False, System.Text.Encoding.UTF8)
                Writer.Write(Builder)
                Writer.Dispose()
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function
        ''' <summary>
        ''' 将字符串数组写入文件
        ''' </summary>
        ''' <param name="StringArray">字符串数组</param>
        ''' <param name="FilePath">文件路径（可以是相对路径）</param>
        ''' <returns>是否写入成功</returns>
        ''' <remarks></remarks>
        Public Shared Function SaveStringArray(ByRef StringArray As List(Of String), ByVal FilePath As String) As Boolean
            Dim Builder As System.Text.StringBuilder
            Dim Writer As System.IO.StreamWriter
            Try
                Builder = New System.Text.StringBuilder()
                For I = 0 To StringArray.Count - 2
                    Builder.Append(StringArray(I) & vbCrLf)
                Next
                If StringArray.Count > 0 Then Builder.Append(StringArray(StringArray.Count - 1))
                Writer = New System.IO.StreamWriter(FilePath, False, System.Text.Encoding.UTF8)
                Writer.Write(Builder)
                Writer.Dispose()
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function
        ''' <summary>
        ''' 将字符串数组写入文件
        ''' </summary>
        ''' <param name="StringArray">字符串数组</param>
        ''' <param name="FilePath">文件路径（可以是相对路径）</param>
        ''' <returns>是否写入成功</returns>
        ''' <remarks></remarks>
        Public Shared Function SaveStringArray(ByRef StringArray As String(), ByVal FilePath As String) As Boolean
            Dim Builder As System.Text.StringBuilder
            Dim Writer As System.IO.StreamWriter
            Try
                Builder = New System.Text.StringBuilder()
                For I = 0 To StringArray.Length - 2
                    Builder.Append(StringArray(I) & vbCrLf)
                Next
                If StringArray.Length > 0 Then Builder.Append(StringArray(StringArray.Length - 1))
                Writer = New System.IO.StreamWriter(FilePath, False, System.Text.Encoding.UTF8)
                Writer.Write(Builder)
                Writer.Dispose()
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function
        ''' <summary>
        ''' 读取文件中的字符串数组
        ''' </summary>
        ''' <param name="StringArray">字符串数组</param>
        ''' <param name="FilePath">文件路径（可以是相对路径）</param>
        ''' <returns>是否读取成功</returns>
        ''' <remarks></remarks>
        Public Shared Function ReadStringArray(ByRef StringArray As ArrayList, ByVal FilePath As String) As Boolean
            Dim Reader As System.IO.StreamReader
            Try
                Reader = New System.IO.StreamReader(FilePath, System.Text.Encoding.UTF8)
                StringArray = New ArrayList(Reader.ReadToEnd().Replace(Chr(13) + Chr(10), Chr(10)).Split(New Char() {Chr(10)}))
                Reader.Dispose()
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function
        ''' <summary>
        ''' 读取文件中的字符串数组
        ''' </summary>
        ''' <param name="StringArray">字符串数组</param>
        ''' <param name="FilePath">文件路径（可以是相对路径）</param>
        ''' <returns>是否读取成功</returns>
        ''' <remarks></remarks>
        Public Shared Function ReadStringArray(ByRef StringArray As List(Of String), ByVal FilePath As String) As Boolean
            Dim Reader As System.IO.StreamReader
            Try
                Reader = New System.IO.StreamReader(FilePath, System.Text.Encoding.UTF8)
                StringArray = New List(Of String)(Reader.ReadToEnd().Replace(Chr(13) + Chr(10), Chr(10)).Split(New Char() {Chr(10)}))
                Reader.Dispose()
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function
        ''' <summary>
        ''' 读取文件中的字符串数组
        ''' </summary>
        ''' <param name="StringArray">字符串数组</param>
        ''' <param name="FilePath">文件路径（可以是相对路径）</param>
        ''' <returns>是否读取成功</returns>
        ''' <remarks></remarks>
        Public Shared Function ReadStringArray(ByRef StringArray As String(), ByVal FilePath As String) As Boolean
            Dim Reader As System.IO.StreamReader
            Try
                Reader = New System.IO.StreamReader(FilePath, System.Text.Encoding.UTF8)
                StringArray = Reader.ReadToEnd().Replace(Chr(13) + Chr(10), Chr(10)).Split(New Char() {Chr(10)})
                Reader.Dispose()
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function
        ''' <summary>
        ''' 获取指定目录下的全部的文件
        ''' </summary>
        ''' <param name="FilePathArray">保存文件路径的字符串数组</param>
        ''' <param name="SearchDirectory">要搜索的文件夹路径</param>
        ''' <returns>是否获取成功</returns>
        ''' <remarks></remarks>
        Public Shared Function GetAllFilePath(ByRef FilePathArray As List(Of String), ByVal SearchDirectory As String) As Boolean
            Dim Directory As List(Of String) = New List(Of String)
            Try
                For Each Temp As String In System.IO.Directory.GetFiles(SearchDirectory)
                    FilePathArray.Add(Temp)
                Next
                For Each Temp As String In System.IO.Directory.GetDirectories(SearchDirectory)
                    Directory.Add(Temp)
                Next
                Dim Index As Integer = 0
                While Directory.Count > Index
                    SearchDirectory = Directory(Index)
                    Index = Index + 1
                    For Each Temp As String In System.IO.Directory.GetFiles(SearchDirectory)
                        FilePathArray.Add(Temp)
                    Next
                    For Each Temp As String In System.IO.Directory.GetDirectories(SearchDirectory)
                        Directory.Add(Temp)
                    Next
                End While
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function
        ''' <summary>
        ''' 获取指定文件的行数
        ''' </summary>
        ''' <param name="FilePath">文件路径（可以是相对路径）</param>
        ''' <returns>行数（获取失败时返回0）</returns>
        ''' <remarks></remarks>
        Public Shared Function GetFileLine(ByVal FilePath As String) As Integer
            Dim Reader As System.IO.StreamReader
            Try
                Reader = New System.IO.StreamReader(FilePath, System.Text.Encoding.UTF8)
                Return Reader.ReadToEnd().Replace(Chr(13) + Chr(10), Chr(10)).Split(New Char() {Chr(10)}).Length
            Catch ex As Exception
                Return 0
            End Try
        End Function
        ''' <summary>
        ''' 获取指定文件的行数
        ''' </summary>
        ''' <param name="FilePath">文件路径（可以是相对路径）</param>
        ''' <param name="Encoding">使用的编码格式（默认UTF-8）</param>
        ''' <returns>行数（获取失败时返回0）</returns>
        ''' <remarks></remarks>
        Public Shared Function GetFileLine(ByVal FilePath As String, ByVal Encoding As System.Text.Encoding) As Integer
            Dim Reader As System.IO.StreamReader
            Try
                Reader = New System.IO.StreamReader(FilePath, Encoding)
                Return Reader.ReadToEnd().Replace(Chr(13) + Chr(10), Chr(10)).Split(New Char() {Chr(10)}).Length
            Catch ex As Exception
                Return 0
            End Try
        End Function
    End Class

    ''' <summary>
    ''' 字符串处理函数
    ''' </summary>
    ''' <remarks></remarks>
    Public NotInheritable Class StringData
        ''' <summary>
        ''' 搜索字符串（从第一个开始字符串的位置，向后搜寻结束字符串，取出中间的部分）
        ''' </summary>
        ''' <param name="SourceCode">要搜索的字符串</param>
        ''' <param name="BeginString">开始字符串</param>
        ''' <param name="EndString">结束字符串</param>
        ''' <returns>搜索结果字符串（无结果时返回空字符串）</returns>
        ''' <remarks></remarks>
        Public Shared Function SearchForward(ByRef SourceCode As String, ByVal BeginString As String, ByVal EndString As String) As String
            If SourceCode.Contains(BeginString) = False Then Return ""
            SourceCode = SourceCode.Substring(SourceCode.IndexOf(BeginString) + BeginString.Length)
            If SourceCode.Contains(EndString) = False Then Return ""
            Return SourceCode.Substring(0, SourceCode.IndexOf(EndString))
        End Function
        ''' <summary>
        ''' 搜索字符串（从最后一个结束字符串的位置，向前搜寻开始字符串，取出中间的部分）
        ''' </summary>
        ''' <param name="SourceCode">要搜索的字符串</param>
        ''' <param name="BeginString">开始字符串</param>
        ''' <param name="EndString">结束字符串</param>
        ''' <returns>搜索结果字符串（无结果时返回空字符串）</returns>
        ''' <remarks></remarks>
        Public Shared Function SearchBackward(ByRef SourceCode As String, ByVal BeginString As String, ByVal EndString As String) As String
            If SourceCode.Contains(EndString) = False Then Return ""
            SourceCode = SourceCode.Substring(0, SourceCode.LastIndexOf(EndString))
            If SourceCode.Contains(BeginString) = False Then Return ""
            Return SourceCode.Substring(SourceCode.LastIndexOf(BeginString) + BeginString.Length)
        End Function
        ''' <summary>
        ''' 搜索字符串（从第一个开始字符串的位置，向后搜寻最后一个结束字符串的位置，取出中间的部分）
        ''' </summary>
        ''' <param name="SourceCode">要搜索的字符串</param>
        ''' <param name="BeginString">开始字符串</param>
        ''' <param name="EndString">结束字符串</param>
        ''' <returns>搜索结果字符串（无结果时返回空字符串）</returns>
        ''' <remarks></remarks>
        Public Shared Function SearchMiddle(ByRef SourceCode As String, ByVal BeginString As String, ByVal EndString As String) As String
            If SourceCode.Contains(BeginString) = False Then Return ""
            SourceCode = SourceCode.Substring(SourceCode.IndexOf(BeginString) + BeginString.Length)
            If SourceCode.Contains(EndString) = False Then Return ""
            Return SourceCode.Substring(0, SourceCode.LastIndexOf(EndString))
        End Function
        ''' <summary>
        ''' 搜索字符串（从第一个开始字符串的位置，向后搜寻结束字符串，取出中间的部分，重复向后搜索，返回数组）
        ''' </summary>
        ''' <param name="SourceCode">要搜索的字符串</param>
        ''' <param name="BeginString">开始字符串</param>
        ''' <param name="EndString">结束字符串</param>
        ''' <returns>搜索结果字符串数组（无结果时返回空数组）</returns>
        ''' <remarks></remarks>
        Public Shared Function SearchAllForward(ByRef SourceCode As String, ByVal BeginString As String, ByVal EndString As String) As List(Of String)
            Dim Temp = New List(Of String)()
            While SourceCode.Contains(BeginString)
                SourceCode = SourceCode.Substring(SourceCode.IndexOf(BeginString) + BeginString.Length)
                If SourceCode.Contains(EndString) = False Then Exit While
                Temp.Add(SourceCode.Substring(0, SourceCode.IndexOf(EndString)))
            End While
            Return Temp
        End Function
        ''' <summary>
        ''' 搜索字符串（从最后一个结束字符串的位置，向前搜寻开始字符串，取出中间的部分，重复向前搜索，返回数组）
        ''' </summary>
        ''' <param name="SourceCode">要搜索的字符串</param>
        ''' <param name="BeginString">开始字符串</param>
        ''' <param name="EndString">结束字符串</param>
        ''' <returns>搜索结果字符串数组（无结果时返回空数组）</returns>
        ''' <remarks></remarks>
        Public Shared Function SearchAllBackward(ByRef SourceCode As String, ByVal BeginString As String, ByVal EndString As String) As List(Of String)
            Dim Temp = New List(Of String)()
            While SourceCode.Contains(EndString)
                SourceCode = SourceCode.Substring(0, SourceCode.LastIndexOf(EndString))
                If SourceCode.Contains(BeginString) = False Then Exit While
                Temp.Add(SourceCode.Substring(SourceCode.LastIndexOf(BeginString) + BeginString.Length))
            End While
            Return Temp
        End Function
        ''' <summary>
        ''' 搜索字符串（从第一个开始字符串的位置，向后搜寻最后一个结束字符串的位置，取出中间的部分，重复向中间搜索，返回数组）
        ''' </summary>
        ''' <param name="SourceCode">要搜索的字符串</param>
        ''' <param name="BeginString">开始字符串</param>
        ''' <param name="EndString">结束字符串</param>
        ''' <returns>搜索结果字符串数组（无结果时返回空数组）</returns>
        ''' <remarks></remarks>
        Public Shared Function SearchAllMiddle(ByRef SourceCode As String, ByVal BeginString As String, ByVal EndString As String) As List(Of String)
            Dim Temp = New List(Of String)()
            While SourceCode.Contains(BeginString)
                SourceCode = SourceCode.Substring(SourceCode.IndexOf(BeginString) + BeginString.Length)
                If SourceCode.Contains(EndString) = False Then Exit While
                Temp.Add(SourceCode.Substring(0, SourceCode.LastIndexOf(EndString)))
            End While
            Return Temp
        End Function
    End Class

    ''' <summary>
    ''' HTML代码处理函数
    ''' </summary>
    ''' <remarks></remarks>
    Public NotInheritable Class Html
        ''' <summary>
        ''' 获取网页源码中指定标签的元素的文本
        ''' </summary>
        ''' <param name="Source">网页源代码</param>
        ''' <param name="HtmlTag">元素标签</param>
        ''' <returns>文本字符串数组</returns>
        ''' <remarks></remarks>
        Public Shared Function GetTextByTagName(ByVal Source As String, ByVal HtmlTag As String) As List(Of String)
            Dim Temp As List(Of String) = New List(Of String)
            Dim Browser As WebBrowser = New WebBrowser
            Browser.DocumentText = ""
            Browser.Document.OpenNew(True).Write(Source)
            Dim Elements As HtmlElementCollection = Browser.Document.GetElementsByTagName(HtmlTag)
            For I = 0 To Elements.Count - 1
                Temp.Add(Elements(I).InnerText)
            Next
            Return Temp
        End Function
        ''' <summary>
        ''' 获取网页源码中指定ID的元素的文本
        ''' </summary>
        ''' <param name="Source">网页源代码</param>
        ''' <param name="Id">元素ID</param>
        ''' <returns>文本字符串</returns>
        ''' <remarks></remarks>
        Public Shared Function GetTextById(ByVal Source As String, ByVal Id As String) As String
            Dim Browser As WebBrowser = New WebBrowser
            Browser.DocumentText = ""
            Browser.Document.OpenNew(True).Write(Source)
            Return Browser.Document.GetElementById(Id).InnerText
        End Function
    End Class


    '扩展已有的命名空间和类库
    ''' <summary>
    ''' My.Application
    ''' </summary>
    ''' <remarks>访问应用程序信息和服务。</remarks>
    Partial Class MyApplication
    End Class

    ''' <summary>
    ''' My.Computer
    ''' </summary>
    ''' <remarks>访问主机及其资源、服务和数据。</remarks>
    Partial Class MyComputer
    End Class

End Namespace
