Namespace My
    '2015/6/2
    'My命名空间的扩展
    '引入此文件，需要确保该工程引用System，System.Drawing，System.Web，System.Windows.Forms

    ''' <summary>
    ''' 编码与解码函数
    ''' </summary>
    ''' <remarks></remarks>
    Public NotInheritable Class Security
        ''' <summary>
        ''' URL编码（例如，"微软.NET"的编码结果为"%e5%be%ae%e8%bd%af.NET"）
        ''' </summary>
        ''' <param name="Source">要加密的字符串</param>
        ''' <param name="ToUpper">是否将结果转换为大写字母形式</param>
        ''' <returns>加密后的结果字符串</returns>
        ''' <remarks></remarks>
        Public Shared Function URL_Encode(ByVal Source As String, Optional ByVal ToUpper As Boolean = False) As String
            If ToUpper Then
                Return System.Web.HttpUtility.UrlEncode(Source, System.Text.Encoding.UTF8).ToUpper
            End If
            Return System.Web.HttpUtility.UrlEncode(Source, System.Text.Encoding.UTF8)
        End Function
        ''' <summary>
        ''' URL编码（例如，"微软.NET"的编码结果为"%e5%be%ae%e8%bd%af.NET"）
        ''' </summary>
        ''' <param name="Source">要加密的字符串</param>
        ''' <param name="Encoding">使用的编码格式（默认UTF-8）</param>
        ''' <param name="ToUpper">是否将结果转换为大写字母形式</param>
        ''' <returns>加密后的结果字符串</returns>
        ''' <remarks></remarks>
        Public Shared Function URL_Encode(ByVal Source As String, ByVal Encoding As System.Text.Encoding, Optional ByVal ToUpper As Boolean = False) As String
            If ToUpper Then
                Return System.Web.HttpUtility.UrlEncode(Source, Encoding).ToUpper
            End If
            Return System.Web.HttpUtility.UrlEncode(Source, Encoding)
        End Function
        ''' <summary>
        ''' URL解码（大小写形式都可以识别）
        ''' </summary>
        ''' <param name="Source">要解密的字符串</param>
        ''' <returns>解密后的结果字符串</returns>
        ''' <remarks></remarks>
        Public Shared Function URL_Decode(ByVal Source As String) As String
            Return System.Web.HttpUtility.UrlDecode(Source, System.Text.Encoding.UTF8)
        End Function
        ''' <summary>
        ''' URL解码（大小写形式都可以识别）
        ''' </summary>
        ''' <param name="Source">要解密的字符串</param>
        ''' <param name="Encoding">使用的编码格式（默认UTF-8）</param>
        ''' <returns>解密后的结果字符串</returns>
        ''' <remarks></remarks>
        Public Shared Function URL_Decode(ByVal Source As String, ByVal Encoding As System.Text.Encoding) As String
            Return System.Web.HttpUtility.UrlDecode(Source, Encoding)
        End Function
        ''' <summary>
        ''' Base64加密算法（加密结果的字符串中包含字母A-Z，a-z，数字0-9，符号+/=）
        ''' </summary>
        ''' <param name="Source">要加密的字符串</param>
        ''' <returns>加密后的结果字符串</returns>
        ''' <remarks></remarks>
        Public Shared Function Base64_Encode(ByVal Source As String) As String
            Return Convert.ToBase64String(System.Text.Encoding.UTF8.GetBytes(Source))
        End Function
        ''' <summary>
        ''' Base64加密算法（加密结果的字符串中包含字母A-Z，a-z，数字0-9，符号+/=）
        ''' </summary>
        ''' <param name="Source">要加密的字符串</param>
        ''' <param name="Encoding">使用的编码格式（默认UTF-8）</param>
        ''' <returns>加密后的结果字符串</returns>
        ''' <remarks></remarks>
        Public Shared Function Base64_Encode(ByVal Source As String, ByVal Encoding As System.Text.Encoding) As String
            Return Convert.ToBase64String(Encoding.GetBytes(Source))
        End Function
        ''' <summary>
        ''' Base64加密算法（用于URL的改进Base64编码，加密结果的字符串中包含字母A-Z，a-z，数字0-9，符号-_=）
        ''' </summary>
        ''' <param name="Source">要加密的字符串</param>
        ''' <returns>加密后的结果字符串</returns>
        ''' <remarks></remarks>
        Public Shared Function Base64_URL_Encode(ByVal Source As String) As String
            Return Convert.ToBase64String(System.Text.Encoding.UTF8.GetBytes(Source)).Replace("+", "-").Replace("/", "_")
        End Function
        ''' <summary>
        ''' Base64加密算法（用于URL的改进Base64编码，加密结果的字符串中包含字母A-Z，a-z，数字0-9，符号-_=）
        ''' </summary>
        ''' <param name="Source">要加密的字符串</param>
        ''' <param name="Encoding">使用的编码格式（默认UTF-8）</param>
        ''' <returns>加密后的结果字符串</returns>
        ''' <remarks></remarks>
        Public Shared Function Base64_URL_Encode(ByVal Source As String, ByVal Encoding As System.Text.Encoding) As String
            Return Convert.ToBase64String(Encoding.GetBytes(Source)).Replace("+", "-").Replace("/", "_")
        End Function
        ''' <summary>
        ''' Base64解密算法（要解密的字符串可以包含字母A-Z，a-z，数字0-9，符号+/=）
        ''' </summary>
        ''' <param name="Source">要解密的字符串</param>
        ''' <returns>解密后的结果字符串</returns>
        ''' <remarks></remarks>
        Public Shared Function Base64_Decode(ByVal Source As String) As String
            Return System.Text.Encoding.UTF8.GetString(Convert.FromBase64String(Source))
        End Function
        ''' <summary>
        ''' Base64解密算法（要解密的字符串可以包含字母A-Z，a-z，数字0-9，符号+/=）
        ''' </summary>
        ''' <param name="Source">要解密的字符串</param>
        ''' <param name="Encoding">使用的编码格式（默认UTF-8）</param>
        ''' <returns>解密后的结果字符串</returns>
        ''' <remarks></remarks>
        Public Shared Function Base64_Decode(ByVal Source As String, ByVal Encoding As System.Text.Encoding) As String
            Return Encoding.GetString(Convert.FromBase64String(Source))
        End Function
        ''' <summary>
        ''' Base64解密算法（用于URL的改进Base64解码，要解密的字符串可以包含字母A-Z，a-z，数字0-9，符号-_=）
        ''' </summary>
        ''' <param name="Source">要解密的字符串</param>
        ''' <returns>解密后的结果字符串</returns>
        ''' <remarks></remarks>
        Public Shared Function Base64_URL_Decode(ByVal Source As String) As String
            Return System.Text.Encoding.UTF8.GetString(Convert.FromBase64String(Source.Replace("-", "+").Replace("_", "/")))
        End Function
        ''' <summary>
        ''' Base64解密算法（用于URL的改进Base64解码，要解密的字符串可以包含字母A-Z，a-z，数字0-9，符号-_=）
        ''' </summary>
        ''' <param name="Source">要解密的字符串</param>
        ''' <param name="Encoding">使用的编码格式（默认UTF-8）</param>
        ''' <returns>解密后的结果字符串</returns>
        ''' <remarks></remarks>
        Public Shared Function Base64_URL_Decode(ByVal Source As String, ByVal Encoding As System.Text.Encoding) As String
            Return Encoding.GetString(Convert.FromBase64String(Source.Replace("-", "+").Replace("_", "/")))
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
        ''' <param name="Encoding">使用的编码格式（默认UTF-8）</param>
        ''' <returns>是否写入成功</returns>
        ''' <remarks></remarks>
        Public Shared Function SaveStringArray(ByRef StringArray As ArrayList, ByVal FilePath As String, ByVal Encoding As System.Text.Encoding) As Boolean
            Dim Builder As System.Text.StringBuilder
            Dim Writer As System.IO.StreamWriter
            Try
                Builder = New System.Text.StringBuilder()
                For I = 0 To StringArray.Count - 2
                    Builder.Append(StringArray(I).ToString() & vbCrLf)
                Next
                If StringArray.Count > 0 Then Builder.Append(StringArray(StringArray.Count - 1))
                Writer = New System.IO.StreamWriter(FilePath, False, Encoding)
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
        ''' <param name="Encoding">使用的编码格式（默认UTF-8）</param>
        ''' <returns>是否写入成功</returns>
        ''' <remarks></remarks>
        Public Shared Function SaveStringArray(ByRef StringArray As List(Of String), ByVal FilePath As String, ByVal Encoding As System.Text.Encoding) As Boolean
            Dim Builder As System.Text.StringBuilder
            Dim Writer As System.IO.StreamWriter
            Try
                Builder = New System.Text.StringBuilder()
                For I = 0 To StringArray.Count - 2
                    Builder.Append(StringArray(I) & vbCrLf)
                Next
                If StringArray.Count > 0 Then Builder.Append(StringArray(StringArray.Count - 1))
                Writer = New System.IO.StreamWriter(FilePath, False, Encoding)
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
        ''' 将字符串数组写入文件
        ''' </summary>
        ''' <param name="StringArray">字符串数组</param>
        ''' <param name="FilePath">文件路径（可以是相对路径）</param>
        ''' <param name="Encoding">使用的编码格式（默认UTF-8）</param>
        ''' <returns>是否写入成功</returns>
        ''' <remarks></remarks>
        Public Shared Function SaveStringArray(ByRef StringArray As String(), ByVal FilePath As String, ByVal Encoding As System.Text.Encoding) As Boolean
            Dim Builder As System.Text.StringBuilder
            Dim Writer As System.IO.StreamWriter
            Try
                Builder = New System.Text.StringBuilder()
                For I = 0 To StringArray.Length - 2
                    Builder.Append(StringArray(I) & vbCrLf)
                Next
                If StringArray.Length > 0 Then Builder.Append(StringArray(StringArray.Length - 1))
                Writer = New System.IO.StreamWriter(FilePath, False, Encoding)
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
        ''' <param name="Encoding">使用的编码格式（默认UTF-8）</param>
        ''' <returns>是否读取成功</returns>
        ''' <remarks></remarks>
        Public Shared Function ReadStringArray(ByRef StringArray As ArrayList, ByVal FilePath As String, ByVal Encoding As System.Text.Encoding) As Boolean
            Dim Reader As System.IO.StreamReader
            Try
                Reader = New System.IO.StreamReader(FilePath, Encoding)
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
        ''' <param name="Encoding">使用的编码格式（默认UTF-8）</param>
        ''' <returns>是否读取成功</returns>
        ''' <remarks></remarks>
        Public Shared Function ReadStringArray(ByRef StringArray As List(Of String), ByVal FilePath As String, ByVal Encoding As System.Text.Encoding) As Boolean
            Dim Reader As System.IO.StreamReader
            Try
                Reader = New System.IO.StreamReader(FilePath, Encoding)
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
        ''' 读取文件中的字符串数组
        ''' </summary>
        ''' <param name="StringArray">字符串数组</param>
        ''' <param name="FilePath">文件路径（可以是相对路径）</param>
        ''' <param name="Encoding">使用的编码格式（默认UTF-8）</param>
        ''' <returns>是否读取成功</returns>
        ''' <remarks></remarks>
        Public Shared Function ReadStringArray(ByRef StringArray As String(), ByVal FilePath As String, ByVal Encoding As System.Text.Encoding) As Boolean
            Dim Reader As System.IO.StreamReader
            Try
                Reader = New System.IO.StreamReader(FilePath, Encoding)
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
        ''' <param name="FileLine">文件行数</param>
        ''' <param name="FilePath">文件路径（可以是相对路径）</param>
        ''' <returns>是否获取成功</returns>
        ''' <remarks></remarks>
        Public Shared Function GetFileLine(ByRef FileLine As Integer, ByVal FilePath As String) As Boolean
            Dim Reader As System.IO.StreamReader
            Try
                Reader = New System.IO.StreamReader(FilePath, System.Text.Encoding.UTF8)
                FileLine = Reader.ReadToEnd().Replace(Chr(13) + Chr(10), Chr(10)).Split(New Char() {Chr(10)}).Length
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function
        ''' <summary>
        ''' 获取指定文件的行数
        ''' </summary>
        ''' <param name="FileLine">文件行数</param>
        ''' <param name="FilePath">文件路径（可以是相对路径）</param>
        ''' <param name="Encoding">使用的编码格式（默认UTF-8）</param>
        ''' <returns>是否获取成功</returns>
        ''' <remarks></remarks>
        Public Shared Function GetFileLine(ByRef FileLine As Integer, ByVal FilePath As String, ByVal Encoding As System.Text.Encoding) As Boolean
            Dim Reader As System.IO.StreamReader
            Try
                Reader = New System.IO.StreamReader(FilePath, Encoding)
                FileLine = Reader.ReadToEnd().Replace(Chr(13) + Chr(10), Chr(10)).Split(New Char() {Chr(10)}).Length
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function
        ''' <summary>
        ''' 将程序嵌入的资源文件写入为磁盘文件（注意必须在解决方案资源管理器中，将图片资源的"属性"-"生成操作"设置为"嵌入的资源"）
        ''' </summary>
        ''' <param name="ResourceName">资源文件名称（注意这个参数的值为资源文件在工程内的相对路径，例如A.jpg文件在Resources文件夹内，则此处的参数应为"Resources.A.jpg"）</param>
        ''' <param name="FilePath">写入到磁盘的文件路径</param>
        ''' <returns>是否执行成功</returns>
        ''' <remarks></remarks>
        Public Shared Function WriteResourceFile(ByVal ResourceName As String, ByVal FilePath As String) As Boolean
            Dim Reader As System.IO.Stream
            Dim Writer As System.IO.FileStream
            Try
                Dim MyAssembly As System.Reflection.Assembly = System.Reflection.Assembly.GetExecutingAssembly()
                Reader = MyAssembly.GetManifestResourceStream(MyAssembly.GetName().Name & "." & ResourceName)
                Dim I(CInt(Reader.Length - 1)) As Byte
                Writer = New System.IO.FileStream(FilePath, System.IO.FileMode.OpenOrCreate, System.IO.FileAccess.ReadWrite)
                Reader.Read(I, 0, I.Length)
                Writer.Write(I, 0, I.Length)
                Reader.Dispose()
                Writer.Dispose()
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function
        ''' <summary>
        ''' 读取程序嵌入的图片类型的资源文件（注意必须在解决方案资源管理器中，将资源文件的"属性"-"生成操作"设置为"嵌入的资源"）
        ''' </summary>
        ''' <param name="ResourcePicture">图片</param>
        ''' <param name="ResourceName">资源文件名称（注意这个参数的值为资源文件在工程内的相对路径，例如A.jpg文件在Resources文件夹内，则此处的参数应为"Resources.A.jpg"）</param>
        ''' <returns>是否读取成功</returns>
        ''' <remarks></remarks>
        Public Shared Function ReadResourcePicture(ByRef ResourcePicture As Bitmap, ByVal ResourceName As String) As Boolean
            Try
                Dim MyAssembly As System.Reflection.Assembly = System.Reflection.Assembly.GetExecutingAssembly()
                ResourcePicture = New Bitmap(MyAssembly.GetManifestResourceStream(MyAssembly.GetName().Name & "." & ResourceName))
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function
        ''' <summary>
        ''' 读取程序嵌入的文本类型的资源文件（注意必须在解决方案资源管理器中，将资源文件的"属性"-"生成操作"设置为"嵌入的资源"）
        ''' </summary>
        ''' <param name="ResourceString">文本</param>
        ''' <param name="ResourceName">资源文件名称（注意这个参数的值为资源文件在工程内的相对路径，例如A.jpg文件在Resources文件夹内，则此处的参数应为"Resources.A.jpg"）</param>
        ''' <returns>是否读取成功</returns>
        ''' <remarks></remarks>
        Public Shared Function ReadResourceString(ByRef ResourceString As String, ByVal ResourceName As String) As Boolean
            Dim Reader As System.IO.Stream
            Try
                Dim MyAssembly As System.Reflection.Assembly = System.Reflection.Assembly.GetExecutingAssembly()
                Reader = MyAssembly.GetManifestResourceStream(MyAssembly.GetName().Name & "." & ResourceName)
                Dim I(CInt(Reader.Length - 1)) As Byte
                Reader.Read(I, 0, I.Length)
                Reader.Dispose()
                ResourceString = System.Text.Encoding.UTF8.GetString(I)
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function
        ''' <summary>
        ''' 读取程序嵌入的文本类型的资源文件（注意必须在解决方案资源管理器中，将资源文件的"属性"-"生成操作"设置为"嵌入的资源"）
        ''' </summary>
        ''' <param name="ResourceString">文本</param>
        ''' <param name="ResourceName">资源文件名称（注意这个参数的值为资源文件在工程内的相对路径，例如A.jpg文件在Resources文件夹内，则此处的参数应为"Resources.A.jpg"）</param>
        ''' <param name="Encoding">使用的编码格式（默认UTF-8）</param>
        ''' <returns>是否读取成功</returns>
        ''' <remarks></remarks>
        Public Shared Function ReadResourceString(ByRef ResourceString As String, ByVal ResourceName As String, ByVal Encoding As System.Text.Encoding) As Boolean
            Dim Reader As System.IO.Stream
            Try
                Dim MyAssembly As System.Reflection.Assembly = System.Reflection.Assembly.GetExecutingAssembly()
                Reader = MyAssembly.GetManifestResourceStream(MyAssembly.GetName().Name & "." & ResourceName)
                Dim I(CInt(Reader.Length - 1)) As Byte
                Reader.Read(I, 0, I.Length)
                Reader.Dispose()
                ResourceString = Encoding.GetString(I)
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function
        ''' <summary>
        ''' 读取程序嵌入的字符串数组类型的资源文件（注意必须在解决方案资源管理器中，将资源文件的"属性"-"生成操作"设置为"嵌入的资源"）
        ''' </summary>
        ''' <param name="ResourceArray">字符串数组</param>
        ''' <param name="ResourceName">资源文件名称（注意这个参数的值为资源文件在工程内的相对路径，例如A.jpg文件在Resources文件夹内，则此处的参数应为"Resources.A.jpg"）</param>
        ''' <returns>是否读取成功</returns>
        ''' <remarks></remarks>
        Public Shared Function ReadResourceArray(ByRef ResourceArray As ArrayList, ByVal ResourceName As String) As Boolean
            Dim Reader As System.IO.Stream
            Try
                Dim MyAssembly As System.Reflection.Assembly = System.Reflection.Assembly.GetExecutingAssembly()
                Reader = MyAssembly.GetManifestResourceStream(MyAssembly.GetName().Name & "." & ResourceName)
                Dim I(CInt(Reader.Length - 1)) As Byte
                Reader.Read(I, 0, I.Length)
                Reader.Dispose()
                ResourceArray = New ArrayList(System.Text.Encoding.UTF8.GetString(I).Replace(Chr(13) + Chr(10), Chr(10)).Split(New Char() {Chr(10)}))
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function
        ''' <summary>
        ''' 读取程序嵌入的字符串数组类型的资源文件（注意必须在解决方案资源管理器中，将资源文件的"属性"-"生成操作"设置为"嵌入的资源"）
        ''' </summary>
        ''' <param name="ResourceArray">字符串数组</param>
        ''' <param name="ResourceName">资源文件名称（注意这个参数的值为资源文件在工程内的相对路径，例如A.jpg文件在Resources文件夹内，则此处的参数应为"Resources.A.jpg"）</param>
        ''' <param name="Encoding">使用的编码格式（默认UTF-8）</param>
        ''' <returns>是否读取成功</returns>
        ''' <remarks></remarks>
        Public Shared Function ReadResourceArray(ByRef ResourceArray As ArrayList, ByVal ResourceName As String, ByVal Encoding As System.Text.Encoding) As Boolean
            Dim Reader As System.IO.Stream
            Try
                Dim MyAssembly As System.Reflection.Assembly = System.Reflection.Assembly.GetExecutingAssembly()
                Reader = MyAssembly.GetManifestResourceStream(MyAssembly.GetName().Name & "." & ResourceName)
                Dim I(CInt(Reader.Length - 1)) As Byte
                Reader.Read(I, 0, I.Length)
                Reader.Dispose()
                ResourceArray = New ArrayList(Encoding.GetString(I).Replace(Chr(13) + Chr(10), Chr(10)).Split(New Char() {Chr(10)}))
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function
        ''' <summary>
        ''' 读取程序嵌入的字符串数组类型的资源文件（注意必须在解决方案资源管理器中，将资源文件的"属性"-"生成操作"设置为"嵌入的资源"）
        ''' </summary>
        ''' <param name="ResourceArray">字符串数组</param>
        ''' <param name="ResourceName">资源文件名称（注意这个参数的值为资源文件在工程内的相对路径，例如A.jpg文件在Resources文件夹内，则此处的参数应为"Resources.A.jpg"）</param>
        ''' <returns>是否读取成功</returns>
        ''' <remarks></remarks>
        Public Shared Function ReadResourceArray(ByRef ResourceArray As List(Of String), ByVal ResourceName As String) As Boolean
            Dim Reader As System.IO.Stream
            Try
                Dim MyAssembly As System.Reflection.Assembly = System.Reflection.Assembly.GetExecutingAssembly()
                Reader = MyAssembly.GetManifestResourceStream(MyAssembly.GetName().Name & "." & ResourceName)
                Dim I(CInt(Reader.Length - 1)) As Byte
                Reader.Read(I, 0, I.Length)
                Reader.Dispose()
                ResourceArray = New List(Of String)(System.Text.Encoding.UTF8.GetString(I).Replace(Chr(13) + Chr(10), Chr(10)).Split(New Char() {Chr(10)}))
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function
        ''' <summary>
        ''' 读取程序嵌入的字符串数组类型的资源文件（注意必须在解决方案资源管理器中，将资源文件的"属性"-"生成操作"设置为"嵌入的资源"）
        ''' </summary>
        ''' <param name="ResourceArray">字符串数组</param>
        ''' <param name="ResourceName">资源文件名称（注意这个参数的值为资源文件在工程内的相对路径，例如A.jpg文件在Resources文件夹内，则此处的参数应为"Resources.A.jpg"）</param>
        ''' <param name="Encoding">使用的编码格式（默认UTF-8）</param>
        ''' <returns>是否读取成功</returns>
        ''' <remarks></remarks>
        Public Shared Function ReadResourceArray(ByRef ResourceArray As List(Of String), ByVal ResourceName As String, ByVal Encoding As System.Text.Encoding) As Boolean
            Dim Reader As System.IO.Stream
            Try
                Dim MyAssembly As System.Reflection.Assembly = System.Reflection.Assembly.GetExecutingAssembly()
                Reader = MyAssembly.GetManifestResourceStream(MyAssembly.GetName().Name & "." & ResourceName)
                Dim I(CInt(Reader.Length - 1)) As Byte
                Reader.Read(I, 0, I.Length)
                Reader.Dispose()
                ResourceArray = New List(Of String)(Encoding.GetString(I).Replace(Chr(13) + Chr(10), Chr(10)).Split(New Char() {Chr(10)}))
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function
        ''' <summary>
        ''' 读取程序嵌入的字符串数组类型的资源文件（注意必须在解决方案资源管理器中，将资源文件的"属性"-"生成操作"设置为"嵌入的资源"）
        ''' </summary>
        ''' <param name="ResourceArray">字符串数组</param>
        ''' <param name="ResourceName">资源文件名称（注意这个参数的值为资源文件在工程内的相对路径，例如A.jpg文件在Resources文件夹内，则此处的参数应为"Resources.A.jpg"）</param>
        ''' <returns>是否读取成功</returns>
        ''' <remarks></remarks>
        Public Shared Function ReadResourceArray(ByRef ResourceArray As String(), ByVal ResourceName As String) As Boolean
            Dim Reader As System.IO.Stream
            Try
                Dim MyAssembly As System.Reflection.Assembly = System.Reflection.Assembly.GetExecutingAssembly()
                Reader = MyAssembly.GetManifestResourceStream(MyAssembly.GetName().Name & "." & ResourceName)
                Dim I(CInt(Reader.Length - 1)) As Byte
                Reader.Read(I, 0, I.Length)
                Reader.Dispose()
                ResourceArray = System.Text.Encoding.UTF8.GetString(I).Replace(Chr(13) + Chr(10), Chr(10)).Split(New Char() {Chr(10)})
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function
        ''' <summary>
        ''' 读取程序嵌入的字符串数组类型的资源文件（注意必须在解决方案资源管理器中，将资源文件的"属性"-"生成操作"设置为"嵌入的资源"）
        ''' </summary>
        ''' <param name="ResourceArray">字符串数组</param>
        ''' <param name="ResourceName">资源文件名称（注意这个参数的值为资源文件在工程内的相对路径，例如A.jpg文件在Resources文件夹内，则此处的参数应为"Resources.A.jpg"）</param>
        ''' <param name="Encoding">使用的编码格式（默认UTF-8）</param>
        ''' <returns>是否读取成功</returns>
        ''' <remarks></remarks>
        Public Shared Function ReadResourceArray(ByRef ResourceArray As String(), ByVal ResourceName As String, ByVal Encoding As System.Text.Encoding) As Boolean
            Dim Reader As System.IO.Stream
            Try
                Dim MyAssembly As System.Reflection.Assembly = System.Reflection.Assembly.GetExecutingAssembly()
                Reader = MyAssembly.GetManifestResourceStream(MyAssembly.GetName().Name & "." & ResourceName)
                Dim I(CInt(Reader.Length - 1)) As Byte
                Reader.Read(I, 0, I.Length)
                Reader.Dispose()
                ResourceArray = Encoding.GetString(I).Replace(Chr(13) + Chr(10), Chr(10)).Split(New Char() {Chr(10)})
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function
        ''' <summary>
        ''' 创建快捷方式
        ''' </summary>
        ''' <param name="TargetPath">快捷方式指向的路径（建议使用绝对路径，使用相对路径时，默认将以桌面作为父目录）</param>
        ''' <param name="LinkFilePath">快捷方式文件的路径（可以是相对路径，如"1.lnk"）</param>
        ''' <param name="Arguments">打开程序的参数（例如"/?"）</param>
        ''' <param name="Description">鼠标悬停在快捷方式上的描述</param>
        ''' <param name="WorkingDirectory">快捷方式的起始位置（不设置此参数时，按照系统默认，自动设置为快捷方式指向的路径的父目录）</param>
        ''' <returns>是否创建成功</returns>
        ''' <remarks></remarks>
        Public Shared Function CreatLinkFile(ByVal TargetPath As String, ByVal LinkFilePath As String, Optional ByVal Arguments As String = "", Optional ByVal Description As String = "", Optional ByVal WorkingDirectory As String = "") As Boolean
            Try
                If System.IO.File.Exists(LinkFilePath) = True Then System.IO.File.Delete(LinkFilePath)
                Dim Shortcut As Object = CreateObject("WScript.Shell").CreateShortcut(LinkFilePath)
                If WorkingDirectory = "" Then WorkingDirectory = System.IO.Directory.GetParent(LinkFilePath).FullName
                With Shortcut
                    .TargetPath = TargetPath
                    .IconLocation = TargetPath
                    .Arguments = Arguments
                    .Description = Description
                    .WorkingDirectory = WorkingDirectory
                    .Save()
                End With
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function
        ''' <summary>
        ''' 读取快捷方式指向的路径
        ''' </summary>
        ''' <param name="TargetPath">快捷方式指向的路径（获得完整的绝对路径）</param>
        ''' <param name="LinkFilePath">快捷方式文件的路径（可以是相对路径，如"1.lnk"）</param>
        ''' <returns>是否读取成功</returns>
        ''' <remarks></remarks>
        Public Shared Function ReadLinkFile(ByRef TargetPath As String, ByVal LinkFilePath As String) As Boolean
            Try
                Dim Shortcut As Object = CreateObject("WScript.Shell").CreateShortcut(LinkFilePath)
                TargetPath = Shortcut.TargetPath
                Return True
            Catch ex As Exception
                Return False
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
        ''' <summary>
        ''' 对两个字符串数组取交集（取出满足在第一个数组里，也在第二个数组里的元素）
        ''' </summary>
        ''' <param name="StringArray1">字符串数组</param>
        ''' <param name="StringArray2">字符串数组</param>
        ''' <returns>取出的元素的集合</returns>
        ''' <remarks></remarks>
        Public Shared Function Intersect(ByRef StringArray1 As ArrayList, ByRef StringArray2 As ArrayList) As ArrayList
            Dim Temp As ArrayList = New ArrayList
            For Each S1 As String In StringArray1
                For Each S2 As String In StringArray2
                    If S1 = S2 Then
                        Temp.Add(S1)
                        Exit For
                    End If
                Next
            Next
            Return Temp
        End Function
        ''' <summary>
        ''' 对两个字符串数组取交集（取出满足在第一个数组里，也在第二个数组里的元素）
        ''' </summary>
        ''' <param name="StringArray1">字符串数组</param>
        ''' <param name="StringArray2">字符串数组</param>
        ''' <returns>取出的元素的集合</returns>
        ''' <remarks></remarks>
        Public Shared Function Intersect(ByRef StringArray1 As List(Of String), ByRef StringArray2 As List(Of String)) As List(Of String)
            Dim Temp As List(Of String) = New List(Of String)
            For Each S1 As String In StringArray1
                For Each S2 As String In StringArray2
                    If S1 = S2 Then
                        Temp.Add(S1)
                        Exit For
                    End If
                Next
            Next
            Return Temp
        End Function
        ''' <summary>
        ''' 对两个字符串数组取交集（取出满足在第一个数组里，也在第二个数组里的元素）
        ''' </summary>
        ''' <param name="StringArray1">字符串数组</param>
        ''' <param name="StringArray2">字符串数组</param>
        ''' <returns>取出的元素的集合</returns>
        ''' <remarks></remarks>
        Public Shared Function Intersect(ByRef StringArray1 As String(), ByRef StringArray2 As String()) As String()
            Dim Temp As List(Of String) = New List(Of String)
            For Each S1 As String In StringArray1
                For Each S2 As String In StringArray2
                    If S1 = S2 Then
                        Temp.Add(S1)
                        Exit For
                    End If
                Next
            Next
            Return Temp.ToArray
        End Function
        ''' <summary>
        ''' 对两个字符串数组取并集（取出第一个数组里的元素，然后取出第二个数组里的元素）
        ''' </summary>
        ''' <param name="StringArray1">字符串数组</param>
        ''' <param name="StringArray2">字符串数组</param>
        ''' <returns>取出的元素的集合</returns>
        ''' <remarks></remarks>
        Public Shared Function Union(ByRef StringArray1 As ArrayList, ByRef StringArray2 As ArrayList) As ArrayList
            Dim Temp As ArrayList = New ArrayList
            For Each S1 As String In StringArray1
                Temp.Add(S1)
            Next
            For Each S2 As String In StringArray2
                Temp.Add(S2)
            Next
            Return Temp
        End Function
        ''' <summary>
        ''' 对两个字符串数组取并集（取出第一个数组里的元素，然后取出第二个数组里的元素）
        ''' </summary>
        ''' <param name="StringArray1">字符串数组</param>
        ''' <param name="StringArray2">字符串数组</param>
        ''' <returns>取出的元素的集合</returns>
        ''' <remarks></remarks>
        Public Shared Function Union(ByRef StringArray1 As List(Of String), ByRef StringArray2 As List(Of String)) As List(Of String)
            Dim Temp As List(Of String) = New List(Of String)
            For Each S1 As String In StringArray1
                Temp.Add(S1)
            Next
            For Each S2 As String In StringArray2
                Temp.Add(S2)
            Next
            Return Temp
        End Function
        ''' <summary>
        ''' 对两个字符串数组取并集（取出第一个数组里的元素，然后取出第二个数组里的元素）
        ''' </summary>
        ''' <param name="StringArray1">字符串数组</param>
        ''' <param name="StringArray2">字符串数组</param>
        ''' <returns>取出的元素的集合</returns>
        ''' <remarks></remarks>
        Public Shared Function Union(ByRef StringArray1 As String(), ByRef StringArray2 As String()) As String()
            Dim Temp As List(Of String) = New List(Of String)
            For Each S1 As String In StringArray1
                Temp.Add(S1)
            Next
            For Each S2 As String In StringArray2
                Temp.Add(S2)
            Next
            Return Temp.ToArray
        End Function
        ''' <summary>
        ''' 对两个字符串数组取差集（取出满足在第一个数组里，但是不在第二个数组里的元素）
        ''' </summary>
        ''' <param name="StringArray1">字符串数组</param>
        ''' <param name="StringArray2">字符串数组</param>
        ''' <returns>取出的元素的集合</returns>
        ''' <remarks></remarks>
        Public Shared Function Except(ByRef StringArray1 As ArrayList, ByRef StringArray2 As ArrayList) As ArrayList
            Dim Temp As ArrayList = New ArrayList
            Dim Contains As Boolean = False
            For Each S1 As String In StringArray1
                Contains = False
                For Each S2 As String In StringArray2
                    If S1 = S2 Then
                        Contains = True
                        Exit For
                    End If
                Next
                If Contains = False Then
                    Temp.Add(S1)
                End If
            Next
            Return Temp
        End Function
        ''' <summary>
        ''' 对两个字符串数组取差集（取出满足在第一个数组里，但是不在第二个数组里的元素）
        ''' </summary>
        ''' <param name="StringArray1">字符串数组</param>
        ''' <param name="StringArray2">字符串数组</param>
        ''' <returns>取出的元素的集合</returns>
        ''' <remarks></remarks>
        Public Shared Function Except(ByRef StringArray1 As List(Of String), ByRef StringArray2 As List(Of String)) As List(Of String)
            Dim Temp As List(Of String) = New List(Of String)
            Dim Contains As Boolean = False
            For Each S1 As String In StringArray1
                Contains = False
                For Each S2 As String In StringArray2
                    If S1 = S2 Then
                        Contains = True
                        Exit For
                    End If
                Next
                If Contains = False Then
                    Temp.Add(S1)
                End If
            Next
            Return Temp
        End Function
        ''' <summary>
        ''' 对两个字符串数组取差集（取出满足在第一个数组里，但是不在第二个数组里的元素）
        ''' </summary>
        ''' <param name="StringArray1">字符串数组</param>
        ''' <param name="StringArray2">字符串数组</param>
        ''' <returns>取出的元素的集合</returns>
        ''' <remarks></remarks>
        Public Shared Function Except(ByRef StringArray1 As String(), ByRef StringArray2 As String()) As String()
            Dim Temp As List(Of String) = New List(Of String)
            Dim Contains As Boolean = False
            For Each S1 As String In StringArray1
                Contains = False
                For Each S2 As String In StringArray2
                    If S1 = S2 Then
                        Contains = True
                        Exit For
                    End If
                Next
                If Contains = False Then
                    Temp.Add(S1)
                End If
            Next
            Return Temp.ToArray
        End Function
        ''' <summary>
        ''' 将一个对象转换为JSON字符串（必须为常见类型的数组对象，或是含有若干属性的自定义类的对象）
        ''' </summary>
        ''' <param name="JsonObject">含有JSON信息的对象（会识别是否为数组、会识别自定义的类的属性、会进行基本的字符串转义以保证满足JSON格式）</param>
        ''' <returns>JSON字符串</returns>
        ''' <remarks></remarks>
        Public Shared Function ChangeObjectToJson(ByRef JsonObject As Object) As String
            'IsPrimitive，基元类型，包含：Boolean、Byte、SByte、Int16、UInt16、Int32、UInt32、Int64、UInt64、Char、Double和Single
            If JsonObject.GetType.IsPrimitive Or JsonObject.GetType Is GetType(System.String) Or JsonObject.GetType Is GetType(System.DateTime) Then
                Return """" & JsonObject.ToString().Replace(vbCr, "\r").Replace(vbLf, "\n").Replace("""", "\""") & """"
            End If
            'IsGenericType包含了类似List(Of Object)的各种类型
            If JsonObject.GetType.IsGenericType Or JsonObject.GetType.IsArray Or JsonObject.GetType Is GetType(ArrayList) Then
                Dim Result As New System.Text.StringBuilder("[")
                For Each Element As Object In JsonObject
                    Result.Append(ChangeObjectToJson(Element) & ",")
                Next
                If Result.Length > 1 Then Result.Remove(Result.Length - 1, 1)
                Result.Append("]")
                Return Result.ToString()
            Else
                Dim Result As New System.Text.StringBuilder("{")
                For Each P As System.Reflection.PropertyInfo In JsonObject.GetType().GetProperties()
                    Result.Append("""" & P.Name & """" & ":")
                    Dim Element As Object = P.GetValue(JsonObject, Nothing)
                    If Element Is Nothing Then
                        Result.Append("null" & ",")
                    Else
                        Dim ElementType As System.Type = Element.GetType()
                        If ElementType.IsPrimitive Or ElementType Is GetType(System.String) Or ElementType Is GetType(System.DateTime) Then
                            Result.Append("""" & P.GetValue(JsonObject, Nothing).ToString().Replace(vbCr, "\r").Replace(vbLf, "\n").Replace("""", "\""") & """" & ",")
                        Else
                            Result.Append(ChangeObjectToJson(P.GetValue(JsonObject, Nothing)) & ",")
                        End If
                    End If
                Next
                If Result.Length > 1 Then Result.Remove(Result.Length - 1, 1)
                Result.Append("}")
                Return Result.ToString()
            End If
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
    ''' 访问应用程序信息和服务
    ''' </summary>
    ''' <remarks></remarks>
    Partial Class MyApplication

        '在这里处理UI线程异常
        '注意：Application_ThreadException方法执行完成后，应用程序仍会继续运行
        Private Shared Sub Application_ThreadException(ByVal sender As Object, ByVal e As Microsoft.VisualBasic.ApplicationServices.UnhandledExceptionEventArgs) Handles Me.UnhandledException
            'AddHandler Me.UnhandledException, New Threading.ThreadExceptionEventHandler(AddressOf MyApplication.Application_ThreadException)
            MessageBox.Show(e.Exception.ToString())
            'MessageBox.Show(e.ExitApplication)'默认为True
            e.ExitApplication = False
        End Sub

        '在这里处理多线程异常
        '注意：CurrentDomain_UnhandledException方法执行完成后，应用程序就会被终止
        Private Shared Sub CurrentDomain_UnhandledException(ByVal sender As Object, ByVal e As UnhandledExceptionEventArgs) Handles CurrentDomain.UnhandledException
            'AddHandler AppDomain.CurrentDomain.UnhandledException, New UnhandledExceptionEventHandler(AddressOf MyApplication.CurrentDomain_UnhandledException)
            MessageBox.Show(DirectCast(e.ExceptionObject, Exception).ToString())
        End Sub
        Friend WithEvents CurrentDomain As System.AppDomain = AppDomain.CurrentDomain

        '测试抛出异常的效果
        Public Shared Sub TestException()
            Dim Temp As Threading.Thread = New Threading.Thread(New Threading.ThreadStart(AddressOf TestExceptionThread))
            Temp.Start()
            Throw New Exception("窗体线程异常")
        End Sub
        Private Shared Sub TestExceptionThread()
            Throw (New Exception("非窗体线程异常"))
        End Sub

    End Class

    ''' <summary>
    ''' 访问主机及其资源、服务和数据
    ''' </summary>
    ''' <remarks></remarks>
    Partial Class MyComputer
        ''' <summary>
        ''' 延时关闭计算机（注意会取代之前可能存在的关机计划）
        ''' </summary>
        ''' <param name="DelaySecond">延时时间（单位秒，最多可以延时10年）</param>
        ''' <returns>是否执行成功</returns>
        ''' <remarks></remarks>
        Public Function ShutDown(Optional ByVal DelaySecond As Integer = 0) As Boolean
            Try
                Dim p As New Process
                p.StartInfo.FileName = "cmd.exe"
                p.StartInfo.UseShellExecute = False
                p.StartInfo.RedirectStandardInput = True
                p.StartInfo.RedirectStandardOutput = True
                p.StartInfo.RedirectStandardError = True
                p.StartInfo.CreateNoWindow = True
                p.Start()
                p.StandardOutput.ReadLine() 'MsgBox(p.StandardOutput.ReadLine()) 'Microsoft Windows
                p.StandardOutput.ReadLine() 'MsgBox(p.StandardOutput.ReadLine()) '版权所有
                p.StandardOutput.ReadLine() 'MsgBox(p.StandardOutput.ReadLine()) '空行
                p.StandardInput.WriteLine("shutdown -a >nul 2>nul")
                p.StandardInput.WriteLine("shutdown -s -t " & DelaySecond & " >nul 2>nul")
                p.StandardInput.WriteLine("exit")
                p.StandardOutput.ReadToEnd()
                p.StandardOutput.Close()
                p.Dispose()
                '因为已有关机计划，这一句不会起作用（仅供参考写法）
                'Shell("shutdown -s -t " & DelaySecond, AppWinStyle.Hide)
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function
        ''' <summary>
        ''' 延时重启计算机（注意会取代之前可能存在的关机计划）
        ''' </summary>
        ''' <param name="DelaySecond">延时时间（单位秒，最多可以延时10年）</param>
        ''' <returns>是否执行成功</returns>
        ''' <remarks></remarks>
        Public Function ShutDownReboot(Optional ByVal DelaySecond As Integer = 0) As Boolean
            Try
                Dim p As New Process
                p.StartInfo.FileName = "cmd.exe"
                p.StartInfo.UseShellExecute = False
                p.StartInfo.RedirectStandardInput = True
                p.StartInfo.RedirectStandardOutput = True
                p.StartInfo.RedirectStandardError = True
                p.StartInfo.CreateNoWindow = True
                p.Start()
                p.StandardOutput.ReadLine() 'MsgBox(p.StandardOutput.ReadLine()) 'Microsoft Windows
                p.StandardOutput.ReadLine() 'MsgBox(p.StandardOutput.ReadLine()) '版权所有
                p.StandardOutput.ReadLine() 'MsgBox(p.StandardOutput.ReadLine()) '空行
                p.StandardInput.WriteLine("shutdown -a >nul 2>nul")
                p.StandardInput.WriteLine("shutdown -r -t " & DelaySecond & " >nul 2>nul")
                p.StandardInput.WriteLine("exit")
                p.StandardOutput.ReadToEnd()
                p.StandardOutput.Close()
                p.Dispose()
                '因为已有关机计划，这一句不会起作用（仅供参考写法）
                'Shell("shutdown -r -t " & DelaySecond, AppWinStyle.Hide)
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function
        ''' <summary>
        ''' 取消关机计划（没有关机计划时则无效果）
        ''' </summary>
        ''' <returns>是否执行成功</returns>
        ''' <remarks></remarks>
        Public Function ShutDownAbort() As Boolean
            Try
                Dim p As New Process
                p.StartInfo.FileName = "cmd.exe"
                p.StartInfo.UseShellExecute = False
                p.StartInfo.RedirectStandardInput = True
                p.StartInfo.RedirectStandardOutput = True
                p.StartInfo.RedirectStandardError = True
                p.StartInfo.CreateNoWindow = True
                p.Start()
                p.StandardOutput.ReadLine() 'MsgBox(p.StandardOutput.ReadLine()) 'Microsoft Windows
                p.StandardOutput.ReadLine() 'MsgBox(p.StandardOutput.ReadLine()) '版权所有
                p.StandardOutput.ReadLine() 'MsgBox(p.StandardOutput.ReadLine()) '空行
                p.StandardInput.WriteLine("shutdown -a >nul 2>nul")
                p.StandardInput.WriteLine("exit")
                p.StandardOutput.ReadToEnd()
                p.StandardOutput.Close()
                p.Dispose()
                '因为没有关机计划，这一句不会起作用（仅供参考写法）
                'Shell("shutdown -a >nul 2>nul", AppWinStyle.Hide)
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function
        ''' <summary>
        ''' 打开指定的程序（多次调用本函数会打开程序的多个实例，新打开的程序会夺取鼠标焦点）
        ''' </summary>
        ''' <param name="TaskName">程序名称（例如"notepad"或"notepad.exe"）</param>
        ''' <returns>是否执行成功</returns>
        ''' <remarks></remarks>
        Public Function TaskRun(ByVal TaskName As String) As Boolean
            Try
                If TaskName.ToLower.EndsWith(".exe") = False Then TaskName = TaskName & ".exe"
                Dim p As New Process
                p.StartInfo.FileName = "cmd.exe"
                p.StartInfo.UseShellExecute = False
                p.StartInfo.RedirectStandardInput = True
                p.StartInfo.RedirectStandardOutput = True
                p.StartInfo.RedirectStandardError = True
                p.StartInfo.CreateNoWindow = True
                p.Start()
                p.StandardOutput.ReadLine() 'MsgBox(p.StandardOutput.ReadLine()) 'Microsoft Windows
                p.StandardOutput.ReadLine() 'MsgBox(p.StandardOutput.ReadLine()) '版权所有
                p.StandardOutput.ReadLine() 'MsgBox(p.StandardOutput.ReadLine()) '空行
                p.StandardInput.WriteLine(TaskName)
                p.StandardOutput.ReadLine()
                '以下这种写法会导致线程一直阻塞，直到TaskName程序结束才继续执行（CMD认为执行的程序可能会无限输出数据，所以一直等待？？？）：
                'p.StandardInput.WriteLine("exit")
                'p.StandardOutput.ReadToEnd()
                p.StandardOutput.Close()
                p.Dispose()
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function
        ''' <summary>
        ''' 打开指定的程序（多次调用本函数会打开程序的多个实例，新打开的程序不会夺取鼠标焦点）
        ''' </summary>
        ''' <param name="TaskName">程序名称（例如"notepad"或"notepad.exe"）</param>
        ''' <returns>是否执行成功</returns>
        ''' <remarks></remarks>
        Public Function ShellRun(ByVal TaskName As String) As Boolean
            Try
                If TaskName.ToLower.EndsWith(".exe") = False Then TaskName = TaskName & ".exe"
                Shell(TaskName, AppWinStyle.NormalNoFocus)
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function
        ''' <summary>
        ''' 关闭指定的程序（程序如果有多个实例，会一并结束，多次调用本函数无特别效果）
        ''' </summary>
        ''' <param name="TaskName">程序名称（例如"notepad"或"notepad.exe"）</param>
        ''' <returns>是否执行成功</returns>
        ''' <remarks></remarks>
        Public Function TaskKill(ByVal TaskName As String) As Boolean
            Try
                If TaskName.ToLower.EndsWith(".exe") = False Then TaskName = TaskName & ".exe"
                Dim p As New Process
                p.StartInfo.FileName = "cmd.exe"
                p.StartInfo.UseShellExecute = False
                p.StartInfo.RedirectStandardInput = True
                p.StartInfo.RedirectStandardOutput = True
                p.StartInfo.RedirectStandardError = True
                p.StartInfo.CreateNoWindow = True
                p.Start()
                p.StandardOutput.ReadLine() 'MsgBox(p.StandardOutput.ReadLine()) 'Microsoft Windows
                p.StandardOutput.ReadLine() 'MsgBox(p.StandardOutput.ReadLine()) '版权所有
                p.StandardOutput.ReadLine() 'MsgBox(p.StandardOutput.ReadLine()) '空行
                p.StandardInput.WriteLine("taskkill /f /t /im " & TaskName)
                p.StandardInput.WriteLine("exit")
                p.StandardOutput.ReadToEnd()
                p.StandardOutput.Close()
                p.Dispose()
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function
        ''' <summary>
        ''' 关闭指定的程序（程序如果有多个实例，会一并结束，多次调用本函数无特别效果）
        ''' </summary>
        ''' <param name="TaskName">程序名称（例如"notepad"或"notepad.exe"）</param>
        ''' <returns>是否执行成功</returns>
        ''' <remarks></remarks>
        Public Function ShellKill(ByVal TaskName As String) As Boolean
            Try
                If TaskName.ToLower.EndsWith(".exe") = False Then TaskName = TaskName & ".exe"
                Shell("taskkill /f /t /im " & TaskName, AppWinStyle.Hide)
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function
        ''' <summary>
        ''' 获取屏幕截图（默认保存格式，实测1920*1080分辨率的截图文件大小为293K，文件大小中等）
        ''' </summary>
        ''' <param name="ScreenFilePath">截图文件保存路径</param>
        ''' <returns>是否执行成功</returns>
        ''' <remarks></remarks>
        Public Function SaveScreen(ByVal ScreenFilePath As String) As Boolean
            Try
                Dim MyRectangle As Rectangle = My.Computer.Screen.Bounds
                Dim MyBmp As New Bitmap(MyRectangle.Width, MyRectangle.Height)
                Using MyGraphics As Graphics = Graphics.FromImage(MyBmp)
                    MyGraphics.CopyFromScreen(0, 0, MyRectangle.Left, MyRectangle.Top, MyRectangle.Size)
                End Using
                MyBmp.Save(ScreenFilePath)
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function
        ''' <summary>
        ''' 获取屏幕截图（bmp格式，实测1920*1080分辨率的截图文件大小为7.91M，文件大小最大）
        ''' </summary>
        ''' <param name="ScreenFilePath">截图文件保存路径</param>
        ''' <returns>是否执行成功</returns>
        ''' <remarks></remarks>
        Public Function SaveScreenBmp(ByVal ScreenFilePath As String) As Boolean
            Try
                Dim MyRectangle As Rectangle = My.Computer.Screen.Bounds
                Dim MyBmp As New Bitmap(MyRectangle.Width, MyRectangle.Height)
                Using MyGraphics As Graphics = Graphics.FromImage(MyBmp)
                    MyGraphics.CopyFromScreen(0, 0, MyRectangle.Left, MyRectangle.Top, MyRectangle.Size)
                End Using
                MyBmp.Save(ScreenFilePath, System.Drawing.Imaging.ImageFormat.Bmp)
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function
        ''' <summary>
        ''' 获取屏幕截图（png格式，实测1920*1080分辨率的截图文件大小为293K，文件大小中等）
        ''' </summary>
        ''' <param name="ScreenFilePath">截图文件保存路径</param>
        ''' <returns>是否执行成功</returns>
        ''' <remarks></remarks>
        Public Function SaveScreenPng(ByVal ScreenFilePath As String) As Boolean
            Try
                Dim MyRectangle As Rectangle = My.Computer.Screen.Bounds
                Dim MyBmp As New Bitmap(MyRectangle.Width, MyRectangle.Height)
                Using MyGraphics As Graphics = Graphics.FromImage(MyBmp)
                    MyGraphics.CopyFromScreen(0, 0, MyRectangle.Left, MyRectangle.Top, MyRectangle.Size)
                End Using
                MyBmp.Save(ScreenFilePath, System.Drawing.Imaging.ImageFormat.Png)
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function
        ''' <summary>
        ''' 获取屏幕截图（jpeg格式，实测1920*1080分辨率的截图文件大小为212K，文件大小最小）
        ''' </summary>
        ''' <param name="ScreenFilePath">截图文件保存路径</param>
        ''' <returns>是否执行成功</returns>
        ''' <remarks></remarks>
        Public Function SaveScreenJpeg(ByVal ScreenFilePath As String) As Boolean
            Try
                Dim MyRectangle As Rectangle = My.Computer.Screen.Bounds
                Dim MyBmp As New Bitmap(MyRectangle.Width, MyRectangle.Height)
                Using MyGraphics As Graphics = Graphics.FromImage(MyBmp)
                    MyGraphics.CopyFromScreen(0, 0, MyRectangle.Left, MyRectangle.Top, MyRectangle.Size)
                End Using
                MyBmp.Save(ScreenFilePath, System.Drawing.Imaging.ImageFormat.Jpeg)
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function
        ''' <summary>
        ''' 获取屏幕截图（gif格式，实测1920*1080分辨率的截图文件大小为232K，文件大小较小）
        ''' </summary>
        ''' <param name="ScreenFilePath">截图文件保存路径</param>
        ''' <returns>是否执行成功</returns>
        ''' <remarks></remarks>
        Public Function SaveScreenGif(ByVal ScreenFilePath As String) As Boolean
            Try
                Dim MyRectangle As Rectangle = My.Computer.Screen.Bounds
                Dim MyBmp As New Bitmap(MyRectangle.Width, MyRectangle.Height)
                Using MyGraphics As Graphics = Graphics.FromImage(MyBmp)
                    MyGraphics.CopyFromScreen(0, 0, MyRectangle.Left, MyRectangle.Top, MyRectangle.Size)
                End Using
                MyBmp.Save(ScreenFilePath, System.Drawing.Imaging.ImageFormat.Gif)
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function
        ''' <summary>
        ''' 获取屏幕截图（ico格式，实测1920*1080分辨率的截图文件大小为294K，文件大小中等）
        ''' </summary>
        ''' <param name="ScreenFilePath">截图文件保存路径</param>
        ''' <returns>是否执行成功</returns>
        ''' <remarks></remarks>
        Public Function SaveScreenIcon(ByVal ScreenFilePath As String) As Boolean
            Try
                Dim MyRectangle As Rectangle = My.Computer.Screen.Bounds
                Dim MyBmp As New Bitmap(MyRectangle.Width, MyRectangle.Height)
                Using MyGraphics As Graphics = Graphics.FromImage(MyBmp)
                    MyGraphics.CopyFromScreen(0, 0, MyRectangle.Left, MyRectangle.Top, MyRectangle.Size)
                End Using
                MyBmp.Save(ScreenFilePath, System.Drawing.Imaging.ImageFormat.Icon)
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function
        ''' <summary>
        ''' 获取屏幕截图（tiff格式，实测1920*1080分辨率的截图文件大小为388K，文件大小较大）
        ''' </summary>
        ''' <param name="ScreenFilePath">截图文件保存路径</param>
        ''' <returns>是否执行成功</returns>
        ''' <remarks></remarks>
        Public Function SaveScreenTiff(ByVal ScreenFilePath As String) As Boolean
            Try
                Dim MyRectangle As Rectangle = My.Computer.Screen.Bounds
                Dim MyBmp As New Bitmap(MyRectangle.Width, MyRectangle.Height)
                Using MyGraphics As Graphics = Graphics.FromImage(MyBmp)
                    MyGraphics.CopyFromScreen(0, 0, MyRectangle.Left, MyRectangle.Top, MyRectangle.Size)
                End Using
                MyBmp.Save(ScreenFilePath, System.Drawing.Imaging.ImageFormat.Tiff)
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function
        ''' <summary>
        ''' 获取屏幕截图（exif格式，实测1920*1080分辨率的截图文件大小为294K，文件大小中等）
        ''' </summary>
        ''' <param name="ScreenFilePath">截图文件保存路径</param>
        ''' <returns>是否执行成功</returns>
        ''' <remarks></remarks>
        Public Function SaveScreenExif(ByVal ScreenFilePath As String) As Boolean
            Try
                Dim MyRectangle As Rectangle = My.Computer.Screen.Bounds
                Dim MyBmp As New Bitmap(MyRectangle.Width, MyRectangle.Height)
                Using MyGraphics As Graphics = Graphics.FromImage(MyBmp)
                    MyGraphics.CopyFromScreen(0, 0, MyRectangle.Left, MyRectangle.Top, MyRectangle.Size)
                End Using
                MyBmp.Save(ScreenFilePath, System.Drawing.Imaging.ImageFormat.Exif)
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function
        ''' <summary>
        ''' 获取屏幕截图（memorybmp格式，实测1920*1080分辨率的截图文件大小为294K，文件大小中等）
        ''' </summary>
        ''' <param name="ScreenFilePath">截图文件保存路径</param>
        ''' <returns>是否执行成功</returns>
        ''' <remarks></remarks>
        Public Function SaveScreenMemoryBmp(ByVal ScreenFilePath As String) As Boolean
            Try
                Dim MyRectangle As Rectangle = My.Computer.Screen.Bounds
                Dim MyBmp As New Bitmap(MyRectangle.Width, MyRectangle.Height)
                Using MyGraphics As Graphics = Graphics.FromImage(MyBmp)
                    MyGraphics.CopyFromScreen(0, 0, MyRectangle.Left, MyRectangle.Top, MyRectangle.Size)
                End Using
                MyBmp.Save(ScreenFilePath, System.Drawing.Imaging.ImageFormat.MemoryBmp)
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function
        ''' <summary>
        ''' 获取屏幕截图（emf格式，实测1920*1080分辨率的截图文件大小为293K，文件大小中等）
        ''' </summary>
        ''' <param name="ScreenFilePath">截图文件保存路径</param>
        ''' <returns>是否执行成功</returns>
        ''' <remarks></remarks>
        Public Function SaveScreenEmf(ByVal ScreenFilePath As String) As Boolean
            Try
                Dim MyRectangle As Rectangle = My.Computer.Screen.Bounds
                Dim MyBmp As New Bitmap(MyRectangle.Width, MyRectangle.Height)
                Using MyGraphics As Graphics = Graphics.FromImage(MyBmp)
                    MyGraphics.CopyFromScreen(0, 0, MyRectangle.Left, MyRectangle.Top, MyRectangle.Size)
                End Using
                MyBmp.Save(ScreenFilePath, System.Drawing.Imaging.ImageFormat.Emf)
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function
        ''' <summary>
        ''' 获取屏幕截图（wmf格式，实测1920*1080分辨率的截图文件大小为293K，文件大小中等）
        ''' </summary>
        ''' <param name="ScreenFilePath">截图文件保存路径</param>
        ''' <returns>是否执行成功</returns>
        ''' <remarks></remarks>
        Public Function SaveScreenWmf(ByVal ScreenFilePath As String) As Boolean
            Try
                Dim MyRectangle As Rectangle = My.Computer.Screen.Bounds
                Dim MyBmp As New Bitmap(MyRectangle.Width, MyRectangle.Height)
                Using MyGraphics As Graphics = Graphics.FromImage(MyBmp)
                    MyGraphics.CopyFromScreen(0, 0, MyRectangle.Left, MyRectangle.Top, MyRectangle.Size)
                End Using
                MyBmp.Save(ScreenFilePath, System.Drawing.Imaging.ImageFormat.Wmf)
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function
        ''' <summary>
        ''' 获取屏幕缩略图（默认保存格式，实测1920*1080【长宽都只保留50%】分辨率，缩略图文件大小为157K，文件大小中等）
        ''' </summary>
        ''' <param name="ScreenFilePath">截图文件保存路径</param>
        ''' <returns>是否执行成功</returns>
        ''' <remarks></remarks>
        Public Function SaveScreenThumbnail(ByVal ScreenFilePath As String) As Boolean
            Try
                Dim MyRectangle As Rectangle = My.Computer.Screen.Bounds
                Dim MyBmp As New Bitmap(MyRectangle.Width, MyRectangle.Height)
                Using MyGraphics As Graphics = Graphics.FromImage(MyBmp)
                    MyGraphics.CopyFromScreen(0, 0, MyRectangle.Left, MyRectangle.Top, MyRectangle.Size)
                End Using
                MyBmp = MyBmp.GetThumbnailImage(MyRectangle.Width / 2, MyRectangle.Height / 2, Nothing, New System.IntPtr(0))
                MyBmp.Save(ScreenFilePath)
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function
        ''' <summary>
        ''' 获取屏幕缩略图（png格式，实测1920*1080【长宽都只保留50%】分辨率，缩略图文件大小为157K，文件大小中等）
        ''' </summary>
        ''' <param name="ScreenFilePath">截图文件保存路径</param>
        ''' <returns>是否执行成功</returns>
        ''' <remarks></remarks>
        Public Function SaveScreenPngThumbnail(ByVal ScreenFilePath As String) As Boolean
            Try
                Dim MyRectangle As Rectangle = My.Computer.Screen.Bounds
                Dim MyBmp As New Bitmap(MyRectangle.Width, MyRectangle.Height)
                Using MyGraphics As Graphics = Graphics.FromImage(MyBmp)
                    MyGraphics.CopyFromScreen(0, 0, MyRectangle.Left, MyRectangle.Top, MyRectangle.Size)
                End Using
                MyBmp = MyBmp.GetThumbnailImage(MyRectangle.Width / 2, MyRectangle.Height / 2, Nothing, New System.IntPtr(0))
                MyBmp.Save(ScreenFilePath, System.Drawing.Imaging.ImageFormat.Png)
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function
        ''' <summary>
        ''' 获取屏幕缩略图（jpeg格式，实测1920*1080【长宽都只保留50%】分辨率，缩略图文件大小为64K，文件大小最小）
        ''' </summary>
        ''' <param name="ScreenFilePath">截图文件保存路径</param>
        ''' <returns>是否执行成功</returns>
        ''' <remarks></remarks>
        Public Function SaveScreenJpegThumbnail(ByVal ScreenFilePath As String) As Boolean
            Try
                Dim MyRectangle As Rectangle = My.Computer.Screen.Bounds
                Dim MyBmp As New Bitmap(MyRectangle.Width, MyRectangle.Height)
                Using MyGraphics As Graphics = Graphics.FromImage(MyBmp)
                    MyGraphics.CopyFromScreen(0, 0, MyRectangle.Left, MyRectangle.Top, MyRectangle.Size)
                End Using
                MyBmp = MyBmp.GetThumbnailImage(MyRectangle.Width / 2, MyRectangle.Height / 2, Nothing, New System.IntPtr(0))
                MyBmp.Save(ScreenFilePath, System.Drawing.Imaging.ImageFormat.Jpeg)
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function
        ''' <summary>
        ''' 发送按键（使用Microsoft.VisualBasic.Devices.Computer类的Keyboard.SendKeys方式）
        ''' </summary>
        ''' <param name="Keys">按键字符串（参照SendKeys的规则）</param>
        ''' <returns>是否执行成功</returns>
        ''' <remarks></remarks>
        Public Function SendKeys(ByVal Keys As String) As Boolean
            Try
                Dim A As New Microsoft.VisualBasic.Devices.Computer()
                A.Keyboard.SendKeys(Keys)
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function
        ''' <summary>
        ''' 发送按键（使用Wscript.Shell的SendKeys方式）
        ''' </summary>
        ''' <param name="Keys">按键字符串（参照SendKeys的规则）</param>
        ''' <returns>是否执行成功</returns>
        ''' <remarks></remarks>
        Public Function SendKeysWshShell(ByVal Keys As String) As Boolean
            Try
                CreateObject("Wscript.Shell").SendKeys(Keys)
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function
        ''' <summary>
        ''' 发送按键（使用System.Windows.Forms.SendKeys.Send方式）
        ''' </summary>
        ''' <param name="Keys">按键字符串（参照SendKeys的规则）</param>
        ''' <returns>是否执行成功</returns>
        ''' <remarks></remarks>
        Public Function SendKeysWinForm(ByVal Keys As String) As Boolean
            Try
                System.Windows.Forms.SendKeys.Send(Keys)
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function
        Private Declare Sub keybd_event Lib "user32.dll" Alias "keybd_event" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
        ''' <summary>
        ''' 按下单个按键（并保持按下状态直到下次按同一个键，连续调用本函数，可执行组合键）
        ''' </summary>
        ''' <param name="Key">键位</param>
        ''' <returns>是否执行成功</returns>
        ''' <remarks></remarks>
        Public Function PressKey(ByVal Key As Keys) As Boolean
            Try
                keybd_event(Key, 0, 0, 0)
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function
        ''' <summary>
        ''' 按下单个按键（并保持按下状态直到下次按同一个键，连续调用本函数，可执行组合键）
        ''' </summary>
        ''' <param name="KeyChar">键位</param>
        ''' <returns>是否执行成功</returns>
        ''' <remarks></remarks>
        Public Function PressKey(ByVal KeyChar As Char) As Boolean
            Try
                keybd_event(Asc(KeyChar), 0, 0, 0)
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function
        ''' <summary>
        ''' 按下多个按键（并保持按下状态直到下次按同一个键，连续调用本函数，可执行组合键）
        ''' </summary>
        ''' <param name="KeyString">键位（只允许字母和数字）</param>
        ''' <returns>是否执行成功</returns>
        ''' <remarks></remarks>
        Public Function PressKey(ByVal KeyString As String) As Boolean
            Try
                Dim Temp As Char() = KeyString.ToCharArray
                For Each TempChar As Char In Temp
                    keybd_event(Asc(TempChar), 0, 0, 0)
                Next
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function
        ''' <summary>
        ''' 释放单个按键（取消按下状态）
        ''' </summary>
        ''' <param name="Key">键位</param>
        ''' <returns>是否执行成功</returns>
        ''' <remarks></remarks>
        Public Function ReleaseKey(ByVal Key As Keys) As Boolean
            Try
                keybd_event(Key, 0, 2, 0)
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function
        ''' <summary>
        ''' 释放单个按键（取消按下状态）
        ''' </summary>
        ''' <param name="KeyChar">键位</param>
        ''' <returns>是否执行成功</returns>
        ''' <remarks></remarks>
        Public Function ReleaseKey(ByVal KeyChar As Char) As Boolean
            Try
                keybd_event(Asc(KeyChar), 0, 2, 0)
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function
        ''' <summary>
        ''' 释放多个按键（取消按下状态）
        ''' </summary>
        ''' <param name="KeyString">键位（只允许字母和数字）</param>
        ''' <returns>是否执行成功</returns>
        ''' <remarks></remarks>
        Public Function ReleaseKey(ByVal KeyString As String) As Boolean
            Try
                Dim Temp As Char() = KeyString.ToCharArray
                For Each TempChar As Char In Temp
                    keybd_event(Asc(TempChar), 0, 2, 0)
                Next
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function
        Private Declare Function mouse_event Lib "user32.dll" Alias "mouse_event" (ByVal dwFlags As MouseEvent, ByVal dX As Int32, ByVal dY As Int32, ByVal dwData As Int32, ByVal dwExtraInfo As Int32) As Boolean
        <Flags()> _
        Private Enum MouseEvent
            Move = &H1
            AbsoluteLocation = &H8000
            LeftButtonDown = &H2
            LeftButtonUp = &H4
            MiddleButtonDown = &H20
            MiddleButtonUp = &H40
            RightButtonDown = &H8
            RightButtonUp = &H10
            Wheel = &H800
            AbsoluteScale = 65535
        End Enum
        ''' <summary>
        ''' 将鼠标位置移动一段距离（移动距离单位为像素，使用Win32 API）
        ''' </summary>
        ''' <param name="x">横向距离</param>
        ''' <param name="y">纵向距离</param>
        ''' <returns>是否执行成功</returns>
        ''' <remarks></remarks>
        Public Function MouseMoveByPixel(Optional ByVal x As Integer = 0, Optional ByVal y As Integer = 0) As Boolean
            Try
                mouse_event(MouseEvent.Move, x, y, 0, 0)
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function
        ''' <summary>
        ''' 将鼠标位置移动一段距离（移动距离单位为屏幕百分比）
        ''' </summary>
        ''' <param name="x">横向距离</param>
        ''' <param name="y">纵向距离</param>
        ''' <returns>是否执行成功</returns>
        ''' <remarks></remarks>
        Public Function MouseMoveByPercent(Optional ByVal x As Double = 0, Optional ByVal y As Double = 0) As Boolean
            Try
                mouse_event(MouseEvent.Move, x * My.Computer.Screen.Bounds.Width, y * My.Computer.Screen.Bounds.Height, 0, 0)
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function
        ''' <summary>
        ''' 移动鼠标到指定位置（定位单位为像素）
        ''' </summary>
        ''' <param name="x">横坐标</param>
        ''' <param name="y">纵坐标</param>
        ''' <returns>是否执行成功</returns>
        ''' <remarks></remarks>
        Public Function MouseMoveToPixel(Optional ByVal x As Integer = 0, Optional ByVal y As Integer = 0) As Boolean
            Try
                mouse_event(MouseEvent.Move Or MouseEvent.AbsoluteLocation, x / My.Computer.Screen.Bounds.Width * MouseEvent.AbsoluteScale, y / My.Computer.Screen.Bounds.Height * MouseEvent.AbsoluteScale, 0, 0)
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function
        ''' <summary>
        ''' 移动鼠标到指定位置（定位单位为屏幕百分比）
        ''' </summary>
        ''' <param name="x">横坐标</param>
        ''' <param name="y">纵坐标</param>
        ''' <returns>是否执行成功</returns>
        ''' <remarks></remarks>
        Public Function MouseMoveToPercent(Optional ByVal x As Double = 0, Optional ByVal y As Double = 0) As Boolean
            Try
                mouse_event(MouseEvent.Move Or MouseEvent.AbsoluteLocation, x * MouseEvent.AbsoluteScale, y * MouseEvent.AbsoluteScale, 0, 0)
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
        Public Function MouseLeftDown() As Boolean
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
        Public Function MouseLeftUp() As Boolean
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
        Public Function MouseLeftClick() As Boolean
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
        Public Function MouseLeftDoubleClick() As Boolean
            Try
                mouse_event(MouseEvent.LeftButtonDown Or MouseEvent.LeftButtonUp Or MouseEvent.LeftButtonDown Or MouseEvent.LeftButtonUp, 0, 0, 0, 0)
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
        Public Function MouseMiddleDown() As Boolean
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
        Public Function MouseMiddleUp() As Boolean
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
        Public Function MouseMiddleClick() As Boolean
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
        Public Function MouseMiddleDoubleClick() As Boolean
            Try
                mouse_event(MouseEvent.MiddleButtonDown Or MouseEvent.MiddleButtonUp Or MouseEvent.MiddleButtonDown Or MouseEvent.MiddleButtonUp, 0, 0, 0, 0)
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
        Public Function MouseRightDown() As Boolean
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
        Public Function MouseRightUp() As Boolean
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
        Public Function MouseRightClick() As Boolean
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
        Public Function MouseRightDoubleClick() As Boolean
            Try
                mouse_event(MouseEvent.RightButtonDown Or MouseEvent.RightButtonUp Or MouseEvent.RightButtonDown Or MouseEvent.RightButtonUp, 0, 0, 0, 0)
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function
        ''' <summary>
        ''' 鼠标滚轮向上滚动（滚动距离单位为像素）
        ''' </summary>
        ''' <param name="ScrollValue">滚动距离</param>
        ''' <returns>是否执行成功</returns>
        ''' <remarks></remarks>
        Public Function MouseWheelUp(ByVal ScrollValue As Integer) As Boolean
            Try
                mouse_event(MouseEvent.Wheel, 0, 0, ScrollValue, 0)
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function
        ''' <summary>
        ''' 鼠标滚轮向下滚动（滚动距离单位为像素）
        ''' </summary>
        ''' <param name="ScrollValue">滚动距离</param>
        ''' <returns>是否执行成功</returns>
        ''' <remarks></remarks>
        Public Function MouseWheelDown(ByVal ScrollValue As Integer) As Boolean
            Try
                mouse_event(MouseEvent.Wheel, 0, 0, -ScrollValue, 0)
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function
        Private Declare Function GetAsyncKeyState Lib "user32.dll" Alias "GetAsyncKeyState" (ByVal vKey As Long) As Integer
        ''' <summary>
        ''' 判断物理键盘设备上的单个键位是否正处于被按下的状态（侦测键盘的硬件中断）
        ''' </summary>
        ''' <param name="Key">键位</param>
        ''' <returns>是否按下</returns>
        ''' <remarks></remarks>
        Public Function KeyBeingPressed(ByVal Key As Keys) As Boolean
            Dim Temp As Integer = GetAsyncKeyState(Key)
            If Temp = -32768 Or Temp = -32767 Then Return True
            Return False
        End Function
        ''' <summary>
        ''' 判断物理键盘设备上的单个键位是否正处于被按下的状态（侦测键盘的硬件中断）
        ''' </summary>
        ''' <param name="KeyChar">键位</param>
        ''' <returns>是否按下</returns>
        ''' <remarks></remarks>
        Public Function KeyBeingPressed(ByVal KeyChar As Char) As Boolean
            Dim Temp As Integer = GetAsyncKeyState(Asc(KeyChar))
            If Temp = -32768 Or Temp = -32767 Then Return True
            Return False
        End Function
        ''' <summary>
        ''' 判断物理键盘设备上的多个键位是否都处于被按下的状态（侦测键盘的硬件中断）
        ''' </summary>
        ''' <param name="KeyString">键位（只允许字母和数字）</param>
        ''' <returns>是否按下</returns>
        ''' <remarks></remarks>
        Public Function KeyBeingPressed(ByVal KeyString As String) As Boolean
            Dim Temp As Char() = KeyString.ToCharArray
            For Each TempChar As Char In Temp
                Dim TempInteger As Integer = GetAsyncKeyState(Asc(TempChar))
                If TempInteger <> -32768 And TempInteger <> -32767 Then Return False
            Next
            Return True
        End Function
        Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As IntPtr
        Private Declare Function GetWindowRect Lib "user32" Alias "GetWindowRect" (ByVal hwnd As IntPtr, ByRef lpRect As RECT) As IntPtr
        Private Structure RECT
            Dim Left As Integer
            Dim Top As Integer
            Dim Right As Integer
            Dim Bottom As Integer
        End Structure
        ''' <summary>
        ''' 根据窗口标题获取窗口的大小和位置（当多个标题相同的窗体存在时，默认获取上一个活动的窗体）
        ''' </summary>
        ''' <param name="WindowTitle">窗口标题（字符串不能有任何差别）</param>
        ''' <returns>窗口的大小和位置（System.Drawing.Rectangle）</returns>
        ''' <remarks></remarks>
        Public Function FindWindow(ByVal WindowTitle As String) As System.Drawing.Rectangle
            Dim hWnd As IntPtr
            Dim Rect As RECT
            hWnd = FindWindow(vbNullString, WindowTitle)
            GetWindowRect(hWnd, Rect)
            Return New System.Drawing.Rectangle(Rect.Left, Rect.Top, Rect.Right - Rect.Left, Rect.Bottom - Rect.Top)
        End Function
        Private Declare Function GetForegroundWindow Lib "user32" () As IntPtr
        ''' <summary>
        ''' 获取当前的焦点窗口的大小和位置
        ''' </summary>
        ''' <returns>窗口的大小和位置（System.Drawing.Rectangle）</returns>
        ''' <remarks></remarks>
        Public Function FindFocusWindow() As System.Drawing.Rectangle
            Dim Rect As RECT
            GetWindowRect(GetForegroundWindow(), Rect)
            Return New System.Drawing.Rectangle(Rect.Left, Rect.Top, Rect.Right - Rect.Left, Rect.Bottom - Rect.Top)
        End Function
        Private Declare Function MoveWindow Lib "user32" Alias "MoveWindow" (ByVal hwnd As IntPtr, ByVal X As Integer, ByVal Y As Integer, ByVal nWidth As Integer, ByVal nHeight As Integer, ByVal bRepaint As Boolean) As Boolean
        ''' <summary>
        ''' 根据窗口标题修改窗口的位置（当多个标题相同的窗体存在时，默认修改上一个活动的窗体；注意可能会把窗口移动到用户鼠标无法触及的位置）
        ''' </summary>
        ''' <param name="WindowTitle">窗口标题（字符串不能有任何差别）</param>
        ''' <param name="X">Left属性</param>
        ''' <param name="Y">Top属性</param>
        ''' <returns>是否修改成功</returns>
        ''' <remarks></remarks>
        Public Function MoveWindow(ByVal WindowTitle As String, Optional ByVal X As Integer = 0, Optional ByVal Y As Integer = 0) As Boolean
            Dim hWnd As IntPtr
            Dim Rect As RECT
            hWnd = FindWindow(vbNullString, WindowTitle)
            GetWindowRect(hWnd, Rect)
            If hWnd = 0 Then Return False
            MoveWindow(hWnd, Rect.Left + X, Rect.Top + Y, Rect.Right - Rect.Left, Rect.Bottom - Rect.Top, True)
            Return True
        End Function
        ''' <summary>
        ''' 根据窗口标题修改窗口的大小（当多个标题相同的窗体存在时，默认修改上一个活动的窗体；注意某些程序的窗口大小，实际上不能被修改的太小）
        ''' </summary>
        ''' <param name="WindowTitle">窗口标题（字符串不能有任何差别）</param>
        ''' <param name="Width">宽度</param>
        ''' <param name="Height">高度</param>
        ''' <returns>是否修改成功</returns>
        ''' <remarks></remarks>
        Public Function ResizeWindow(ByVal WindowTitle As String, Optional ByVal Width As Integer = 0, Optional ByVal Height As Integer = 0) As Boolean
            Dim hWnd As IntPtr
            Dim Rect As RECT
            hWnd = FindWindow(vbNullString, WindowTitle)
            GetWindowRect(hWnd, Rect)
            If hWnd = 0 Then Return False
            MoveWindow(hWnd, Rect.Left, Rect.Top, Width, Height, True)
            Return True
        End Function
        ''' <summary>
        ''' 将鼠标位置移动一段距离（移动距离单位为像素，使用System.Windows.Forms.Cursor）
        ''' </summary>
        ''' <param name="x">横向距离</param>
        ''' <param name="y">纵向距离</param>
        ''' <returns>是否执行成功</returns>
        ''' <remarks></remarks>
        Public Function MouseMove(ByVal x As Integer, ByVal y As Integer) As Boolean
            Try
                Dim P As System.Drawing.Point = System.Windows.Forms.Cursor.Position
                P.Offset(x, y)
                System.Windows.Forms.Cursor.Position = P
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function
        ''' <summary>
        ''' 获取鼠标的当前位置（System.Windows.Forms.Cursor.Position）
        ''' </summary>
        ''' <returns>位置坐标（System.Drawing.Point）</returns>
        ''' <remarks></remarks>
        Public Function MousePosition() As System.Drawing.Point
            Return System.Windows.Forms.Cursor.Position
        End Function
        ''' <summary>
        ''' 获取鼠标的当前位置的屏幕颜色
        ''' </summary>
        ''' <returns>颜色值（System.Drawing.Color）</returns>
        ''' <remarks></remarks>
        Public Function MousePositionColor() As System.Drawing.Color
            Dim MyRectangle As Rectangle = My.Computer.Screen.Bounds
            Dim MyBmp As New Bitmap(MyRectangle.Width, MyRectangle.Height)
            Using MyGraphics As Graphics = Graphics.FromImage(MyBmp)
                MyGraphics.CopyFromScreen(0, 0, MyRectangle.Left, MyRectangle.Top, MyRectangle.Size)
            End Using
            Return MyBmp.GetPixel(System.Windows.Forms.Cursor.Position.X, System.Windows.Forms.Cursor.Position.Y)
        End Function
        ''' <summary>
        ''' 强制把所有运行的程序窗口最小化，显示桌面（效果类似Win7系统鼠标点击屏幕右下角）
        ''' </summary>
        ''' <returns>是否执行成功</returns>
        ''' <remarks></remarks>
        Public Function ShowDesktop() As Boolean
            Try
                CreateObject("Shell.Application").ToggleDesktop()
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function
    End Class
End Namespace