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
            Return Convert.ToBase64String(System.Text.Encoding.UTF8.GetBytes(Source))
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
            Return System.Text.Encoding.UTF8.GetString(Convert.FromBase64String(Source))
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
