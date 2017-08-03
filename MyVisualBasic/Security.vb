Namespace My

    ''' <summary>
    ''' 编码解码、加密解密相关函数
    ''' </summary>
    ''' <remarks></remarks>
    Partial Public NotInheritable Class Security

        ''' <summary>
        ''' URL编码
        ''' </summary>
        ''' <param name="Source">要编码的字符串</param>
        ''' <returns>编码后的结果字符串</returns>
        ''' <remarks></remarks>
        Public Shared Function URL_Encode(ByVal Source As String) As String
            Return System.Web.HttpUtility.UrlEncode(Source, System.Text.Encoding.UTF8)
        End Function
        ''' <summary>
        ''' URL编码
        ''' </summary>
        ''' <param name="Source">要编码的字符串</param>
        ''' <param name="Encoding">使用特定的字符编码（默认UTF-8）</param>
        ''' <returns>编码后的结果字符串</returns>
        ''' <remarks></remarks>
        Public Shared Function URL_Encode(ByVal Source As String, ByVal Encoding As System.Text.Encoding) As String
            Return System.Web.HttpUtility.UrlEncode(Source, Encoding)
        End Function

        ''' <summary>
        ''' URL解码
        ''' </summary>
        ''' <param name="Source">要解码的字符串</param>
        ''' <returns>解码后的结果字符串</returns>
        ''' <remarks></remarks>
        Public Shared Function URL_Decode(ByVal Source As String) As String
            Return System.Web.HttpUtility.UrlDecode(Source, System.Text.Encoding.UTF8)
        End Function
        ''' <summary>
        ''' URL解码
        ''' </summary>
        ''' <param name="Source">要解码的字符串</param>
        ''' <param name="Encoding">使用特定的字符编码（默认UTF-8）</param>
        ''' <returns>解码后的结果字符串</returns>
        ''' <remarks></remarks>
        Public Shared Function URL_Decode(ByVal Source As String, ByVal Encoding As System.Text.Encoding) As String
            Return System.Web.HttpUtility.UrlDecode(Source, Encoding)
        End Function



        ''' <summary>
        ''' Base64编码（编码结果的字符串中包含字母A-Z，a-z，数字0-9，符号+/=）
        ''' </summary>
        ''' <param name="Source">要编码的字符串</param>
        ''' <returns>编码后的结果字符串</returns>
        ''' <remarks></remarks>
        Public Shared Function Base64_Encode(ByVal Source As String) As String
            Return Convert.ToBase64String(System.Text.Encoding.UTF8.GetBytes(Source))
        End Function
        ''' <summary>
        ''' Base64编码（编码结果的字符串中包含字母A-Z，a-z，数字0-9，符号+/=）
        ''' </summary>
        ''' <param name="Source">要编码的字符串</param>
        ''' <param name="Encoding">使用特定的字符编码（默认UTF-8）</param>
        ''' <returns>编码后的结果字符串</returns>
        ''' <remarks></remarks>
        Public Shared Function Base64_Encode(ByVal Source As String, ByVal Encoding As System.Text.Encoding) As String
            Return Convert.ToBase64String(Encoding.GetBytes(Source))
        End Function

        ''' <summary>
        ''' Base64解码（要解码的字符串可以包含字母A-Z，a-z，数字0-9，符号+/=）
        ''' </summary>
        ''' <param name="Source">要解码的字符串</param>
        ''' <returns>解码后的结果字符串</returns>
        ''' <remarks></remarks>
        Public Shared Function Base64_Decode(ByVal Source As String) As String
            Return System.Text.Encoding.UTF8.GetString(Convert.FromBase64String(Source))
        End Function
        ''' <summary>
        ''' Base64解码（要解码的字符串可以包含字母A-Z，a-z，数字0-9，符号+/=）
        ''' </summary>
        ''' <param name="Source">要解码的字符串</param>
        ''' <param name="Encoding">使用特定的字符编码（默认UTF-8）</param>
        ''' <returns>解码后的结果字符串</returns>
        ''' <remarks></remarks>
        Public Shared Function Base64_Decode(ByVal Source As String, ByVal Encoding As System.Text.Encoding) As String
            Return Encoding.GetString(Convert.FromBase64String(Source))
        End Function



        ''' <summary>
        ''' Base64编码（用于URL的改进Base64编码，编码结果的字符串中包含字母A-Z，a-z，数字0-9，符号-_=）
        ''' </summary>
        ''' <param name="Source">要编码的字符串</param>
        ''' <returns>编码后的结果字符串</returns>
        ''' <remarks></remarks>
        Public Shared Function Base64_URL_Encode(ByVal Source As String) As String
            Return Convert.ToBase64String(System.Text.Encoding.UTF8.GetBytes(Source)).Replace("+", "-").Replace("/", "_")
        End Function
        ''' <summary>
        ''' Base64编码（用于URL的改进Base64编码，编码结果的字符串中包含字母A-Z，a-z，数字0-9，符号-_=）
        ''' </summary>
        ''' <param name="Source">要编码的字符串</param>
        ''' <param name="Encoding">使用特定的字符编码（默认UTF-8）</param>
        ''' <returns>编码后的结果字符串</returns>
        ''' <remarks></remarks>
        Public Shared Function Base64_URL_Encode(ByVal Source As String, ByVal Encoding As System.Text.Encoding) As String
            Return Convert.ToBase64String(Encoding.GetBytes(Source)).Replace("+", "-").Replace("/", "_")
        End Function

        ''' <summary>
        ''' Base64解码（用于URL的改进Base64解码，要解码的字符串可以包含字母A-Z，a-z，数字0-9，符号-_=）
        ''' </summary>
        ''' <param name="Source">要解码的字符串</param>
        ''' <returns>解码后的结果字符串</returns>
        ''' <remarks></remarks>
        Public Shared Function Base64_URL_Decode(ByVal Source As String) As String
            Return System.Text.Encoding.UTF8.GetString(Convert.FromBase64String(Source.Replace("-", "+").Replace("_", "/")))
        End Function
        ''' <summary>
        ''' Base64解码（用于URL的改进Base64解码，要解码的字符串可以包含字母A-Z，a-z，数字0-9，符号-_=）
        ''' </summary>
        ''' <param name="Source">要解码的字符串</param>
        ''' <param name="Encoding">使用特定的字符编码（默认UTF-8）</param>
        ''' <returns>解码后的结果字符串</returns>
        ''' <remarks></remarks>
        Public Shared Function Base64_URL_Decode(ByVal Source As String, ByVal Encoding As System.Text.Encoding) As String
            Return Encoding.GetString(Convert.FromBase64String(Source.Replace("-", "+").Replace("_", "/")))
        End Function


        ''' <summary>
        ''' MD5加密（摘要结果为32位16进制字符串）
        ''' </summary>
        ''' <param name="Source">要加密的Byte数组</param>
        ''' <param name="ToUpper">是否将结果转换为大写字母形式</param>
        ''' <returns>加密后的结果字符串</returns>
        ''' <remarks></remarks>
        Public Shared Function MD5_Encode(ByVal Source As Byte(), Optional ByVal ToUpper As Boolean = True) As String
            If ToUpper Then
                Return BitConverter.ToString((New System.Security.Cryptography.MD5CryptoServiceProvider).ComputeHash(Source)).Replace("-", "").ToUpper()
            Else
                Return BitConverter.ToString((New System.Security.Cryptography.MD5CryptoServiceProvider).ComputeHash(Source)).Replace("-", "").ToLower()
            End If
        End Function
        ''' <summary>
        ''' MD5加密（摘要结果为32位16进制字符串）
        ''' </summary>
        ''' <param name="Source">要加密的字符串</param>
        ''' <param name="ToUpper">是否将结果转换为大写字母形式</param>
        ''' <returns>加密后的结果字符串</returns>
        ''' <remarks></remarks>
        Public Shared Function MD5_Encode(ByVal Source As String, Optional ByVal ToUpper As Boolean = True) As String
            If ToUpper Then
                Return System.Web.Security.FormsAuthentication.HashPasswordForStoringInConfigFile(Source, "MD5").ToUpper()
            Else
                Return System.Web.Security.FormsAuthentication.HashPasswordForStoringInConfigFile(Source, "MD5").ToLower()
            End If
        End Function



        ''' <summary>
        ''' SHA1加密（摘要结果为40位16进制字符串）
        ''' </summary>
        ''' <param name="Source">要加密的Byte数组</param>
        ''' <param name="ToUpper">是否将结果转换为大写字母形式</param>
        ''' <returns>加密后的结果字符串</returns>
        ''' <remarks></remarks>
        Public Shared Function SHA1_Encode(ByVal Source As Byte(), Optional ByVal ToUpper As Boolean = True) As String
            If ToUpper Then
                Return BitConverter.ToString((New System.Security.Cryptography.SHA1CryptoServiceProvider).ComputeHash(Source)).Replace("-", "").ToUpper()
            Else
                Return BitConverter.ToString((New System.Security.Cryptography.SHA1CryptoServiceProvider).ComputeHash(Source)).Replace("-", "").ToLower()
            End If
        End Function
        ''' <summary>
        ''' SHA1加密（摘要结果为40位16进制字符串）
        ''' </summary>
        ''' <param name="Source">要加密的字符串</param>
        ''' <param name="ToUpper">是否将结果转换为大写字母形式</param>
        ''' <returns>加密后的结果字符串</returns>
        ''' <remarks></remarks>
        Public Shared Function SHA1_Encode(ByVal Source As String, Optional ByVal ToUpper As Boolean = True) As String
            If ToUpper Then
                Return BitConverter.ToString((New System.Security.Cryptography.SHA1CryptoServiceProvider).ComputeHash(System.Text.Encoding.UTF8.GetBytes(Source))).Replace("-", "").ToUpper()
            Else
                Return BitConverter.ToString((New System.Security.Cryptography.SHA1CryptoServiceProvider).ComputeHash(System.Text.Encoding.UTF8.GetBytes(Source))).Replace("-", "").ToLower()
            End If
        End Function



        ''' <summary>
        ''' SHA256加密（摘要结果为64位16进制字符串）
        ''' </summary>
        ''' <param name="Source">要加密的Byte数组</param>
        ''' <param name="ToUpper">是否将结果转换为大写字母形式</param>
        ''' <returns>加密后的结果字符串</returns>
        ''' <remarks></remarks>
        Public Shared Function SHA256_Encode(ByVal Source As Byte(), Optional ByVal ToUpper As Boolean = True) As String
            If ToUpper Then
                Return BitConverter.ToString((New System.Security.Cryptography.SHA256Managed).ComputeHash(Source)).Replace("-", "").ToUpper()
            Else
                Return BitConverter.ToString((New System.Security.Cryptography.SHA256Managed).ComputeHash(Source)).Replace("-", "").ToLower()
            End If
        End Function
        ''' <summary>
        ''' SHA256加密（摘要结果为64位16进制字符串）
        ''' </summary>
        ''' <param name="Source">要加密的字符串</param>
        ''' <param name="ToUpper">是否将结果转换为大写字母形式</param>
        ''' <returns>加密后的结果字符串</returns>
        ''' <remarks></remarks>
        Public Shared Function SHA256_Encode(ByVal Source As String, Optional ByVal ToUpper As Boolean = True) As String
            If ToUpper Then
                Return BitConverter.ToString((New System.Security.Cryptography.SHA256Managed).ComputeHash(System.Text.Encoding.UTF8.GetBytes(Source))).Replace("-", "").ToUpper()
            Else
                Return BitConverter.ToString((New System.Security.Cryptography.SHA256Managed).ComputeHash(System.Text.Encoding.UTF8.GetBytes(Source))).Replace("-", "").ToLower()
            End If
        End Function



        ''' <summary>
        ''' SHA384加密（摘要结果为96位16进制字符串）
        ''' </summary>
        ''' <param name="Source">要加密的Byte数组</param>
        ''' <param name="ToUpper">是否将结果转换为大写字母形式</param>
        ''' <returns>加密后的结果字符串</returns>
        ''' <remarks></remarks>
        Public Shared Function SHA384_Encode(ByVal Source As Byte(), Optional ByVal ToUpper As Boolean = True) As String
            If ToUpper Then
                Return BitConverter.ToString((New System.Security.Cryptography.SHA384Managed).ComputeHash(Source)).Replace("-", "").ToUpper()
            Else
                Return BitConverter.ToString((New System.Security.Cryptography.SHA384Managed).ComputeHash(Source)).Replace("-", "").ToLower()
            End If
        End Function
        ''' <summary>
        ''' SHA384加密（摘要结果为96位16进制字符串）
        ''' </summary>
        ''' <param name="Source">要加密的字符串</param>
        ''' <param name="ToUpper">是否将结果转换为大写字母形式</param>
        ''' <returns>加密后的结果字符串</returns>
        ''' <remarks></remarks>
        Public Shared Function SHA384_Encode(ByVal Source As String, Optional ByVal ToUpper As Boolean = True) As String
            If ToUpper Then
                Return BitConverter.ToString((New System.Security.Cryptography.SHA384Managed).ComputeHash(System.Text.Encoding.UTF8.GetBytes(Source))).Replace("-", "").ToUpper()
            Else
                Return BitConverter.ToString((New System.Security.Cryptography.SHA384Managed).ComputeHash(System.Text.Encoding.UTF8.GetBytes(Source))).Replace("-", "").ToLower()
            End If
        End Function



        ''' <summary>
        ''' SHA512加密（摘要结果为128位16进制字符串）
        ''' </summary>
        ''' <param name="Source">要加密的Byte数组</param>
        ''' <param name="ToUpper">是否将结果转换为大写字母形式</param>
        ''' <returns>加密后的结果字符串</returns>
        ''' <remarks></remarks>
        Public Shared Function SHA512_Encode(ByVal Source As Byte(), Optional ByVal ToUpper As Boolean = True) As String
            If ToUpper Then
                Return BitConverter.ToString((New System.Security.Cryptography.SHA512Managed).ComputeHash(Source)).Replace("-", "").ToUpper()
            Else
                Return BitConverter.ToString((New System.Security.Cryptography.SHA512Managed).ComputeHash(Source)).Replace("-", "").ToLower()
            End If
        End Function
        ''' <summary>
        ''' SHA512加密（摘要结果为128位16进制字符串）
        ''' </summary>
        ''' <param name="Source">要加密的字符串</param>
        ''' <param name="ToUpper">是否将结果转换为大写字母形式</param>
        ''' <returns>加密后的结果字符串</returns>
        ''' <remarks></remarks>
        Public Shared Function SHA512_Encode(ByVal Source As String, Optional ByVal ToUpper As Boolean = True) As String
            If ToUpper Then
                Return BitConverter.ToString((New System.Security.Cryptography.SHA512Managed).ComputeHash(System.Text.Encoding.UTF8.GetBytes(Source))).Replace("-", "").ToUpper()
            Else
                Return BitConverter.ToString((New System.Security.Cryptography.SHA512Managed).ComputeHash(System.Text.Encoding.UTF8.GetBytes(Source))).Replace("-", "").ToLower()
            End If
        End Function



        ''' <summary>
        ''' DES加密
        ''' </summary>
        ''' <param name="Source">要加密的Byte数组</param>
        ''' <param name="SecretKey">加密密钥（8的整数倍字节数的字符串）</param>
        ''' <returns>加密后的结果Byte数组（失败返回空Byte数组）</returns>
        ''' <remarks></remarks>
        Public Shared Function DES_Encode(ByVal Source As Byte(), ByVal SecretKey As String) As Byte()
            Dim DES As New System.Security.Cryptography.DESCryptoServiceProvider()
            Try
                DES.Key = System.Text.Encoding.UTF8.GetBytes(SecretKey)
                DES.IV = System.Text.Encoding.UTF8.GetBytes(SecretKey)
            Catch ex As System.ArgumentException '加密失败（通常是由于SecretKey的字节数不是8的倍数）
                Return New Byte() {}
            End Try
            Dim MemoryStream As New System.IO.MemoryStream()
            Dim CryptoStream As New System.Security.Cryptography.CryptoStream(MemoryStream, DES.CreateEncryptor(), System.Security.Cryptography.CryptoStreamMode.Write)
            CryptoStream.Write(Source, 0, Source.Length)
            CryptoStream.FlushFinalBlock()
            Return MemoryStream.ToArray()
        End Function
        ''' <summary>
        ''' DES加密
        ''' </summary>
        ''' <param name="Source">要加密的字符串</param>
        ''' <param name="SecretKey">加密密钥（8的整数倍字节数的字符串）</param>
        ''' <returns>加密后的结果Byte数组（失败返回空Byte数组）</returns>
        ''' <remarks></remarks>
        Public Shared Function DES_Encode(ByVal Source As String, ByVal SecretKey As String) As Byte()
            Dim DES As New System.Security.Cryptography.DESCryptoServiceProvider()
            Try
                DES.Key = System.Text.Encoding.UTF8.GetBytes(SecretKey)
                DES.IV = System.Text.Encoding.UTF8.GetBytes(SecretKey)
            Catch ex As System.ArgumentException '加密失败（通常是由于SecretKey的字节数不是8的倍数）
                Return New Byte() {}
            End Try
            Dim MemoryStream As New System.IO.MemoryStream()
            Dim CryptoStream As New System.Security.Cryptography.CryptoStream(MemoryStream, DES.CreateEncryptor(), System.Security.Cryptography.CryptoStreamMode.Write)
            CryptoStream.Write(System.Text.Encoding.UTF8.GetBytes(Source), 0, System.Text.Encoding.UTF8.GetBytes(Source).Length)
            CryptoStream.FlushFinalBlock()
            Return MemoryStream.ToArray()
        End Function

        ''' <summary>
        ''' DES解密
        ''' </summary>
        ''' <param name="Source">要解密的Byte数组</param>
        ''' <param name="SecretKey">解密密钥（8的整数倍字节数的字符串）</param>
        ''' <returns>解密后的结果Byte数组（失败返回空Byte数组）</returns>
        ''' <remarks></remarks>
        Public Shared Function DES_Decode(ByVal Source As Byte(), ByVal SecretKey As String) As Byte()
            Dim DES As New System.Security.Cryptography.DESCryptoServiceProvider()
            Try
                DES.Key = System.Text.Encoding.UTF8.GetBytes(SecretKey)
                DES.IV = System.Text.Encoding.UTF8.GetBytes(SecretKey)
            Catch ex As System.ArgumentException '解密失败（通常是由于SecretKey的字节数不是8的倍数）
                Return New Byte() {}
            End Try
            Dim MemoryStream As New System.IO.MemoryStream()
            Dim CryptoStream As New System.Security.Cryptography.CryptoStream(MemoryStream, DES.CreateDecryptor, System.Security.Cryptography.CryptoStreamMode.Write)
            CryptoStream.Write(Source, 0, Source.Length)
            Try
                CryptoStream.FlushFinalBlock()
            Catch ex As System.Security.Cryptography.CryptographicException '解密失败（通常是由于SecretKey不匹配）
                Return New Byte() {}
            End Try
            Return MemoryStream.ToArray()
        End Function
        ''' <summary>
        ''' DES解密
        ''' </summary>
        ''' <param name="Source">要解密的Byte数组</param>
        ''' <param name="SecretKey">解密密钥（8的整数倍字节数的字符串）</param>
        ''' <returns>解密后的结果字符串（失败返回空字符串）</returns>
        ''' <remarks></remarks>
        Public Shared Function DES_Decode_String(ByVal Source As Byte(), ByVal SecretKey As String) As String
            Dim DES As New System.Security.Cryptography.DESCryptoServiceProvider()
            Try
                DES.Key = System.Text.Encoding.UTF8.GetBytes(SecretKey)
                DES.IV = System.Text.Encoding.UTF8.GetBytes(SecretKey)
            Catch ex As System.ArgumentException '解密失败（通常是由于SecretKey的字节数不是8的倍数）
                Return ""
            End Try
            Dim MemoryStream As New System.IO.MemoryStream()
            Dim CryptoStream As New System.Security.Cryptography.CryptoStream(MemoryStream, DES.CreateDecryptor, System.Security.Cryptography.CryptoStreamMode.Write)
            CryptoStream.Write(Source, 0, Source.Length)
            Try
                CryptoStream.FlushFinalBlock()
            Catch ex As System.Security.Cryptography.CryptographicException '解密失败（通常是由于SecretKey不匹配）
                Return ""
            End Try
            Return System.Text.Encoding.UTF8.GetString(MemoryStream.ToArray())
        End Function



        ''' <summary>
        ''' RSA加密（使用本函数库内置的密钥）
        ''' </summary>
        ''' <param name="Source">要加密的字符串</param>
        ''' <returns>加密后的结果Byte数组</returns>
        ''' <remarks></remarks>
        Public Shared Function RSA_Encode(ByVal Source As String) As Byte()
            Dim RSA As New System.Security.Cryptography.RSACryptoServiceProvider
            RSA.FromXmlString("<RSAKeyValue><Modulus>pZGIiC3CxVYpTJ4dLylSy2TLXW+R9EyRZ39ekSosvRKf7iPuz4oPlHqjssh4Glbj/vTUIMFzHFC/9zC56GggNLfZBjh6fc3adq5cXGKlU74kAyM2z7gdYlUHtLT/GwDp4YcQKeSb9GjcvsXbUp0mrzI/axzueLIqK+R07rnv3yc=</Modulus><Exponent>AQAB</Exponent></RSAKeyValue>")
            Return RSA.Encrypt(System.Text.Encoding.UTF8.GetBytes(Source), True)
        End Function

        ''' <summary>
        ''' RSA解密（使用本函数库内置的密钥）
        ''' </summary>
        ''' <param name="Source">要解密的Byte数组</param>
        ''' <returns>解密后的结果字符串（失败返回空字符串）</returns>
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



        ''' <summary>
        ''' 二进制形式化（只由0和1组成的，8的整数倍位数的2进制字符串）
        ''' </summary>
        ''' <param name="Source">要编码的Byte数组</param>
        ''' <param name="ContainSpace">是否将结果每8位以空格隔开</param>
        ''' <returns>编码后的结果字符串</returns>
        ''' <remarks></remarks>
        Public Shared Function Binary_Encode(ByVal Source As Byte(), Optional ByVal ContainSpace As Boolean = True) As String
            Dim Result As String = ""
            If ContainSpace Then
                For Each SourceByte In Source
                    Result = Result & Convert.ToString(SourceByte, 2).PadLeft(8, "0") & " "
                Next
            Else
                For Each SourceByte In Source
                    Result = Result & Convert.ToString(SourceByte, 2).PadLeft(8, "0")
                Next
            End If
            Return Result.Trim()
        End Function
        ''' <summary>
        ''' 二进制形式化（只由0和1组成的，8的整数倍位数的2进制字符串）
        ''' </summary>
        ''' <param name="Source">要编码的字符串</param>
        ''' <param name="ContainSpace">是否将结果每8位以空格隔开</param>
        ''' <returns>编码后的结果字符串</returns>
        ''' <remarks></remarks>
        Public Shared Function Binary_Encode(ByVal Source As String, Optional ByVal ContainSpace As Boolean = True) As String
            Dim Result As String = ""
            If ContainSpace Then
                For Each SourceByte In System.Text.Encoding.UTF8.GetBytes(Source)
                    Result = Result & Convert.ToString(SourceByte, 2).PadLeft(8, "0") & " "
                Next
            Else
                For Each SourceByte In System.Text.Encoding.UTF8.GetBytes(Source)
                    Result = Result & Convert.ToString(SourceByte, 2).PadLeft(8, "0")
                Next
            End If
            Return Result.Trim()
        End Function
        ''' <summary>
        ''' 二进制形式化（只由0和1组成的，8的整数倍位数的2进制字符串）
        ''' </summary>
        ''' <param name="Source">要编码的字符串</param>
        ''' <param name="Encoding">使用特定的字符编码（默认UTF-8）</param>
        ''' <param name="ContainSpace">是否将结果每8位以空格隔开</param>
        ''' <returns>编码后的结果字符串</returns>
        ''' <remarks></remarks>
        Public Shared Function Binary_Encode(ByVal Source As String, ByVal Encoding As System.Text.Encoding, Optional ByVal ContainSpace As Boolean = True) As String
            Dim Result As String = ""
            If ContainSpace Then
                For Each SourceByte In Encoding.GetBytes(Source)
                    Result = Result & Convert.ToString(SourceByte, 2).PadLeft(8, "0") & " "
                Next
            Else
                For Each SourceByte In Encoding.GetBytes(Source)
                    Result = Result & Convert.ToString(SourceByte, 2).PadLeft(8, "0")
                Next
            End If
            Return Result.Trim()
        End Function

        ''' <summary>
        ''' 二进制反形式化（将只由0和1组成的，8的整数倍位数的2进制字符串，转换为原始意义的Byte数组）
        ''' </summary>
        ''' <param name="Source">要解码的字符串</param>
        ''' <returns>解码后的结果Byte数组</returns>
        ''' <remarks></remarks>
        Public Shared Function Binary_Decode(ByVal Source As String) As Byte()
            Source = Source.Replace(" ", "")
            Dim Result(Source.Length / 8 - 1) As Byte
            For I = 0 To Source.Length / 8 - 1
                Dim SourceByte As String = Source.Substring(I * 8, 8)
                Result(I) = Convert.ToByte(SourceByte, 2)
            Next
            Return Result
        End Function
        ''' <summary>
        ''' 二进制反形式化（将只由0和1组成的，8的整数倍位数的2进制字符串，转换为原始意义的字符串）
        ''' </summary>
        ''' <param name="Source">要解码的字符串</param>
        ''' <returns>解码后的结果字符串</returns>
        ''' <remarks></remarks>
        Public Shared Function Binary_Decode_String(ByVal Source As String) As String
            Source = Source.Replace(" ", "")
            Dim Result(Source.Length / 8 - 1) As Byte
            For I = 0 To Source.Length / 8 - 1
                Dim SourceByte As String = Source.Substring(I * 8, 8)
                Result(I) = Convert.ToByte(SourceByte, 2)
            Next
            Return System.Text.Encoding.UTF8.GetString(Result)
        End Function
        ''' <summary>
        ''' 二进制反形式化（将只由0和1组成的，8的整数倍位数的2进制字符串，转换为原始意义的字符串）
        ''' </summary>
        ''' <param name="Source">要解码的字符串</param>
        ''' <param name="Encoding">使用特定的字符编码（默认UTF-8）</param>
        ''' <returns>解码后的结果字符串</returns>
        ''' <remarks></remarks>
        Public Shared Function Binary_Decode_String(ByVal Source As String, ByVal Encoding As System.Text.Encoding) As String
            Source = Source.Replace(" ", "")
            Dim Result(Source.Length / 8 - 1) As Byte
            For I = 0 To Source.Length / 8 - 1
                Dim SourceByte As String = Source.Substring(I * 8, 8)
                Result(I) = Convert.ToByte(SourceByte, 2)
            Next
            Return Encoding.GetString(Result)
        End Function



        ''' <summary>
        ''' 二进制非处理（将二进制形式的0替换为1，1替换为0）
        ''' </summary>
        ''' <param name="Source">要编码的Byte数组</param>
        ''' <returns>编码后的结果Byte数组</returns>
        ''' <remarks></remarks>
        Public Shared Function Binary_Not(ByVal Source As Byte()) As Byte()
            Dim Temp As String = ""
            For Each SourceByte In Source
                Temp = Temp & Convert.ToString(SourceByte, 2).PadLeft(8, "0")
            Next
            Temp = Temp.Replace("0", "2")
            Temp = Temp.Replace("1", "0")
            Temp = Temp.Replace("2", "1")
            Dim Result(Temp.Length / 8 - 1) As Byte
            For I = 0 To Temp.Length / 8 - 1
                Dim SourceByte As String = Temp.Substring(I * 8, 8).Remove(0, Temp.Substring(I * 8, 8).IndexOf("1"))
                Result(I) = Convert.ToByte(SourceByte, 2)
            Next
            Return Result
        End Function



        ''' <summary>
        ''' 十六进制编码（由0-F组成的，2的整数倍位数的16进制字符串）
        ''' </summary>
        ''' <param name="Source">要编码的Byte数组</param>
        ''' <param name="ToUpper">是否将结果转换为大写字母形式</param>
        ''' <returns>编码后的结果字符串</returns>
        ''' <remarks></remarks>
        Public Shared Function Hex_Encode(ByVal Source As Byte(), Optional ByVal ToUpper As Boolean = True) As String
            If ToUpper Then
                Return BitConverter.ToString(Source).Replace("-", "").ToUpper()
            Else
                Return BitConverter.ToString(Source).Replace("-", "").ToLower()
            End If
        End Function
        ''' <summary>
        ''' 十六进制编码（由0-F组成的，2的整数倍位数的16进制字符串）
        ''' </summary>
        ''' <param name="Source">要编码的字符串</param>
        ''' <param name="ToUpper">是否将结果转换为大写字母形式</param>
        ''' <returns>编码后的结果字符串</returns>
        ''' <remarks></remarks>
        Public Shared Function Hex_Encode(ByVal Source As String, Optional ByVal ToUpper As Boolean = True) As String
            If ToUpper Then
                Return BitConverter.ToString(System.Text.Encoding.UTF8.GetBytes(Source)).Replace("-", "").ToUpper()
            Else
                Return BitConverter.ToString(System.Text.Encoding.UTF8.GetBytes(Source)).Replace("-", "").ToLower()
            End If
        End Function
        ''' <summary>
        ''' 十六进制编码（由0-F组成的，2的整数倍位数的16进制字符串）
        ''' </summary>
        ''' <param name="Source">要编码的字符串</param>
        ''' <param name="Encoding">使用特定的字符编码（默认UTF-8）</param>
        ''' <param name="ToUpper">是否将结果转换为大写字母形式</param>
        ''' <returns>编码后的结果字符串</returns>
        ''' <remarks></remarks>
        Public Shared Function Hex_Encode(ByVal Source As String, ByVal Encoding As System.Text.Encoding, Optional ByVal ToUpper As Boolean = True) As String
            If ToUpper Then
                Return BitConverter.ToString(Encoding.GetBytes(Source)).Replace("-", "").ToUpper()
            Else
                Return BitConverter.ToString(Encoding.GetBytes(Source)).Replace("-", "").ToLower()
            End If
        End Function

        ''' <summary>
        ''' 十六进制解码（将由0-F组成的，2的整数倍位数的16进制字符串，转换为原始意义的Byte数组）
        ''' </summary>
        ''' <param name="Source">要解码的字符串</param>
        ''' <returns>解码后的结果Byte数组</returns>
        ''' <remarks></remarks>
        Public Shared Function Hex_Decode(ByVal Source As String) As Byte()
            Dim Result(Source.Length / 2 - 1) As Byte
            For I = 0 To Source.Length / 2 - 1
                Dim SourceByte As String = Source.Substring(I * 2, 2)
                Result(I) = Convert.ToByte(SourceByte, 16)
            Next
            Return Result
        End Function
        ''' <summary>
        ''' 十六进制解码（将由0-F组成的，2的整数倍位数的16进制字符串，转换为原始意义的字符串）
        ''' </summary>
        ''' <param name="Source">要解码的字符串</param>
        ''' <returns>解码后的结果字符串</returns>
        ''' <remarks></remarks>
        Public Shared Function Hex_Decode_String(ByVal Source As String) As String
            Dim Result(Source.Length / 2 - 1) As Byte
            For I = 0 To Source.Length / 2 - 1
                Dim SourceByte As String = Source.Substring(I * 2, 2)
                Result(I) = Convert.ToByte(SourceByte, 16)
            Next
            Return System.Text.Encoding.UTF8.GetString(Result)
        End Function
        ''' <summary>
        ''' 十六进制解码（将由0-F组成的，2的整数倍位数的16进制字符串，转换为原始意义的字符串）
        ''' </summary>
        ''' <param name="Source">要解码的字符串</param>
        ''' <param name="Encoding">使用特定的字符编码（默认UTF-8）</param>
        ''' <returns>解码后的结果字符串</returns>
        ''' <remarks></remarks>
        Public Shared Function Hex_Decode_String(ByVal Source As String, ByVal Encoding As System.Text.Encoding) As String
            Dim Result(Source.Length / 2 - 1) As Byte
            For I = 0 To Source.Length / 2 - 1
                Dim SourceByte As String = Source.Substring(I * 2, 2)
                Result(I) = Convert.ToByte(SourceByte, 16)
            Next
            Return System.Text.Encoding.UTF8.GetString(Result)
        End Function



        ''' <summary>
        ''' 根据文件的名称、大小、哈希值，生成文件的ED2K下载链接
        ''' </summary>
        ''' <param name="FileName">文件名称（不必准确）</param>
        ''' <param name="FileLength">文件大小（必须准确）</param>
        ''' <param name="FileHash">文件哈希值（必须准确，32位字符串）</param>
        ''' <returns>生成的ED2K链接结果字符串（失败返回空字符串）</returns>
        ''' <remarks></remarks>
        Public Shared Function Generate_ED2K(ByVal FileName As String, ByVal FileLength As Integer, ByVal FileHash As String) As String
            If FileHash.Length <> 32 Then
                Return ""
            End If
            FileName = System.Web.HttpUtility.UrlEncode(FileName, System.Text.Encoding.UTF8)
            Return "ed2k://|file|" & FileName & "|" & FileLength & "|" & FileHash & "|/"
        End Function

    End Class

End Namespace