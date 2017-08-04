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
        ''' <param name="Source">要编码的Int32值</param>
        ''' <param name="ToUpper">是否将结果转换为大写字母形式</param>
        ''' <returns>编码后的结果字符串</returns>
        ''' <remarks></remarks>
        Public Shared Function Hex_Encode(ByVal Source As Int32, Optional ByVal ToUpper As Boolean = True) As String
            If ToUpper Then
                Return Source.ToString("X8")
            Else
                Return Source.ToString("x8")
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
        ''' 十六进制解码（将由0-F组成的，2的整数倍位数的16进制字符串，转换为原始意义的Int32值）
        ''' </summary>
        ''' <param name="Source">要解码的字符串</param>
        ''' <returns>解码后的结果Int32值</returns>
        ''' <remarks></remarks>
        Public Shared Function Hex_Decode_Int32(ByVal Source As String) As String
            Return Int32.Parse(Source, System.Globalization.NumberStyles.HexNumber)
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



        Public Class MD4_Hash_Algorithm
            Private Const BlockSize As Integer = 512 / 8
            Private Context As UInt32() = New UInt32() {&H67452301UI, &HEFCDAB89UI, &H98BADCFEUI, &H10325476UI}
            Private ProcessedCount As Integer = 0
            Private InputBuffer(BlockSize - 1) As Byte
            Private WorkBuffer(BlockSize / 4 - 1) As UInt32
            Public Sub New(ByVal InputBytes As Byte())
                Update(InputBytes)
            End Sub
            Private Sub Update(ByVal InputBytes As Byte())
                Dim UnhashedBufferLength As Integer = ProcessedCount Mod BlockSize
                Dim PartLength As Integer = BlockSize - UnhashedBufferLength
                Dim Index As Integer = 0
                If InputBytes.Length >= PartLength Then
                    Array.Copy(InputBytes, Index, InputBuffer, UnhashedBufferLength, PartLength)
                    Transform(InputBuffer, 0)
                    Index = PartLength
                    Do While (Index + BlockSize - 1) < InputBytes.Length
                        Transform(InputBytes, Index)
                        Index += BlockSize
                    Loop
                    UnhashedBufferLength = 0
                End If
                If Index < InputBytes.Length Then
                    Array.Copy(InputBytes, Index, InputBuffer, UnhashedBufferLength, InputBytes.Length - Index)
                End If
                ProcessedCount += InputBytes.Length
            End Sub
            Private Sub Transform(ByRef Block As Byte(), ByVal Offset As Integer)
                For I = 0 To 15
                    If Offset >= Block.Length Then
                        Exit For
                    End If
                    WorkBuffer(I) = (CType(Block(Offset + 0), UInt32) And Byte.MaxValue)
                    WorkBuffer(I) = WorkBuffer(I) Or ((CType(Block(Offset + 1), UInt32) And Byte.MaxValue) << 8)
                    WorkBuffer(I) = WorkBuffer(I) Or ((CType(Block(Offset + 2), UInt32) And Byte.MaxValue) << 16)
                    WorkBuffer(I) = WorkBuffer(I) Or ((CType(Block(Offset + 3), UInt32) And Byte.MaxValue) << 24)
                    Offset += 4
                Next
                Dim A As UInt32 = Context(0)
                Dim B As UInt32 = Context(1)
                Dim C As UInt32 = Context(2)
                Dim D As UInt32 = Context(3)
                A = Round1(A, B, C, D, WorkBuffer(0), 3)
                D = Round1(D, A, B, C, WorkBuffer(1), 7)
                C = Round1(C, D, A, B, WorkBuffer(2), 11)
                B = Round1(B, C, D, A, WorkBuffer(3), 19)
                A = Round1(A, B, C, D, WorkBuffer(4), 3)
                D = Round1(D, A, B, C, WorkBuffer(5), 7)
                C = Round1(C, D, A, B, WorkBuffer(6), 11)
                B = Round1(B, C, D, A, WorkBuffer(7), 19)
                A = Round1(A, B, C, D, WorkBuffer(8), 3)
                D = Round1(D, A, B, C, WorkBuffer(9), 7)
                C = Round1(C, D, A, B, WorkBuffer(10), 11)
                B = Round1(B, C, D, A, WorkBuffer(11), 19)
                A = Round1(A, B, C, D, WorkBuffer(12), 3)
                D = Round1(D, A, B, C, WorkBuffer(13), 7)
                C = Round1(C, D, A, B, WorkBuffer(14), 11)
                B = Round1(B, C, D, A, WorkBuffer(15), 19)
                A = Round2(A, B, C, D, WorkBuffer(0), 3)
                D = Round2(D, A, B, C, WorkBuffer(4), 5)
                C = Round2(C, D, A, B, WorkBuffer(8), 9)
                B = Round2(B, C, D, A, WorkBuffer(12), 13)
                A = Round2(A, B, C, D, WorkBuffer(1), 3)
                D = Round2(D, A, B, C, WorkBuffer(5), 5)
                C = Round2(C, D, A, B, WorkBuffer(9), 9)
                B = Round2(B, C, D, A, WorkBuffer(13), 13)
                A = Round2(A, B, C, D, WorkBuffer(2), 3)
                D = Round2(D, A, B, C, WorkBuffer(6), 5)
                C = Round2(C, D, A, B, WorkBuffer(10), 9)
                B = Round2(B, C, D, A, WorkBuffer(14), 13)
                A = Round2(A, B, C, D, WorkBuffer(3), 3)
                D = Round2(D, A, B, C, WorkBuffer(7), 5)
                C = Round2(C, D, A, B, WorkBuffer(11), 9)
                B = Round2(B, C, D, A, WorkBuffer(15), 13)
                A = Round3(A, B, C, D, WorkBuffer(0), 3)
                D = Round3(D, A, B, C, WorkBuffer(8), 9)
                C = Round3(C, D, A, B, WorkBuffer(4), 11)
                B = Round3(B, C, D, A, WorkBuffer(12), 15)
                A = Round3(A, B, C, D, WorkBuffer(2), 3)
                D = Round3(D, A, B, C, WorkBuffer(10), 9)
                C = Round3(C, D, A, B, WorkBuffer(6), 11)
                B = Round3(B, C, D, A, WorkBuffer(14), 15)
                A = Round3(A, B, C, D, WorkBuffer(1), 3)
                D = Round3(D, A, B, C, WorkBuffer(9), 9)
                C = Round3(C, D, A, B, WorkBuffer(5), 11)
                B = Round3(B, C, D, A, WorkBuffer(13), 15)
                A = Round3(A, B, C, D, WorkBuffer(3), 3)
                D = Round3(D, A, B, C, WorkBuffer(11), 9)
                C = Round3(C, D, A, B, WorkBuffer(7), 11)
                B = Round3(B, C, D, A, WorkBuffer(15), 15)
                Context(0) = &HFFFFFFFFUI And (Context(0) + Convert.ToInt64(A))
                Context(1) = &HFFFFFFFFUI And (Context(1) + Convert.ToInt64(B))
                Context(2) = &HFFFFFFFFUI And (Context(2) + Convert.ToInt64(C))
                Context(3) = &HFFFFFFFFUI And (Context(3) + Convert.ToInt64(D))
            End Sub
            Private Function Round1(ByVal P1 As UInt32, ByVal P2 As UInt32, ByVal P3 As UInt32, ByVal P4 As UInt32, ByVal X As UInt32, ByVal S As Integer) As UInt32
                Dim T As UInt32 = &HFFFFFFFFUI And (&HFFFFFFFFUI And (Convert.ToInt64(P1) + ((P2 And P3) Or ((Not P2) And P4))) + Convert.ToInt64(X))
                Return T << S Or T >> (32 - S)
            End Function
            Private Function Round2(ByVal P1 As UInt32, ByVal P2 As UInt32, ByVal P3 As UInt32, ByVal P4 As UInt32, ByVal X As UInt32, ByVal S As Integer) As UInt32
                Dim T As UInt32 = &HFFFFFFFFUI And (&HFFFFFFFFUI And (Convert.ToInt64(P1) + ((P2 And (P3 Or P4)) Or (P3 And P4))) + Convert.ToInt64(X) + &H5A827999UI)
                Return T << S Or T >> (32 - S)
            End Function
            Private Function Round3(ByVal P1 As UInt32, ByVal P2 As UInt32, ByVal P3 As UInt32, ByVal P4 As UInt32, ByVal X As UInt32, ByVal S As Integer) As UInt32
                Dim T As UInt32 = &HFFFFFFFFUI And (&HFFFFFFFFUI And (Convert.ToInt64(P1) + (P2 Xor P3 Xor P4)) + Convert.ToInt64(X) + &H6ED9EBA1UI)
                Return T << S Or T >> (32 - S)
            End Function
            Public Function DigestResult() As Byte()
                Dim UnhashedBufferLength As Integer = CType(ProcessedCount Mod BlockSize, Integer)
                Dim PaddingLength As Integer
                If UnhashedBufferLength < 56 Then
                    PaddingLength = 56 - UnhashedBufferLength
                Else
                    PaddingLength = 120 - UnhashedBufferLength
                End If
                Dim Tail(PaddingLength + 8 - 1) As Byte
                Tail(0) = CType(128, Byte)
                BitConverter.GetBytes(ProcessedCount * 8).CopyTo(Tail, PaddingLength)
                Update(Tail)
                Dim Result(16 - 1) As Byte
                For I As Integer = 0 To 3
                    BitConverter.GetBytes(Context(I)).CopyTo(Result, I * 4)
                Next
                Return Result
            End Function
        End Class
        ''' <summary>
        ''' MD4加密（摘要结果为32位16进制字符串）
        ''' </summary>
        ''' <param name="Source">要加密的Byte数组</param>
        ''' <param name="ToUpper">是否将结果转换为大写字母形式</param>
        ''' <returns>加密后的结果字符串</returns>
        ''' <remarks></remarks>
        Public Shared Function MD4_Encode(ByVal Source As Byte(), Optional ByVal ToUpper As Boolean = True) As String
            Dim MD4 As MD4_Hash_Algorithm = New MD4_Hash_Algorithm(Source)
            If ToUpper Then
                Return BitConverter.ToString(MD4.DigestResult()).Replace("-", "").ToUpper()
            Else
                Return BitConverter.ToString(MD4.DigestResult()).Replace("-", "").ToLower()
            End If
        End Function
        ''' <summary>
        ''' MD4加密（摘要结果为32位16进制字符串）
        ''' </summary>
        ''' <param name="Source">要加密的字符串</param>
        ''' <param name="ToUpper">是否将结果转换为大写字母形式</param>
        ''' <returns>加密后的结果字符串</returns>
        ''' <remarks></remarks>
        Public Shared Function MD4_Encode(ByVal Source As String, Optional ByVal ToUpper As Boolean = True) As String
            Dim MD4 As MD4_Hash_Algorithm = New MD4_Hash_Algorithm(System.Text.Encoding.UTF8.GetBytes(Source))
            If ToUpper Then
                Return BitConverter.ToString(MD4.DigestResult()).Replace("-", "").ToUpper()
            Else
                Return BitConverter.ToString(MD4.DigestResult()).Replace("-", "").ToLower()
            End If
        End Function



        ''' <summary>
        ''' MD4混合加密（用于ED2K链接中的文件哈希值的计算，摘要结果为32位16进制字符串）
        ''' </summary>
        ''' <param name="Source">要加密的Byte数组</param>
        ''' <param name="ToUpper">是否将结果转换为大写字母形式</param>
        ''' <returns>加密后的结果字符串</returns>
        ''' <remarks></remarks>
        Public Shared Function MD4_ED2K_Encode(ByVal Source As Byte(), Optional ByVal ToUpper As Boolean = True) As String
            Dim ChunkSize As Integer = 9728000
            Dim ChunkCount As Integer = Math.Ceiling(Source.Length / ChunkSize)
            Dim ChunkBuffer(ChunkSize - 1) As Byte
            Dim ChunkHash As New List(Of Byte)
            For I = 0 To ChunkCount - 1
                If Source.Length - (ChunkSize * I) > ChunkSize Then
                    Array.Copy(Source, ChunkSize * I, ChunkBuffer, 0, ChunkSize)
                    Dim Chunk_MD4 As MD4_Hash_Algorithm = New MD4_Hash_Algorithm(ChunkBuffer)
                    ChunkHash.AddRange(Chunk_MD4.DigestResult())
                    For J = 0 To ChunkBuffer.Length - 1
                        ChunkBuffer(J) = 0
                    Next
                Else
                    ReDim ChunkBuffer(Source.Length - (ChunkSize * I) - 1)
                    Array.Copy(Source, ChunkSize * I, ChunkBuffer, 0, ChunkBuffer.Length)
                    Dim Chunk_MD4 As MD4_Hash_Algorithm = New MD4_Hash_Algorithm(ChunkBuffer)
                    ChunkHash.AddRange(Chunk_MD4.DigestResult())
                End If
            Next
            If ChunkCount = 1 Then
                If ToUpper Then
                    Return BitConverter.ToString(ChunkHash.ToArray()).Replace("-", "").ToUpper()
                Else
                    Return BitConverter.ToString(ChunkHash.ToArray()).Replace("-", "").ToLower()
                End If
            Else
                Dim Total_MD4 As MD4_Hash_Algorithm = New MD4_Hash_Algorithm(ChunkHash.ToArray())
                If ToUpper Then
                    Return BitConverter.ToString(Total_MD4.DigestResult()).Replace("-", "").ToUpper()
                Else
                    Return BitConverter.ToString(Total_MD4.DigestResult()).Replace("-", "").ToLower()
                End If
            End If
        End Function

        ''' <summary>
        ''' 根据文件的名称、大小、哈希值，生成文件的ED2K下载链接
        ''' </summary>
        ''' <param name="FileName">文件名称（不能为空字符串）</param>
        ''' <param name="FileLength">文件大小（必须大于0）</param>
        ''' <param name="FileHash">文件哈希值（用于ED2K链接中的MD4混合哈希值）</param>
        ''' <returns>生成的ED2K链接结果字符串（失败返回空字符串）</returns>
        ''' <remarks></remarks>
        Public Shared Function Generate_ED2K_Link(ByVal FileName As String, ByVal FileLength As Integer, ByVal FileHash As String) As String
            If FileName = "" OrElse FileHash.Length <> 32 OrElse FileLength <= 0 Then
                Return ""
            End If
            FileName = System.Web.HttpUtility.UrlEncode(FileName, System.Text.Encoding.UTF8)
            Return "ed2k://|file|" & FileName & "|" & FileLength & "|" & FileHash & "|/"
        End Function
        ''' <summary>
        ''' 根据文件的名称和字节内容，生成文件的ED2K下载链接
        ''' </summary>
        ''' <param name="FileName">文件名称（不能为空字符串）</param>
        ''' <param name="Source">文件内容（不能为空Byte数组）</param>
        ''' <returns>生成的ED2K链接结果字符串（失败返回空字符串）</returns>
        ''' <remarks></remarks>
        Public Shared Function Generate_ED2K_Link(ByVal FileName As String, ByVal Source As Byte()) As String
            If FileName = "" OrElse Source.Length = 0 Then
                Return ""
            End If
            FileName = System.Web.HttpUtility.UrlEncode(FileName, System.Text.Encoding.UTF8)
            Dim ChunkSize As Integer = 9728000
            Dim ChunkCount As Integer = Math.Ceiling(Source.Length / ChunkSize)
            Dim ChunkBuffer(ChunkSize - 1) As Byte
            Dim ChunkHash As New List(Of Byte)
            For I = 0 To ChunkCount - 1
                If Source.Length - (ChunkSize * I) > ChunkSize Then
                    Array.Copy(Source, ChunkSize * I, ChunkBuffer, 0, ChunkSize)
                    Dim Chunk_MD4 As MD4_Hash_Algorithm = New MD4_Hash_Algorithm(ChunkBuffer)
                    ChunkHash.AddRange(Chunk_MD4.DigestResult())
                    For J = 0 To ChunkBuffer.Length - 1
                        ChunkBuffer(J) = 0
                    Next
                Else
                    ReDim ChunkBuffer(Source.Length - (ChunkSize * I) - 1)
                    Array.Copy(Source, ChunkSize * I, ChunkBuffer, 0, ChunkBuffer.Length)
                    Dim Chunk_MD4 As MD4_Hash_Algorithm = New MD4_Hash_Algorithm(ChunkBuffer)
                    ChunkHash.AddRange(Chunk_MD4.DigestResult())
                End If
            Next
            Dim FileHash As String
            If ChunkCount = 1 Then
                FileHash = BitConverter.ToString(ChunkHash.ToArray()).Replace("-", "")
            Else
                Dim Total_MD4 As MD4_Hash_Algorithm = New MD4_Hash_Algorithm(ChunkHash.ToArray())
                FileHash = BitConverter.ToString(Total_MD4.DigestResult()).Replace("-", "")
            End If
            Return "ed2k://|file|" & FileName & "|" & Source.Length & "|" & FileHash & "|/"
        End Function

    End Class

End Namespace