Namespace My

    ''' <summary>
    ''' 磁盘文件读写相关函数
    ''' </summary>
    ''' <remarks></remarks>
    Partial Public NotInheritable Class IO

        ''' <summary>
        ''' 读取文件为Byte数组
        ''' </summary>
        ''' <param name="FilePath">文件路径（可以是相对路径）</param>
        ''' <returns>结果Byte数组（失败返回空Byte数组）</returns>
        ''' <remarks></remarks>
        Public Shared Function ReadByte(ByVal FilePath As String) As Byte()
            Dim Reader As System.IO.FileStream
            Try
                Reader = New System.IO.FileStream(FilePath, System.IO.FileMode.Open, System.IO.FileAccess.Read)
                Dim Temp(Reader.Length - 1) As Byte
                Reader.Read(Temp, 0, Reader.Length)
                Reader.Dispose()
                Return Temp
            Catch ex As Exception
                Return New Byte() {}
            End Try
        End Function

        ''' <summary>
        ''' 将Byte数组写入文件（覆盖）
        ''' </summary>
        ''' <param name="Source">Byte数组</param>
        ''' <param name="FilePath">文件路径（可以是相对路径）</param>
        ''' <returns>是否写入成功</returns>
        ''' <remarks></remarks>
        Public Shared Function WriteByte(ByVal Source As Byte(), ByVal FilePath As String) As Boolean
            Dim Writer As System.IO.FileStream
            Try
                Writer = New System.IO.FileStream(FilePath, System.IO.FileMode.OpenOrCreate, System.IO.FileAccess.ReadWrite)
                Writer.Write(Source, 0, Source.Length)
                Writer.Dispose()
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function

        ''' <summary>
        ''' 将Byte数组写入文件（追加）
        ''' </summary>
        ''' <param name="Source">Byte数组</param>
        ''' <param name="FilePath">文件路径（可以是相对路径）</param>
        ''' <returns>是否写入成功</returns>
        ''' <remarks></remarks>
        Public Shared Function AppendByte(ByVal Source As Byte(), ByVal FilePath As String) As Boolean
            Dim Writer As System.IO.FileStream
            Try
                Writer = New System.IO.FileStream(FilePath, System.IO.FileMode.OpenOrCreate, System.IO.FileAccess.ReadWrite)
                Writer.Write(Source, Writer.Length, Source.Length)
                Writer.Dispose()
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function



        ''' <summary>
        ''' 读取文件为字符串（UTF-8）
        ''' </summary>
        ''' <param name="FilePath">文件路径（可以是相对路径）</param>
        ''' <returns>结果字符串（失败返回空字符串）</returns>
        ''' <remarks></remarks>
        Public Shared Function ReadString(ByVal FilePath As String) As String
            Dim Reader As System.IO.StreamReader
            Try
                Reader = New System.IO.StreamReader(FilePath, System.Text.Encoding.UTF8)
                Dim Temp As String = Reader.ReadToEnd()
                Reader.Dispose()
                Return Temp
            Catch ex As Exception
                Return ""
            End Try
        End Function

        ''' <summary>
        ''' 将字符串写入文件（覆盖，不包含UTF8的BOM头）
        ''' </summary>
        ''' <param name="Source">字符串</param>
        ''' <param name="FilePath">文件路径（可以是相对路径）</param>
        ''' <returns>是否写入成功</returns>
        ''' <remarks></remarks>
        Public Shared Function WriteString(ByVal Source As String, ByVal FilePath As String) As Boolean
            Dim Writer As System.IO.StreamWriter
            Try
                Writer = New System.IO.StreamWriter(FilePath, False, New System.Text.UTF8Encoding(False))
                Writer.Write(Source)
                Writer.Dispose()
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function

        ''' <summary>
        ''' 将字符串写入文件（追加，不包含UTF8的BOM头）
        ''' </summary>
        ''' <param name="Source">字符串</param>
        ''' <param name="FilePath">文件路径（可以是相对路径）</param>
        ''' <returns>是否写入成功</returns>
        ''' <remarks></remarks>
        Public Shared Function AppendString(ByVal Source As String, ByVal FilePath As String) As Boolean
            Dim Writer As System.IO.StreamWriter
            Try
                Writer = New System.IO.StreamWriter(FilePath, True, New System.Text.UTF8Encoding(False))
                Writer.Write(Source)
                Writer.Dispose()
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function



        ''' <summary>
        ''' 读取文件中的字符串数组（UTF-8）
        ''' </summary>
        ''' <param name="FilePath">文件路径（可以是相对路径）</param>
        ''' <returns>结果字符串数组（失败返回空String数组）</returns>
        ''' <remarks></remarks>
        Public Shared Function ReadStringArray(ByVal FilePath As String) As String()
            Dim Reader As System.IO.StreamReader
            Try
                Reader = New System.IO.StreamReader(FilePath, System.Text.Encoding.UTF8)
                Dim Temp As String = Reader.ReadToEnd()
                Reader.Dispose()
                Return Temp.Replace(Chr(13) + Chr(10), Chr(10)).Split(New Char() {Chr(10)})
            Catch ex As Exception
                Return New String() {}
            End Try
        End Function

        ''' <summary>
        ''' 将字符串数组写入文件（覆盖，不包含UTF8的BOM头）
        ''' </summary>
        ''' <param name="StringArray">字符串数组</param>
        ''' <param name="FilePath">文件路径（可以是相对路径）</param>
        ''' <returns>是否写入成功</returns>
        ''' <remarks></remarks>
        Public Shared Function WriteStringArray(ByVal StringArray As String(), ByVal FilePath As String) As Boolean
            Dim Builder As System.Text.StringBuilder
            Dim Writer As System.IO.StreamWriter
            Try
                Builder = New System.Text.StringBuilder()
                For I = 0 To StringArray.Length - 2
                    Builder.Append(StringArray(I) & vbCrLf)
                Next
                If StringArray.Length > 0 Then
                    Builder.Append(StringArray(StringArray.Length - 1))
                End If
                Writer = New System.IO.StreamWriter(FilePath, False, New System.Text.UTF8Encoding(False))
                Writer.Write(Builder)
                Writer.Dispose()
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function

        ''' <summary>
        ''' 将字符串数组写入文件（追加，不包含UTF8的BOM头）
        ''' </summary>
        ''' <param name="StringArray">字符串数组</param>
        ''' <param name="FilePath">文件路径（可以是相对路径）</param>
        ''' <returns>是否写入成功</returns>
        ''' <remarks></remarks>
        Public Shared Function AppendStringArray(ByVal StringArray As String(), ByVal FilePath As String) As Boolean
            Dim Builder As System.Text.StringBuilder
            Dim Writer As System.IO.StreamWriter
            Try
                Builder = New System.Text.StringBuilder()
                For I = 0 To StringArray.Length - 1
                    Builder.Append(vbCrLf & StringArray(I))
                Next
                Writer = New System.IO.StreamWriter(FilePath, True, New System.Text.UTF8Encoding(False))
                Writer.Write(Builder)
                Writer.Dispose()
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function



        ''' <summary>
        ''' 获取指定目录下的全部文件的路径列表
        ''' </summary>
        ''' <param name="SearchDirectory">要搜索的文件夹路径（默认为程序运行的当前文件夹）</param>
        ''' <returns>包含所有文件路径的结果字符串数组（失败返回空String数组）</returns>
        ''' <remarks></remarks>
        Public Shared Function ListFile(Optional ByVal SearchDirectory As String = ".\") As String()
            Dim FileList As List(Of String) = New List(Of String)
            Dim Directory As List(Of String) = New List(Of String)
            Try
                For Each Temp As String In System.IO.Directory.GetFiles(SearchDirectory)
                    FileList.Add(Temp)
                Next
                For Each Temp As String In System.IO.Directory.GetDirectories(SearchDirectory)
                    Directory.Add(Temp)
                Next
                Dim Index As Integer = 0
                While Directory.Count > Index
                    SearchDirectory = Directory(Index)
                    Index = Index + 1
                    For Each Temp As String In System.IO.Directory.GetFiles(SearchDirectory)
                        FileList.Add(Temp)
                    Next
                    For Each Temp As String In System.IO.Directory.GetDirectories(SearchDirectory)
                        Directory.Add(Temp)
                    Next
                End While
                Return FileList.ToArray()
            Catch ex As Exception
                Return New String() {}
            End Try
        End Function



        ''' <summary>
        ''' 创建快捷方式文件（覆盖）
        ''' </summary>
        ''' <param name="TargetPath">快捷方式指向的路径（可以是相对路径，如"1.exe"）</param>
        ''' <param name="LinkFilePath">快捷方式文件的路径（可以是相对路径，如"1.lnk"）</param>
        ''' <param name="Arguments">打开程序的参数（例如"/?"）</param>
        ''' <param name="Description">鼠标悬停在快捷方式上的描述</param>
        ''' <param name="WorkingDirectory">快捷方式的起始位置（默认设置为快捷方式指向的路径的父目录）</param>
        ''' <returns>是否创建成功</returns>
        ''' <remarks></remarks>
        Public Shared Function WriteLinkFile(ByVal TargetPath As String, ByVal LinkFilePath As String, Optional ByVal Arguments As String = "", Optional ByVal Description As String = "", Optional ByVal WorkingDirectory As String = "") As Boolean
            Try
                If System.IO.File.Exists(LinkFilePath) = True Then
                    System.IO.File.Delete(LinkFilePath)
                End If
                If TargetPath.Contains(":") = False Then
                    TargetPath = System.IO.Directory.GetCurrentDirectory + "\" + TargetPath
                End If
                If WorkingDirectory = "" Then
                    WorkingDirectory = System.IO.Directory.GetParent(LinkFilePath).FullName
                End If
                Dim Shortcut As Object = CreateObject("WScript.Shell").CreateShortcut(LinkFilePath)
                Shortcut.TargetPath = TargetPath
                Shortcut.IconLocation = TargetPath
                Shortcut.Arguments = Arguments
                Shortcut.Description = Description
                Shortcut.WorkingDirectory = WorkingDirectory
                Shortcut.Save()
                Shortcut = Nothing
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function

        ''' <summary>
        ''' 读取快捷方式指向的路径（获得完整的绝对路径）
        ''' </summary>
        ''' <param name="LinkFilePath">快捷方式文件的路径（可以是相对路径，如"1.lnk"）</param>
        ''' <returns>结果字符串（失败返回空字符串""）</returns>
        ''' <remarks></remarks>
        Public Shared Function ReadLinkFile(ByVal LinkFilePath As String) As String
            Try
                Dim Shortcut As Object = CreateObject("WScript.Shell").CreateShortcut(LinkFilePath)
                Return Shortcut.TargetPath
            Catch ex As Exception
                Return ""
            End Try
        End Function

    End Class

End Namespace