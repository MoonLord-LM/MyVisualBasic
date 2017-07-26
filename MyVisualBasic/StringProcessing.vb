Namespace My
    
    ''' <summary>
    ''' 字符串处理相关函数
    ''' </summary>
    ''' <remarks></remarks>
    Partial Public NotInheritable Class StringProcessing

        ''' <summary>
        ''' 搜索字符串（搜寻第一个开始字符串，再向后搜寻第一个结束字符串，取出中间的部分）
        ''' </summary>
        ''' <param name="Source">要搜索的字符串</param>
        ''' <param name="BeginString">开始位置的字符串（结果中不包含这一部分）</param>
        ''' <param name="EndString">结束位置的字符串（结果中不包含这一部分）</param>
        ''' <returns>结果字符串（失败返回空字符串）</returns>
        ''' <remarks></remarks>
        Public Shared Function Find(ByVal Source As String, ByVal BeginString As String, ByVal EndString As String) As String
            If Source.Contains(BeginString) = False Then
                Return ""
            End If
            Source = Source.Substring(Source.IndexOf(BeginString) + BeginString.Length)
            If Source.Contains(EndString) = False Then
                Return ""
            End If
            Return Source.Substring(0, Source.IndexOf(EndString))
        End Function

        ''' <summary>
        ''' 搜索字符串（重复：搜寻开始字符串，再向后搜寻结束字符串，取出中间的部分）
        ''' </summary>
        ''' <param name="Source">要搜索的字符串</param>
        ''' <param name="BeginString">开始位置的字符串（结果中不包含这一部分）</param>
        ''' <param name="EndString">结束位置的字符串（结果中不包含这一部分）</param>
        ''' <returns>结果字符串数组（失败返回空String数组）</returns>
        ''' <remarks></remarks>
        Public Shared Function FindAll(ByVal Source As String, ByVal BeginString As String, ByVal EndString As String) As String()
            Dim Result As List(Of String) = New List(Of String)
            While (Source.Contains(BeginString))
                Source = Source.Substring(Source.IndexOf(BeginString) + BeginString.Length)
                If Source.Contains(EndString) = False Then
                    Exit While
                Else
                    Result.Add(Source.Substring(0, Source.IndexOf(EndString)))
                    Source = Source.Substring(Source.IndexOf(EndString) + EndString.Length)
                End If
            End While
            Return Result.ToArray()
        End Function



        ''' <summary>
        ''' 搜索字符串（搜寻最后一个裁剪标志字符串，取出之后的部分）
        ''' </summary>
        ''' <param name="Source">要搜索的字符串</param>
        ''' <param name="CutString">裁剪位置的标志字符串（结果中不包含这一部分）</param>
        ''' <returns>结果字符串（失败返回空字符串）</returns>
        ''' <remarks></remarks>
        Public Shared Function FindAfter(ByVal Source As String, ByVal CutString As String) As String
            If Source.Contains(CutString) = False Then
                Return ""
            End If
            Return Source.Substring(Source.LastIndexOf(CutString) + CutString.Length)
        End Function

        ''' <summary>
        ''' 搜索字符串（搜寻第一个裁剪标志字符串，取出之前的部分）
        ''' </summary>
        ''' <param name="Source">要搜索的字符串</param>
        ''' <param name="CutString">裁剪位置的标志字符串（结果中不包含这一部分）</param>
        ''' <returns>结果字符串（失败返回空字符串）</returns>
        ''' <remarks></remarks>
        Public Shared Function FindBefore(ByVal Source As String, ByVal CutString As String) As String
            If Source.Contains(CutString) = False Then
                Return ""
            End If
            Return Source.Substring(0, Source.IndexOf(CutString))
        End Function



        ''' <summary>
        ''' 搜索字符串数组，取出不为空字符串""的元素，返回新数组（包括空格和制表符）
        ''' </summary>
        ''' <param name="Source">要搜索的字符串数组</param>
        ''' <returns>结果字符串数组（失败返回空String数组）</returns>
        ''' <remarks></remarks>
        Public Shared Function SelectNotEmpty(ByVal Source As String()) As String()
            Dim Result As List(Of String) = New List(Of String)
            For I = 0 To Source.Length - 1
                If Source(I) <> "" Then
                    Result.Add(Source(I))
                End If
            Next
            Return Result.ToArray()
        End Function

    End Class

End Namespace