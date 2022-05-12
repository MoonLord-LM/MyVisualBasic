Namespace My

    ''' <summary>
    ''' 范例代码（暂未整理到 MyVisualBasic 库中的可参考代码）
    ''' </summary>
    ''' <remarks></remarks>
    Partial Public NotInheritable Class Example

        ''' <summary>
        ''' 内存检测（不断申请256MB空间的Int数组，并反复读写内容，检测内存质量）
        ''' </summary>
        ''' <remarks></remarks>
        Public Shared Sub MemoryTest()
            System.GC.Collect()
            Dim Tested As New ArrayList()
            Dim BlockSize As UInt32 = 256 * 1024 * 1024
            Dim BlockSizeMB As UInt32 = Int(BlockSize / 1024 / 1024)
            Dim TestedGB As UInt32 = 0
            While True 'My.Computer.Info.AvailablePhysicalMemory > BlockSize
                Dim Testing As UInt32() = Nothing
                '初始0
                Try
                    ReDim Preserve Testing(BlockSize / 4 - 1)
                Catch ex As System.OutOfMemoryException
                    ex = New OutOfMemoryException("【内存检测通过】：" & vbCrLf & "测试通过 " & Tested.Count & " * " & BlockSizeMB & " MB（" & TestedGB & "GB）。")
                    Tested.Clear()
                    System.GC.Collect()
                    Throw ex
                End Try
                For I = 0 To Testing.Length - 1
                    If Testing(I) <> 0 Then
                        Throw New Exception("【发现内存错误】初始0值异常：" & vbCrLf & "测试通过 " & Tested.Count & " * " & BlockSizeMB & " MB（" & TestedGB & "GB），错误位置 " & Int(I * 4 / 1024 / 1024) & " MB。")
                    End If
                Next
                '重置1
                For I = 0 To Testing.Length - 1
                    Testing(I) = UInt32.MaxValue
                Next
                For I = 0 To Testing.Length - 1
                    If Testing(I) <> UInt32.MaxValue Then
                        Throw New Exception("【发现内存错误】重置1值异常：" & vbCrLf & "测试通过 " & Tested.Count & " * " & BlockSizeMB & " MB（" & TestedGB & "GB），错误位置 " & Int(I * 4 / 1024 / 1024) & " MB。")
                    End If
                Next
                '重置0
                For I = 0 To Testing.Length - 1
                    Testing(I) = UInt32.MinValue
                Next
                For I = 0 To Testing.Length - 1
                    If Testing(I) <> 0 Then
                        Throw New Exception("【发现内存错误】重置0值异常：" & vbCrLf & "测试通过 " & Tested.Count & " * " & BlockSizeMB & " MB（" & TestedGB & "GB），错误位置 " & Int(I * 4 / 1024 / 1024) & " MB。")
                    End If
                Next
                Tested.Add(Testing)
                TestedGB = Tested.Count * BlockSizeMB / 1024
            End While
        End Sub



        ''' <summary>
        ''' 获取上一次调用Win32 API产生的错误信息（实测：错误信息会一直保留，直到下一次调用Win32 API）
        ''' </summary>
        ''' <returns>错误信息（默认为"0 操作成功完成。"）</returns>
        ''' <remarks></remarks>
        Public Shared Function Win32Error() As String
            Dim ErrorCode As Int32 = Runtime.InteropServices.Marshal.GetLastWin32Error()
            Dim ErrorMessage As String = New System.ComponentModel.Win32Exception(Runtime.InteropServices.Marshal.GetLastWin32Error()).Message
            Return ErrorCode & " " & ErrorMessage
        End Function



        ''' <summary>
        ''' 在鼠标位置附近，创建一个无边框窗口，实时显示全屏的截图
        ''' </summary>
        ''' <remarks></remarks>
        Public Shared Sub CaptureScreen()
            Dim CaptureForm As New CaptureForm()
            CaptureForm.Show()
        End Sub
        <Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
        Private Class CaptureForm
            Inherits System.Windows.Forms.Form
            Protected Overrides Sub Dispose(ByVal disposing As Boolean)
                Try
                    If disposing AndAlso components IsNot Nothing Then
                        components.Dispose()
                    End If
                Finally
                    MyBase.Dispose(disposing)
                End Try
            End Sub
            Private components As System.ComponentModel.IContainer
            Private Sub InitializeComponent()
                Me.components = New System.ComponentModel.Container()
                Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
                Me.SuspendLayout()
                Me.Timer1.Enabled = True
                Me.Timer1.Interval = 16
                Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None
                Me.ResumeLayout(False)
            End Sub
            Friend WithEvents Timer1 As System.Windows.Forms.Timer
            Private Sub FormMain_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
                Me.TopMost = True
                Me.FormBorderStyle = Windows.Forms.FormBorderStyle.None
                Me.SetStyle(ControlStyles.OptimizedDoubleBuffer, True)
                Me.SetStyle(ControlStyles.AllPaintingInWmPaint, True)
                Me.Opacity = 0
            End Sub
            Private Sub Timer1_Tick(sender As System.Object, e As System.EventArgs) Handles Timer1.Tick
                Timer1.Enabled = False
                If Not BackgroundImage Is Nothing Then BackgroundImage.Dispose()
                Me.BackgroundImage = My.Screen.ImageThumbnail(0.5)
                Me.Size = Me.BackgroundImage.Size
                Me.Location = My.Mouse.Position() + New Size(10, 10)
                Me.Opacity = 1
                Timer1.Enabled = True
            End Sub
        End Class



        ''' <summary>
        ''' 获取一个窗体的所有的可见子窗体句柄，保存信息并截图
        ''' </summary>
        ''' <param name="hWnd">窗口句柄（IntPtr）</param>
        ''' <remarks></remarks>
        Public Shared Sub AnalysisChildren(ByVal hWnd As IntPtr)
            My.Window.SetFocus(hWnd)
            Dim SavePath As String = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory)
            Dim WindowsName As String = My.Window.GetTitle(hWnd)
            SavePath = SavePath & "\" & WindowsName & My.Time.Stamp()
            System.IO.Directory.CreateDirectory(SavePath)
            My.IO.WriteString("", SavePath & "\Name_ClassName_Left_Top_Width_Height_Child_Father" & ".txt")
            Dim Children As New List(Of IntPtr)(My.Window.ListChildren(hWnd))
            Dim I As Int32 = 0
            While (I < Children.Count)
                Dim Child As IntPtr = Children(I)
                Dim Rectangle As Rectangle = My.Window.GetRectangle(Child)
                If Rectangle.Width > 0 And Rectangle.Height > 0 Then
                    Dim Name As String = My.Window.GetTitle(Child)
                    Dim ClassName As String = My.Window.GetClassName(Child)
                    Dim Father As IntPtr = My.Window.FindParent(Child)
                    Dim Image As Image = My.Screen.Image(Rectangle)
                    Image.Save(SavePath & "\" & Name & "_" & ClassName & "_" & Rectangle.Left & "_" & Rectangle.Top & "_" & Rectangle.Width & "_" & Rectangle.Height & "_" & Child.ToString() & "_" & Father.ToString() & ".png")
                    Image.Dispose()
                End If
                I = I + 1
            End While
        End Sub



        ''' <summary>
        ''' 对当前目录下的，所有用.NET Reflector从“.vb”转换生成的“.cs”文件，进行简单的修正
        ''' </summary>
        ''' <remarks></remarks>
        Public Shared Sub CheckVBToCSharp()
            Dim Dialog1 As New FolderBrowserDialog
            Dialog1.Description = "请选择反编译出的 C# 代码的路径"
            Dialog1.SelectedPath = "F:\Desktop\新建文件夹\MyVisualBasic\My"
            If Dialog1.ShowDialog() = DialogResult.OK Then
                Dim CSFiles As String() = My.IO.ListFile(Dialog1.SelectedPath)
                For FI = 0 To CSFiles.Length - 1
                    Dim File As String = CSFiles(FI)
                    Dim Codes As String() = My.IO.ReadStringArray(File)
                    For CI = 0 To Codes.Length - 1
                        Dim Code As String = Codes(CI)
                        If (Code = "") Then
                            Codes(CI) = " "
                        End If
                        If Code.Contains("ProjectData.SetProjectError(exception1);") Then
                            Codes(CI) = ""
                            Continue For
                        End If
                        If Code.Contains("Exception exception = exception1;") Then
                            Codes(CI) = ""
                            Continue For
                        End If
                        If Code.Contains("ProjectData.ClearProjectError();") Then
                            Codes(CI) = ""
                            Continue For
                        End If
                        If Code.Contains("using Microsoft.VisualBasic.CompilerServices;") Then
                            Codes(CI) = ""
                            Continue For
                        End If
                        Code = Code.Replace("namespace MyVisualBasic.My", "namespace My")
                        Code = Code.Replace("catch (Exception exception1)", "catch (Exception ex)")
                        Code = Code.Replace("catch (Exception exception3)", "catch (Exception ex)")
                        Code = Code.Replace("catch (Exception exception5)", "catch (Exception ex)")
                        Code = Code.Replace("string str;", "string result;")
                        Code = Code.Replace("str = ", "result = ")
                        Code = Code.Replace("return str;", "return result;")
                        Code = Code.Replace("byte[] buffer;", "byte[] result;")
                        Code = Code.Replace("buffer = ", "result = ")
                        Code = Code.Replace("return buffer;", "return result;")
                        Code = Code.Replace("bool flag;", "bool result;")
                        Code = Code.Replace("flag = ", "result = ")
                        Code = Code.Replace("return flag;", "return result;")
                        Codes(CI) = Code
                    Next
                    Codes = My.StringProcessing.SelectNotEmpty(Codes)
                    My.IO.WriteStringArray(Codes, File)
                Next
            End If
        End Sub



        ''' <summary>
        ''' 文件列表式备份（保存文件的绝对路径、名称、后缀、大小、哈希值、ED2K链接等）
        ''' </summary>
        ''' <param name="CalculateHash">是否读取文件内容，计算哈希值（较慢）</param>
        ''' <remarks></remarks>
        Public Shared Sub FileListBackup(Optional ByVal CalculateHash As Boolean = False)
            Dim Dialog1 As New FolderBrowserDialog
            Dialog1.Description = "请选择要备份的文件夹路径"
            If Dialog1.ShowDialog() = DialogResult.OK Then
                Dim DirInfo As New System.IO.DirectoryInfo(Dialog1.SelectedPath)
                Dim SavePath As String = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory)
                SavePath = SavePath & "\" & "【" & DirInfo.Name & "】" & My.Time.Stamp() & ".log"
                My.IO.WriteString("【备份时间：" & Now & "】" & vbCrLf, SavePath)
                My.IO.AppendString("【备份路径：" & DirInfo.FullName & "】" & vbCrLf, SavePath)
                My.IO.AppendString("【本地时区：" & TimeZone.CurrentTimeZone.StandardName & " " & TimeZone.CurrentTimeZone.GetUtcOffset(Now).ToString() & "】" & vbCrLf, SavePath)
                Dim Files As String() = My.IO.ListFile(Dialog1.SelectedPath)
                My.IO.AppendString("【文件数目：" & Files.Length & "】" & vbCrLf, SavePath)
                If CalculateHash = False Then
                    My.IO.AppendString("【属性说明：完整路径、文件名、文件后缀<快捷方式指向位置>、文件大小、创建时间、修改时间、访问时间】" & vbCrLf, SavePath)
                Else
                    My.IO.AppendString("【属性说明：完整路径、文件名、文件后缀<快捷方式指向位置>、文件大小、创建时间、修改时间、访问时间、ED2K下载链接、MD5值、SHA1值、SHA256值、SHA384值、SHA512值】" & vbCrLf, SavePath)
                End If
                For I = 0 To Files.Length - 1
                    Dim File As String = Files(I)
                    Dim FileInfo As New System.IO.FileInfo(File)
                    Dim Buffer As New System.Text.StringBuilder()
                    Buffer.Append(vbCrLf)
                    Buffer.Append("""" & (I + 1) & """")
                    Buffer.Append(" ")
                    Buffer.Append("""" & FileInfo.FullName & """")
                    Buffer.Append(" ")
                    Buffer.Append("""" & FileInfo.Name & """")
                    Buffer.Append(" ")
                    If FileInfo.Extension.ToLower() <> ".lnk" Then
                        Buffer.Append("""" & FileInfo.Extension & """")
                        Buffer.Append(" ")
                    Else
                        Dim Link As String = My.IO.ReadLinkFile(File)
                        Buffer.Append("""" & FileInfo.Extension & "<" & Link & ">" & """")
                        Buffer.Append(" ")
                        File = Link
                    End If
                    Buffer.Append("""" & FileInfo.Length & """")
                    Buffer.Append(" ")
                    Buffer.Append("""" & FileInfo.CreationTime() & """")
                    Buffer.Append(" ")
                    Buffer.Append("""" & FileInfo.LastWriteTime() & """")
                    Buffer.Append(" ")
                    Buffer.Append("""" & FileInfo.LastAccessTime() & """")
                    Buffer.Append(" ")
                    If CalculateHash Then
                        Dim Source As Byte() = My.IO.ReadByte(File)
                        Buffer.Append("""" & My.Security.Generate_ED2K_Link(FileInfo.Name, Source) & """")
                        Buffer.Append(" ")
                        Buffer.Append("""" & My.Security.MD5_Encode(Source) & """")
                        Buffer.Append(" ")
                        Buffer.Append("""" & My.Security.SHA1_Encode(Source) & """")
                        Buffer.Append(" ")
                        Buffer.Append("""" & My.Security.SHA256_Encode(Source) & """")
                        Buffer.Append(" ")
                        Buffer.Append("""" & My.Security.SHA384_Encode(Source) & """")
                        Buffer.Append(" ")
                        Buffer.Append("""" & My.Security.SHA512_Encode(Source) & """")
                        Buffer.Append(" ")
                    End If
                    My.IO.AppendString(Buffer.ToString(), SavePath)
                Next
            End If

        End Sub

    End Class

End Namespace