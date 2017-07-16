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
            While My.Computer.Info.AvailablePhysicalMemory > BlockSize
                Dim Testing As UInt32() = Nothing
                '初始0
                Try
                    ReDim Preserve Testing(BlockSize / 4 - 1)
                Catch ex As System.OutOfMemoryException
                    Dim Result = New OutOfMemoryException("【内存检测通过】：" & vbCrLf & "测试通过 " & Tested.Count & " * " & BlockSizeMB & " MB（" & TestedGB & "GB）。")
                    Tested.Clear()
                    System.GC.Collect()
                    Throw Result
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
            <System.Diagnostics.DebuggerNonUserCode()> _
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
            <System.Diagnostics.DebuggerStepThrough()> _
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
            End Sub
            Private Sub Timer1_Tick(sender As System.Object, e As System.EventArgs) Handles Timer1.Tick
                Timer1.Enabled = False
                If Not BackgroundImage Is Nothing Then BackgroundImage.Dispose()
                Me.BackgroundImage = My.Screen.ImageThumbnail(0.5)
                Me.Size = Me.BackgroundImage.Size
                Me.Location = My.Mouse.Position() + New Point(10, 10)
                Timer1.Enabled = True
            End Sub
        End Class

        ''' <summary>
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

    End Class

End Namespace