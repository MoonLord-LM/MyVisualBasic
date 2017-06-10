# MyVisualBasic
A function library to extend the My namespace of VB.NET.  
  
## [简介]
扩展 VB.NET 中的 My 命名空间的函数库  
对Win32 API，VBScript，Batch等进行了封装，并分类整理了常用的函数功能  
默认字符编码：UTF-8  
项目需要引用：System，System.Drawing，System.Web，System.Windows.Forms  
  
## [用法]
引入【MyVisualBisic】文件夹中的.vb文件，获得函数扩展  
各个.vb文件之间互不依赖，可根据需要选择使用，推荐直接引入整个文件夹  
  
## [说明]
<table>
    <tr>
        <td><a href="MyVisualBisic\Security.vb">Security.vb</a></td>
		<td>编码解码、加密解密相关函数</td>
    </tr>
    <tr>
        <td><a href="MyVisualBisic\Http.vb">Http.vb</a></td>
		<td>HTTP网络请求相关函数</td>
    </tr>
    <tr>
        <td><a href="MyVisualBisic\IO.vb">IO.vb</a></td>
		<td>磁盘文件读写相关函数</td>
    </tr>
    <tr>
        <td><a href="MyVisualBisic\Resource.vb">Resource.vb</a></td>
		<td>程序资源文件相关函数</td>
    </tr>
    <tr>
        <td><a href="MyVisualBisic\StringProcessing.vb">StringProcessing.vb</a></td>
		<td>字符串处理相关函数</td>
    </tr>
    <tr>
        <td><a href="MyVisualBisic\Power.vb">Power.vb</a></td>
		<td>电源管理计划相关函数</td>
    </tr>
    <tr>
        <td><a href="MyVisualBisic\Task.vb">Task.vb</a></td>
		<td>进程管理相关函数</td>
    </tr>
    <tr>
        <td><a href="MyVisualBisic\Screen.vb">Screen.vb</a></td>
		<td>屏幕截图相关函数</td>
    </tr>
    <tr>
        <td><a href="MyVisualBisic\Keyboard.vb">Keyboard.vb</a></td>
		<td>模拟键盘操作相关函数</td>
    </tr>
    <tr>
        <td><a href="MyVisualBisic\Time.vb">Time.vb</a></td>
		<td>时间管理、转换相关函数</td>
    </tr>
    <tr>
        <td><a href="MyVisualBisic\Mouse.vb">Mouse.vb</a></td>
		<td>模拟鼠标操作相关函数</td>
    </tr>
    <tr>
        <td><a href="MyVisualBisic\Window.vb">Window.vb</a></td>
		<td>窗口管理、控制相关函数</td>
    </tr>
</table>
  
## [参考]
- 引入方法：解决方案资源管理器，显示所有文件，刷新，右键“.vb”文件，包括在项目内  
- 空行规则：同一方面的函数放在同一源文件中，重载函数不用空行隔开，相关函数隔开一行，无关函数隔开三行  
- 注释示例：“结果字符串（失败返回空字符串）”，“结果Byte数组（失败返回空Byte数组）”，“使用特定的字符编码（默认UTF-8）”  
  
## [教程]
01. 在Visual Studio中显示行号：工具，选项，文本编辑器，所有语言，显示行号  
02. VB.NET中，数组与ArrayList互相转换：(New ArrayList(New String() {})).ToArray(String.Empty.GetType())  
03. VB.NET中，数组元素个数的声明与其它语言不同，Dim Array(2) As String : MsgBox(Array.Length)，输出3  
04. VB.NET中，双引号使用两个双引号来转义替代，如""""表示1个双引号的字符串，字符串用&符号来连接  
05. VB.NET中，想要在“调试”状态下，程序也能正常捕获UI异常，需要：项目，属性，应用程序，取消“启用应用程序框架”，将“启动对象”设置为自定义的Main函数（参考本函数库中的Program.vb文件）  
06. 从 .NET Framework 2.0 版开始，将无法通过 try-catch 块捕获 StackOverflowException 对象，并且默认情况下将立即终止相应的进程，而 OutOfMemoryException 则可以捕获并处理  
07. System.Drawing.Imaging.ImageFormat的图片保存质量及文件大小降序排列，实测结果：  
    Bmp（最大）> Tiff > Exif/Icon/MemoryBmp > Png/Emf/Wmf（默认） > Gif > Jpeg（最小）  
08. VB.NET中，SendKeys函数不能模拟发送PrintScreen键（全屏截图），必须使用底层的keybd_event函数实现才可以：  
    My.Computer.Keyboard.SendKeys(Keys.PrintScreen) '内置函数，无效  
    System.Windows.Forms.SendKeys.Send(Keys.PrintScreen) '内置函数，无效  
    My.Keyboard.Click(Keys.PrintScreen) '本函数库，有效  
09. 在Windows中，底层的keybd_event函数，也不能发送某些（跳转到当前用户的界面之外的）特殊组合键：  
    My.Keyboard.Click(Keys.LWin, Keys.D) 'Win+D 显示桌面，有效  
    My.Keyboard.Click(Keys.LWin, Keys.L) 'Win+L 锁定电脑，无效  
    My.Keyboard.Click(Keys.ControlKey, Keys.ShiftKey, Keys.Escape) 'Ctrl+Shift+Esc 打开任务管理器，有效  
    My.Keyboard.Click(Keys.ControlKey, Keys.Menu, Keys.Delete) 'Ctrl+Alt+Delete 跳转系统界面，无效  
10. 无法模拟“Win+L”的问题，本函数库提供了一个替代方案，调用“user32.dll”中的“LockWorkStation”：  
    My.Power.Lock() '锁定电脑，有效  
11. VB.NET中，需要将函数指针作为参数传递时，可以用“Delegate Function”定义一个函数类型，然后用“AddressOf”获得函数的指针  
12. VB.NET中，调用.dll文件时，Alias后的函数名才是.dll中真正起作用的函数的名称，Alias不存在时，才会寻找同名函数  
13. VB.NET中，频繁修改窗体内容（如修改背景图片），会导致内存泄露和卡顿闪烁的问题，优化方案：  
	If Not BackgroundImage Is Nothing Then BackgroundImage.Dispose() '设置新的背景图片之前，判断其是否为空，销毁旧的图片  
    SetStyle(ControlStyles.OptimizedDoubleBuffer, True) '首先在缓冲区中绘制，而不是直接绘制到屏幕上，这样可以减少闪烁  
	补充方案：
    System.GC.Collect() '在适当的时机和代码位置，强制进行即时垃圾回收  
    SetStyle(ControlStyles.AllPaintingInWmPaint, True) '忽略擦除背景的窗口消息，不擦除之前的背景，以减少闪烁  
  
## [示例]
	'创建快捷方式  
    My.IO.WriteLinkFile("MyVisualBasic.exe", "快捷方式名称.lnk", "参数", "描述")  
	
    '将字符串转换为类似01010101的二进制形式的字符串，并写入到txt文件中  
    My.IO.WriteString(My.Security.Binary_Encode("字符串"), "binary.txt")  
    
	'获取当前目录下所有的文件列表，并保存到txt文件中  
    My.IO.WriteStringArray(My.IO.ListFile(), "list.txt")  
    
	'获取网页源码，并分离出其中所有的href属性值，返回字符串数组  
    My.StringProcessing.FindAll(My.Http.GetString("http://www.baidu.com"), "href=""", """")  

	'打开和关闭记事本程序，在cmd窗口同步阻塞  
	My.Task.RunAsync("notepad") : My.Task.Run("cmd") : My.Task.KillAsync("notepad")  

	'将完整的屏幕截图保存为png文件，并将60%比例的屏幕缩略图保存为jpg文件  
	My.Screen.Image().Save("100.png") : My.Screen.ImageThumbnail(0.6).Save("60.jpg", Imaging.ImageFormat.Jpeg)  

	'模拟键盘敲击，发送组合键：切换输入法Ctrl+Shift，关闭当前窗口Alt+F4，QQ屏幕截图Ctrl+Alt+A  
	My.Keyboard.Click(Keys.ControlKey, Keys.ShiftKey)  
    My.Keyboard.Click(Keys.Menu, Keys.F4)  
    My.Keyboard.Click(Keys.ControlKey, Keys.Menu, Keys.A)  

	'模拟键盘敲击，输入一段字符串，输入每个字符的时间间隔为50毫秒  
	My.Keyboard.Input("1!2@3#4$5%6^7&8*9(0)-_=+Aa", 50)  

	'模拟连续复制粘贴字符，输入一段字符串，输入每个字符的时间间隔为100毫秒  
	My.Keyboard.PasteDelay("这是一段中文字符。", 100)  

	'模拟用户操作，打开“画图”程序，粘贴截图，并将文件保存到桌面，关闭“画图”程序  
    Dim Screenshot As Bitmap = My.Screen.Image()  
    Dim SavePath As String = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory)  
    My.Task.RunAsync("mspaint.exe")  
    My.Time.Wait(0.5)  
    My.Keyboard.Paste(Screenshot)  
    My.Keyboard.Click(Keys.ControlKey, Keys.S)  
    My.Time.Wait(0.5)  
    My.Keyboard.Paste(SavePath & "\截图" & My.Time.Stamp() & ".png")  
    My.Keyboard.Click(Keys.Enter)  
    My.Time.Wait(0.5)  
    My.Task.KillAsync("mspaint.exe")  

	'模拟用户操作，移动鼠标到桌面右下角（显示桌面），单击2下，并将鼠标移回初始位置  
    Dim Position As Point = My.Mouse.Position()  
    My.Mouse.MoveToPercent(1, 1)  
    My.Mouse.LeftClick()  
    My.Time.Wait(1)  
    My.Mouse.LeftClick()  
    My.Mouse.MoveToPosition(Position)  