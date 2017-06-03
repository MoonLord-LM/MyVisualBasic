# MyVisualBasic
A function library to extend the My namespace of VB.NET.  
  
## [简介]
扩展 VB.NET 中的My命名空间的函数库  
对Win32 API，VBScript，Batch等进行了封装，并分类整理了常用的函数功能  
默认字符编码：UTF-8  
项目需要引用：System，System.Drawing，System.Web，System.Windows.Forms  
  
## [用法]
引入【MyVisualBisic】文件夹中的.vb文件，获得函数扩展  
各个.vb文件之间互不依赖，可根据需要选择使用  
  
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
</table>
  
## [参考]
- 引入方法：解决方案资源管理器，显示所有文件，刷新，右键“.vb”文件，包括在项目内  
- 空行规则：同一方面的函数放在同一源文件中，重载函数不用空行隔开，相关函数隔开一行，无关函数隔开三行  
- 注释示例：“结果字符串（失败返回空字符串）”，“结果Byte数组（失败返回空Byte数组）”，“使用特定的字符编码（默认UTF-8）”  
  
## [教程]
1. 在Visual Studio中显示行号：工具，选项，文本编辑器，所有语言，显示行号  
2. VB.NET中，数组与ArrayList互相转换：(New ArrayList(New String() {})).ToArray(String.Empty.GetType())  
3. VB.NET中，数组元素个数的声明与其它语言不同，Dim Array(2) As String : MsgBox(Array.Length)，输出3  
4. VB.NET中，双引号使用两个双引号来转义替代，如""""表示1个双引号的字符串，字符串用&符号来连接  
5. VB.NET中，想要在“调试”状态下，程序也能正常捕获UI异常，需要：项目，属性，应用程序，取消“启用应用程序框架”，然后用自定义的Main函数启动程序（参考本函数库中的Program.vb）  
6. 从 .NET Framework 2.0 版开始，将无法通过 try-catch 块捕获 StackOverflowException 对象，并且默认情况下将立即终止相应的进程  
7. System.Drawing.Imaging.ImageFormat的图片保存质量及文件大小降序排列，实测结果：  
    Bmp（最大）> Tiff > Exif/Icon/MemoryBmp > Png/Emf/Wmf（默认） > Gif > Jpeg（最小）  
8. VB.NET中，SendKeys函数不能模拟发送PrintScreen键（全屏截图），必须使用底层的keybd_event函数实现才可以：  
   My.Computer.Keyboard.SendKeys(Keys.PrintScreen) '内置函数，无效  
   System.Windows.Forms.SendKeys.Send(Keys.PrintScreen) '内置函数，无效  
   My.Keyboard.Click(Keys.PrintScreen) '本函数库，有效  
9. VB.NET中，底层的keybd_event函数，也不能发送某些（跳转到当前用户的界面之外的）特殊组合键：  
   My.Keyboard.Click(New Keys() {Keys.LWin, Keys.D}) 'Win+D 显示桌面，有效  
   My.Keyboard.Click(New Keys() {Keys.LWin, Keys.L}) 'Win+L 锁定用户，无效  
   My.Keyboard.Click(New Keys() {Keys.ControlKey, Keys.ShiftKey, Keys.Escape}) 'Ctrl+Shift+Esc 打开任务管理器，有效  
   My.Keyboard.Click(New Keys() {Keys.ControlKey, Keys.Menu, Keys.Delete}) 'Ctrl+Alt+Delete 跳转系统界面，无效  
  
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
	My.Screen.Image().Save("100.png") : My.Screen.Thumbnail(0.6).Save("60.jpg", Imaging.ImageFormat.Jpeg)  

	'模拟键盘敲击，发送组合键：切换输入法Ctrl+Shift，关闭当前窗口Alt+F4，QQ屏幕截图Ctrl+Alt+A  
	My.Keyboard.Click(New Keys() {Keys.ControlKey, Keys.ShiftKey})  
    My.Keyboard.Click(New Keys() {Keys.Menu, Keys.F4})  
    My.Keyboard.Click(New Keys() {Keys.ControlKey, Keys.Menu, Keys.A})  

	'模拟键盘敲击，输入一段字符串，输入每个字符的时间间隔为50毫秒  
	My.Keyboard.Input("1!2@3#4$5%6^7&8*9(0)-_=+Aa", 50)  

	'模拟连续复制粘贴字符，输入一段字符串，输入每个字符的时间间隔为100毫秒  
	My.Keyboard.Paste("这是一段中文字符。", 100)  