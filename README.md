# MyVisualBasic
A function library to extend the My namespace of VB.NET.  
  
## [简介]
扩展 VB.NET 中的My命名空间的函数库  
分类整理了常用函数，对Win32 API，VBScript，Batch等进行了封装  
默认字符编码：UTF-8  
项目需要引用：System，System.Drawing，System.Web，System.Windows.Forms  
  
## [用法]
使用本函数库：引入【MyVisualBisic】文件夹中的.vb文件，获得函数扩展  
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
</table>
  
## [参考]
- 引入方法：解决方案资源管理器，显示所有文件，刷新，右键“.vb”文件，包括在项目内  
- 空行规则：同一方面的函数放在同一源文件，重载函数不用空行隔开，相关函数隔开一行，无关函数隔开三行  
- 注释规则：“结果字符串（失败返回空字符串）”，“结果Byte数组（失败返回空Byte数组）”，“使用特定的字符编码（默认UTF-8）”  
  
## [教程]
1. 在Visual Studio中显示行号：工具，选项，文本编辑器，所有语言，显示行号  
2. VB.NET中，数组与ArrayList互相转换：(New ArrayList(New String() {})).ToArray(String.Empty.GetType())  
3. VB.NET中，数组元素个数的声明与其它语言不同，Dim Array(2) As String : MsgBox(Array.Length)，输出3  
4. VB.NET中，双引号使用两个双引号来转义替代，如""""表示1个双引号的字符串，字符串用&符号来连接  
5. VB.NET中，想要在“调试”状态下，程序也能正常捕获UI异常，需要：项目，属性，应用程序，取消“启用应用程序框架”，然后用自定义的Main函数启动程序（参考Program.vb）  
6. 从 .NET Framework 2.0 版开始，将无法通过 try-catch 块捕获 StackOverflowException 对象，并且默认情况下将立即终止相应的进程  
  
## [示例]
	'创建快捷方式  
    My.IO.WriteLinkFile("MyVisualBasic.exe", "快捷方式名称.lnk", "参数", "描述")  
	
    '将字符串转换为类似01010101的二进制形式的字符串，并写入到txt文件中  
    My.IO.WriteString(My.Security.Binary_Encode("字符串"), "binary.txt")  
    
	'获取当前目录下所有的文件列表，并保存到txt文件中  
    My.IO.WriteStringArray(My.IO.ListFile(), "list.txt")  
    
	'获取网页源码，并分离出其中所有的href属性值，返回字符串数组  
    My.StringProcessing.FindAll(My.Http.GetString("http://www.baidu.com"), "href=""", """")  