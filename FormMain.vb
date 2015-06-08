Public Class FormMain

    Private Sub FormMain_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        'CreateObject("Wscript.Shell").SendKeys("^+{ESC}")
        'keybd_event(Keys.ControlKey, 0, 0, 0)
        'keybd_event(Keys.ShiftKey, 0, 0, 0)
        'keybd_event(Keys.Escape, 0, 0, 0)
        'keybd_event(Keys.Escape, 0, KEYEVENTF_KEYUP, 0)
        'keybd_event(Keys.ShiftKey, 0, KEYEVENTF_KEYUP, 0)
        'keybd_event(Keys.ControlKey, 0, KEYEVENTF_KEYUP, 0)
        'My.Computer.SendKeys("^+{ESC}")
        Dim A As New User
        A.UserName = "ABC"
        A.Password = "123"
        Dim B As New User
        B.UserName = "DEF"
        B.Password = "456"
        A.InviteUser = B
        Dim C As New User
        B.InviteUser = C
        Dim D As New List(Of User)
        D.Add(C)
        'Clipboard.Clear()
        'Clipboard.SetText(Json.ChangeObjectToString(A))
        'MsgBox(Json.ChangeObjectToString(A))

        'MsgBox(Json.ChangeObjectToString(D))
        'MsgBox(Json.ChangeObjectToString(New User() {C, C}))
        'MsgBox(Json.ChangeObjectToString(New String() {"123", "456", "789"}))

        Me.Visible = False
        'MsgBox(System.Windows.Forms.Cursor.Position.X)
        'MsgBox(System.Windows.Forms.Cursor.Position.Y)
        System.Windows.Forms.Cursor.Position.Offset(New Point(1000, 1000))
        'MsgBox(Control.MousePosition.X)
        'MsgBox(Control.MousePosition.Y)
    End Sub

    Private Sub Timer1_Tick(sender As System.Object, e As System.EventArgs) Handles Timer1.Tick
        Dim W As System.Drawing.Rectangle = My.Computer.FindWindow("内存清理.bat - 记事本")
        W = My.Computer.FindFocusWindow
        Console.Write("左上角坐标(" & W.Left & "," & W.Top & ")" & vbCrLf & "右下角坐标(" & W.Right & "," & W.Bottom & ")" & vbCrLf & "窗口高" & W.Height & "窗口宽" & W.Width)
    End Sub
    Public Class User
        Public Property UserName As String = ""
        Public Property Password As String = ""
        Public Property InviteUser As User = Nothing
        Public Property Boy As Boolean = Nothing
        Public Property Age As Integer = Nothing
        Public Property Time As DateTime = Now
    End Class

    Public Class Json
        '栈溢出错误（如果Object的属性互相引用，则会发生）
        Public Shared Function ChangeObjectToString(ByRef JsonObject As Object) As String
            'MsgBox(JsonObject.GetType.BaseType.ToString)
            MsgBox(JsonObject.GetType.ToString)
            If JsonObject.GetType.IsPrimitive Or JsonObject.GetType Is GetType(System.String) Or JsonObject.GetType Is GetType(System.DateTime) Then
                Return """" & JsonObject.ToString & """"
            End If
            'IsGenericType包含了类似List(Of Object)的各种类型
            If JsonObject.GetType.IsArray Or JsonObject.GetType.IsGenericType Or JsonObject.GetType Is GetType(ArrayList) Then
                Dim Result As New System.Text.StringBuilder("[")
                For Each Element As Object In JsonObject
                    Result.Append(ChangeObjectToString(Element) & ",")
                Next
                If Result.Length > 1 Then Result.Remove(Result.Length - 1, 1)
                Result.Append("]")
                Return Result.ToString()
            Else
                Dim Result As New System.Text.StringBuilder("{")
                For Each P As System.Reflection.PropertyInfo In JsonObject.GetType().GetProperties()
                    Result.Append("""" & P.Name & """" & ":")
                    Dim Element As Object = P.GetValue(JsonObject, Nothing)
                    If Element Is Nothing Then
                        Result.Append("null" & ",")
                    Else
                        Dim ElementType As System.Type = Element.GetType()
                        'IsPrimitive，基元类型，包含：Boolean、Byte、SByte、Int16、UInt16、Int32、UInt32、Int64、UInt64、Char、Double和Single
                        If ElementType.IsPrimitive Or ElementType Is GetType(System.String) Or ElementType Is GetType(System.DateTime) Then
                            Result.Append("""" & P.GetValue(JsonObject, Nothing).ToString() & """" & ",")
                        Else
                            Result.Append(ChangeObjectToString(P.GetValue(JsonObject, Nothing)) & ",")
                        End If
                    End If
                Next
                If Result.Length > 1 Then Result.Remove(Result.Length - 1, 1)
                Result.Append("}")
                Return Result.ToString()
            End If
        End Function
    End Class

End Class
