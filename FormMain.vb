Public Class FormMain

    Private Sub FormMain_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        'MsgBox(My.Html.GetTextByTagName("<html><div>123</div><div>456</div></html>", "html")(0))
        'MsgBox(My.Html.GetTextById("<html><div id=" & """" & "a" & """" & ">123</div><div>456</div></html>", "a"))
        Dim Temp As String = My.Http.GetString("http://dy.www.yxdown.com/news/tag/listmore.json?callback=viewMoreCallback&key=erciyuan&previd=168308")
        My.Http.DownloadFile("http://dy.www.yxdown.com/news/tag/listmore.json?callback=viewMoreCallback&key=erciyuan&previd=2000000", "1.HTML")
        Dim List As List(Of String) = My.StringData.SearchAllForward(Temp, "a href=\" & """", "\" & """")
        My.IO.SaveStringArray(List, "2.HTML")
        MsgBox(List.Count)
        For I = 214 To List.Count - 1
            If I Mod 20 = 0 Then MsgBox(I)
            My.Http.DownloadFile(List(I), "Save\" & List(I).Replace("/", "_").Replace(":", "_"))
        Next
    End Sub
End Class
