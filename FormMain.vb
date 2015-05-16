Public Class FormMain

    Private Sub FormMain_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        MsgBox(My.Security.RSA_Encode("你妹啊").Length)
        MsgBox(My.Security.RSA_Decode(My.Security.RSA_Encode("你妹啊")))
    End Sub
End Class
