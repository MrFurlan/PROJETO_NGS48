Public Class Mensagem
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        txtMensagem.Text = HttpContext.Current.Session("ssMessage")
    End Sub

    Protected Sub cmdOk_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim jscript As String
        jscript = "<script language='javascript' type='text/javascript'>"
        jscript += ";self.close();"
        jscript += "</script>"
        ScriptManager.RegisterStartupScript(Me, Me.UpdatePanel1.GetType(), Guid.NewGuid().ToString(), jscript, False)
    End Sub

End Class