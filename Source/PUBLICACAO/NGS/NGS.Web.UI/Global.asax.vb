Imports System.IO
Imports System.Web
Imports System.Web.Security
Imports System.Web.SessionState
Imports System.Web.Routing
Imports System.Xml
Imports NGS.Lib.Uteis

Public Class Global_asax
    Inherits System.Web.HttpApplication

#Region "Eventos"

    Sub Application_Start(ByVal sender As Object, ByVal e As EventArgs)
        'RegistrarNomesAmigaveis(RouteTable.Routes)
        If Not Directory.Exists(Server.MapPath("~/Files")) Then Directory.CreateDirectory(Server.MapPath("~/Files"))
        If Not Directory.Exists(Server.MapPath("~/Conecta")) Then Directory.CreateDirectory(Server.MapPath("~/Conecta"))
        If Not Directory.Exists(Server.MapPath("~/FotosClientes")) Then Directory.CreateDirectory(Server.MapPath("~/FotosClientes"))
        If Not Directory.Exists(Server.MapPath("~/PlanilhasExcel")) Then Directory.CreateDirectory(Server.MapPath("~/PlanilhasExcel"))
        If Not Directory.Exists(Server.MapPath("~/Receituario")) Then Directory.CreateDirectory(Server.MapPath("~/Receituario"))
        If Not Directory.Exists(Server.MapPath("~/RemessaBancaria")) Then Directory.CreateDirectory(Server.MapPath("~/RemessaBancaria"))
        If Not Directory.Exists(Server.MapPath("~/Pamcard")) Then Directory.CreateDirectory(Server.MapPath("~/Pamcard"))
        If Not Directory.Exists(Server.MapPath("~/Sintegra")) Then Directory.CreateDirectory(Server.MapPath("~/Sintegra"))
        If Not Directory.Exists(Server.MapPath("~/SpedFiscal")) Then Directory.CreateDirectory(Server.MapPath("~/SpedFiscal"))
        If Not Directory.Exists(Server.MapPath("~/SpedPisCofins")) Then Directory.CreateDirectory(Server.MapPath("~/SpedPisCofins"))
        If Not Directory.Exists(Server.MapPath(String.Format("~/Sped/Ecd/{0}", DateTime.Now.Year - 1))) Then Directory.CreateDirectory(Server.MapPath(String.Format("~/Sped/Ecd/{0}", DateTime.Now.Year - 1)))

    End Sub

    Sub Session_Start(ByVal sender As Object, ByVal e As EventArgs)
        'Response.Redirect("~/Index.aspx")
    End Sub

    Sub Application_BeginRequest(ByVal sender As Object, ByVal e As EventArgs)
        ' Fires at the beginning of each request
    End Sub

    Sub Application_AuthenticateRequest(ByVal sender As Object, ByVal e As EventArgs)
        ' Fires upon attempting to authenticate the use
    End Sub

    Sub Application_Error(ByVal sender As Object, ByVal e As EventArgs)
        ' Fires when an error occurs
    End Sub

    Sub Session_End(ByVal sender As Object, ByVal e As EventArgs)
        ' Fires when the session ends
    End Sub

    Sub Application_End(ByVal sender As Object, ByVal e As EventArgs)
        ' Fires when the application ends
    End Sub

    Sub Application_PostAuthorizeRequest()
        ' Habilitar o uso de sessões no Web API
        If IsWebApiRequest() Then
            HttpContext.Current.SetSessionStateBehavior(SessionStateBehavior.Required)
        End If
    End Sub

    Private Function IsWebApiRequest() As Boolean
        Return HttpContext.Current.Request.AppRelativeCurrentExecutionFilePath.StartsWith("~/api")
    End Function

#End Region

#Region "Métodos"

    Private Sub RegistrarNomesAmigaveis(ByVal routes As RouteCollection)
        Dim rotas As New XmlDocument
        rotas.Load(Server.MapPath("~/ScreenSchema.xml"))

        For Each r As XmlElement In rotas("rotas")
            routes.MapPageRoute(r.Attributes("id").InnerText, _
                                         r.Attributes("amigavel").InnerText, _
                                         r.Attributes("url").InnerText)
        Next
    End Sub

#End Region

End Class