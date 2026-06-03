Imports System.Web.Http
Imports System.Web.Optimization

Public Class WebApiApplication
    Inherits System.Web.HttpApplication

    Sub Application_Start()

        AreaRegistration.RegisterAllAreas()
        GlobalConfiguration.Configure(AddressOf WebApiConfig.Register)
        FilterConfig.RegisterGlobalFilters(GlobalFilters.Filters)
        RouteConfig.RegisterRoutes(RouteTable.Routes)
        BundleConfig.RegisterBundles(BundleTable.Bundles)

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

    Sub Session_Start(ByVal sender As Object, ByVal e As EventArgs)
        'Response.Redirect("~/Index.aspx")
        If HttpContext.Current.Session("ssKeyCodeActive") Is Nothing Then
            HttpContext.Current.Session("ssKeyCodeActive") = False
        End If
    End Sub

    Sub Session_End(ByVal sender As Object, ByVal e As EventArgs)
        ' Fires when the session ends
    End Sub

End Class
