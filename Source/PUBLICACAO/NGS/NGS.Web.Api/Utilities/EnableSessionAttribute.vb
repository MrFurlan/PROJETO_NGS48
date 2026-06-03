Imports System.Web
Imports System.Web.Http.Controllers
Imports System.Web.Http.Filters

Public Class EnableSessionAttribute
    Inherits ActionFilterAttribute

    Public Overrides Sub OnActionExecuting(actionContext As HttpActionContext)
        If Not HttpContext.Current.Session IsNot Nothing Then
            Throw New InvalidOperationException("HttpContext.Current.Session is null. HttpContext.SetSessionStateBehavior must be called before the HttpApplication.AcquireRequestState event is raised.")
        End If

        MyBase.OnActionExecuting(actionContext)
    End Sub
End Class
