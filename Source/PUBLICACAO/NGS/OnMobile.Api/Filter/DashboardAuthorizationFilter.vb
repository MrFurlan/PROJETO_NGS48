Imports Hangfire.Dashboard
Imports Microsoft.AspNetCore.Mvc.Filters
Imports System.Threading.Tasks

Public Class DashboardAuthorizationFilter
    Implements IDashboardAuthorizationFilter, IDashboardAsyncAuthorizationFilter

    Public Function Authorize(context As DashboardContext) As Boolean Implements IDashboardAuthorizationFilter.Authorize
        Return True
    End Function

    Public Async Function AuthorizeAsync(context As DashboardContext) As Task(Of Boolean) Implements IDashboardAsyncAuthorizationFilter.AuthorizeAsync
        Return Await Task.FromResult(True)
    End Function
End Class
