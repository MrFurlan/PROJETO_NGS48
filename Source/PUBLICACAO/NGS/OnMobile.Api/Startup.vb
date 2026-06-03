Imports Hangfire
Imports Hangfire.Dashboard
Imports Hangfire.SqlServer
Imports Microsoft.Owin
Imports OnMobile.Api
Imports OnMobile.Api.Controllers
Imports Owin

<Assembly: OwinStartup(GetType(Global.Startup))>

Namespace Global
    Public Class Startup
        <Obsolete>
        Public Sub Configuration(app As IAppBuilder)

            'GlobalConfiguration.Configuration.UseSqlServerStorage(ConfigurationManager.ConnectionStrings("baxi").ConnectionString)
            'GlobalConfiguration.Configuration.UseSqlServerStorage("Data Source=MRFURLAN; Initial Catalog=Baxi; User Id=sa; Password=pwd_ngs123")
            GlobalConfiguration.Configuration.UseSqlServerStorage("Data Source=SERVIDOR-POWER-; Initial Catalog=Verde; User Id=sa; Password=PY_pwd@Ove312")

            ' Configuração do Dashboard com autorização personalizada
            Dim dashboardOptions As New DashboardOptions() With {
                .AsyncAuthorization = {New DashboardAuthorizationFilter()},
                .AppPath = VirtualPathUtility.ToAbsolute("~"),
                .Authorization = {
                                     New BasicAuthAuthorizationFilter(New BasicAuthAuthorizationFilterOptions With {
                                    .SslRedirect = True,
                                    .RequireSsl = False,
                                    .LoginCaseSensitive = True,
                                    .Users = {New BasicAuthAuthorizationUser With {
                                                    .Login = "admin",
                                                    .PasswordClear = "admin"
                                                                                    }
                                    }
                                         })}}

            'Dim options = New DashboardOptions With {.AppPath = VirtualPathUtility.ToAbsolute("~")}
            app.UseHangfireDashboard("/hangfire", dashboardOptions)
            app.UseHangfireServer()

            RecurringJob.AddOrUpdate(Function() OnMobileController.Sincronizar(), Cron.HourInterval(3))
            RecurringJob.AddOrUpdate(Function() OnMobileController.SincronizarPedido(), Cron.Hourly)

        End Sub

    End Class
End Namespace