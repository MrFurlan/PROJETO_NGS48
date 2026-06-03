Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Web.Http
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Serialization

Public Module WebApiConfig
    Public Sub Register(ByVal config As HttpConfiguration)
        ' Web API configuration and services

        ' Configuração do JSON Formatter
        Dim jsonFormatter = config.Formatters.JsonFormatter
        jsonFormatter.SerializerSettings.ContractResolver = New CamelCasePropertyNamesContractResolver()
        jsonFormatter.SerializerSettings.Formatting = Formatting.Indented

        ' Web API routes
        config.MapHttpAttributeRoutes()

        config.Routes.MapHttpRoute(
            name:="DefaultApi",
            routeTemplate:="api/{controller}/{id}",
            defaults:=New With {.id = RouteParameter.Optional}
        )

        ' Remover o XmlFormatter para suportar apenas JSON
        config.Formatters.Remove(config.Formatters.XmlFormatter)

        ' Adicionar filtro de autorização global
        ' Tenha cuidado ao adicionar isso globalmente, pois todas as rotas exigirão autenticação
        config.Filters.Add(New AuthorizeAttribute())

    End Sub
End Module
