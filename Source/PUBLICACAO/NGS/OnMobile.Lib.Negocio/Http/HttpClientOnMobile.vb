Imports System.Configuration
Imports System.Net.Http
Imports System.Net.Http.Headers
Imports System.Text
Imports Newtonsoft.Json

Public Class HttpClientOnMobile
    Public Property _BaseUrl As String = ConfigurationManager.AppSettings("baseUrl")
    Public Property _Token As String = ConfigurationManager.AppSettings("onMobileToken")
    Public Async Function PostAsync(ByVal path As String, ByVal body As Object) As Task

        'Dim log As New LogPedidoIntegracaoOnMobile
        'log.Data = Date.Now
        'log.Endpoint = path
        'log.Request = JsonConvert.SerializeObject(body)
        'log.Metodo = "POST"

        Try
            Dim client As New HttpClient
            Dim data = JsonConvert.SerializeObject(body)

            Dim buffer = Encoding.UTF8.GetBytes(data)
            Dim bytes = New ByteArrayContent(buffer)
            bytes.Headers.ContentType = New Headers.MediaTypeHeaderValue("application/json")

            client.DefaultRequestHeaders.Authorization = New AuthenticationHeaderValue("Bearer", _Token)

            Dim response = Await client.PostAsync($"{_BaseUrl}/{path}", bytes)
            Dim stringContent As String = response.Content.ReadAsStringAsync().Result

            'Log.Status = response.StatusCode
            'Log.Response = JsonConvert.SerializeObject(stringContent.Replace("'", ""))
            'Log.Salvar()
        Catch ex As Exception
            'log.Status = 500
            'log.Erro = ex.Message
            'log.Salvar()
            Throw ex
        End Try



    End Function

    Public Async Function GetPedidoOnSolftAsync() As Task(Of PedidoOnSoftResponse)
        'Dim log As New LogPedidoIntegracaoOnMobile
        'log.Data = Date.Now
        'log.Endpoint = "WsPedido?limit=1&page=1"
        'Log.Metodo = "GET"

        Try
            Dim client As New HttpClient
            Dim url = $"https://api.onmobile.com.br/WsPedido?limit=1&page=1"

            client.DefaultRequestHeaders.Authorization = New AuthenticationHeaderValue("Bearer", _Token)
            client.DefaultRequestHeaders.Accept.Add(New MediaTypeWithQualityHeaderValue("application/json"))

            Dim response = Await client.GetAsync(url)
            Dim stringContent As String = response.Content.ReadAsStringAsync().Result

            'log.Status = response.StatusCode
            'og.Response = JsonConvert.SerializeObject(stringContent)

            Dim pedidosOnSoft = JsonConvert.DeserializeObject(Of PedidoOnSoftResponse)(stringContent)

            'log.Salvar()

            Return pedidosOnSoft
        Catch ex As Exception
            'log.Status = 500
            'log.Erro = ex.Message
            'log.Salvar()
            Throw New Exception(ex.Message)
        End Try

        Return Nothing
    End Function

    Public Async Function GetClientesPreCadastroOnSolftAsync() As Task(Of WsClienteIntegracao)
        'Dim log As New LogPedidoIntegracaoOnMobile
        'log.Data = Date.Now
        'log.Endpoint = "WsPedido?limit=1&page=1"
        'Log.Metodo = "GET"

        Try
            Dim client As New HttpClient
            Dim url = $"https://api.onmobile.com.br/WsClienteIntegracao?limit=1&page=1"

            client.DefaultRequestHeaders.Authorization = New AuthenticationHeaderValue("Bearer", _Token)
            client.DefaultRequestHeaders.Accept.Add(New MediaTypeWithQualityHeaderValue("application/json"))

            Dim response = Await client.GetAsync(url)
            Dim stringContent As String = response.Content.ReadAsStringAsync().Result

            'log.Status = response.StatusCode
            'og.Response = JsonConvert.SerializeObject(stringContent)

            Dim clientesPreCadastro = JsonConvert.DeserializeObject(Of WsClienteIntegracao)(stringContent)

            'log.Salvar()

            Return clientesPreCadastro
        Catch ex As Exception
            'log.Status = 500
            'log.Erro = ex.Message
            'log.Salvar()
            Throw New Exception(ex.Message)
        End Try

        Return Nothing
    End Function

    Public Async Function GetWsLoadAsync() As Task

        'Dim log As New LogPedidoIntegracaoOnMobile
        ''log.Data = Date.Now
        'log.Endpoint = "WsLoad"
        'log.Metodo = "GET"

        Try
            Dim client As New HttpClient

            client.DefaultRequestHeaders.Authorization = New AuthenticationHeaderValue("Bearer", _Token)
            client.DefaultRequestHeaders.Accept.Add(New MediaTypeWithQualityHeaderValue("application/json"))

            Dim response = Await client.GetAsync($"{_BaseUrl}/WsLoad")
            Dim stringContent As String = response.Content.ReadAsStringAsync().Result
            response.EnsureSuccessStatusCode()

            'log.Status = response.StatusCode
            'log.Response = JsonConvert.SerializeObject(stringContent)

            'log.Salvar()

        Catch ex As Exception
            'log.Status = 500
            'log.Erro = ex.Message
            'log.Salvar()
        End Try
    End Function

    Public Async Function DeleteAsync(ByVal path As String) As Task
        'Dim log As New LogPedidoIntegracaoOnMobile
        'log.Data = Date.Now
        'log.Endpoint = path
        'log.Metodo = "DELETE"

        Try
            Dim client As New HttpClient

            client.DefaultRequestHeaders.Authorization = New AuthenticationHeaderValue("Bearer", _Token)
            client.DefaultRequestHeaders.Accept.Add(New MediaTypeWithQualityHeaderValue("application/json"))

            Dim response = Await client.DeleteAsync($"{_BaseUrl}/{path}")
            Dim stringContent As String = response.Content.ReadAsStringAsync().Result


            Dim e = response.StatusCode
            'log.Status = response.StatusCode
            'log.Response = JsonConvert.SerializeObject(stringContent)
            'log.Salvar()
        Catch ex As Exception
            'log.Status = 500
            'log.Erro = ex.Message
            'log.Salvar()
        End Try
    End Function
End Class
