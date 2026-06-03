Imports System.Security.Claims
Imports System.Text
Imports System.Web.Http
Imports System.Web
Imports System.Web.Services
Imports System.Net.Http
Imports System.Net
Imports NGS.Lib.Negocio

Public Class NotaFiscalController
    Inherits ApiController

    Private Shared ReadOnly repository As INotaFiscalRepositorio = New NotaFiscalRepositorio()
    Private ReadOnly _httpContext As HttpContextBase

    Public Sub New()

    End Sub

    Public Sub New(httpContext As HttpContextBase)
        _httpContext = httpContext
    End Sub

    ' GET api/values
    <AllowAnonymous>
    <Route("api/NotaFiscal")>
    Public Function GetValues() As IEnumerable(Of String)
        Return New String() {"value1", "value2"}
    End Function

    'http://localhost:44323/api/NotaFiscal/EmitirNotaFiscal?xml=1%20-%202548_E8_14691
    '<Authorize>
    <HttpGet>
    <AllowAnonymous>
    <Route("api/NotaFiscal/EmitirNotaFiscal")>
    Public Function EmitirNotaFiscal(<FromUri> xml As String, <FromUri> usarProdutoXML As Boolean, <FromUri> notaDeTerceiro As Boolean, <FromUri> pedido As String) As IHttpActionResult

        Try

            ' Your logic here
            Dim identity As ClaimsIdentity = TryCast(User.Identity, ClaimsIdentity)
            Dim username As String = identity?.FindFirst(ClaimTypes.Name)?.Value
            'HttpContext.Current.Session("UserName") = chaveNota
            UsuarioServidor.KeyCodeActive = True

            AcessaBanco()
            Return Ok(Newtonsoft.Json.JsonConvert.SerializeObject(New With {
                .Success = True,
                .Message = "Nota fiscal emitida com sucesso.",
                .Data = repository.EmitirNotaFiscal(xml, usarProdutoXML, notaDeTerceiro, pedido)
            }))


        Catch ex As Exception

            Dim errorResponse = New With {
                .Message = "Erro interno do servidor.",
                .ExceptionMessage = ex.Message,
                .ExceptionType = ex.GetType().Name,
                .StackTrace = ex.StackTrace
            }

            Return Content(HttpStatusCode.InternalServerError, errorResponse)


        End Try

    End Function

    'http://localhost:44323/api/NotaFiscal/EmitirNotaFiscal?xml=1%20-%202548_E8_14691
    '<Authorize>
    <HttpGet>
    <AllowAnonymous>
    <Route("api/NotaFiscal/EmitirNotaFiscal_Teste")>
    Public Function EmitirNotaFiscal_Teste(<FromUri> xml As String) As IHttpActionResult
        If String.IsNullOrEmpty(xml) Then
            Return BadRequest(Newtonsoft.Json.JsonConvert.SerializeObject(New With {.Message = "XML is required."}))
        End If

        ' Simule o processamento do XML
        Dim isProcessed As Boolean = ProcessXml(xml)

        If isProcessed Then
            Return Ok(Newtonsoft.Json.JsonConvert.SerializeObject(New With {
                .Success = True,
                .Message = "Nota fiscal emitida com sucesso.",
                .Data = New With {
                    .NotaFiscalId = "12345",
                    .Status = "Emitida"
                }
            }))
        Else
            Return BadRequest(Newtonsoft.Json.JsonConvert.SerializeObject(New With {
                .Success = False,
                .Message = "Erro ao emitir nota fiscal.",
                .ErrorCode = "ERR001"
            }))
        End If
    End Function

    Private Function ProcessXml(xml As String) As Boolean
        ' Simule a lógica de processamento do XML
        Return True ' Simule o processamento bem-sucedido
    End Function

    <AllowAnonymous>
    <Route("api/NotaFiscal/AcessaBanco/")>
    Public Function AcessaBanco() As IHttpActionResult
        Try
            ' Leia os dados do arquivo (JSON ou XML)
            ' Para JSON:
            Dim data As DataModel = DataAcess.ReadDataFromXmlFile()
            ' Para XML:
            ' Dim dataList As List(Of NotaFiscal) = ReadDataFromXmlFile()

            ' Retorne os dados lidos
            Return Ok(data)
        Catch ex As Exception
            Return InternalServerError(ex)
        End Try
    End Function

End Class
