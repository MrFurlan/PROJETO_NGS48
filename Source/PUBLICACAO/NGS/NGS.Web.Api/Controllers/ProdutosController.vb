Imports System.Net
Imports System.Net.Http
Imports System.Web.Http

Public Class ProdutosController
    Inherits ApiController

    Private Shared ReadOnly repository As IProdutoRepositorio = New ProdutoRepositorio()

    Public Function GetAllProdutos() As IEnumerable(Of Produto)
        Return repository.GetAll()
    End Function

    Public Function GetProdutos(id As Integer) As Produto
        Dim item As Produto = repository.GetProduct(id)
        If item Is Nothing Then
            Throw New HttpResponseException(HttpStatusCode.NotFound)
        End If
        Return item
    End Function

    <Route("api/produtos/categoria/{category}")>
    Public Function GetProdutosByCategory(category As String) As IEnumerable(Of Produto)
        Return repository.GetAll().Where(Function(p) String.Equals(p.Categoria, category, StringComparison.OrdinalIgnoreCase))
    End Function

    Public Function PostProdutos(item As Produto) As Produto
        item = repository.Add(item)
        Return item
    End Function

    Public Function PostProdutosHttp(item As Produto) As HttpResponseMessage
        item = repository.Add(item)
        Dim response As HttpResponseMessage = Request.CreateResponse(Of Produto)(HttpStatusCode.Created, item)

        Dim uri As String = Url.Link("DefaultApi", New With {Key .id = item.Id})
        response.Headers.Location = New Uri(uri)
        Return response
    End Function

    Public Sub PutProdutos(id As Integer, product As Produto)
        product.Id = id
        If Not repository.Update(product) Then
            Throw New HttpResponseException(HttpStatusCode.NotFound)
        End If
    End Sub

    Public Sub DeleteProdutos(id As Integer)
        Dim item As Produto = repository.GetProduct(id)
        If item Is Nothing Then
            Throw New HttpResponseException(HttpStatusCode.NotFound)
        End If

        repository.Remove(id)
    End Sub


    'EmitirNotaFiscal

End Class