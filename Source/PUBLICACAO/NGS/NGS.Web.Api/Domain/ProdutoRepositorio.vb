Public Class ProdutoRepositorio
    Implements IProdutoRepositorio

    Private products As New List(Of Produto)()
    Private _nextId As Integer = 1

    Public Sub New()
        Add(New Produto With {.Nome = "Tomato soup", .Categoria = "Groceries", .Preco = 1.39D})
        Add(New Produto With {.Nome = "Yo-yo", .Categoria = "Toys", .Preco = 3.75D})
        Add(New Produto With {.Nome = "Hammer", .Categoria = "Hardware", .Preco = 16.99D})
    End Sub

    Public Function GetAll() As IEnumerable(Of Produto) Implements IProdutoRepositorio.GetAll
        Return products
    End Function

    Public Function GetProduct(id As Integer) As Produto Implements IProdutoRepositorio.GetProduct
        Return products.Find(Function(p) p.Id = id)
    End Function

    Public Function Add(item As Produto) As Produto Implements IProdutoRepositorio.Add
        If item Is Nothing Then
            Throw New ArgumentNullException(NameOf(item))
        End If
        item.Id = _nextId
        _nextId += 1
        products.Add(item)
        Return item
    End Function

    Public Sub Remove(id As Integer) Implements IProdutoRepositorio.Remove
        products.RemoveAll(Function(p) p.Id = id)
    End Sub

    Public Function Update(item As Produto) As Boolean Implements IProdutoRepositorio.Update
        If item Is Nothing Then
            Throw New ArgumentNullException(NameOf(item))
        End If
        Dim index As Integer = products.FindIndex(Function(p) p.Id = item.Id)
        If index = -1 Then
            Return False
        End If
        products.RemoveAt(index)
        products.Add(item)
        Return True
    End Function
End Class
