Public Interface IProdutoRepositorio
    Function GetAll() As IEnumerable(Of Produto)
    Function GetProduct(id As Integer) As Produto
    Function Add(item As Produto) As Produto
    Sub Remove(id As Integer)
    Function Update(item As Produto) As Boolean
End Interface
