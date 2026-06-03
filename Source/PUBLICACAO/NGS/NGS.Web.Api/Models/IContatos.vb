
Public Interface IContatosRepositorio

    Sub Adicionar(item As Contato)
    Function GetTodos() As IEnumerable(Of Contato)
    Function Encontrar(chave As String) As Contato
    Sub Remover(Id As String)
    Sub Atualizar(item As Contato)

End Interface
