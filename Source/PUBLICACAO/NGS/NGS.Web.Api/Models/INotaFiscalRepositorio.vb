Imports NGS.Lib.Negocio

Public Interface INotaFiscalRepositorio
    Function EmitirNotaFiscal(ByVal pNomeArquivo As String, ByVal usarProdutoXML As Boolean, ByVal notaDeTerceiro As Boolean, ByVal pedido As String) As String
End Interface
