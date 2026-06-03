
Public Class ProdutoDashboard

    ' Propriedades comuns
    Public Property EmpresaId As String
    Public Property Filtro As ProdutoDashboardFiltro
    Public Property Data As String
    Public Property ProdutoId As String
    Public Property Nome As String
    Public Property Seguimento As String
    Public Property Cidade As String
    Public Property Estado As String

    ' Propriedades cliente
    Public Property ClienteId As String
    Public Property ClienteNome As String
    Public Property ClienteFantasia As String

    ' Propriedades representate
    Public Property RepresentanteId As String
    Public Property RepresentanteNome As String
    Public Property RepresentanteFantasia As String

    ' Propriedades para Faturamento
    Public Property Valor As Decimal

    ' Propriedades para Detalhamento
    Public Property QuantidadePedido As Decimal
    Public Property QuantidadeFiscal As Decimal
    Public Property PesoFiscal As Decimal
    Public Property PesoTotal As Decimal
    Public Property ValorTotal As Decimal

    Public Property Periodo As String
    Public Property Pedidos As String
    Public Property VolumeKg As Decimal
    Public Property Faturamento As Decimal
    Public Property VolumeMedio As Decimal
    Public Property FaturamentoMedio As Decimal

End Class


