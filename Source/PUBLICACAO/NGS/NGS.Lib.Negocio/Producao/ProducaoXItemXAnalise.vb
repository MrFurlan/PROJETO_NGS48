<Serializable()> _
Public Class ProducaoXItemXAnalise

#Region "Builders"

    Public Sub New()

    End Sub

    Public Sub New(ByVal producao As Integer, ByVal produto As String, ByVal analise As Integer)
        Dim sql As String = " SELECT Producao_id, Produto_Id, Analise_Id, Etapa, Quantidade, Indice" & vbCrLf & _
                            "    FROM NewProducao                                                  " & vbCrLf

        Dim ds As DataSet = New AcessaBanco().ConsultaDataSet(sql, "ProducaoXItem")

        If ds IsNot Nothing AndAlso ds.Tables(0).Rows.Count > 0 Then
            Dim obj As New Producao()
            For Each row As DataRow In ds.Tables(0).Rows
                With obj
                    .CodigoProducao = row("Producao_id")
                    .UnidadeDeNegocio = row("UnidadeDeNegocio")
                    .CodigoEmpresa = row("Empresa")
                    .EndEmpresa = row("EndEmpresa")
                    .CodigoDeposito = row("Deposito")
                    .EndDeposito = row("EndDeposito")
                    .Movimento = row("Movimento")
                    .Etapa = row("Etapa")
                    .Safra = row("Safra")
                    .Observacao = row("Observacao")
                End With
            Next
        End If

    End Sub

#End Region

#Region "Fields"

    Private _CodigoProducao As Integer
    Private _CodigoProduto As String
    Private _CodigoAnalise As Integer
    Private _Etapa As Integer
    Private _Quantidade As Decimal
    Private _Indice As String

#End Region

#Region "Methods"

    Public Property CodigoProducao() As Integer
        Get
            Return _CodigoProducao
        End Get
        Set(ByVal value As Integer)
            _CodigoProducao = value
        End Set
    End Property

    Public Property CodigoProduto() As String
        Get
            Return _CodigoProduto
        End Get
        Set(ByVal value As String)
            _CodigoProduto = value
        End Set
    End Property

    Public Property CodigoAnalise() As Integer
        Get
            Return _CodigoAnalise
        End Get
        Set(ByVal value As Integer)
            _CodigoAnalise = value
        End Set
    End Property

    Public Property Etapa() As Integer
        Get
            Return _Etapa
        End Get
        Set(ByVal value As Integer)
            _Etapa = value
        End Set
    End Property

    Public Property Quantidade() As Decimal
        Get
            Return _Quantidade
        End Get
        Set(ByVal value As Decimal)
            _Quantidade = value
        End Set
    End Property

    Public Property Indice() As Decimal
        Get
            Return _Indice
        End Get
        Set(ByVal value As Decimal)
            _Indice = value
        End Set
    End Property

#End Region

End Class
