Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Collections.Generic
Imports NGS.Lib.Uteis

<Serializable()> _
Public Class ListPedidoFixacaoXNotaFiscal
    Inherits List(Of PedidoFixacaoXNotaFiscal)

#Region "Construtor"
    Public Sub New(pFixacao As Fixacao)
        _Fixacao = pFixacao
    End Sub
#End Region

#Region "Fields"
    Private _Fixacao As Fixacao
#End Region

#Region "Property"
    Public ReadOnly Property Fixacao As Fixacao
        Get
            Return _Fixacao
        End Get
    End Property
#End Region

#Region "Methods"
    Public Sub SalvarSql(ByVal Sqls As ArrayList)
        For Each Nota As PedidoFixacaoXNotaFiscal In Me
            If Fixacao.IUD = "I" AndAlso Not Nota.Liberada Then Exit Sub
            If Fixacao.IUD = "D" Or Fixacao.IUD = "I" Then Nota.IUD = Fixacao.IUD
            Nota.SalvarSql(Sqls)
        Next
    End Sub
#End Region

#Region "Funcoes"
    Public Sub CarregarNotasParaSelecao()
        Me.Clear()
        Dim strSQL As String = ""
        strSQL = "SELECT NFxI.Empresa_Id, NFxI.EndEmpresa_Id," & vbCrLf & _
                 "       NFxI.Cliente_Id, NFxI.EndCliente_Id," & vbCrLf & _
                 "	     NFxI.Nota_Id, NFxI.Serie_Id, NFxI.EntradaSaida_Id," & vbCrLf & _
                 "       NFxI.CFOP_Id, NFxI.Sequencia_Id," & vbCrLf & _
                 "       NF.DataDaNota," & vbCrLf & _
                 "	     NFxI.QuantidadeFiscal AS QtdeNotaFiscal," & vbCrLf & _
                 "	     NFxI.Unitario," & vbCrLf & _
                 "	     NFxI.Valor AS ValorNotaFiscal," & vbCrLf & _
                 "	     isnull(SbDev.Quantidade,0) AS QtdeDevolucao," & vbCrLf & _
                 "	     isnull(SbDev.Valor,0)      AS ValorDevolucao," & vbCrLf & _
                 "	     isnull(SbFix.Quantidade,0) AS QtdeFixacao," & vbCrLf & _
                 "	     isnull(SbFix.Valor,0)      AS ValorFixacao," & vbCrLf & _
                 "	     --isnull(SbFix.ValorFixacao,0)      AS ValorFixacaoReal," & vbCrLf & _
                 "       NFxI.QuantidadeFiscal - isnull(sbdev.Quantidade,0) - isnull(sbfix.Quantidade,0) AS QdteSaldo," & vbCrLf & _
                 "       NFxI.Valor - isnull(sbdev.Valor,0) - isnull(sbfix.Valor,0) AS ValorSaldo," & vbCrLf & _
                 "	     isnull((SELECT Percentual From NotasFiscaisXEncargos" & vbCrLf & _
                 "					WHERE Empresa_Id    = NFxI.Empresa_Id" & vbCrLf & _
                 "					AND EndEmpresa_Id   = NFxI.EndEmpresa_Id" & vbCrLf & _
                 "					AND Cliente_Id      = NFxI.Cliente_Id" & vbCrLf & _
                 "					AND EndCliente_Id   = NFxI.EndCliente_Id" & vbCrLf & _
                 "					AND EntradaSaida_Id = NFxI.EntradaSaida_Id" & vbCrLf & _
                 "					AND Serie_Id        = NFxI.Serie_Id" & vbCrLf & _
                 "					AND Nota_Id         = NFxI.Nota_Id" & vbCrLf & _
                 "					AND Produto_Id      = NFxI.Produto_Id" & vbCrLf & _
                 "					AND CFOP_Id         = NFxI.CFOP_Id" & vbCrLf & _
                 "					AND Sequencia_Id    = NFxI.Sequencia_Id" & vbCrLf & _
                 "					AND Encargo_id      = 'FACS'),0) AS PercentualFacs," & vbCrLf & _
                 "	     isnull((SELECT Percentual From NotasFiscaisXEncargos" & vbCrLf & _
                 "					WHERE Empresa_Id    = NFxI.Empresa_Id" & vbCrLf & _
                 "					AND EndEmpresa_Id   = NFxI.EndEmpresa_Id" & vbCrLf & _
                 "					AND Cliente_Id      = NFxI.Cliente_Id" & vbCrLf & _
                 "					AND EndCliente_Id   = NFxI.EndCliente_Id" & vbCrLf & _
                 "					AND EntradaSaida_Id = NFxI.EntradaSaida_Id" & vbCrLf & _
                 "					AND Serie_Id        = NFxI.Serie_Id" & vbCrLf & _
                 "					AND Nota_Id         = NFxI.Nota_Id" & vbCrLf & _
                 "					AND Produto_Id      = NFxI.Produto_Id" & vbCrLf & _
                 "					AND CFOP_Id         = NFxI.CFOP_Id" & vbCrLf & _
                 "					AND Sequencia_Id    = NFxI.Sequencia_Id" & vbCrLf & _
                 "					AND Encargo_id      = 'FETHAB'),0) AS PercentualFethab" & vbCrLf & _
                 "  FROM NotasFiscaisXItens NFxI" & vbCrLf & _
                 " INNER JOIN NotasFiscais NF" & vbCrLf & _
                 " 	  ON NFxI.Empresa_Id      = NF.Empresa_Id" & vbCrLf & _
                 "	 AND NFxI.EndEmpresa_Id   = NF.EndEmpresa_Id" & vbCrLf & _
                 "	 AND NFxI.Cliente_Id      = NF.Cliente_Id" & vbCrLf & _
                 "	 AND NFxI.EndCliente_Id   = NF.EndCliente_Id" & vbCrLf & _
                 "	 AND NFxI.EntradaSaida_Id = NF.EntradaSaida_Id" & vbCrLf & _
                 "	 AND NFxI.Serie_Id        = NF.Serie_Id" & vbCrLf & _
                 "	 AND NFxI.Nota_Id         = NF.Nota_Id" & vbCrLf & _
                 " INNER JOIN SubOperacoes SO" & vbCrLf & _
                 "    ON NF.Operacao    = SO.Operacao_Id" & vbCrLf & _
                 "	 AND NF.SubOperacao = SO.SubOperacoes_Id" & vbCrLf & _
                 "    LEFT JOIN (" & vbCrLf & _
                 "                SELECT nfd.EmpresaDevolucao_Id as Empresa_Id," & vbCrLf & _
                 "                       nfd.EndEmpresaDevolucao_Id as EndEmpresa_Id," & vbCrLf & _
                 "                       nfd.ClienteDevolucao_Id as Cliente_Id," & vbCrLf & _
                 "                       nfd.EndClienteDevolucao_Id as EndCliente_Id," & vbCrLf & _
                 "                       nfd.EntradaSaida_Id," & vbCrLf & _
                 "                       nfd.Serie_Id," & vbCrLf & _
                 "                       nfd.Nota_Id," & vbCrLf & _
                 "                       nfd.Produto_Id," & vbCrLf & _
                 "                       sum(nfd.Quantidade) as Quantidade," & vbCrLf & _
                 "                       sum(nfd.Valor) as Valor" & vbCrLf & _
                 "   			      FROM NotaFiscalDevolucaoXNotaFiscal nfd" & vbCrLf & _
                 "                 Inner Join NotasFiscais nf" & vbCrLf & _
                 "                    On nf.Empresa_id      = nfd.EmpresaDevolucao_Id" & vbCrLf & _
                 "                   and nf.EndEmpresa_Id   = nfd.EndEmpresaDevolucao_Id" & vbCrLf & _
                 "                   and nf.Cliente_Id      = nfd.ClienteDevolucao_Id" & vbCrLf & _
                 "                   and nf.EndCliente_Id   = nfd.EndClienteDevolucao_Id" & vbCrLf & _
                 "                   and nf.EntradaSaida_Id = nfd.EntradaSaida_Id" & vbCrLf & _
                 "                   and nf.Serie_Id        = nfd.Serie_Id" & vbCrLf & _
                 "                   and nf.Nota_Id         = nfd.Nota_Id" & vbCrLf & _
                 "                 Where nf.situacao        in (1,4,7) " & vbCrLf & _
                 "                   and nf.TipoDeDocumento = 1 " & vbCrLf & _
                 "                   and nf.Empresa_Id     ='" & Fixacao.ItemPedido.Pedido.CodigoEmpresa & "'" & vbCrLf & _
                 " 	                 and nf.EndEmpresa_Id  = " & Fixacao.ItemPedido.Pedido.EnderecoEmpresa & vbCrLf & _
                 "	                 and nf.Pedido         = " & Fixacao.ItemPedido.Pedido.Codigo & vbCrLf & _
                 "                 Group by nfd.EmpresaDevolucao_Id, nfd.EndEmpresaDevolucao_Id, nfd.ClienteDevolucao_Id," & vbCrLf & _
                 "                          nfd.EndClienteDevolucao_Id, nfd.EntradaSaida_Id, nfd.Serie_Id, nfd.Nota_Id," & vbCrLf & _
                 "                           nfd.Produto_Id" & vbCrLf & _
                 "               ) SbDev" & vbCrLf & _
                 "      On NFxI.Empresa_id      = SbDev.Empresa_Id" & vbCrLf & _
                 "     and NFxI.EndEmpresa_Id   = SbDev.EndEmpresa_Id" & vbCrLf & _
                 "     and NFxI.Cliente_Id      = SbDev.Cliente_Id" & vbCrLf & _
                 "     and NFxI.EndCliente_Id   = SbDev.EndCliente_Id" & vbCrLf & _
                 "     and NFxI.EntradaSaida_Id = SbDev.EntradaSaida_Id" & vbCrLf & _
                 "     and NFxI.Serie_Id        = SbDev.Serie_Id" & vbCrLf & _
                 "     and NFxI.Nota_Id         = SbDev.Nota_Id" & vbCrLf & _
                 "     and NFxI.Produto_Id      = SbDev.Produto_Id" & vbCrLf & _
                 "    LEFT JOIN (" & vbCrLf & _
                 "                SELECT FxNF.Empresa_Id," & vbCrLf & _
                 "                       FxNF.EndEmpresa_Id," & vbCrLf & _
                 "                       FxNF.Cliente_Id," & vbCrLf & _
                 "                       FxNF.EndCliente_Id," & vbCrLf & _
                 "                       FxNF.EntradaSaida_Id," & vbCrLf & _
                 "                       FxNF.Serie_Id," & vbCrLf & _
                 "                       FxNF.Nota_Id," & vbCrLf & _
                 "                       FxNF.Produto_Id," & vbCrLf & _
                 "                       sum(FxNF.Quantidade) as Quantidade," & vbCrLf & _
                 "                       sum(FxNF.Valor) as Valor--," & vbCrLf & _
                 "                       --sum(FxNF.ValorFixacao) as ValorFixacao" & vbCrLf & _
                 "  			       FROM FixacaoXNotaFiscal FxNF" & vbCrLf & _
                 "                 Inner Join NotasFiscais nf" & vbCrLf & _
                 "                    On nf.Empresa_id      = FxNF.Empresa_Id" & vbCrLf & _
                 "                   and nf.EndEmpresa_Id   = FxNF.EndEmpresa_Id" & vbCrLf & _
                 "                   and nf.Cliente_Id      = FxNF.Cliente_Id" & vbCrLf & _
                 "                   and nf.EndCliente_Id   = FxNF.EndCliente_Id" & vbCrLf & _
                 "                   and nf.EntradaSaida_Id = FxNF.EntradaSaida_Id" & vbCrLf & _
                 "                   and nf.Serie_Id        = FxNF.Serie_Id" & vbCrLf & _
                 "                   and nf.Nota_Id         = FxNF.Nota_Id" & vbCrLf & _
                 "                 Where nf.situacao        in (1,4,7) " & vbCrLf & _
                 "                   and nf.TipoDeDocumento = 1 " & vbCrLf & _
                 "                   and nf.Empresa_Id     ='" & Fixacao.ItemPedido.Pedido.CodigoEmpresa & "'" & vbCrLf & _
                 " 	                 and nf.EndEmpresa_Id  = " & Fixacao.ItemPedido.Pedido.EnderecoEmpresa & vbCrLf & _
                 "	                 and nf.Cliente_Id     ='" & Fixacao.ItemPedido.Pedido.CodigoCliente & "'" & vbCrLf & _
                 "	                 and nf.EndCliente_Id  = " & Fixacao.ItemPedido.Pedido.EnderecoCliente & vbCrLf & _
                 "	                 and nf.Pedido         = " & Fixacao.ItemPedido.Pedido.Codigo & vbCrLf & _
                 "                 Group by FxNF.Empresa_Id, FxNF.EndEmpresa_Id, FxNF.Cliente_Id, FxNF.EndCliente_Id, FxNF.EntradaSaida_Id," & vbCrLf & _
                 "                          FxNF.Serie_Id, FxNF.Nota_Id, FxNF.Produto_Id" & vbCrLf & _
                 "               ) SbFix" & vbCrLf & _
                 "      On NFxI.Empresa_id      = SbFix.Empresa_Id" & vbCrLf & _
                 "     and NFxI.EndEmpresa_Id   = SbFix.EndEmpresa_Id" & vbCrLf & _
                 "     and NFxI.Cliente_Id      = SbFix.Cliente_Id" & vbCrLf & _
                 "     and NFxI.EndCliente_Id   = SbFix.EndCliente_Id" & vbCrLf & _
                 "     and NFxI.EntradaSaida_Id = SbFix.EntradaSaida_Id" & vbCrLf & _
                 "     and NFxI.Serie_Id        = SbFix.Serie_Id" & vbCrLf & _
                 "     and NFxI.Nota_Id         = SbFix.Nota_Id" & vbCrLf & _
                 "     and NFxI.Produto_Id      = SbFix.Produto_Id" & vbCrLf & _
                 "   WHERE NF.Situacao        in (1,4,7) " & vbCrLf & _
                 "     AND NF.TipoDeDocumento = 1 " & vbCrLf & _
                 "     AND SO.Devolucao        = 'N'" & vbCrLf & _
                 "     AND SO.Classe           = '" & eClassesOperacoes.AFIXAR.ToString & "'" & vbCrLf & _
                 "     AND NFxI.Empresa_Id     ='" & Fixacao.ItemPedido.Pedido.CodigoEmpresa & "'" & vbCrLf & _
                 " 	   AND NFxI.EndEmpresa_Id  = " & Fixacao.ItemPedido.Pedido.EnderecoEmpresa & vbCrLf & _
                 "	   AND NFxI.Cliente_Id     ='" & Fixacao.ItemPedido.Pedido.CodigoCliente & "'" & vbCrLf & _
                 "	   AND NFxI.EndCliente_Id  = " & Fixacao.ItemPedido.Pedido.EnderecoCliente & vbCrLf & _
                 "	   AND NFxI.Pedido         = " & Fixacao.ItemPedido.Pedido.Codigo & vbCrLf & _
                 "	   AND NFxI.Produto_Id     ='" & Fixacao.ItemPedido.CodigoProduto & "'" & vbCrLf & _
                 "	   AND (" & vbCrLf & _
                 "          NFxI.QuantidadeFiscal - isnull(sbdev.Quantidade,0) - isnull(sbfix.Quantidade,0) > 0 " & vbCrLf & _
                 "          or" & vbCrLf & _
                 "          NFxI.Valor - isnull(sbdev.Valor,0) - isnull(sbfix.Valor,0) > 0" & vbCrLf & _
                 "         )  " & vbCrLf

        Dim objBanco As New AcessaBanco
        Dim dsFixXNF As DataSet = objBanco.ConsultaDataSet(strSQL, "FixacaoXNotaFiscal")

        For Each row As DataRow In dsFixXNF.Tables(0).Rows

            Dim FixXNF As New PedidoFixacaoXNotaFiscal(Fixacao)

            Dim NF As New NotaFiscal()
            NF.CodigoEmpresa = row("Empresa_Id")
            NF.EnderecoEmpresa = row("EndEmpresa_Id")
            NF.CodigoCliente = row("Cliente_Id")
            NF.EnderecoCliente = row("EndCliente_Id")
            NF.EntradaSaida = IIf(row("EntradaSaida_Id") = "S", 1, 0)
            NF.Serie = row("Serie_Id")
            NF.Codigo = row("Nota_Id")

            Dim NFxI As New NotaFiscalXItem(NF)
            FixXNF.NotaFiscalXItem = NFxI
            FixXNF.NotaFiscalXItem.CFOP = row("CFOP_Id")
            FixXNF.NotaFiscalXItem.Sequencia = row("Sequencia_Id")
            FixXNF.QtdeNotaFiscal = row("QtdeNotaFiscal")
            FixXNF.UnitarioNotaFiscal = row("Unitario")
            FixXNF.ValorNotafiscal = row("ValorNotaFiscal")
            FixXNF.QtdeDevolucao = row("QtdeDevolucao")
            FixXNF.ValorDevolucao = row("ValorDevolucao")
            FixXNF.QtdeFixacao = row("QtdeFixacao")
            FixXNF.ValorFixacao = row("ValorFixacao")
            FixXNF.QtdeAFixar = row("QdteSaldo")
            FixXNF.ValorAFixar = row("ValorSaldo")
            FixXNF.PercentualFacs = row("PercentualFacs")
            FixXNF.PercentualFethab = row("PercentualFethab")
            FixXNF.Liberada = False
            Me.Add(FixXNF)
        Next
    End Sub

    Public Sub CarregarNotasUsadasNaFixacao()
        Me.Clear()
        Dim strSQL As String = ""
        strSQL = "    SELECT FxNF.Empresa_Id," & vbCrLf & _
                 "           FxNF.EndEmpresa_Id," & vbCrLf & _
                 "           FxNF.Cliente_Id," & vbCrLf & _
                 "           FxNF.EndCliente_Id," & vbCrLf & _
                 "           FxNF.EntradaSaida_Id," & vbCrLf & _
                 "           FxNF.Serie_Id," & vbCrLf & _
                 "           FxNF.Nota_Id," & vbCrLf & _
                 "           FxNF.Fixacao_id," & vbCrLf & _
                 "           FxNF.Produto_Id," & vbCrLf & _
                 "           FxNF.CFOP_Id," & vbCrLf & _
                 "           FxNF.Sequencia_Id," & vbCrLf & _
                 "           FxNF.Quantidade," & vbCrLf & _
                 "           FxNF.Valor" & vbCrLf & _
                 "      FROM FixacaoXNotaFiscal FxNF" & vbCrLf & _
                 "     Inner Join NotasFiscais nf" & vbCrLf & _
                 "        On nf.Empresa_id      = FxNF.Empresa_Id" & vbCrLf & _
                 "       and nf.EndEmpresa_Id   = FxNF.EndEmpresa_Id" & vbCrLf & _
                 "       and nf.Cliente_Id      = FxNF.Cliente_Id" & vbCrLf & _
                 "       and nf.EndCliente_Id   = FxNF.EndCliente_Id" & vbCrLf & _
                 "       and nf.EntradaSaida_Id = FxNF.EntradaSaida_Id" & vbCrLf & _
                 "       and nf.Serie_Id        = FxNF.Serie_Id" & vbCrLf & _
                 "       and nf.Nota_Id         = FxNF.Nota_Id" & vbCrLf & _
                 "     Where FxNF.Empresa_Id     ='" & Fixacao.ItemPedido.Pedido.CodigoEmpresa & "'" & vbCrLf & _
                 " 	     AND FxNF.EndEmpresa_Id  = " & Fixacao.ItemPedido.Pedido.EnderecoEmpresa & vbCrLf & _
                 "	     AND FxNF.Cliente_Id     ='" & Fixacao.ItemPedido.Pedido.CodigoCliente & "'" & vbCrLf & _
                 "	     AND FxNF.EndCliente_Id  = " & Fixacao.ItemPedido.Pedido.EnderecoCliente & vbCrLf & _
                 "	     AND FxNF.Pedido_id      = " & Fixacao.ItemPedido.Pedido.Codigo & vbCrLf & _
                 "	     AND FxNF.Produto_Id     ='" & Fixacao.ItemPedido.CodigoProduto & "'" & vbCrLf & _
                 "       AND FxNF.Fixacao_id     = " & Fixacao.Codigo

        Dim objBanco As New AcessaBanco
        Dim dsFixXNF As DataSet = objBanco.ConsultaDataSet(strSQL, "FixacaoXNotaFiscal")

        For Each row As DataRow In dsFixXNF.Tables(0).Rows
            Dim FixXNF As New PedidoFixacaoXNotaFiscal(Fixacao)

            Dim NF As New NotaFiscal()
            NF.CodigoEmpresa = row("Empresa_Id")
            NF.EnderecoEmpresa = row("EndEmpresa_Id")
            NF.CodigoCliente = row("Cliente_Id")
            NF.EnderecoCliente = row("EndCliente_Id")
            NF.EntradaSaida = IIf(row("EntradaSaida_Id") = "S", 1, 0)
            NF.Serie = row("Serie_Id")
            NF.Codigo = row("Nota_Id")

            Dim NFxI As New NotaFiscalXItem(NF)
            FixXNF.NotaFiscalXItem = NFxI
            FixXNF.NotaFiscalXItem.CodigoProduto = row("Produto_Id")
            FixXNF.NotaFiscalXItem.CFOP = row("CFOP_Id")
            FixXNF.NotaFiscalXItem.Sequencia = row("Sequencia_Id")
            FixXNF.QtdeNotaFiscal = 0
            FixXNF.UnitarioNotaFiscal = 0
            FixXNF.ValorNotafiscal = 0
            FixXNF.QtdeDevolucao = 0
            FixXNF.ValorDevolucao = 0
            FixXNF.QtdeFixacao = row("Quantidade")
            FixXNF.ValorFixacao = row("Valor")
            FixXNF.QtdeAFixar = 0
            FixXNF.ValorAFixar = 0
            FixXNF.Liberada = False
            Me.Add(FixXNF)
        Next
    End Sub
#End Region

End Class

<Serializable()> _
Public Class PedidoFixacaoXNotaFiscal

#Region "Construtor"
    Public Sub New(pFixacao As Fixacao)
        _Fixacao = pFixacao
    End Sub
#End Region

#Region "Fields"
    Private _IUD As String = ""
    Private _Fixacao As Fixacao

    'Private _CodigoEmpresa As String = ""
    'Private _EnderecoEmpresa As Integer
    'Private _CodigoCliente As String = ""
    'Private _EnderecoCliente As Integer
    'Private _EntradaSaida As String = ""
    'Private _Serie As String = ""
    'Private _NumeroNota As Integer
    Private _NotaFiscalXItem As [Lib].Negocio.NotaFiscalXItem

    Private _QtdeNotaFiscal As Decimal
    Private _UnitarioNotaFiscal As Decimal
    Private _ValorNotafiscal As Decimal
    Private _ValorFixacao As Decimal

    '#ValorFixacaoReal
    'Private _ValorFixacaoReal As Decimal

    Private _QtdeDevolucao As Decimal
    Private _ValorDevolucao As Decimal
    Private _QtdeFixacao As Decimal
    Private _QtdeAFixar As Decimal
    Private _ValorAFixar As Decimal
    Private _PercentualFacs As Decimal
    Private _PercentualFethab As Decimal
    Private _PercentualFundersul As Decimal
    Private _PercentualFudems As Decimal
    Private _Liberada As Boolean
    Private _QtdeAFixarLiberado As Decimal
    Private _ValorAFixarLiberado As Decimal
#End Region

#Region "Property"
    Public Property IUD() As String
        Get
            Return _IUD
        End Get
        Set(ByVal value As String)
            _IUD = value
        End Set
    End Property

    'Fixacao
    Public ReadOnly Property Fixacao As Fixacao
        Get
            Return _Fixacao
        End Get
    End Property

    'NotaxItem

    Public Property NotaFiscalXItem() As [Lib].[Negocio].NotaFiscalXItem
        Get
            Return _NotaFiscalXItem
        End Get
        Set(ByVal value As [Lib].[Negocio].NotaFiscalXItem)
            _NotaFiscalXItem = value
        End Set
    End Property

    Public Property QtdeNotaFiscal() As Decimal
        Get
            Return _QtdeNotaFiscal
        End Get
        Set(ByVal value As Decimal)
            _QtdeNotaFiscal = value
        End Set
    End Property

    Public Property UnitarioNotaFiscal() As Decimal
        Get
            Return _UnitarioNotaFiscal
        End Get
        Set(ByVal value As Decimal)
            _UnitarioNotaFiscal = value
        End Set
    End Property

    Public Property ValorNotafiscal() As Decimal
        Get
            Return _ValorNotafiscal
        End Get
        Set(ByVal value As Decimal)
            _ValorNotafiscal = value
        End Set
    End Property

    Public Property QtdeDevolucao() As Decimal
        Get
            Return _QtdeDevolucao
        End Get
        Set(ByVal value As Decimal)
            _QtdeDevolucao = value
        End Set
    End Property

    Public Property ValorDevolucao() As Decimal
        Get
            Return _ValorDevolucao
        End Get
        Set(ByVal value As Decimal)
            _ValorDevolucao = value
        End Set
    End Property

    Public Property QtdeFixacao() As Decimal
        Get
            Return _QtdeFixacao
        End Get
        Set(ByVal value As Decimal)
            _QtdeFixacao = value
        End Set
    End Property

    Public Property ValorFixacao() As Decimal
        Get
            Return _ValorFixacao
        End Get
        Set(ByVal value As Decimal)
            _ValorFixacao = value
        End Set
    End Property

    '#ValorFixacaoReal
    'Public Property ValorFixacaoReal() As Decimal
    '    Get
    '        Return _ValorFixacaoReal
    '    End Get
    '    Set(ByVal value As Decimal)
    '        _ValorFixacaoReal = value
    '    End Set
    'End Property

    Public Property QtdeAFixar() As Decimal
        Get
            Return _QtdeAFixar
        End Get
        Set(ByVal value As Decimal)
            _QtdeAFixar = value
        End Set
    End Property

    Public Property ValorAFixar() As Decimal
        Get
            Return _ValorAFixar
        End Get
        Set(ByVal value As Decimal)
            _ValorAFixar = value
        End Set
    End Property

    Public Property PercentualFacs() As Decimal
        Get
            Return _PercentualFacs
        End Get
        Set(ByVal value As Decimal)
            _PercentualFacs = value
        End Set
    End Property

    Public Property PercentualFethab() As Decimal
        Get
            Return _PercentualFethab
        End Get
        Set(ByVal value As Decimal)
            _PercentualFethab = value
        End Set
    End Property

    Public Property PercentualFundersul() As Decimal
        Get
            Return _PercentualFundersul
        End Get
        Set(ByVal value As Decimal)
            _PercentualFundersul = value
        End Set
    End Property

    Public Property PercentualFudems() As Decimal
        Get
            Return _PercentualFudems
        End Get
        Set(ByVal value As Decimal)
            _PercentualFudems = value
        End Set
    End Property


    Public Property Liberada() As Boolean
        Get
            Return _Liberada
        End Get
        Set(ByVal value As Boolean)
            _Liberada = value
        End Set
    End Property

    Public Property QtdeAFixarLiberado() As Decimal
        Get
            Return _QtdeAFixarLiberado
        End Get
        Set(ByVal value As Decimal)
            _QtdeAFixarLiberado = value
        End Set
    End Property

    Public Property ValorAFixarLiberado() As Decimal
        Get
            Return _ValorAFixarLiberado
        End Get
        Set(ByVal value As Decimal)
            _ValorAFixarLiberado = value
        End Set
    End Property
#End Region

#Region "Methods"
    Public Function Salvar() As Boolean
        If IUD = Nothing Then Return True
        Dim Banco As New AcessaBanco
        Dim Sqls As New ArrayList

        Sqls.Clear()
        SalvarSql(Sqls)

        If Banco.GravaBanco(Sqls) Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Sub SalvarSql(ByRef Sqls As ArrayList)
        Dim sql As String = ""

        Dim NFxI = Me.NotaFiscalXItem

        Select Case Me.IUD
            Case "I"
                sql = "Insert Into FixacaoXNotaFiscal " & vbCrLf & _
                      "            (Empresa_Id, EndEmpresa_Id, Pedido_Id, Produto_Id, Fixacao_Id, " & vbCrLf & _
                      "             Cliente_Id, EndCliente_Id, EntradaSaida_Id, Serie_Id, Nota_Id, CFOP_Id, Sequencia_Id, " & vbCrLf & _
                      "             Quantidade, Valor) " & vbCrLf & _
                      "     Values ('" & NFxI.NotaFiscal.CodigoEmpresa & "'," & NFxI.NotaFiscal.EnderecoEmpresa & "," & Fixacao.ItemPedido.Pedido.Codigo & ",'" & Fixacao.ItemPedido.CodigoProduto & "'," & Fixacao.Codigo & "," & vbCrLf & _
                      "             '" & NFxI.NotaFiscal.CodigoCliente & "'," & NFxI.NotaFiscal.EnderecoCliente & ",'" & IIf(NFxI.NotaFiscal.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "','" & NFxI.NotaFiscal.Serie & "'," & vbCrLf & _
                      "              " & NFxI.NotaFiscal.Codigo & "," & NFxI.CFOP & "," & NFxI.Sequencia & "," & vbCrLf & _
                      "              " & Str(QtdeAFixarLiberado) & "," & Str(ValorAFixarLiberado) & ")" & vbCrLf
                Sqls.Add(sql)

            Case "D"
                sql = "Delete From FixacaoXNotaFiscal " & vbCrLf & _
                      "Where Empresa_Id      ='" & NFxI.NotaFiscal.CodigoEmpresa & "'" & vbCrLf & _
                      "  and EndEmpresa_Id   = " & NFxI.NotaFiscal.EnderecoEmpresa & vbCrLf & _
                      "  and Pedido_Id       = " & Fixacao.ItemPedido.Pedido.Codigo & vbCrLf & _
                      "  and Produto_Id      ='" & Fixacao.ItemPedido.CodigoProduto & "'" & vbCrLf & _
                      "  and Fixacao_Id      = " & Fixacao.Codigo & vbCrLf & _
                      "  and Cliente_Id      ='" & NFxI.NotaFiscal.CodigoCliente & "'" & vbCrLf & _
                      "  and EndCliente_Id   = " & NFxI.NotaFiscal.EnderecoCliente & vbCrLf & _
                      "  and EntradaSaida_Id ='" & IIf(NFxI.NotaFiscal.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "'" & vbCrLf & _
                      "  and Serie_Id        ='" & NFxI.NotaFiscal.Serie & "'" & vbCrLf & _
                      "  and Nota_Id         = " & NFxI.NotaFiscal.Codigo & vbCrLf

                Sqls.Add(sql)
        End Select
    End Sub
#End Region

End Class
