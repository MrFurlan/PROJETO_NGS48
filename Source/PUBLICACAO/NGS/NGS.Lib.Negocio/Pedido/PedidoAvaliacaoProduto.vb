Imports Microsoft.VisualBasic
Imports System.Data
Imports NGS.Lib.Uteis

<Serializable()> _
Public Class PedidoAvaliacaoProduto

#Region "Construtor"
    Public Sub New()

    End Sub
#End Region

#Region "Fields"
    Private _ValorOficial As Double
    Private _ValorMoeda As Double
    Private _Banco As New AcessaBanco
#End Region

#Region "Property"
    Public Property ValorOficial() As Double
        Get
            Return _ValorOficial
        End Get
        Set(ByVal value As Double)
            _ValorOficial = value
        End Set
    End Property

    Public Property ValorMoeda() As Double
        Get
            Return _ValorMoeda
        End Get
        Set(ByVal value As Double)
            _ValorMoeda = value
        End Set
    End Property
#End Region

#Region "Methods"
    Public Sub AvaliarProduto(ByVal CodigoEmpresa As String, ByVal EndEmpresa As Integer, ByVal CodigoSafra As String, ByVal codigoProduto As String, ByVal DataAvaliacao As Date)
        Dim sql As String
        sql = "select top 1 ordem, data" & vbCrLf & _
              "  from (" & vbCrLf & _
              "		Select 1 ordem, max(Data_Id) as data" & vbCrLf & _
              "		  from ProdutosXPrecos" & vbCrLf & _
              "		 where Data_Id   <= '" & DataAvaliacao.ToString("yyyy-MM-dd") & "'" & vbCrLf & _
              "        and Produto_id = '" & codigoProduto & "'" & vbCrLf
        '"     union" & vbCrLf & _
        '"		select 2 ordem, max(pedidos.DataPedido) as Data" & vbCrLf & _
        '"		  from pedidos" & vbCrLf & _
        '"		  INNER JOIN PedidosXItens " & vbCrLf & _
        '"			ON Pedidos.Empresa_Id    = PedidosXItens.Empresa_Id " & vbCrLf & _
        '"		   AND Pedidos.EndEmpresa_Id = PedidosXItens.EndEmpresa_Id " & vbCrLf & _
        '"		   AND Pedidos.Pedido_Id     = PedidosXItens.Pedido_Id " & vbCrLf & _
        '"		   AND Pedidos.Sequencia_Id  = PedidosXItens.Sequencia_Id " & vbCrLf & _
        '"      Inner Join SubOperacoes SO" & vbCrLf & _
        '"         ON SO.Operacao_id     = Pedidos.Operacao" & vbCrLf & _
        '"        and SO.SubOperacoes_id = Pedidos.SubOperacao" & vbCrLf & _
        '"		 Inner Join Produtos Pd" & vbCrLf & _
        '"			ON PedidosXItens.Produto_id = Pd.Produto_id" & vbCrLf & _
        '"		 where Pedidos.Empresa_id    ='" & CodigoEmpresa & "'" & vbCrLf & _
        '"         and Pedidos.EndEmpresa_id = " & EndEmpresa & vbCrLf & _
        '"		   and Pd.Produto_Id         ='" & codigoProduto & "'" & vbCrLf & _
        '"		   and Pedidos.DataPedido   <='" & DataAvaliacao.ToString("yyyy-MM-dd") & "'" & vbCrLf & _
        '"		   and So.Entradasaida       = 'E'" & vbCrLf & _
        '"		   and So.Devolucao          = 'N'" & vbCrLf & _
        '"		   and So.Classe            in ('"&  eClassesOperacoes.COMPRAS.ToString &"','"&  eClassesOperacoes.GLOBAL.ToString &"','"&  eClassesOperacoes.IMPORTACOES.ToString &"')" & vbCrLf & _
        '"		   and Pedidos.Situacao      = 1" & vbCrLf & _
        '"		union" & vbCrLf & _
        '"		select 3 ordem, max(pedidos.DataPedido) as Data" & vbCrLf & _
        '"		  from pedidos" & vbCrLf & _
        '"		  INNER JOIN PedidosXItens " & vbCrLf & _
        '"			ON Pedidos.Empresa_Id    = PedidosXItens.Empresa_Id " & vbCrLf & _
        '"		   AND Pedidos.EndEmpresa_Id = PedidosXItens.EndEmpresa_Id " & vbCrLf & _
        '"		   AND Pedidos.Pedido_Id     = PedidosXItens.Pedido_Id " & vbCrLf & _
        '"		   AND Pedidos.Sequencia_Id  = PedidosXItens.Sequencia_Id " & vbCrLf & _
        '"      Inner Join SubOperacoes SO" & vbCrLf & _
        '"         ON SO.Operacao_id     = Pedidos.Operacao" & vbCrLf & _
        '"        and SO.SubOperacoes_id = Pedidos.SubOperacao" & vbCrLf & _
        '"  	 Inner Join Produtos Pd" & vbCrLf & _
        '"		    ON PedidosXItens.Produto_id = Pd.Produto_id" & vbCrLf & _
        '"		 where Pedidos.Empresa_id    = '" & CodigoEmpresa & "'" & vbCrLf & _
        '"         and Pedidos.EndEmpresa_id = " & EndEmpresa & vbCrLf & _
        '"		   and Pd.Produto_Id         = '" & codigoProduto & "'" & vbCrLf & _
        '"		   and Pedidos.DataPedido   <= '" & DataAvaliacao.ToString("yyyy-MM-dd") & "' " & vbCrLf & _
        '"		   and So.Entradasaida       = 'E' " & vbCrLf & _
        '"		   and So.Devolucao          = 'N' " & vbCrLf & _
        '"		   and So.Classe            in ('"&  eClassesOperacoes.COMPRAS.ToString &"','"&  eClassesOperacoes.GLOBAL.ToString &"','"&  eClassesOperacoes.IMPORTACOES.ToString &"')" & vbCrLf & _
        '"		   and Pedidos.Situacao      = 1" & vbCrLf & _
        sql &= "		 union" & vbCrLf & _
              "		Select 4 ordem, max(Data_Id) as data" & vbCrLf & _
              "		  from ProdutosXPrecos" & vbCrLf & _
              "		 where Data_Id    <= '" & DataAvaliacao.ToString("yyyy-MM-dd") & "'" & vbCrLf & _
              "        and Produto_id  = '" & codigoProduto & "'" & vbCrLf & _
              "     ) sb" & vbCrLf & _
              " where data is not null" & vbCrLf & _
              " order by data, ordem" & vbCrLf

        Dim ds As DataSet = _Banco.ConsultaDataSet(sql, "Pedidos")
        If ds.Tables(0).Rows.Count = 0 Then
            _ValorMoeda = 0
            _ValorOficial = 0
            Exit Sub
        End If
        Dim row As DataRow = ds.Tables(0).Rows(0)

        Select Case row("Ordem")
            Case 1
                sql = "Select Moeda, Valor" & vbCrLf & _
                      "  from ProdutosXPrecos" & vbCrLf & _
                      " where Data_Id    = '" & CDate(row("Data")).ToString("yyyy-MM-dd") & "'" & vbCrLf & _
                      "   and Produto_id = '" & codigoProduto & "'" & vbCrLf
                ds = _Banco.ConsultaDataSet(sql, "Ordem1")

                Dim M As New Moeda(ds.Tables(0).Rows(0)("Moeda"))
                Dim c As New Cotacao(3, CDate(row("Data")).ToString("yyyy-MM-dd"))

                If M.Classificacao = eTiposMoeda.Oficial Then
                    _ValorOficial = ds.Tables(0).Rows(0)("Valor")
                    _ValorMoeda = _ValorOficial / c.Indice
                Else
                    _ValorMoeda = ds.Tables(0).Rows(0)("Valor")
                    _ValorOficial = _ValorMoeda * c.Indice
                End If
            Case 2
                sql = "Select pxi.UnitarioOficial, pxi.UnitarioMoeda" & vbCrLf & _
                      "  from pedidos" & vbCrLf & _
                      " INNER JOIN PedidoXItemXLancamento pxi " & vbCrLf & _
                      "    ON Pedidos.Empresa_Id    = pxi.Empresa_Id " & vbCrLf & _
                      "   AND Pedidos.EndEmpresa_Id = pxi.EndEmpresa_Id " & vbCrLf & _
                      "   AND Pedidos.Pedido_Id     = pxi.Pedido_Id " & vbCrLf & _
                      " Inner Join SubOperacoes SO" & vbCrLf & _
                      "    ON SO.Operacao_id     = Pedidos.Operacao" & vbCrLf & _
                      "   and SO.SubOperacoes_id = Pedidos.SubOperacao" & vbCrLf & _
                      " Inner Join Produtos Pd" & vbCrLf & _
                      "    ON pxi.Produto_id = Pd.Produto_id" & vbCrLf & _
                      " where Pedidos.safra         ='" & CodigoSafra & "'" & vbCrLf & _
                      "   and Pedidos.Empresa_id    ='" & CodigoEmpresa & "'" & vbCrLf & _
                      "   and Pedidos.EndEmpresa_Id = " & EndEmpresa & vbCrLf & _
                      "   and Pd.Produto_Id         ='" & codigoProduto & "'" & vbCrLf & _
                      "   and Pedidos.DataPedido    ='" & CDate(row("Data")).ToString("yyyy-MM-dd") & "' " & vbCrLf & _
                      "   and So.Entradasaida       ='E' " & vbCrLf & _
                      "   and So.Devolucao          ='N' " & vbCrLf & _
                      "   and So.Classe            in ('" & eClassesOperacoes.COMPRAS.ToString & "','" & eClassesOperacoes.GLOBAL.ToString & "','" & eClassesOperacoes.IMPORTACOES.ToString & "')" & vbCrLf & _
                      "   and Pedidos.Situacao      = 1" & vbCrLf
                ds = _Banco.ConsultaDataSet(sql, "Ordem2")
                _ValorMoeda = ds.Tables(0).Rows(0)("UnitarioMoeda")
                _ValorOficial = ds.Tables(0).Rows(0)("UnitarioOficial")
            Case 3
                sql = "Select pxi.UnitarioOficial, pxi.UnitarioMoeda" & vbCrLf & _
                      "  from pedidos" & vbCrLf & _
                      " INNER JOIN PedidoXItemXLancamento pxi " & vbCrLf & _
                      "    ON Pedidos.Empresa_Id    = pxi.Empresa_Id " & vbCrLf & _
                      "   AND Pedidos.EndEmpresa_Id = pxi.EndEmpresa_Id " & vbCrLf & _
                      "   AND Pedidos.Pedido_Id     = pxi.Pedido_Id " & vbCrLf & _
                      " Inner Join SubOperacoes SO" & vbCrLf & _
                      "    ON SO.Operacao_id     = Pedidos.Operacao" & vbCrLf & _
                      "   and SO.SubOperacoes_id = Pedidos.SubOperacao" & vbCrLf & _
                      " Inner Join Produtos Pd" & vbCrLf & _
                      "    ON pxi.Produto_id = Pd.Produto_id" & vbCrLf & _
                      " where Pedidos.Empresa_id    ='" & CodigoEmpresa & "'" & vbCrLf & _
                      "   and Pedidos.EndEmpresa_Id = " & EndEmpresa & vbCrLf & _
                      "   and Pd.Produto_Id         ='" & codigoProduto & "'" & vbCrLf & _
                      "   and Pedidos.DataPedido    ='" & CDate(row("Data")).ToString("yyyy-MM-dd") & "' " & vbCrLf & _
                      "   and So.Entradasaida       ='E' " & vbCrLf & _
                      "   and So.Devolucao          ='N' " & vbCrLf & _
                      "   and So.Classe            in ('" & eClassesOperacoes.COMPRAS.ToString & "','" & eClassesOperacoes.GLOBAL.ToString & "','" & eClassesOperacoes.IMPORTACOES.ToString & "')" & vbCrLf & _
                      "   and Pedidos.Situacao      = 1" & vbCrLf
                ds = _Banco.ConsultaDataSet(sql, "Ordem3")
                _ValorMoeda = ds.Tables(0).Rows(0)("UnitarioMoeda")
                _ValorOficial = ds.Tables(0).Rows(0)("UnitarioOficial")
            Case 4
                sql = "Select Moeda, Valor" & vbCrLf & _
                      "  from ProdutosXPrecos" & vbCrLf & _
                      " where Data_Id    = '" & CDate(row("Data")).ToString("yyyy-MM-dd") & "'" & vbCrLf & _
                      "   and Produto_id = '" & codigoProduto & "'" & vbCrLf
                ds = _Banco.ConsultaDataSet(sql, "Ordem1")

                Dim M As New Moeda(ds.Tables(0).Rows(0)("Moeda"))
                Dim c As New Cotacao(3, CDate(row("Data")).ToString("yyyy-MM-dd"))

                If M.Classificacao = eTiposMoeda.Oficial Then
                    _ValorOficial = ds.Tables(0).Rows(0)("Valor")
                    _ValorMoeda = _ValorOficial / c.Indice
                Else
                    _ValorMoeda = ds.Tables(0).Rows(0)("Valor")
                    _ValorOficial = _ValorMoeda * c.Indice
                End If
        End Select
    End Sub

    Public Function AvaliarPedidos(ByVal CodigoEmpresa As String, ByVal EndEmpresa As Integer, ByVal CodigoSafra As String) As Boolean
        Dim sql As String
        sql = " SELECT distinct PxI.Produto_Id, PxI.DataEntrega" & vbCrLf & _
              "   FROM Pedidos P" & vbCrLf & _
              "  Inner Join Suboperacoes SO" & vbCrLf & _
              "     on SO.Operacao_Id     = P.Operacao" & vbCrLf & _
              "    and SO.Suboperacoes_Id = P.Suboperacao" & vbCrLf & _
              "  INNER JOIN PedidoXItemXLancamento PxI " & vbCrLf & _
              "     ON P.Empresa_Id     = PxI.Empresa_Id " & vbCrLf & _
              "    AND P.EndEmpresa_Id  = PxI.EndEmpresa_Id " & vbCrLf & _
              "    AND P.Pedido_Id      = PxI.Pedido_Id " & vbCrLf & _
              "  WHERE P.Empresa_Id    = '" & CodigoEmpresa & "'" & vbCrLf & _
              "    AND P.EndEmpresa_Id =  " & EndEmpresa & vbCrLf & _
              "    AND P.Safra         = '" & CodigoSafra & "'" & vbCrLf & _
              "    AND So.Entradasaida = 'S'" & vbCrLf & _
              "    AND So.Devolucao    = 'N'" & vbCrLf & _
              "    AND P.Situacao      =  1 " & vbCrLf & _
              "    AND So.Classe in ('" & eClassesOperacoes.VENDAS.ToString & "','" & eClassesOperacoes.REMESSAS.ToString & "','" & eClassesOperacoes.EXPORTACOES.ToString & "','" & eClassesOperacoes.VENDASAORDEM.ToString & "','" & eClassesOperacoes.GLOBAL.ToString & "')"

        Dim ds As DataSet = _Banco.ConsultaDataSet(sql, "Pedidos")
        Dim sqls As New ArrayList

        For Each row As DataRow In ds.Tables(0).Rows
            AvaliarProduto(CodigoEmpresa, EndEmpresa, CodigoSafra, row("Produto_Id"), row("DataEntrega"))

            sql = "Update pxi set " & vbCrLf & _
                  "     UnitarioOficialCompra = " & Str(_ValorOficial) & vbCrLf & _
                  "    ,UnitarioMoedaCompra   = " & Str(_ValorMoeda) & vbCrLf & _
                  "  From Pedidos " & vbCrLf & _
                  " Inner Join PedidoXItemXLancamento pxi" & vbCrLf & _
                  "    ON Pedidos.Empresa_Id     = pxi.Empresa_Id " & vbCrLf & _
                  "   AND Pedidos.EndEmpresa_Id  = pxi.EndEmpresa_Id " & vbCrLf & _
                  "   AND Pedidos.Pedido_Id      = pxi.Pedido_Id " & vbCrLf & _
                  " Where pedidos.Empresa_Id        ='" & CodigoEmpresa & "'" & vbCrLf & _
                  "   and pedidos.EndEmpresa_Id     = " & EndEmpresa & vbCrLf & _
                  "   and Pedidos.safra             ='" & CodigoSafra & "'" & _
                  "   And pxi.Produto_Id  ='" & row("Produto_Id") & "'" & vbCrLf & _
                  "   And pxi.DataEntrega ='" & CDate(row("DataEntrega")).ToString("yyyy-MM-dd") & "'" & vbCrLf & _
                  "   AND (UnitarioOficialCompra <> " & Str(_ValorOficial) & " or UnitarioMoedaCompra <> " & Str(_ValorMoeda) & ") "
            sqls.Add(sql)
        Next

        Return _Banco.GravaBanco(sqls)
    End Function
#End Region

End Class