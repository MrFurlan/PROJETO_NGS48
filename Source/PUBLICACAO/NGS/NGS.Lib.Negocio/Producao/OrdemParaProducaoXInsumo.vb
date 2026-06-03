Public Class ListOrdemParaProducaoXInsumo
    Inherits List(Of OrdemParaProducaoXInsumo)
    Implements IBaseEntity

#Region "Contrutor"
    Public Sub New()
    End Sub

    Public Sub New(ByVal objOrdemDeProducao As OrdemParaProducao)

        Me.OrdemParaProducao = objOrdemDeProducao

        Dim db As New AcessaBanco()

        Try
            Dim strSQL As String

            strSQL = "Select Empresa_Id, EndEmpresa_Id, Ordem_Id, ProdutoProducao_Id, Produto_Id, Quantidade" & vbCrLf &
                     "From OrdemDeProducaoXInsumo " & vbCrLf &
                     "Where Empresa_id    = '" & Me.OrdemParaProducao.CodigoEmpresa & "'" & vbCrLf &
                     "  and EndEmpresa_Id = " & Me.OrdemParaProducao.EnderecoEmpresa & vbCrLf

            If Me.OrdemParaProducao.Codigo > 0 Then strSQL &= "  and Ordem_Id      = " & Me.OrdemParaProducao.Codigo & vbCrLf

            Dim ds As DataSet = db.ConsultaDataSet(strSQL, "OrdemParaProducaoXInsumo")

            For Each row As DataRow In ds.Tables(0).Rows
                Dim oPC As New OrdemParaProducaoXInsumo(Me.OrdemParaProducao)
                oPC.CodigoProdutoProducao = row("ProdutoProducao_Id").ToString()
                oPC.CodigoProduto = row("Produto_Id").ToString()
                oPC.Produto = New Produto(row("Produto_Id").ToString())
                oPC.Quantidade = row("Quantidade")

                Me.Add(oPC)
            Next

        Catch ex As Exception
            Throw ex
        Finally
            db = Nothing
        End Try
    End Sub

    Public Sub SalvarSql(ByRef Sqls As ArrayList)
        For Each item As OrdemParaProducaoXInsumo In Me
            item.SalvarSql(Sqls)
        Next
    End Sub
#End Region

#Region "Fields"
    Private _OrdemParaProducao As OrdemParaProducao
#End Region

#Region "Property"
    Public Property OrdemParaProducao() As OrdemParaProducao
        Get
            Return _OrdemParaProducao
        End Get
        Set(ByVal value As OrdemParaProducao)
            _OrdemParaProducao = value
        End Set
    End Property

#End Region

End Class

<Serializable()> _
Public Class OrdemParaProducaoXInsumo
    Implements IBaseEntity

#Region "Fields"
    Private _IUD As String = ""

    Private _OrdemParaProducao As OrdemParaProducao

    Private _CodigoProduto As String
    Private _Produto As Produto
    Private _CodigoProdutoProducao As String

    Private _Quantidade As Decimal

    Private _SubOperacaoInsumo As SubOperacao

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

    Public Property OrdemParaProducao() As OrdemParaProducao
        Get
            Return _OrdemParaProducao
        End Get
        Set(ByVal value As OrdemParaProducao)
            _OrdemParaProducao = value
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

    Public Property Produto() As Produto
        Get
            If _Produto Is Nothing And Not Me.CodigoProduto Is Nothing Then _Produto = New Produto(Me.CodigoProduto)
            Return _Produto
        End Get
        Set(ByVal value As Produto)
            _Produto = value
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

    Public Property SubOperacaoInsumo() As SubOperacao
        Get
            If _SubOperacaoInsumo Is Nothing OrElse _SubOperacaoInsumo.Codigo = 0 Then _SubOperacaoInsumo = New SubOperacao(40, 52)

            Return _SubOperacaoInsumo
        End Get
        Set(value As SubOperacao)
            _SubOperacaoInsumo = value
        End Set
    End Property

    Public Property CodigoProdutoProducao As String
        Get
            Return _CodigoProdutoProducao
        End Get
        Set(value As String)
            _CodigoProdutoProducao = value
        End Set
    End Property

#End Region

#Region "Construtor"
    Public Sub New()
    End Sub

    Public Sub New(ByVal OrdemParaProducao As OrdemParaProducao)
        Me.OrdemParaProducao = OrdemParaProducao
    End Sub
#End Region

#Region "Methods"

    Public Function buscarEstoqueProduto(ByVal Empresa As String, ByVal EndEmpresa As Integer, ByVal codigoProduto As String, ByVal DataAnterior As String, ByVal DataAtual As String) As Decimal
        Dim db As New AcessaBanco()
        Dim ds As New DataSet
        Dim quantidade As Decimal = 0

        Try
            Dim strSQL As String = "SELECT 0 AS Entradas, 0 AS Saidas, SUM(SaldoAnterior) AS Saldo" & vbCrLf & _
                                    "INTO #Estoque" & vbCrLf & _
                                    "FROM (SELECT COALESCE(SUM(E.Entradas - E.Saidas), 0) AS SaldoAnterior " & vbCrLf & _
                                    "FROM Producao E " & vbCrLf & _
                                    "INNER JOIN SubOperacoes SO " & vbCrLf & _
                                    "ON SO.Operacao_Id = E.Operacao_Id " & vbCrLf & _
                                    "AND SO.SubOperacoes_Id = E.SubOperacao_Id " & vbCrLf & _
                                    "INNER JOIN Produtos Prd " & vbCrLf & _
                                    "ON Prd.Produto_Id = E.Produto_Id " & vbCrLf & _
                                    "AND Prd.Situacao  = 1 " & vbCrLf & _
                                    "WHERE E.Empresa_Id = '" & Empresa & "' " & vbCrLf & _
                                    "AND E.EndEmpresa_Id = " & EndEmpresa & vbCrLf & _
                                    "AND E.Produto_Id = '" & codigoProduto & "' " & vbCrLf & _
                                    "AND SO.EstoqueFiscal = 'S' " & vbCrLf & _
                                    "AND E.FisicoFiscal_Id = 2 " & vbCrLf & _
                                    "AND E.Movimento_Id < '" & DataAnterior & "' " & vbCrLf & _
                                    "UNION ALL " & vbCrLf & _
                                    "SELECT COALESCE(SUM(NFI.QuantidadeFiscal), 0) AS SaldoAnterior " & vbCrLf & _
                                    "FROM NotasFiscaisXItens NFI " & vbCrLf & _
                                    "INNER JOIN NotasFiscais NF " & vbCrLf & _
                                    "ON  NFI.Empresa_Id = NF.Empresa_Id " & vbCrLf & _
                                    "AND NFI.EndEmpresa_Id = NF.EndEmpresa_Id " & vbCrLf & _
                                    "AND NFI.Cliente_Id = NF.Cliente_Id " & vbCrLf & _
                                    "AND NFI.EndCliente_Id = NF.EndCliente_Id " & vbCrLf & _
                                    "AND NFI.EntradaSaida_Id = NF.EntradaSaida_Id " & vbCrLf & _
                                    "AND NFI.Serie_Id = NF.Serie_Id " & vbCrLf & _
                                    "AND NFI.Nota_Id = NF.Nota_Id " & vbCrLf & _
                                    "INNER JOIN SubOperacoes SO " & vbCrLf & _
                                    "ON  SO.Operacao_Id = NFI.Operacao " & vbCrLf & _
                                    "AND SO.SubOperacoes_Id = NFI.SubOperacao " & vbCrLf & _
                                    "INNER JOIN Produtos Prd " & vbCrLf & _
                                    "ON Prd.Produto_Id = NFI.Produto_Id " & vbCrLf & _
                                    "AND Prd.Situacao  = 1 " & vbCrLf & _
                                    "WHERE NF.Situacao=1  AND NFI.QuantidadeFiscal <> 0 " & vbCrLf & _
                                    "AND NFI.Empresa_Id = '" & Empresa & "' " & vbCrLf & _
                                    "AND NFI.EndEmpresa_Id = " & EndEmpresa & vbCrLf & _
                                    " And NFI.Produto_Id = '" & codigoProduto & "' " & vbCrLf & _
                                    "AND SO.EstoqueFiscal = 'S' " & vbCrLf & _
                                    "AND NFI.EntradaSaida_Id = 'E' AND NF.Situacao=1 " & vbCrLf & _
                                    "AND NF.Serie_id <> '101' AND NF.Movimento < '" & DataAnterior & "' " & vbCrLf & _
                                    "UNION ALL " & vbCrLf & _
                                    "SELECT COALESCE(SUM(NFI.QuantidadeFiscal) * (-1), 0) AS SaldoAnterior " & vbCrLf & _
                                    "FROM NotasFiscaisXItens NFI " & vbCrLf & _
                                    "INNER JOIN NotasFiscais NF " & vbCrLf & _
                                    "ON  NFI.Empresa_Id = NF.Empresa_Id " & vbCrLf & _
                                    "AND NFI.EndEmpresa_Id = NF.EndEmpresa_Id " & vbCrLf & _
                                    "AND NFI.Cliente_Id = NF.Cliente_Id " & vbCrLf & _
                                    "AND NFI.EndCliente_Id = NF.EndCliente_Id " & vbCrLf & _
                                    "AND NFI.EntradaSaida_Id = NF.EntradaSaida_Id " & vbCrLf & _
                                    "AND NFI.Serie_Id = NF.Serie_Id " & vbCrLf & _
                                    "AND NFI.Nota_Id = NF.Nota_Id " & vbCrLf & _
                                    "INNER JOIN SubOperacoes SO " & vbCrLf & _
                                    "ON  SO.Operacao_Id = NFI.Operacao " & vbCrLf & _
                                    "AND SO.SubOperacoes_Id = NFI.SubOperacao " & vbCrLf & _
                                    "INNER JOIN Produtos Prd " & vbCrLf & _
                                    "ON Prd.Produto_Id = NFI.Produto_Id " & vbCrLf & _
                                    "AND Prd.Situacao  = 1 " & vbCrLf & _
                                    "WHERE NF.Situacao=1  AND NFI.QuantidadeFiscal <> 0 " & vbCrLf & _
                                    "AND NFI.Empresa_Id = '" & Empresa & "' " & vbCrLf & _
                                    "AND NFI.EndEmpresa_Id = " & EndEmpresa & vbCrLf & _
                                    " And NFI.Produto_Id = '" & codigoProduto & "' " & vbCrLf & _
                                    "AND SO.EstoqueFiscal = 'S' " & vbCrLf & _
                                    " AND NF.Serie_id <> '101' AND NF.Movimento < '" & DataAnterior & "' " & vbCrLf & _
                                    "AND NFI.EntradaSaida_Id = 'S' AND NF.Situacao=1 ) SA " & vbCrLf & _
                                    "UNION ALL " & vbCrLf & _
                                    "SELECT MOV.Entradas ,MOV.Saidas , 0 AS Saldo " & vbCrLf & _
                                    "FROM (SELECT COALESCE(E.Entradas, 0) AS Entradas, COALESCE(E.Saidas, 0) AS Saidas" & vbCrLf & _
                                    "FROM Producao E " & vbCrLf & _
                                    "INNER JOIN SubOperacoes SO " & vbCrLf & _
                                    "ON SO.Operacao_Id = E.Operacao_Id " & vbCrLf & _
                                    "AND SO.SubOperacoes_Id = E.SubOperacao_Id " & vbCrLf & _
                                    "INNER JOIN Produtos Prd " & vbCrLf & _
                                    "ON Prd.Produto_Id = E.Produto_Id " & vbCrLf & _
                                    "AND Prd.Situacao  = 1 " & vbCrLf & _
                                    "WHERE E.Empresa_Id = '" & Empresa & "' " & vbCrLf & _
                                    "AND E.EndEmpresa_Id = " & EndEmpresa & vbCrLf & _
                                    "AND E.Produto_Id = '" & codigoProduto & "' " & vbCrLf & _
                                    "AND SO.EstoqueFiscal = 'S' " & vbCrLf & _
                                    "AND E.FisicoFiscal_Id = 2 " & vbCrLf & _
                                    "AND E.Movimento_Id BETWEEN CONVERT(DATETIME, '" & DataAnterior & "', 102) " & vbCrLf & _
                                    "AND CONVERT(DATETIME, '" & DataAtual & "', 102) " & vbCrLf & _
                                    "UNION ALL " & vbCrLf & _
                                    "SELECT 0 AS Entradas, SUM(opXI.Quantidade) AS Saidas" & vbCrLf & _
                                    "FROM OrdemDeProducaoXInsumo opXI" & vbCrLf & _
                                    "	INNER JOIN OrdemDeProducao op" & vbCrLf & _
                                    "			on op.Empresa_id     = opXI.Empresa_id" & vbCrLf & _
                                    "			and op.EndEmpresa_Id = opXI.EndEmpresa_Id" & vbCrLf & _
                                    "			and op.Ordem_Id      = opXI.Ordem_Id" & vbCrLf & _
                                    "WHERE opXI.Empresa_Id    = '" & Empresa & "'" & vbCrLf & _
                                    "  and opXI.EndEmpresa_Id = " & EndEmpresa & vbCrLf & _
                                    "  and opXI.Produto_id    = '" & codigoProduto & "' " & vbCrLf & _
                                    "  and op.Situacao        = 1" & vbCrLf & _
                                    "  and op.Estoque         = 0" & vbCrLf & _
                                    "UNION ALL " & vbCrLf & _
                                    "SELECT " & vbCrLf & _
                                    "sum(CASE WHEN NFI.EntradaSaida_Id = 'E' THEN COALESCE(NFI.QuantidadeFiscal, 0) " & vbCrLf & _
                                    "ELSE 0 END) AS Entradas, " & vbCrLf & _
                                    "sum(CASE WHEN NFI.EntradaSaida_Id = 'S' THEN COALESCE(NFI.QuantidadeFiscal, 0) " & vbCrLf & _
                                    "ELSE 0 END) AS Saidas " & vbCrLf & _
                                    "FROM NotasFiscaisXItens NFI " & vbCrLf & _
                                    "INNER JOIN NotasFiscais NF " & vbCrLf & _
                                    "ON  NFI.Empresa_Id = NF.Empresa_Id " & vbCrLf & _
                                    "AND NFI.EndEmpresa_Id = NF.EndEmpresa_Id " & vbCrLf & _
                                    "AND NFI.Cliente_Id = NF.Cliente_Id " & vbCrLf & _
                                    "AND NFI.EndCliente_Id = NF.EndCliente_Id " & vbCrLf & _
                                    "AND NFI.EntradaSaida_Id = NF.EntradaSaida_Id " & vbCrLf & _
                                    "AND NFI.Serie_Id = NF.Serie_Id " & vbCrLf & _
                                    "AND NFI.Nota_Id = NF.Nota_Id " & vbCrLf & _
                                    "LEFT JOIN NotaFiscalXLote NFL" & vbCrLf & _
                                    "ON  NFL.Empresa_Id = NFI.Empresa_Id " & vbCrLf & _
                                    "AND NFL.EndEmpresa_Id = NFI.EndEmpresa_Id " & vbCrLf & _
                                    "AND NFL.Cliente_Id = NFI.Cliente_Id " & vbCrLf & _
                                    "AND NFL.EndCliente_Id = NFI.EndCliente_Id " & vbCrLf & _
                                    "AND NFL.EntradaSaida_Id = NFI.EntradaSaida_Id " & vbCrLf & _
                                    "AND NFL.Serie_Id = NFI.Serie_Id " & vbCrLf & _
                                    "AND NFL.Nota_Id = NFI.Nota_Id " & vbCrLf & _
                                    "AND NFL.Produto_Id = NFI.Produto_Id " & vbCrLf & _
                                    "AND NFL.CFOP_Id = NFI.CFOP_Id " & vbCrLf & _
                                    "AND NFL.Sequencia_Id = NFI.Sequencia_Id " & vbCrLf & _
                                    "INNER JOIN SubOperacoes SO " & vbCrLf & _
                                    "ON  SO.Operacao_Id = NFI.Operacao " & vbCrLf & _
                                    "AND SO.SubOperacoes_Id = NFI.SubOperacao " & vbCrLf & _
                                    "INNER JOIN Produtos Prd " & vbCrLf & _
                                    "ON Prd.Produto_Id = NFI.Produto_Id " & vbCrLf & _
                                    "AND Prd.Situacao  = 1 " & vbCrLf & _
                                    "WHERE NF.Situacao=1  AND NFI.QuantidadeFiscal <> 0 " & vbCrLf & _
                                    "AND NFI.Empresa_Id = '" & Empresa & "' " & vbCrLf & _
                                    "AND NFI.EndEmpresa_Id = " & EndEmpresa & vbCrLf & _
                                    " And NFI.Produto_Id = '" & codigoProduto & "' " & vbCrLf & _
                                    "AND SO.EstoqueFiscal = 'S' " & vbCrLf & _
                                    "AND NF.Movimento BETWEEN CONVERT(DATETIME, '" & DataAnterior & "', 102) " & vbCrLf & _
                                    "AND CONVERT(DATETIME, '" & DataAtual & "', 102)  AND NF.Situacao=1 " & vbCrLf & _
                                    "group by NF.Operacao, NF.SubOperacao" & vbCrLf & _
                                    " ) MOV " & vbCrLf

            strSQL &= "select (Sum(Saldo) + sum(Entradas) - Sum(Saidas)) as Saldo from #Estoque" & vbCrLf

            ds = db.ConsultaDataSet(strSQL, "verEstoque")

            For Each row As DataRow In ds.Tables(0).Rows
                quantidade = row("Saldo")
            Next

        Catch ex As Exception
            Throw ex
        Finally
            db = Nothing
        End Try

        Return quantidade
    End Function


    Public Function Salvar() As Boolean
        Dim objBanco As New AcessaBanco()
        Dim sqls As New ArrayList

        SalvarSql(sqls)
        Return objBanco.GravaBanco(sqls)
    End Function

    Public Sub SalvarSql(ByRef Sqls As ArrayList)
        Dim strSql As String = ""

        Select Case Me.IUD
            Case "I"

                strSql = " IF EXISTS( " & vbCrLf &
                            "           SELECT * FROM OrdemDeProducaoXInsumo" & vbCrLf &
                            "           WHERE Empresa_Id = '" & OrdemParaProducao.CodigoEmpresa & "' " & vbCrLf &
                            "               AND EndEmpresa_Id = " & OrdemParaProducao.EnderecoEmpresa & " " & vbCrLf &
                            "               AND Ordem_Id = " & OrdemParaProducao.Codigo & " " & vbCrLf &
                            "               AND Produto_Id = '" & Me.CodigoProduto & "') " & vbCrLf &
                            "BEGIN" & vbCrLf &
                            "           UPDATE OrdemDeProducaoXInsumo SET Quantidade = " & Str(Me.Quantidade) & " " & vbCrLf &
                            "           WHERE Empresa_Id = '" & OrdemParaProducao.CodigoEmpresa & "' " & vbCrLf &
                            "               AND EndEmpresa_Id = " & OrdemParaProducao.EnderecoEmpresa & " " & vbCrLf &
                            "               AND Ordem_Id = " & OrdemParaProducao.Codigo & " " & vbCrLf &
                            "               AND ProdutoProducao_Id = '" & Me.CodigoProdutoProducao & "' " & vbCrLf &
                            "               AND Produto_Id = '" & Me.CodigoProduto & "' " & vbCrLf &
                            "END" & vbCrLf &
                            "ELSE" & vbCrLf &
                            "BEGIN" & vbCrLf &
                            "   INSERT INTO OrdemDeProducaoXInsumo(Empresa_Id, EndEmpresa_Id, Ordem_Id, ProdutoProducao_Id, Produto_Id, Quantidade)" & vbCrLf &
                                                     "VALUES('" & OrdemParaProducao.CodigoEmpresa & "'," & OrdemParaProducao.EnderecoEmpresa & "," & OrdemParaProducao.Codigo & ",'" & Me.CodigoProdutoProducao & "', '" & Me.CodigoProduto & "'," & Str(Me.Quantidade) & ")" & vbCrLf &
                            "END"

            Case "U"
                strSql = "Update OrdemDeProducaoXInsumo Set " & vbCrLf &
                         "   Quantidade                         = " & Str(Me.Quantidade) & vbCrLf &
                         " WHERE Empresa_Id                     = '" & OrdemParaProducao.CodigoEmpresa & "'" & vbCrLf &
                         "  AND EndEmpresa_Id                   = " & OrdemParaProducao.EnderecoEmpresa & vbCrLf &
                         "  AND Ordem_Id                        = " & OrdemParaProducao.Codigo & vbCrLf &
                         "  AND ProdutoProducao_Id              = '" & Me.CodigoProdutoProducao & "' " & vbCrLf &
                         "  AND Produto_Id                      = '" & Me.CodigoProduto & "'"

            Case "D"

                strSql = "Delete OrdemDeProducaoXInsumo " & vbCrLf &
                         " WHERE Empresa_Id    = '" & OrdemParaProducao.CodigoEmpresa & "'" & vbCrLf &
                         "  AND EndEmpresa_Id = " & OrdemParaProducao.EnderecoEmpresa & vbCrLf &
                         "  AND Ordem_Id      = " & OrdemParaProducao.Codigo & vbCrLf &
                         "  AND ProdutoProducao_Id              = '" & Me.CodigoProdutoProducao & "' " & vbCrLf &
                         "  AND Produto_Id    = '" & Me.CodigoProduto & "'"

            Case "IN_CON_INS"

                strSql = "INSERT INTO OrdemDeProducaoXInsumo(Empresa_Id, EndEmpresa_Id, Ordem_Id, ProdutoProducao_Id, Produto_Id, Quantidade)" & vbCrLf &
                         "Values('" & OrdemParaProducao.CodigoEmpresa & "'," & OrdemParaProducao.EnderecoEmpresa & "," & OrdemParaProducao.Codigo & ",'" & Me.CodigoProdutoProducao & "', '" & Me.CodigoProduto & "'," & Str(Me.Quantidade) & ")"

        End Select

        Sqls.Add(strSql)

        If Me.OrdemParaProducao.Estoque AndAlso Me.Produto.ControlarEstoque Then GeraEstoqueInsumo(Sqls)

    End Sub

    Private Sub GeraEstoqueInsumo(ByRef sqls As ArrayList)
        Dim strSql As String = ""

        If Me.SubOperacaoInsumo.EstoqueFisico AndAlso Me.Quantidade > 0 Then 'Estoque Físico

            strSql = "DELETE FROM Producao " &
                    " WHERE Empresa_Id = '" & Me.OrdemParaProducao.CodigoEmpresa & "' AND EndEmpresa_Id = " & Me.OrdemParaProducao.EnderecoEmpresa & " AND OrdemDeProducao = '" & Me.OrdemParaProducao.Codigo & "' AND Produto_Id = '" & Me.CodigoProduto & "' AND FisicoFiscal_Id = 1; "

            sqls.Add(strSql)

            strSql = "INSERT INTO Producao " &
                     "(Empresa_Id, EndEmpresa_Id, Deposito_Id, EndDeposito_Id, Produto_Id, Operacao_Id, SubOperacao_Id, " & vbCrLf &
                     "Movimento_Id, FisicoFiscal_Id, ProdutoDerivado_Id, Lote_Id, Classificacao_Id, Embalagem_Id, TipoDeEmbalagem_Id, " & vbCrLf &
                     " CapacidadeEmbalagem_Id, Etapa, Safra, Entradas, Saidas, Observacao, UsuarioInclusao, UsuarioInclusaoData, OrdemDeProducao) " & vbCrLf &
                     "VALUES ('" & Me.OrdemParaProducao.CodigoEmpresa & "', " & Me.OrdemParaProducao.EnderecoEmpresa & ", '" & Me.OrdemParaProducao.CodigoEmpresa & "', " & Me.OrdemParaProducao.EnderecoEmpresa & ", " & vbCrLf &
                     "'" & Me.CodigoProduto & "', " & Me.SubOperacaoInsumo.CodigoOperacao & ", " & Me.SubOperacaoInsumo.Codigo & ", '" & Me.OrdemParaProducao.MovimentoEstoque.ToString("yyyy-MM-dd") & "', " & vbCrLf &
                     " 1, '" & Me.CodigoProdutoProducao & "', '', '" & Me.OrdemParaProducao.Codigo & "', 0, '', 0, 1, 'NENHUM', 0, " & Str(Me.Quantidade) & ", '', " & vbCrLf &
                     "'" & UsuarioServidor.NomeUsuario & "', GETDATE(), " & Me.OrdemParaProducao.Codigo & ")"

            sqls.Add(strSql)

        Else

            strSql = "DELETE FROM Producao " &
                    " WHERE Empresa_Id = '" & Me.OrdemParaProducao.CodigoEmpresa & "' AND EndEmpresa_Id = " & Me.OrdemParaProducao.EnderecoEmpresa & " AND OrdemDeProducao = '" & Me.OrdemParaProducao.Codigo & "' AND Produto_Id = '" & Me.CodigoProduto & "' AND FisicoFiscal_Id = 1; "

            sqls.Add(strSql)

        End If

        If Me.SubOperacaoInsumo.EstoqueFiscal AndAlso Me.Quantidade > 0 Then 'Estoque Fiscal

            strSql = "DELETE FROM Producao " &
                    " WHERE Empresa_Id = '" & Me.OrdemParaProducao.CodigoEmpresa & "' AND EndEmpresa_Id = " & Me.OrdemParaProducao.EnderecoEmpresa & " AND OrdemDeProducao = '" & Me.OrdemParaProducao.Codigo & "' AND Produto_Id = '" & Me.CodigoProduto & "' AND FisicoFiscal_Id = 2; "

            sqls.Add(strSql)

            strSql = "INSERT INTO Producao " &
                     "(Empresa_Id, EndEmpresa_Id, Deposito_Id, EndDeposito_Id, Produto_Id, Operacao_Id, SubOperacao_Id, " & vbCrLf &
                     "Movimento_Id, FisicoFiscal_Id, ProdutoDerivado_Id, Lote_Id, Classificacao_Id, Embalagem_Id, TipoDeEmbalagem_Id, " & vbCrLf &
                     " CapacidadeEmbalagem_Id, Etapa, Safra, Entradas, Saidas, Observacao, UsuarioInclusao, UsuarioInclusaoData, OrdemDeProducao) " & vbCrLf &
                     "VALUES ('" & Me.OrdemParaProducao.CodigoEmpresa & "', " & Me.OrdemParaProducao.EnderecoEmpresa & ", '" & Me.OrdemParaProducao.CodigoEmpresa & "', " & Me.OrdemParaProducao.EnderecoEmpresa & ", " & vbCrLf &
                     "'" & Me.CodigoProduto & "', " & Me.SubOperacaoInsumo.CodigoOperacao & ", " & Me.SubOperacaoInsumo.Codigo & ", '" & Me.OrdemParaProducao.MovimentoEstoque.ToString("yyyy-MM-dd") & "', " & vbCrLf &
                     " 2, '" & Me.CodigoProdutoProducao & "', '', '" & Me.OrdemParaProducao.Codigo & "', 0, '', 0, 1, 'NENHUM', 0, " & Str(Me.Quantidade) & ", '', " & vbCrLf &
                     "'" & UsuarioServidor.NomeUsuario & "', GETDATE(), " & Me.OrdemParaProducao.Codigo & ")"

            sqls.Add(strSql)

        Else

            strSql = "DELETE FROM Producao " &
                     " WHERE Empresa_Id = '" & Me.OrdemParaProducao.CodigoEmpresa & "' AND EndEmpresa_Id = " & Me.OrdemParaProducao.EnderecoEmpresa & " AND OrdemDeProducao = '" & Me.OrdemParaProducao.Codigo & "' AND Produto_Id = '" & Me.CodigoProduto & "' AND FisicoFiscal_Id = 2; "

            sqls.Add(strSql)

        End If

    End Sub

#End Region

End Class
