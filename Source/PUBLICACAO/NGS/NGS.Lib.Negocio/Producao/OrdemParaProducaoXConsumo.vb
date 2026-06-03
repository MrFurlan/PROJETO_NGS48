Public Class ListOrdemParaProducaoXConsumo
    Inherits List(Of OrdemParaProducaoXConsumo)
    Implements IBaseEntity

#Region "Contrutor"
    Public Sub New()
    End Sub

    Public Sub New(ByVal objOrdemDeProducao As OrdemParaProducao)

        Me.OrdemParaProducao = objOrdemDeProducao

        Dim db As New AcessaBanco()

        Try
            Dim strSQL As String

            strSQL = "Select Empresa_Id, EndEmpresa_Id, Ordem_Id, ProdutoProducao_Id, Produto_Id, Quantidade, QuantidadeDeAjuste, Sinal " & vbCrLf &
                     "From OrdemDeProducaoXConsumo " & vbCrLf &
                     "Where Empresa_id    = '" & Me.OrdemParaProducao.CodigoEmpresa & "'" & vbCrLf &
                     "  and EndEmpresa_Id = " & Me.OrdemParaProducao.EnderecoEmpresa & vbCrLf

            If Me.OrdemParaProducao.Codigo > 0 Then strSQL &= "  and Ordem_Id      = " & Me.OrdemParaProducao.Codigo & vbCrLf

            Dim ds As DataSet = db.ConsultaDataSet(strSQL, "OrdemParaProducaoXConsumo")

            For Each row As DataRow In ds.Tables(0).Rows
                Dim oPC As New OrdemParaProducaoXConsumo(Me.OrdemParaProducao)

                oPC.CodigoProdutoProducao = row("ProdutoProducao_Id").ToString()
                oPC.CodigoProduto = row("Produto_Id").ToString()
                oPC.Produto = New Produto(row("Produto_Id").ToString())
                oPC.Quantidade = row("Quantidade")
                oPC.QuantidadeDeAjuste = row("QuantidadeDeAjuste")
                oPC.Sinal = row("Sinal").ToString()

                Me.Add(oPC)
            Next

        Catch ex As Exception
            Throw ex
        Finally
            db = Nothing
        End Try
    End Sub

    Public Sub SalvarSql(ByRef Sqls As ArrayList)
        For Each item As OrdemParaProducaoXConsumo In Me
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

<Serializable()>
Public Class OrdemParaProducaoXConsumo
    Implements IBaseEntity

#Region "Fields"
    Private _IUD As String = ""

    Private _OrdemParaProducao As OrdemParaProducao

    Private _CodigoProduto As String
    Private _Produto As Produto
    Private _CodigoProdutoProducao As String

    Private _Quantidade As Decimal
    Private _QuantidadeDeAjuste As Decimal

    Private _Percentual As Decimal

    Private _Sinal As String

    Private _SubOperacaoConsumo As SubOperacao

    Private _ItensDeConsumoXLote As ListOrdemParaProducaoXConsumoXLote

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

    Public Property QuantidadeDeAjuste() As Decimal
        Get
            Return _QuantidadeDeAjuste
        End Get
        Set(ByVal value As Decimal)
            _QuantidadeDeAjuste = value
        End Set
    End Property

    Public Property Percentual() As Decimal
        Get
            If _Percentual = 0 Then

                Dim objPercentual = New ListProdutoXConsumos()

                If _CodigoProdutoProducao Is Nothing AndAlso _CodigoProduto.Length > 0 Then
                    For Each produtoDaProducao As OrdemParaProducaoXProduto In Me.OrdemParaProducao.ProdutosdaProducao
                        objPercentual.CarregarConsumo(produtoDaProducao.CodigoProduto, _CodigoProduto)
                        If objPercentual.Count > 0 Then
                            _Percentual = objPercentual(0).Percentual
                            Exit For
                        End If
                    Next
                ElseIf _CodigoProdutoProducao.Length > 0 AndAlso _CodigoProduto.Length > 0 Then

                    objPercentual.CarregarConsumo(_CodigoProdutoProducao, _CodigoProduto)

                    If objPercentual.Count > 0 Then
                        _Percentual = objPercentual(0).Percentual
                    End If

                End If
            End If

            Return _Percentual
        End Get
        Set(ByVal value As Decimal)
            _QuantidadeDeAjuste = value
        End Set
    End Property

    Public Property Sinal() As String
        Get
            Return _Sinal
        End Get
        Set(ByVal value As String)
            _Sinal = value
        End Set
    End Property

    Public Property SubOperacaoConsumo() As SubOperacao
        Get
            If _SubOperacaoConsumo Is Nothing OrElse _SubOperacaoConsumo.Codigo = 0 Then _SubOperacaoConsumo = New SubOperacao(40, 52)

            Return _SubOperacaoConsumo
        End Get
        Set(value As SubOperacao)
            _SubOperacaoConsumo = value
        End Set
    End Property

    Public Property ItensDeConsumoXLote() As ListOrdemParaProducaoXConsumoXLote
        Get
            If _ItensDeConsumoXLote Is Nothing Then _ItensDeConsumoXLote = New ListOrdemParaProducaoXConsumoXLote(Me)

            Return _ItensDeConsumoXLote
        End Get
        Set(ByVal value As ListOrdemParaProducaoXConsumoXLote)
            _ItensDeConsumoXLote = value
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

    Public Function buscarLoteDeFornecedor(ByVal Empresa As String, ByVal EndEmpresa As Integer, ByVal codigoProduto As String, Optional ByVal grupo As Boolean = False) As DataSet
        Dim db As New AcessaBanco()
        Dim ds As New DataSet

        'Tirei fora conforme conversa com Rose em 18/01/2021 para Reservar mesmo ainda não estando em estoque
        '                                    "  and Estoque       = 1" & vbCrLf & _

        Try
            Dim strSQL As String = "SELECT nXt.Produto_Id AS Produto, nXt.Lote_Id As Lote, nXt.Fabricado, nXt.Validade," & vbCrLf &
                                    "    sum(Case" & vbCrLf &
                                    "            when nXt.EntradaSaida_Id = 'E'" & vbCrLf &
                                    "            then nXt.Quantidade - isnull(nXt.QuantidadeDeConsumo,0)" & vbCrLf &
                                    "            else nXt.Quantidade * -1 " & vbCrLf &
                                    "        end) Quantidade" & vbCrLf &
                                    "INTO #Consumo" & vbCrLf &
                                    "FROM NotaFiscalXLote nXt" & vbCrLf &
                                    "	inner join NotasFiscais n" & vbCrLf &
                                    "		    ON n.Empresa_Id       = nXt.Empresa_Id" & vbCrLf &
                                    "		    and n.EndEmpresa_iD   = nXt.EndEmpresa_Id" & vbCrLf &
                                    "			and n.Cliente_Id      = nXt.Cliente_Id" & vbCrLf &
                                    "			and n.EndCliente_Id   = nXt.EndCliente_Id" & vbCrLf &
                                    "			and n.EntradaSaida_Id = nXt.EntradaSaida_Id" & vbCrLf &
                                    "			and n.Serie_Id        = nXt.Serie_Id" & vbCrLf &
                                    "			and n.Nota_Id         = nXt.Nota_Id" & vbCrLf &
                                    "WHERE nXt.Empresa_Id    = '" & Empresa & "'" & vbCrLf &
                                    "  and nXt.EndEmpresa_Id = " & EndEmpresa & vbCrLf &
                                    "  and n.Situacao        = 1 " & vbCrLf

            If grupo Then
                strSQL &= "  and LEFT(nXt.Produto_id,5) in(" & codigoProduto & ")" & vbCrLf
            Else
                strSQL &= "  and nXt.Produto_id    in(" & codigoProduto & ")" & vbCrLf
            End If
            strSQL &= "GROUP BY nXt.Produto_Id, nXt.Lote_Id, nXt.Fabricado, nXt.Validade" & vbCrLf &
                        "" & vbCrLf &
                        "" & vbCrLf


            strSQL &= "Insert into #Consumo (Produto, Lote, Fabricado, Validade, Quantidade)" & vbCrLf &
                        "select OPXP.Produto_Id, OPXP.Lote, OP.Movimento AS Fabricado, OP.Validade," & vbCrLf &
                        "		SUM(CASE" & vbCrLf &
                        "			WHEN OPXP.QuantidadeDeAjuste > 0" & vbCrLf &
                        "				THEN OPXP.QuantidadeDeAjuste" & vbCrLf &
                        "				ELSE OPXP.Quantidade" & vbCrLf &
                        "			END) AS Quantidade" & vbCrLf &
                        "FROM OrdemDeProducao OP " & vbCrLf &
                        "INNER JOIN OrdemDeProducaoXProduto OPXP " & vbCrLf &
                        "   ON OP.Empresa_Id		= OPXP.Empresa_Id " & vbCrLf &
                        "   AND OP.EndEmpresa_Id	= OPXP.EndEmpresa_Id " & vbCrLf &
                        "   AND OP.Ordem_Id			= OPXP.Ordem_Id " & vbCrLf &
                        "WHERE OP.Empresa_Id           = '" & Empresa & "'" & vbCrLf &
                        "   AND OP.EndEmpresa_Id       = " & EndEmpresa & vbCrLf &
                        "   AND OP.Estoque             = 1" & vbCrLf
            If grupo Then
                strSQL &= "     AND LEFT(OPXP.Produto_Id,5) IN (" & codigoProduto & ")" & vbCrLf
            Else
                strSQL &= "     AND OPXP.Produto_Id       IN (" & codigoProduto & ")" & vbCrLf
            End If
            strSQL &= "         AND OP.Situacao      = 1" & vbCrLf &
                        "GROUP BY OPXP.Produto_Id, OPXP.Lote, OP.Movimento, OP.Validade" & vbCrLf &
                        "" & vbCrLf &
                        "" & vbCrLf


            strSQL &= "Insert into #Consumo (Produto, Lote, Fabricado, Validade, Quantidade)" & vbCrLf &
                        "SELECT Produto_Id AS Produto, Lote_Id As Lote, Movimento AS Fabricado, Validade, Quantidade" & vbCrLf &
                        "FROM OrdemDeProducaoInterna" & vbCrLf &
                        "WHERE Empresa_Id    = '" & Empresa & "'" & vbCrLf &
                        "  and EndEmpresa_Id = " & EndEmpresa & vbCrLf
            If grupo Then
                strSQL &= "  and LEFT(Produto_id,5) in(" & codigoProduto & ")" & vbCrLf
            Else
                strSQL &= "  and Produto_id    in(" & codigoProduto & ")" & vbCrLf
            End If
            strSQL &= "  and EntradaSaida  = 'E'" & vbCrLf &
                        "" & vbCrLf &
                        "" & vbCrLf &
                        "" & vbCrLf


            strSQL &= "INSERT INTO #Consumo (Produto, Lote, Fabricado, Validade, Quantidade)" & vbCrLf &
                        "Select p.Produto_Id AS Produto, p.Lote_Id As Lote, p.Fabricado, p.Validade," & vbCrLf &
                        "		sum(Case" & vbCrLf &
                        "				when so.EntradaSaida = 'E'" & vbCrLf &
                        "				then p.Entradas" & vbCrLf &
                        "				else p.Saidas * -1" & vbCrLf &
                        "			end) Quantidade" & vbCrLf &
                        "FROM Producao p" & vbCrLf &
                        "		inner join SubOperacoes so" & vbCrLf &
                        "				on so.Operacao_Id      = p.Operacao_Id" & vbCrLf &
                        "				and so.SubOperacoes_Id = p.SubOperacao_Id" & vbCrLf &
                        "WHERE p.Empresa_Id    = '" & Empresa & "'" & vbCrLf &
                        "  and p.EndEmpresa_Id = " & EndEmpresa & vbCrLf &
                        "  and p.FisicoFiscal_Id = 2" & vbCrLf &
                        "  and Len(isnull(p.Lote_Id,'')) > 0 " & vbCrLf &
                        "  and isnull(p.OrdemDeProducao,0) = 0" & vbCrLf &
                        "  and not p.Fabricado is null" & vbCrLf
            If grupo Then
                strSQL &= "  and LEFT(p.Produto_Id,5) in(" & codigoProduto & ")" & vbCrLf
            Else
                strSQL &= "  and p.Produto_Id    in(" & codigoProduto & ")" & vbCrLf
            End If
            strSQL &= "group by p.Produto_Id, p.Lote_Id, p.Fabricado, p.Validade" & vbCrLf &
                        "" & vbCrLf &
                        "" & vbCrLf


            strSQL &= "SELECT OP.Empresa_Id, OP.EndEmpresa_Id, OPXP.Produto_Id, OPXP.Lote, OP.Movimento As Fabricado" & vbCrLf &
                        "INTO #OrdemPRD " & vbCrLf &
                        "FROM OrdemDeProducao OP " & vbCrLf &
                        "INNER JOIN OrdemDeProducaoXProduto OPXP " & vbCrLf &
                        "   ON OP.Empresa_Id		= OPXP.Empresa_Id " & vbCrLf &
                        "   AND OP.EndEmpresa_Id	= OPXP.EndEmpresa_Id " & vbCrLf &
                        "   AND OP.Ordem_Id			= OPXP.Ordem_Id " & vbCrLf &
                        "WHERE OP.Empresa_Id    = '" & Empresa & "'" & vbCrLf &
                        "   AND OP.EndEmpresa_Id = " & EndEmpresa & vbCrLf &
                        "   AND OP.Estoque       = 1" & vbCrLf &
                        "   AND OP.Situacao      = 1" & vbCrLf
            If grupo Then
                strSQL &= "  AND LEFT(OPXP.Produto_Id,5) in(" & codigoProduto & ")" & vbCrLf
            Else
                strSQL &= "  AND OPXP.Produto_Id    in(" & codigoProduto & ")" & vbCrLf
            End If
            strSQL &= "" & vbCrLf &
                        "" & vbCrLf


            strSQL &= "INSERT INTO #OrdemPRD (Empresa_Id, EndEmpresa_Id, Produto_Id, Lote, Fabricado)" & vbCrLf &
                        "SELECT Empresa_Id, EndEmpresa_Id, Produto_Id, Lote_Id, Movimento" & vbCrLf &
                        "FROM OrdemDeProducaoInterna" & vbCrLf &
                        "WHERE Empresa_Id    = '" & Empresa & "'" & vbCrLf &
                        "  and EndEmpresa_Id = " & EndEmpresa & vbCrLf
            If grupo Then
                strSQL &= "  and LEFT(Produto_Id,5) in(" & codigoProduto & ")" & vbCrLf
            Else
                strSQL &= "  and Produto_Id    in(" & codigoProduto & ")" & vbCrLf
            End If
            strSQL &= "" & vbCrLf &
                    "" & vbCrLf


            strSQL &= "SELECT opXL.Produto_Id AS Produto, opXL.Lote_Id As Lote," & vbCrLf &
                        "		case" & vbCrLf &
                        "			when SUBSTRING(opXL.Lote_Id,4,1) = '.'" & vbCrLf &
                        "				then odP.Fabricado" & vbCrLf &
                        "				else nXt.Fabricado" & vbCrLf &
                        "			end Fabricado," & vbCrLf &
                        "opXL.Validade, SUM(opXL.Quantidade) AS Quantidade" & vbCrLf &
                        "   INTO #ConsumoLote" & vbCrLf &
                        "   FROM OrdemDeProducaoXConsumoXLote opXL" & vbCrLf &
                        "	INNER JOIN OrdemDeProducao op" & vbCrLf &
                        "		ON op.Empresa_id            = opXL.Empresa_id" & vbCrLf &
                        "		AND op.EndEmpresa_Id        = opXL.EndEmpresa_Id" & vbCrLf &
                        "		AND op.Ordem_Id             = opXL.Ordem_Id" & vbCrLf &
                        "	LEFT JOIN (select Empresa_id, EndEmpresa_Id, Produto_Id, Lote_Id, Max(Fabricado) AS Fabricado " & vbCrLf &
                        "				from NotaFiscalXLote" & vbCrLf &
                        "				group by Empresa_id, EndEmpresa_Id, Produto_Id, Lote_Id) AS nXt" & vbCrLf &
                        "		on nXt.Empresa_Id       = opXL.Empresa_id" & vbCrLf &
                        "		and nXt.EndEmpresa_Id   = opXL.EndEmpresa_Id" & vbCrLf &
                        "		and nXt.Produto_Id      = opXL.Produto_Id" & vbCrLf &
                        "		and nXt.Lote_Id         = opXL.Lote_Id" & vbCrLf &
                        "	LEFT JOIN (select Empresa_id, EndEmpresa_Id, Produto_Id, Lote, Max(Fabricado) AS Fabricado" & vbCrLf &
                        "				from #OrdemPRD" & vbCrLf &
                        "				group by Empresa_id, EndEmpresa_Id, Produto_Id, Lote) AS odP" & vbCrLf &
                        "		on odP.Empresa_Id     = opXL.Empresa_id" & vbCrLf &
                        "		and odP.EndEmpresa_Id = opXL.EndEmpresa_Id" & vbCrLf &
                        "		and odp.Produto_Id    = opXL.Produto_Id" & vbCrLf &
                        "		and odP.Lote          = opXL.Lote_Id" & vbCrLf &
                        "   WHERE opXL.Empresa_Id    = '" & Empresa & "'" & vbCrLf &
                        "       and opXL.EndEmpresa_Id = " & EndEmpresa & vbCrLf &
                        "       and op.Estoque         = 1" & vbCrLf
            If grupo Then
                strSQL &= "  and LEFT(opXL.Produto_id,5) in(" & codigoProduto & ")" & vbCrLf
            Else
                strSQL &= "  and opXL.Produto_id    in(" & codigoProduto & ")" & vbCrLf
            End If
            strSQL &= "  and op.Situacao        = 1" & vbCrLf &
                        "GROUP BY opXL.Produto_Id, opXL.Lote_Id, nXt.Fabricado, odP.Fabricado, opXL.Validade" & vbCrLf &
                        "" & vbCrLf &
                        "" & vbCrLf &
                        "" & vbCrLf


            strSQL &= "SELECT Produto_Id AS Produto, Lote_Id As Lote, Movimento AS Fabricado, Validade, Quantidade" & vbCrLf &
                        "INTO #ConsumoLoteInterno" & vbCrLf &
                        "FROM OrdemDeProducaoInterna" & vbCrLf &
                        "WHERE Empresa_Id    = '" & Empresa & "'" & vbCrLf &
                        "  and EndEmpresa_Id = " & EndEmpresa & vbCrLf
            If grupo Then
                strSQL &= "  and LEFT(Produto_id,5) in(" & codigoProduto & ")" & vbCrLf
            Else
                strSQL &= "  and Produto_id    in(" & codigoProduto & ")" & vbCrLf
            End If
            strSQL &= "  and EntradaSaida  = 'S'" & vbCrLf &
                                    "" & vbCrLf &
                                    "" & vbCrLf


            strSQL &= "SELECT c.Produto, c.Lote, c.Fabricado, c.Validade, sum(c.Quantidade - isnull(cl.Quantidade,0) - isnull(clI.Quantidade,0)) as Quantidade, 0 AS Consumo" & vbCrLf &
                        "FROM #Consumo c" & vbCrLf &
                        "	LEFT JOIN #ConsumoLote cl" & vbCrLf &
                        "			on cl.Produto    = c.Produto" & vbCrLf &
                        "			and cl.Lote      = c.Lote" & vbCrLf &
                        "			and cl.Fabricado = c.Fabricado" & vbCrLf &
                        "			and cl.Validade  = c.Validade" & vbCrLf &
                        "	LEFT JOIN #ConsumoLoteInterno clI" & vbCrLf &
                        "			on clI.Produto    = c.Produto" & vbCrLf &
                        "			and clI.Lote      = c.Lote" & vbCrLf &
                        "			and clI.Fabricado = c.Fabricado" & vbCrLf &
                        "			and clI.Validade  = c.Validade" & vbCrLf &
                        "	WHERE c.Validade >= CONVERT(DATE, CONVERT(VARCHAR(10),  GETDATE(), 102)) " & vbCrLf &
                        "GROUP BY c.Produto, c.Lote, c.Fabricado, c.Validade" & vbCrLf &
                        "HAVING sum(c.Quantidade - isnull(cl.Quantidade,0) - isnull(clI.Quantidade,0)) > 0" & vbCrLf &
                        "ORDER BY c.Validade asc"


            ds = db.ConsultaDataSet(strSQL, "LoteDeFornecedor")
        Catch ex As Exception
            Throw ex
        Finally
            db = Nothing
        End Try

        Return ds
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

                strSql = "INSERT INTO OrdemDeProducaoXConsumo(Empresa_Id, EndEmpresa_Id, Ordem_Id, ProdutoProducao_Id, Produto_Id, Quantidade, QuantidadeDeAjuste, Sinal)" & vbCrLf &
                         "Values('" & OrdemParaProducao.CodigoEmpresa & "'," & OrdemParaProducao.EnderecoEmpresa & "," & OrdemParaProducao.Codigo & ",'" & Me.CodigoProdutoProducao & "', '" & Me.CodigoProduto & "'," & Str(Me.Quantidade) & "," & Str(Me.QuantidadeDeAjuste) & ",'" & Me.Sinal & "')"

                Sqls.Add(strSql)

                salvarRelacionados(Sqls)

            Case "U"
                strSql = "Update OrdemDeProducaoXConsumo Set " & vbCrLf &
                         "   Quantidade                     = " & Str(Me.Quantidade) & vbCrLf &
                         "  ,QuantidadeDeAjuste             = " & Str(Me.QuantidadeDeAjuste) & vbCrLf &
                         "  ,Sinal                          = '" & Me.Sinal & "'" & vbCrLf &
                         " Where Empresa_Id                 = '" & OrdemParaProducao.CodigoEmpresa & "'" & vbCrLf &
                         "   and EndEmpresa_Id              = " & OrdemParaProducao.EnderecoEmpresa & vbCrLf &
                         "   and Ordem_Id                   = " & OrdemParaProducao.Codigo & vbCrLf &
                         "   and ProdutoProducao_Id         = '" & Me.CodigoProdutoProducao & "'" & vbCrLf &
                         "   and Produto_Id                 = '" & Me.CodigoProduto & "'"

                Sqls.Add(strSql)

                salvarRelacionados(Sqls)

            Case "D"
                salvarRelacionados(Sqls)

                strSql = "Delete OrdemDeProducaoXConsumo " & vbCrLf &
                         " Where Empresa_Id                 = '" & OrdemParaProducao.CodigoEmpresa & "'" & vbCrLf &
                         "   and EndEmpresa_Id              = " & OrdemParaProducao.EnderecoEmpresa & vbCrLf &
                         "   and Ordem_Id                   = " & OrdemParaProducao.Codigo & vbCrLf &
                         "   and ProdutoProducao_Id         = '" & Me.CodigoProdutoProducao & "'" & vbCrLf &
                         "   and Produto_Id                 = '" & Me.CodigoProduto & "'"

                Sqls.Add(strSql)

            Case "DEL_CONS"

                salvarRelacionados(Sqls)

                strSql = "Delete OrdemDeProducaoXConsumo " & vbCrLf &
                         " Where Empresa_Id    = '" & OrdemParaProducao.CodigoEmpresa & "'" & vbCrLf &
                         "   and EndEmpresa_Id = " & OrdemParaProducao.EnderecoEmpresa & vbCrLf &
                         "   and Ordem_Id      = " & OrdemParaProducao.Codigo & vbCrLf &
                         "   and ProdutoProducao_Id      = '" & Me.CodigoProdutoProducao & "'" & vbCrLf &
                         "   and Produto_Id    = '" & Me.CodigoProduto & "'"

                Sqls.Add(strSql)

        End Select

    End Sub

    Private Sub salvarRelacionados(ByRef sqls As ArrayList)

        If Me.ItensDeConsumoXLote.Count > 0 Then Me.ItensDeConsumoXLote.SalvarSql(sqls)

    End Sub

#End Region

End Class
