Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Collections.Generic
Imports NGS.Lib.Uteis

<Serializable()>
Public Class ListNotaFiscalDevolucaoXNotaFiscal
    Inherits List(Of NotaFiscalDevolucaoXNotaFiscal)

#Region "Fields"
    Private _itemNota As NotaFiscalXItem
    Private _Valida As Boolean
    Private _SomaQtde As Decimal
    Private _SomaVlr As Decimal
    Dim _MsgDevolucao As String = ""
#End Region

#Region "Property"
    Public ReadOnly Property ItemNota As NotaFiscalXItem
        Get
            Return _itemNota
        End Get
    End Property

    Public Property Valida() As Boolean
        Get
            Return _Valida
        End Get
        Set(ByVal value As Boolean)
            _Valida = value
        End Set
    End Property

    Public ReadOnly Property SomaQtde() As Decimal
        Get
            Return (From x In Me Select x.QuantidadeDevolucao).Sum()
        End Get
    End Property

    Public ReadOnly Property SomaVlr() As Decimal
        Get
            Return (From x In Me Select x.ValorDevolucao).Sum()
        End Get
    End Property

    Public Property MsgDevolucao() As String
        Get
            If _MsgDevolucao.Length = 0 Then

                If Me.Count > 1 Then
                    _MsgDevolucao = "Devolucao das Notas do Produto " & _itemNota.Produto.Nome & ": "
                Else
                    _MsgDevolucao = "Devolucao da Nota do Produto " & _itemNota.Produto.Nome & ": "
                End If

                Dim virgula As String = String.Empty
                For Each row As NotaFiscalDevolucaoXNotaFiscal In Me
                    If row.QuantidadeDevolucao > 0 Or row.ValorDevolucao > 0 Then
                        _MsgDevolucao &= row.NumeroNota & "-" & row.Serie & virgula
                        virgula = " , "
                        If row.QuantidadeNota <> row.QuantidadeDevolucao Or row.ValorNota <> row.ValorDevolucao Then
                            _MsgDevolucao &= " Parcial"
                        End If
                    End If
                Next
            End If
            Return _MsgDevolucao
        End Get
        Set(ByVal value As String)
            _MsgDevolucao = value
        End Set
    End Property
#End Region

#Region "Construtor"
    Public Sub New(ByRef nfxi As NotaFiscalXItem)
        _itemNota = nfxi
    End Sub

    Public Sub New()

    End Sub
#End Region

#Region "Methods"
    Public Sub SalvarSql(ByVal Sqls As ArrayList)
        If _itemNota.IUD = "U" Then
            Dim sql As String
            sql = " Delete NotaFiscalDevolucaoXNotaFiscal " & vbCrLf &
                  " Where EmpresaDevolucao_Id      ='" & _itemNota.NotaFiscal.CodigoEmpresa & "'" & vbCrLf &
                  "   and EndEmpresaDevolucao_Id   = " & _itemNota.NotaFiscal.EnderecoEmpresa & vbCrLf &
                  "   and ClienteDevolucao_Id      ='" & _itemNota.NotaFiscal.CodigoCliente & "'" & vbCrLf &
                  "   and EndClienteDevolucao_Id   = " & _itemNota.NotaFiscal.EnderecoCliente & vbCrLf &
                  "   and EntradaSaidaDevolucao_Id ='" & _itemNota.NotaFiscal.EntradaSaida.ToString.Substring(0, 1) & "'" & vbCrLf &
                  "   and SerieDevolucao_Id        ='" & _itemNota.NotaFiscal.Serie & "'" & vbCrLf &
                  "   and NotaDevolucao_Id         = " & _itemNota.NotaFiscal.Codigo & vbCrLf &
                  "   and SequenciaDevolucao_id    = " & _itemNota.Sequencia & vbCrLf &
                  "   and CFOPDevolucao_id         = " & _itemNota.CFOP & vbCrLf &
                  "   and Produto_Id               ='" & _itemNota.CodigoProduto & "'" & vbCrLf
            Sqls.Add(sql)
        End If

        For Each nfd As NotaFiscalDevolucaoXNotaFiscal In Me
            If _itemNota.IUD = "U" Then nfd.IUD = "I"
            If _itemNota.IUD = "D" Or _itemNota.IUD = "I" Then nfd.IUD = _itemNota.IUD

            If nfd.IUD <> "" And (nfd.QuantidadeDevolucao > 0 Or nfd.ValorDevolucao > 0) Then
                nfd.SalvarSql(Sqls)
            End If
        Next
    End Sub

    Public Sub CarregarNotasParaSelecao(Optional ByVal UnitarioMedio As Decimal = 0)
        Dim Sql As String
        Sql = "SELECT NF.Empresa_Id," & vbCrLf &
              "       NF.EndEmpresa_Id," & vbCrLf &
              "       NF.Cliente_Id," & vbCrLf &
              "       NF.EndCliente_Id," & vbCrLf &
              "       NF.EntradaSaida_Id," & vbCrLf &
              "       NF.Serie_Id," & vbCrLf &
              "       NF.Nota_Id," & vbCrLf &
              "       NF.DataDaNota," & vbCrLf &
              "       NFxI.Produto_Id," & vbCrLf &
              "       NFxI.Sequencia_Id," & vbCrLf &
              "       NFxI.CFOP_Id," & vbCrLf &
              "       NFxI.Unitario as UnitarioNota," & vbCrLf &
              "       NFxI.QuantidadeFiscal as QuantidadeNota," & vbCrLf &
              "       NFxI.Valor as ValorNota," & vbCrLf &
              "       isnull(sb.Quantidade,0) as QdteDevolvido," & vbCrLf &
              "       isnull(sb.Valor,0)      as ValorDevolvido," & vbCrLf &
              "       isnull(sbfix.QtdeFix,0) as QtdeFix," & vbCrLf &
              "       CONVERT(DECIMAL(25,13), isnull(sbfix.VlrFix,0))  as VlrFix" & vbCrLf &
              "  FROM NotasFiscais NF" & vbCrLf &
              " INNER JOIN NotasFiscaisXItens NFxI" & vbCrLf &
              "    ON NF.Empresa_Id      = NFxI.Empresa_Id" & vbCrLf &
              "   AND NF.EndEmpresa_Id   = NFxI.EndEmpresa_Id" & vbCrLf &
              "   AND NF.Cliente_Id      = NFxI.Cliente_Id" & vbCrLf &
              "   AND NF.EndCliente_Id   = NFxI.EndCliente_Id" & vbCrLf &
              "   AND NF.EntradaSaida_Id = NFxI.EntradaSaida_Id" & vbCrLf &
              "   AND NF.Serie_Id        = NFxI.Serie_Id" & vbCrLf &
              "   AND NF.Nota_Id         = NFxI.Nota_Id" & vbCrLf &
              " Inner Join Suboperacoes SO" & vbCrLf &
              "    ON NF.Operacao        = SO.Operacao_Id" & vbCrLf &
              "   AND NF.SubOperacao     = SO.SubOperacoes_Id" & vbCrLf &
              "  LEFT JOIN (" & vbCrLf &
              "               SELECT fxn.Empresa_Id, fxn.EndEmpresa_Id," & vbCrLf &
              "                      fxn.Cliente_Id, fxn.EndCliente_Id," & vbCrLf &
              "                      fxn.EntradaSaida_Id, fxn.Serie_Id, fxn.Nota_Id," & vbCrLf &
              "                      fxn.Produto_Id," & vbCrLf &
              "			             Sum(fxn.Quantidade) as QtdeFix," & vbCrLf &
              "					     --SUM(fxn.Quantidade * fix.UnitarioOficial) as VlrFix" & vbCrLf &
              "                      --SUM(fxn.valor) as VlrFix " & vbCrLf &
              "                      SUM(fxn.Quantidade * CASE WHEN fix.UnitarioOficial < NFxI.Unitario THEN  fix.UnitarioOficial ELSE NFxI.Unitario END) as VlrFix " & vbCrLf &
              "                 FROM FixacaoXNotaFiscal fxn" & vbCrLf &
              "                INNER JOIN PedidosXItensXFixacoes fix" & vbCrLf &
              "			          ON fxn.Empresa_Id    = fix.Empresa_Id" & vbCrLf &
              "				     AND fxn.EndEmpresa_Id = fix.EndEmpresa_Id" & vbCrLf &
              "				     AND fxn.Pedido_Id     = fix.Pedido_Id" & vbCrLf &
              "				     AND fxn.Fixacao_Id    = fix.Fixacao_Id" & vbCrLf &
              "				     AND fxn.Produto_Id    = fix.Produto_Id" & vbCrLf &
              "                 JOIN NotasfiscaisXItens NfxI" & vbCrLf &
              "	                  ON NFxI.Empresa_id      = fxn.Empresa_Id" & vbCrLf &
              "	                 AND NFxI.EndEmpresa_Id   = fxn.EndEmpresa_Id" & vbCrLf &
              "	                 AND NFxI.Cliente_Id      = fxn.Cliente_Id" & vbCrLf &
              "	                 AND NFxI.EndCliente_Id   = fxn.EndCliente_Id" & vbCrLf &
              "	                 AND NFxI.EntradaSaida_Id = fxn.EntradaSaida_Id" & vbCrLf &
              "	                 AND NFxI.Serie_Id        = fxn.Serie_Id" & vbCrLf &
              "	                 AND NFxI.Nota_Id         = fxn.Nota_Id" & vbCrLf &
              "	                 AND NFxI.Produto_Id      = fxn.Produto_Id" & vbCrLf &
              "	                 AND NFxI.Sequencia_id    = fxn.Sequencia_Id" & vbCrLf &
              " 	             AND NFxI.CFOP_id         = fxn.CFOP_Id" & vbCrLf &
              "              GROUP BY fxn.Empresa_Id, fxn.EndEmpresa_Id, fxn.Cliente_Id, fxn.EndCliente_Id, fxn.EntradaSaida_Id, fxn.Serie_Id, fxn.Nota_Id, fxn.Produto_Id" & vbCrLf &
              "             ) SbFix" & vbCrLf &
              "    On NFxI.Empresa_id      = SbFix.Empresa_Id" & vbCrLf &
              "   and NFxI.EndEmpresa_Id   = SbFix.EndEmpresa_Id" & vbCrLf &
              "   and NFxI.Cliente_Id      = SbFix.Cliente_Id" & vbCrLf &
              "   and NFxI.EndCliente_Id   = SbFix.EndCliente_Id" & vbCrLf &
              "   and NFxI.EntradaSaida_Id = SbFix.EntradaSaida_Id" & vbCrLf &
              "   and NFxI.Serie_Id        = SbFix.Serie_Id" & vbCrLf &
              "   and NFxI.Nota_Id         = SbFix.Nota_Id" & vbCrLf &
              "   and NFxI.Produto_Id      = SbFix.Produto_Id" & vbCrLf &
              "  Left Join (" & vbCrLf &
              "		 	    SELECT nfd.EmpresaDevolucao_Id as Empresa_Id," & vbCrLf &
              "                    nfd.EndEmpresaDevolucao_Id as EndEmpresa_Id," & vbCrLf &
              "                    nfd.ClienteDevolucao_Id as Cliente_Id," & vbCrLf &
              "                    nfd.EndClienteDevolucao_Id as EndCliente_Id," & vbCrLf &
              "                    nfd.EntradaSaida_Id," & vbCrLf &
              "                    nfd.Serie_Id," & vbCrLf &
              "                    nfd.Nota_Id," & vbCrLf &
              "                    nfd.Produto_Id," & vbCrLf &
              "                    nfd.Sequencia_Id," & vbCrLf &
              "                    nfd.CFOP_Id," & vbCrLf &
              "                    sum(nfd.Quantidade) as Quantidade," & vbCrLf &
              "                    sum(nfd.Valor) as Valor" & vbCrLf &
              "			      FROM NotaFiscalDevolucaoXNotaFiscal nfd" & vbCrLf &
              "              Inner Join NotasFiscais nf" & vbCrLf &
              "                 On nf.Empresa_id      = nfd.EmpresaDevolucao_Id" & vbCrLf &
              "                and nf.EndEmpresa_Id   = nfd.EndEmpresaDevolucao_Id" & vbCrLf &
              "                and nf.Cliente_Id      = nfd.ClienteDevolucao_Id" & vbCrLf &
              "                and nf.EndCliente_Id   = nfd.EndClienteDevolucao_Id" & vbCrLf &
              "                and nf.EntradaSaida_Id = nfd.EntradaSaidaDevolucao_Id" & vbCrLf &
              "                and nf.Serie_Id        = nfd.SerieDevolucao_Id" & vbCrLf &
              "                and nf.Nota_Id         = nfd.NotaDevolucao_Id" & vbCrLf &
              "               JOIN NotasFiscaisXItens AS NFXI " & vbCrLf &
              "                 ON NF.Empresa_Id      = NFXI.Empresa_Id " & vbCrLf &
              "                AND NF.EndEmpresa_Id   = NFXI.EndEmpresa_Id " & vbCrLf &
              "                AND NF.Cliente_Id      = NFXI.Cliente_Id " & vbCrLf &
              "                AND NF.EndCliente_Id   = NFXI.EndCliente_Id " & vbCrLf &
              "                AND NF.EntradaSaida_Id = NFXI.EntradaSaida_Id " & vbCrLf &
              "                AND NF.Serie_Id        = NFXI.Serie_Id " & vbCrLf &
              "                AND NF.Nota_Id         = NFXI.Nota_Id" & vbCrLf &
              "              Where nf.situacao in (1, 4)" & vbCrLf &
              "                And NF.Empresa_Id ='" & _itemNota.NotaFiscal.CodigoEmpresa & "'" & vbCrLf &
              "                And NF.EndEmpresa_Id ='" & _itemNota.NotaFiscal.EnderecoEmpresa & "'" & vbCrLf &
              "                And not (NF.Nota_Id  =" & _itemNota.NotaFiscal.Codigo & " And NF.Serie_id ='" & _itemNota.NotaFiscal.Serie & "')" & vbCrLf &
              "                And NFXI.Produto_Id ='" & _itemNota.CodigoProduto & "'" & vbCrLf

        If _itemNota.Lote.Length > 0 Then
            Sql &= "                AND nfxi.Lote           ='" & _itemNota.Lote & "'" & vbCrLf &
                   "                AND nfxi.Classificacao  ='" & _itemNota.Classificacao & "'"
        End If

        Sql &= "              Group by nfd.EmpresaDevolucao_Id, nfd.EndEmpresaDevolucao_Id, nfd.ClienteDevolucao_Id, nfd.EndClienteDevolucao_Id, nfd.EntradaSaida_Id, nfd.Serie_Id, nfd.Nota_Id, nfd.Produto_Id, nfd.Sequencia_Id, nfd.cfop_id" & vbCrLf &
              "             ) Sb" & vbCrLf &
              "    On NFxI.Empresa_id      = Sb.Empresa_Id" & vbCrLf &
              "   and NFxI.EndEmpresa_Id   = Sb.EndEmpresa_Id" & vbCrLf &
              "   and NFxI.Cliente_Id      = sb.Cliente_Id" & vbCrLf &
              "   and NFxI.EndCliente_Id   = Sb.EndCliente_Id" & vbCrLf &
              "   and NFxI.EntradaSaida_Id = Sb.EntradaSaida_Id" & vbCrLf &
              "   and NFxI.Serie_Id        = sb.Serie_Id" & vbCrLf &
              "   and NFxI.Nota_Id         = sb.Nota_Id" & vbCrLf &
              "   and NFxI.Produto_Id      = sb.Produto_Id" & vbCrLf &
              "   and NFxI.Sequencia_id    = sb.Sequencia_Id" & vbCrLf &
              "   and NFxI.CFOP_id         = sb.CFOP_Id" & vbCrLf &
              " WHERE NF.Situacao IN (1,4) " & vbCrLf &
              "   And NF.Empresa_Id ='" & _itemNota.NotaFiscal.CodigoEmpresa & "'" & vbCrLf &
              "   And NF.EndEmpresa_Id ='" & _itemNota.NotaFiscal.EnderecoEmpresa & "'" & vbCrLf

        If _itemNota.CodigoPedido > 0 Then
            Sql &= "   AND NF.Pedido        = " & _itemNota.CodigoPedido & vbCrLf
        End If

        Sql &= "   AND NF.Cliente_id    ='" & _itemNota.NotaFiscal.CodigoCliente & "'" & vbCrLf &
              "   AND NF.EndCliente_Id = " & _itemNota.NotaFiscal.EnderecoCliente & vbCrLf &
              "   and NFxI.Produto_Id  ='" & _itemNota.CodigoProduto & "'" & vbCrLf &
              "   AND SO.Devolucao     ='N'" & vbCrLf &
              "   And ((NFxI.QuantidadeFiscal - isnull(sb.Quantidade,0) - isnull(sbfix.QtdeFix,0) > 0) or (NFxI.Valor - isnull(sb.Valor,0) - isnull(sbfix.VlrFix,0) > 0))" & vbCrLf &
              "   AND nf.DataDaNota <= '" & _itemNota.NotaFiscal.DataNota.ToString("yyyy/MM/dd") & "'" & vbCrLf

        If _itemNota.Lote.Length > 0 Then
            Sql &= "  and nfxi.Lote           ='" & _itemNota.Lote & "'" & vbCrLf &
                   "  and nfxi.Classificacao  ='" & _itemNota.Classificacao & "'"
        End If

        If _itemNota.NotaFiscal.SubOperacao.Classe = eClassesOperacoes.GLOBAL Then
            Sql &= "   And so.classe = '" & eClassesOperacoes.GLOBAL.ToString & "'" & vbCrLf
        End If

        If _itemNota.NotaFiscal.SubOperacao.Classe = eClassesOperacoes.REMESSAS Then
            Sql &= "   And so.classe = '" & eClassesOperacoes.REMESSAS.ToString & "'" & vbCrLf
        End If

        If _itemNota.NotaFiscal.SubOperacao.Classe = eClassesOperacoes.AFIXAR Then
            Sql &= "   And so.classe = '" & eClassesOperacoes.AFIXAR.ToString & "'" & vbCrLf
        End If

        If _itemNota.NotaFiscal.SubOperacao.QuantidadeFiscal Then
            Sql &= "   AND NFxI.QuantidadeFiscal > 0 "
        ElseIf Not _itemNota.NotaFiscal.SubOperacao.QuantidadeFiscal AndAlso UnitarioMedio > 0 Then
            Sql &= "   AND NFxI.QuantidadeFiscal > 0 "
        ElseIf Not _itemNota.NotaFiscal.SubOperacao.QuantidadeFiscal AndAlso UnitarioMedio = 0 Then
            Sql &= "   AND NFxI.QuantidadeFiscal = 0 "
        End If

        Sql &= " Order by NF.Movimento DESC"

        Dim Banco As New AcessaBanco
        Dim ds As New DataSet
        ds = Banco.ConsultaDataSet(Sql, "Notas")

        Dim qtdenota As Decimal = _itemNota.QuantidadeFiscal
        Dim vlrnota As Decimal = _itemNota.ValorTotal

        Dim UnitarioMedioPedido As Decimal = 0

        If _itemNota.CodigoPedido > 0 Then
            UnitarioMedioPedido = _itemNota.Pedido.Itens.Where(Function(s) s.CodigoProduto = _itemNota.CodigoProduto).FirstOrDefault.UnitarioMedioFaturamento
        Else
            UnitarioMedioPedido = vlrnota / qtdenota
        End If

        Dim VlrSaldoNF As Decimal = 0

        If qtdenota > 0 Then
            For Each row As DataRow In ds.Tables(0).Rows
                Dim Nfd As New NotaFiscalDevolucaoXNotaFiscal(_itemNota)
                Nfd.Sequencia = row("Sequencia_Id")
                Nfd.CodigoCFOP = row("CFOP_Id")
                Nfd.EntradaSaida = row("EntradaSaida_Id")
                Nfd.Serie = row("Serie_Id")
                Nfd.NumeroNota = row("Nota_Id")
                Nfd.DataDaNota = row("DataDaNota")

                Nfd.QuantidadeNota = row("QuantidadeNota")
                Nfd.UnitarioNota = row("UnitarioNota")
                Nfd.ValorNota = row("ValorNota")
                Nfd.QuantidadeDevolvido = row("QdteDevolvido")
                Nfd.ValorDevolvido = row("ValorDevolvido")

                Nfd.QuantidadeFixado = row("QtdeFix")
                Nfd.ValorFixado = row("VlrFix")

                If qtdenota > Nfd.QuantidadeSaldo Then
                    If Nfd.QuantidadeSaldo > 0 Then
                        Nfd.QuantidadeDevolucao = Nfd.QuantidadeSaldo
                        Nfd.ValorDevolucao = Nfd.ValorSaldo
                        qtdenota -= Nfd.QuantidadeSaldo
                    End If
                Else
                    Nfd.QuantidadeDevolucao = qtdenota
                    Nfd.ValorDevolucao = qtdenota * Nfd.UnitarioNota
                    qtdenota = 0
                End If

                Nfd.NotaOrigem = False

                If Me.Any(Function(x) x.EntradaSaida = Nfd.EntradaSaida And x.NumeroNota = Nfd.NumeroNota AndAlso x.Serie = Nfd.Serie) Then
                    Me.RemoveAll(Function(x) x.EntradaSaida = Nfd.EntradaSaida And x.NumeroNota = Nfd.NumeroNota AndAlso x.Serie = Nfd.Serie)
                End If

                Me.Add(Nfd)

            Next
        Else
            For Each row As DataRow In ds.Tables(0).Rows
                Dim Nfd As New NotaFiscalDevolucaoXNotaFiscal(_itemNota)
                Nfd.Sequencia = row("Sequencia_Id")
                Nfd.CodigoCFOP = row("CFOP_Id")
                Nfd.EntradaSaida = row("EntradaSaida_Id")
                Nfd.Serie = row("Serie_Id")
                Nfd.NumeroNota = row("Nota_Id")
                Nfd.DataDaNota = row("DataDaNota")

                Nfd.QuantidadeNota = row("QuantidadeNota")
                Nfd.UnitarioNota = row("UnitarioNota")
                Nfd.ValorNota = row("ValorNota")
                Nfd.QuantidadeDevolvido = row("QdteDevolvido")
                Nfd.ValorDevolvido = row("ValorDevolvido")

                Nfd.QuantidadeFixado = row("QtdeFix")
                Nfd.ValorFixado = row("VlrFix")

                If UnitarioMedio <= 0 Then
                    VlrSaldoNF = Nfd.ValorSaldo - Math.Round(Nfd.QuantidadeSaldo * IIf(Nfd.UnitarioNota < UnitarioMedioPedido, Nfd.UnitarioNota, UnitarioMedioPedido), 2, MidpointRounding.AwayFromZero)
                Else
                    VlrSaldoNF = Nfd.ValorSaldo - Math.Round(Nfd.QuantidadeSaldo * IIf(Nfd.UnitarioNota < UnitarioMedio, Nfd.UnitarioNota, UnitarioMedio), 2, MidpointRounding.AwayFromZero)
                End If

                If vlrnota > VlrSaldoNF Then
                    'If Nfd.QuantidadeSaldo = 0 OrElse Nfd.QuantidadeSaldo = Nfd.QuantidadeNota Then
                    Nfd.ValorDevolucao = VlrSaldoNF
                    vlrnota -= VlrSaldoNF
                    'End If
                Else
                    Nfd.ValorDevolucao = vlrnota
                    vlrnota = 0
                End If

                Nfd.NotaOrigem = False

                If Me.Any(Function(x) x.EntradaSaida = Nfd.EntradaSaida And x.NumeroNota = Nfd.NumeroNota AndAlso x.Serie = Nfd.Serie) Then
                    Me.RemoveAll(Function(x) x.EntradaSaida = Nfd.EntradaSaida And x.NumeroNota = Nfd.NumeroNota AndAlso x.Serie = Nfd.Serie)
                End If

                Me.Add(Nfd)

            Next
        End If

        If qtdenota <> 0 Or vlrnota <> 0 Then
            _Valida = False
        Else
            _Valida = True
        End If
    End Sub

    Public Sub CarregarNotasUsadasNaDevolucao()
        Dim Sql As String
        Sql = "SELECT nfd.EntradaSaida_Id," & vbCrLf &
              "       nfd.Serie_Id," & vbCrLf &
              "       nfd.Nota_Id," & vbCrLf &
              "       nfd.Sequencia_id," & vbCrLf &
              "       nfd.CFOP_Id," & vbCrLf &
              "       nf.DataDaNota," & vbCrLf &
              "       nfd.Produto_Id," & vbCrLf &
              "       nfd.Quantidade," & vbCrLf &
              "       case" & vbCrLf &
              "          when nfd.Quantidade = 0 OR valor = 0" & vbCrLf &
              "             then 0" & vbCrLf &
              "             else round(nfd.Quantidade / valor,10)" & vbCrLf &
              "        end as unitario," & vbCrLf &
              "       nfd.Valor," & vbCrLf &
              "       0 AS NotaOrigem" & vbCrLf &
              "	 FROM NotaFiscalDevolucaoXNotaFiscal nfd" & vbCrLf &
              " Inner Join NotasFiscais NF" & vbCrLf &
              "    On nfd.EmpresaDevolucao_Id      = NF.Empresa_Id" & vbCrLf &
              "   and nfd.EndEmpresaDevolucao_Id   = NF.EndEmpresa_Id" & vbCrLf &
              "   and nfd.ClienteDevolucao_Id      = NF.Cliente_Id" & vbCrLf &
              "   and nfd.EndClienteDevolucao_Id   = NF.EndCliente_Id" & vbCrLf &
              "   and nfd.EntradaSaida_Id          = NF.EntradaSaida_Id" & vbCrLf &
              "   and nfd.Serie_Id                 = NF.Serie_Id" & vbCrLf &
              "   and nfd.Nota_Id                  = NF.Nota_Id" & vbCrLf &
              " WHERE nfd.EmpresaDevolucao_Id      ='" & _itemNota.NotaFiscal.CodigoEmpresa & "'" & vbCrLf &
              "   AND nfd.EndEmpresaDevolucao_Id   = " & _itemNota.NotaFiscal.EnderecoEmpresa & vbCrLf &
              "   AND nfd.ClienteDevolucao_Id      ='" & _itemNota.NotaFiscal.CodigoCliente & "'" & vbCrLf &
              "   AND nfd.EndClienteDevolucao_Id   = " & _itemNota.NotaFiscal.EnderecoCliente & vbCrLf &
              "   AND nfd.EntradaSaidaDevolucao_Id ='" & IIf(_itemNota.NotaFiscal.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "'" & vbCrLf &
              "   AND nfd.SerieDevolucao_Id        ='" & _itemNota.NotaFiscal.Serie & "'" & vbCrLf &
              "   AND nfd.NotaDevolucao_Id         = " & _itemNota.NotaFiscal.Codigo & vbCrLf &
              "   AND nfd.Produto_id               = " & _itemNota.CodigoProduto & vbCrLf &
              "   AND nfd.SequenciaDevolucao_Id    = " & _itemNota.Sequencia & vbCrLf &
              "Union All" & vbCrLf &
              "SELECT nfd.EntradaSaidaDevolucao_Id," & vbCrLf &
              "       nfd.SerieDevolucao_Id," & vbCrLf &
              "       nfd.NotaDevolucao_Id," & vbCrLf &
              "       nfd.SequenciaDevolucao_Id," & vbCrLf &
              "       nfd.CFOPDevolucao_Id," & vbCrLf &
              "       nf.DataDaNota," & vbCrLf &
              "       nfd.Produto_Id," & vbCrLf &
              "       nfd.Quantidade," & vbCrLf &
              "       case" & vbCrLf &
              "          when nfd.Quantidade = 0 OR valor = 0 " & vbCrLf &
              "             then 0" & vbCrLf &
              "             else round(nfd.Quantidade / valor,10)" & vbCrLf &
              "        end as unitario," & vbCrLf &
              "       nfd.Valor," & vbCrLf &
              "       1 AS NotaOrigem" & vbCrLf &
              "	 FROM NotaFiscalDevolucaoXNotaFiscal nfd" & vbCrLf &
              " Inner Join NotasFiscais NF" & vbCrLf &
              "    On nfd.EmpresaDevolucao_Id      = NF.Empresa_Id" & vbCrLf &
              "   and nfd.EndEmpresaDevolucao_Id   = NF.EndEmpresa_Id" & vbCrLf &
              "   and nfd.ClienteDevolucao_Id      = NF.Cliente_Id" & vbCrLf &
              "   and nfd.EndClienteDevolucao_Id   = NF.EndCliente_Id" & vbCrLf &
              "   and nfd.EntradaSaidaDevolucao_Id = NF.EntradaSaida_Id" & vbCrLf &
              "   and nfd.SerieDevolucao_Id        = NF.Serie_Id" & vbCrLf &
              "   and nfd.NotaDevolucao_Id         = NF.Nota_Id" & vbCrLf &
              " WHERE nfd.EmpresaDevolucao_Id      ='" & _itemNota.NotaFiscal.CodigoEmpresa & "'" & vbCrLf &
              "   AND nfd.EndEmpresaDevolucao_Id   = " & _itemNota.NotaFiscal.EnderecoEmpresa & vbCrLf &
              "   AND nfd.ClienteDevolucao_Id      ='" & _itemNota.NotaFiscal.CodigoCliente & "'" & vbCrLf &
              "   AND nfd.EndClienteDevolucao_Id   = " & _itemNota.NotaFiscal.EnderecoCliente & vbCrLf &
              "   AND nfd.EntradaSaida_Id ='" & IIf(_itemNota.NotaFiscal.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "'" & vbCrLf &
              "   AND nfd.Serie_Id        ='" & _itemNota.NotaFiscal.Serie & "'" & vbCrLf &
              "   AND nfd.Nota_Id         = " & _itemNota.NotaFiscal.Codigo & vbCrLf &
              "   AND nfd.Produto_id      = " & _itemNota.CodigoProduto & vbCrLf &
              "   AND nfd.Sequencia_Id    = " & _itemNota.Sequencia & vbCrLf

        Dim Banco As New AcessaBanco
        Dim ds As New DataSet
        ds = Banco.ConsultaDataSet(Sql, "Notas")

        For Each row As DataRow In ds.Tables(0).Rows
            Dim Nfd As New NotaFiscalDevolucaoXNotaFiscal(_itemNota)
            Nfd.Sequencia = row("Sequencia_Id")
            Nfd.CodigoCFOP = row("CFOP_Id")
            Nfd.EntradaSaida = row("EntradaSaida_Id")
            Nfd.Serie = row("Serie_Id")
            Nfd.NumeroNota = row("Nota_Id")
            Nfd.DataDaNota = row("DataDaNota")
            Nfd.QuantidadeNota = row("Quantidade")
            Nfd.UnitarioNota = row("unitario")
            Nfd.ValorNota = row("Valor")

            Nfd.QuantidadeDevolucao = row("Quantidade")
            Nfd.ValorDevolucao = row("Valor")

            Nfd.NotaOrigem = row("NotaOrigem")

            Me.Add(Nfd)
        Next
    End Sub
#End Region

End Class

<Serializable()>
Public Class NotaFiscalDevolucaoXNotaFiscal

#Region "Construtor"
    Public Sub New(ByRef nfxi As NotaFiscalXItem)
        _ItemNota = nfxi
    End Sub
#End Region

#Region "Fields"
    Private _IUD As String = ""
    Private _ItemNota As NotaFiscalXItem

    Private _EntradaSaida As String
    Private _Serie As String
    Private _NumeroNota As Integer
    Private _DataDaNota As Date
    Private _Sequencia As Integer
    Private _CodigoCFOP As Integer

    Private _Nota As NotaFiscal

    Private _QuantidadeNota As Decimal
    Private _UnitarioNota As Decimal
    Private _ValorNota As Decimal

    Private _QuantidadeDevolvido As Decimal
    Private _ValorDevolvido As Decimal

    Private _QuantidadeFixado As Decimal
    Private _ValorFixado As Decimal

    Private _QuantidadeDevolucao As Decimal
    Private _ValorDevolucao As Decimal
    Private _ValorLiquidoDevolucao As Decimal

    Private _NotaOrigem As Boolean = False
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

    Public ReadOnly Property ItemNota As NotaFiscalXItem
        Get
            Return _ItemNota
        End Get
    End Property

    Public Property EntradaSaida() As String
        Get
            Return _EntradaSaida
        End Get
        Set(ByVal value As String)
            _EntradaSaida = value
            _Nota = Nothing
        End Set
    End Property

    Public Property Serie() As String
        Get
            Return _Serie
        End Get
        Set(ByVal value As String)
            _Serie = value
            _Nota = Nothing
        End Set
    End Property

    Public Property NumeroNota() As Integer
        Get
            Return _NumeroNota
        End Get
        Set(ByVal value As Integer)
            _NumeroNota = value
            _Nota = Nothing
        End Set
    End Property

    Public Property DataDaNota() As Date
        Get
            Return _DataDaNota
        End Get
        Set(ByVal value As Date)
            _DataDaNota = value
        End Set
    End Property

    Public Property Sequencia() As Integer
        Get
            Return _Sequencia
        End Get
        Set(ByVal value As Integer)
            _Sequencia = value
        End Set
    End Property

    Public Property CodigoCFOP() As Integer
        Get
            Return _CodigoCFOP
        End Get
        Set(ByVal value As Integer)
            _CodigoCFOP = value
        End Set
    End Property

    Public ReadOnly Property Nota() As NotaFiscal
        Get
            If _Nota Is Nothing Then
                Dim NFConsulta As New NotaFiscal
                NFConsulta.CodigoEmpresa = _ItemNota.NotaFiscal.CodigoEmpresa
                NFConsulta.EnderecoEmpresa = _ItemNota.NotaFiscal.EnderecoEmpresa
                NFConsulta.CodigoCliente = _ItemNota.NotaFiscal.CodigoCliente
                NFConsulta.EnderecoCliente = _ItemNota.NotaFiscal.EnderecoCliente
                NFConsulta.EntradaSaida = IIf(_EntradaSaida = "E", eEntradaSaida.Entrada, eEntradaSaida.Saida)
                NFConsulta.Serie = _Serie
                NFConsulta.Codigo = _NumeroNota

                _Nota = New NotaFiscal(NFConsulta)
            End If
            Return _Nota
        End Get
    End Property

    Public Property QuantidadeNota() As Decimal
        Get
            Return _QuantidadeNota
        End Get
        Set(ByVal value As Decimal)
            _QuantidadeNota = value
        End Set
    End Property

    Public Property UnitarioNota() As Decimal
        Get
            Return _UnitarioNota
        End Get
        Set(ByVal value As Decimal)
            _UnitarioNota = value
        End Set
    End Property

    Public Property ValorNota() As Decimal
        Get
            Return _ValorNota
        End Get
        Set(ByVal value As Decimal)
            _ValorNota = value
        End Set
    End Property

    Public Property QuantidadeDevolvido() As Decimal
        Get
            Return _QuantidadeDevolvido
        End Get
        Set(ByVal value As Decimal)
            _QuantidadeDevolvido = value
        End Set
    End Property

    Public Property ValorDevolvido() As Decimal
        Get
            Return _ValorDevolvido
        End Get
        Set(ByVal value As Decimal)
            _ValorDevolvido = value
        End Set
    End Property

    Public Property QuantidadeDevolucao() As Decimal
        Get
            Return _QuantidadeDevolucao
        End Get
        Set(ByVal value As Decimal)
            _QuantidadeDevolucao = value
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

    Public Property QuantidadeFixado As Decimal
        Get
            Return _QuantidadeFixado
        End Get
        Set(value As Decimal)
            _QuantidadeFixado = value
        End Set
    End Property

    Public Property ValorFixado As Decimal
        Get
            Return _ValorFixado
        End Get
        Set(value As Decimal)
            _ValorFixado = value
        End Set
    End Property

    Public ReadOnly Property QuantidadeSaldo() As Decimal
        Get
            Return Me.QuantidadeNota - Me.QuantidadeDevolvido - Me.QuantidadeFixado
        End Get
    End Property

    Public ReadOnly Property ValorSaldo() As Decimal
        Get
            Return Me.ValorNota - Me.ValorDevolvido - Me.ValorFixado
        End Get
    End Property

    Public Property NotaOrigem() As Boolean
        Get
            Return _NotaOrigem
        End Get
        Set(ByVal value As Boolean)
            _NotaOrigem = value
        End Set
    End Property

    Public Property ValorLiquidoDevolucao As Decimal
        Get
            Return _ValorLiquidoDevolucao
        End Get
        Set(value As Decimal)
            _ValorLiquidoDevolucao = value
        End Set
    End Property

#End Region

#Region "Methods"
    Public Sub SalvarSql(ByRef Sqls As ArrayList)
        If IUD = "I" And QuantidadeDevolucao = 0 And ValorDevolucao = 0 Then Exit Sub

        Dim Sql As String
        Select Case IUD
            Case "I"
                sql = "INSERT INTO NotaFiscalDevolucaoXNotaFiscal" & vbCrLf &
                      "  (EmpresaDevolucao_Id, EndEmpresaDevolucao_Id, ClienteDevolucao_Id, EndClienteDevolucao_Id, EntradaSaidaDevolucao_Id, SerieDevolucao_Id, NotaDevolucao_Id, SequenciaDevolucao_Id, CFOPDevolucao_Id," & vbCrLf &
                      "   Produto_Id, EntradaSaida_Id, Serie_Id, Nota_Id, Sequencia_Id, CFOP_Id, Quantidade, Valor, ValorLiquido) " & vbCrLf &
                      " VALUES ('" & _ItemNota.NotaFiscal.CodigoEmpresa & "'," & _ItemNota.NotaFiscal.EnderecoEmpresa & ", " & vbCrLf &
                      "'" & _ItemNota.NotaFiscal.CodigoCliente & "', " & _ItemNota.NotaFiscal.EnderecoCliente & ", " & vbCrLf

                If _NotaOrigem Then
                    Sql &= "'" & _EntradaSaida & "', '" & _Serie & "', " & _NumeroNota & "," & _Sequencia & "," & _CodigoCFOP & "," & vbCrLf &
                          "'" & _ItemNota.CodigoProduto & "'," & vbCrLf &
                          "'" & _ItemNota.NotaFiscal.EntradaSaida.ToString.Substring(0, 1) & "','" & _ItemNota.NotaFiscal.Serie & "'," & _ItemNota.NotaFiscal.Codigo & "," & _ItemNota.Sequencia & "," & _ItemNota.CFOP & "," & Str(_QuantidadeDevolucao) & "," & Str(_ValorDevolucao) & "," & Str(_ValorLiquidoDevolucao) & ")" & vbCrLf

                Else
                    sql &= "'" & _ItemNota.NotaFiscal.EntradaSaida.ToString.Substring(0, 1) & "', '" & _ItemNota.NotaFiscal.Serie & "', " & _ItemNota.NotaFiscal.Codigo & "," & _ItemNota.Sequencia & "," & _ItemNota.CFOP & "," & vbCrLf &
                          "'" & _ItemNota.CodigoProduto & "'," & vbCrLf &
                          "'" & _EntradaSaida & "','" & _Serie & "'," & _NumeroNota & "," & _Sequencia & "," & _CodigoCFOP & "," & Str(_QuantidadeDevolucao) & "," & Str(_ValorDevolucao) & "," & Str(_ValorLiquidoDevolucao) & ")" & vbCrLf
                End If

                Sqls.Add(sql)

                For Each nfDev In _ItemNota.NotasDevolucao.Where(Function(x) x.ValorDevolucao > 0)

                    If Not nfDev.Nota Is Nothing AndAlso nfDev.Nota.Codigo > 0 Then
                        If Not nfDev.Nota.VencimentosNota Is Nothing AndAlso nfDev.Nota.VencimentosNota.Count > 0 Then
                            For Each vencNF In nfDev.Nota.VencimentosNota
                                If vencNF IsNot Nothing And (vencNF.CodigoProvisao = eProvisao.Provisao OrElse vencNF.ValorDoDocumento = 0) Then
                                    sql = " Delete NotaFiscalXTitulo " & vbCrLf &
                                          " Where Empresa_Id      ='" & _ItemNota.NotaFiscal.CodigoEmpresa & "'" & vbCrLf &
                                          "   and EndEmpresa_Id   = " & _ItemNota.NotaFiscal.EnderecoEmpresa & vbCrLf &
                                          "   and Cliente_Id      ='" & _ItemNota.NotaFiscal.CodigoCliente & "'" & vbCrLf &
                                          "   and EndCliente_Id   = " & _ItemNota.NotaFiscal.EnderecoCliente & vbCrLf &
                                          "   and EntradaSaida_Id ='" & _EntradaSaida & "'" & vbCrLf &
                                          "   and Serie_Id        ='" & _Serie & "'" & vbCrLf &
                                          "   and Nota_Id         = " & _NumeroNota & vbCrLf &
                                          "   and Titulo_Id         = " & vencNF.Codigo & vbCrLf

                                    Sqls.Add(sql)
                                End If
                            Next
                        End If
                    End If
                Next

            Case "U"
                sql = "Update NotaFiscalDevolucaoXNotaFiscal set" & vbCrLf &
                      "     Quantidade                  = " & Str(_QuantidadeDevolucao) & vbCrLf &
                      "    ,Valor                       = " & Str(_ValorDevolucao) & vbCrLf &
                      "    ,ValorLiquido                = " & Str(_ValorLiquidoDevolucao) & vbCrLf &
                      " Where EmpresaDevolucao_Id       ='" & _ItemNota.NotaFiscal.CodigoEmpresa & "'" & vbCrLf &
                      "   and EndEmpresaDevolucao_Id    = " & _ItemNota.NotaFiscal.EnderecoEmpresa & vbCrLf &
                      "   and ClienteDevolucao_Id       ='" & _ItemNota.NotaFiscal.CodigoCliente & "'" & vbCrLf &
                      "   and EndClienteDevolucao_Id    = " & _ItemNota.NotaFiscal.EnderecoCliente & vbCrLf &
                      "   and EntradaSaidaDevolucao_Id  ='" & _ItemNota.NotaFiscal.EntradaSaida.ToString.Substring(0, 1) & "'" & vbCrLf &
                      "   and SerieDevolucao_Id         ='" & _ItemNota.NotaFiscal.Serie & "'" & vbCrLf &
                      "   and NotaDevolucao_Id          = " & _ItemNota.NotaFiscal.Codigo & vbCrLf &
                      "   and SequenciaDevolucao_Id     = " & _ItemNota.Sequencia & vbCrLf &
                      "   and CFOPDevolucao_Id          = " & _ItemNota.CFOP & vbCrLf &
                      "   and Produto_Id                ='" & _ItemNota.CodigoProduto & "'" & vbCrLf &
                      "   and EntradaSaida_Id           ='" & _EntradaSaida & "'" & vbCrLf &
                      "   and Serie_Id                  ='" & _Serie & "'" & vbCrLf &
                      "   and Nota_Id                   = " & _NumeroNota & vbCrLf &
                      "   and Sequencia_Id              = " & _Sequencia & vbCrLf &
                      "   and CFOP_Id                   = " & _CodigoCFOP & vbCrLf
                Sqls.Add(sql)
            Case "D"
                sql = " Delete NotaFiscalDevolucaoXNotaFiscal " & vbCrLf &
                      " Where EmpresaDevolucao_Id      ='" & _ItemNota.NotaFiscal.CodigoEmpresa & "'" & vbCrLf &
                      "   and EndEmpresaDevolucao_Id   = " & _ItemNota.NotaFiscal.EnderecoEmpresa & vbCrLf &
                      "   and ClienteDevolucao_Id      ='" & _ItemNota.NotaFiscal.CodigoCliente & "'" & vbCrLf &
                      "   and EndClienteDevolucao_Id   = " & _ItemNota.NotaFiscal.EnderecoCliente & vbCrLf

                If _NotaOrigem Then
                    sql &= "   and EntradaSaidaDevolucao_Id ='" & _EntradaSaida & "'" & vbCrLf &
                    "   and SerieDevolucao_Id        ='" & _Serie & "'" & vbCrLf &
                    "   and NotaDevolucao_Id         = " & _NumeroNota & vbCrLf &
                    "   and SequenciaDevolucao_id    = " & _Sequencia & vbCrLf &
                    "   and CFOPDevolucao_id         = " & _CodigoCFOP & vbCrLf &
                    "   and Produto_Id               ='" & _ItemNota.CodigoProduto & "'" & vbCrLf &
                    "   and EntradaSaida_Id          ='" & _ItemNota.NotaFiscal.EntradaSaida.ToString.Substring(0, 1) & "'" & vbCrLf &
                    "   and Serie_Id                 ='" & _ItemNota.NotaFiscal.Serie & "'" & vbCrLf &
                    "   and Nota_Id                  = " & _ItemNota.NotaFiscal.Codigo & vbCrLf &
                    "   and Sequencia_id             = " & _ItemNota.Sequencia & vbCrLf &
                    "   and CFOP_Id                  = " & _ItemNota.CFOP
                Else
                    sql &= "   and EntradaSaidaDevolucao_Id ='" & _ItemNota.NotaFiscal.EntradaSaida.ToString.Substring(0, 1) & "'" & vbCrLf &
                    "   and SerieDevolucao_Id        ='" & _ItemNota.NotaFiscal.Serie & "'" & vbCrLf &
                    "   and NotaDevolucao_Id         = " & _ItemNota.NotaFiscal.Codigo & vbCrLf &
                    "   and SequenciaDevolucao_id    = " & _ItemNota.Sequencia & vbCrLf &
                    "   and CFOPDevolucao_id         = " & _ItemNota.CFOP & vbCrLf &
                    "   and Produto_Id               ='" & _ItemNota.CodigoProduto & "'" & vbCrLf &
                    "   and EntradaSaida_Id          ='" & _EntradaSaida & "'" & vbCrLf &
                    "   and Serie_Id                 ='" & _Serie & "'" & vbCrLf &
                    "   and Nota_Id                  = " & _NumeroNota & vbCrLf &
                    "   and Sequencia_id             = " & _Sequencia & vbCrLf &
                    "   and CFOP_Id                  = " & _CodigoCFOP
                End If


                Sqls.Add(sql)
        End Select
    End Sub
#End Region

End Class

Public Class RetornoNotasDevolucao
    Public indexItem As Integer
    Public ReadOnly Property Unitario As Decimal
        Get
            If Quantidade = 0 Then
                Return 0
            Else
                Return Valor / Quantidade
            End If
        End Get
    End Property
    Public Property Quantidade As Decimal
    Public Property Valor As Decimal
End Class