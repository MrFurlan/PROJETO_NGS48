Imports Microsoft.VisualBasic
Imports System.Collections.Generic
Imports System.Data
Imports System.Web
Imports NGS.Lib.Uteis
Imports BoletoNet
Imports NGS.Lib.Negocio

<Serializable()>
Public Class ListTitulo
    Inherits List(Of Titulo)

#Region "Construtores"
    'Where
    Public Sub New(ByVal Where As String)
        Dim sql As String
        sql = " Select Registro_id " & vbCrLf &
              "   from contasareceber " & vbCrLf &
              "  Where " & Where & vbCrLf &
              "  Union " & vbCrLf &
              " Select Registro_id " & vbCrLf &
              "   from contasaPagar " & vbCrLf &
              "  Where " & Where

        Dim ds As DataSet
        Dim Banco As New AcessaBanco
        ds = Banco.ConsultaDataSet(sql, "Titulos")
        For Each row As DataRow In ds.Tables(0).Rows
            Dim Tit As New Titulo(row("Registro_Id"))
            Me.Add(Tit)
        Next
        Banco = Nothing
    End Sub

    'Lista Vazia
    Public Sub New()

    End Sub

    'Lista Titulo
    Public Sub New(ByVal pTitulo As Titulo)
        'Dim Banco As New AcessaBanco
        'Dim strSQL As String
        'strSQL = "	SELECT Titulo_id"
        'strSQL = "	  From Titulos"

        'strSQL = "   Where 1 = 1"

        ''If pTitulo.VencMov <> -1 Then
        ''    If pTitulo.VencMov = 0 Then
        ''        strSQL &= " and Vencimento between '" & pTitulo.Vencimento.ToString("yyyy-MM-dd") & "' and '" & pTitulo.VencimentoFinal.ToString("yyyy-MM-dd") & "'"
        ''    Else
        ''        strSQL &= " and Movimento between '" & pTitulo.Movimento.ToString("yyyy-MM-dd") & "' and '" & pTitulo.MovimentoFinal.ToString("yyyy-MM-dd") & "'"
        ''    End If
        ''End If

        'If pTitulo.CodigoSituacao <> 0 Then
        '    strSQL &= " and Situacao_Id = " & pTitulo.Situacao.ToString
        'End If

        'If pTitulo.CodigoProvisao <> 0 Then
        '    strSQL &= " and Provisao_Id = " & pTitulo.Provisao.ToString
        'End If

        'If pTitulo.ReceberPagar <> 0 Then
        '    strSQL &= " and ReceberPagar = '" & IIf(CInt(pTitulo.ReceberPagar) = 1, "R", "P") & "'"
        'End If

        'If Not pTitulo.codigoEmpresaOrigem Is Nothing Then
        '    strSQL &= " and EmpresaOrigem  = '" & pTitulo.EmpresaOrigem & "' "
        '    strSQL &= " and EnderecoOrigem =  " & pTitulo.EnderecoOrigem & " "
        'End If

        'If Not pTitulo.EmpresaDestino Is Nothing Then
        '    strSQL &= " and EmpresaDestino  = '" & pTitulo.EmpresaDestino & "' "
        '    strSQL &= " and EnderecoDestino =  " & pTitulo.EnderecoDestino & " "
        'End If

        'If Not pTitulo.Destinatario Is Nothing Then
        '    strSQL &= " and Destinatario    = '" & pTitulo.Destinatario & "' "
        '    strSQL &= " and EndDestinatario =  " & pTitulo.EndDestinatario & " "
        'End If

        'If Not pTitulo.ClienteCredito Is Nothing Then
        '    strSQL &= " and ClienteCredito    = '" & pTitulo.ClienteCredito & "' "
        '    strSQL &= " and EndClienteCredito =  " & pTitulo.EndClienteCredito & " "
        'End If

        'If Not pTitulo.ClienteDebito Is Nothing Then
        '    strSQL &= " and ClienteDebito    = '" & pTitulo.ClienteDebito & "' "
        '    strSQL &= " and EndClienteDebito =  " & pTitulo.EndClienteDebito & " "
        'End If

        'If pTitulo.FaturaNumero > 0 Then
        '    strSQL &= " and FaturaNumero =  " & pTitulo.FaturaNumero & " "
        'End If

        'If pTitulo.FaturaSerie Is Nothing Then
        '    strSQL &= " and FaturaSerie =  '" & pTitulo.FaturaSerie & "'"
        'End If

        'If pTitulo.Pedido Is Nothing Then
        '    strSQL &= " and Pedido =  '" & pTitulo.Pedido & "'"
        'End If

        'If pTitulo.Titulo <> 0 Then
        '    strSQL &= " and Titulo_id =  " & pTitulo.Titulo & ""
        'End If

        'Banco.dr = Banco.ConsultaBanco(strSQL)
        'While Banco.dr.read
        '    Me.Add(New Titulo(Banco.dr("Registro")))
        'End While
        'Banco.dr.close()
    End Sub

    'Lista Pedido
    Sub New(ByRef CodigoPedido As Integer, ByVal ES As String, ByVal Empresa As String, ByVal EndEmpresa As Integer)
        Dim Banco As New AcessaBanco
        Dim ds As New DataSet
        Dim Sql As String = ""

        Sql = " Select Registro_id, Moeda " & vbCrLf &
              "   From Contasapagar" & vbCrLf &
              "  Where Pedido   = " & CodigoPedido & vbCrLf &
              "    and Empresa = '" & Empresa & "'" & vbCrLf &
              "    and EndEmpresa = " & EndEmpresa &
              "    and Situacao = 1" & vbCrLf &
              "  Union " & vbCrLf &
              " select Registro_id, Moeda " & vbCrLf &
              "   from ContasaReceber" & vbCrLf &
              "  where Pedido   = " & CodigoPedido & vbCrLf &
              "    and Empresa = '" & Empresa & "'" & vbCrLf &
              "    and EndEmpresa = " & EndEmpresa &
              "    and Situacao = 1" & vbCrLf

        ds = Banco.ConsultaDataSet(Sql, "Titulos")

        If ds.Tables(0).Rows.Count = 0 Then
            Exit Sub
        Else
            Dim moeda As New Moeda(ds.Tables(0).Rows(0)("Moeda"))
            _TipoMoeda = moeda.Classificacao
        End If

        For Each row As DataRow In ds.Tables(0).Rows
            Dim Tit As New Titulo(row("Registro_id"))
            Me.Add(Tit)
        Next

        ES_Pedido = ES
    End Sub

    'Lista Nota
    Sub New(ByRef pNF As NotaFiscal)
        _NF = pNF
        If _NF.CodigoTipoDeDocumento = 1 AndAlso _NF.Pedido IsNot Nothing AndAlso _NF.Pedido.Codigo > 0 Then
            _TipoMoeda = _NF.Pedido.Moeda.Classificacao
            ES_Pedido = _NF.Pedido.SubOperacao.EntradaSaida.ToString.Substring(0, 1)
        Else
            _TipoMoeda = eTiposMoeda.Oficial
            ES_Pedido = NF.SubOperacao.EntradaSaida.ToString.Substring(0, 1)
        End If
        Dim Banco As New AcessaBanco
        Dim ds As New DataSet
        Dim Sql As String = ""

        Sql = "SELECT Titulo_Id" & vbCrLf &
              "  FROM NotaFiscalXTitulo" & vbCrLf &
              " Where Empresa_Id      ='" & pNF.CodigoEmpresa & "'" & vbCrLf &
              "   and EndEmpresa_Id   = " & pNF.EnderecoEmpresa & vbCrLf &
              "   and EntradaSaida_Id ='" & IIf(pNF.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "'" & vbCrLf &
              "   and Serie_Id        ='" & pNF.Serie & "'" & vbCrLf &
              "   and Nota_Id         = " & pNF.Codigo & vbCrLf &
              "   and Cliente_Id      ='" & pNF.CodigoCliente & "'" & vbCrLf &
              "   and EndCliente_Id   = " & pNF.EnderecoCliente & vbCrLf &
              "   and Nota_id         > 0"

        ds = Banco.ConsultaDataSet(Sql, "Titulos")

        For Each row As DataRow In ds.Tables(0).Rows
            Dim Tit As New Titulo(row("Titulo_Id"))
            If Tit.CodigoSituacao = 1 OrElse Tit.CodigoSituacao = 101 Then Me.Add(Tit)
        Next

        If Me.Count = 0 AndAlso pNF.CodigoPedido > 0 AndAlso pNF.Pedido.AgruparFinanceiro Then
            Sql = "SELECT NFxT.Titulo_Id" & vbCrLf &
                  "  FROM NotaFiscalXTitulo NFxT" & vbCrLf &
                  " Inner Join NotasFiscais NF" & vbCrLf &
                  "    on NFxT.Empresa_Id      = NF.Empresa_Id" & vbCrLf &
                  "   and NFxT.EndEmpresa_Id   = NF.EndEmpresa_Id" & vbCrLf &
                  "   and NFxT.Cliente_Id      = NF.Cliente_Id" & vbCrLf &
                  "   and NFxT.EndCliente_Id   = NF.EndCliente_Id" & vbCrLf &
                  "   and NFxT.EntradaSaida_Id = NF.EntradaSaida_Id" & vbCrLf &
                  "   and NFxT.Serie_Id        = NF.Serie_Id" & vbCrLf &
                  "   and NFxT.Nota_Id         = NF.Nota_Id " & vbCrLf &
                  " Inner Join (" & vbCrLf &
                  "             SELECT Registro_Id as Titulo_Id, Provisao, isnull(UsuarioLiberacao,'') as UsuarioLiberacao" & vbCrLf &
                  "			      FROM ContasAPagar" & vbCrLf &
                  "			     Union" & vbCrLf &
                  "			    SELECT Registro_Id, Provisao, isnull(UsuarioLiberacao,'') as UsuarioLiberacao" & vbCrLf &
                  "			      FROM ContasAreceber" & vbCrLf &
                  "             ) Titulos" & vbCrLf &
                  "    on Titulos.Titulo_Id = NFxT.Titulo_Id" & vbCrLf &
                  " Where NF.Pedido = (Select Pedido" & vbCrLf &
                  "                      from NotasFiscais" & vbCrLf &
                  "                     Where Empresa_Id      ='" & pNF.CodigoEmpresa & "'" & vbCrLf &
                  "                       and EndEmpresa_Id   = " & pNF.EnderecoEmpresa & vbCrLf &
                  "                       and Cliente_Id      ='" & pNF.Pedido.CodigoCliente & "'" & vbCrLf &
                  "                       and EndCliente_Id   = " & pNF.Pedido.EnderecoCliente & vbCrLf &
                  "                       and EntradaSaida_Id ='" & IIf(pNF.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "'" & vbCrLf &
                  "                       and Serie_Id        ='" & pNF.Serie & "'" & vbCrLf &
                  "                       and Nota_Id         = " & pNF.Codigo & vbCrLf &
                  "                    )" & vbCrLf &
                  "  and Titulos.Provisao         = 2" & vbCrLf &
                  "  and Titulos.UsuarioLiberacao = ''" & vbCrLf &
                  "  and NF.Nota_id > 0"

            ds = Banco.ConsultaDataSet(Sql, "Titulos")

            For Each row As DataRow In ds.Tables(0).Rows
                Dim Tit As New Titulo(row("Titulo_Id"))
                If Tit.CodigoSituacao = 1 OrElse Tit.CodigoSituacao = 101 Then Me.Add(Tit)
            Next
        End If
    End Sub
#End Region

#Region "Fields"
    Private _NF As NotaFiscal
    Private _Pedido As Pedido
    Private _Fixacao As Fixacao

    Private _msg As String
    Private _TipoMoeda As eTiposMoeda
    Private _ES_Pedido As String
    Private _RP_Pedido As String
    Private _ReajFinanceiro As ReajusteFinanceiro

    Private _FRETE As FaturaDeFrete

#End Region

#Region "Property"
    Public Property NF() As NotaFiscal
        Get
            Return _NF
        End Get
        Set(ByVal value As NotaFiscal)
            _NF = value
        End Set
    End Property

    Public Property Pedido As Pedido
        Get
            Return _Pedido
        End Get
        Set(ByVal value As Pedido)
            _Pedido = value
        End Set
    End Property

    Public Property Fixacao As Fixacao
        Get
            Return _Fixacao
        End Get
        Set(ByVal value As Fixacao)
            _Fixacao = value
        End Set
    End Property


    Public Property Mensagem() As String
        Get
            Return _msg
        End Get
        Set(ByVal value As String)
            _msg = value
        End Set
    End Property

    Public Property TipoMoeda() As eTiposMoeda
        Get
            Return _TipoMoeda
        End Get
        Set(ByVal value As eTiposMoeda)
            _TipoMoeda = value
        End Set
    End Property

    Public Property ES_Pedido() As String
        Get
            Return _ES_Pedido
        End Get
        Set(ByVal value As String)
            _ES_Pedido = value
            If value = "S" Then
                _RP_Pedido = "R"
            Else
                _RP_Pedido = "P"
            End If
        End Set
    End Property

    Public Property RP_Pedido() As String
        Get
            Return _RP_Pedido
        End Get
        Set(ByVal value As String)
            _RP_Pedido = value
        End Set
    End Property

    Public Property ReajFinanceiro As ReajusteFinanceiro
        Get
            Return _ReajFinanceiro
        End Get
        Set(value As ReajusteFinanceiro)
            _ReajFinanceiro = value
        End Set
    End Property

    Public Property FRETE As FaturaDeFrete
        Get
            Return _FRETE
        End Get
        Set(value As FaturaDeFrete)
            _FRETE = value
        End Set
    End Property
#End Region

#Region "Métodos"
    Public Function Salvar(ByRef Sqls As ArrayList, Optional ByVal UsaNumerador As Boolean = True) As Boolean
        Dim Banco As New AcessaBanco
        Sqls.Clear()
        Me.SalvarSQL(Sqls)

        If Sqls.Count = 0 OrElse Banco.GravaBanco(Sqls) Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Sub SalvarSQL(ByRef Sqls As ArrayList, Optional ByVal UsaNumerador As Boolean = True)
        Dim i As Integer
        Dim NumeroTitulo As Integer
        Dim Num As Numerador = Nothing
        If Not UsaNumerador Then
            Num = New Numerador(1)
            NumeroTitulo = Num.Sequencia
            i = 1
        End If

        For Each Tit In Me.Where(Function(s) Not s.CodigoProvisao = 1)
            If Not Tit.IUD = Nothing Then
                'If Tit.UsarSequencia Then
                '    UsaNumerador = False
                'Else
                If Tit.IUD = "I" And Not UsaNumerador Then
                    Tit.Codigo = NumeroTitulo + i
                    i += 1
                End If
                'End If
                Tit.SalvarSql(Sqls, UsaNumerador)
                If Not NF Is Nothing AndAlso Not Tit.IUD = "U" AndAlso Tit.CodigoProvisao = eProvisao.Previsao Then
                    Sqls.Add(NotaxTituloSql(Tit.Codigo, Tit.IUD))
                End If
            End If
        Next

        If Not UsaNumerador And i > 0 Then Sqls.Add(Num.IncrementarNumeradorSql(True, i))
    End Sub

    Public Function NotaxTituloSql(ByRef CodigoTitulo As Integer, ByVal pIUD As String) As String
        Dim sql As String = ""
        Select Case pIUD
            Case "I"
                sql = "Insert into NotaFiscalXTitulo(Empresa_Id, EndEmpresa_Id, Cliente_Id, EndCliente_Id, EntradaSaida_Id, Serie_Id, Nota_Id, Titulo_Id)" & vbCrLf &
                      " values('" & NF.CodigoEmpresa & "'," & NF.EnderecoEmpresa & ", " & vbCrLf &
                      "'" & NF.Pedido.CodigoCliente & "'," & NF.Pedido.EnderecoCliente & "," & vbCrLf &
                      "'" & NF.EntradaSaida.ToString.Substring(0, 1) & "','" & NF.Serie & "'," & NF.Codigo & "," & CodigoTitulo & ")"
            Case "D"
                sql = " Delete NotaFiscalXTitulo" & vbCrLf &
                      "  Where Empresa_Id      ='" & NF.CodigoEmpresa & "'" & vbCrLf &
                      "    and EndEmpresa_Id   = " & NF.EnderecoEmpresa & vbCrLf &
                      "    and Cliente_Id      ='" & NF.Pedido.CodigoCliente & "'" & vbCrLf &
                      "    and EndCliente_Id   = " & NF.Pedido.EnderecoCliente & vbCrLf &
                      "    and EntradaSaida_Id ='" & NF.EntradaSaida.ToString.Substring(0, 1) & "'" & vbCrLf &
                      "    and Serie_Id        ='" & NF.Serie & "'" & vbCrLf &
                      "    and Nota_Id         = " & NF.Codigo & vbCrLf &
                      "    and Titulo_Id       = " & CodigoTitulo
        End Select

        Return sql
    End Function
#End Region

#Region "Methods Usados pela Nota Fiscal / Pedido / Fixacao"

    Public Sub ExcluirNota()

        Dim ValorNota As Decimal = 0

        For Each tit As Titulo In NF.VencimentosNota
            tit.IUD = "D"

            If NF.Pedido.Moeda.Classificacao = eTiposMoeda.MoedaEstrangeira Then
                ValorNota += tit.MoedaValorDoDocumento
            Else
                ValorNota += tit.ValorDoDocumento
            End If
        Next

        Dim ProvisionadosCount As Integer = NF.VencimentosPedido.NumeroDeTitulosProvisionados()

        If ProvisionadosCount > 0 Then
            AjustarProvisionadoExclusaoNF(ValorNota)
            Exit Sub
        End If

        Dim ValorTotalPedido As Decimal = NF.VencimentosPedido.ConsultaValorNaLista(RP_Pedido, _TipoMoeda)
        Dim ValorParaCriacaoDeUmNovoTitulo As Decimal = NF.VencimentosNota.ConsultaValorNaLista(RP_Pedido, _TipoMoeda, 2)
        Dim ValorTotalNota As Decimal = NF.TotalNotaValorModificado

        If ValorParaCriacaoDeUmNovoTitulo <> ValorTotalNota Then
            ValorParaCriacaoDeUmNovoTitulo = ValorParaCriacaoDeUmNovoTitulo - ValorTotalNota
            If ValorParaCriacaoDeUmNovoTitulo < 0 Then
                ValorParaCriacaoDeUmNovoTitulo = 0
            End If
        End If

        Dim n As Numerador = New Numerador(1)
        Dim NumeroTitulo As Integer = n.Sequencia + 1

        If ValorParaCriacaoDeUmNovoTitulo > 0 Then

            Dim Tit As New Titulo
            Tit.IUD = "I"
            Tit.Codigo = NumeroTitulo
            NumeroTitulo += 1
            Tit.Sequencia = 0
            Tit.CodigoProvisao = IIf(NF.SubOperacao.Devolucao, 2, 3)

            If RP_Pedido = "R" Then
                Tit.ReceberPagar = "R"
                Tit.ContaContabilCliente = NF.Itens(0).Produto.CarteiraVenda.CodigoContaCliente
                Tit.CodigoCarteira = NF.Itens(0).Produto.CodigoCarteiraVenda
            Else
                Tit.ReceberPagar = "P"
                Tit.ContaContabilCliente = NF.Itens(0).Produto.CarteiraCompra.CodigoContaCliente
                Tit.CodigoCarteira = NF.Itens(0).Produto.CodigoCarteiraCompra
            End If

            Tit.Tributo = ""
            Tit.CodigoEmpresaPedido = NF.Pedido.CodigoEmpresa
            Tit.EndEmpresaPedido = NF.Pedido.EnderecoEmpresa
            Tit.CodigoPedido = NF.CodigoPedido
            Tit.CodigoPedidoFixacao = 0
            Tit.CodigoProcuracao = NF.CodigoProcuracao
            Tit.Movimento = Date.Now.ToString("dd-MM-yyyy")
            Tit.DataMoeda = Date.Now.ToString("dd-MM-yyyy")
            Tit.CodigoIndexador = NF.Pedido.CodigoIndexador
            Tit.CodigoMoeda = NF.Pedido.CodigoMoeda
            Tit.CodigoTipoPgto = 1
            Tit.CodigoSituacao = 1
            Tit.Lote = 70

            Tit.Vencimento = Date.Now.ToString("dd-MM-yyyy")
            Tit.Prorrogacao = Date.Now.ToString("dd-MM-yyyy")

            Tit.Baixa = Date.Now.ToString("dd-MM-yyyy")
            Tit.CodigoUnidadeDeNegocio = NF.Pedido.CodigoUnidadeNegocio
            Tit.CodigoEmpresa = NF.CodigoEmpresa
            Tit.EnderecoEmpresa = NF.EnderecoEmpresa
            Tit.CodigoCliente = NF.Pedido.CodigoCliente
            Tit.EndCliente = NF.Pedido.EnderecoCliente
            Tit.CodigoBancoCliente = 0
            Tit.CodigoAgenciaCliente = ""
            Tit.DigitoAgenciaCliente = ""
            Tit.ContaCliente = ""
            Tit.DigitoContaCliente = ""
            Tit.Cheque = False
            Tit.Slips = False
            Tit.Recibo = False
            Tit.Aviso = False
            Tit.ReciboDeposito = False

            If _TipoMoeda = eTiposMoeda.Oficial Then
                Tit.ValorDoDocumento = ValorParaCriacaoDeUmNovoTitulo
            Else
                Tit.MoedaValorDoDocumento = ValorParaCriacaoDeUmNovoTitulo
            End If

            Tit.Historico = ""
            Tit.CodigoDeBarras = ""
            Tit.CodigoDigitado = False
            Tit.CodigoDeBarrasPreImpresso = False

            Tit.CodigoDestinatario = NF.Pedido.CodigoCliente
            Tit.EndDestinatario = NF.Pedido.EnderecoCliente
            Tit.NomeDoDestinatario = ""
            Tit.Destinacao = ""
            Tit.Solicitacao = 0

            Tit.Agrupado = eAgrupamentoFinanceiro.NaoAgrupado
            Tit.RegistroMestre = 0
            Tit.Observacoes = ""
            Tit.SituacaoBancaria = 0
            Tit.NumeroDoCheque = 0
            Tit.CodigoAdiantamento = 0
            Tit.TaxaAdto = 0

            Tit.UsuarioInclusao = NF.UsuarioInclusao
            Tit.UsuarioInclusaoData = Date.Now.ToString("dd-MM-yyyy")

            Tit.CodigoCarteiraDoTitulo = 0

            NF.VencimentosPedido.Add(Tit)

        End If

    End Sub

    Public Sub DevolucaoNota(ByRef Sqls As ArrayList)
        'Dim ValorNota As Decimal = IIf(_TipoMoeda = eTiposMoeda.Oficial, NF.TotalNota, Math.Round(NF.TotalNota / NF.IndiceNota, 2))
        'Dim ValorNota As Decimal = IIf(_TipoMoeda = eTiposMoeda.Oficial, NF.TotalNota, (From x In NF.Itens Select x.SaldoValorOficial / x.IndiceProdutoNota).Sum)
        Dim ValorNota As Decimal = 0
        If _TipoMoeda = eTiposMoeda.Oficial Then
            ValorNota = NF.TotalNota
        Else
            ValorNota = (From x In NF.Itens Select x.SaldoValorOficial / x.IndiceProdutoNota).Sum
        End If

        Dim valorAuxNF As Decimal = ValorNota
        Dim utilizaTitPrevisao As Boolean = False

        'recupera o titulo em provisao do pedido
        Dim TitProvisao As Titulo = NF.VencimentosPedido.Where(Function(s) s.CodigoProvisao = eProvisao.Provisao).OrderBy(Function(s) s.Prorrogacao).FirstOrDefault()
        'Pensar como fazer - furlan - 31/07/2015
        '1 - pegar todos os titulos em previsão(2) ref. a nota da lista do que tá sendo devolvido e order decrescente
        '2 - verificar se tem algum titulo em provisão(3)
        '3 - pegar o total da nota que está sendo emitida e ir diminuindo dos titulos, para os titulos que zerar mudar a situação para 3(excluido)

        'Procura na NF q está sendo devolvida.
        For Each Item As NotaFiscalXItem In NF.Itens
            For Each itemNotaDevolucao As NotaFiscalDevolucaoXNotaFiscal In Item.NotasDevolucao.Where(Function(s) s.ValorDevolucao > 0)
                For Each TitNotaDevolucao As Titulo In itemNotaDevolucao.Nota.VencimentosNota.Where(Function(s) s.CodigoProvisao = eProvisao.Previsao AndAlso s.CodigoSituacao = eSituacao.Normal).OrderBy(Function(s) s.Prorrogacao)
                    utilizaTitPrevisao = False
                    If TitProvisao Is Nothing Then
                        If ValorNota < TitNotaDevolucao.ValorDoDocumento Then
                            TitProvisao = New Titulo(TitNotaDevolucao.Codigo, TitNotaDevolucao.ReceberPagar)
                            TitProvisao.IUD = "I"
                        Else
                            TitProvisao = TitNotaDevolucao
                            TitProvisao.IUD = "U"
                            utilizaTitPrevisao = True
                            ValorNota -= TitNotaDevolucao.ValorDoDocumento
                        End If
                        TituloNew(TitProvisao, Sqls)
                    Else
                        TitProvisao.IUD = "U"
                    End If

                    If ValorNota >= TitNotaDevolucao.ValorDoDocumento Then
                        ValorNota -= TitNotaDevolucao.ValorDoDocumento
                        TitNotaDevolucao.IUD = "U"
                        If Not utilizaTitPrevisao Then TitNotaDevolucao.CodigoSituacao = eSituacao.Excluido
                    Else
                        TitNotaDevolucao.ValorDoDocumento -= ValorNota
                        TitNotaDevolucao.IUD = "U"
                        ValorNota = 0
                    End If
                    If Not utilizaTitPrevisao Then NF.VencimentosNota.Add(TitNotaDevolucao)
                    If ValorNota = 0 Then
                        If String.IsNullOrWhiteSpace(TitProvisao.Historico) Then
                            TitProvisao.Historico = TitProvisao.Historico & vbCrLf & "** Titulo de: " & (TitProvisao.ValorDoDocumento + valorAuxNF).ToString("N2") &
                           " Gerado pela NF de devolução Nº " & NF.Codigo & "-" & NF.Serie
                            'Else
                            '    TitProvisao.Historico = TitProvisao.Historico & vbCrLf & "** Aumento de capital do titulo de: " & TitProvisao.ValorDoDocumento.ToString("N2") & " para: " & (TitProvisao.ValorDoDocumento + valorAuxNF).ToString("N2") & _
                            '   " devido a NF de devolução Nº " & NF.Codigo & "-" & NF.Serie
                        End If
                        TitProvisao.ValorDoDocumento += valorAuxNF
                        NF.VencimentosNota.Add(TitProvisao)
                        Exit For
                    End If
                Next
            Next
        Next
        'procura na lista de titulos do pedido.
        If ValorNota > 0 Then
            For Each Tit As Titulo In NF.VencimentosPedido.Where(Function(s) s.CodigoProvisao = eProvisao.Previsao AndAlso s.CodigoSituacao = eSituacao.Normal).OrderBy(Function(s) s.Prorrogacao)
                utilizaTitPrevisao = False
                If TitProvisao Is Nothing Then
                    If ValorNota < Tit.ValorDoDocumento Then
                        TitProvisao = New Titulo(Tit.Codigo, Tit.ReceberPagar)
                        TitProvisao.IUD = "I"
                    Else
                        TitProvisao = Tit
                        TitProvisao.IUD = "U"
                        ValorNota -= Tit.ValorDoDocumento
                    End If
                    TituloNew(TitProvisao, Sqls)
                Else
                    TitProvisao.IUD = "U"
                End If

                If ValorNota >= Tit.ValorDoDocumento Then
                    ValorNota -= Tit.ValorDoDocumento
                    Tit.IUD = "U"
                    If Not utilizaTitPrevisao Then Tit.CodigoSituacao = eSituacao.Excluido
                Else
                    Tit.ValorDoDocumento -= ValorNota
                    Tit.IUD = "U"
                    ValorNota = 0
                End If
                If Not utilizaTitPrevisao Then NF.VencimentosNota.Add(Tit)
                If ValorNota = 0 Then
                    If String.IsNullOrWhiteSpace(TitProvisao.Historico) Then
                        TitProvisao.Historico = TitProvisao.Historico & vbCrLf & "** Titulo de: " & (TitProvisao.ValorDoDocumento + valorAuxNF).ToString("N2") &
                       " Gerado pela NF de devolução Nº " & NF.Codigo & "-" & NF.Serie
                        'Else
                        '    TitProvisao.Historico = TitProvisao.Historico & vbCrLf & "** Aumento de capital do titulo de: " & TitProvisao.ValorDoDocumento.ToString("N2") & " para: " & (TitProvisao.ValorDoDocumento + valorAuxNF).ToString("N2") & _
                        '   " devido a NF de devolução Nº " & NF.Codigo & "-" & NF.Serie
                    End If

                    TitProvisao.ValorDoDocumento += valorAuxNF
                    NF.VencimentosNota.Add(TitProvisao)
                    Exit For
                End If
            Next
        End If
    End Sub

    Private Sub TituloNew(ByVal Tit As Titulo, ByRef Sqls As ArrayList)

        If Tit.IUD = "I" Then
            Dim n As Numerador = New Numerador(1)
            Dim NumeroTitulo As Integer = n.Sequencia + 1
            Tit.Codigo = NumeroTitulo
            Sqls.Add(n.IncrementarNumeradorSql(True, 1))
        End If
        Tit.Sequencia = 0
        Tit.CodigoProvisao = eProvisao.Provisao
        Tit.ValorDoDocumento = 0

        If RP_Pedido = "R" Then
            Tit.ReceberPagar = "R"
            Tit.ContaContabilCliente = NF.Itens(0).Produto.CarteiraVenda.CodigoContaCliente
            Tit.CodigoCarteira = NF.Itens(0).Produto.CodigoCarteiraVenda
        Else
            Tit.ReceberPagar = "P"
            Tit.ContaContabilCliente = NF.Itens(0).Produto.CarteiraCompra.CodigoContaCliente
            Tit.CodigoCarteira = NF.Itens(0).Produto.CodigoCarteiraCompra
        End If

        Tit.Tributo = ""

        Tit.DataMoeda = NF.Movimento.ToString("dd-MM-yyyy")
        Tit.CodigoIndexador = NF.Pedido.CodigoIndexador
        Tit.CodigoMoeda = NF.Pedido.CodigoMoeda
        Tit.Movimento = NF.Movimento.ToString("dd-MM-yyyy")
        Tit.Vencimento = NF.Movimento.ToString("dd-MM-yyyy")
        Tit.Prorrogacao = NF.Movimento.ToString("dd-MM-yyyy")

        Tit.CodigoTipoPgto = 1
        Tit.CodigoSituacao = 1
        Tit.Lote = 70

        Tit.Baixa = NF.Movimento.ToString("dd-MM-yyyy")
        Tit.CodigoUnidadeDeNegocio = NF.Pedido.CodigoUnidadeNegocio
        Tit.CodigoEmpresa = NF.CodigoEmpresa
        Tit.EnderecoEmpresa = NF.EnderecoEmpresa
        Tit.CodigoCliente = NF.Pedido.CodigoCliente
        Tit.EndCliente = NF.Pedido.EnderecoCliente
        Tit.CodigoBancoCliente = 0
        Tit.CodigoAgenciaCliente = ""
        Tit.DigitoAgenciaCliente = ""
        Tit.ContaCliente = ""
        Tit.DigitoContaCliente = ""
        Tit.Cheque = False
        Tit.Slips = False
        Tit.Recibo = False
        Tit.Aviso = False
        Tit.ReciboDeposito = False
        Tit.CodigoEmpresaPedido = NF.Pedido.CodigoEmpresa
        Tit.EndEmpresaPedido = NF.Pedido.EnderecoEmpresa
        Tit.CodigoPedido = NF.CodigoPedido
        Tit.CodigoPedidoFixacao = 0
        Tit.CodigoProcuracao = NF.CodigoProcuracao
        Tit.Historico = String.Empty
        Tit.CodigoDeBarras = ""
        Tit.CodigoDigitado = False
        Tit.CodigoDeBarrasPreImpresso = False

        Tit.CodigoDestinatario = NF.Pedido.CodigoCliente
        Tit.EndDestinatario = NF.Pedido.EnderecoCliente
        Tit.NomeDoDestinatario = ""
        Tit.Destinacao = ""
        Tit.Solicitacao = 0

        Tit.Agrupado = eAgrupamentoFinanceiro.NaoAgrupado
        Tit.RegistroMestre = 0
        Tit.Observacoes = ""
        Tit.SituacaoBancaria = 0
        Tit.NumeroDoCheque = 0
        Tit.CodigoAdiantamento = 0
        Tit.TaxaAdto = 0

        Tit.UsuarioInclusao = NF.UsuarioInclusao
        Tit.UsuarioInclusaoData = NF.DataInclusao

        Tit.CodigoCarteiraDoTitulo = 0

    End Sub

    Public Sub ParcelarNota(ByRef FormaPagamento As Integer)
        Dim Banco As New AcessaBanco
        Dim ds As New DataSet

        Dim bDataFixa As Boolean = False
        Dim dataFixa As New DateTime

        Dim Sql As String = ""
        Dim SqlA As String = ""


        Dim ValorProvisionado As Decimal

        'TRATANDO COMPLEMENTO EM DÓLAR PARA RTGRÃOS - FURLAN 04/11/2024
        If (Left(NF.CodigoEmpresa, 8) = "24450490" OrElse Left(NF.CodigoEmpresa, 8) = "44979506") AndAlso NF.SubOperacao.QuantidadeFiscal = False AndAlso NF.SubOperacao.FinalidadeDaNota = 2 AndAlso _TipoMoeda = eTiposMoeda.MoedaEstrangeira Then
            ValorProvisionado = NF.TotalNota
        Else
            ValorProvisionado = NF.VencimentosPedido.ConsultaValorNaLista(RP_Pedido, TipoMoeda, 3)
        End If

        Dim IPIICMSST As Decimal
        Dim ValorAParcelar As Decimal
        Dim ValorNota As Decimal

        'Alterada a validação para contemplar notas de complemento de valor em pedidos que sejam em dólar
        'Ao buscar o IndiceNota é feita uma verificação se a NF têm quantidade fiscal para que tenha índice senão será zero
        'If _TipoMoeda = eTiposMoeda.MoedaEstrangeira AndAlso NF.IndiceNota > 0 Then


        If NF.Pedido.SubOperacao.Classe = eClassesOperacoes.AFIXAR AndAlso NF.SubOperacao.Classe = eClassesOperacoes.COMPLEMENTACOES Then
            ValorNota = NF.Pedido.Itens(0).Fixacoes.Where(Function(s) s.Codigo = NF.Itens(0).CodigoFixacao).FirstOrDefault.TotalOficial
        ElseIf (Left(NF.CodigoEmpresa, 8) = "24450490" OrElse Left(NF.CodigoEmpresa, 8) = "44979506") AndAlso NF.SubOperacao.QuantidadeFiscal = False AndAlso NF.SubOperacao.FinalidadeDaNota = 2 AndAlso _TipoMoeda = eTiposMoeda.MoedaEstrangeira Then
            'TRATANDO COMPLEMENTO EM DÓLAR PARA RTGRÃOS - FURLAN 04/11/2024
            'LIBERADO PARA VERDE - FURLAN 02/05/2025
            ValorNota = NF.TotalNota
        ElseIf _TipoMoeda = eTiposMoeda.MoedaEstrangeira Then
            ValorNota = (From x In NF.Itens Select x.ValorLiquidoMoeda).Sum
        Else
            ValorNota = NF.TotalNota
        End If

        If NF.VencimentosNota.Count > 0 Then
            ValorAParcelar = ValorNota
        ElseIf ValorProvisionado < ValorNota Then
            ValorAParcelar = ValorProvisionado
        Else
            ValorAParcelar = ValorNota
        End If

        If NF.VencimentosNota.Count = 0 Then NF.VencimentosNota.AjustarTitulosProvisionados(ValorAParcelar)

        If ValorAParcelar <= 0 Then Exit Sub

        For Each item As NotaFiscalXItem In NF.Itens
            For Each Enc As NotaFiscalXItemXEncargo In item.Encargos
                'If Enc.Codigo = "IPI" Then IPIICMSST += IIf(_TipoMoeda = eTiposMoeda.Oficial, Enc.Valor, Enc.Valor / item.IndiceProdutoNota)
                'If Enc.Codigo = "ICMS-ST" Then IPIICMSST += IIf(_TipoMoeda = eTiposMoeda.Oficial, Enc.Valor, Enc.Valor / item.IndiceProdutoNota)

                If Left(NF.CodigoEmpresa, 8) = "40938762" Then
                    Continue For
                Else
                    If Enc.Codigo = "IPI" Then
                        If _TipoMoeda = eTiposMoeda.Oficial Then
                            IPIICMSST += Enc.Valor
                        Else
                            IPIICMSST += Enc.Valor / item.IndiceProdutoNota
                        End If
                    End If

                    If Enc.Codigo = "ICMS-ST" Then
                        If _TipoMoeda = eTiposMoeda.Oficial Then
                            IPIICMSST += Enc.Valor
                        Else
                            IPIICMSST += Enc.Valor / item.IndiceProdutoNota
                        End If
                    End If
                End If
            Next
        Next

        If NF.TotalNota = 0 Then
            _msg = "Valor da Nota está zerado"
            Exit Sub
        End If

        Sql &= "	DECLARE " & vbCrLf &
               "	@Diferenca numeric(18,2)," & vbCrLf &
               "	@ValorIPI numeric(18,2)," & vbCrLf &
               "	@ValorTotal numeric(18,2)," & vbCrLf &
               "	@Data varchar(10)," & vbCrLf &
               "	@FPagto int" & vbCrLf

        '--Informa o valor do IPI a ser cobrado"
        '-- Na Compra dilui em todas as parcelas nas outras cobra na primeira parcela
        If NF.SubOperacao.Classe = eClassesOperacoes.COMPRAS Then
            Sql &= "	set @ValorIPI   =  0 " & vbCrLf &
                   "	set @ValorTotal =    " & Str(ValorAParcelar)
        Else
            Sql &= "	set @ValorIPI   =  " & Str(IPIICMSST) & vbCrLf &
                   "	set @ValorTotal =  " & Str(ValorAParcelar - IPIICMSST) & vbCrLf
        End If


        SqlA = "SELECT Pagamento_Id, Sequencia_Id, isnull(Dias,0) as Dias" & vbCrLf &
                "FROM  PagamentosXParcelas" & vbCrLf &
                "WHERE Pagamento_Id = " & FormaPagamento

        ds = Banco.ConsultaDataSet(SqlA, "VencimentosNF")

        bDataFixa = NF.Pedido.DataFixa(dataFixa)

        If (Left(NF.CodigoEmpresa, 8) = "05366261" _
            OrElse Left(NF.CodigoEmpresa, 8) = "38198213" _
            OrElse Left(NF.CodigoEmpresa, 8) = "40938762" _
            OrElse Left(NF.CodigoEmpresa, 8) = "44005444" _
            OrElse Left(NF.CodigoEmpresa, 8) = "44979506" _
            OrElse Left(NF.CodigoEmpresa, 8) = "24450490" _
            OrElse Left(NF.CodigoEmpresa, 8) = "62747840" _
            OrElse Left(NF.CodigoEmpresa, 8) = "62780383" _
            OrElse Left(NF.CodigoEmpresa, 8) = "63358210" _
            OrElse Left(NF.CodigoEmpresa, 8) = "48984539") AndAlso
            (ds.Tables("VencimentosNF").Rows.Count > 1 AndAlso ds.Tables("VencimentosNF").Rows(1).Item(2) = 0) Then
            Sql &= "select pgto.Pagamento_Id, pgto.Descricao, 0 as Sequencia_Id, pgto.Parcelas, 0 as Dias, cr.Prorrogacao AS Vencimento," & vbCrLf &
                    "	   round(@ValorTotal / pgto.Parcelas, 2) as ValorParcela, @ValorTotal as ValorTotal" & vbCrLf &
                    "INTO #Temp1 " & vbCrLf

            If NF.EntradaSaida = eEntradaSaida.Entrada Then
                Sql &= "from ContasAPagar cr" & vbCrLf
            Else
                Sql &= "from ContasAReceber cr" & vbCrLf
            End If

            Sql &= "		inner Join Pedidos p" & vbCrLf &
                    "				on p.Empresa_Id     = cr.EmpresaPedido" & vbCrLf &
                    "				and p.EndEmpresa_Id = cr.EndEmpresaPedido" & vbCrLf &
                    "				and p.Pedido_Id     = cr.Pedido" & vbCrLf &
                    "		inner join Pagamentos pgto" & vbCrLf &
                    "				on pgto.Pagamento_Id = " & FormaPagamento.ToString & vbCrLf &
                    "where cr.EmpresaPedido  = '" & NF.CodigoEmpresa & "'" & vbCrLf &
                    "AND cr.EndEmpresaPedido = " & NF.EnderecoEmpresa & vbCrLf &
                    "AND cr.Pedido           = " & NF.CodigoPedido & vbCrLf &
                    "AND cr.Provisao         = 3" & vbCrLf

        Else
            If Left(NF.CodigoEmpresa, 8) = "15204808" Then
                Sql &= "	set @Data       = '" & NF.Pedido.Vencimentos.Where(Function(s) s.Provisao = 3).OrderBy(Function(s) s.Vencimento).First.DataProrrogacao.ToString("yyyy-MM-dd") & "'" & vbCrLf &
                       "	set @FPagto     =  " & FormaPagamento.ToString & vbCrLf
            ElseIf Left(NF.CodigoEmpresa, 8) = "05366261" _
                OrElse Left(NF.CodigoEmpresa, 8) = "38198213" _
                OrElse Left(NF.CodigoEmpresa, 8) = "40938762" _
                OrElse Left(NF.CodigoEmpresa, 8) = "44005444" _
                OrElse Left(NF.CodigoEmpresa, 8) = "44979506" _
                OrElse Left(NF.CodigoEmpresa, 8) = "24450490" _
                OrElse Left(NF.CodigoEmpresa, 8) = "62747840" _
                OrElse Left(NF.CodigoEmpresa, 8) = "62780383" _
                OrElse Left(NF.CodigoEmpresa, 8) = "63358210" _
                OrElse Left(NF.CodigoEmpresa, 8) = "48984539" Then
                If ds.Tables("VencimentosNF").Rows.Count = 1 AndAlso ds.Tables("VencimentosNF").Rows(0).Item(2) = 0 Then
                    Sql &= "	set @Data       = '" & NF.Pedido.Vencimentos.Where(Function(s) s.Provisao = 3).OrderBy(Function(s) s.Vencimento).First.DataProrrogacao.ToString("yyyy-MM-dd") & "'" & vbCrLf &
                           "	set @FPagto     =  " & FormaPagamento.ToString & vbCrLf
                Else
                    Sql &= "	set @Data       = '" & NF.Movimento.ToString("yyyy-MM-dd") & "'" & vbCrLf &
                            "	set @FPagto     =  " & FormaPagamento.ToString & vbCrLf
                End If
            Else
                Sql &= "	set @Data       = '" & NF.Movimento.ToString("yyyy-MM-dd") & "'" & vbCrLf &
                       "	set @FPagto     =  " & FormaPagamento.ToString & vbCrLf
            End If

            '--Seleciona no banco a forma de pagamento e as Parcelas"
            Sql &= "	SELECT Pagamentos.Pagamento_Id," & vbCrLf &
                   "		   Pagamentos.Descricao," & vbCrLf &
                   "		   PagamentosXParcelas.Sequencia_Id, " & vbCrLf &
                   "		   Pagamentos.Parcelas," & vbCrLf &
                   "		   PagamentosXParcelas.Dias," & vbCrLf
            '--Soma a Data Base o Numero de dias referente ao numero da Parcela"
            Sql &= "		   DATEADD(DAY, CONVERT(INT, PagamentosXParcelas.Dias), CONVERT(DATETIME, @Data)) as Vencimento," & vbCrLf
            '--Divide o valor total pelo numero de parcelas p/ descubrir o valor da parcela"
            Sql &= "		   round(@ValorTotal / Pagamentos.Parcelas, 2) as ValorParcela," & vbCrLf
            '-- Armazena o Valor Total para calcular a diferenca na divisao das parcelas"
            Sql &= "		   @ValorTotal as ValorTotal" & vbCrLf &
                   "	  INTO #Temp1 " & vbCrLf &
                   "	  FROM Pagamentos " & vbCrLf &
                   "	 INNER JOIN PagamentosXParcelas " & vbCrLf &
                   "		ON Pagamentos.Pagamento_Id = PagamentosXParcelas.Pagamento_Id " & vbCrLf &
                   "	 where Pagamentos.Pagamento_Id = @FPagto " & vbCrLf &
                   "	 order by PagamentosXParcelas.Sequencia_Id " & vbCrLf
        End If

        '	--Calcula o valor Parcelado para ver se ha diferenca com o valor total caso haja armazena o valor na variavel "
        Sql &= "	set @Diferenca = (Select top(1) ValorTotal - (ValorParcela * Parcelas) from #Temp1)" & vbCrLf


        '	--Atualiza o valor da primeira parcela acrescida do IPI + a Diferenca"
        '       "	where #temp1.Sequencia_Id = 1" & vbCrLf - Ajustei pois não estava fazendo o arredondamento correto - Furlan - 07-06-2022
        '"	where #temp1.Vencimento = (SELECT MAX(#temp1.Vencimento) FROM #temp1) " & vbCrLf
        Sql &= "	update #temp1 set" & vbCrLf &
               "	ValorParcela = ValorParcela + @ValorIPI + @Diferenca" & vbCrLf &
               "	where #temp1.Sequencia_Id = 1 " & vbCrLf

        Sql &= "	select Pagamento_id, Descricao, Sequencia_id, Parcelas, Dias, Vencimento, ValorParcela, ValorTotal + @ValorIPI as ValorTotal from #temp1"


        ds = Banco.ConsultaDataSet(Sql, "Titulos")

        Dim n As Numerador = New Numerador(1)
        Dim NumeroTitulo As Integer = n.Sequencia + 1
        Dim parcela As Integer

        Dim dValorParcela As Decimal
        Dim dDiferencaAParcelar As Decimal

        Dim dataParcelaAtual = Today()

        NF.VencimentosNota.Clear()

        For Each Row As DataRow In ds.Tables("Titulos").Rows
            Dim Tit As New Titulo

            If NF.Pedido.Moeda.Classificacao = eTiposMoeda.MoedaEstrangeira Then
                If NF.Pedido.CodigoIndexador = 99 OrElse NF.Pedido.IndexadorFixo Then
                    Tit.IndiceTitulo = NF.Pedido.IndiceFixado
                Else
                    Tit.IndiceTitulo = NF.Itens.ValorLiquido_Oficial / ValorAParcelar
                End If

                Tit.IndiceFixo = True
            Else
                Tit.IndiceFixo = False
            End If


            Tit.IUD = "I"
            Tit.Codigo = NumeroTitulo
            NumeroTitulo += 1
            parcela += 1
            Tit.Sequencia = 0
            Tit.CodigoProvisao = 2

            If NF.EntradaSaida = eEntradaSaida.Entrada Then
                Tit.ReceberPagar = "P"
                If NF.SubOperacao.Devolucao Then
                    Tit.ContaContabilCliente = NF.Itens(0).Produto.CarteiraVenda.CodigoContaCliente
                    Tit.CodigoCarteira = NF.Itens(0).Produto.CodigoCarteiraVenda
                Else
                    Tit.ContaContabilCliente = NF.Itens(0).Produto.CarteiraCompra.CodigoContaCliente
                    Tit.CodigoCarteira = NF.Itens(0).Produto.CodigoCarteiraCompra
                End If
            Else
                Tit.ReceberPagar = "R"
                If NF.SubOperacao.Devolucao Then
                    Tit.ContaContabilCliente = NF.Itens(0).Produto.CarteiraCompra.CodigoContaCliente
                    Tit.CodigoCarteira = NF.Itens(0).Produto.CodigoCarteiraCompra
                Else
                    Tit.ContaContabilCliente = NF.Itens(0).Produto.CarteiraVenda.CodigoContaCliente
                    Tit.CodigoCarteira = NF.Itens(0).Produto.CodigoCarteiraVenda
                End If
            End If

            Tit.Tributo = ""

            Tit.CodigoEmpresaPedido = NF.Pedido.CodigoEmpresa
            Tit.EndEmpresaPedido = NF.Pedido.EnderecoEmpresa
            Tit.CodigoPedido = NF.CodigoPedido
            Tit.CodigoPedidoFixacao = 0
            Tit.CodigoProcuracao = NF.CodigoProcuracao
            Tit.DataMoeda = NF.Movimento.ToString("dd-MM-yyyy")
            Tit.CodigoIndexador = NF.Pedido.CodigoIndexador
            Tit.CodigoMoeda = NF.Pedido.CodigoMoeda
            Tit.Movimento = NF.Movimento.ToString("dd-MM-yyyy")

            If bDataFixa Then
                Tit.Vencimento = Funcoes.ValidaDataUtil(NF.Empresa.Codigo, NF.Empresa.CodigoEndereco, dataFixa)
                Tit.Prorrogacao = Funcoes.ValidaDataUtil(NF.Empresa.Codigo, NF.Empresa.CodigoEndereco, dataFixa)
            ElseIf CDate(Row("Vencimento")) < dataParcelaAtual AndAlso NF.EntradaSaida = eEntradaSaida.Saida Then
                Tit.Vencimento = Funcoes.ValidaDataUtil(NF.Empresa.Codigo, NF.Empresa.CodigoEndereco, dataParcelaAtual)
                Tit.Prorrogacao = Funcoes.ValidaDataUtil(NF.Empresa.Codigo, NF.Empresa.CodigoEndereco, dataParcelaAtual)
            Else
                Tit.Vencimento = Funcoes.ValidaDataUtil(NF.Empresa.Codigo, NF.Empresa.CodigoEndereco, CDate(Row("Vencimento")))
                Tit.Prorrogacao = Funcoes.ValidaDataUtil(NF.Empresa.Codigo, NF.Empresa.CodigoEndereco, CDate(Row("Vencimento")))
            End If

            Tit.Baixa = Funcoes.ValidaDataUtil(NF.Empresa.Codigo, NF.Empresa.CodigoEndereco, CDate(Row("Vencimento")))

            Tit.CodigoTipoPgto = 1
            Tit.CodigoSituacao = 1
            Tit.Lote = 70
            Tit.CodigoUnidadeDeNegocio = NF.Pedido.CodigoUnidadeNegocio
            Tit.CodigoEmpresa = NF.CodigoEmpresa
            Tit.EnderecoEmpresa = NF.EnderecoEmpresa
            Tit.CodigoCliente = NF.Pedido.CodigoCliente
            Tit.EndCliente = NF.Pedido.EnderecoCliente
            Tit.CodigoBancoCliente = 0
            Tit.CodigoAgenciaCliente = ""
            Tit.DigitoAgenciaCliente = ""
            Tit.ContaCliente = ""
            Tit.DigitoContaCliente = ""
            Tit.Cheque = False
            Tit.Slips = False
            Tit.Recibo = False
            Tit.Aviso = False
            Tit.ReciboDeposito = False

            'Ajuste de Centavos
            dValorParcela += Row("ValorParcela")

            If parcela = ds.Tables("Titulos").Rows.Count() Then
                dDiferencaAParcelar = ValorAParcelar - dValorParcela
            End If

            'TRATANDO COMPLEMENTO EM DÓLAR PARA RTGRÃOS - FURLAN 04/11/2024
            If (Left(NF.CodigoEmpresa, 8) = "24450490" OrElse Left(NF.CodigoEmpresa, 8) = "44979506") AndAlso NF.SubOperacao.QuantidadeFiscal = 0 AndAlso NF.SubOperacao.FinalidadeDaNota = 2 AndAlso _TipoMoeda = eTiposMoeda.MoedaEstrangeira Then
                Tit.Moeda.Classificacao = eTiposMoeda.Oficial
                Tit.IndiceTitulo = 0
                Tit.ValorDoDocumento = dValorParcela
                Tit.Moeda.Classificacao = eTiposMoeda.MoedaEstrangeira
            ElseIf _TipoMoeda = eTiposMoeda.Oficial Then
                Tit.ValorDoDocumento = Row("ValorParcela") + dDiferencaAParcelar
            Else
                Tit.MoedaValorDoDocumento = Row("ValorParcela") + dDiferencaAParcelar
            End If

            If _TipoMoeda = eTiposMoeda.MoedaEstrangeira Then
                If ((Tit.ValorDoDocumento - NF.TotalNota) < 0.5) OrElse ((NF.TotalNota - Tit.ValorDoDocumento) < 0.5) Then
                    Tit.ValorDoDocumento = NF.TotalNota
                End If
            End If

            Tit.Historico = "REF. NF " & NF.Codigo & "-" & NF.Serie & "-" & IIf(NF.EntradaSaida = eEntradaSaida.Saida, "S", "E") & ", Parcela " & parcela & "/" & Row("Parcelas") & ", Pedido: " & NF.CodigoPedido & " / " & NF.Cliente.Nome
            Tit.CodigoDeBarras = ""
            Tit.CodigoDigitado = False
            Tit.CodigoDeBarrasPreImpresso = False

            Tit.CodigoDestinatario = NF.Pedido.CodigoCliente
            Tit.EndDestinatario = NF.Pedido.EnderecoCliente
            Tit.NomeDoDestinatario = ""
            Tit.Destinacao = ""
            Tit.Solicitacao = 0

            Tit.Agrupado = eAgrupamentoFinanceiro.NaoAgrupado
            Tit.RegistroMestre = 0
            Tit.Observacoes = ""
            Tit.SituacaoBancaria = 0
            Tit.NumeroDoCheque = 0
            Tit.CodigoAdiantamento = 0
            Tit.TaxaAdto = 0

            Tit.UsuarioInclusao = NF.UsuarioInclusao
            Tit.UsuarioInclusaoData = NF.DataInclusao

            Tit.CodigoCarteiraDoTitulo = 0

            Me.Add(Tit)
        Next
    End Sub

    Public Sub ParcelarNotasFiscaisGerais(ByRef FormaPagamento As Integer, pValorDaNota As Decimal, pValorParcelado As Decimal, pValorPago As Decimal, Optional pCodigoProvisao As Integer = 3)
        'Dim RetornoFunction As New ArrayList
        'RetornoFunction.Add(True)
        'RetornoFunction.Add("")
        'Posicao 0 - True/False
        'Posicao 1 - Mensagem Sucesso/Erro

        Dim Banco As New AcessaBanco
        Dim ds As New DataSet
        Dim Sql As String = ""


        '*************************************************
        If NF.IUD = "I" Then
            NF.VencimentosNota.Clear()
        Else
            For Each row In NF.VencimentosNota.Where(Function(s) s.CodigoProvisao <> 1 AndAlso s.IUD = "I").ToList()
                NF.VencimentosNota.Remove(row)
            Next
        End If

        Dim lstVencimentos = NF.VencimentosNota.Where(Function(s) s.CodigoProvisao = 1 And s.CodigoSituacao = 1).ToList()
        Dim parcela As Integer = lstVencimentos.Count
        '*************************************************

        Dim IPIICMSST As Decimal
        Dim ValorAParcelar As Decimal = pValorDaNota - pValorPago

        'If ValorAParcelar <= 0 Then
        '    RetornoFunction(1) = "Não existe valor para parcelamento"
        'End If

        For Each item As NotaFiscalXItem In NF.Itens
            For Each Enc As NotaFiscalXItemXEncargo In item.Encargos
                'If Enc.Codigo = "IPI" Then IPIICMSST += IIf(_TipoMoeda = eTiposMoeda.Oficial, Enc.Valor, Enc.Valor / item.IndiceProdutoNota)
                'If Enc.Codigo = "ICMS-ST" Then IPIICMSST += IIf(_TipoMoeda = eTiposMoeda.Oficial, Enc.Valor, Enc.Valor / item.IndiceProdutoNota)
                If NF.NFG OrElse (NF.Pedido IsNot Nothing AndAlso NF.Pedido.Moeda.Classificacao = eTiposMoeda.Oficial) Then
                    If Enc.Codigo = "IPI" Then
                        IPIICMSST += Enc.Valor
                    End If

                    If Enc.Codigo = "ICMS-ST" Then
                        IPIICMSST += Enc.Valor
                    End If
                Else
                    If Enc.Codigo = "IPI" Then
                        IPIICMSST += Enc.Valor / item.IndiceProdutoNota
                    End If

                    If Enc.Codigo = "ICMS-ST" Then
                        IPIICMSST += Enc.Valor / item.IndiceProdutoNota
                    End If
                End If
            Next
        Next

        Sql &= "	DECLARE " & vbCrLf &
               "	@Diferenca numeric(18,2)," & vbCrLf &
               "	@ValorIPI numeric(18,2)," & vbCrLf &
               "	@ValorTotal numeric(18,2)," & vbCrLf &
               "	@Data varchar(10)," & vbCrLf &
               "	@FPagto int" & vbCrLf

        '--Informa o valor do IPI a ser cobrado"
        '-- Na Compra dilui em todas as parcelas nas outras cobra na primeira parcela
        If NF.SubOperacao.Classe = eClassesOperacoes.COMPRAS Then
            Sql &= "	set @ValorIPI   =  0 " & vbCrLf &
                   "	set @ValorTotal =    " & Str(ValorAParcelar)
        Else
            Sql &= "	set @ValorIPI   =  " & Str(IPIICMSST) & vbCrLf &
                   "	set @ValorTotal =  " & Str(ValorAParcelar - IPIICMSST) & vbCrLf
        End If

        Sql &= "	set @Data       = '" & NF.Movimento.ToString("yyyy-MM-dd") & "'" & vbCrLf &
               "	set @FPagto     =  " & FormaPagamento.ToString & vbCrLf

        '--Seleciona no banco a forma de pagamento e as Parcelas"
        Sql &= "	SELECT Pagamentos.Pagamento_Id," & vbCrLf &
               "		   Pagamentos.Descricao," & vbCrLf &
               "		   PagamentosXParcelas.Sequencia_Id, " & vbCrLf &
               "		   Pagamentos.Parcelas - " & lstVencimentos.Count & " as Parcelas," & vbCrLf &
               "		   PagamentosXParcelas.Dias," & vbCrLf
        '--Soma a Data Base o Numero de dias referente ao numero da Parcela"
        Sql &= "		   convert(datetime,(@Data)) + PagamentosXParcelas.Dias as Vencimento," & vbCrLf
        '--Divide o valor total pelo numero de parcelas p/ descubrir o valor da parcela"
        Sql &= "		   round(@ValorTotal / (Pagamentos.Parcelas -" & lstVencimentos.Count & "), 2) as ValorParcela," & vbCrLf
        '-- Armazena o Valor Total para calcular a diferenca na divisao das parcelas"
        Sql &= "		   @ValorTotal as ValorTotal" & vbCrLf &
               "	  INTO #Temp1 " & vbCrLf &
               "	  FROM Pagamentos " & vbCrLf &
               "	 INNER JOIN PagamentosXParcelas " & vbCrLf &
               "		ON Pagamentos.Pagamento_Id = PagamentosXParcelas.Pagamento_Id " & vbCrLf &
               "       and PagamentosXParcelas.Sequencia_Id <= Pagamentos.Parcelas - " & lstVencimentos.Count & vbCrLf &
               "	 where Pagamentos.Pagamento_Id = @FPagto " & vbCrLf &
               "	 order by PagamentosXParcelas.Sequencia_Id " & vbCrLf

        '	--Calcula o valor Parcelado para ver se ha diferenca com o valor total caso haja armazena o valor na variavel "
        Sql &= "	set @Diferenca = (Select top(1) ValorTotal - (ValorParcela * Parcelas) from #Temp1)" & vbCrLf


        '	--Atualiza o valor da primeira parcela acrescida do IPI + a Diferenca"
        Sql &= "	update #temp1 set" & vbCrLf &
               "	ValorParcela = ValorParcela + @ValorIPI + @Diferenca" & vbCrLf &
               "	where #temp1.Sequencia_Id = 1" & vbCrLf &
               "	select Pagamento_id, Descricao, Sequencia_id, Parcelas, Dias, Vencimento, ValorParcela, ValorTotal + @ValorIPI as ValorTotal from #temp1" & vbCrLf

        ds = Banco.ConsultaDataSet(Sql, "Titulos")


        Dim n As Numerador = New Numerador(1)
        Dim NumeroTitulo As Integer = n.Sequencia + 1

        lstVencimentos = NF.VencimentosNota.Where(Function(s) s.CodigoProvisao <> 1 And s.CodigoSituacao = 1).ToList()

        Dim i As Integer = 0
        If lstVencimentos.Count > ds.Tables(0).Rows.Count Then
            For Each row In Me.Where(Function(s) s.IUD <> "D" And s.CodigoProvisao <> 1)
                If i < ds.Tables(0).Rows.Count Then
                    row.IUD = "U"
                    PreencheTitulo(ds.Tables(0).Rows(i), pCodigoProvisao, row, row.Codigo)
                Else
                    row.IUD = "D"
                End If
                i += 1
            Next
        Else
            For Each row In ds.Tables(0).Rows
                If lstVencimentos.Count > i Then
                    lstVencimentos(i).IUD = "U"
                    lstVencimentos(i).ValorDoDocumento = row("ValorParcela")
                    'PreencheTitulo(row, pCodigoProvisao, lstVencimentos(i), lstVencimentos(i).Codigo)
                Else
                    Dim tit As New Titulo()
                    tit.IUD = "I"
                    PreencheTitulo(row, pCodigoProvisao, tit, NumeroTitulo)
                    NumeroTitulo += 1
                    Me.Add(tit)
                End If
                i += 1
            Next
        End If
    End Sub

    Private Sub PreencheTitulo(row As DataRow, pCodigoProvisao As String, tit As Titulo, NumeroTitulo As Integer)
        tit.Codigo = NumeroTitulo
        tit.Sequencia = 0
        tit.CodigoProvisao = pCodigoProvisao
        tit.Historico = NF.Observacoes

        If NF.EntradaSaida = eEntradaSaida.Entrada Then
            tit.ReceberPagar = "P"
            If NF.SubOperacao.Devolucao Then
                tit.ContaContabilCliente = NF.Itens(0).Produto.CarteiraVenda.CodigoContaCliente
                tit.CodigoCarteira = NF.Itens(0).Produto.CodigoCarteiraVenda
            Else
                tit.ContaContabilCliente = NF.Itens(0).Produto.CarteiraCompra.CodigoContaCliente
                tit.CodigoCarteira = NF.Itens(0).Produto.CodigoCarteiraCompra
            End If
        Else
            tit.ReceberPagar = "R"
            If NF.SubOperacao.Devolucao Then
                tit.ContaContabilCliente = NF.Itens(0).Produto.CarteiraCompra.CodigoContaCliente
                tit.CodigoCarteira = NF.Itens(0).Produto.CodigoCarteiraCompra
            Else
                tit.ContaContabilCliente = NF.Itens(0).Produto.CarteiraVenda.CodigoContaCliente
                tit.CodigoCarteira = NF.Itens(0).Produto.CodigoCarteiraVenda
            End If
        End If

        tit.Tributo = ""
        tit.CodigoEmpresaPedido = NF.CodigoEmpresa
        tit.EndEmpresaPedido = NF.EnderecoEmpresa

        tit.DataMoeda = NF.DataNota
        tit.Movimento = NF.Movimento.ToString("dd-MM-yyyy")
        tit.IndiceFixo = True
        tit.CodigoIndexador = 2
        tit.CodigoMoeda = 1
        tit.Vencimento = Funcoes.ValidaDataUtil(NF.Empresa.Codigo, NF.Empresa.CodigoEndereco, CDate(row("Vencimento")))
        tit.Prorrogacao = Funcoes.ValidaDataUtil(NF.Empresa.Codigo, NF.Empresa.CodigoEndereco, CDate(row("Vencimento")))
        tit.Baixa = Funcoes.ValidaDataUtil(NF.Empresa.Codigo, NF.Empresa.CodigoEndereco, CDate(row("Vencimento")))

        tit.CodigoPedido = NF.CodigoPedido
        tit.CodigoPedidoFixacao = 0
        tit.CodigoProcuracao = NF.CodigoProcuracao

        tit.CodigoTipoPgto = 1
        tit.CodigoSituacao = 1
        tit.Lote = 70

        tit.CodigoUnidadeDeNegocio = NF.CodigoUnidadeDeNegocio
        tit.CodigoEmpresa = NF.CodigoEmpresa
        tit.EnderecoEmpresa = NF.EnderecoEmpresa
        tit.CodigoCliente = NF.CodigoCliente
        tit.EndCliente = NF.EnderecoCliente
        tit.CodigoBancoCliente = 0
        tit.CodigoAgenciaCliente = ""
        tit.DigitoAgenciaCliente = ""
        tit.ContaCliente = ""
        tit.DigitoContaCliente = ""
        tit.Cheque = False
        tit.Slips = False
        tit.Recibo = False
        tit.Aviso = False
        tit.ReciboDeposito = False

        If _TipoMoeda = eTiposMoeda.Oficial Then
            tit.ValorDoDocumento = row("ValorParcela")
        Else
            tit.MoedaValorDoDocumento = row("ValorParcela")
        End If

        'tit.Historico = "REF. NF " & NF.Codigo & "-" & NF.Serie & "-" & IIf(NF.EntradaSaida = eEntradaSaida.Saida, "S", "E") & ", Parcela " & parcela & "/" & row("Parcelas") & ", Pedido: " & NF.CodigoPedido & " / " & NF.Cliente.Nome

        tit.CodigoDeBarras = ""
        tit.CodigoDigitado = False
        tit.CodigoDeBarrasPreImpresso = False

        tit.CodigoDestinatario = NF.CodigoCliente
        tit.EndDestinatario = NF.EnderecoCliente
        tit.NomeDoDestinatario = ""
        tit.Destinacao = ""
        tit.Solicitacao = 0

        tit.Agrupado = eAgrupamentoFinanceiro.NaoAgrupado
        tit.RegistroMestre = 0
        tit.Observacoes = ""
        tit.SituacaoBancaria = 0
        tit.NumeroDoCheque = 0
        tit.CodigoAdiantamento = 0
        tit.TaxaAdto = 0

        tit.UsuarioInclusao = NF.UsuarioInclusao
        tit.UsuarioInclusaoData = NF.DataInclusao

        tit.CodigoCarteiraDoTitulo = 0

    End Sub

    Public Sub ParcelarFixacao(ByRef Fixacao As Fixacao, ByRef Pedido As Pedido)
        Dim Banco As New AcessaBanco
        Dim ds As New DataSet
        Dim Sql As String = ""

        Dim ValorAParcelarOficial As Decimal
        Dim ValorAParcelarMoeda As Decimal
        Dim ValorAParcelar As Decimal

        For Each Enc As FixacaoXEncargo In Fixacao.Encargos
            If Enc.CodigoEncargo = "LIQUIDO" Then
                ValorAParcelarOficial += Enc.ValorOficial
                ValorAParcelarMoeda += Enc.ValorMoeda
            End If
        Next

        If Pedido.Moeda.Classificacao = eTiposMoeda.Oficial Then
            ValorAParcelar = ValorAParcelarOficial
        Else
            ValorAParcelar = ValorAParcelarMoeda
        End If

        Sql = "	DECLARE " & vbCrLf &
               "	@Diferenca numeric(18,2)," & vbCrLf &
               "	@ValorTotal numeric(18,2)," & vbCrLf &
               "	@Data varchar(10)," & vbCrLf &
               "	@FPagto int" & vbCrLf

        '-- Na Compra dilui em todas as parcelas nas outras cobra na primeira parcela
        Sql &= "	set @ValorTotal =    " & Str(ValorAParcelar)

        Sql &= "	set @Data       = '" & Fixacao.Movimento.ToString("yyyy-MM-dd") & "'" & vbCrLf &
               "	set @FPagto     =  " & Fixacao.CondicaoPagamento & vbCrLf

        '--Seleciona no banco a forma de pagamento e as Parcelas"
        Sql &= "	SELECT Pagamentos.Pagamento_Id," & vbCrLf &
               "		   Pagamentos.Descricao," & vbCrLf &
               "		   PagamentosXParcelas.Sequencia_Id, " & vbCrLf &
               "		   Pagamentos.Parcelas," & vbCrLf &
               "		   PagamentosXParcelas.Dias," & vbCrLf
        '--Soma a Data Base o Numero de dias referente ao numero da Parcela"
        Sql &= "		   convert(datetime,(@Data)) + PagamentosXParcelas.Dias as Vencimento," & vbCrLf
        '--Divide o valor total pelo numero de parcelas p/ descubrir o valor da parcela"
        Sql &= "		   round(@ValorTotal / Pagamentos.Parcelas, 2) as ValorParcela," & vbCrLf
        '-- Armazena o Valor Total para calcular a diferenca na divisao das parcelas"
        Sql &= "		   @ValorTotal as ValorTotal" & vbCrLf &
               "	  INTO #Temp1 " & vbCrLf &
               "	  FROM Pagamentos " & vbCrLf &
               "	 INNER JOIN PagamentosXParcelas " & vbCrLf &
               "		ON Pagamentos.Pagamento_Id = PagamentosXParcelas.Pagamento_Id " & vbCrLf &
               "	 where Pagamentos.Pagamento_Id = @FPagto " & vbCrLf &
               "	 order by PagamentosXParcelas.Sequencia_Id " & vbCrLf

        '	--Calcula o valor Parcelado para ver se ha diferenca com o valor total caso haja armazena o valor na variavel "
        Sql &= "	set @Diferenca = (Select top(1) ValorTotal - (ValorParcela * Parcelas) from #Temp1)" & vbCrLf


        '	--Atualiza o valor da primeira parcela acrescida do IPI + a Diferenca"
        Sql &= "	update #Temp1 set" & vbCrLf &
               "	ValorParcela = ValorParcela + @Diferenca" & vbCrLf &
               "	where #Temp1.Sequencia_Id = 1" & vbCrLf &
               "	select Pagamento_Id, Descricao, Sequencia_Id, Parcelas, Dias, Vencimento, ValorParcela, ValorTotal from #Temp1" & vbCrLf

        ds = Banco.ConsultaDataSet(Sql, "Titulos")

        Dim n As Numerador = New Numerador(1)
        Dim NumeroTitulo As Integer = n.Sequencia + 1
        Dim parcela As Integer

        For Each Row As DataRow In ds.Tables("Titulos").Rows
            Dim Tit As New Titulo
            Tit.IUD = "I"
            Tit.Codigo = NumeroTitulo
            NumeroTitulo += 1
            parcela += 1
            Tit.Sequencia = 0
            Tit.CodigoProvisao = 2

            If Pedido.SubOperacao.EntradaSaida = eEntradaSaida.Entrada Then
                Tit.ReceberPagar = "P"
            Else
                Tit.ReceberPagar = "R"
            End If

            Tit.ContaContabilCliente = Fixacao.ItemPedido.Produto.CarteiraCompra.CodigoContaCliente
            Tit.CodigoCarteira = Fixacao.ItemPedido.Produto.CodigoCarteiraCompra

            Tit.Tributo = ""
            Tit.CodigoUnidadeDeNegocio = Pedido.CodigoUnidadeNegocio
            Tit.CodigoEmpresa = Pedido.CodigoEmpresa
            Tit.EnderecoEmpresa = Pedido.EnderecoEmpresa
            Tit.CodigoCliente = Pedido.CodigoCliente
            Tit.EndCliente = Pedido.EnderecoCliente
            Tit.CodigoEmpresaPedido = Pedido.CodigoEmpresa
            Tit.EndEmpresaPedido = Pedido.EnderecoEmpresa
            Tit.CodigoPedido = Pedido.Codigo
            Tit.CodigoPedidoFixacao = Fixacao.Codigo
            Tit.CodigoProcuracao = Fixacao.Procuracao
            Tit.Movimento = Fixacao.Movimento.ToString("dd-MM-yyyy")
            Tit.DataMoeda = Fixacao.Movimento.ToString("dd-MM-yyyy")
            Tit.CodigoIndexador = Pedido.CodigoIndexador
            Tit.CodigoMoeda = Pedido.CodigoMoeda
            Tit.Vencimento = Row("Vencimento")
            Tit.Prorrogacao = Row("Vencimento")
            Tit.Baixa = Row("Vencimento")
            Tit.CodigoTipoPgto = 1
            Tit.CodigoSituacao = 1
            Tit.Lote = 70
            Tit.CodigoBancoCliente = 0
            Tit.CodigoAgenciaCliente = ""
            Tit.DigitoAgenciaCliente = ""
            Tit.ContaCliente = ""
            Tit.DigitoContaCliente = ""
            Tit.Cheque = False
            Tit.Slips = False
            Tit.Recibo = False
            Tit.Aviso = False
            Tit.ReciboDeposito = False
            Tit.IndiceTitulo = Fixacao.IndiceFixado

            If Pedido.Moeda.Classificacao = eTiposMoeda.Oficial Then
                Tit.ValorDoDocumento = Row("ValorParcela")
            Else
                Tit.MoedaValorDoDocumento = Row("ValorParcela")
            End If

            Tit.Historico = "REF. FIXACAO " & Fixacao.Codigo & ", Parcela " & parcela & "/" & Row("Parcelas") & ", Pedido: " & Pedido.Codigo & " - " & Pedido.Cliente.Nome
            Tit.CodigoDeBarras = ""
            Tit.CodigoDigitado = False
            Tit.CodigoDeBarrasPreImpresso = False

            Tit.CodigoDestinatario = Pedido.CodigoCliente
            Tit.EndDestinatario = Pedido.EnderecoCliente
            Tit.NomeDoDestinatario = Pedido.Cliente.Nome
            Tit.Destinacao = ""
            Tit.Solicitacao = 0

            Tit.Agrupado = eAgrupamentoFinanceiro.NaoAgrupado
            Tit.RegistroMestre = 0
            Tit.Observacoes = ""
            Tit.SituacaoBancaria = 0
            Tit.NumeroDoCheque = 0
            Tit.CodigoAdiantamento = 0
            Tit.TaxaAdto = 0

            Tit.UsuarioInclusao = HttpContext.Current.Session("ssNomeUsuario")
            Tit.UsuarioInclusaoData = Now()

            Tit.CodigoCarteiraDoTitulo = 0

            Me.Add(Tit)
        Next
    End Sub

    Public Sub AdicionarValor(ByVal Valor As Decimal)
        Dim ValorPorParcela As Decimal = Math.Round((Valor / Me.Count) - 0.005, 2)
        Dim Diferenca As Decimal = (ValorPorParcela * Me.Count) - Valor
        For i As Integer = 0 To Me.Count - 1
            Me(i).IUD = "U"
            If _TipoMoeda = eTiposMoeda.Oficial Then
                Me(i).ValorDoDocumento += ValorPorParcela + IIf(i = 0, Diferenca, 0)
            Else
                Me(i).MoedaValorDoDocumento += ValorPorParcela + IIf(i = 0, Diferenca, 0)
            End If
        Next
    End Sub

    Public Sub AjustarProvisionadoExclusaoNF(ByVal Valor As Decimal)
        If NF.VencimentosPedido.Count = 0 Then Exit Sub

        Dim Saldo As Decimal = Valor

        For Each tit As Titulo In NF.VencimentosPedido
            If tit.CodigoProvisao = 3 And tit.ReceberPagar = RP_Pedido And Saldo > 0 Then

                If _TipoMoeda = eTiposMoeda.Oficial Then
                    '****************** Oficial ******************************
                    tit.IUD = "U"
                    tit.ValorDoDocumento += Saldo
                Else
                    '****************** Moeda ********************************
                    tit.IUD = "U"
                    tit.MoedaValorDoDocumento += Saldo

                    If NF.Pedido.CodigoIndexador = 99 OrElse NF.Pedido.IndexadorFixo Then
                        tit.ValorDoDocumento = Math.Round(NF.Pedido.IndiceFixado * tit.MoedaValorDoDocumento, 2, MidpointRounding.AwayFromZero)
                    Else
                        Dim indiceDolar As Decimal = New Cotacao(NF.Pedido.CodigoIndexador, NF.DataNota).Indice 'PTAX

                        tit.ValorDoDocumento = Math.Round(indiceDolar * tit.MoedaValorDoDocumento, 10, MidpointRounding.AwayFromZero)
                    End If
                End If
            End If
        Next
    End Sub


    Public Sub AjustarTitulosProvisionados(ByVal Valor As Decimal)
        If NF.VencimentosPedido.Count = 0 Then Exit Sub

        'Dim ValorProvisionado As Decimal
        'ValorProvisionado = NF.VencimentosPedido.ConsultaValorNaLista(RP_Pedido, TipoMoeda, 3)

        'If ValorProvisionado = Valor Then
        '    For Each tit As NGS.Titulo In NF.VencimentosPedido
        '        If tit.CodigoProvisao = 3 And tit.ReceberPagar = RP_Pedido Then
        '            tit.IUD = "D"
        '            If _TipoMoeda = eTiposMoeda.Oficial Then
        '                tit.ValorDoDocumento = 0
        '            Else
        '                tit.MoedaValorDoDocumento = 0
        '            End If
        '        End If
        '    Next
        'Else
        '    Dim ProvisionadosCount As Integer = NF.VencimentosPedido.NumeroDeTitulosProvisionados()

        '    Dim ValorPorParcela As Decimal = Math.Round((Valor / ProvisionadosCount) - 0.005, 2)
        '    Dim Diferenca As Decimal = Math.Round(Valor - (ValorPorParcela * ProvisionadosCount), 2)
        '    Dim Saldo As Decimal = Valor

        '    While Saldo <> 0
        '        If ProvisionadosCount > 0 Then
        '            For Each tit As NGS.Titulo In NF.VencimentosPedido
        '                If tit.CodigoProvisao = 3 And tit.ReceberPagar = RP_Pedido Then
        '                    If _TipoMoeda = eTiposMoeda.Oficial Then
        '                        '****************** Oficial ******************************
        '                        If tit.ValorDoDocumento >= ValorPorParcela Then
        '                            tit.IUD = "U"
        '                            tit.ValorDoDocumento -= ValorPorParcela
        '                            Saldo -= ValorPorParcela
        '                            If Diferenca > 0 And Diferenca < tit.ValorDoDocumento Then
        '                                tit.IUD = "U"
        '                                tit.ValorDoDocumento -= Diferenca
        '                                Saldo -= Diferenca
        '                                Diferenca = 0
        '                            ElseIf Diferenca > 0 And Diferenca > tit.ValorDoDocumento Then
        '                                tit.IUD = "D"
        '                                Diferenca -= tit.ValorDoDocumento
        '                                Saldo -= tit.ValorDoDocumento
        '                                tit.ValorDoDocumento = 0
        '                            End If
        '                            If tit.ValorDoDocumento = 0 Then tit.IUD = "D"
        '                        End If
        '                    Else
        '                        '****************** Moeda ******************************
        '                        If tit.MoedaValorDoDocumento >= ValorPorParcela Then
        '                            tit.IUD = "U"
        '                            tit.MoedaValorDoDocumento -= ValorPorParcela
        '                            Saldo -= ValorPorParcela
        '                            If Diferenca > 0 And Diferenca < tit.MoedaValorDoDocumento Then
        '                                tit.IUD = "U"
        '                                tit.MoedaValorDoDocumento -= Diferenca
        '                                Saldo -= Diferenca
        '                                Diferenca = 0
        '                            ElseIf Diferenca > 0 And Diferenca > tit.MoedaValorDoDocumento Then
        '                                tit.IUD = "D"
        '                                Diferenca -= tit.MoedaValorDoDocumento
        '                                Saldo -= tit.MoedaValorDoDocumento
        '                                tit.MoedaValorDoDocumento = 0
        '                            End If
        '                            If tit.MoedaValorDoDocumento = 0 Then tit.IUD = "D"
        '                        End If
        '                    End If
        '                End If
        '            Next
        '        End If
        '    End While
        'End If

        Dim Saldo As Decimal = Valor

        For Each tit As Titulo In NF.VencimentosPedido
            If tit.CodigoProvisao = 3 And tit.ReceberPagar = RP_Pedido And Saldo > 0 Then
                'TRATANDO COMPLEMENTO EM DÓLAR PARA RTGRÃOS - FURLAN 04/11/2024
                If (Left(NF.CodigoEmpresa, 8) = "24450490" OrElse Left(NF.CodigoEmpresa, 8) = "44979506") AndAlso NF.SubOperacao.QuantidadeFiscal = False AndAlso NF.SubOperacao.FinalidadeDaNota = 2 AndAlso _TipoMoeda = eTiposMoeda.MoedaEstrangeira Then
                    If tit.ValorDoDocumento > Saldo Then
                        tit.IUD = "U"
                        tit.ValorDoDocumento -= Saldo
                        Saldo = 0
                    Else
                        tit.IUD = "D"
                        Saldo -= tit.ValorDoDocumento
                        tit.ValorDoDocumento = 0
                    End If
                ElseIf _TipoMoeda = eTiposMoeda.Oficial Then
                    '****************** Oficial ******************************
                    If tit.ValorDoDocumento > Saldo Then
                        tit.IUD = "U"
                        tit.ValorDoDocumento -= Saldo
                        Saldo = 0
                    Else
                        tit.IUD = "D"
                        Saldo -= tit.ValorDoDocumento
                        tit.ValorDoDocumento = 0
                    End If
                Else
                    '****************** Moeda ******************************
                    If tit.MoedaValorDoDocumento > Saldo Then
                        tit.IUD = "U"
                        tit.MoedaValorDoDocumento -= Saldo
                        Saldo = 0

                        If NF.Pedido.CodigoIndexador = 99 OrElse NF.Pedido.IndexadorFixo Then
                            tit.ValorDoDocumento = Math.Round(NF.Pedido.IndiceFixado * tit.MoedaValorDoDocumento, 2, MidpointRounding.AwayFromZero)
                        Else
                            Dim indiceDolar As Decimal = New Cotacao(NF.Pedido.CodigoIndexador, NF.DataNota).Indice 'PTAX

                            tit.ValorDoDocumento = Math.Round(indiceDolar * tit.MoedaValorDoDocumento, 10, MidpointRounding.AwayFromZero)
                        End If
                    Else
                        tit.IUD = "D"
                        Saldo -= tit.MoedaValorDoDocumento
                        tit.MoedaValorDoDocumento = 0
                    End If
                End If
            End If
        Next
    End Sub

    Public Sub AjustarTitulosPrevisionados(ByVal Valor As Decimal)
        If NF.VencimentosPedido.Count = 0 Then Exit Sub
        Dim ValorPrevisionado As Decimal
        ValorPrevisionado = NF.VencimentosPedido.ConsultaValorNaLista(RP_Pedido, TipoMoeda, 2)

        If ValorPrevisionado = Valor Then
            For Each tit As Titulo In NF.VencimentosPedido
                If tit.CodigoProvisao = 2 And tit.ReceberPagar = RP_Pedido Then
                    tit.IUD = "D"
                End If
            Next
        Else
            Dim PrevisionadosCount As Integer = NF.VencimentosPedido.NumeroDeTitulosPrevisionados()


            Dim ValorPorParcela As Decimal = Math.Round((Valor / PrevisionadosCount) - 0.005, 2)
            Dim Diferenca As Decimal = Math.Round(Valor - (ValorPorParcela * PrevisionadosCount), 2)
            Dim Saldo As Decimal = Valor

            While Saldo <> 0
                If PrevisionadosCount > 0 Then
                    For Each tit As Titulo In NF.VencimentosPedido
                        If tit.CodigoProvisao = 2 And tit.ReceberPagar = RP_Pedido Then
                            If _TipoMoeda = eTiposMoeda.Oficial Then
                                '****************** Oficial ******************************
                                If tit.ValorDoDocumento >= ValorPorParcela Then
                                    tit.IUD = "U"
                                    tit.ValorDoDocumento -= ValorPorParcela
                                    Saldo -= ValorPorParcela
                                    If Diferenca > 0 And Diferenca < tit.ValorDoDocumento Then
                                        tit.IUD = "U"
                                        tit.ValorDoDocumento -= Diferenca
                                        Saldo -= Diferenca
                                        Diferenca = 0
                                    ElseIf Diferenca > 0 And Diferenca > tit.ValorDoDocumento Then
                                        tit.IUD = "D"
                                        Diferenca -= tit.ValorDoDocumento
                                        Saldo -= tit.ValorDoDocumento
                                        tit.ValorDoDocumento = 0
                                    End If
                                    If tit.ValorDoDocumento = 0 Then tit.IUD = "D"
                                End If
                            Else
                                '****************** Moeda ******************************
                                If tit.MoedaValorDoDocumento >= ValorPorParcela Then
                                    tit.IUD = "U"
                                    tit.MoedaValorDoDocumento -= ValorPorParcela
                                    Saldo -= ValorPorParcela
                                    If Diferenca > 0 And Diferenca < tit.MoedaValorDoDocumento Then
                                        tit.IUD = "U"
                                        tit.MoedaValorDoDocumento -= Diferenca
                                        Saldo -= Diferenca
                                        Diferenca = 0
                                    ElseIf Diferenca > 0 And Diferenca > tit.MoedaValorDoDocumento Then
                                        tit.IUD = "D"
                                        Diferenca -= tit.MoedaValorDoDocumento
                                        Saldo -= tit.MoedaValorDoDocumento
                                        tit.MoedaValorDoDocumento = 0
                                    End If
                                    If tit.MoedaValorDoDocumento = 0 Then tit.IUD = "D"
                                End If
                            End If
                        End If
                    Next
                End If
            End While
        End If
    End Sub

    Public Function AjustaParcelas(ByVal Apartir As Integer, ByVal ValorOriginal As Decimal, ByVal ValorNovo As Decimal) As String
        'ULTIMA LINHA
        If (Apartir = Me.Count - 1) Then
            If TipoMoeda = eTiposMoeda.Oficial Then
                Me(Apartir).ValorDoDocumento = ValorOriginal
            Else
                Me(Apartir).MoedaValorDoDocumento = ValorOriginal
            End If

            If ValorNovo = ValorOriginal Then
                Return ""
            Else
                Return "Parcela Unica ou Ultima Parcela nao pode ser alterada"
            End If
        End If

        Dim saldo As Decimal = ValorOriginal - ValorNovo
        Dim numParcelas As Integer = (Me.Count - 1 - Apartir)
        Dim parcelas As Decimal = Math.Round(saldo / numParcelas, 2)
        Dim diferenca As Decimal = saldo - (parcelas * numParcelas)

        saldo = 0
        If (Apartir < Me.Count - 1) And ValorNovo > ValorOriginal Then
            For i As Integer = Apartir + 1 To Me.Count - 1
                If TipoMoeda = eTiposMoeda.Oficial Then
                    saldo += Me(i).ValorDoDocumento
                Else
                    saldo += Me(i).MoedaValorDoDocumento
                End If
            Next

            If saldo <= ValorNovo - ValorOriginal Then
                If TipoMoeda = eTiposMoeda.Oficial Then
                    Me(Apartir).ValorDoDocumento = ValorOriginal
                Else
                    Me(Apartir).MoedaValorDoDocumento = ValorOriginal
                End If
                Return "Valor Informado ultrapassa o Valor da transacao"
            End If
        End If

        For i As Integer = Apartir + 1 To Me.Count - 1
            If TipoMoeda = eTiposMoeda.Oficial Then
                If i = Me.Count - 1 Then
                    Me(i).ValorDoDocumento += parcelas + diferenca
                Else
                    Me(i).ValorDoDocumento += parcelas
                End If
            Else
                If i = Me.Count - 1 Then
                    Me(i).MoedaValorDoDocumento += parcelas + diferenca
                Else
                    Me(i).MoedaValorDoDocumento += parcelas
                End If
            End If

        Next
        Return ""
    End Function

    '*****************  Uteis  **********************
    Private Function NumeroDeTitulosProvisionados() As Integer
        Dim count As Integer
        For Each tit As Titulo In Me
            If tit.CodigoProvisao = 3 And tit.ReceberPagar = RP_Pedido Then count += 1
        Next
        Return count
    End Function

    Private Function NumeroDeTitulosPrevisionados() As Integer
        Dim count As Integer
        For Each tit As Titulo In Me
            If tit.CodigoProvisao = 2 And tit.ReceberPagar = RP_Pedido Then count += 1
        Next
        Return count
    End Function

    Private Function ConsultaValorNaLista(ByVal ReceberPagar As String, ByVal Moeda As eTiposMoeda, Optional ByVal CodigoProvisao As Integer = 0) As Decimal
        Dim _valor As Decimal = 0
        If CodigoProvisao = 0 Then
            For Each tit As Titulo In Me
                If tit.ReceberPagar = ReceberPagar Then _valor += IIf(Moeda = eTiposMoeda.MoedaEstrangeira, tit.MoedaValorDoDocumento, tit.ValorDoDocumento)
            Next
        Else
            For Each tit As Titulo In Me
                If tit.ReceberPagar = ReceberPagar And tit.CodigoProvisao = CodigoProvisao Then _valor += IIf(Moeda = eTiposMoeda.MoedaEstrangeira, tit.MoedaValorDoDocumento, tit.ValorDoDocumento)
            Next
        End If

        Return _valor
    End Function
#End Region

End Class

<Serializable()>
Public Class Titulo
    Implements IBaseEntity

#Region "Uteis"
    Private _Carregando As Boolean
    Private _Controlando As Boolean = True
    Private _MsgControle As String = ""

#End Region

#Region "Contrutor"
    Public Sub New()

    End Sub

    Public Sub New(ByVal pTitulo As Integer, Optional ByVal pReceberPagar As String = "")
        Dim sql As String

        sql = "SELECT * " & vbCrLf &
                   "  FROM (" & vbCrLf &
                   "		SELECT 'P' as ReceberPagar, Registro_Id, Sequencia_Id, Provisao, isnull(Carteira,0) as Carteira, Tributo, Indexador, Moeda, TipoPagto, Situacao, Lote, Movimento, Vencimento, Prorrogacao," & vbCrLf &
                   "			   DataMoeda, Baixa, UnidadeDeNegocio, Empresa, EndEmpresa, Cliente, EndCliente, BancoCliente, AgenciaCliente, DigitoAgenciaCliente, ContaCliente," & vbCrLf &
                   "			   DigitoContaCliente, ContaContabilCliente, EmpresaPagadora, EndEmpresaPagadora, isnull(BancoPagador,0) AS BancoPagador, AgenciaPagadora, DigitoAgenciaPagadora," & vbCrLf &
                   "			   ContaPagadora, DigitoContaPagadora, ContaContabilPagadora, Cheque, Slips, isnull(Recibo,'N') AS Recibo, isnull(Aviso,'N') AS Aviso, isnull(ReciboDeposito,'N') AS ReciboDeposito, isnull(EmpresaPedido,'') as EmpresaPedido, isnull(EndEmpresaPedido,0) as EndEmpresaPedido," & vbCrLf &
                   "			   isnull(Pedido,0) as Pedido, isnull(PedidoFixacao,0) as PedidoFixacao, isnull(Procuracao,0) as Procuracao, ValorDoDocumento, Descontos, Deducoes, Juros, Acrescimos, ValorLiquido, isnull(MoedaValorDoDocumento,0) as MoedaValorDoDocumento," & vbCrLf &
                   "			   isnull(MoedaDescontos,0) as MoedaDescontos, isnull(MoedaDeducoes,0) as MoedaDeducoes, isnull(MoedaJuros,0) as MoedaJuros, isnull(MoedaAcrescimos,0) as MoedaAcrescimos, isnull(MoedaValorLiquido,0) as MoedaValorLiquido, Historico, CodigoDeBarras, CodigoDigitado, isnull(CodigoDeBarraPreImpresso,0) AS CodigoDeBarraPreImpresso, Destinatario," & vbCrLf &
                   "			   EndDestinatario, isnull(NomeDoDestinatario,'') as NomeDoDestinatario, isnull(Destinacao,'') as Destinacao, solicitacao, UsuarioInclusao, isnull(UsuarioInclusaoData,Movimento) as UsuarioInclusaoData,  isnull(UsuarioAlteracao,'') as UsuarioAlteracao, UsuarioAlteracaoData," & vbCrLf &
                   "			   isnull(UsuarioCancelamento,'') as UsuarioCancelamento, UsuarioCancelamentoData, isnull(UsuarioLiberacao,'') as UsuarioLiberacao, UsuarioLiberacaoData, isnull(UsuarioBaixa,'') as UsuarioBaixa, UsuarioBaixaData, isnull(Grupado,'N') AS Grupado," & vbCrLf &
                   "			   isnull(RegistroMestre,0) as RegistroMestre, convert(varchar(4000),Observacoes) as Observacoes, isnull(SituacaoBancaria,0) as SituacaoBancaria, isnull(NumeroDoCheque,0) as NumeroDoCheque, isnull(Adiantamento,0) as Adiantamento," & vbCrLf &
                   "               VencimentoAdto, isnull(TaxaAdto,0) as TaxaAdto, isnull(UsuarioLiberacaoBloqueio,'') as UsuarioLiberacaoBloqueio," & vbCrLf &
                   "			   UsuarioLiberacaoBloqueioDate, isnull(UsuarioLiberacaoPedido,'') as UsuarioLiberacaoPedido, UsuarioLiberacaoPedidoDate, isnull(UsuarioLiberacaoCheque,'') as UsuarioLiberacaoCheque," & vbCrLf &
                   "			   UsuarioLiberacaoChequeDate, isnull(ContratoBancario,'') AS ContratoBancario, isnull(CarteiraAdto,'') AS CarteiraAdto, isnull(CarteiraDoTitulo,0) AS CarteiraDoTitulo, " & vbCrLf &
                   "               isnull(ContratoDeFinanciamento,'') AS ContratoDeFinanciamento, DataEnvio, isnull(Ocorrencia,0) AS Ocorrencia, isnull(motivo,'') AS motivo, isnull(DigitoNossoNumero,'') AS DigitoNossoNumero, " & vbCrLf &
                   "               isnull(NossoNumero,0) AS NossoNumero, isnull(Timbrado,'') AS Timbrado, isnull(ModalidadeDePagamento,0) AS ModalidadeDePagamento, CreditoNaConta, isnull(NumeroDaRemessa,0) AS NumeroDaRemessa, " & vbCrLf &
                   "               isnull(SituacaoRemessaBancaria,0) AS SituacaoRemessaBancaria, isnull(TipoContaCliente,'') AS TipoContaCliente, isnull(TituloOrigem,0) AS TituloOrigem, isnull(ValorRecompra,0) AS ValorRecompra, " & vbCrLf &
                   "               isnull(TipoContaPagadora,'') AS TipoContaPagadora, isnull(Processo,'') AS Processo, isnull(BoletoBancario,0) AS BoletoBancario, isnull(UsuarioBoletoBancario,'') as UsuarioBoletoBancario, UsuarioBoletoBancarioDate, " & vbCrLf &
                   "               isnull(HistoricoRemessa,'') AS HistoricoRemessa, isnull(ObservacoesControleInterno,'') AS ObservacoesControleInterno, isnull(FolhaDePagamento,0) AS FolhaDePagamento, isnull(UsuarioFolhaDePagamento,'') as UsuarioFolhaDePagamento, UsuarioFolhaDePagamentoDate " & vbCrLf &
                   "		  FROM ContasAPagar" & vbCrLf &
                   "         UNION" & vbCrLf &
                   "		SELECT 'R' as ReceberPagar, Registro_Id, Sequencia_Id, Provisao, isnull(Carteira,0) as Carteira, Tributo, Indexador, Moeda, TipoPagto, Situacao, Lote, Movimento, Vencimento, Prorrogacao," & vbCrLf &
                   "			   DataMoeda, Baixa, UnidadeDeNegocio, Empresa, EndEmpresa, Cliente, EndCliente, BancoCliente, AgenciaCliente, DigitoAgenciaCliente, ContaCliente," & vbCrLf &
                   "			   DigitoContaCliente, ContaContabilCliente, EmpresaPagadora, EndEmpresaPagadora, isnull(BancoPagador,0) AS BancoPagador, AgenciaPagadora, DigitoAgenciaPagadora," & vbCrLf &
                   "			   ContaPagadora, DigitoContaPagadora, ContaContabilPagadora, Cheque, Slips, isnull(Recibo,'N') AS Recibo, isnull(Aviso,'N') AS Aviso, isnull(ReciboDeposito,'N') AS ReciboDeposito, isnull(EmpresaPedido,'') as EmpresaPedido, isnull(EndEmpresaPedido,0) as EndEmpresaPedido," & vbCrLf &
                   "			   isnull(Pedido,0) as Pedido, isnull(PedidoFixacao,0) as PedidoFixacao, isnull(Procuracao,0) as Procuracao, ValorDoDocumento, Descontos, Deducoes, Juros, Acrescimos, ValorLiquido, isnull(MoedaValorDoDocumento,0) as MoedaValorDoDocumento," & vbCrLf &
                   "			   isnull(MoedaDescontos,0) as MoedaDescontos, isnull(MoedaDeducoes,0) as MoedaDeducoes, isnull(MoedaJuros,0) as MoedaJuros, isnull(MoedaAcrescimos,0) as MoedaAcrescimos, isnull(MoedaValorLiquido,0) as MoedaValorLiquido, Historico, CodigoDeBarras, CodigoDigitado, isnull(CodigoDeBarraPreImpresso,0) AS CodigoDeBarraPreImpresso, Destinatario," & vbCrLf &
                   "			   EndDestinatario, isnull(NomeDoDestinatario,'') as NomeDoDestinatario, isnull(Destinacao,'') as Destinacao, solicitacao, UsuarioInclusao, isnull(UsuarioInclusaoData,Movimento) as UsuarioInclusaoData, isnull(UsuarioAlteracao,'') as UsuarioAlteracao, UsuarioAlteracaoData," & vbCrLf &
                   "			   isnull(UsuarioCancelamento,'') as UsuarioCancelamento, UsuarioCancelamentoData, isnull(UsuarioLiberacao,'') as UsuarioLiberacao, UsuarioLiberacaoData, isnull(UsuarioBaixa,'') as UsuarioBaixa, UsuarioBaixaData, isnull(Grupado,'N') AS Grupado," & vbCrLf &
                   "			   isnull(RegistroMestre,0) as RegistroMestre, convert(varchar(4000),Observacoes) as Observacoes, isnull(SituacaoBancaria,0) as SituacaoBancaria, isnull(NumeroDoCheque,0) as NumeroDoCheque, isnull(Adiantamento,0) as Adiantamento," & vbCrLf &
                   "               VencimentoAdto, isnull(TaxaAdto,0) as TaxaAdto, isnull(UsuarioLiberacaoBloqueio,'') as UsuarioLiberacaoBloqueio," & vbCrLf &
                   "			   UsuarioLiberacaoBloqueioDate, isnull(UsuarioLiberacaoPedido,'') as UsuarioLiberacaoPedido, UsuarioLiberacaoPedidoDate, isnull(UsuarioLiberacaoCheque,'') as UsuarioLiberacaoCheque," & vbCrLf &
                   "			   UsuarioLiberacaoChequeDate, isnull(ContratoBancario,'') AS ContratoBancario, isnull(CarteiraAdto,'') AS CarteiraAdto, isnull(CarteiraDoTitulo,0) AS CarteiraDoTitulo, " & vbCrLf &
                   "               isnull(ContratoDeFinanciamento,'') AS ContratoDeFinanciamento, DataEnvio, isnull(Ocorrencia,0) AS Ocorrencia, isnull(motivo,'') AS motivo, isnull(DigitoNossoNumero,'') AS DigitoNossoNumero, " & vbCrLf &
                   "               isnull(NossoNumero,0) AS NossoNumero, isnull(Timbrado,'') AS Timbrado, isnull(ModalidadeDePagamento,0) AS ModalidadeDePagamento, CreditoNaConta, isnull(NumeroDaRemessa,0) AS NumeroDaRemessa, " & vbCrLf &
                   "               isnull(SituacaoRemessaBancaria,0) AS SituacaoRemessaBancaria, isnull(TipoContaCliente,'') AS TipoContaCliente, isnull(TituloOrigem,0) AS TituloOrigem, isnull(ValorRecompra,0) AS ValorRecompra, " & vbCrLf &
                   "               isnull(TipoContaPagadora,'') AS TipoContaPagadora, isnull(Processo,'') AS Processo, isnull(BoletoBancario,0) AS BoletoBancario, isnull(UsuarioBoletoBancario,'') as UsuarioBoletoBancario, UsuarioBoletoBancarioDate, " & vbCrLf &
                   "               isnull(HistoricoRemessa,'') AS HistoricoRemessa, isnull(ObservacoesControleInterno,'') AS ObservacoesControleInterno, isnull(FolhaDePagamento,0) AS FolhaDePagamento, isnull(UsuarioFolhaDePagamento,'') as UsuarioFolhaDePagamento, UsuarioFolhaDePagamentoDate " & vbCrLf &
                   "		  FROM ContasAReceber" & vbCrLf &
                   "         UNION" & vbCrLf &
                   "		SELECT 'M' as ReceberPagar, Registro_Id, Sequencia_Id, Provisao, isnull(Carteira,0) as Carteira, Tributo, Indexador, Moeda, TipoPagto, Situacao, Lote, Movimento, Vencimento, Prorrogacao," & vbCrLf &
                   "			   DataMoeda, Baixa, UnidadeDeNegocio, Empresa, EndEmpresa, Cliente, EndCliente, BancoCliente, AgenciaCliente, DigitoAgenciaCliente, ContaCliente," & vbCrLf &
                   "			   DigitoContaCliente, ContaContabilCliente, EmpresaPagadora, EndEmpresaPagadora, isnull(BancoPagador,0) AS BancoPagador, AgenciaPagadora, DigitoAgenciaPagadora," & vbCrLf &
                   "			   ContaPagadora, DigitoContaPagadora, ContaContabilPagadora, Cheque, Slips, isnull(Recibo,'N') AS Recibo, isnull(Aviso,'N') AS Aviso, isnull(ReciboDeposito,'N') AS ReciboDeposito, isnull(EmpresaPedido,'') as EmpresaPedido, isnull(EndEmpresaPedido,0) as EndEmpresaPedido," & vbCrLf &
                   "			   isnull(Pedido,0) as Pedido, isnull(PedidoFixacao,0) as PedidoFixacao, isnull(Procuracao,0) as Procuracao, ValorDoDocumento, Descontos, Deducoes, Juros, Acrescimos, ValorLiquido, isnull(MoedaValorDoDocumento,0) as MoedaValorDoDocumento," & vbCrLf &
                   "			   isnull(MoedaDescontos,0) as MoedaDescontos, isnull(MoedaDeducoes,0) as MoedaDeducoes, isnull(MoedaJuros,0) as MoedaJuros, isnull(MoedaAcrescimos,0) as MoedaAcrescimos, isnull(MoedaValorLiquido,0) as MoedaValorLiquido, Historico, CodigoDeBarras, CodigoDigitado, isnull(CodigoDeBarraPreImpresso,0) AS CodigoDeBarraPreImpresso, Destinatario," & vbCrLf &
                   "			   EndDestinatario, isnull(NomeDoDestinatario,'') as NomeDoDestinatario, isnull(Destinacao,'') as Destinacao, solicitacao, UsuarioInclusao, isnull(UsuarioInclusaoData,Movimento) as UsuarioInclusaoData, isnull(UsuarioAlteracao,'') as UsuarioAlteracao, UsuarioAlteracaoData," & vbCrLf &
                   "			   isnull(UsuarioCancelamento,'') as UsuarioCancelamento, UsuarioCancelamentoData, isnull(UsuarioLiberacao,'') as UsuarioLiberacao, UsuarioLiberacaoData, isnull(UsuarioBaixa,'') as UsuarioBaixa, UsuarioBaixaData, isnull(Grupado,'N') AS Grupado," & vbCrLf &
                   "			   isnull(RegistroMestre,0) as RegistroMestre, convert(varchar(4000),Observacoes) as Observacoes, isnull(SituacaoBancaria,0) as SituacaoBancaria, isnull(NumeroDoCheque,0) as NumeroDoCheque, isnull(Adiantamento,0) as Adiantamento," & vbCrLf &
                   "               VencimentoAdto, isnull(TaxaAdto,0) as TaxaAdto, isnull(UsuarioLiberacaoBloqueio,'') as UsuarioLiberacaoBloqueio," & vbCrLf &
                   "			   UsuarioLiberacaoBloqueioDate, isnull(UsuarioLiberacaoPedido,'') as UsuarioLiberacaoPedido, UsuarioLiberacaoPedidoDate, isnull(UsuarioLiberacaoCheque,'') as UsuarioLiberacaoCheque," & vbCrLf &
                   "			   UsuarioLiberacaoChequeDate, isnull(ContratoBancario,'') AS ContratoBancario, isnull(CarteiraAdto,'') AS CarteiraAdto, isnull(CarteiraDoTitulo,0) AS CarteiraDoTitulo, " & vbCrLf &
                   "               isnull(ContratoDeFinanciamento,'') AS ContratoDeFinanciamento, DataEnvio, isnull(Ocorrencia,0) AS Ocorrencia, isnull(motivo,'') AS motivo, isnull(DigitoNossoNumero,'') AS DigitoNossoNumero, " & vbCrLf &
                   "               isnull(NossoNumero,0) AS NossoNumero, isnull(Timbrado,'') AS Timbrado, isnull(ModalidadeDePagamento,0) AS ModalidadeDePagamento, CreditoNaConta, isnull(NumeroDaRemessa,0) AS NumeroDaRemessa, " & vbCrLf &
                   "               isnull(SituacaoRemessaBancaria,0) AS SituacaoRemessaBancaria, isnull(TipoContaCliente,'') AS TipoContaCliente, isnull(TituloOrigem,0) AS TituloOrigem, isnull(ValorRecompra,0) AS ValorRecompra, " & vbCrLf &
                   "               isnull(TipoContaPagadora,'') AS TipoContaPagadora, isnull(Processo,'') AS Processo, isnull(BoletoBancario,0) AS BoletoBancario, isnull(UsuarioBoletoBancario,'') as UsuarioBoletoBancario, UsuarioBoletoBancarioDate, " & vbCrLf &
                   "               isnull(HistoricoRemessa,'') AS HistoricoRemessa, isnull(ObservacoesControleInterno,'') AS ObservacoesControleInterno, isnull(FolhaDePagamento,0) AS FolhaDePagamento, isnull(UsuarioFolhaDePagamento,'') as UsuarioFolhaDePagamento, UsuarioFolhaDePagamentoDate " & vbCrLf &
                   "		  FROM MovimentacoesFinanceiras" & vbCrLf &
                   "        ) sb" & vbCrLf &
                   " WHERE sb.Registro_Id = " & pTitulo

        If Not String.IsNullOrWhiteSpace(pReceberPagar) Then
            sql &= " AND sb.ReceberPagar = '" & pReceberPagar & "'"
        End If

        Dim Banco As New AcessaBanco
        Dim ds As DataSet = Banco.ConsultaDataSet(sql, "Titulo")

        For Each row As DataRow In ds.Tables(0).Rows
            _ReceberPagar = row("ReceberPagar")
            _Codigo = row("Registro_Id")
            _Sequencia = row("Sequencia_Id")

            _DataMoeda = row("DataMoeda")
            If Not IsDBNull(row("Baixa")) Then _Baixa = row("Baixa")

            _CodigoIndexador = row("Indexador")
            _CodigoMoeda = row("Moeda")
            _Movimento = row("Movimento")
            _Vencimento = row("Vencimento")
            _Prorrogacao = row("Prorrogacao")
            _CodigoEmpresaPedido = row("EmpresaPedido")
            _EndEmpresaPedido = row("EndEmpresaPedido")

            _codigoPedidoFixacao = row("PedidoFixacao")
            _CodigoPedido = row("Pedido")

            _CodigoProvisao = row("Provisao")
            _CodigoCarteira = row("Carteira")
            _Tributo = row("Tributo")
            _CodigoTipoPagto = row("TipoPagto")
            _CodigoSituacao = row("Situacao")
            _Lote = row("Lote")
            _CodigoEmpresa = row("Empresa")
            _EndEmpresa = row("EndEmpresa")
            _CodigoUnidadeDeNegocio = row("UnidadeDeNegocio")
            _CodigoCliente = row("Cliente")
            _EndCliente = row("EndCliente")
            _CodigoBancoCliente = row("BancoCliente")
            _CodigoAgenciaCliente = row("AgenciaCliente")
            _DigitoAgenciaCliente = row("DigitoAgenciaCliente")
            _ContaCliente = row("ContaCliente")
            _DigitoContaCliente = row("DigitoContaCliente")
            _ContaContabilCliente = row("ContaContabilCliente")
            _CodigoEmpresaPagadora = row("EmpresaPagadora")
            _EndEmpresaPagadora = row("EndEmpresaPagadora")
            _CodigoBancoPagador = row("BancoPagador")
            _codigoAgenciaPagadora = row("AgenciaPagadora")
            _DigitoAgenciaPagadora = row("DigitoAgenciaPagadora")
            _ContaPagadora = row("ContaPagadora")
            _DigitoContaPagadora = row("DigitoContaPagadora")
            _ContaContabilPagadora = row("ContaContabilPagadora")
            _Cheque = row("Cheque") = "S"
            _Slips = row("Slips") = "S"
            _Recibo = row("Recibo") = "S"
            _Aviso = row("Aviso") = "S"
            _ReciboDeposito = row("ReciboDeposito") = "S"
            _CodigoProcuracao = row("Procuracao")
            _ValorDoDocumento = row("ValorDoDocumento")
            _Descontos = row("Descontos")
            _Deducoes = row("Deducoes")
            _Juros = row("Juros")
            _Acrescimos = row("Acrescimos")
            _ValorLiquido = row("ValorLiquido")
            _MoedaValorDoDocumento = row("MoedaValorDoDocumento")
            _MoedaDescontos = row("MoedaDescontos")
            _MoedaDeducoes = row("MoedaDeducoes")
            _MoedaJuros = row("MoedaJuros")
            _MoedaAcrescimos = row("MoedaAcrescimos")
            _MoedaValorLiquido = row("MoedaValorLiquido")
            _Historico = row("Historico")
            If Not IsDBNull(row("CodigoDeBarras")) Then _CodigoDeBarras = row("CodigoDeBarras")
            If Not IsDBNull(row("CodigoDigitado")) Then _CodigoDigitado = row("CodigoDigitado") = "S"
            _CodigoDeBarrasPreImpresso = row("CodigoDeBarraPreImpresso")
            _CodigoDestinatario = row("Destinatario")
            _EndDestinatario = row("EndDestinatario")
            _NomeDoDestinatario = row("NomeDoDestinatario")
            _Destinacao = row("Destinacao")
            _Solicitacao = row("solicitacao")
            _Agrupado = Conversoes.GetValueFromDescription(Of eAgrupamentoFinanceiro)(row("Grupado"))
            _RegistroMestre = row("RegistroMestre")
            If Not IsDBNull(row("Observacoes")) Then _Observacoes = row("Observacoes")
            _SituacaoBancaria = row("SituacaoBancaria")
            _NumeroDoCheque = row("NumeroDoCheque")
            _CodigoAdiantamento = row("Adiantamento")
            If Not IsDBNull(row("VencimentoAdto")) Then _VencimentoAdto = row("VencimentoAdto")
            _TaxaAdto = row("TaxaAdto")
            _CarteiraAdto = row("CarteiraAdto")
            _UsuarioLiberacaoBloqueio = row("UsuarioLiberacaoBloqueio")
            If Not IsDBNull(row("UsuarioLiberacaoBloqueioDate")) Then _UsuarioLiberacaoBloqueioDate = row("UsuarioLiberacaoBloqueioDate")
            _UsuarioLiberacaoPedido = row("UsuarioLiberacaoPedido")
            If Not IsDBNull(row("UsuarioLiberacaoPedidoDate")) Then _UsuarioLiberacaoPedidoDate = row("UsuarioLiberacaoPedidoDate")
            _UsuarioLiberacaoCheque = row("UsuarioLiberacaoCheque")
            If Not IsDBNull(row("UsuarioLiberacaoChequeDate")) Then _UsuarioLiberacaoChequeDate = row("UsuarioLiberacaoChequeDate")
            _ContratoBancario = row("ContratoBancario")
            _UsuarioInclusao = row("UsuarioInclusao")
            _UsuarioInclusaoData = row("UsuarioInclusaoData")
            _UsuarioAlteracao = row("UsuarioAlteracao")
            If Not IsDBNull(row("UsuarioAlteracaoData")) Then _UsuarioAlteracaoData = row("UsuarioAlteracaoData")
            _UsuarioCancelamento = row("UsuarioCancelamento")
            If Not IsDBNull(row("UsuarioCancelamentoData")) Then _UsuarioCancelamentoData = row("UsuarioCancelamentoData")
            _UsuarioLiberacao = row("UsuarioLiberacao")
            If Not IsDBNull(row("UsuarioLiberacaoData")) Then _UsuarioLiberacaoData = row("UsuarioLiberacaoData")
            _UsuarioBaixa = row("UsuarioBaixa")
            If Not IsDBNull(row("UsuarioBaixaData")) Then _UsuarioBaixaData = row("UsuarioBaixaData")
            '_UsarSequencia = True
            _CodigoCarteiraDoTitulo = row("CarteiraDoTitulo")

            _ContratoDeFinanciamento = row("ContratoDeFinanciamento")
            If Not IsDBNull(row("DataEnvio")) Then _DataEnvio = row("DataEnvio")
            _CodigoOcorrencia = row("Ocorrencia")
            _Motivo = row("motivo")
            _DigitoNossoNumero = row("DigitoNossoNumero")
            _NossoNumero = row("NossoNumero")
            _Timbrado = row("Timbrado")
            _ModalidadeDePagamento = row("ModalidadeDePagamento")
            If Not IsDBNull(row("CreditoNaConta")) Then _CreditoNaConta = row("CreditoNaConta")
            _NumeroDaRemessa = row("NumeroDaRemessa")
            _SituacaoRemessaBancaria = row("SituacaoRemessaBancaria")
            _TipoContaCliente = row("TipoContaCliente")
            _CodigoTituloOrigem = row("TituloOrigem")
            _ValorRecompra = row("ValorRecompra")
            _TipoContaPagadora = row("TipoContaPagadora")
            _Processo = row("Processo")

            _BoletoBancario = row("BoletoBancario")
            _UsuarioBoletoBancario = row("UsuarioBoletoBancario")
            If Not IsDBNull(row("UsuarioBoletoBancarioDate")) Then _UsuarioBoletoBancarioDate = row("UsuarioBoletoBancarioDate")

            _HistoricoRemessa = row("HistoricoRemessa")

            _FolhaDePagamento = row("FolhaDePagamento")
            _UsuarioFolhaDePagamento = row("UsuarioFolhaDePagamento")
            If Not IsDBNull(row("UsuarioFolhaDePagamentoDate")) Then _UsuarioFolhaDePagamentoDate = row("UsuarioFolhaDePagamentoDate")

            _ObservacoesControleInterno = row("ObservacoesControleInterno")
        Next
    End Sub
#End Region

#Region "Fields"

    Private _IUD As String

    Private _ReceberPagar As String
    Private _Codigo As Integer 'Registro_Id
    Private _Sequencia As Integer 'Sequencia_Id

    Private _CodigoProvisao As Integer 'Provisao
    Private _Provisao As TipoProvisao
    Private _DescricaoProvisao As String = ""

    Private _CodigoCarteira As String
    Private _Carteira As CarteiraFinanceira 'Carteira

    Private _Tributo As String 'Tributo

    Private _CodigoIndexador As Integer
    Private _Indexador As Indexador 'Indexador
    Private _CodigoMoeda As Integer 'Moeda
    Private _Moeda As Moeda
    Private _DescricaoMoeda As String = ""
    Private _DataMoeda As Date 'DataMoeda

    Private _CodigoTipoPagto As Integer 'TipoPagto
    Private _TipoPagto As TipoDePagamento

    Private _CodigoSituacao As Integer 'Situacao
    Private _Situacao As Situacao
    Private _DescSituacao As String

    Private _Lote As Integer 'Lote
    Private _Movimento As Date 'Movimento
    Private _Vencimento As Date 'Vencimento
    Private _Prorrogacao As Date 'Prorrogacao

    Private _Baixa As Date 'Baixa

    Private _CodigoUnidadeDeNegocio As String 'UnidadeDeNegocio
    Private _UnidadeDeNegocio As Cliente 'UnidadeDeNegocio

    Private _CodigoEmpresa As String
    Private _EndEmpresa As Integer 'EndEmpresa
    Private _Empresa As Cliente 'Empresa

    Private _CodigoCliente As String 'Cliente
    Private _EndCliente As Integer 'EndCliente
    Private _Cliente As Cliente 'Cliente

    '*********  Banco Cliente  *********
    Private _CodigoBancoCliente As Integer 'BancoCliente
    Private _BancoCliente As Banco
    Private _CodigoAgenciaCliente As String 'AgenciaCliente
    Private _DigitoAgenciaCliente As String 'DigitoAgenciaCliente
    Private _AgenciaCliente As Agencia
    Private _ContaCliente As String 'ContaCliente
    Private _DigitoContaCliente As String 'DigitoContaCliente
    Private _ContaContabilCliente As String 'ContaContabilCliente

    Private _CodigoEmpresaPagadora As String 'EmpresaPagadora
    Private _EndEmpresaPagadora As Integer 'EndEmpresaPagadora
    Private _EmpresaPagadora As Cliente 'EmpresaPagadora
    Private _UnidadeDeNegocioEmpresaPagadora As GruposXEmpresas 'UnidadeDeNegocio EmpresaPagadora

    Private _CodigoBancoPagador As Integer 'BancoPagador
    Private _BancoPagador As Banco
    Private _codigoAgenciaPagadora As String 'AgenciaPagadora
    Private _DigitoAgenciaPagadora As String 'DigitoAgenciaPagadora
    Private _AgenciaPagadora As Agencia
    Private _ContaPagadora As String 'ContaPagadora
    Private _DigitoContaPagadora As String 'DigitoContaPagadora
    Private _ContaContabilPagadora As String 'ContaContabilPagadora

    Private _Cheque As Boolean 'Cheque
    Private _Slips As Boolean 'Slips
    Private _Recibo As Boolean 'Recibo
    Private _Aviso As Boolean 'Aviso
    Private _ReciboDeposito As Boolean 'ReciboDeposito

    Private _CodigoEmpresaPedido As String
    Private _EndEmpresaPedido As Integer
    Private _EmpresaPedido As Cliente

    Private _CodigoPedido As Integer
    Private _Pedido As Pedido
    Private _codigoPedidoFixacao As Integer
    Private _PedidoFixacao As Pedido

    Private _CodigoProcuracao As Integer
    Private _Procuracao As Procuracao

    '**************************************
    '**********  Valores  *****************
    '**************************************
    Private _ValorDoDocumento As Decimal
    Private _Descontos As Decimal
    Private _Deducoes As Decimal
    Private _Juros As Decimal
    Private _Acrescimos As Decimal
    Private _ValorLiquido As Decimal

    Private _MoedaValorDoDocumento As Decimal
    Private _MoedaDescontos As Decimal
    Private _MoedaDeducoes As Decimal
    Private _MoedaJuros As Decimal
    Private _MoedaAcrescimos As Decimal
    Private _MoedaValorLiquido As Decimal

    Private _Historico As String
    Private _CodigoDeBarras As String
    Private _CodigoDigitado As Boolean

    Private _CodigoDestinatario As String
    Private _EndDestinatario As Integer
    Private _Destinatario As Cliente

    Private _NomeDoDestinatario As String
    Private _Destinacao As String
    Private _Solicitacao As Integer

    Private _UsuarioInclusao As String
    Private _UsuarioInclusaoData As Date
    Private _UsuarioAlteracao As String
    Private _UsuarioAlteracaoData As Date
    Private _UsuarioCancelamento As String
    Private _UsuarioCancelamentoData As Date
    Private _UsuarioLiberacao As String
    Private _UsuarioLiberacaoData As Date
    Private _UsuarioBaixa As String
    Private _UsuarioBaixaData As Date

    Private _Agrupado As eAgrupamentoFinanceiro
    Private _RegistroMestre As Integer
    Private _Observacoes As String
    Private _SituacaoBancaria As Integer

    Private _TaxaAdto As Decimal
    Private _VencimentoAdto As Date
    Private _CodigoAdiantamento As Integer
    Private _CarteiraAdto As String

    Private _NumeroDoCheque As Integer

    Private _UsuarioLiberacaoBloqueio As String
    Private _UsuarioLiberacaoBloqueioDate As Date
    Private _UsuarioLiberacaoPedido As String
    Private _UsuarioLiberacaoPedidoDate As Date
    Private _UsuarioLiberacaoCheque As String
    Private _UsuarioLiberacaoChequeDate As Date

    Private _CodigoCarteiraDoTitulo As Integer 'Código Carteira do Título

    Private _ContratoDeFinanciamento As String 'Contrato de Financiamento
    Private _DataEnvio As Date 'Data de envio
    Private _CodigoOcorrencia As Integer 'Código de Ocorrência
    Private _Motivo As String 'Motivo
    Private _DigitoNossoNumero As String 'Dígito Nosso Número
    Private _NossoNumero As String 'Nosso Número
    Private _Timbrado As String 'Timbrado
    Private _ModalidadeDePagamento As Integer 'Modalidade de Pagamento
    Private _CreditoNaConta As Date 'Crédito na Conta
    Private _NumeroDaRemessa As Long 'Numero da Remessa
    Private _SituacaoRemessaBancaria As Integer 'Situação da Remessa Bancária
    Private _TipoContaCliente As String 'Tipo Conta Cliente

    Private _ContratoBancario As String 'ContratoBancario  
    Private _CodigoDeBarrasPreImpresso As Boolean   'Código de Barras Pré Impresso

    Private _CodigoTituloOrigem As Integer 'Título Origem
    Private _ValorRecompra As Decimal 'Valor Recompra
    Private _TipoContaPagadora As String 'Tipo da Conta Pagadora

    Private _IndiceFixo As Boolean = False
    Private _IndiceTitulo As Decimal

    Private _Processo As String

    Private _BoletoBancario As Boolean = False
    Private _UsuarioBoletoBancario As String
    Private _UsuarioBoletoBancarioDate As Date

    Private _HistoricoRemessa As String

    Private _FolhaDePagamento As Boolean = False
    Private _UsuarioFolhaDePagamento As String
    Private _UsuarioFolhaDePagamentoDate As Date

    Private _ObservacoesControleInterno As String

    'Private _UsarSequencia As Boolean = False
    Private _ListHistoricoFinanceiro As ListHistoricoFinanceiro = New ListHistoricoFinanceiro()

    Private _Quantidade As Decimal

    Private _Valores As Novo.ListTituloXContaContabil
    Private _CodigoContaContabilCliFor As String = ""
    Private _ContaContabilCliFor As PlanoDeConta

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

    'Public Property UsarSequencia() As Boolean
    '    Get
    '        Return _UsarSequencia
    '    End Get
    '    Set(ByVal value As Boolean)
    '        _UsarSequencia = value
    '    End Set
    'End Property

    Public Property Carregando() As Boolean
        Get
            Return _Carregando
        End Get
        Set(ByVal value As Boolean)
            _Carregando = value
        End Set
    End Property

    Public Property ReceberPagar() As String
        Get
            Return _ReceberPagar
        End Get
        Set(ByVal value As String)
            _ReceberPagar = value
        End Set
    End Property

    Public Property Codigo() As Integer
        Get
            Return _Codigo
        End Get
        Set(ByVal value As Integer)
            _Codigo = value
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

    Public Property CodigoProvisao() As Integer
        Get
            Return _CodigoProvisao
        End Get
        Set(ByVal value As Integer)
            _CodigoProvisao = value
            _Provisao = Nothing
        End Set
    End Property

    Public Property Provisao() As TipoProvisao
        Get
            If _Provisao Is Nothing And _CodigoProvisao > 0 Then
                _Provisao = New TipoProvisao(_CodigoProvisao)
                _DescricaoProvisao = _Provisao.Descricao
            End If
            Return _Provisao
        End Get
        Set(ByVal value As TipoProvisao)
            _Provisao = value
        End Set
    End Property

    Public ReadOnly Property DescricaoProvisao() As String
        Get
            If _CodigoProvisao > 0 And _DescricaoProvisao = "" Then Return Provisao.Descricao
            Return _DescricaoProvisao
        End Get
    End Property

    Public Property CodigoCarteira() As String
        Get
            Return _CodigoCarteira
        End Get
        Set(ByVal value As String)
            _CodigoCarteira = value
            _Carteira = Nothing
        End Set
    End Property

    Public Property Carteira() As CarteiraFinanceira
        Get
            If _Carteira Is Nothing And _CodigoCarteira.Length > 0 Then _Carteira = New CarteiraFinanceira(_CodigoCarteira)
            Return _Carteira
        End Get
        Set(ByVal value As CarteiraFinanceira)
            _Carteira = value
        End Set
    End Property

    Public Property Tributo() As String
        Get
            Return _Tributo
        End Get
        Set(ByVal value As String)
            _Tributo = value
        End Set
    End Property

    Public Property CodigoIndexador() As Integer
        Get
            Return _CodigoIndexador
        End Get
        Set(ByVal value As Integer)
            _CodigoIndexador = value
            If Not Me.IndiceFixo Then
                _IndiceTitulo = New Cotacao(_CodigoIndexador, Me.DataMoeda).Indice
            End If
        End Set
    End Property

    Public Property Indexador() As Indexador
        Get
            If _Indexador Is Nothing And _CodigoIndexador > 0 Then _Indexador = New Indexador(_CodigoIndexador)
            Return _Indexador
        End Get
        Set(ByVal value As Indexador)
            _Indexador = value
        End Set
    End Property

    Public Property CodigoMoeda() As Integer
        Get
            Return _CodigoMoeda
        End Get
        Set(ByVal value As Integer)
            _CodigoMoeda = value
        End Set
    End Property

    Public Property Moeda() As Moeda
        Get
            If _Moeda Is Nothing And _CodigoMoeda > 0 Then
                _Moeda = New Moeda(_CodigoMoeda)
                _DescricaoMoeda = _Moeda.Descricao
            End If
            Return _Moeda
        End Get
        Set(ByVal value As Moeda)
            _Moeda = value
        End Set
    End Property

    Public ReadOnly Property DescricaoMoeda() As String
        Get
            If _CodigoMoeda > 0 And _DescricaoMoeda = "" Then Return Moeda.Descricao
            Return _DescricaoMoeda
        End Get
    End Property

    Public Property DataMoeda() As Date
        Get
            Return _DataMoeda
        End Get
        Set(ByVal value As Date)
            _DataMoeda = value
            If Not _IndiceFixo And Me.CodigoIndexador > 0 Then
                _IndiceTitulo = New Cotacao(Me.CodigoIndexador, _DataMoeda).Indice
            End If
        End Set
    End Property

    Public Property IndiceFixo() As Boolean
        Get
            Return _IndiceFixo
        End Get
        Set(value As Boolean)
            _IndiceFixo = value
        End Set
    End Property

    Public Property IndiceTitulo() As Decimal
        Get
            Return _IndiceTitulo
        End Get
        Set(ByVal value As Decimal)
            _IndiceTitulo = value
        End Set
    End Property

    Public Property CodigoTipoPgto() As Integer
        Get
            Return _CodigoTipoPagto
        End Get
        Set(ByVal value As Integer)
            _CodigoTipoPagto = value
            _TipoPagto = Nothing
        End Set
    End Property

    Public Property TipoPagto() As TipoDePagamento
        Get
            If _TipoPagto Is Nothing And _CodigoTipoPagto > 0 Then _TipoPagto = New TipoDePagamento(_CodigoTipoPagto)
            Return _TipoPagto
        End Get
        Set(ByVal value As TipoDePagamento)
            _TipoPagto = value
        End Set
    End Property

    Public Property CodigoSituacao() As Integer
        Get
            Return _CodigoSituacao
        End Get
        Set(ByVal value As Integer)
            _CodigoSituacao = value
            _Situacao = Nothing
        End Set
    End Property

    Public Property Situacao() As Situacao
        Get
            If _Situacao Is Nothing And _CodigoSituacao > 0 Then _Situacao = New Situacao(_CodigoSituacao)
            Return _Situacao
        End Get
        Set(ByVal value As Situacao)
            _Situacao = value
        End Set
    End Property

    Public ReadOnly Property DescSituacao() As String
        Get
            Return Situacao.Codigo & "-" & Situacao.Descricao
        End Get
    End Property


    Public Property Lote() As Integer
        Get
            Return _Lote
        End Get
        Set(ByVal value As Integer)
            _Lote = value
        End Set
    End Property

    Public Property Movimento() As Date
        Get
            Return _Movimento
        End Get
        Set(ByVal value As Date)
            _Movimento = value
        End Set
    End Property

    Public Property Vencimento() As Date
        Get
            Return _Vencimento
        End Get
        Set(ByVal value As Date)
            _Vencimento = value
            If Not Carregando Then
                If Moeda.Classificacao = eTiposMoeda.Oficial Then
                    AtualizaLiquidoOficial()
                Else
                    AtualizaLiquidoMoeda()
                End If
            End If

        End Set
    End Property

    Public Property Prorrogacao() As Date
        Get
            Return _Prorrogacao
        End Get
        Set(ByVal value As Date)
            _Prorrogacao = value
        End Set
    End Property

    Public Property Baixa() As Date
        Get
            Return _Baixa
        End Get
        Set(ByVal value As Date)
            _Baixa = value
        End Set
    End Property

    Public Property CodigoEmpresa() As String
        Get
            Return _CodigoEmpresa
        End Get
        Set(ByVal value As String)
            _CodigoEmpresa = value
            _Empresa = Nothing
        End Set
    End Property

    Public Property EnderecoEmpresa() As Integer
        Get
            Return _EndEmpresa
        End Get
        Set(ByVal value As Integer)
            _EndEmpresa = value
            _Empresa = Nothing
        End Set
    End Property

    Public Property Empresa() As Cliente
        Get
            If _Empresa Is Nothing And _CodigoEmpresa.Length > 0 Then _Empresa = New Cliente(_CodigoEmpresa, _EndEmpresa)
            Return _Empresa
        End Get
        Set(ByVal value As Cliente)
            _Empresa = value
        End Set
    End Property

    Public Property ContratoBancario() As String
        Get
            Return _ContratoBancario
        End Get
        Set(value As String)
            _ContratoBancario = value
        End Set
    End Property

    Public Property CodigoUnidadeDeNegocio() As String
        Get
            Return _CodigoUnidadeDeNegocio
        End Get
        Set(ByVal value As String)
            _CodigoUnidadeDeNegocio = value
            _UnidadeDeNegocio = Nothing
        End Set
    End Property

    Public Property UnidadeDeNegocio() As Cliente
        Get
            If _UnidadeDeNegocio Is Nothing And _CodigoUnidadeDeNegocio > 0 Then _UnidadeDeNegocio = New Cliente(_CodigoUnidadeDeNegocio, 0)
            Return _UnidadeDeNegocio
        End Get
        Set(ByVal value As Cliente)
            _UnidadeDeNegocio = value
        End Set
    End Property

    Public Property CodigoCliente() As String
        Get
            Return _CodigoCliente
        End Get
        Set(ByVal value As String)
            _CodigoCliente = value
            _Cliente = Nothing
        End Set
    End Property

    Public Property EndCliente() As Integer
        Get
            Return _EndCliente
        End Get
        Set(ByVal value As Integer)
            _EndCliente = value
        End Set
    End Property

    Public Property Cliente() As Cliente
        Get
            If _Cliente Is Nothing And _CodigoCliente.Length > 0 Then _Cliente = New Cliente(_CodigoCliente, _EndCliente)
            Return _Cliente
        End Get
        Set(ByVal value As Cliente)
            _Cliente = value
        End Set
    End Property

    '******* Dados Bancarios ************
    Public Property CodigoBancoCliente() As Integer
        Get
            Return _CodigoBancoCliente
        End Get
        Set(ByVal value As Integer)
            _CodigoBancoCliente = value
            _BancoCliente = Nothing
            _AgenciaCliente = Nothing
        End Set
    End Property

    Public ReadOnly Property BancoCliente() As Banco
        Get
            If _BancoCliente Is Nothing And _CodigoBancoCliente > 0 Then _BancoCliente = New Banco(_CodigoBancoCliente)
            Return _BancoCliente
        End Get
    End Property

    Public Property CodigoAgenciaCliente() As String
        Get
            Return _CodigoAgenciaCliente
        End Get
        Set(ByVal value As String)
            _CodigoAgenciaCliente = value
            _AgenciaCliente = Nothing
        End Set
    End Property

    Public Property DigitoAgenciaCliente() As String
        Get
            Return _DigitoAgenciaCliente
        End Get
        Set(ByVal value As String)
            _DigitoAgenciaCliente = value
            _AgenciaCliente = Nothing
        End Set
    End Property

    Public ReadOnly Property AgenciaCliente() As Agencia
        Get
            If _CodigoBancoCliente > 0 And _CodigoAgenciaCliente.Length > 0 And DigitoAgenciaCliente.Length > 0 And _AgenciaCliente Is Nothing Then _AgenciaCliente = New Agencia(_CodigoBancoCliente, _CodigoAgenciaCliente, _DigitoAgenciaCliente)
            Return _AgenciaCliente
        End Get
    End Property

    Public Property ContaCliente() As String
        Get
            Return _ContaCliente
        End Get
        Set(ByVal value As String)
            _ContaCliente = value
        End Set
    End Property

    Public Property DigitoContaCliente() As String
        Get
            Return _DigitoContaCliente
        End Get
        Set(ByVal value As String)
            _DigitoContaCliente = value
        End Set
    End Property

    Public Property ContaContabilCliente() As String
        Get
            Return _ContaContabilCliente
        End Get
        Set(ByVal value As String)
            _ContaContabilCliente = value
        End Set
    End Property
    '************************************

    Public Property CodigoEmpresaPagadora() As String
        Get
            Return _CodigoEmpresaPagadora
        End Get
        Set(ByVal value As String)
            _CodigoEmpresaPagadora = value
            _EmpresaPagadora = Nothing
            _UnidadeDeNegocioEmpresaPagadora = Nothing
        End Set
    End Property

    Public Property EndEmpresaPagadora() As Integer
        Get
            Return _EndEmpresaPagadora
        End Get
        Set(ByVal value As Integer)
            _EndEmpresaPagadora = value
            _EmpresaPagadora = Nothing
            _UnidadeDeNegocioEmpresaPagadora = Nothing
        End Set
    End Property

    Public Property EmpresaPagadora() As Cliente
        Get
            If _EmpresaPagadora Is Nothing And _CodigoEmpresaPagadora.Length > 0 Then _EmpresaPagadora = New Cliente(_CodigoEmpresaPagadora, _EndEmpresaPagadora)
            Return _EmpresaPagadora
        End Get
        Set(ByVal value As Cliente)
            _EmpresaPagadora = value
        End Set
    End Property

    Public ReadOnly Property UnidadeDeNegocioEmpresaPagadora As GruposXEmpresas
        Get
            If _UnidadeDeNegocioEmpresaPagadora Is Nothing And _CodigoEmpresaPagadora.Length > 0 Then _UnidadeDeNegocioEmpresaPagadora = New GruposXEmpresas(_CodigoEmpresaPagadora, _EndEmpresaPagadora)
            Return _UnidadeDeNegocioEmpresaPagadora
        End Get
    End Property

    Public Property CodigoBancoPagador() As Integer
        Get
            Return _CodigoBancoPagador
        End Get
        Set(ByVal value As Integer)
            _CodigoBancoPagador = value
            _BancoPagador = Nothing
            _AgenciaPagadora = Nothing
        End Set
    End Property

    Public ReadOnly Property BancoPagador() As Banco
        Get
            If _BancoPagador Is Nothing And _CodigoBancoPagador > 0 Then _BancoPagador = New Banco(_CodigoBancoCliente)
            Return _BancoPagador
        End Get
    End Property

    Public Property CodigoAgenciaPagadora() As String
        Get
            Return _codigoAgenciaPagadora
        End Get
        Set(ByVal value As String)
            _codigoAgenciaPagadora = value
            _AgenciaPagadora = Nothing
        End Set
    End Property

    Public Property DigitoAgenciaPagadora() As String
        Get
            Return _DigitoAgenciaPagadora
        End Get
        Set(ByVal value As String)
            _DigitoAgenciaPagadora = value
            _AgenciaPagadora = Nothing
        End Set
    End Property

    Public ReadOnly Property AgenciaPagadora() As Agencia
        Get
            If _CodigoBancoPagador > 0 And _codigoAgenciaPagadora.Length > 0 And DigitoAgenciaPagadora.Length > 0 And _AgenciaPagadora Is Nothing Then _AgenciaPagadora = New Agencia(_CodigoBancoPagador, _codigoAgenciaPagadora, _DigitoAgenciaPagadora)
            Return _AgenciaPagadora
        End Get
    End Property

    Public Property ContaPagadora() As String
        Get
            Return _ContaPagadora
        End Get
        Set(ByVal value As String)
            _ContaPagadora = value
        End Set
    End Property

    Public Property DigitoContaPagadora() As String
        Get
            Return _DigitoContaPagadora
        End Get
        Set(ByVal value As String)
            _DigitoContaPagadora = value
        End Set
    End Property

    Public Property ContaContabilPagadora() As String
        Get
            Return _ContaContabilPagadora
        End Get
        Set(ByVal value As String)
            _ContaContabilPagadora = value
        End Set
    End Property

    Public Property Cheque() As Boolean
        Get
            Return _Cheque
        End Get
        Set(ByVal value As Boolean)
            _Cheque = value
        End Set
    End Property

    Public Property Slips() As Boolean
        Get
            Return _Slips
        End Get
        Set(ByVal value As Boolean)
            _Slips = value
        End Set
    End Property

    Public Property Recibo() As Boolean
        Get
            Return _Recibo
        End Get
        Set(ByVal value As Boolean)
            _Recibo = value
        End Set
    End Property

    Public Property Aviso() As Boolean
        Get
            Return _Aviso
        End Get
        Set(ByVal value As Boolean)
            _Aviso = value
        End Set
    End Property

    Public Property ReciboDeposito() As Boolean
        Get
            Return _ReciboDeposito
        End Get
        Set(ByVal value As Boolean)
            _ReciboDeposito = value
        End Set
    End Property

    Public Property CodigoEmpresaPedido() As String
        Get
            Return _CodigoEmpresaPedido
        End Get
        Set(ByVal value As String)
            _CodigoEmpresaPedido = value
            _EmpresaPedido = Nothing
        End Set
    End Property

    Public Property EndEmpresaPedido() As Integer
        Get
            Return _EndEmpresaPedido
        End Get
        Set(ByVal value As Integer)
            _EndEmpresaPedido = value
            _EmpresaPedido = Nothing
        End Set
    End Property

    Public Property EmpresaPedido() As Cliente
        Get
            If _EmpresaPedido Is Nothing And _CodigoEmpresaPedido.Length > 0 Then _EmpresaPedido = New Cliente(_CodigoEmpresaPedido, _EndEmpresaPedido)
            Return _EmpresaPedido
        End Get
        Set(ByVal value As Cliente)
            _EmpresaPedido = value
        End Set
    End Property

    Public Property CodigoPedido() As Integer
        Get
            Return _CodigoPedido
        End Get
        Set(ByVal value As Integer)
            _CodigoPedido = value
            _Pedido = Nothing

            'Edson 25/11/2015
            'If value > 0 Then
            '    _Pedido = New Pedido(_CodigoEmpresaPedido, _EndEmpresaPedido, _CodigoPedido)
            '    If _Pedido.Moeda Is Nothing OrElse _Pedido.Moeda.Classificacao = eTiposMoeda.Oficial Then
            '        _IndiceFixo = True
            '        _IndiceTitulo = 0
            '    ElseIf _Pedido.IndiceFixado > 0 Then
            '        _IndiceFixo = True
            '        _IndiceTitulo = _Pedido.IndiceFixado
            '    Else
            '        _IndiceFixo = False
            '        _IndiceTitulo = New Cotacao(_Pedido.CodigoIndexador, _DataMoeda).Indice
            '    End If
            'Else
            '    _IndiceFixo = True
            '    _IndiceTitulo = New Cotacao(Me.CodigoIndexador, Me.DataMoeda).Indice
            'End If
        End Set
    End Property

    Public Property Pedido() As Pedido
        Get
            If _Pedido Is Nothing And Me.CodigoPedido > 0 Then _Pedido = New Pedido(Me.CodigoEmpresaPedido, Me.EndEmpresaPedido, Me.CodigoPedido)
            Return _Pedido
        End Get
        Set(ByVal value As Pedido)
            _Pedido = value
        End Set
    End Property

    Public Property CodigoPedidoFixacao() As Integer
        Get
            Return _codigoPedidoFixacao
        End Get
        Set(ByVal value As Integer)
            _codigoPedidoFixacao = value
            _PedidoFixacao = Nothing
        End Set
    End Property

    Public Property PedidoFixacao() As Pedido
        Get
            If _PedidoFixacao Is Nothing And _CodigoEmpresaPedido.Length > 0 And _CodigoPedido > 0 Then _PedidoFixacao = New Pedido(_CodigoEmpresaPedido, _EndEmpresaPedido, _CodigoEmpresaPedido)
            Return _Pedido
        End Get
        Set(ByVal value As Pedido)
            _PedidoFixacao = value
        End Set
    End Property

    Public Property CodigoProcuracao() As Integer
        Get
            Return _CodigoProcuracao
        End Get
        Set(ByVal value As Integer)
            _CodigoProcuracao = value
            _Procuracao = Nothing
        End Set
    End Property

    Public Property Procuracao() As Procuracao
        Get
            If _Procuracao Is Nothing And _CodigoProcuracao > 0 Then _Procuracao = New Procuracao(_CodigoEmpresa, _EndEmpresa, _CodigoProcuracao)
            Return _Procuracao
        End Get
        Set(ByVal value As Procuracao)
            _Procuracao = value
        End Set
    End Property

    '***************** Valores Oficiais *****************
    Public Property ValorDoDocumento() As Decimal
        Get
            Return _ValorDoDocumento
        End Get
        Set(ByVal value As Decimal)
            'NÃO ESTAVA RECALCULANDO O VALOR DO TÍTULO SE FOSSE DÓLAR - FURLAN - 28/05/2024
            'If Moeda.Classificacao = eTiposMoeda.Oficial Then
            _ValorDoDocumento = value
            If Not Carregando Then AtualizaLiquidoOficial()
            'End If
        End Set
    End Property

    Public Property Descontos() As Decimal
        Get
            Return _Descontos
        End Get
        Set(ByVal value As Decimal)
            If Moeda.Classificacao = eTiposMoeda.Oficial Then
                _Descontos = value
                If Not Carregando Then AtualizaLiquidoOficial()
            End If
        End Set
    End Property

    Public Property Deducoes() As Decimal
        Get
            Return _Deducoes
        End Get
        Set(ByVal value As Decimal)
            If Moeda.Classificacao = eTiposMoeda.Oficial Then
                _Deducoes = value
                If Not Carregando Then AtualizaLiquidoOficial()
            End If
        End Set
    End Property

    Public Property Juros() As Decimal
        Get
            Return _Juros
        End Get
        Set(ByVal value As Decimal)
            If Moeda.Classificacao = eTiposMoeda.Oficial Then
                _Juros = value
                If Not Carregando Then AtualizaLiquidoOficial()
            End If
        End Set
    End Property

    Public Property Acrescimos() As Decimal
        Get
            Return _Acrescimos
        End Get
        Set(ByVal value As Decimal)
            If Moeda.Classificacao = eTiposMoeda.Oficial Then
                _Acrescimos = value
                If Not Carregando Then AtualizaLiquidoOficial()
            End If
        End Set
    End Property

    Public ReadOnly Property ValorLiquido() As Decimal
        Get
            Return _ValorLiquido
        End Get
    End Property

    '**************** Valores Moeda *********************
    Public Property MoedaValorDoDocumento() As Decimal
        Get
            Return _MoedaValorDoDocumento
        End Get
        Set(ByVal value As Decimal)
            'USEI O MESMO CRITÉRIO DO VALOR EM REAIS, DEVE ATUALIZAR - FURLAN - 28/05/2025
            'If Moeda.Classificacao = eTiposMoeda.MoedaEstrangeira Then
            _MoedaValorDoDocumento = value
            If Not Carregando Then AtualizaLiquidoMoeda()
            'End If
        End Set
    End Property

    Public Property MoedaDescontos() As Decimal
        Get
            Return _MoedaDescontos
        End Get
        Set(ByVal value As Decimal)
            If Moeda.Classificacao = eTiposMoeda.MoedaEstrangeira Then
                _MoedaDescontos = value
                If Not Carregando Then AtualizaLiquidoMoeda()
            End If
        End Set
    End Property

    Public Property MoedaDeducoes() As Decimal
        Get
            Return _MoedaDeducoes
        End Get
        Set(ByVal value As Decimal)
            If Moeda.Classificacao = eTiposMoeda.MoedaEstrangeira Then
                _MoedaDeducoes = value
                If Not Carregando Then AtualizaLiquidoMoeda()
            End If
        End Set
    End Property

    Public Property MoedaJuros() As Decimal
        Get
            Return _MoedaJuros
        End Get
        Set(ByVal value As Decimal)
            If Moeda.Classificacao = eTiposMoeda.MoedaEstrangeira Then
                _MoedaJuros = value
                If Not Carregando Then AtualizaLiquidoMoeda()
            End If
        End Set
    End Property

    Public Property MoedaAcrescimos() As Decimal
        Get
            Return _MoedaAcrescimos
        End Get
        Set(ByVal value As Decimal)
            If Moeda.Classificacao = eTiposMoeda.MoedaEstrangeira Then
                _MoedaAcrescimos = value
                If Not Carregando Then AtualizaLiquidoMoeda()
            End If
        End Set
    End Property

    Public ReadOnly Property MoedaValorLiquido() As Decimal
        Get
            Return _MoedaValorLiquido
        End Get
    End Property
    '****************************************************

    Public Property Historico() As String
        Get
            Return _Historico
        End Get
        Set(ByVal value As String)
            _Historico = value
        End Set
    End Property

    Public Property CodigoDeBarras() As String
        Get
            Return _CodigoDeBarras
        End Get
        Set(ByVal value As String)
            _CodigoDeBarras = value
        End Set
    End Property

    Public Property CodigoDigitado() As Boolean
        Get
            Return _CodigoDigitado
        End Get
        Set(ByVal value As Boolean)
            _CodigoDigitado = value
        End Set
    End Property

    Public Property CodigoDeBarrasPreImpresso() As Boolean
        Get
            Return _CodigoDeBarrasPreImpresso
        End Get
        Set(ByVal value As Boolean)
            _CodigoDeBarrasPreImpresso = value
        End Set
    End Property

    Public Property CodigoDestinatario() As String
        Get
            Return _CodigoDestinatario
        End Get
        Set(ByVal value As String)
            _CodigoDestinatario = value
        End Set
    End Property

    Public Property EndDestinatario() As Integer
        Get
            Return _EndDestinatario
        End Get
        Set(ByVal value As Integer)
            _EndDestinatario = value
        End Set
    End Property

    Public Property Destinatario() As Cliente
        Get
            If _Destinatario Is Nothing And _CodigoDestinatario.Length > 0 Then _Destinatario = New Cliente(_CodigoDestinatario, _EndDestinatario)
            Return _Destinatario
        End Get
        Set(ByVal value As Cliente)
            _Destinatario = value
        End Set
    End Property

    '************ Pra Que **********
    Public Property NomeDoDestinatario() As String
        Get
            Return _NomeDoDestinatario
        End Get
        Set(ByVal value As String)
            _NomeDoDestinatario = value
        End Set
    End Property

    Public Property Destinacao() As String
        Get
            Return _Destinacao
        End Get
        Set(ByVal value As String)
            _Destinacao = value
        End Set
    End Property

    Public Property Solicitacao() As Integer
        Get
            Return _Solicitacao
        End Get
        Set(ByVal value As Integer)
            _Solicitacao = value
        End Set
    End Property
    '*******************************

    Public Property Agrupado() As eAgrupamentoFinanceiro
        Get
            Return _Agrupado
        End Get
        Set(ByVal value As eAgrupamentoFinanceiro)
            _Agrupado = value
        End Set
    End Property

    Public Property RegistroMestre() As Integer
        Get
            Return _RegistroMestre
        End Get
        Set(ByVal value As Integer)
            _RegistroMestre = value
        End Set
    End Property

    Public Property Observacoes() As String
        Get
            Return _Observacoes
        End Get
        Set(ByVal value As String)
            _Observacoes = value
        End Set
    End Property

    Public Property SituacaoBancaria() As Integer
        Get
            Return _SituacaoBancaria
        End Get
        Set(ByVal value As Integer)
            _SituacaoBancaria = value
        End Set
    End Property

    Public Property NumeroDoCheque() As Integer
        Get
            Return _NumeroDoCheque
        End Get
        Set(ByVal value As Integer)
            _NumeroDoCheque = value
        End Set
    End Property

    '***********  Pra Que  ****************
    Public Property CodigoAdiantamento() As Integer
        Get
            Return _CodigoAdiantamento
        End Get
        Set(ByVal value As Integer)
            _CodigoAdiantamento = value
        End Set
    End Property

    Public Property VencimentoAdto() As Date
        Get
            Return _VencimentoAdto
        End Get
        Set(ByVal value As Date)
            _VencimentoAdto = value
        End Set
    End Property

    Public Property TaxaAdto() As Decimal
        Get
            Return _TaxaAdto
        End Get
        Set(ByVal value As Decimal)
            _TaxaAdto = value
        End Set
    End Property
    '**************************************
    Public Property CarteiraAdto() As String
        Get
            Return _CarteiraAdto
        End Get
        Set(ByVal value As String)
            _CarteiraAdto = value
        End Set
    End Property

    Public Property UsuarioLiberacaoBloqueio() As String
        Get
            Return _UsuarioLiberacaoBloqueio
        End Get
        Set(ByVal value As String)
            _UsuarioLiberacaoBloqueio = value
        End Set
    End Property

    Public Property UsuarioLiberacaoBloqueioDate() As Date
        Get
            Return _UsuarioLiberacaoBloqueioDate
        End Get
        Set(ByVal value As Date)
            _UsuarioLiberacaoBloqueioDate = value
        End Set
    End Property

    Public Property UsuarioLiberacaoPedido() As String
        Get
            Return _UsuarioLiberacaoPedido
        End Get
        Set(ByVal value As String)
            _UsuarioLiberacaoPedido = value
        End Set
    End Property

    Public Property UsuarioLiberacaoPedidoDate() As Date
        Get
            Return _UsuarioLiberacaoPedidoDate
        End Get
        Set(ByVal value As Date)
            _UsuarioLiberacaoPedidoDate = value
        End Set
    End Property

    Public Property UsuarioLiberacaoCheque() As String
        Get
            Return _UsuarioLiberacaoCheque
        End Get
        Set(ByVal value As String)
            _UsuarioLiberacaoCheque = value
        End Set
    End Property

    Public Property UsuarioLiberacaoChequeDate() As Date
        Get
            Return _UsuarioLiberacaoChequeDate
        End Get
        Set(ByVal value As Date)
            _UsuarioLiberacaoChequeDate = value
        End Set
    End Property

    Public Property UsuarioInclusao() As String
        Get
            Return _UsuarioInclusao
        End Get
        Set(ByVal value As String)
            _UsuarioInclusao = value
        End Set
    End Property

    Public Property UsuarioInclusaoData() As Date
        Get
            Return _UsuarioInclusaoData
        End Get
        Set(ByVal value As Date)
            _UsuarioInclusaoData = value
        End Set
    End Property

    Public Property UsuarioAlteracao() As String
        Get
            Return _UsuarioAlteracao
        End Get
        Set(ByVal value As String)
            _UsuarioAlteracao = value
        End Set
    End Property

    Public Property UsuarioAlteracaoData() As Date
        Get
            Return _UsuarioAlteracaoData
        End Get
        Set(ByVal value As Date)
            _UsuarioAlteracaoData = value
        End Set
    End Property

    Public Property UsuarioCancelamento() As String
        Get
            Return _UsuarioCancelamento
        End Get
        Set(ByVal value As String)
            _UsuarioCancelamento = value
        End Set
    End Property

    Public Property UsuarioCancelamentoData() As Date
        Get
            Return _UsuarioCancelamentoData
        End Get
        Set(ByVal value As Date)
            _UsuarioCancelamentoData = value
        End Set
    End Property

    Public Property UsuarioLiberacao() As String
        Get
            Return _UsuarioLiberacao
        End Get
        Set(ByVal value As String)
            _UsuarioLiberacao = value
        End Set
    End Property

    Public Property UsuarioLiberacaoData() As Date
        Get
            Return _UsuarioLiberacaoData
        End Get
        Set(ByVal value As Date)
            _UsuarioLiberacaoData = value
        End Set
    End Property

    Public Property UsuarioBaixa() As String
        Get
            Return _UsuarioBaixa
        End Get
        Set(ByVal value As String)
            _UsuarioBaixa = value
        End Set
    End Property

    Public Property UsuarioBaixaData() As Date
        Get
            Return _UsuarioBaixaData
        End Get
        Set(ByVal value As Date)
            _UsuarioBaixaData = value
        End Set
    End Property

    Public Property CodigoCarteiraDoTitulo() As Integer
        Get
            Return _CodigoCarteiraDoTitulo
        End Get
        Set(ByVal value As Integer)
            _CodigoCarteiraDoTitulo = value
        End Set
    End Property

    Public Property ContratoDeFinanciamento() As String
        Get
            Return _ContratoDeFinanciamento
        End Get
        Set(ByVal value As String)
            _ContratoDeFinanciamento = value
        End Set
    End Property

    Public Property DataEnvio() As Date
        Get
            Return _DataEnvio
        End Get
        Set(ByVal value As Date)
            _DataEnvio = value
        End Set
    End Property

    Public Property CodigoOcorrencia() As Integer
        Get
            Return _CodigoOcorrencia
        End Get
        Set(ByVal value As Integer)
            _CodigoOcorrencia = value
        End Set
    End Property

    Public Property Motivo() As String
        Get
            Return _Motivo
        End Get
        Set(ByVal value As String)
            _Motivo = value
        End Set
    End Property

    Public Property DigitoNossoNumero() As String
        Get
            Return _DigitoNossoNumero
        End Get
        Set(ByVal value As String)
            _DigitoNossoNumero = value
        End Set
    End Property

    Public Property NossoNumero() As String
        Get
            Return _NossoNumero
        End Get
        Set(ByVal value As String)
            _NossoNumero = value
        End Set
    End Property

    Public Property Timbrado() As String
        Get
            Return _Timbrado
        End Get
        Set(ByVal value As String)
            _Timbrado = value
        End Set
    End Property

    Public Property ModalidadeDePagamento() As Integer
        Get
            Return _ModalidadeDePagamento
        End Get
        Set(ByVal value As Integer)
            _ModalidadeDePagamento = value
        End Set
    End Property

    Public Property CreditoNaConta() As Date
        Get
            Return _CreditoNaConta
        End Get
        Set(ByVal value As Date)
            _CreditoNaConta = value
        End Set
    End Property

    Public Property NumeroDaRemessa() As Long
        Get
            Return _NumeroDaRemessa
        End Get
        Set(ByVal value As Long)
            _NumeroDaRemessa = value
        End Set
    End Property

    Public Property SituacaoRemessaBancaria() As Integer
        Get
            Return _SituacaoRemessaBancaria
        End Get
        Set(ByVal value As Integer)
            _SituacaoRemessaBancaria = value
        End Set
    End Property

    Public Property TipoContaCliente() As String
        Get
            Return _TipoContaCliente
        End Get
        Set(ByVal value As String)
            _TipoContaCliente = value
        End Set
    End Property

    Public Property CodigoTituloOrigem() As Integer
        Get
            Return _CodigoTituloOrigem
        End Get
        Set(ByVal value As Integer)
            _CodigoTituloOrigem = value
        End Set
    End Property

    Public Property ValorRecompra() As Decimal
        Get
            Return _ValorRecompra
        End Get
        Set(ByVal value As Decimal)
            _ValorRecompra = value
        End Set
    End Property

    Public Property TipoContaPagadora() As String
        Get
            Return _TipoContaPagadora
        End Get
        Set(ByVal value As String)
            _TipoContaPagadora = value
        End Set
    End Property

    Public Property Processo() As String
        Get
            Return _Processo
        End Get
        Set(ByVal value As String)
            _Processo = value
        End Set
    End Property

    Public Property BoletoBancario() As Boolean
        Get
            Return _BoletoBancario
        End Get
        Set(value As Boolean)
            _BoletoBancario = value
        End Set
    End Property

    Public Property UsuarioBoletoBancario() As String
        Get
            Return _UsuarioBoletoBancario
        End Get
        Set(ByVal value As String)
            _UsuarioBoletoBancario = value
        End Set
    End Property

    Public Property UsuarioBoletoBancarioDate() As Date
        Get
            Return _UsuarioBoletoBancarioDate
        End Get
        Set(ByVal value As Date)
            _UsuarioBoletoBancarioDate = value
        End Set
    End Property

    Public Property HistoricoRemessa() As String
        Get
            Return _HistoricoRemessa
        End Get
        Set(ByVal value As String)
            _HistoricoRemessa = value
        End Set
    End Property

    Public Property UsuarioFolhaDePagamento() As String
        Get
            Return _UsuarioFolhaDePagamento
        End Get
        Set(ByVal value As String)
            _UsuarioFolhaDePagamento = value
        End Set
    End Property

    Public Property UsuarioFolhaDePagamentoDate() As Date
        Get
            Return _UsuarioFolhaDePagamentoDate
        End Get
        Set(ByVal value As Date)
            _UsuarioFolhaDePagamentoDate = value
        End Set
    End Property

    Public Property ObservacoesControleInterno() As String
        Get
            Return _ObservacoesControleInterno
        End Get
        Set(ByVal value As String)
            _ObservacoesControleInterno = value
        End Set
    End Property

    Public Property ListHistoricoFinanceiro() As ListHistoricoFinanceiro
        Get
            Return _ListHistoricoFinanceiro
        End Get
        Set(ByVal value As ListHistoricoFinanceiro)
            _ListHistoricoFinanceiro = value
        End Set
    End Property

    Public Property FolhaDePagamento() As Boolean
        Get
            Return _FolhaDePagamento
        End Get
        Set(value As Boolean)
            _FolhaDePagamento = value
        End Set
    End Property

    Public ReadOnly Property isAdiantamento As Boolean
        Get
            If ReceberPagar = "R" And ContaContabilCliente.Substring(0, 1) = "1" Then Return True
            If ReceberPagar = "P" And ContaContabilCliente.Substring(0, 1) = "2" Then Return True
            Return False
        End Get
    End Property

    Public ReadOnly Property isBaixaAdiantamento As Boolean
        Get

            If CodigoProvisao = 3 Then

                If ReceberPagar = "R" And ContaContabilCliente.Substring(0, 1) = "1" Then Return True
                If ReceberPagar = "P" And ContaContabilCliente.Substring(0, 1) = "2" Then Return True

                If ContaContabilPagadora.Length > 0 Then

                    If ReceberPagar = "R" And ContaContabilPagadora.Substring(0, 1) = "1" Then Return True
                    If ReceberPagar = "P" And ContaContabilPagadora.Substring(0, 1) = "2" Then Return True

                End If

            End If

            Return False
        End Get
    End Property

    Public Property Quantidade() As Decimal
        Get
            Return _Quantidade
        End Get
        Set(ByVal value As Decimal)
            _Quantidade = value
        End Set
    End Property

    Public Property Valores() As Novo.ListTituloXContaContabil
        Get
            If _Valores Is Nothing And _CodigoContaContabilCliFor.Length > 0 Then
                _Valores = New Novo.ListTituloXContaContabil(Me)
            End If
            Return _Valores
        End Get
        Set(ByVal value As Novo.ListTituloXContaContabil)
            _Valores = value
        End Set
    End Property

    Public Property CodigoContaContabilCliFor() As String
        Get
            Return _CodigoContaContabilCliFor
        End Get
        Set(ByVal value As String)
            _CodigoContaContabilCliFor = value
            _ContaContabilCliFor = Nothing

            If Me.Controlando Then
                Dim Prod As New Novo.TituloXContaContabil(Me)
                Prod.CodigoContaEncargo = ContaContabilCliFor.Conta
                Prod.ContaEncargo = ContaContabilCliFor
                Prod.Descricao = ContaContabilCliFor.Titulo
                Prod.DC = IIf(Me.ReceberPagar = "R", "C", "D")

                If Not Me.Valores.EncargoValorDocumento Is Nothing Then
                    Prod.ValorMoeda = Me.Valores.EncargoValorDocumento.ValorMoeda
                    Prod.ValorOficial = Me.Valores.EncargoValorDocumento.ValorOficial
                    Me.Valores.Remove(Me.Valores.EncargoValorDocumento)
                End If

                Me.Valores.Add(Prod)
                Me.Valores.EncargoValorDocumento = Prod
            End If
        End Set
    End Property

    Public Property ContaContabilCliFor() As PlanoDeConta
        Get
            If _ContaContabilCliFor Is Nothing And _CodigoContaContabilCliFor.Length > 0 Then _ContaContabilCliFor = New PlanoDeConta("", 0, _CodigoContaContabilCliFor)
            Return _ContaContabilCliFor
        End Get
        Set(ByVal value As PlanoDeConta)
            _ContaContabilCliFor = value
        End Set
    End Property

    Public ReadOnly Property Controlando() As Boolean
        Get
            Return _Controlando
        End Get
    End Property
    Public Property MsgControle As String
        Get
            Return _MsgControle
        End Get
        Set(value As String)
            _MsgControle = value
        End Set
    End Property

#End Region

#Region "Methods"

    Public Function Copy() As Titulo
        Return DirectCast(Me.MemberwiseClone(), Titulo)
    End Function

    Public Sub CalculaJuros(ByVal ate As Date)
        Dim dias As Double
        dias = ate.Subtract(Vencimento).Days
        If dias <= 0 Then Exit Sub

        Dim iq As Decimal 'taxa que eu quero 
        Dim it As Decimal 'taxa que eu tenho 
        Dim q As Decimal  'número de dias em que quero expressar a taxa 
        Dim t As Decimal  'número de dias em que a taxa que tenho está expressa 


        'iq = {[(1+(it/100))^(q/t)]-1}*100 
        'Então, se temos it=6 ao mês (30 dias) e quero iq para 1 dia:
        'it = 6% a.m.     q = 1     t = 30     
        'iq = {[(1+(6/100))^(1/30)]-1}*100 = [(1,06^0,0333)-1]*100 = (1,001944-1)*100 = 0,1944% ao dia
        If CodigoPedido > 0 Then
            it = IIf(TaxaAdto > 0, TaxaAdto, IIf(Pedido.Taxa > 0, Pedido.Taxa, New Safra(Pedido.CodigoSafra).Taxa))
        Else
            it = TaxaAdto
        End If

        If it = 0 Then Exit Sub
        q = 1
        t = 30
        iq = (((1 + (it / 100)) ^ (q / t)) - 1) * 100
        'Vf = vi(1 + i) ^ n
        If Moeda.Classificacao = eTiposMoeda.Oficial Then
            Juros = Math.Abs(_ValorLiquido - (_ValorLiquido * (1 + iq / 100) ^ dias))
        Else
            MoedaJuros = Math.Abs(_MoedaValorLiquido - (_MoedaValorLiquido * (1 + iq / 100) ^ dias))
        End If
    End Sub

    Private Sub AtualizaLiquidoOficial()
        _ValorLiquido = Me.ValorDoDocumento + Me.Juros + Me.Acrescimos - Me.Descontos - Me.Deducoes
        If _IndiceTitulo = 0 Then Exit Sub
        _MoedaValorDoDocumento = Math.Round(Me.ValorDoDocumento / Me.IndiceTitulo, 2)
        _MoedaJuros = Math.Round(Me.Juros / Me.IndiceTitulo, 2)
        _MoedaAcrescimos = Math.Round(Me.Acrescimos / Me.IndiceTitulo, 2)
        _MoedaDescontos = Math.Round(Me.Descontos / Me.IndiceTitulo, 2)
        _MoedaDeducoes = Math.Round(Me.Deducoes / Me.IndiceTitulo, 2)
        _MoedaValorLiquido = Me.MoedaValorDoDocumento + Me.MoedaJuros + Me.MoedaAcrescimos - Me.MoedaDescontos - Me.MoedaDeducoes
    End Sub

    Private Sub AtualizaLiquidoMoeda()
        _MoedaValorLiquido = Me.MoedaValorDoDocumento + Me.MoedaJuros + Me.MoedaAcrescimos - Me.MoedaDescontos - Me.MoedaDeducoes
        If _IndiceTitulo = 0 Then Exit Sub
        _ValorDoDocumento = Math.Round(Me.MoedaValorDoDocumento * Me.IndiceTitulo, 2)
        _Juros = Math.Round(Me.MoedaJuros * Me.IndiceTitulo, 2)
        _Acrescimos = Math.Round(Me.MoedaAcrescimos * Me.IndiceTitulo, 2)
        _Descontos = Math.Round(Me.MoedaDescontos * Me.IndiceTitulo, 2)
        _Deducoes = Math.Round(Me.MoedaDeducoes * Me.IndiceTitulo, 2)
        _ValorLiquido = Me.ValorDoDocumento + Me.Juros + Me.Acrescimos - Me.Descontos - Me.Deducoes
    End Sub

    Public Function DolarizaBaixa(ByVal pData As String, ByVal Valor As String, ByVal Indexador As String) As Decimal
        Dim SqlL As String
        Dim indice As Decimal

        SqlL = "SELECT Indice" &
               "  FROM Cotacoes" &
               " WHERE Data_Id      ='" & pData & "'" & vbCrLf &
               "   AND Indexador_Id =" & Indexador
        ' ''"   AND Realizado    ='R'
        ''acrescentar  Realizado = 'R' no AND

        Dim Banco As New AcessaBanco
        For Each Dr As DataRow In Banco.ConsultaDataSet(SqlL, "Cot").Tables(0).Rows
            indice = Dr("Indice")
        Next

        Return indice
    End Function

    Public Sub DesligarControles()
        _Controlando = False
    End Sub

    Public Sub LigarControles()
        _Controlando = True
    End Sub

    Public Function Salvar() As Boolean
        Dim Banco As New AcessaBanco
        Dim sqls As New ArrayList
        SalvarSql(sqls)
        Return Banco.GravaBanco(sqls)
    End Function

    Public Function SalvarSql(ByRef sqls As ArrayList, Optional ByVal UsaNumerador As Boolean = True) As ArrayList
        Dim strSQL As String
        Dim ObjBanco As New AcessaBanco

        'Recupera o Numero do Titulo e atualiza o mesmo na tabela 
        If UsaNumerador And Me.IUD = "I" Then
            Dim numTitulo As New Numerador(1)
            _Codigo = numTitulo.Sequencia + 1
            sqls.Add(numTitulo.IncrementarNumeradorSql(1))
        End If

        DeletaContabilizacao(sqls)

        Select Case Me.IUD
            Case "I"
                strSQL = "INSERT INTO "

                If _ReceberPagar = "M" Then
                    strSQL &= "MovimentacoesFinanceiras" & vbCrLf
                ElseIf _ReceberPagar = "R" Then
                    strSQL &= "ContasAReceber" & vbCrLf
                Else
                    strSQL &= "ContasAPagar" & vbCrLf
                End If

                strSQL &= "  (Registro_Id, Sequencia_Id, Provisao, Carteira, Tributo, Indexador, Moeda, TipoPagto, Situacao, Lote, Movimento, Vencimento, Prorrogacao, DataMoeda, Baixa, " & vbCrLf &
                         "   UnidadeDeNegocio, Empresa, EndEmpresa, Cliente, EndCliente, BancoCliente, AgenciaCliente, DigitoAgenciaCliente, ContaCliente, DigitoContaCliente, " & vbCrLf &
                         "   ContaContabilCliente, EmpresaPagadora, EndEmpresaPagadora, BancoPagador, AgenciaPagadora, DigitoAgenciaPagadora, ContaPagadora, " & vbCrLf &
                         "   DigitoContaPagadora, ContaContabilPagadora, Cheque, Slips, Recibo, Aviso, ReciboDeposito, EmpresaPedido, EndEmpresaPedido, Pedido, PedidoFixacao, " & vbCrLf &
                         "   Procuracao, ValorDoDocumento, Descontos, Deducoes, Juros, Acrescimos, ValorLiquido, MoedaValorDoDocumento, MoedaDescontos, MoedaDeducoes, " & vbCrLf &
                         "   MoedaJuros, MoedaAcrescimos, MoedaValorLiquido, Historico, CodigoDeBarras, CodigoDigitado, CodigoDeBarraPreImpresso, Destinatario, EndDestinatario, NomeDoDestinatario, Destinacao, " & vbCrLf &
                         "   solicitacao, UsuarioInclusao, UsuarioInclusaoData, UsuarioAlteracao, UsuarioAlteracaoData, UsuarioCancelamento, UsuarioCancelamentoData, " & vbCrLf &
                         "   UsuarioLiberacao, UsuarioLiberacaoData, UsuarioBaixa, UsuarioBaixaData, Grupado, RegistroMestre, Observacoes, SituacaoBancaria, NumeroDoCheque, " & vbCrLf &
                         "   Adiantamento, VencimentoAdto, TaxaAdto, UsuarioLiberacaoBloqueio, UsuarioLiberacaoBloqueioDate, UsuarioLiberacaoPedido, UsuarioLiberacaoPedidoDate, " & vbCrLf &
                         "   UsuarioLiberacaoCheque, UsuarioLiberacaoChequeDate, ContratoBancario, CarteiraAdto, CarteiraDoTitulo, ContratoDeFinanciamento, DataEnvio, Ocorrencia, motivo, " & vbCrLf &
                         "   DigitoNossoNumero, NossoNumero, Timbrado, ModalidadeDePagamento, CreditoNaConta, NumeroDaRemessa, SituacaoRemessaBancaria, TipoContaCliente, TituloOrigem, " & vbCrLf &
                         "   ValorRecompra, TipoContaPagadora, Processo, BoletoBancario, UsuarioBoletoBancario, UsuarioBoletoBancarioDate, HistoricoRemessa, ObservacoesControleInterno, " & vbCrLf &
                         "   FolhaDePagamento, UsuarioFolhaDePagamento, UsuarioFolhaDePagamentoDate) " & vbCrLf &
                         " VALUES (" & _Codigo & "," &
                         _Sequencia & ", " &
                         _CodigoProvisao & ",'" &
                         _CodigoCarteira & "','" &
                         _Tributo & "'," &
                         _CodigoIndexador & "," &
                         _CodigoMoeda & ", " &
                         _CodigoTipoPagto & ", " &
                         _CodigoSituacao & "," &
                         _Lote & ",'" &
                         _Movimento.ToString("yyyy-MM-dd") & "','" &
                         _Vencimento.ToString("yyyy-MM-dd") & "','" &
                         _Prorrogacao.ToString("yyyy-MM-dd") & "','" &
                         _DataMoeda.ToString("yyyy-MM-dd") & "',"

                If _Baixa = Nothing Then
                    strSQL &= "NULL,"
                Else
                    strSQL &= "'" & _Baixa.ToString("yyyy-MM-dd") & "',"
                End If

                strSQL &= "'" & _CodigoUnidadeDeNegocio & "','" &
                                _CodigoEmpresa & "'," &
                                _EndEmpresa & ",'" &
                                _CodigoCliente & "'," &
                                _EndCliente & "," &
                                _CodigoBancoCliente & ",'" &
                                _CodigoAgenciaCliente & "','" &
                                _DigitoAgenciaCliente & "','" &
                                _ContaCliente & "','" &
                                _DigitoContaCliente & "'," & "'" &
                                _ContaContabilCliente & "','" &
                                _CodigoEmpresaPagadora & "'," &
                                _EndEmpresaPagadora & "," &
                                _CodigoBancoPagador & ",'" &
                                _codigoAgenciaPagadora & "','" &
                                _DigitoAgenciaPagadora & "','" &
                                _ContaPagadora & "'," & "'" &
                                _DigitoContaPagadora & "','" &
                                _ContaContabilPagadora & "','" &
                                IIf(_Cheque, "S", "N") & "','" &
                                IIf(_Slips, "S", "N") & "','" &
                                IIf(_Recibo, "S", "N") & "','" &
                                IIf(_Aviso, "S", "N") & "','" &
                                IIf(_ReciboDeposito, "S", "N") & "',"

                If _CodigoPedido = 0 Then
                    strSQL &= "NULL,NULL,NULL,"
                Else
                    strSQL &= "'" & _CodigoEmpresaPedido & "'," &
                              _EndEmpresaPedido & "," &
                              _CodigoPedido & ","
                End If

                strSQL &= _codigoPedidoFixacao.ToSqlNULL & "," &
                                _CodigoProcuracao.ToSqlNULL & "," &
                                Str(ValorDoDocumento) & "," &
                                Str(_Descontos) & "," &
                                Str(_Deducoes) & "," &
                                Str(_Juros) & "," &
                                Str(_Acrescimos) & "," &
                                Str(_ValorLiquido) & "," &
                                Str(_MoedaValorDoDocumento) & "," &
                                Str(_MoedaDescontos) & "," &
                                Str(_MoedaDeducoes) & "," &
                                Str(_MoedaJuros) & "," &
                                Str(_MoedaAcrescimos) & "," &
                                Str(_MoedaValorLiquido) & ",'" &
                                _Historico & "',"

                If _CodigoDeBarras IsNot Nothing Then
                    strSQL &= "'" & _CodigoDeBarras.Replace(" ", "").Replace(".", "") & "',"
                Else
                    strSQL &= "'',"
                End If

                strSQL &= IIf(_CodigoDigitado, "'S'", "'N'") & "," &
                          IIf(_CodigoDeBarrasPreImpresso, 1, 0) & ",'" &
                          _CodigoDestinatario & "'," &
                          _EndDestinatario & ",'" &
                          _NomeDoDestinatario & "','" &
                          _Destinacao & "'," &
                          _Solicitacao & ",'" &
                          _UsuarioInclusao & "',"

                If _UsuarioInclusaoData = Nothing Then
                    strSQL &= "NULL,"
                Else
                    strSQL &= "'" & _UsuarioInclusaoData.ToString("yyyy-MM-dd") & "',"
                End If

                strSQL &= "'" & _UsuarioAlteracao & "',"

                If _UsuarioAlteracaoData = Nothing Then
                    strSQL &= "NULL,"
                Else
                    strSQL &= "'" & _UsuarioAlteracaoData.ToString("yyyy-MM-dd") & "',"
                End If

                strSQL &= "'" & _UsuarioCancelamento & "',"

                If _UsuarioCancelamentoData = Nothing Then
                    strSQL &= "NULL,"
                Else
                    strSQL &= "'" & _UsuarioCancelamentoData.ToString("yyyy-MM-dd") & "',"
                End If

                strSQL &= "'" & _UsuarioLiberacao & "',"

                If _UsuarioLiberacaoData = Nothing Then
                    strSQL &= "NULL,"
                Else
                    strSQL &= "'" & _UsuarioLiberacaoData.ToString("yyyy-MM-dd") & "',"
                End If

                strSQL &= "'" & _UsuarioBaixa & "',"

                If _UsuarioBaixaData = Nothing Then
                    strSQL &= "NULL,"
                Else
                    strSQL &= "'" & _UsuarioBaixaData.ToString("yyyy-MM-dd") & "',"
                End If

                strSQL &= "'" & Textos.GetEnumDescription(_Agrupado) & "'," & _RegistroMestre & ",'" & _Observacoes & "'," & _SituacaoBancaria & "," & _NumeroDoCheque & "," & _CodigoAdiantamento & ","

                If _VencimentoAdto = Nothing Then
                    strSQL &= "NULL,"
                Else
                    strSQL &= "'" & _VencimentoAdto.ToString("yyyy-MM-dd") & "',"
                End If

                strSQL &= Str(_TaxaAdto) & ",'" & _UsuarioLiberacaoBloqueio & "',"

                If _UsuarioLiberacaoBloqueioDate = Nothing Then
                    strSQL &= "NULL,"
                Else
                    strSQL &= "'" & _UsuarioLiberacaoBloqueioDate.ToString("yyyy-MM-dd") & "',"
                End If

                strSQL &= "'" & _UsuarioLiberacaoPedido & "',"

                If _UsuarioLiberacaoPedidoDate = Nothing Then
                    strSQL &= "NULL,"
                Else
                    strSQL &= "'" & _UsuarioLiberacaoPedidoDate.ToString("yyyy-MM-dd") & "',"
                End If


                strSQL &= "'" & _UsuarioLiberacaoCheque & "',"


                If _UsuarioLiberacaoChequeDate = Nothing Then
                    strSQL &= "NULL,"
                Else
                    strSQL &= "'" & _UsuarioLiberacaoChequeDate.ToString("yyyy-MM-dd") & "',"
                End If

                strSQL &= "'" & _ContratoBancario & "','" & _CarteiraAdto & "'," & _CodigoCarteiraDoTitulo & ","


                strSQL &= "'" & _ContratoDeFinanciamento & "',"
                If _DataEnvio = Nothing Then
                    strSQL &= "NULL,"
                Else
                    strSQL &= "'" & _DataEnvio.ToString("yyyy-MM-dd") & "',"
                End If
                strSQL &= _CodigoOcorrencia & ","
                strSQL &= "'" & _Motivo & "',"
                strSQL &= "'" & _DigitoNossoNumero & "',"
                If _NossoNumero = Nothing OrElse String.IsNullOrEmpty(_NossoNumero) Then
                    strSQL &= "NULL,"
                Else
                    strSQL &= _NossoNumero & ","
                End If
                strSQL &= "'" & _Timbrado & "',"
                strSQL &= _ModalidadeDePagamento & ","
                If _CreditoNaConta = Nothing Then
                    strSQL &= "NULL,"
                Else
                    strSQL &= "'" & _CreditoNaConta.ToString("yyyy-MM-dd") & "',"
                End If
                strSQL &= _NumeroDaRemessa & ","
                strSQL &= _SituacaoRemessaBancaria & ","
                strSQL &= "'" & _TipoContaCliente & "',"
                strSQL &= _CodigoTituloOrigem & ","
                strSQL &= Str(_ValorRecompra) & ","
                strSQL &= "'" & _TipoContaPagadora & "',"
                strSQL &= "'" & _Processo & "',"

                strSQL &= IIf(_BoletoBancario, 1, 0) & ","
                strSQL &= "'" & _UsuarioBoletoBancario & "',"
                If _UsuarioBoletoBancarioDate = Nothing Then
                    strSQL &= "NULL,"
                Else
                    strSQL &= "'" & _UsuarioBoletoBancarioDate.ToString("yyyy-MM-dd") & "',"
                End If

                strSQL &= "'" & _HistoricoRemessa & "',"

                strSQL &= "'" & _ObservacoesControleInterno & "',"

                strSQL &= IIf(_FolhaDePagamento, 1, 0) & ","
                strSQL &= "'" & _UsuarioFolhaDePagamento & "',"
                If _UsuarioFolhaDePagamentoDate = Nothing Then
                    strSQL &= "NULL)"
                Else
                    strSQL &= "'" & _UsuarioFolhaDePagamentoDate.ToString("yyyy-MM-dd") & "')"
                End If

                sqls.Add(strSQL)

                If _CodigoProvisao = eProvisao.Baixa Then Contabilizar(sqls)

            Case "U"
                strSQL = " UPDATE "

                If _ReceberPagar = "M" Then
                    strSQL &= "MovimentacoesFinanceiras" & vbCrLf
                ElseIf _ReceberPagar = "R" Then
                    strSQL &= "ContasAReceber" & vbCrLf
                Else
                    strSQL &= "ContasAPagar" & vbCrLf
                End If

                strSQL &= " SET  ValorDoDocumento       = " & Str(_ValorDoDocumento) & vbCrLf &
                "  ,Vencimento             ='" & _Vencimento.ToString("yyyy-MM-dd") & "'" & vbCrLf &
                "  ,Provisao               = " & _CodigoProvisao & vbCrLf &
                "  ,Carteira               = " & IIf(_CodigoCarteira.Length > 0, "'" & _CodigoCarteira & "'", "NULL") & vbCrLf &
                "  ,Tributo                ='" & Tributo & "'" & vbCrLf &
                "  ,Indexador              = " & _CodigoIndexador & vbCrLf &
                "  ,Moeda                  = " & CodigoMoeda & vbCrLf &
                "  ,TipoPagto              = " & _CodigoTipoPagto & vbCrLf &
                "  ,Situacao               = " & _CodigoSituacao & vbCrLf &
                "  ,Lote                   = " & _Lote & vbCrLf &
                "  ,Movimento              ='" & _Movimento.ToString("yyyy-MM-dd") & "'" & vbCrLf &
                "  ,Prorrogacao            ='" & _Prorrogacao.ToString("yyyy-MM-dd") & "'" & vbCrLf &
                "  ,DataMoeda              ='" & _DataMoeda.ToString("yyyy-MM-dd") & "'" & vbCrLf &
                "  ,Baixa                  = " & IIf(_Baixa = Nothing, "NULL", "'" & _Baixa.ToString("yyyy-MM-dd") & "'") & vbCrLf &
                "  ,UnidadeDeNegocio       ='" & _CodigoUnidadeDeNegocio & "'" & vbCrLf &
                "  ,Empresa                ='" & _CodigoEmpresa & "'" & vbCrLf &
                "  ,EndEmpresa             = " & _EndEmpresa & vbCrLf &
                "  ,Cliente                ='" & _CodigoCliente & "'" & vbCrLf &
                "  ,EndCliente             = " & _EndCliente & vbCrLf &
                "  ,BancoCliente           = " & _CodigoBancoCliente & vbCrLf &
                "  ,AgenciaCliente         ='" & _CodigoAgenciaCliente & "'" & vbCrLf &
                "  ,DigitoAgenciaCliente   ='" & _DigitoAgenciaCliente & "'" & vbCrLf &
                "  ,ContaCliente           ='" & _ContaCliente & "'" & vbCrLf &
                "  ,DigitoContaCliente     ='" & DigitoContaCliente & "'" & vbCrLf &
                "  ,ContaContabilCliente   ='" & _ContaContabilCliente & "'" & vbCrLf &
                "  ,EmpresaPagadora        ='" & _CodigoEmpresaPagadora & "'" & vbCrLf &
                "  ,EndEmpresaPagadora     = " & _EndEmpresaPagadora & vbCrLf &
                "  ,BancoPagador           = " & _CodigoBancoPagador & vbCrLf &
                "  ,AgenciaPagadora        ='" & _codigoAgenciaPagadora & "'" & vbCrLf &
                "  ,DigitoAgenciaPagadora  ='" & _DigitoAgenciaPagadora & "'" & vbCrLf &
                "  ,ContaPagadora          ='" & _ContaPagadora & "'" & vbCrLf &
                "  ,DigitoContaPagadora    ='" & _DigitoContaPagadora & "'" & vbCrLf &
                "  ,ContaContabilPagadora  ='" & _ContaContabilPagadora & "'" & vbCrLf &
                "  ,Cheque                 ='" & IIf(Cheque, "S", "N") & "'" & vbCrLf &
                "  ,Slips                  ='" & IIf(Slips, "S", "N") & "'" & vbCrLf &
                "  ,Recibo                 ='" & IIf(Recibo, "S", "N") & "'" & vbCrLf &
                "  ,Aviso                  ='" & IIf(Aviso, "S", "N") & "'" & vbCrLf &
                "  ,ReciboDeposito         ='" & IIf(ReciboDeposito, "S", "N") & "'" & vbCrLf

                If _CodigoPedido = 0 Then
                    strSQL &= ",EmpresaPedido = NULL,  " & vbCrLf &
                              "EndEmpresaPedido = NULL, " & vbCrLf &
                              "Pedido = NULL" & vbCrLf
                Else
                    strSQL &= ",EmpresaPedido = '" & _CodigoEmpresaPedido & "'," & vbCrLf &
                              "EndEmpresaPedido = " & _EndEmpresaPedido & "," & vbCrLf &
                              "Pedido = " & _CodigoPedido & vbCrLf
                End If

                strSQL &= " ,PedidoFixacao          = " & _codigoPedidoFixacao.ToSqlNULL & vbCrLf
                strSQL &= "  ,Procuracao             = " & _CodigoProcuracao.ToSqlNULL & vbCrLf
                strSQL &= "  ,Descontos              = " & Str(_Descontos) & vbCrLf
                strSQL &= "  ,Deducoes               = " & Str(_Deducoes) & vbCrLf
                strSQL &= "  ,Juros                  = " & Str(_Juros) & vbCrLf
                strSQL &= "  ,Acrescimos             = " & Str(_Acrescimos) & vbCrLf
                strSQL &= "  ,ValorLiquido           = " & Str(_ValorLiquido) & vbCrLf
                strSQL &= "  ,MoedaValorDoDocumento  = " & Str(_MoedaValorDoDocumento) & vbCrLf
                strSQL &= "  ,MoedaDescontos         = " & Str(_MoedaDescontos) & vbCrLf
                strSQL &= "  ,MoedaDeducoes          = " & Str(_MoedaDeducoes) & vbCrLf
                strSQL &= "  ,MoedaJuros             = " & Str(_MoedaJuros) & vbCrLf
                strSQL &= "  ,MoedaAcrescimos        = " & Str(_MoedaAcrescimos) & vbCrLf
                strSQL &= "  ,MoedaValorLiquido      = " & Str(_MoedaValorLiquido) & vbCrLf
                strSQL &= "  ,Historico              ='" & _Historico & "'" & vbCrLf
                strSQL &= "  ,CodigoDeBarras         ='" & _CodigoDeBarras.Replace(" ", "").Replace(".", "") & "'" & vbCrLf
                strSQL &= "  ,CodigoDigitado         ='" & IIf(_CodigoDigitado, "S", "N") & "'" & vbCrLf
                strSQL &= "  ,CodigoDeBarraPreImpresso = " & IIf(_CodigoDeBarrasPreImpresso, 1, 0) & vbCrLf
                strSQL &= "  ,Destinatario           ='" & _CodigoDestinatario & "'" & vbCrLf
                strSQL &= "  ,EndDestinatario        = " & _EndDestinatario & vbCrLf
                strSQL &= "  ,NomeDoDestinatario     ='" & _NomeDoDestinatario & "'" & vbCrLf
                strSQL &= "  ,Destinacao             ='" & _Destinacao & "'" & vbCrLf
                strSQL &= "  ,solicitacao            = " & _Solicitacao & vbCrLf
                strSQL &= "  ,UsuarioInclusao        ='" & _UsuarioInclusao & "'" & vbCrLf
                strSQL &= "  ,UsuarioInclusaoData    = " & IIf(_UsuarioInclusaoData = Nothing, "NULL", "'" & _UsuarioInclusaoData.ToString("yyyy-MM-dd") & "'") & vbCrLf
                strSQL &= "  ,UsuarioAlteracao       ='" & _UsuarioAlteracao & "'" & vbCrLf
                strSQL &= "  ,UsuarioAlteracaoData   = " & IIf(_UsuarioAlteracaoData = Nothing, "NULL", "'" & _UsuarioAlteracaoData.ToString("yyyy-MM-dd") & "'") & vbCrLf
                strSQL &= "  ,UsuarioCancelamento    ='" & _UsuarioCancelamento & "'" & vbCrLf
                strSQL &= "  ,UsuarioCancelamentoData= " & IIf(_UsuarioCancelamentoData = Nothing, "NULL", "'" & _UsuarioCancelamentoData.ToString("yyyy-MM-dd") & "'") & vbCrLf
                strSQL &= "  ,UsuarioLiberacao       ='" & _UsuarioLiberacao & "'" & vbCrLf
                strSQL &= "  ,UsuarioLiberacaoData   = " & IIf(_UsuarioLiberacaoData = Nothing, "NULL", "'" & _UsuarioLiberacaoData.ToString("yyyy-MM-dd") & "'") & vbCrLf
                strSQL &= "  ,UsuarioBaixa           ='" & _UsuarioBaixa & "'" & vbCrLf
                strSQL &= "  ,UsuarioBaixaData       = " & IIf(_UsuarioBaixaData = Nothing, "NULL", "'" & _UsuarioBaixaData.ToString("yyyy-MM-dd") & "'") & vbCrLf
                strSQL &= "  ,Grupado                ='" & Textos.GetEnumDescription(_Agrupado) & "'" & vbCrLf
                strSQL &= "  ,RegistroMestre         = " & _RegistroMestre & vbCrLf
                strSQL &= "  ,Observacoes            ='" & _Observacoes & "'" & vbCrLf
                strSQL &= "  ,SituacaoBancaria       = " & _SituacaoBancaria & vbCrLf
                strSQL &= "  ,NumeroDoCheque         = " & _NumeroDoCheque & vbCrLf
                strSQL &= "  ,Adiantamento           = " & _CodigoAdiantamento & vbCrLf
                strSQL &= "  ,VencimentoAdto         = " & IIf(_VencimentoAdto = Nothing, "NULL", "'" & _VencimentoAdto.ToString("yyyy-MM-dd") & "'") & vbCrLf
                strSQL &= "  ,TaxaAdto               = " & Str(_TaxaAdto) & vbCrLf
                strSQL &= "  ,CarteiraAdto           = '" & _CarteiraAdto & "'" & vbCrLf
                strSQL &= "  ,UsuarioLiberacaoBloqueio    ='" & _UsuarioLiberacaoBloqueio & "'" & vbCrLf
                strSQL &= "  ,UsuarioLiberacaoBloqueioDate= " & IIf(_UsuarioLiberacaoBloqueioDate = Nothing, "NULL", "'" & _UsuarioLiberacaoBloqueioDate.ToString("yyyy-MM-dd") & "'") & vbCrLf
                strSQL &= "  ,UsuarioLiberacaoPedido      ='" & _UsuarioLiberacaoPedido & "'" & vbCrLf
                strSQL &= "  ,UsuarioLiberacaoPedidoDate  = " & IIf(_UsuarioLiberacaoPedidoDate = Nothing, "NULL", "'" & _UsuarioLiberacaoPedidoDate.ToString("yyyy-MM-dd") & "'") & vbCrLf
                strSQL &= "  ,UsuarioLiberacaoCheque      ='" & _UsuarioLiberacaoCheque & "'" & vbCrLf
                strSQL &= "  ,UsuarioLiberacaoChequeDate  = " & IIf(_UsuarioLiberacaoChequeDate = Nothing, "NULL", "'" & _UsuarioLiberacaoChequeDate.ToString("yyyy-MM-dd") & "'") & vbCrLf
                strSQL &= "  ,ContratoBancario            = '" & _ContratoBancario & "'" & vbCrLf
                strSQL &= "  ,CarteiraDoTitulo            = " & _CodigoCarteiraDoTitulo & vbCrLf
                strSQL &= "  ,ContratoDeFinanciamento     = '" & _ContratoDeFinanciamento & "'" & vbCrLf
                strSQL &= "  ,DataEnvio                   = " & IIf(_DataEnvio = Nothing, "NULL", "'" & _DataEnvio.ToString("yyyy-MM-dd") & "'") & vbCrLf
                strSQL &= "  ,Ocorrencia                  = " & _CodigoOcorrencia & vbCrLf
                strSQL &= "  ,motivo                      = '" & _Motivo & "'" & vbCrLf
                strSQL &= "  ,DigitoNossoNumero           = '" & _DigitoNossoNumero & "'" & vbCrLf
                strSQL &= "  ,NossoNumero                 = " & IIf(_NossoNumero = Nothing OrElse String.IsNullOrEmpty(_NossoNumero), "NULL", "" & _NossoNumero & "") & vbCrLf
                strSQL &= "  ,Timbrado                    = '" & _Timbrado & "'" & vbCrLf
                strSQL &= "  ,ModalidadeDePagamento       = " & _ModalidadeDePagamento & vbCrLf
                strSQL &= "  ,CreditoNaConta              = " & IIf(_CreditoNaConta = Nothing, "NULL", "'" & _CreditoNaConta.ToString("yyyy-MM-dd") & "'") & vbCrLf
                strSQL &= "  ,NumeroDaRemessa             = " & _NumeroDaRemessa & vbCrLf
                strSQL &= "  ,SituacaoRemessaBancaria     = " & _SituacaoRemessaBancaria & vbCrLf
                strSQL &= "  ,TipoContaCliente            = '" & _TipoContaCliente & "'" & vbCrLf
                strSQL &= "  ,TituloOrigem                = " & _CodigoTituloOrigem & vbCrLf
                strSQL &= "  ,ValorRecompra               = " & Str(_ValorRecompra) & vbCrLf
                strSQL &= "  ,TipoContaPagadora           = '" & _TipoContaPagadora & "'" & vbCrLf
                strSQL &= "  ,Processo                    = '" & _Processo & "'" & vbCrLf
                strSQL &= "  ,BoletoBancario              = " & IIf(_BoletoBancario, 1, 0) & vbCrLf
                strSQL &= "  ,UsuarioBoletoBancario       ='" & _UsuarioBoletoBancario & "'" & vbCrLf
                strSQL &= "  ,UsuarioBoletoBancarioDate   = " & IIf(_UsuarioBoletoBancarioDate = Nothing, "NULL", "'" & _UsuarioBoletoBancarioDate.ToString("yyyy-MM-dd") & "'") & vbCrLf
                strSQL &= "  ,HistoricoRemessa            ='" & _HistoricoRemessa & "'" & vbCrLf
                strSQL &= "  ,FolhaDePagamento              = " & IIf(_FolhaDePagamento, 1, 0) & vbCrLf
                strSQL &= "  ,UsuarioFolhaDePagamento       ='" & _UsuarioFolhaDePagamento & "'" & vbCrLf
                strSQL &= "  ,UsuarioFolhaDePagamentoDate   = " & IIf(_UsuarioFolhaDePagamentoDate = Nothing, "NULL", "'" & _UsuarioFolhaDePagamentoDate.ToString("yyyy-MM-dd") & "'") & vbCrLf
                strSQL &= "  ,ObservacoesControleInterno  ='" & ObservacoesControleInterno & "'" & vbCrLf
                strSQL &= " WHERE Registro_Id = " & _Codigo & " " & vbCrLf

                sqls.Add(strSQL)

                If _CodigoProvisao = eProvisao.Baixa Then Contabilizar(sqls)

            Case "D"
                strSQL = "Delete "

                If _ReceberPagar = "M" Then
                    strSQL &= "MovimentacoesFinanceiras" & vbCrLf
                ElseIf _ReceberPagar = "R" Then
                    strSQL &= "ContasAReceber" & vbCrLf
                Else
                    strSQL &= "ContasAPagar" & vbCrLf
                End If

                strSQL &= " WHERE Registro_Id = " & _Codigo & ""

                sqls.Add(strSQL)
            Case "C"
                strSQL = "UPDATE "

                If _ReceberPagar = "M" Then
                    strSQL &= "MovimentacoesFinanceiras"
                ElseIf _ReceberPagar = "R" Then
                    strSQL &= "ContasAReceber"
                Else
                    strSQL &= "ContasAPagar"
                End If

                strSQL &= "  Set Situacao = 3" & vbCrLf
                strSQL &= "  ,UsuarioCancelamento        ='" & _UsuarioCancelamento & "'" & vbCrLf
                strSQL &= "  ,UsuarioCancelamentoData    = " & IIf(_UsuarioCancelamentoData = Nothing, "NULL", "'" & _UsuarioCancelamentoData.ToString("yyyy-MM-dd") & "'") & vbCrLf
                strSQL &= "  ,ObservacoesControleInterno ='" & _ObservacoesControleInterno & "'" & vbCrLf
                strSQL &= "WHERE Registro_Id = " & _Codigo

                sqls.Add(strSQL)
        End Select

        'Adiciona na tabela o Histórico Financeiro da movimentação do título
        If _ListHistoricoFinanceiro.Any() Then
            For Each hf As HistoricoFinanceiro In _ListHistoricoFinanceiro
                sqls.Add(hf.AdicionarHistoricoNF())
            Next
        End If

        Return sqls
    End Function

    Public Function DeletaContabilizacao(ByRef sqls As ArrayList) As ArrayList
        Dim Sql As String

        'Deleta Razão
        Sql = "DELETE FROM razao" & vbCrLf &
              " WHERE Titulo = " & _Codigo
        sqls.Add(Sql)

        Return sqls
    End Function

    Public Function Contabilizar(ByRef sqls As ArrayList) As ArrayList
        Dim Sql As String

        'Valor do Documento
        If Me.Agrupado = eAgrupamentoFinanceiro.Agrupado OrElse Me.Agrupado = eAgrupamentoFinanceiro.NaoAgrupado Then
            Sql = "INSERT INTO Razao " & vbCrLf &
                  "       (Empresa_Id, " & vbCrLf &
                  "       EndEmpresa_Id, " & vbCrLf &
                  "       Conta_Id, " & vbCrLf &
                  "       Cliente_Id, " & vbCrLf &
                  "       EndCliente_Id, " & vbCrLf &
                  "       Movimento_Id, " & vbCrLf &
                  "       Lote_Id, " & vbCrLf &
                  "       Sequencia_Id, " & vbCrLf &
                  "       Titulo, " & vbCrLf &
                  "       UnidadeDeNegocio, " & vbCrLf &
                  "       Indexador, " & vbCrLf &
                  "       DataMoeda, " & vbCrLf &
                  "       DebitoOficial, " & vbCrLf &
                  "       CreditoOficial, " & vbCrLf &
                  "       DebitoMoeda, " & vbCrLf &
                  "       CreditoMoeda, " & vbCrLf &
                  "       Conciliacao, " & vbCrLf &
                  "       DataDaBaixa, " & vbCrLf &
                  "       Historico, " & vbCrLf &
                  "       PrevistoRealizado," & vbCrLf &
                  "       Processo," & vbCrLf &
                  "       UsuarioInclusao," & vbCrLf &
                  "       UsuarioInclusaoData)" & vbCrLf &
                  "VALUES ('" & _CodigoEmpresa & "'," & vbCrLf &
                  "         " & _EndEmpresa & "," & vbCrLf

            If Not String.IsNullOrEmpty(_Tributo) Then
                Dim Encargo As New Encargo(_Tributo)

                Sql &= "'" & Encargo.ContaDebito & "'" & vbCrLf

                Dim objPlaContaTributo As New [Lib].Negocio.PlanoDeConta("", 0, Encargo.ContaDebito)
                If objPlaContaTributo.TemCliente Then
                    Sql &= ", '" & _CodigoCliente & "'" 'Cliente
                    Sql &= ", " & _EndCliente           'Endereco do Cliente
                Else
                    Sql &= ", ''"                       'Cliente
                    Sql &= ", 0"                        'Endereco do Cliente
                End If
            Else
                Sql &= "'" & Me.Carteira.CodigoContaCliente & "'" & vbCrLf

                If Me.Carteira.CodigoContaCliente.Length > 0 AndAlso Me.Carteira.ContaCliente.TemCliente Then
                    Sql &= ", '" & _CodigoCliente & "'"     'Cliente
                    Sql &= ", " & _EndCliente               'Endereco do Cliente
                Else
                    Sql &= ", ''"                           'Cliente
                    Sql &= ", 0"                            'Endereco do Cliente
                End If
            End If

            If Not _Baixa = Nothing Then
                Sql &= ", '" & _Baixa.ToString("yyyy-MM-dd") & "'"      'Data de Movimento
            Else
                Sql &= ", '" & _Movimento.ToString("yyyy-MM-dd") & "'"  'Data de Movimento
            End If

            Sql &= ", 0070"
            Sql &= ", " & _Codigo                        'Sequencia no Razao = Registro do Titulo
            Sql &= ", " & _Codigo                        'Numero do Titulo
            Sql &= ", '" & _CodigoUnidadeDeNegocio & "'" 'Unidade de Negócio
            Sql &= ", " & _CodigoIndexador               'Indexador
            Sql &= ", " & IIf(_Baixa = Nothing, "'" & _Movimento.ToString("yyyy-MM-dd") & "'", "'" & _Baixa.ToString("yyyy-MM-dd") & "'") & vbCrLf  'Data da Moeda

            'Valor Oficial
            If _ReceberPagar = "P" Then
                Sql &= ", " & Str(_ValorDoDocumento) 'Valor Débito Oficial
                Sql &= ", 0.0"                       'Valor Crédito Oficial
            Else
                Sql &= ", 0.0"                       'Valor Debito Oficial
                Sql &= ", " & Str(_ValorDoDocumento) 'Valor Crédito Oficial
            End If
            'Valor Moeda
            If _ReceberPagar = "P" Then
                Sql &= ", " & Str(_MoedaValorDoDocumento) 'Valor Débito Moeda
                Sql &= ", 0.0"                            'Valor Crédito Moeda
            Else
                Sql &= ", 0.0"                            'Valor Debito Moeda
                Sql &= ", " & Str(_MoedaValorDoDocumento) 'Valor Crédito Moeda
            End If

            'Arrumar isso - Furlan
            'If Raz_DebitoCredito = "C" AndAlso chkConciliado.Checked Then
            '    Sql &= ", 'B'"                                                             'Conciliação
            '    Sql &= ", '" & CDate(txtDataConciliacao.Value).ToString("yyyy/MM/dd") & "'" 'Data Conciliação
            'Else
            Sql &= ", NULL "                                                           'Conciliação
            Sql &= ", NULL "                                                           'Data Conciliação
            'End If

            Sql &= ", '" & Funcoes.EliminarCaracteresEspeciais(_Historico.Trim & ". " & _Observacoes.Trim) & "'"  'Histórico
            Sql &= ", 'P'"              'Previsto/Realizado

            If _ReceberPagar = "M" Then ''Processo
                Sql &= ",'MovimentacoesFinanceiras'"
            ElseIf _ReceberPagar = "R" Then
                Sql &= ",'ContasAReceber'"
            Else
                Sql &= ",'ContasAPagar'"
            End If

            Sql &= ", '" & HttpContext.Current.Session("ssNomeUsuario") & "'"             'Usuario que Baixou
            Sql &= ", '" & CDate(Today).ToString("yyyy/MM/dd") & "')" 'Data da Baixa

            sqls.Add(Sql)

            'Descontos
            If _Descontos > 0 Then
                Sql = "INSERT INTO Razao " & vbCrLf &
                      "       (Empresa_Id, " & vbCrLf &
                      "       EndEmpresa_Id, " & vbCrLf &
                      "       Conta_Id, " & vbCrLf &
                      "       Cliente_Id, " & vbCrLf &
                      "       EndCliente_Id, " & vbCrLf &
                      "       Movimento_Id, " & vbCrLf &
                      "       Lote_Id, " & vbCrLf &
                      "       Sequencia_Id, " & vbCrLf &
                      "       Titulo, " & vbCrLf &
                      "       UnidadeDeNegocio, " & vbCrLf &
                      "       Indexador, " & vbCrLf &
                      "       DataMoeda, " & vbCrLf &
                      "       DebitoOficial, " & vbCrLf &
                      "       CreditoOficial, " & vbCrLf &
                      "       DebitoMoeda, " & vbCrLf &
                      "       CreditoMoeda, " & vbCrLf &
                      "       Conciliacao, " & vbCrLf &
                      "       DataDaBaixa, " & vbCrLf &
                      "       Historico, " & vbCrLf &
                      "       PrevistoRealizado," & vbCrLf &
                      "       Processo," & vbCrLf &
                      "       UsuarioInclusao," & vbCrLf &
                      "       UsuarioInclusaoData)" & vbCrLf &
                      "VALUES ('" & _CodigoEmpresa & "'," & vbCrLf &
                      "         " & _EndEmpresa & "," & vbCrLf &
                      "        '" & Me.Carteira.CodigoContaDesconto & "'" & vbCrLf

                If Me.Carteira.CodigoContaDesconto.Length > 0 AndAlso Me.Carteira.ContaDesconto.TemCliente Then
                    Sql &= ", '" & _CodigoCliente & "'"     'Cliente
                    Sql &= ", " & _EndCliente               'Endereco do Cliente
                Else
                    Sql &= ", ''"                           'Cliente
                    Sql &= ", 0"                            'Endereco do Cliente
                End If

                If Not _Baixa = Nothing Then
                    Sql &= ", '" & _Baixa.ToString("yyyy-MM-dd") & "'"      'Data de Movimento
                Else
                    Sql &= ", '" & _Movimento.ToString("yyyy-MM-dd") & "'"  'Data de Movimento
                End If

                Sql &= ", 0070"
                Sql &= ", " & _Codigo                        'Sequencia no Razao = Registro do Titulo
                Sql &= ", " & _Codigo                        'Numero do Titulo
                Sql &= ", '" & _CodigoUnidadeDeNegocio & "'" 'Unidade de Negócio
                Sql &= ", " & _CodigoIndexador               'Indexador
                Sql &= ", " & IIf(_Baixa = Nothing, "'" & _Movimento.ToString("yyyy-MM-dd") & "'", "'" & _Baixa.ToString("yyyy-MM-dd") & "'") & vbCrLf  'Data da Moeda

                'Valor Oficial
                If _ReceberPagar = "P" Then
                    Sql &= ", 0.0"                     'Valor Debito Oficial
                    Sql &= ", " & Str(_Descontos)      'Valor Crédito Oficial
                Else
                    Sql &= ", " & Str(_Descontos)      'Valor Débito Oficial
                    Sql &= ", 0.0"                     'Valor Crédito Oficial
                End If

                'Valor Moeda
                If _ReceberPagar = "P" Then
                    Sql &= ", 0.0"                     'Valor Debito Moeda
                    Sql &= ", " & Str(_MoedaDescontos) 'Valor Crédito Moeda
                Else
                    Sql &= ", " & Str(_MoedaDescontos) 'Valor Débito Moeda
                    Sql &= ", 0.0"                     'Valor Crédito Moeda
                End If

                'Arrumar isso - Furlan
                'If Raz_DebitoCredito = "C" AndAlso chkConciliado.Checked Then
                '    Sql &= ", 'B'"                                                             'Conciliação
                '    Sql &= ", '" & CDate(txtDataConciliacao.Value).ToString("yyyy/MM/dd") & "'" 'Data Conciliação
                'Else
                Sql &= ", NULL "                                                           'Conciliação
                Sql &= ", NULL "                                                           'Data Conciliação
                'End If

                Sql &= ", '" & Funcoes.EliminarCaracteresEspeciais(_Historico.Trim & ". " & _Observacoes.Trim) & "'"  'Histórico
                Sql &= ", 'P'"              'Previsto/Realizado

                If _ReceberPagar = "M" Then ''Processo
                    Sql &= ",'MovimentacoesFinanceiras'"
                ElseIf _ReceberPagar = "R" Then
                    Sql &= ",'ContasAReceber'"
                Else
                    Sql &= ",'ContasAPagar'"
                End If

                Sql &= ", '" & HttpContext.Current.Session("ssNomeUsuario") & "'"             'Usuario que Baixou
                Sql &= ", '" & CDate(Today).ToString("yyyy/MM/dd") & "')" 'Data da Baixa

                sqls.Add(Sql)
            End If

            'Deduções
            If _Deducoes > 0 Then
                Sql = "INSERT INTO Razao " & vbCrLf &
                      "       (Empresa_Id, " & vbCrLf &
                      "       EndEmpresa_Id, " & vbCrLf &
                      "       Conta_Id, " & vbCrLf &
                      "       Cliente_Id, " & vbCrLf &
                      "       EndCliente_Id, " & vbCrLf &
                      "       Movimento_Id, " & vbCrLf &
                      "       Lote_Id, " & vbCrLf &
                      "       Sequencia_Id, " & vbCrLf &
                      "       Titulo, " & vbCrLf &
                      "       UnidadeDeNegocio, " & vbCrLf &
                      "       Indexador, " & vbCrLf &
                      "       DataMoeda, " & vbCrLf &
                      "       DebitoOficial, " & vbCrLf &
                      "       CreditoOficial, " & vbCrLf &
                      "       DebitoMoeda, " & vbCrLf &
                      "       CreditoMoeda, " & vbCrLf &
                      "       Conciliacao, " & vbCrLf &
                      "       DataDaBaixa, " & vbCrLf &
                      "       Historico, " & vbCrLf &
                      "       PrevistoRealizado," & vbCrLf &
                      "       Processo," & vbCrLf &
                      "       UsuarioInclusao," & vbCrLf &
                      "       UsuarioInclusaoData)" & vbCrLf &
                      "VALUES ('" & _CodigoEmpresa & "'," & vbCrLf &
                      "         " & _EndEmpresa & "," & vbCrLf &
                      "        '" & Me.Carteira.CodigoContaDeducao & "'" & vbCrLf

                If Me.Carteira.CodigoContaDeducao.Length > 0 AndAlso Me.Carteira.ContaDeducao.TemCliente Then
                    Sql &= ", '" & _CodigoCliente & "'"     'Cliente
                    Sql &= ", " & _EndCliente               'Endereco do Cliente
                Else
                    Sql &= ", ''"                           'Cliente
                    Sql &= ", 0"                            'Endereco do Cliente
                End If

                If Not _Baixa = Nothing Then
                    Sql &= ", '" & _Baixa.ToString("yyyy-MM-dd") & "'"      'Data de Movimento
                Else
                    Sql &= ", '" & _Movimento.ToString("yyyy-MM-dd") & "'"  'Data de Movimento
                End If

                Sql &= ", 0070"
                Sql &= ", " & _Codigo                        'Sequencia no Razao = Registro do Titulo
                Sql &= ", " & _Codigo                        'Numero do Titulo
                Sql &= ", '" & _CodigoUnidadeDeNegocio & "'" 'Unidade de Negócio
                Sql &= ", " & _CodigoIndexador               'Indexador
                Sql &= ", " & IIf(_Baixa = Nothing, "'" & _Movimento.ToString("yyyy-MM-dd") & "'", "'" & _Baixa.ToString("yyyy-MM-dd") & "'") & vbCrLf  'Data da Moeda

                'Valor Oficial
                If _ReceberPagar = "P" Then
                    Sql &= ", 0.0"                     'Valor Debito Oficial
                    Sql &= ", " & Str(_Deducoes)      'Valor Crédito Oficial
                Else
                    Sql &= ", " & Str(_Deducoes)      'Valor Débito Oficial
                    Sql &= ", 0.0"                     'Valor Crédito Oficial
                End If

                'Valor Moeda
                If _ReceberPagar = "P" Then
                    Sql &= ", 0.0"                     'Valor Debito Moeda
                    Sql &= ", " & Str(_MoedaDeducoes) 'Valor Crédito Moeda
                Else
                    Sql &= ", " & Str(_MoedaDeducoes) 'Valor Débito Moeda
                    Sql &= ", 0.0"                     'Valor Crédito Moeda
                End If

                'Arrumar isso - Furlan
                'If Raz_DebitoCredito = "C" AndAlso chkConciliado.Checked Then
                '    Sql &= ", 'B'"                                                             'Conciliação
                '    Sql &= ", '" & CDate(txtDataConciliacao.Value).ToString("yyyy/MM/dd") & "'" 'Data Conciliação
                'Else
                Sql &= ", NULL "                                                           'Conciliação
                Sql &= ", NULL "                                                           'Data Conciliação
                'End If

                Sql &= ", '" & Funcoes.EliminarCaracteresEspeciais(_Historico.Trim & ". " & _Observacoes.Trim) & "'"  'Histórico
                Sql &= ", 'P'"              'Previsto/Realizado

                If _ReceberPagar = "M" Then ''Processo
                    Sql &= ",'MovimentacoesFinanceiras'"
                ElseIf _ReceberPagar = "R" Then
                    Sql &= ",'ContasAReceber'"
                Else
                    Sql &= ",'ContasAPagar'"
                End If

                Sql &= ", '" & HttpContext.Current.Session("ssNomeUsuario") & "'"             'Usuario que Baixou
                Sql &= ", '" & CDate(Today).ToString("yyyy/MM/dd") & "')" 'Data da Baixa

                sqls.Add(Sql)
            End If

            'Juros
            If _Juros > 0 Then
                Sql = "INSERT INTO Razao " & vbCrLf &
                      "       (Empresa_Id, " & vbCrLf &
                      "       EndEmpresa_Id, " & vbCrLf &
                      "       Conta_Id, " & vbCrLf &
                      "       Cliente_Id, " & vbCrLf &
                      "       EndCliente_Id, " & vbCrLf &
                      "       Movimento_Id, " & vbCrLf &
                      "       Lote_Id, " & vbCrLf &
                      "       Sequencia_Id, " & vbCrLf &
                      "       Titulo, " & vbCrLf &
                      "       UnidadeDeNegocio, " & vbCrLf &
                      "       Indexador, " & vbCrLf &
                      "       DataMoeda, " & vbCrLf &
                      "       DebitoOficial, " & vbCrLf &
                      "       CreditoOficial, " & vbCrLf &
                      "       DebitoMoeda, " & vbCrLf &
                      "       CreditoMoeda, " & vbCrLf &
                      "       Conciliacao, " & vbCrLf &
                      "       DataDaBaixa, " & vbCrLf &
                      "       Historico, " & vbCrLf &
                      "       PrevistoRealizado," & vbCrLf &
                      "       Processo," & vbCrLf &
                      "       UsuarioInclusao," & vbCrLf &
                      "       UsuarioInclusaoData)" & vbCrLf &
                      "VALUES ('" & _CodigoEmpresa & "'," & vbCrLf &
                      "         " & _EndEmpresa & "," & vbCrLf &
                      "        '" & Me.Carteira.CodigoContaJuro & "'" & vbCrLf

                If Me.Carteira.CodigoContaJuro.Length > 0 AndAlso Me.Carteira.ContaJuro.TemCliente Then
                    Sql &= ", '" & _CodigoCliente & "'"     'Cliente
                    Sql &= ", " & _EndCliente               'Endereco do Cliente
                Else
                    Sql &= ", ''"                           'Cliente
                    Sql &= ", 0"                            'Endereco do Cliente
                End If

                If Not _Baixa = Nothing Then
                    Sql &= ", '" & _Baixa.ToString("yyyy-MM-dd") & "'"      'Data de Movimento
                Else
                    Sql &= ", '" & _Movimento.ToString("yyyy-MM-dd") & "'"  'Data de Movimento
                End If

                Sql &= ", 0070"
                Sql &= ", " & _Codigo                        'Sequencia no Razao = Registro do Titulo
                Sql &= ", " & _Codigo                        'Numero do Titulo
                Sql &= ", '" & _CodigoUnidadeDeNegocio & "'" 'Unidade de Negócio
                Sql &= ", " & _CodigoIndexador               'Indexador
                Sql &= ", " & IIf(_Baixa = Nothing, "'" & _Movimento.ToString("yyyy-MM-dd") & "'", "'" & _Baixa.ToString("yyyy-MM-dd") & "'") & vbCrLf  'Data da Moeda

                'Valor Oficial
                If _ReceberPagar = "P" Then
                    Sql &= ", " & Str(_Juros)      'Valor Débito Oficial
                    Sql &= ", 0.0"                 'Valor Crédito Oficial
                Else
                    Sql &= ", 0.0"                 'Valor Debito Oficial
                    Sql &= ", " & Str(_Juros)      'Valor Crédito Oficial
                End If

                'Valor Moeda
                If _ReceberPagar = "P" Then
                    Sql &= ", " & Str(_MoedaJuros) 'Valor Débito Moeda
                    Sql &= ", 0.0"                 'Valor Crédito Moeda
                Else
                    Sql &= ", 0.0"                 'Valor Debito Moeda
                    Sql &= ", " & Str(_MoedaJuros) 'Valor Crédito Moeda
                End If

                'Arrumar isso - Furlan
                'If Raz_DebitoCredito = "C" AndAlso chkConciliado.Checked Then
                '    Sql &= ", 'B'"                                                             'Conciliação
                '    Sql &= ", '" & CDate(txtDataConciliacao.Value).ToString("yyyy/MM/dd") & "'" 'Data Conciliação
                'Else
                Sql &= ", NULL "                                                           'Conciliação
                Sql &= ", NULL "                                                           'Data Conciliação
                'End If

                Sql &= ", '" & Funcoes.EliminarCaracteresEspeciais(_Historico.Trim & ". " & _Observacoes.Trim) & "'"  'Histórico
                Sql &= ", 'P'"              'Previsto/Realizado

                If _ReceberPagar = "M" Then ''Processo
                    Sql &= ",'MovimentacoesFinanceiras'"
                ElseIf _ReceberPagar = "R" Then
                    Sql &= ",'ContasAReceber'"
                Else
                    Sql &= ",'ContasAPagar'"
                End If

                Sql &= ", '" & HttpContext.Current.Session("ssNomeUsuario") & "'"             'Usuario que Baixou
                Sql &= ", '" & CDate(Today).ToString("yyyy/MM/dd") & "')" 'Data da Baixa

                sqls.Add(Sql)
            End If

            'Acréscimos
            If _Acrescimos > 0 Then
                Sql = "INSERT INTO Razao " & vbCrLf &
                      "       (Empresa_Id, " & vbCrLf &
                      "       EndEmpresa_Id, " & vbCrLf &
                      "       Conta_Id, " & vbCrLf &
                      "       Cliente_Id, " & vbCrLf &
                      "       EndCliente_Id, " & vbCrLf &
                      "       Movimento_Id, " & vbCrLf &
                      "       Lote_Id, " & vbCrLf &
                      "       Sequencia_Id, " & vbCrLf &
                      "       Titulo, " & vbCrLf &
                      "       UnidadeDeNegocio, " & vbCrLf &
                      "       Indexador, " & vbCrLf &
                      "       DataMoeda, " & vbCrLf &
                      "       DebitoOficial, " & vbCrLf &
                      "       CreditoOficial, " & vbCrLf &
                      "       DebitoMoeda, " & vbCrLf &
                      "       CreditoMoeda, " & vbCrLf &
                      "       Conciliacao, " & vbCrLf &
                      "       DataDaBaixa, " & vbCrLf &
                      "       Historico, " & vbCrLf &
                      "       PrevistoRealizado," & vbCrLf &
                      "       Processo," & vbCrLf &
                      "       UsuarioInclusao," & vbCrLf &
                      "       UsuarioInclusaoData)" & vbCrLf &
                      "VALUES ('" & _CodigoEmpresa & "'," & vbCrLf &
                      "         " & _EndEmpresa & "," & vbCrLf &
                      "        '" & Me.Carteira.CodigoContaAcrescimo & "'" & vbCrLf

                If Me.Carteira.CodigoContaAcrescimo.Length > 0 AndAlso Me.Carteira.ContaAcrescimo.TemCliente Then
                    Sql &= ", '" & _CodigoCliente & "'"     'Cliente
                    Sql &= ", " & _EndCliente               'Endereco do Cliente
                Else
                    Sql &= ", ''"                           'Cliente
                    Sql &= ", 0"                            'Endereco do Cliente
                End If

                If Not _Baixa = Nothing Then
                    Sql &= ", '" & _Baixa.ToString("yyyy-MM-dd") & "'"      'Data de Movimento
                Else
                    Sql &= ", '" & _Movimento.ToString("yyyy-MM-dd") & "'"  'Data de Movimento
                End If

                Sql &= ", 0070"
                Sql &= ", " & _Codigo                        'Sequencia no Razao = Registro do Titulo
                Sql &= ", " & _Codigo                        'Numero do Titulo
                Sql &= ", '" & _CodigoUnidadeDeNegocio & "'" 'Unidade de Negócio
                Sql &= ", " & _CodigoIndexador               'Indexador
                Sql &= ", " & IIf(_Baixa = Nothing, "'" & _Movimento.ToString("yyyy-MM-dd") & "'", "'" & _Baixa.ToString("yyyy-MM-dd") & "'") & vbCrLf  'Data da Moeda

                'Valor Oficial
                If _ReceberPagar = "P" Then
                    Sql &= ", " & Str(_Acrescimos)      'Valor Débito Oficial
                    Sql &= ", 0.0"                      'Valor Crédito Oficial
                Else
                    Sql &= ", 0.0"                      'Valor Debito Oficial
                    Sql &= ", " & Str(_Acrescimos)      'Valor Crédito Oficial
                End If

                'Valor Moeda
                If _ReceberPagar = "P" Then
                    Sql &= ", " & Str(_MoedaAcrescimos) 'Valor Débito Moeda
                    Sql &= ", 0.0"                      'Valor Crédito Moeda
                Else
                    Sql &= ", 0.0"                      'Valor Debito Moeda
                    Sql &= ", " & Str(_MoedaAcrescimos) 'Valor Crédito Moeda
                End If

                'Arrumar isso - Furlan
                'If Raz_DebitoCredito = "C" AndAlso chkConciliado.Checked Then
                '    Sql &= ", 'B'"                                                             'Conciliação
                '    Sql &= ", '" & CDate(txtDataConciliacao.Value).ToString("yyyy/MM/dd") & "'" 'Data Conciliação
                'Else
                Sql &= ", NULL "                                                           'Conciliação
                Sql &= ", NULL "                                                           'Data Conciliação
                'End If

                Sql &= ", '" & Funcoes.EliminarCaracteresEspeciais(_Historico.Trim & ". " & _Observacoes.Trim) & "'"  'Histórico
                Sql &= ", 'P'"              'Previsto/Realizado

                If _ReceberPagar = "M" Then ''Processo
                    Sql &= ",'MovimentacoesFinanceiras'"
                ElseIf _ReceberPagar = "R" Then
                    Sql &= ",'ContasAReceber'"
                Else
                    Sql &= ",'ContasAPagar'"
                End If

                Sql &= ", '" & HttpContext.Current.Session("ssNomeUsuario") & "'"             'Usuario que Baixou
                Sql &= ", '" & CDate(Today).ToString("yyyy/MM/dd") & "')" 'Data da Baixa

                sqls.Add(Sql)
            End If
        End If


        'Líquido
        If Me.Agrupado = eAgrupamentoFinanceiro.Mestre OrElse Me.Agrupado = eAgrupamentoFinanceiro.NaoAgrupado Then
            Sql = "INSERT INTO Razao " & vbCrLf &
                  "       (Empresa_Id, " & vbCrLf &
                  "       EndEmpresa_Id, " & vbCrLf &
                  "       Conta_Id, " & vbCrLf &
                  "       Cliente_Id, " & vbCrLf &
                  "       EndCliente_Id, " & vbCrLf &
                  "       Movimento_Id, " & vbCrLf &
                  "       Lote_Id, " & vbCrLf &
                  "       Sequencia_Id, " & vbCrLf &
                  "       Titulo, " & vbCrLf &
                  "       UnidadeDeNegocio, " & vbCrLf &
                  "       Indexador, " & vbCrLf &
                  "       DataMoeda, " & vbCrLf &
                  "       DebitoOficial, " & vbCrLf &
                  "       CreditoOficial, " & vbCrLf &
                  "       DebitoMoeda, " & vbCrLf &
                  "       CreditoMoeda, " & vbCrLf &
                  "       Conciliacao, " & vbCrLf &
                  "       DataDaBaixa, " & vbCrLf &
                  "       Historico, " & vbCrLf &
                  "       PrevistoRealizado," & vbCrLf &
                  "       Processo," & vbCrLf &
                  "       UsuarioInclusao," & vbCrLf &
                  "       UsuarioInclusaoData)" & vbCrLf &
                  "VALUES ('" & _CodigoEmpresaPagadora & "'," & vbCrLf &
                  "         " & _EndEmpresaPagadora & "," & vbCrLf &
                  "        '" & _ContaContabilPagadora & "'" & vbCrLf

            Dim objPlaContaLiquido As New [Lib].Negocio.PlanoDeConta("", 0, _ContaContabilPagadora)
            If objPlaContaLiquido.TemCliente Then
                Sql &= ", '" & _CodigoCliente & "'"     'Cliente
                Sql &= ", " & _EndCliente               'Endereco do Cliente
            Else
                Sql &= ", ''"                           'Cliente
                Sql &= ", 0"                            'Endereco do Cliente
            End If

            If Not _Baixa = Nothing Then
                Sql &= ", '" & _Baixa.ToString("yyyy-MM-dd") & "'"      'Data de Movimento
            Else
                Sql &= ", '" & _Movimento.ToString("yyyy-MM-dd") & "'"  'Data de Movimento
            End If

            Sql &= ", 0070"
            Sql &= ", " & _Codigo                        'Sequencia no Razao = Registro do Titulo
            Sql &= ", " & _Codigo                        'Numero do Titulo

            Sql &= ", '" & Me.UnidadeDeNegocioEmpresaPagadora.CodigoUnidade & "'" 'Unidade de Negócio
            Sql &= ", " & _CodigoIndexador               'Indexador
            Sql &= ", " & IIf(_Baixa = Nothing, "'" & _Movimento.ToString("yyyy-MM-dd") & "'", "'" & _Baixa.ToString("yyyy-MM-dd") & "'") & vbCrLf  'Data da Moeda

            'Valor Oficial
            If _ReceberPagar = "P" Then
                Sql &= ", 0.0"                   'Valor Debito Oficial
                Sql &= ", " & Str(_ValorLiquido) 'Valor Crédito Oficial
            Else
                Sql &= ", " & Str(_ValorLiquido) 'Valor Débito Oficial
                Sql &= ", 0.0"                   'Valor Crédito Oficial
            End If

            'Valor Moeda
            If _ReceberPagar = "P" Then
                Sql &= ", 0.0"                        'Valor Debito Moeda
                Sql &= ", " & Str(_MoedaValorLiquido) 'Valor Crédito Moeda
            Else
                Sql &= ", " & Str(_MoedaValorLiquido) 'Valor Débito Moeda
                Sql &= ", 0.0"                        'Valor Crédito Moeda
            End If

            'Arrumar isso - Furlan
            'If Raz_DebitoCredito = "C" AndAlso chkConciliado.Checked Then
            '    Sql &= ", 'B'"                                                             'Conciliação
            '    Sql &= ", '" & CDate(txtDataConciliacao.Value).ToString("yyyy/MM/dd") & "'" 'Data Conciliação
            'Else
            Sql &= ", NULL "                                                           'Conciliação
            Sql &= ", NULL "                                                           'Data Conciliação
            'End If

            Sql &= ", '" & Funcoes.EliminarCaracteresEspeciais(_Historico.Trim & ". " & _Observacoes.Trim) & "'"  'Histórico
            Sql &= ", 'P'"              'Previsto/Realizado

            If _ReceberPagar = "M" Then ''Processo
                Sql &= ",'MovimentacoesFinanceiras'"
            ElseIf _ReceberPagar = "R" Then
                Sql &= ",'ContasAReceber'"
            Else
                Sql &= ",'ContasAPagar'"
            End If

            Sql &= ", '" & HttpContext.Current.Session("ssNomeUsuario") & "'"             'Usuario que Baixou
            Sql &= ", '" & CDate(Today).ToString("yyyy/MM/dd") & "')" 'Data da Baixa

            sqls.Add(Sql)
        End If

        '-------------------------------------------
        'Transferencias Financeiras
        '-------------------------------------------
        Sql = " SELECT EmpresaDebito, EnderecoDebito, EmpresaCredito, EnderecoCredito, EmpresaContabil, EnderecoContabil, " & vbCrLf &
              "        ContaContabil, ClienteContabil,EndClienteContabil, " & vbCrLf &
              IIf(_ReceberPagar = "P", "DebitoCredito", "case when DebitoCredito='D'then 'C' else 'D' end DebitoCredito ") & vbCrLf &
              "   FROM TransferenciasFinanceiras " & vbCrLf &
              "  WHERE EmpresaDebito   ='" & _CodigoEmpresa & "'" & vbCrLf &
              "    and EnderecoDebito  = " & _EndEmpresa & vbCrLf &
              "    and EmpresaCredito  ='" & _CodigoEmpresaPagadora & "'" & vbCrLf &
              "    and EnderecoCredito = " & _EndEmpresaPagadora

        Dim Banco As New AcessaBanco
        Dim ds As DataSet

        ds = Banco.ConsultaDataSet(Sql, "TransferenciasFinanceiras")

        For Each DrT As DataRow In ds.Tables(0).Rows
            Sql = "INSERT INTO Razao " & vbCrLf &
                  "       (Empresa_Id, " & vbCrLf &
                  "       EndEmpresa_Id, " & vbCrLf &
                  "       Conta_Id, " & vbCrLf &
                  "       Cliente_Id, " & vbCrLf &
                  "       EndCliente_Id, " & vbCrLf &
                  "       Movimento_Id, " & vbCrLf &
                  "       Lote_Id, " & vbCrLf &
                  "       Sequencia_Id, " & vbCrLf &
                  "       Titulo, " & vbCrLf &
                  "       UnidadeDeNegocio, " & vbCrLf &
                  "       Indexador, " & vbCrLf &
                  "       DataMoeda, " & vbCrLf &
                  "       DebitoOficial, " & vbCrLf &
                  "       CreditoOficial, " & vbCrLf &
                  "       DebitoMoeda, " & vbCrLf &
                  "       CreditoMoeda, " & vbCrLf &
                  "       Conciliacao, " & vbCrLf &
                  "       DataDaBaixa, " & vbCrLf &
                  "       Historico, " & vbCrLf &
                  "       PrevistoRealizado," & vbCrLf &
                  "       Processo," & vbCrLf &
                  "       UsuarioInclusao," & vbCrLf &
                  "       UsuarioInclusaoData)" & vbCrLf &
                  "VALUES ('" & DrT("EmpresaContabil") & "'," & vbCrLf &
                  "         " & DrT("EnderecoContabil") & "," & vbCrLf &
                  "        '" & DrT("ContaContabil") & "'" & vbCrLf

            Sql &= ", '" & DrT("ClienteContabil") & "'"  'Cliente
            Sql &= ", " & DrT("EndClienteContabil")      'Endereco do Cliente

            If Not _Baixa = Nothing Then
                Sql &= ", '" & _Baixa.ToString("yyyy-MM-dd") & "'"      'Data de Movimento
            Else
                Sql &= ", '" & _Movimento.ToString("yyyy-MM-dd") & "'"  'Data de Movimento
            End If

            Sql &= ", 0070"
            Sql &= ", " & _Codigo                        'Sequencia no Razao = Registro do Titulo
            Sql &= ", " & _Codigo                        'Numero do Titulo

            Dim UnidadePagadora = New GruposXEmpresas(DrT("EmpresaContabil"), DrT("EnderecoContabil"))
            Sql &= ", '" & UnidadePagadora.CodigoUnidade & "'" 'Unidade de Negócio

            Sql &= ", " & _CodigoIndexador               'Indexador
            Sql &= ", " & IIf(_Baixa = Nothing, "'" & _Movimento.ToString("yyyy-MM-dd") & "'", "'" & _Baixa.ToString("yyyy-MM-dd") & "'") & vbCrLf  'Data da Moeda

            'Valor Oficial
            If DrT("DebitoCredito") = "D" Then
                Sql &= ", " & Str(_ValorLiquido)   'Valor Débito Oficial
                Sql &= ", 0.0"                     'Valor Crédito Oficial
            Else
                Sql &= ", 0.0"                     'Valor Debito Oficial
                Sql &= ", " & Str(_ValorLiquido)   'Valor Crédito Oficial
            End If

            'Valor Moeda
            If DrT("DebitoCredito") = "D" Then
                Sql &= ", " & Str(_MoedaValorLiquido) 'Valor Débito Moeda
                Sql &= ", 0.0"                        'Valor Crédito Moeda
            Else
                Sql &= ", 0.0"                        'Valor Debito Moeda
                Sql &= ", " & Str(_MoedaValorLiquido) 'Valor Crédito Moeda
            End If

            'Arrumar isso - Furlan
            'If Raz_DebitoCredito = "C" AndAlso chkConciliado.Checked Then
            '    Sql &= ", 'B'"                                                             'Conciliação
            '    Sql &= ", '" & CDate(txtDataConciliacao.Value).ToString("yyyy/MM/dd") & "'" 'Data Conciliação
            'Else
            Sql &= ", NULL "                                                           'Conciliação
            Sql &= ", NULL "                                                           'Data Conciliação
            'End If

            Sql &= ", '" & Funcoes.EliminarCaracteresEspeciais(_Historico.Trim & ". " & _Observacoes.Trim) & "'"  'Histórico
            Sql &= ", 'P'"              'Previsto/Realizado

            If _ReceberPagar = "M" Then ''Processo
                Sql &= ",'MovimentacoesFinanceiras'"
            ElseIf _ReceberPagar = "R" Then
                Sql &= ",'ContasAReceber'"
            Else
                Sql &= ",'ContasAPagar'"
            End If

            Sql &= ", '" & HttpContext.Current.Session("ssNomeUsuario") & "'"             'Usuario que Baixou
            Sql &= ", '" & CDate(Today).ToString("yyyy/MM/dd") & "')" 'Data da Baixa

            sqls.Add(Sql)
        Next

        Return sqls
    End Function

    Public Function AddNotaxTituloSql(ByRef CodigoTitulo As Integer, ByVal pIUD As String, ByRef NF As NotaFiscal) As String
        Dim sql As String = ""
        Select Case pIUD
            Case "I"
                sql = "Insert into NotaFiscalXTitulo(Empresa_Id, EndEmpresa_Id, Cliente_Id, EndCliente_Id, EntradaSaida_Id, Serie_Id, Nota_Id, Titulo_Id)" & vbCrLf &
                      " values('" & NF.CodigoEmpresa & "'," & NF.EnderecoEmpresa & ", " & vbCrLf &
                      "'" & NF.Pedido.CodigoCliente & "'," & NF.Pedido.EnderecoCliente & "," & vbCrLf &
                      "'" & NF.EntradaSaida.ToString.Substring(0, 1) & "','" & NF.Serie & "'," & NF.Codigo & "," & CodigoTitulo & ")"
            Case "D"
                sql = " Delete NotaFiscalXTitulo" & vbCrLf &
                      "  Where Empresa_Id      ='" & NF.CodigoEmpresa & "'" & vbCrLf &
                      "    and EndEmpresa_Id   = " & NF.EnderecoEmpresa & vbCrLf &
                      "    and Cliente_Id      ='" & NF.Pedido.CodigoCliente & "'" & vbCrLf &
                      "    and EndCliente_Id   = " & NF.Pedido.EnderecoCliente & vbCrLf &
                      "    and EntradaSaida_Id ='" & NF.EntradaSaida.ToString.Substring(0, 1) & "'" & vbCrLf &
                      "    and Serie_Id        ='" & NF.Serie & "'" & vbCrLf &
                      "    and Nota_Id         = " & NF.Codigo & vbCrLf &
                      "    and Titulo_Id       = " & CodigoTitulo
        End Select

        Return sql
    End Function
#End Region

End Class