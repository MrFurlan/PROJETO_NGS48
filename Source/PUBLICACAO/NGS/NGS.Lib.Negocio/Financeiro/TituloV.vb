Imports Microsoft.VisualBasic
Imports System.Collections.Generic
Imports System.Data
Imports System.Web
Imports NGS.Lib.Uteis


'*********************************************************************************************************************************************************************
'****************************************    LIST DA CLASSE BASE DE TITULO     ***************************************************************************************
'*********************************************************************************************************************************************************************
<Serializable()> _
Public Class ListTituloV
    Inherits List(Of TituloV)

#Region "Construtores"
    'Where
    Public Sub New(ByVal Where As String)
        Dim sql As String
        sql = " Select Registro_id " & vbCrLf & _
              "   from contasareceber " & vbCrLf & _
              "  Where " & Where & vbCrLf & _
              "  Union all" & vbCrLf & _
              " Select Registro_id " & vbCrLf & _
              "   from contasaPagar " & vbCrLf & _
              "  Where " & Where

        Dim ds As DataSet
        Dim Banco As New AcessaBanco
        ds = Banco.ConsultaDataSet(sql, "Titulos")
        For Each row As DataRow In ds.Tables(0).Rows
            Dim Tit As New TituloV(row("Registro_Id"))
            Me.Add(Tit)
        Next
        Banco = Nothing
    End Sub

    'Lista Vazia
    Public Sub New()

    End Sub

    'Lista Pedido
    Sub New(ByRef CodigoPedido As Integer, ByVal PagRec As String, ByVal Empresa As String, ByVal EndEmpresa As Integer)
        Dim Banco As New AcessaBanco
        Dim ds As New DataSet
        Dim Sql As String = ""

        Sql = " Select Registro_id, Moeda " & vbCrLf & _
              "   from (" & vbCrLf & _
              "         Select 'P' as Tipo, Registro_id, Moeda " & vbCrLf & _
              "           From Contasapagar" & vbCrLf & _
              "          Where Pedido   = " & CodigoPedido & vbCrLf & _
              "            and Empresa = '" & Empresa & "'" & vbCrLf & _
              "            and EndEmpresa = " & EndEmpresa & _
              "            and Situacao = 1" & vbCrLf & _
              "          Union all" & vbCrLf & _
              "         select 'R', Registro_id, Moeda " & vbCrLf & _
              "           from ContasaReceber" & vbCrLf & _
              "          where Pedido   = " & CodigoPedido & vbCrLf & _
              "            and Empresa = '" & Empresa & "'" & vbCrLf & _
              "            and EndEmpresa = " & EndEmpresa & _
              "            and Situacao = 1" & vbCrLf & _
              "        ) sb" & vbCrLf & _
              "  Where tipo ='" & PagRec & "'"

        ds = Banco.ConsultaDataSet(Sql, "Titulos")

        If ds.Tables(0).Rows.Count = 0 Then
            Exit Sub
        Else
            Dim moeda As New Moeda(ds.Tables(0).Rows(0)("Moeda"))
            _TipoMoeda = moeda.Classificacao
        End If

        For Each row As DataRow In ds.Tables(0).Rows
            Dim Tit As New TituloV(row("Registro_id"))
            Me.Add(Tit)
        Next

        ES_Pedido = IIf(PagRec = "P", "E", "S")
    End Sub

    'Lista Nota
    Sub New(ByRef pNF As NotaFiscal)
        _NF = pNF
        If _NF.CodigoTipoDeDocumento = 1 AndAlso NF.Pedido IsNot Nothing AndAlso NF.Pedido.Codigo > 0 Then
            _TipoMoeda = NF.Pedido.Moeda.Classificacao
            ES_Pedido = NF.Pedido.SubOperacao.EntradaSaida.ToString.Substring(0, 1)
        Else
            _TipoMoeda = eTiposMoeda.Oficial
            ES_Pedido = NF.SubOperacao.EntradaSaida.ToString.Substring(0, 1)
        End If
        Dim Banco As New AcessaBanco
        Dim ds As New DataSet
        Dim Sql As String = ""

        Sql = "SELECT Titulo_Id" & vbCrLf & _
              "  FROM NotaFiscalXTitulo" & vbCrLf & _
              " Where Empresa_Id      ='" & pNF.CodigoEmpresa & "'" & vbCrLf & _
              "   and EndEmpresa_Id   = " & pNF.EnderecoEmpresa & vbCrLf & _
              "   and EntradaSaida_Id ='" & IIf(pNF.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "'" & vbCrLf & _
              "   and Serie_Id        ='" & pNF.Serie & "'" & vbCrLf & _
              "   and Nota_Id         = " & pNF.Codigo & vbCrLf & _
              "   and Cliente_Id      ='" & pNF.CodigoCliente & "'" & vbCrLf & _
              "   and EndCliente_Id   = " & pNF.EnderecoCliente & vbCrLf & _
              "   and Nota_id         > 0"

        ds = Banco.ConsultaDataSet(Sql, "Titulos")

        For Each row As DataRow In ds.Tables(0).Rows
            Dim Tit As New TituloV(row("Titulo_Id"))
            If Tit.CodigoSituacao = 1 Then Me.Add(Tit)
        Next

        If Me.Count = 0 AndAlso pNF.CodigoPedido > 0 AndAlso pNF.Pedido.AgruparFinanceiro Then
            Sql = "SELECT NFxT.Titulo_Id" & vbCrLf & _
                  "  FROM NotaFiscalXTitulo NFxT" & vbCrLf & _
                  " Inner Join NotasFiscais NF" & vbCrLf & _
                  "    on NFxT.Empresa_Id      = NF.Empresa_Id" & vbCrLf & _
                  "   and NFxT.EndEmpresa_Id   = NF.EndEmpresa_Id" & vbCrLf & _
                  "   and NFxT.Cliente_Id      = NF.Cliente_Id" & vbCrLf & _
                  "   and NFxT.EndCliente_Id   = NF.EndCliente_Id" & vbCrLf & _
                  "   and NFxT.EntradaSaida_Id = NF.EntradaSaida_Id" & vbCrLf & _
                  "   and NFxT.Serie_Id        = NF.Serie_Id" & vbCrLf & _
                  "   and NFxT.Nota_Id         = NF.Nota_Id " & vbCrLf & _
                  " Inner Join (" & vbCrLf & _
                  "             SELECT Registro_Id as Titulo_Id, Provisao, isnull(UsuarioLiberacao,'') as UsuarioLiberacao" & vbCrLf & _
                  "			      FROM ContasAPagar" & vbCrLf & _
                  "			     Union" & vbCrLf & _
                  "			    SELECT Registro_Id, Provisao, isnull(UsuarioLiberacao,'') as UsuarioLiberacao" & vbCrLf & _
                  "			      FROM ContasAreceber" & vbCrLf & _
                  "             ) Titulos" & vbCrLf & _
                  "    on Titulos.Titulo_Id = NFxT.Titulo_Id" & vbCrLf & _
                  " Where NF.Pedido = (Select Pedido" & vbCrLf & _
                  "                      from NotasFiscais" & vbCrLf & _
                  "                     Where Empresa_Id      ='" & pNF.CodigoEmpresa & "'" & vbCrLf & _
                  "                       and EndEmpresa_Id   = " & pNF.EnderecoEmpresa & vbCrLf & _
                  "                       and Cliente_Id      ='" & pNF.Pedido.CodigoCliente & "'" & vbCrLf & _
                  "                       and EndCliente_Id   = " & pNF.Pedido.EnderecoCliente & vbCrLf & _
                  "                       and EntradaSaida_Id ='" & IIf(pNF.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "'" & vbCrLf & _
                  "                       and Serie_Id        ='" & pNF.Serie & "'" & vbCrLf & _
                  "                       and Nota_Id         = " & pNF.Codigo & vbCrLf & _
                  "                    )" & vbCrLf & _
                  "  and Titulos.Provisao         = 2" & vbCrLf & _
                  "  and Titulos.UsuarioLiberacao = ''" & vbCrLf & _
                  "  and NF.Nota_id > 0"

            ds = Banco.ConsultaDataSet(Sql, "Titulos")

            For Each row As DataRow In ds.Tables(0).Rows
                Dim Tit As New TituloV(row("Titulo_Id"))
                If Tit.CodigoSituacao = 1 Then Me.Add(Tit)
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
                If Tit.IUD = "I" And Not UsaNumerador Then
                    Tit.Codigo = NumeroTitulo + i
                    i += 1
                End If

                Tit.SalvarSql(Sqls, UsaNumerador)
                If Not NF Is Nothing AndAlso Not Tit.IUD = "U" Then
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
                sql = "Insert into NotaFiscalXTitulo(Empresa_Id, EndEmpresa_Id, Cliente_Id, EndCliente_Id, EntradaSaida_Id, Serie_Id, Nota_Id, Titulo_Id)" & vbCrLf & _
                      " values('" & NF.CodigoEmpresa & "'," & NF.EnderecoEmpresa & ", " & vbCrLf & _
                      "'" & NF.Pedido.CodigoCliente & "'," & NF.Pedido.EnderecoCliente & "," & vbCrLf & _
                      "'" & NF.EntradaSaida.ToString.Substring(0, 1) & "','" & NF.Serie & "'," & NF.Codigo & "," & CodigoTitulo & ")"
            Case "D"
                sql = " Delete NotaFiscalXTitulo" & vbCrLf & _
                      "  Where Empresa_Id      ='" & NF.CodigoEmpresa & "'" & vbCrLf & _
                      "    and EndEmpresa_Id   = " & NF.EnderecoEmpresa & vbCrLf & _
                      "    and Cliente_Id      ='" & NF.Pedido.CodigoCliente & "'" & vbCrLf & _
                      "    and EndCliente_Id   = " & NF.Pedido.EnderecoCliente & vbCrLf & _
                      "    and EntradaSaida_Id ='" & NF.EntradaSaida.ToString.Substring(0, 1) & "'" & vbCrLf & _
                      "    and Serie_Id        ='" & NF.Serie & "'" & vbCrLf & _
                      "    and Nota_Id         = " & NF.Codigo & vbCrLf & _
                      "    and Titulo_Id       = " & CodigoTitulo
        End Select

        Return sql
    End Function
#End Region

#Region "Methods Usados pela Nota Fiscal / Pedido / Fixacao"
    Public Sub ParcelarNota(ByRef FormaPagamento As Integer)
        Dim Banco As New AcessaBanco
        Dim ds As New DataSet
        Dim Sql As String = ""


        Dim ValorNota As Decimal = NF.Itens.ValorLiquido_Oficial
        Dim ValorNotaMoeda As Decimal = NF.Itens.ValorLiquido_Moeda

        If NF.TotalNota = 0 Then
            _msg = "Valor da Nota está zerado"
            Exit Sub
        End If

        Dim n As Numerador = New Numerador(1)
        Dim NumeroTitulo As Integer = n.Sequencia + 1
        Dim QtdeParcelas As Integer = NF.Pedido.ParcelasTitulosVituais.Count
        Dim Parcela As Integer

        If NF.Pedido.ParcelasTitulosVituais.Count > 0 Then
            NF.VencimentosNota.Clear()

            For Each PTV In NF.Pedido.ParcelasTitulosVituais
                If ValorNota + ValorNotaMoeda = 0 Then Exit For

                Dim Tit As New TituloV
                Tit.IndiceFixo = True

                Tit.IUD = "I"
                Tit.Codigo = NumeroTitulo
                NumeroTitulo += 1
                Parcela += 1
                Tit.Sequencia = 0
                Tit.CodigoProvisao = 2

                If NF.EntradaSaida = eEntradaSaida.Entrada Then
                    Tit.ReceberPagar = "P"
                    Tit.ContaContabilCliente = IIf(NF.SubOperacao.Devolucao, NF.Itens(0).Produto.CarteiraVenda.CodigoContaCliente, NF.Itens(0).Produto.CarteiraCompra.CodigoContaCliente)
                    Tit.CodigoCarteira = IIf(NF.SubOperacao.Devolucao, NF.Itens(0).Produto.CodigoCarteiraVenda, NF.Itens(0).Produto.CodigoCarteiraCompra)
                Else
                    Tit.ReceberPagar = "R"
                    Tit.ContaContabilCliente = IIf(NF.SubOperacao.Devolucao, NF.Itens(0).Produto.CarteiraCompra.CodigoContaCliente, NF.Itens(0).Produto.CarteiraVenda.CodigoContaCliente)
                    Tit.CodigoCarteira = IIf(NF.SubOperacao.Devolucao, NF.Itens(0).Produto.CodigoCarteiraCompra, NF.Itens(0).Produto.CodigoCarteiraVenda)
                End If

                Tit.Tributo = ""
                Tit.CodigoEmpresaPedido = NF.Pedido.CodigoEmpresa
                Tit.EndEmpresaPedido = NF.Pedido.EnderecoEmpresa
                Tit.CodigoPedido = NF.CodigoPedido
                Tit.CodigoPedidoFixacao = 0
                Tit.CodigoProcuracao = NF.CodigoProcuracao

                Tit.DataMoeda = PTV.Vencimento
                Tit.CodigoIndexador = NF.Pedido.CodigoIndexador
                Tit.CodigoMoeda = NF.Pedido.CodigoMoeda
                Tit.Movimento = PTV.Vencimento

                Tit.Vencimento = Funcoes.ValidaDataUtil(NF.Empresa.Codigo, NF.Empresa.CodigoEndereco, PTV.Vencimento)
                Tit.Prorrogacao = Tit.Vencimento
                Tit.Baixa = Tit.Vencimento

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


                If _TipoMoeda = eTiposMoeda.Oficial Then
                    If ValorNota > PTV.ValorOficial Then
                        Tit.ValorDoDocumento = PTV.ValorOficial
                        ValorNota -= PTV.ValorOficial
                    Else
                        Tit.ValorDoDocumento = ValorNota
                        ValorNota = 0
                        ValorNotaMoeda = 0
                    End If
                Else
                    If ValorNota > PTV.ValorOficial Then
                        Tit.ValorDoDocumento = PTV.ValorOficial
                        ValorNota -= PTV.ValorOficial
                    Else
                        Tit.ValorDoDocumento = ValorNota
                        ValorNota = 0
                    End If

                    If ValorNotaMoeda > PTV.ValorMoeda Then
                        Tit.MoedaValorDoDocumento = PTV.ValorMoeda
                        ValorNotaMoeda -= PTV.ValorMoeda
                    Else
                        Tit.MoedaValorDoDocumento = ValorNotaMoeda
                        ValorNotaMoeda = 0
                    End If
                End If

                Tit.Historico = "REF. NF " & IIf(NF.SubOperacao.Devolucao, " de Devolucao ", "") & NF.Codigo & "-" & NF.Serie & "-" & IIf(NF.EntradaSaida = eEntradaSaida.Saida, "S", "E") & ", Parcela " & Parcela & "/" & QtdeParcelas & ", Pedido: " & NF.CodigoPedido & " / " & NF.Cliente.Nome
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

                Tit.UsuarioInclusao = NF.UsuarioInclusao
                Tit.UsuarioInclusaoData = NF.DataInclusao
                Me.Add(Tit)
            Next
        Else
            Dim venc As Date = Funcoes.ValidaDataUtil(NF.Empresa.Codigo, NF.Empresa.CodigoEndereco, Date.Now.AddDays(3))

            Dim Tit As New TituloV
            Tit.IndiceFixo = True

            Tit.IUD = "I"
            Tit.Codigo = NumeroTitulo
            NumeroTitulo += 1
            Parcela += 1
            Tit.Sequencia = 0
            Tit.CodigoProvisao = 2

            If NF.EntradaSaida = eEntradaSaida.Entrada Then
                Tit.ReceberPagar = "P"
                Tit.ContaContabilCliente = IIf(NF.SubOperacao.Devolucao, NF.Itens(0).Produto.CarteiraVenda.CodigoContaCliente, NF.Itens(0).Produto.CarteiraCompra.CodigoContaCliente)
                Tit.CodigoCarteira = IIf(NF.SubOperacao.Devolucao, NF.Itens(0).Produto.CodigoCarteiraVenda, NF.Itens(0).Produto.CodigoCarteiraCompra)
            Else
                Tit.ReceberPagar = "R"
                Tit.ContaContabilCliente = IIf(NF.SubOperacao.Devolucao, NF.Itens(0).Produto.CarteiraCompra.CodigoContaCliente, NF.Itens(0).Produto.CarteiraVenda.CodigoContaCliente)
                Tit.CodigoCarteira = IIf(NF.SubOperacao.Devolucao, NF.Itens(0).Produto.CodigoCarteiraCompra, NF.Itens(0).Produto.CodigoCarteiraVenda)
            End If

            Tit.Tributo = ""
            Tit.CodigoEmpresaPedido = NF.Pedido.CodigoEmpresa
            Tit.EndEmpresaPedido = NF.Pedido.EnderecoEmpresa
            Tit.CodigoPedido = NF.CodigoPedido
            Tit.CodigoPedidoFixacao = 0
            Tit.CodigoProcuracao = NF.CodigoProcuracao

            Tit.DataMoeda = venc
            Tit.CodigoIndexador = NF.Pedido.CodigoIndexador
            Tit.CodigoMoeda = NF.Pedido.CodigoMoeda
            Tit.Movimento = venc

            Tit.Vencimento = venc
            Tit.Prorrogacao = venc
            Tit.Baixa = venc

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


            If _TipoMoeda = eTiposMoeda.Oficial Then
                Tit.ValorDoDocumento = ValorNota
            Else
                Tit.ValorDoDocumento = ValorNota
                Tit.MoedaValorDoDocumento = ValorNotaMoeda
            End If

            Tit.Historico = "REF. NF " & NF.Codigo & "-" & NF.Serie & "-" & IIf(NF.EntradaSaida = eEntradaSaida.Saida, "S", "E") & ", Parcela " & Parcela & ", Pedido: " & NF.CodigoPedido & " / " & NF.Cliente.Nome
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

            Tit.UsuarioInclusao = NF.UsuarioInclusao
            Tit.UsuarioInclusaoData = NF.DataInclusao
            Me.Add(Tit)
        End If

    End Sub

    Public Sub ParcelarNotaDevolucao(ByRef FormaPagamento As Integer)
        Dim Banco As New AcessaBanco
        Dim ds As New DataSet
        Dim Sql As String = ""


        Dim ValorNota As Decimal = NF.Itens.ValorLiquido_Oficial
        Dim ValorNotaMoeda As Decimal = NF.Itens.ValorLiquido_Moeda

        Dim Vencimento As Date
        Dim LimVlrOficial As Decimal
        Dim LimVlrMoeda As Decimal

        If NF.TotalNota = 0 Then
            _msg = "Valor da Nota está zerado"
            Exit Sub
        End If

        Dim n As Numerador = New Numerador(1)
        Dim NumeroTitulo As Integer = n.Sequencia + 1
        Dim Parcela As Integer

        Dim cot As Cotacao

        NF.VencimentosNota.Clear()
        For Each Parc In NF.Pedido.Parcelas.OrderByDescending(Function(s) s.CodigoParcela)
            If ValorNota = 0 Then Exit For

            Dim TV As PedidoXTituloVirtual = NF.Pedido.ParcelasTitulosVituais.Where(Function(s) s.NrParcela = Parc.CodigoParcela).FirstOrDefault
            If TV IsNot Nothing Then
                Vencimento = TV.Vencimento
                cot = New Cotacao(NF.Pedido.CodigoIndexador, Vencimento)

                If NF.Pedido.Moeda.Classificacao = eTiposMoeda.Oficial Then
                    LimVlrOficial = Parc.Valor - TV.ValorOficial
                    LimVlrMoeda = Funcoes.ConverteMoeda(Parc.Valor - TV.ValorOficial, cot.Indice, eTiposMoeda.MoedaEstrangeira, False, False, 2)
                Else
                    LimVlrOficial = Funcoes.ConverteMoeda(Parc.Valor - TV.ValorMoeda, cot.Indice, eTiposMoeda.Oficial, False, False, 2)
                    LimVlrMoeda = Parc.Valor - TV.ValorMoeda
                End If
            Else
                Vencimento = Funcoes.ValidaDataUtil(NF.Empresa.Codigo, NF.Empresa.CodigoEndereco, Parc.DataVencimento)
                cot = New Cotacao(NF.Pedido.CodigoIndexador, Vencimento)

                If NF.Pedido.Moeda.Classificacao = eTiposMoeda.Oficial Then
                    LimVlrOficial = Parc.Valor
                    LimVlrMoeda = Funcoes.ConverteMoeda(Parc.Valor, cot.Indice, eTiposMoeda.MoedaEstrangeira, False, False, 2)
                Else
                    LimVlrOficial = Funcoes.ConverteMoeda(Parc.Valor, cot.Indice, eTiposMoeda.Oficial, False, False, 2)
                    LimVlrMoeda = Parc.Valor
                End If
            End If


            If LimVlrOficial = 0 Then Continue For

            Dim Tit As New TituloV
            Tit.IndiceFixo = True

            Tit.IUD = "I"
            Tit.Codigo = NumeroTitulo
            NumeroTitulo += 1
            Parcela += 1
            Tit.Sequencia = 0
            Tit.CodigoProvisao = 2

            If NF.EntradaSaida = eEntradaSaida.Entrada Then
                Tit.ReceberPagar = "P"
                Tit.ContaContabilCliente = IIf(NF.SubOperacao.Devolucao, NF.Itens(0).Produto.CarteiraVenda.CodigoContaCliente, NF.Itens(0).Produto.CarteiraCompra.CodigoContaCliente)
                Tit.CodigoCarteira = IIf(NF.SubOperacao.Devolucao, NF.Itens(0).Produto.CodigoCarteiraVenda, NF.Itens(0).Produto.CodigoCarteiraCompra)
            Else
                Tit.ReceberPagar = "R"
                Tit.ContaContabilCliente = IIf(NF.SubOperacao.Devolucao, NF.Itens(0).Produto.CarteiraCompra.CodigoContaCliente, NF.Itens(0).Produto.CarteiraVenda.CodigoContaCliente)
                Tit.CodigoCarteira = IIf(NF.SubOperacao.Devolucao, NF.Itens(0).Produto.CodigoCarteiraCompra, NF.Itens(0).Produto.CodigoCarteiraVenda)
            End If

            Tit.Tributo = ""
            Tit.CodigoEmpresaPedido = NF.Pedido.CodigoEmpresa
            Tit.EndEmpresaPedido = NF.Pedido.EnderecoEmpresa
            Tit.CodigoPedido = NF.CodigoPedido
            Tit.CodigoPedidoFixacao = 0
            Tit.CodigoProcuracao = NF.CodigoProcuracao

            Tit.DataMoeda = Vencimento
            Tit.CodigoIndexador = NF.Pedido.CodigoIndexador
            Tit.CodigoMoeda = NF.Pedido.CodigoMoeda
            Tit.Movimento = Vencimento

            Tit.Vencimento = Vencimento
            Tit.Prorrogacao = Vencimento
            Tit.Baixa = Vencimento

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

            If _TipoMoeda = eTiposMoeda.Oficial Then
                If ValorNota > LimVlrOficial Then
                    Tit.ValorDoDocumento = LimVlrOficial
                    ValorNota -= LimVlrOficial
                Else
                    Tit.ValorDoDocumento = ValorNota
                    ValorNota = 0
                    ValorNotaMoeda = 0
                End If
            Else
                If ValorNota > LimVlrOficial Then
                    Tit.ValorDoDocumento = LimVlrOficial
                    ValorNota -= LimVlrOficial
                Else
                    Tit.ValorDoDocumento = ValorNota
                    ValorNota = 0
                End If

                If ValorNotaMoeda > LimVlrMoeda Then
                    Tit.MoedaValorDoDocumento = LimVlrMoeda
                    ValorNotaMoeda -= LimVlrMoeda
                Else
                    Tit.MoedaValorDoDocumento = ValorNotaMoeda
                    ValorNotaMoeda = 0
                End If
            End If

            Tit.Historico = "REF. NF " & NF.Codigo & "-" & NF.Serie & "-" & IIf(NF.EntradaSaida = eEntradaSaida.Saida, "S", "E") & ", Parcela " & Parcela & ", Pedido: " & NF.CodigoPedido & " / " & NF.Cliente.Nome
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

            Tit.UsuarioInclusao = NF.UsuarioInclusao
            Tit.UsuarioInclusaoData = NF.DataInclusao
            Me.Add(Tit)
        Next
    End Sub

    Public Sub ParcelarNotasFiscaisGerais(ByRef FormaPagamento As Integer, pValorDaNota As Decimal, pValorParcelado As Decimal, pValorPago As Decimal, Optional pProvisao As Integer = 3)
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

        Sql &= "	DECLARE " & vbCrLf & _
               "	@Diferenca numeric(18,2)," & vbCrLf & _
               "	@ValorIPI numeric(18,2)," & vbCrLf & _
               "	@ValorTotal numeric(18,2)," & vbCrLf & _
               "	@Data varchar(10)," & vbCrLf & _
               "	@FPagto int" & vbCrLf

        '--Informa o valor do IPI a ser cobrado"
        '-- Na Compra dilui em todas as parcelas nas outras cobra na primeira parcela
        If NF.SubOperacao.Classe = eClassesOperacoes.COMPRAS Then
            Sql &= "	set @ValorIPI   =  0 " & vbCrLf & _
                   "	set @ValorTotal =    " & Str(ValorAParcelar)
        Else
            Sql &= "	set @ValorIPI   =  " & Str(IPIICMSST) & vbCrLf & _
                   "	set @ValorTotal =  " & Str(ValorAParcelar - IPIICMSST) & vbCrLf
        End If

        Sql &= "	set @Data       = '" & NF.Movimento.ToString("yyyy-MM-dd") & "'" & vbCrLf & _
               "	set @FPagto     =  " & FormaPagamento.ToString & vbCrLf

        '--Seleciona no banco a forma de pagamento e as Parcelas"
        Sql &= "	SELECT Pagamentos.Pagamento_Id," & vbCrLf & _
               "		   Pagamentos.Descricao," & vbCrLf & _
               "		   PagamentosXParcelas.Sequencia_Id, " & vbCrLf & _
               "		   Pagamentos.Parcelas - " & lstVencimentos.Count & " as Parcelas," & vbCrLf & _
               "		   PagamentosXParcelas.Dias," & vbCrLf
        '--Soma a Data Base o Numero de dias referente ao numero da Parcela"
        Sql &= "		   convert(datetime,(@Data)) + PagamentosXParcelas.Dias as Vencimento," & vbCrLf
        '--Divide o valor total pelo numero de parcelas p/ descubrir o valor da parcela"
        Sql &= "		   round(@ValorTotal / (Pagamentos.Parcelas -" & lstVencimentos.Count & "), 2) as ValorParcela," & vbCrLf
        '-- Armazena o Valor Total para calcular a diferenca na divisao das parcelas"
        Sql &= "		   @ValorTotal as ValorTotal" & vbCrLf & _
               "	  INTO #Temp1 " & vbCrLf & _
               "	  FROM Pagamentos " & vbCrLf & _
               "	 INNER JOIN PagamentosXParcelas " & vbCrLf & _
               "		ON Pagamentos.Pagamento_Id = PagamentosXParcelas.Pagamento_Id " & vbCrLf & _
               "       and PagamentosXParcelas.Sequencia_Id <= Pagamentos.Parcelas - " & lstVencimentos.Count & vbCrLf & _
               "	 where Pagamentos.Pagamento_Id = @FPagto " & vbCrLf & _
               "	 order by PagamentosXParcelas.Sequencia_Id " & vbCrLf

        '	--Calcula o valor Parcelado para ver se ha diferenca com o valor total caso haja armazena o valor na variavel "
        Sql &= "	set @Diferenca = (Select top(1) ValorTotal - (ValorParcela * Parcelas) from #Temp1)" & vbCrLf


        '	--Atualiza o valor da primeira parcela acrescida do IPI + a Diferenca"
        Sql &= "	update #temp1 set" & vbCrLf & _
               "	ValorParcela = ValorParcela + @ValorIPI + @Diferenca" & vbCrLf & _
               "	where #temp1.Sequencia_Id = 1" & vbCrLf & _
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
                    PreencheTitulo(ds.Tables(0).Rows(i), row, row.Codigo, pProvisao)
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
                    Dim tit As New TituloV()
                    tit.IUD = "I"
                    PreencheTitulo(row, tit, NumeroTitulo, pProvisao)
                    NumeroTitulo += 1
                    Me.Add(tit)
                End If
                i += 1
            Next
        End If
    End Sub

    Private Sub PreencheTitulo(row As DataRow, tit As TituloV, NumeroTitulo As Integer, pProvisao As Integer)
        tit.Codigo = NumeroTitulo
        tit.Sequencia = 0
        tit.CodigoProvisao = pProvisao
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

        tit.UsuarioInclusao = NF.UsuarioInclusao
        tit.UsuarioInclusaoData = NF.DataInclusao
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

        Sql = "	DECLARE " & vbCrLf & _
               "	@Diferenca numeric(18,2)," & vbCrLf & _
               "	@ValorTotal numeric(18,2)," & vbCrLf & _
               "	@Data varchar(10)," & vbCrLf & _
               "	@FPagto int" & vbCrLf

        '-- Na Compra dilui em todas as parcelas nas outras cobra na primeira parcela
        Sql &= "	set @ValorTotal =    " & Str(ValorAParcelar)

        Sql &= "	set @Data       = '" & Fixacao.Movimento.ToString("yyyy-MM-dd") & "'" & vbCrLf & _
               "	set @FPagto     =  " & Fixacao.CondicaoPagamento & vbCrLf

        '--Seleciona no banco a forma de pagamento e as Parcelas"
        Sql &= "	SELECT Pagamentos.Pagamento_Id," & vbCrLf & _
               "		   Pagamentos.Descricao," & vbCrLf & _
               "		   PagamentosXParcelas.Sequencia_Id, " & vbCrLf & _
               "		   Pagamentos.Parcelas," & vbCrLf & _
               "		   PagamentosXParcelas.Dias," & vbCrLf
        '--Soma a Data Base o Numero de dias referente ao numero da Parcela"
        Sql &= "		   convert(datetime,(@Data)) + PagamentosXParcelas.Dias as Vencimento," & vbCrLf
        '--Divide o valor total pelo numero de parcelas p/ descubrir o valor da parcela"
        Sql &= "		   round(@ValorTotal / Pagamentos.Parcelas, 2) as ValorParcela," & vbCrLf
        '-- Armazena o Valor Total para calcular a diferenca na divisao das parcelas"
        Sql &= "		   @ValorTotal as ValorTotal" & vbCrLf & _
               "	  INTO #Temp1 " & vbCrLf & _
               "	  FROM Pagamentos " & vbCrLf & _
               "	 INNER JOIN PagamentosXParcelas " & vbCrLf & _
               "		ON Pagamentos.Pagamento_Id = PagamentosXParcelas.Pagamento_Id " & vbCrLf & _
               "	 where Pagamentos.Pagamento_Id = @FPagto " & vbCrLf & _
               "	 order by PagamentosXParcelas.Sequencia_Id " & vbCrLf

        '	--Calcula o valor Parcelado para ver se ha diferenca com o valor total caso haja armazena o valor na variavel "
        Sql &= "	set @Diferenca = (Select top(1) ValorTotal - (ValorParcela * Parcelas) from #Temp1)" & vbCrLf


        '	--Atualiza o valor da primeira parcela acrescida do IPI + a Diferenca"
        Sql &= "	update #Temp1 set" & vbCrLf & _
               "	ValorParcela = ValorParcela + @Diferenca" & vbCrLf & _
               "	where #Temp1.Sequencia_Id = 1" & vbCrLf & _
               "	select Pagamento_Id, Descricao, Sequencia_Id, Parcelas, Dias, Vencimento, ValorParcela, ValorTotal from #Temp1" & vbCrLf

        ds = Banco.ConsultaDataSet(Sql, "Titulos")

        Dim n As Numerador = New Numerador(1)
        Dim NumeroTitulo As Integer = n.Sequencia + 1
        Dim parcela As Integer

        For Each Row As DataRow In ds.Tables("Titulos").Rows
            Dim Tit As New TituloV
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

            Tit.UsuarioInclusao = HttpContext.Current.Session("ssNomeUsuario")
            Tit.UsuarioInclusaoData = Now()
            Me.Add(Tit)
        Next
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

    'Apagar Depois
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
                For Each TitNotaDevolucao As Titulo In itemNotaDevolucao.Nota.VencimentosNota.Where(Function(s) s.CodigoProvisao = eProvisao.Previsao).OrderBy(Function(s) s.Prorrogacao)
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
                            TitProvisao.Historico = TitProvisao.Historico & vbCrLf & "** Titulo de: " & (TitProvisao.ValorDoDocumento + valorAuxNF).ToString("N2") & _
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
            For Each Tit As Titulo In NF.VencimentosPedido.Where(Function(s) s.CodigoProvisao = eProvisao.Previsao).OrderBy(Function(s) s.Prorrogacao)
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
                        TitProvisao.Historico = TitProvisao.Historico & vbCrLf & "** Titulo de: " & (TitProvisao.ValorDoDocumento + valorAuxNF).ToString("N2") & _
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

        Tit.UsuarioInclusao = NF.UsuarioInclusao
        Tit.UsuarioInclusaoData = NF.DataInclusao
    End Sub
#End Region

End Class

'*********************************************************************************************************************************************************************
'*************************************************** CLASSE BASE DE TITULO     ***************************************************************************************
'*********************************************************************************************************************************************************************
<Serializable()> _
Public Class TituloV
    Implements IBaseEntity

#Region "Uteis"
    Private _Carregando As Boolean
#End Region

#Region "Contrutor"
    Public Sub New()

    End Sub

    Public Sub New(ByVal pTitulo As Integer, Optional ByVal pReceberPagar As String = "")
        Dim sql As String

        sql = "SELECT * " & vbCrLf & _
              "  FROM (" & vbCrLf & _
              "         SELECT 'P' AS ReceberPagar, Registro_Id, Sequencia_Id, Provisao, Carteira, isnull(CarteiraAdto,'') AS CarteiraAdto, Tributo, Indexador, Moeda, TipoPagto, Situacao, Lote, Movimento, Vencimento, Prorrogacao, DataMoeda, Baixa, UnidadeDeNegocio, Empresa, EndEmpresa, " & vbCrLf & _
              "                Cliente, EndCliente," & vbCrLf & _
              "                BancoCliente, AgenciaCliente, DigitoAgenciaCliente, ContaCliente, DigitoContaCliente, isnull(TipoContaCliente,'') as TipoContaCliente, ContaContabilCliente," & vbCrLf & _
              "                EmpresaPagadora, EndEmpresaPagadora," & vbCrLf & _
              "                isnull(BancoPagador,0) as BancoPagador, AgenciaPagadora, DigitoAgenciaPagadora, ContaPagadora, DigitoContaPagadora, isnull(TipoContaPagadora,'') as TipoContaPagadora, ContaContabilPagadora," & vbCrLf & _
              "                 Cheque, Slips,  isnull(Recibo,'N') AS Recibo,  isnull(Aviso,'N') AS Aviso,  isnull(ReciboDeposito,'N') AS ReciboDeposito,  isnull(EmpresaPedido,'') as EmpresaPedido,  isnull(EndEmpresaPedido,0) as EndEmpresaPedido, isnull(Pedido,0) as Pedido, isnull(PedidoFixacao,0) as PedidoFixacao, isnull(Procuracao,0) as Procuracao," & vbCrLf & _
              "                ValorDoDocumento, Descontos, Deducoes, Juros, Acrescimos, ValorLiquido,isnull(MoedaValorDoDocumento,0) as MoedaValorDoDocumento, isnull(MoedaDescontos,0) as MoedaDescontos, isnull(MoedaDeducoes,0) as MoedaDeducoes, isnull(MoedaJuros,0) as MoedaJuros, isnull(MoedaAcrescimos,0) as MoedaAcrescimos, isnull(MoedaValorLiquido,0) as MoedaValorLiquido, Historico," & vbCrLf & _
              "                CodigoDeBarras, CodigoDigitado, Destinatario, EndDestinatario, isnull(NomeDoDestinatario,'') as NomeDoDestinatario, isnull(Destinacao,'') as Destinacao, solicitacao, UsuarioInclusao, isnull(UsuarioInclusaoData,Movimento) as UsuarioInclusaoData, isnull(UsuarioAlteracao,'') as UsuarioAlteracao, UsuarioAlteracaoData," & vbCrLf & _
              "                isnull(UsuarioCancelamento,'') as UsuarioCancelamento, UsuarioCancelamentoData, isnull(UsuarioLiberacao,'') as UsuarioLiberacao, UsuarioLiberacaoData, isnull(UsuarioBaixa,'') as UsuarioBaixa, UsuarioBaixaData, isnull(Grupado,'N') AS Grupado, isnull(RegistroMestre,0) as RegistroMestre, isnull(Observacoes,'') as Observacoes, isnull(SituacaoBancaria,0) as SituacaoBancaria, isnull(NumeroDoCheque,0) as NumeroDoCheque," & vbCrLf & _
              "                isnull(UsuarioLiberacaoBloqueio,'') as UsuarioLiberacaoBloqueio, UsuarioLiberacaoBloqueioDate, isnull(UsuarioLiberacaoPedido,'') as UsuarioLiberacaoPedido, UsuarioLiberacaoPedidoDate, isnull(UsuarioLiberacaoCheque,'') as UsuarioLiberacaoCheque," & vbCrLf & _
              "                UsuarioLiberacaoChequeDate, CarteiraDoTitulo, ContratoDeFinanciamento, DataEnvio, Ocorrencia, motivo, DigitoNossoNumero, NossoNumero, SituacaoRemessaBancaria," & vbCrLf & _
              "                isnull(ContratoBancario,'') AS ContratoBancario, isnull(CodigoDeBarraPreImpresso,0) AS CodigoDeBarraPreImpresso, TituloOrigem, ValorRecompra" & vbCrLf & _
              "           FROM ContasAPagar" & vbCrLf & _
              "          union all" & vbCrLf & _
              "         SELECT 'R' AS ReceberPagar, Registro_Id, Sequencia_Id, Provisao, Carteira, isnull(CarteiraAdto,'') AS CarteiraAdto, Tributo, Indexador, Moeda, TipoPagto, Situacao, Lote, Movimento, Vencimento, Prorrogacao, DataMoeda, Baixa, UnidadeDeNegocio, Empresa, EndEmpresa," & vbCrLf & _
              "                Cliente, EndCliente," & vbCrLf & _
              "                BancoCliente, AgenciaCliente, DigitoAgenciaCliente, ContaCliente, DigitoContaCliente, isnull(TipoContaCliente,'') as TipoContaCliente, ContaContabilCliente," & vbCrLf & _
              "                EmpresaPagadora, EndEmpresaPagadora," & vbCrLf & _
              "                isnull(BancoPagador,0) as BancoPagador, AgenciaPagadora, DigitoAgenciaPagadora, ContaPagadora, DigitoContaPagadora, isnull(TipoContaPagadora,'') as TipoContaPagadora, ContaContabilPagadora," & vbCrLf & _
              "                Cheque, Slips,  isnull(Recibo,'N') AS Recibo,  isnull(Aviso,'N') AS Aviso,  isnull(ReciboDeposito,'N') AS ReciboDeposito,  isnull(EmpresaPedido,'') as EmpresaPedido,  isnull(EndEmpresaPedido,0) as EndEmpresaPedido, isnull(Pedido,0) as Pedido, isnull(PedidoFixacao,0) as PedidoFixacao, isnull(Procuracao,0) as Procuracao," & vbCrLf & _
              "                ValorDoDocumento, Descontos, Deducoes, Juros, Acrescimos, ValorLiquido, isnull(MoedaValorDoDocumento,0) as MoedaValorDoDocumento, isnull(MoedaDescontos,0) as MoedaDescontos, isnull(MoedaDeducoes,0) as MoedaDeducoes, isnull(MoedaJuros,0) as MoedaJuros, isnull(MoedaAcrescimos,0) as MoedaAcrescimos, isnull(MoedaValorLiquido,0) as MoedaValorLiquido, Historico," & vbCrLf & _
              "                CodigoDeBarras, CodigoDigitado, Destinatario, EndDestinatario, isnull(NomeDoDestinatario,'') as NomeDoDestinatario, isnull(Destinacao,'') as Destinacao, solicitacao, UsuarioInclusao, isnull(UsuarioInclusaoData,Movimento) as UsuarioInclusaoData, isnull(UsuarioAlteracao,'') as UsuarioAlteracao, UsuarioAlteracaoData," & vbCrLf & _
              "                isnull(UsuarioCancelamento,'') as UsuarioCancelamento, UsuarioCancelamentoData, isnull(UsuarioLiberacao,'') as UsuarioLiberacao, UsuarioLiberacaoData, isnull(UsuarioBaixa,'') as UsuarioBaixa, UsuarioBaixaData, isnull(Grupado,'N') AS Grupado, isnull(RegistroMestre,0) as RegistroMestre, isnull(Observacoes,'') as Observacoes, isnull(SituacaoBancaria,0) as SituacaoBancaria, isnull(NumeroDoCheque,0) as NumeroDoCheque," & vbCrLf & _
              "                isnull(UsuarioLiberacaoBloqueio,'') as UsuarioLiberacaoBloqueio, UsuarioLiberacaoBloqueioDate, isnull(UsuarioLiberacaoPedido,'') as UsuarioLiberacaoPedido, UsuarioLiberacaoPedidoDate, isnull(UsuarioLiberacaoCheque,'') as UsuarioLiberacaoCheque," & vbCrLf & _
              "                UsuarioLiberacaoChequeDate, CarteiraDoTitulo, ContratoDeFinanciamento, DataEnvio, Ocorrencia, motivo, DigitoNossoNumero, NossoNumero, SituacaoRemessaBancaria," & vbCrLf & _
              "                isnull(ContratoBancario,'') AS ContratoBancario, isnull(CodigoDeBarraPreImpresso,0) AS CodigoDeBarraPreImpresso, TituloOrigem, ValorRecompra" & vbCrLf & _
              "           FROM ContasAReceber" & vbCrLf & _
              "        ) sb" & vbCrLf & _
              " WHERE sb.Registro_Id = " & pTitulo

        If Not String.IsNullOrWhiteSpace(pReceberPagar) Then
            sql &= " AND sb.ReceberPagar = '" & pReceberPagar & "'"
        End If

        Dim Banco As New AcessaBanco
        Dim ds As DataSet = Banco.ConsultaDataSet(sql, "Titulo")

        For Each row As DataRow In ds.Tables(0).Rows
            Me.Carregando = True
            Me.ReceberPagar = row("ReceberPagar")
            Me.Codigo = row("Registro_Id")
            Me.Sequencia = row("Sequencia_Id")

            Me.DataMoeda = row("DataMoeda")
            If Not IsDBNull(row("Baixa")) Then Me.Baixa = row("Baixa")

            Me.CodigoIndexador = row("Indexador")
            Me.CodigoMoeda = row("Moeda")
            Me.Movimento = row("Movimento")
            Me.Vencimento = row("Vencimento")
            Me.Prorrogacao = row("Prorrogacao")
            Me.CodigoEmpresaPedido = row("EmpresaPedido")
            Me.EndEmpresaPedido = row("EndEmpresaPedido")

            Me.CodigoPedidoFixacao = row("PedidoFixacao")
            Me.CodigoPedido = row("Pedido")

            Me.CodigoProvisao = row("Provisao")
            Me.CodigoCarteira = row("Carteira")
            Me.CodigoCarteiraAdto = row("CarteiraAdto")
            Me.Tributo = row("Tributo")
            Me.CodigoTipoPgto = row("TipoPagto")
            Me.CodigoSituacao = row("Situacao")
            Me.Lote = row("Lote")
            Me.CodigoEmpresa = row("Empresa")
            Me.EnderecoEmpresa = row("EndEmpresa")
            Me.CodigoUnidadeDeNegocio = row("UnidadeDeNegocio")
            Me.CodigoCliente = row("Cliente")
            Me.EndCliente = row("EndCliente")

            Me.CodigoBancoCliente = row("BancoCliente")
            Me.CodigoAgenciaCliente = row("AgenciaCliente")
            Me.DigitoAgenciaCliente = row("DigitoAgenciaCliente")
            Me.ContaCliente = row("ContaCliente")
            Me.DigitoContaCliente = row("DigitoContaCliente")
            Me.TipoDaContaCliente = row("TipoContaCliente")
            Me.ContaContabilCliente = row("ContaContabilCliente")

            Me.CodigoEmpresaPagadora = row("EmpresaPagadora")
            Me.EndEmpresaPagadora = row("EndEmpresaPagadora")
            Me.CodigoBancoPagador = row("BancoPagador")
            Me.CodigoAgenciaPagadora = row("AgenciaPagadora")
            Me.DigitoAgenciaPagadora = row("DigitoAgenciaPagadora")
            Me.ContaPagadora = row("ContaPagadora")
            Me.DigitoContaPagadora = row("DigitoContaPagadora")
            Me.TipoDaContaPagadora = row("TipoContaPagadora")
            Me.ContaContabilPagadora = row("ContaContabilPagadora")

            Me.Cheque = row("Cheque") = "S"
            Me.Slips = row("Slips") = "S"
            Me.Recibo = row("Recibo") = "S"
            Me.Aviso = row("Aviso") = "S"
            Me.ReciboDeposito = row("ReciboDeposito") = "S"
            Me.CodigoProcuracao = row("Procuracao")
            Me.ValorDoDocumento = row("ValorDoDocumento")
            Me.Descontos = row("Descontos")
            Me.Deducoes = row("Deducoes")
            Me.Juros = row("Juros")
            Me.Acrescimos = row("Acrescimos")
            _ValorLiquido = row("ValorLiquido")
            Me.MoedaValorDoDocumento = row("MoedaValorDoDocumento")
            Me.MoedaDescontos = row("MoedaDescontos")
            Me.MoedaDeducoes = row("MoedaDeducoes")
            Me.MoedaJuros = row("MoedaJuros")
            Me.MoedaAcrescimos = row("MoedaAcrescimos")
            _MoedaValorLiquido = row("MoedaValorLiquido")
            Me.Historico = row("Historico")
            If Not IsDBNull(row("CodigoDeBarras")) Then Me.CodigoDeBarras = row("CodigoDeBarras")
            If Not IsDBNull(row("CodigoDigitado")) Then Me.CodigoDigitado = row("CodigoDigitado") = "S"
            Me.CodigoDeBarrasPreImpresso = row("CodigoDeBarraPreImpresso")
            Me.CodigoDestinatario = row("Destinatario")
            Me.EndDestinatario = row("EndDestinatario")
            Me.NomeDoDestinatario = row("NomeDoDestinatario")
            Me.Destinacao = row("Destinacao")
            Me.Solicitacao = row("solicitacao")
            Me.Agrupado = row("Grupado") = "S"
            Me.RegistroMestre = row("RegistroMestre")
            If Not IsDBNull(row("Observacoes")) Then _Observacoes = row("Observacoes")
            Me.SituacaoBancaria = row("SituacaoBancaria")
            Me.NumeroDoCheque = row("NumeroDoCheque")
            Me.UsuarioLiberacaoBloqueio = row("UsuarioLiberacaoBloqueio")
            If Not IsDBNull(row("UsuarioLiberacaoBloqueioDate")) Then Me.UsuarioLiberacaoBloqueioDate = row("UsuarioLiberacaoBloqueioDate")
            Me.UsuarioLiberacaoPedido = row("UsuarioLiberacaoPedido")
            If Not IsDBNull(row("UsuarioLiberacaoPedidoDate")) Then Me.UsuarioLiberacaoPedidoDate = row("UsuarioLiberacaoPedidoDate")
            Me.UsuarioLiberacaoCheque = row("UsuarioLiberacaoCheque")
            If Not IsDBNull(row("UsuarioLiberacaoChequeDate")) Then Me.UsuarioLiberacaoChequeDate = row("UsuarioLiberacaoChequeDate")
            Me.ContratoBancario = row("ContratoBancario")
            Me.UsuarioInclusao = row("UsuarioInclusao")
            Me.UsuarioInclusaoData = row("UsuarioInclusaoData")
            Me.UsuarioAlteracao = row("UsuarioAlteracao")
            If Not IsDBNull(row("UsuarioAlteracaoData")) Then Me.UsuarioAlteracaoData = row("UsuarioAlteracaoData")
            Me.UsuarioCancelamento = row("UsuarioCancelamento")
            If Not IsDBNull(row("UsuarioCancelamentoData")) Then Me.UsuarioCancelamentoData = row("UsuarioCancelamentoData")
            Me.UsuarioLiberacao = row("UsuarioLiberacao")
            If Not IsDBNull(row("UsuarioLiberacaoData")) Then Me.UsuarioLiberacaoData = row("UsuarioLiberacaoData")
            Me.UsuarioBaixa = row("UsuarioBaixa")
            If Not IsDBNull(row("UsuarioBaixaData")) Then Me.UsuarioBaixaData = row("UsuarioBaixaData")
            Me.Carregando = False
        Next
    End Sub
#End Region

#Region "Fields"
    Private _IUD As String
    Private _TituloOriginal As TituloV

    Private _ReceberPagar As String
    Private _Codigo As Integer 'Registro_Id
    Private _Sequencia As Integer 'Sequencia_Id

    Private _CodigoProvisao As Integer 'Provisao
    Private _Provisao As TipoProvisao
    Private _DescricaoProvisao As String = ""

    Private _CodigoCarteira As String
    Private _Carteira As CarteiraFinanceira 'Carteira

    Private _CodigoCarteiraAdto As String
    Private _CarteiraAdto As CarteiraFinanceira 'Carteira


    Private _Tributo As String 'Tributo

    Private _CodigoIndexador As Integer
    Private _Indexador As Indexador 'Indexador
    Private _CodigoMoeda As Integer 'Moeda
    Private _Moeda As Moeda
    Private _DescricaoMoeda As String = ""
    Private _DataMoeda As Date 'DataMoeda

    Private _CodigoTipoPgto As Integer 'TipoPagto
    Private _TipoPagto As TipoDePagamento

    Private _CodigoSituacao As Integer 'Situacao
    Private _Situacao As Situacao
    Private _DescSituacao As String

    Private _Lote As Integer 'Lote
    Private _Movimento As Date 'Movimento
    Private _Vencimento As Date 'Vencimento
    Private _Prorrogacao As Date 'Prorrogacao

    Private _Baixa As Date 'Baixa

    Private _CodigoEmpresa As String
    Private _EndEmpresa As Integer 'EndEmpresa
    Private _Empresa As Cliente 'Empresa

    Private _ContratoBancario As String 'ContratoBancario

    Private _CodigoUnidadeDeNegocio As String 'UnidadeDeNegocio
    Private _UnidadeDeNegocio As Cliente 'UnidadeDeNegocio

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
    Private _TipoDaContaCliente As String 'Tipo da Conta C - Conta Corrente P - Pupança 
    Private _ContaContabilCliente As String 'ContaContabilCliente
    Private _ObjContaContabilCliente As PlanoDeConta

    Private _CodigoEmpresaPagadora As String 'EmpresaPagadora
    Private _EndEmpresaPagadora As Integer 'EndEmpresaPagadora
    Private _EmpresaPagadora As Cliente 'EmpresaPagadora

    Private _CodigoBancoPagador As Integer 'BancoPagador
    Private _BancoPagador As Banco
    Private _codigoAgenciaPagadora As String 'AgenciaPagadora
    Private _DigitoAgenciaPagadora As String 'DigitoAgenciaPagadora
    Private _AgenciaPagadora As Agencia
    Private _ContaPagadora As String 'ContaPagadora
    Private _DigitoContaPagadora As String 'DigitoContaPagadora
    Private _ContaContabilPagadora As String = "" 'ContaContabilPagadora
    Private _TipoDaContaPagadora As String 'Tipo da Conta C - Conta Corrente P - Pupança 
    Private _ObjContaContabilPagadora As PlanoDeConta

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
    Private _CodigoDeBarrasPreImpresso As Boolean

    Private _CodigoDestinatario As String
    Private _EndDestinatario As Integer
    Private _Destinatario As Cliente

    Private _NomeDoDestinatario As String
    Private _Destinacao As String
    Private _Solicitacao As Integer

    Private _Agrupado As Boolean
    Private _RegistroMestre As Integer
    Private _Observacoes As String
    Private _SituacaoBancaria As Integer
    Private _NumeroDoCheque As Integer

    Private _IndiceFixo As Boolean = False
    Private _IndiceTitulo As Decimal

    Private _UsuarioLiberacaoBloqueio As String
    Private _UsuarioLiberacaoBloqueioDate As Date
    Private _UsuarioLiberacaoPedido As String
    Private _UsuarioLiberacaoPedidoDate As Date
    Private _UsuarioLiberacaoCheque As String
    Private _UsuarioLiberacaoChequeDate As Date
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

    '*******************************************************
    '****************** Listas *****************************
    '*******************************************************
    Private _Adiantamentos As ListAdiantamento
    Private _BaixasAdiantamento As ListAdiantamentoBaixa

    '**** Nota Fiscal **************************************
    Private _NotaTitulo As NotaFiscalXTitulo

    '**** Documentos Financeiros *****************************
    Private _FinanceiroXDocumentos As ListFinanceiroXDocumento


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

    Public Property TituloOriginal As TituloV
        Get
            If _TituloOriginal Is Nothing AndAlso Codigo > 0 Then _TituloOriginal = New TituloV(Me.Codigo)
            Return _TituloOriginal
        End Get
        Set(value As TituloV)
            _TituloOriginal = value
        End Set
    End Property

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

    Public Property CodigoCarteiraAdto() As String
        Get
            Return _CodigoCarteiraAdto
        End Get
        Set(ByVal value As String)
            _CodigoCarteiraAdto = value
            _CarteiraAdto = Nothing
        End Set
    End Property

    Public Property CarteiraAdto() As CarteiraFinanceira
        Get
            If _CarteiraAdto Is Nothing AndAlso _CodigoCarteiraAdto IsNot Nothing AndAlso _CodigoCarteiraAdto.Length > 0 Then _CarteiraAdto = New CarteiraFinanceira(_CodigoCarteiraAdto)
            Return _CarteiraAdto
        End Get
        Set(ByVal value As CarteiraFinanceira)
            _CarteiraAdto = value
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
            If Not Carregando AndAlso Not Me.IndiceFixo Then
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
            If Not Carregando AndAlso Not Me.IndiceFixo AndAlso Me.CodigoIndexador > 0 Then
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
            Return _CodigoTipoPgto
        End Get
        Set(ByVal value As Integer)
            _CodigoTipoPgto = value
            _TipoPagto = Nothing
        End Set
    End Property

    Public Property TipoPagto() As TipoDePagamento
        Get
            If _TipoPagto Is Nothing And Me.CodigoTipoPgto > 0 Then _TipoPagto = New TipoDePagamento(Me.CodigoTipoPgto)
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
            If _CodigoBancoCliente > 0 And _CodigoAgenciaCliente.Length > 0 And _AgenciaCliente Is Nothing Then _AgenciaCliente = New Agencia(_CodigoBancoCliente, _CodigoAgenciaCliente, _DigitoAgenciaCliente)
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

    Public Property TipoDaContaCliente() As String
        Get
            Return _TipoDaContaCliente
        End Get
        Set(ByVal value As String)
            _TipoDaContaCliente = value
        End Set
    End Property

    Public ReadOnly Property BancoXConta() As BancosXContas
        Get
            Return New BancosXContas(CodigoBancoCliente, CodigoAgenciaCliente, DigitoAgenciaCliente, ContaCliente, DigitoContaCliente)
        End Get
        'Set(ByVal value As BancosXContas)
        '    newPropertyValue = value
        'End Set
    End Property





    Public Property ContaContabilCliente() As String
        Get
            Return _ContaContabilCliente
        End Get
        Set(ByVal value As String)
            _ContaContabilCliente = value
        End Set
    End Property

    Public Property ObjContaContabilCliente() As PlanoDeConta
        Get
            If _ObjContaContabilCliente Is Nothing And Me.ContaContabilPagadora.Length > 0 Then _ObjContaContabilCliente = New PlanoDeConta("", 0, Me.ContaContabilCliente)
            Return _ObjContaContabilCliente
        End Get
        Set(ByVal value As PlanoDeConta)
            _ObjContaContabilCliente = value
        End Set
    End Property

    '************************************
    '************************************
    '************************************

    Public Property CodigoEmpresaPagadora() As String
        Get
            Return _CodigoEmpresaPagadora
        End Get
        Set(ByVal value As String)
            _CodigoEmpresaPagadora = value
            _EmpresaPagadora = Nothing
        End Set
    End Property

    Public Property EndEmpresaPagadora() As Integer
        Get
            Return _EndEmpresaPagadora
        End Get
        Set(ByVal value As Integer)
            _EndEmpresaPagadora = value
            _EmpresaPagadora = Nothing
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
            If _CodigoBancoPagador > 0 And _codigoAgenciaPagadora.Length > 0 And _AgenciaPagadora Is Nothing Then _AgenciaPagadora = New Agencia(_CodigoBancoPagador, _codigoAgenciaPagadora, _DigitoAgenciaPagadora)
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

    Public Property TipoDaContaPagadora() As String
        Get
            Return _TipoDaContaPagadora
        End Get
        Set(ByVal value As String)
            _TipoDaContaPagadora = value
        End Set
    End Property

    Public Property ContaContabilPagadora() As String
        Get
            Return _ContaContabilPagadora
        End Get
        Set(ByVal value As String)
            _ContaContabilPagadora = value
            Me.ObjContaContabilPagadora = Nothing
        End Set
    End Property

    Public Property ObjContaContabilPagadora() As PlanoDeConta
        Get
            If _ObjContaContabilPagadora Is Nothing And Me.ContaContabilPagadora.Length > 0 Then _ObjContaContabilPagadora = New PlanoDeConta("", 0, Me.ContaContabilPagadora)
            Return _ObjContaContabilPagadora
        End Get
        Set(ByVal value As PlanoDeConta)
            _ObjContaContabilPagadora = value
        End Set
    End Property

    '************************************
    '************************************
    '************************************

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
            _ValorDoDocumento = value
            If Not Carregando AndAlso Moeda.Classificacao = eTiposMoeda.Oficial Then AtualizaLiquidoOficial()
        End Set
    End Property

    Public Property Descontos() As Decimal
        Get
            Return _Descontos
        End Get
        Set(ByVal value As Decimal)
            _Descontos = value
            If Not Carregando AndAlso Moeda.Classificacao = eTiposMoeda.Oficial Then AtualizaLiquidoOficial()
        End Set
    End Property

    Public Property Deducoes() As Decimal
        Get
            Return _Deducoes
        End Get
        Set(ByVal value As Decimal)
            _Deducoes = value
            If Not Carregando AndAlso Moeda.Classificacao = eTiposMoeda.Oficial Then AtualizaLiquidoOficial()
        End Set
    End Property

    Public Property Juros() As Decimal
        Get
            Return _Juros
        End Get
        Set(ByVal value As Decimal)
            _Juros = value
            If Not Carregando AndAlso Moeda.Classificacao = eTiposMoeda.Oficial Then AtualizaLiquidoOficial()
        End Set
    End Property

    Public Property Acrescimos() As Decimal
        Get
            Return _Acrescimos
        End Get
        Set(ByVal value As Decimal)
            _Acrescimos = value
            If Not Carregando AndAlso Moeda.Classificacao = eTiposMoeda.Oficial Then AtualizaLiquidoOficial()
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
            _MoedaValorDoDocumento = value
            If Not Carregando AndAlso Moeda.Classificacao = eTiposMoeda.MoedaEstrangeira Then AtualizaLiquidoMoeda()
        End Set
    End Property

    Public Property MoedaDescontos() As Decimal
        Get
            Return _MoedaDescontos
        End Get
        Set(ByVal value As Decimal)
            _MoedaDescontos = value
            If Not Carregando AndAlso Moeda.Classificacao = eTiposMoeda.MoedaEstrangeira Then AtualizaLiquidoMoeda()
        End Set
    End Property

    Public Property MoedaDeducoes() As Decimal
        Get
            Return _MoedaDeducoes
        End Get
        Set(ByVal value As Decimal)
            _MoedaDeducoes = value
            If Not Carregando AndAlso Moeda.Classificacao = eTiposMoeda.MoedaEstrangeira Then AtualizaLiquidoMoeda()
        End Set
    End Property

    Public Property MoedaJuros() As Decimal
        Get
            Return _MoedaJuros
        End Get
        Set(ByVal value As Decimal)
            _MoedaJuros = value
            If Not Carregando AndAlso Moeda.Classificacao = eTiposMoeda.MoedaEstrangeira Then AtualizaLiquidoMoeda()
        End Set
    End Property

    Public Property MoedaAcrescimos() As Decimal
        Get
            Return _MoedaAcrescimos
        End Get
        Set(ByVal value As Decimal)
            _MoedaAcrescimos = value
            If Not Carregando AndAlso Moeda.Classificacao = eTiposMoeda.MoedaEstrangeira Then AtualizaLiquidoMoeda()
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

    Public Property Agrupado() As Boolean
        Get
            Return _Agrupado
        End Get
        Set(ByVal value As Boolean)
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

    '*******************************************************
    '****************** Listas *****************************
    '*******************************************************
    Public Property Adiantamentos As ListAdiantamento
        Get
            If _Adiantamentos Is Nothing _
            AndAlso Me.CodigoCliente.Length > 0 _
            AndAlso Me.CodigoEmpresa.Length > 0 _
            AndAlso ( _
                     (Me.Carteira IsNot Nothing AndAlso Me.Carteira.BaixaAdiantamento) _
                     OrElse
                     (Me.CarteiraAdto IsNot Nothing AndAlso Me.CarteiraAdto.BaixaAdiantamento) _
                     ) Then

                Dim Conta As String = ""
                If Me.Carteira IsNot Nothing AndAlso Me.Carteira.BaixaAdiantamento Then Conta = Me.Carteira.CodigoContaCliente
                If Me.CarteiraAdto IsNot Nothing AndAlso Me.CarteiraAdto.BaixaAdiantamento Then Conta = Me.CarteiraAdto.CodigoContaCliente

                _Adiantamentos = New ListAdiantamento(Me.Empresa, Me.Cliente, 1, Me.Moeda.Classificacao, Me.CodigoPedido, Conta)
            End If
            Return _Adiantamentos
        End Get
        Set(value As ListAdiantamento)
            _Adiantamentos = value
        End Set
    End Property

    Public Property BaixasAdiantamento() As ListAdiantamentoBaixa
        Get
            If _BaixasAdiantamento Is Nothing Then _BaixasAdiantamento = New ListAdiantamentoBaixa(Me.Codigo)
            Return _BaixasAdiantamento
        End Get
        Set(ByVal value As ListAdiantamentoBaixa)
            _BaixasAdiantamento = value
        End Set
    End Property

    '** Nota Fiscal ************************************************************
    Public Property NotaTitulo As NotaFiscalXTitulo
        Get
            If _NotaTitulo Is Nothing And _Codigo > 0 Then
                _NotaTitulo = New NotaFiscalXTitulo(Me)
            End If
            Return _NotaTitulo
        End Get
        Set(value As NotaFiscalXTitulo)
            _NotaTitulo = value
        End Set
    End Property

    Public Property FinanceiroXDocumentos() As ListFinanceiroXDocumento
        Get
            If _FinanceiroXDocumentos Is Nothing Then _FinanceiroXDocumentos = New ListFinanceiroXDocumento(Me)
            Return _FinanceiroXDocumentos
        End Get
        Set(ByVal value As ListFinanceiroXDocumento)
            _FinanceiroXDocumentos = value
        End Set
    End Property

#End Region

#Region "Methods"

    Public Sub CalculaJuros(ByVal ate As Date)
        Dim dias As Double
        dias = ate.Subtract(Vencimento).Days
        If dias <= 0 Then Exit Sub

        Dim iq As Decimal 'taxa que eu quero 
        Dim it As Decimal 'taxa que eu tenho 
        Dim q As Decimal  'número de dias em que quero expressar a taxa 
        Dim t As Decimal  'número de dias em que a taxa que tenho está expressa 

        Dim TaxaAdto As Decimal 'tem q ser a taxa do adiantamento


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

    Public Function Salvar() As Boolean
        Dim Banco As New AcessaBanco
        Dim sqls As New ArrayList
        SalvarSql(sqls)
        Return Banco.GravaBanco(sqls)
    End Function

    Public Function SalvarSql(ByRef sqls As ArrayList, Optional ByVal UsaNumerador As Boolean = True) As ArrayList
        Dim strSQL As String
        Dim ObjBanco As New AcessaBanco

        'Recupera o Numero do Titulo e atualiza o mesmo na tabela Numerador
        Dim numTitulo As New Numerador(1)
        If UsaNumerador And Me.IUD = "I" Then
            _Codigo = numTitulo.Sequencia + 1
            sqls.Add(numTitulo.IncrementarNumeradorSql(1))
        End If

        Select Case Me.IUD
            Case "I"
                strSQL = "INSERT INTO " & IIf(Me.ReceberPagar = "P", "ContasAPagar", "ContasAReceber") & vbCrLf & _
                         "  (Registro_Id, Sequencia_Id, Provisao, Carteira, CarteiraAdto, Tributo, Indexador, Moeda, TipoPagto, Situacao, Lote, Movimento, Vencimento, Prorrogacao, DataMoeda, Baixa, " & vbCrLf & _
                         "   UnidadeDeNegocio, Empresa, EndEmpresa, Cliente, EndCliente, BancoCliente, AgenciaCliente, DigitoAgenciaCliente, ContaCliente, DigitoContaCliente, TipoContaCliente," & vbCrLf & _
                         "   ContaContabilCliente, EmpresaPagadora, EndEmpresaPagadora, BancoPagador, AgenciaPagadora, DigitoAgenciaPagadora, ContaPagadora, " & vbCrLf & _
                         "   DigitoContaPagadora, ContaContabilPagadora, Cheque, Slips, Recibo, Aviso, ReciboDeposito, EmpresaPedido, EndEmpresaPedido, Pedido, PedidoFixacao, " & vbCrLf & _
                         "   Procuracao, ValorDoDocumento, Descontos, Deducoes, Juros, Acrescimos, ValorLiquido, MoedaValorDoDocumento, MoedaDescontos, MoedaDeducoes, " & vbCrLf & _
                         "   MoedaJuros, MoedaAcrescimos, MoedaValorLiquido, Historico, CodigoDeBarras, CodigoDigitado, CodigoDeBarraPreImpresso, Destinatario, EndDestinatario, NomeDoDestinatario, Destinacao, " & vbCrLf & _
                         "   solicitacao, UsuarioInclusao, UsuarioInclusaoData, UsuarioAlteracao, UsuarioAlteracaoData, UsuarioCancelamento, UsuarioCancelamentoData, " & vbCrLf & _
                         "   UsuarioLiberacao, UsuarioLiberacaoData, UsuarioBaixa, UsuarioBaixaData, Grupado, RegistroMestre, Observacoes, SituacaoBancaria, NumeroDoCheque, " & vbCrLf & _
                         "   UsuarioLiberacaoBloqueio, UsuarioLiberacaoBloqueioDate, UsuarioLiberacaoPedido, UsuarioLiberacaoPedidoDate, " & vbCrLf & _
                         "   UsuarioLiberacaoCheque, UsuarioLiberacaoChequeDate, ContratoBancario) " & vbCrLf & _
                         " VALUES (" & Me.Codigo & "," & Me.Sequencia & ", " & Me.CodigoProvisao & "," & Me.CodigoCarteira.ToSqlNULL & "," & Me.CodigoCarteiraAdto.ToSqlNULL & ",'" & Me.Tributo & "'," & Me.CodigoIndexador & "," & Me.CodigoMoeda & ", " & Me.CodigoTipoPgto & ", " & Me.CodigoSituacao & "," & Me.Lote & ",'" & Me.Movimento.ToSqlDate & "','" & Me.Vencimento.ToSqlDate & "','" & Me.Prorrogacao.ToSqlDate & "','" & Me.DataMoeda.ToSqlDate & "'," & Me.Baixa.ToSqlDateNULL & "," & vbCrLf & _
                         "'" & Me.CodigoUnidadeDeNegocio & "','" & Me.CodigoEmpresa & "'," & Me.EnderecoEmpresa & ",'" & Me.CodigoCliente & "'," & Me.EndCliente & "," & Me.CodigoBancoCliente & ",'" & Me.CodigoAgenciaCliente & "','" & Me.DigitoAgenciaCliente & "','" & Me.ContaCliente & "','" & Me.DigitoContaCliente & "','" & Me.TipoDaContaCliente & "','" & Me.ContaContabilCliente & "','" & Me.CodigoEmpresaPagadora & "'," & Me.EndEmpresaPagadora & "," & Me.CodigoBancoPagador & ",'" & Me.CodigoAgenciaPagadora & "','" & Me.DigitoAgenciaPagadora & "','" & Me.ContaPagadora & "'," & "'" & Me.DigitoContaPagadora & "','" & Me.ContaContabilPagadora & "'," & vbCrLf & _
                         "'" & IIf(Me.Cheque, "S", "N") & "','" & IIf(Me.Slips, "S", "N") & "','" & IIf(Me.Recibo, "S", "N") & "','" & IIf(Me.Aviso, "S", "N") & "','" & IIf(Me.ReciboDeposito, "S", "N") & "'," & vbCrLf

                If Me.CodigoPedido = 0 Then
                    strSQL &= "NULL,NULL,NULL,"
                Else
                    strSQL &= "'" & Me.CodigoEmpresaPedido & "'," & Me.EndEmpresaPedido & "," & Me.CodigoPedido & ","
                End If

                strSQL &= Me.CodigoPedidoFixacao.ToSqlNULL & "," & Me.CodigoProcuracao.ToSqlNULL & "," & _
                          Str(Me.ValorDoDocumento) & "," & Str(Me.Descontos) & "," & Str(Me.Deducoes) & "," & Str(Me.Juros) & "," & Str(Me.Acrescimos) & "," & Str(Me.ValorLiquido) & "," & vbCrLf & _
                          Str(Me.MoedaValorDoDocumento) & "," & Str(Me.MoedaDescontos) & "," & Str(Me.MoedaDeducoes) & "," & Str(Me.MoedaJuros) & "," & Str(Me.MoedaAcrescimos) & "," & Str(Me.MoedaValorLiquido) & "," & vbCrLf & _
                          "'" & Me.Historico & "',"

                If Me.CodigoDeBarras IsNot Nothing Then
                    strSQL &= "'" & _CodigoDeBarras.Replace(" ", "").Replace(".", "") & "',"
                Else
                    strSQL &= "'',"
                End If

                strSQL &= IIf(Me.CodigoDigitado, "'S'", "'N'") & "," & IIf(Me.CodigoDeBarrasPreImpresso, 1, 0) & ",'" & Me.CodigoDestinatario & "'," & Me.EndDestinatario & ",'" & Me.NomeDoDestinatario & "','" & Me.Destinacao & "'," & Me.Solicitacao & "," & vbCrLf & _
                          "'" & Me.UsuarioInclusao & "'," & Me.UsuarioInclusaoData.ToSqlDateNULL & "," & vbCrLf & _
                          "'" & Me.UsuarioAlteracao & "'," & Me.UsuarioAlteracaoData.ToSqlDateNULL & "," & vbCrLf & _
                          "'" & Me.UsuarioCancelamento & "'," & Me.UsuarioCancelamentoData.ToSqlDateNULL & "," & vbCrLf & _
                          "'" & Me.UsuarioLiberacao & "'," & Me.UsuarioLiberacaoData.ToSqlDateNULL & "," & vbCrLf & _
                          "'" & Me.UsuarioBaixa & "'," & Me.UsuarioBaixaData.ToSqlDateNULL & "," & vbCrLf & _
                          "'" & IIf(Me.Agrupado, "S", "N") & "'," & Me.RegistroMestre & ",'" & Me.Observacoes & "'," & Me.SituacaoBancaria & "," & Me.NumeroDoCheque & "," & vbCrLf & _
                          "'" & Me.UsuarioLiberacaoBloqueio & "'," & Me.UsuarioLiberacaoBloqueioDate.ToSqlDateNULL & "," & vbCrLf & _
                          "'" & Me.UsuarioLiberacaoPedido & "'," & Me.UsuarioLiberacaoPedidoDate.ToSqlDateNULL & "," & vbCrLf & _
                          "'" & Me.UsuarioLiberacaoCheque & "'," & Me.UsuarioLiberacaoChequeDate.ToSqlDateNULL & "," & vbCrLf & _
                          "'" & Me.ContratoBancario & "')"

                sqls.Add(strSQL)
            Case "U"
                strSQL = "UPDATE " & IIf(Me.ReceberPagar = "P", "ContasAPagar", "ContasAReceber") & " SET" & vbCrLf & _
                         "   ValorDoDocumento       = " & Str(Me.ValorDoDocumento) & vbCrLf & _
                         "  ,Vencimento             ='" & Me.Vencimento.ToSqlDate & "'" & vbCrLf & _
                         "  ,Provisao               = " & Me.CodigoProvisao & vbCrLf & _
                         "  ,Carteira               = " & Me.CodigoCarteira.ToSqlNULL & vbCrLf & _
                         "  ,CarteiraAdto           = " & Me.CodigoCarteiraAdto.ToSqlNULL & vbCrLf & _
                         "  ,Tributo                ='" & Me.Tributo & "'" & vbCrLf & _
                         "  ,Indexador              = " & Me.CodigoIndexador & vbCrLf & _
                         "  ,Moeda                  = " & Me.CodigoMoeda & vbCrLf & _
                         "  ,TipoPagto              = " & Me.CodigoTipoPgto & vbCrLf & _
                         "  ,Situacao               = " & Me.CodigoSituacao & vbCrLf & _
                         "  ,Lote                   = " & Me.Lote & vbCrLf & _
                         "  ,Movimento              ='" & Me.Movimento.ToSqlDate & "'" & vbCrLf & _
                         "  ,Prorrogacao            ='" & Me.Prorrogacao.ToSqlDate & "'" & vbCrLf & _
                         "  ,DataMoeda              ='" & Me.DataMoeda.ToSqlDate & "'" & vbCrLf & _
                         "  ,Baixa                  = " & Me.Baixa.ToSqlDateNULL & vbCrLf & _
                         "  ,UnidadeDeNegocio       ='" & Me.CodigoUnidadeDeNegocio & "'" & vbCrLf & _
                         "  ,Empresa                ='" & Me.CodigoEmpresa & "'" & vbCrLf & _
                         "  ,EndEmpresa             = " & Me.EnderecoEmpresa & vbCrLf & _
                         "  ,Cliente                ='" & Me.CodigoCliente & "'" & vbCrLf & _
                         "  ,EndCliente             = " & Me.EndCliente & vbCrLf & _
                         "  ,BancoCliente           = " & Me.CodigoBancoCliente & vbCrLf & _
                         "  ,AgenciaCliente         ='" & Me.CodigoAgenciaCliente & "'" & vbCrLf & _
                         "  ,DigitoAgenciaCliente   ='" & Me.DigitoAgenciaCliente & "'" & vbCrLf & _
                         "  ,ContaCliente           ='" & Me.ContaCliente & "'" & vbCrLf & _
                         "  ,DigitoContaCliente     ='" & Me.DigitoContaCliente & "'" & vbCrLf & _
                         "  ,TipoDaConta            ='" & Me.TipoDaContaCliente & "'" & vbCrLf & _
                         "  ,ContaContabilCliente   ='" & Me.ContaContabilCliente & "'" & vbCrLf & _
                         "  ,EmpresaPagadora        ='" & Me.CodigoEmpresaPagadora & "'" & vbCrLf & _
                         "  ,EndEmpresaPagadora     = " & Me.EndEmpresaPagadora & vbCrLf & _
                         "  ,BancoPagador           = " & Me.CodigoBancoPagador & vbCrLf & _
                         "  ,AgenciaPagadora        ='" & Me.CodigoAgenciaPagadora & "'" & vbCrLf & _
                         "  ,DigitoAgenciaPagadora  ='" & Me.DigitoAgenciaPagadora & "'" & vbCrLf & _
                         "  ,ContaPagadora          ='" & Me.ContaPagadora & "'" & vbCrLf & _
                         "  ,DigitoContaPagadora    ='" & Me.DigitoContaPagadora & "'" & vbCrLf & _
                         "  ,ContaContabilPagadora  ='" & Me.ContaContabilPagadora & "'" & vbCrLf & _
                         "  ,Cheque                 ='" & IIf(Me.Cheque, "S", "N") & "'" & vbCrLf & _
                         "  ,Slips                  ='" & IIf(Me.Slips, "S", "N") & "'" & vbCrLf & _
                         "  ,Recibo                 ='" & IIf(Me.Recibo, "S", "N") & "'" & vbCrLf & _
                         "  ,Aviso                  ='" & IIf(Me.Aviso, "S", "N") & "'" & vbCrLf & _
                         "  ,ReciboDeposito         ='" & IIf(Me.ReciboDeposito, "S", "N") & "'" & vbCrLf

                If _CodigoPedido = 0 Then
                    strSQL &= "  ,EmpresaPedido    = NULL" & vbCrLf & _
                              "  ,EndEmpresaPedido = NULL" & vbCrLf & _
                              "  ,Pedido           = NULL" & vbCrLf
                Else
                    strSQL &= "  ,EmpresaPedido    ='" & Me.CodigoEmpresaPedido & "'" & vbCrLf & _
                              "  ,EndEmpresaPedido = " & Me.EndEmpresaPedido & vbCrLf & _
                              "  ,Pedido           = " & Me.CodigoPedido & vbCrLf
                End If
                strSQL &= "  ,PedidoFixacao          = " & Me.CodigoPedidoFixacao.ToSqlNULL & vbCrLf & _
                          "  ,Procuracao             = " & Me.CodigoProcuracao.ToSqlNULL & vbCrLf & _
                          "  ,Descontos              = " & Str(Me.Descontos) & vbCrLf & _
                          "  ,Deducoes               = " & Str(Me.Deducoes) & vbCrLf & _
                          "  ,Juros                  = " & Str(Me.Juros) & vbCrLf & _
                          "  ,Acrescimos             = " & Str(Me.Acrescimos) & vbCrLf & _
                          "  ,ValorLiquido           = " & Str(Me.ValorLiquido) & vbCrLf & _
                          "  ,MoedaValorDoDocumento  = " & Str(Me.MoedaValorDoDocumento) & vbCrLf & _
                          "  ,MoedaDescontos         = " & Str(Me.MoedaDescontos) & vbCrLf & _
                          "  ,MoedaDeducoes          = " & Str(Me.MoedaDeducoes) & vbCrLf & _
                          "  ,MoedaJuros             = " & Str(Me.MoedaJuros) & vbCrLf & _
                          "  ,MoedaAcrescimos        = " & Str(Me.MoedaAcrescimos) & vbCrLf & _
                          "  ,MoedaValorLiquido      = " & Str(Me.MoedaValorLiquido) & vbCrLf & _
                          "  ,Historico              ='" & Me.Historico & "'" & vbCrLf & _
                          "  ,CodigoDeBarras         ='" & Me.CodigoDeBarras.Replace(" ", "").Replace(".", "") & "'" & vbCrLf & _
                          "  ,CodigoDigitado         ='" & IIf(Me.CodigoDigitado, "S", "N") & "'" & vbCrLf & _
                          "  ,CodigoDeBarraPreImpresso = " & IIf(Me.CodigoDeBarrasPreImpresso, 1, 0) & vbCrLf & _
                          "  ,Destinatario           ='" & Me.CodigoDestinatario & "'" & vbCrLf & _
                          "  ,EndDestinatario        = " & Me.EndDestinatario & vbCrLf & _
                          "  ,NomeDoDestinatario     ='" & Me.NomeDoDestinatario & "'" & vbCrLf & _
                          "  ,Destinacao             ='" & Me.Destinacao & "'" & vbCrLf & _
                          "  ,solicitacao            = " & Me.Solicitacao & vbCrLf & _
                          "  ,UsuarioInclusao        ='" & Me.UsuarioInclusao & "'" & vbCrLf & _
                          "  ,UsuarioInclusaoData    = " & Me.UsuarioInclusaoData.ToSqlDateNULL & vbCrLf & _
                          "  ,UsuarioAlteracao       ='" & Me.UsuarioAlteracao & "'" & vbCrLf & _
                          "  ,UsuarioAlteracaoData   = " & Me.UsuarioAlteracaoData.ToSqlDateNULL & vbCrLf & _
                          "  ,UsuarioCancelamento    ='" & Me.UsuarioCancelamento & "'" & vbCrLf & _
                          "  ,UsuarioCancelamentoData= " & Me.UsuarioCancelamentoData.ToSqlDateNULL & vbCrLf & _
                          "  ,UsuarioLiberacao       ='" & Me.UsuarioLiberacao & "'" & vbCrLf & _
                          "  ,UsuarioLiberacaoData   = " & Me.UsuarioLiberacaoData.ToSqlDateNULL & vbCrLf & _
                          "  ,UsuarioBaixa           ='" & Me.UsuarioBaixa & "'" & vbCrLf & _
                          "  ,UsuarioBaixaData       = " & Me.UsuarioBaixaData.ToSqlDateNULL & vbCrLf & _
                          "  ,Grupado                ='" & IIf(Me.Agrupado, "S", "N") & "'" & vbCrLf & _
                          "  ,RegistroMestre         = " & Me.RegistroMestre & vbCrLf & _
                          "  ,Observacoes            ='" & Me.Observacoes & "'" & vbCrLf & _
                          "  ,SituacaoBancaria       = " & Me.SituacaoBancaria & vbCrLf & _
                          "  ,NumeroDoCheque         = " & Me.NumeroDoCheque & vbCrLf & _
                          "  ,UsuarioLiberacaoBloqueio    ='" & Me.UsuarioLiberacaoBloqueio & "'" & vbCrLf & _
                          "  ,UsuarioLiberacaoBloqueioDate= " & Me.UsuarioLiberacaoBloqueioDate.ToSqlDateNULL & vbCrLf & _
                          "  ,UsuarioLiberacaoPedido      ='" & Me.UsuarioLiberacaoPedido & "'" & vbCrLf & _
                          "  ,UsuarioLiberacaoPedidoDate  = " & Me.UsuarioLiberacaoPedidoDate.ToSqlDateNULL & vbCrLf & _
                          "  ,UsuarioLiberacaoCheque      ='" & Me.UsuarioLiberacaoCheque & "'" & vbCrLf & _
                          "  ,UsuarioLiberacaoChequeDate  = " & Me.UsuarioLiberacaoChequeDate.ToSqlDateNULL & vbCrLf & _
                          "  ,ContratoBancario            ='" & Me.ContratoBancario & "'" & vbCrLf & _
                          " WHERE Registro_Id = " & Me.Codigo & " " & vbCrLf
                sqls.Add(strSQL)
            Case "D"
                strSQL = "Delete " & IIf(Me.ReceberPagar = "P", "ContasAPagar", "ContasAReceber") & vbCrLf & _
                         " WHERE Registro_Id = " & Me.Codigo

                sqls.Add(strSQL)
            Case "C"
                strSQL = "UPDATE " & IIf(Me.ReceberPagar = "P", "ContasAPagar", "ContasAReceber") & " Set" & vbCrLf & _
                         "  Situacao = 3" & vbCrLf & _
                         " WHERE Registro_Id = " & Me.Codigo
                sqls.Add(strSQL)
        End Select


        Return sqls
    End Function
#End Region

End Class