Imports Microsoft.VisualBasic
Imports System.Collections.Generic
Imports System.Data
Imports Microsoft.VisualBasic.Financial
Imports NGS.Lib.Uteis
Imports System.Configuration

Namespace Novo
    '***********************************************************************************************************************************************************
    '**************************************************************** CLASSE LISTA TITULO **********************************************************************
    '***********************************************************************************************************************************************************
    <Serializable()> _
    Public Class ListTituloNovo
        Inherits List(Of Novo.TituloNovo)
        Implements IBaseEntity

#Region "Construtores"
        'Lista Vazia
        Public Sub New()

        End Sub

        'Lista com Where
        Public Sub New(ByVal Where As String, Optional ByVal pTitulo As Novo.TituloNovo = Nothing)
            If Not pTitulo Is Nothing Then
                _Titulo = pTitulo
            End If
            Dim sql As String
            sql = " Select Titulo_Id " & vbCrLf & _
                  "   from Titulos " & vbCrLf & _
                  "  Where " & Where & vbCrLf

            Dim ds As DataSet
            Dim Banco As New AcessaBanco
            ds = Banco.ConsultaDataSet(sql, "Titulos")
            For Each row As DataRow In ds.Tables(0).Rows
                Dim Tit As New Novo.TituloNovo(row("Titulo_Id"))
                Me.Add(Tit)
            Next
            Banco = Nothing
        End Sub

        'Lista Pedido
        'Sub New(ByVal CodigoEmpresa As String, ByVal CodigoEndEmpresa As Integer, ByVal CodigoPedido As Integer, Optional ByVal CodigoProvisao As Integer = 0, Optional CodigoFaturamento As Integer = 0)
        '    If CodigoPedido = 0 Then
        '        Me.Clear()
        '        Exit Sub
        '    End If

        '    Dim Banco As New AcessaBanco
        '    Dim ds As New DataSet
        '    Dim Sql As String = ""

        '    Sql = " Select Titulo_id" & vbCrLf & _
        '          "   From Titulos" & vbCrLf & _
        '          "  Where Pedido   = " & CodigoPedido & vbCrLf

        '    If CodigoFaturamento = 0 Then
        '        Sql &= "    and Situacao = 1" & vbCrLf
        '    End If

        '    If CodigoProvisao > 0 Then
        '        Sql &= "    and Provisao = " & CodigoProvisao & vbCrLf
        '    End If

        '    ds = Banco.ConsultaDataSet(Sql, "Titulos")

        '    If ds.Tables(0).Rows.Count = 0 Then
        '        Exit Sub
        '    End If

        '    For Each row As DataRow In ds.Tables(0).Rows
        '        Dim Tit As New Novo.TituloNovo(row("Titulo_id"))
        '        Me.Add(Tit)
        '    Next
        'End Sub

        Sub New(ByVal CodigoEmpresa As String, ByVal CodigoEndEmpresa As Integer, ByVal CodigoPedido As Integer, Optional ByVal CodigoProvisao As Integer = 0, Optional CodigoFaturamento As Integer = 0)
            If CodigoPedido = 0 Then
                Me.Clear()
                Exit Sub
            End If

            Dim Banco As New AcessaBanco
            Dim ds As New DataSet
            Dim Sql As String = ""
            Dim SqlWhere As String = ""

            SqlWhere = " Where Pedido = " & CodigoPedido & vbCrLf

            If CodigoFaturamento = 0 Then
                SqlWhere &= " and Situacao = 1" & vbCrLf
            End If

            If CodigoProvisao > 0 Then
                SqlWhere &= " and Provisao = " & CodigoProvisao & vbCrLf
            End If

            Sql = " Select Registro_Id " & vbCrLf & _
                  "   From ContasAPagar "

            ds = Banco.ConsultaDataSet(Sql + SqlWhere, "Titulos")

            If Not ds.Tables(0).Rows.Count = 0 Then
                Sql = " Select Registro_Id " & vbCrLf & _
                      "   From ContasAReceber "

                ds = Banco.ConsultaDataSet(Sql + SqlWhere, "Titulos")
            End If

            For Each row As DataRow In ds.Tables(0).Rows
                Dim Tit As New Novo.TituloNovo(row("Titulo_id"))
                Me.Add(Tit)
            Next
        End Sub

        'Lista Nota
        Sub New(ByRef pNF As NotaFiscal)
            Dim Banco As New AcessaBanco
            Dim ds As New DataSet
            Dim Sql As String = ""

            Sql = "SELECT Titulo_Id " & vbCrLf & _
                  "  FROM NotaFiscalXTitulo" & vbCrLf & _
                  " Where Empresa_Id      ='" & pNF.CodigoEmpresa & "'" & vbCrLf & _
                  "   and EndEmpresa_Id   = " & pNF.EnderecoEmpresa & vbCrLf & _
                  "   and Cliente_Id      ='" & pNF.CodigoCliente & "'" & vbCrLf & _
                  "   and EndCliente_Id   = " & pNF.EnderecoCliente & vbCrLf & _
                  "   and EntradaSaida_Id ='" & IIf(pNF.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "'" & vbCrLf & _
                  "   and Serie_Id        ='" & pNF.Serie & "'" & vbCrLf & _
                  "   and Nota_Id         = " & pNF.Codigo

            ds = Banco.ConsultaDataSet(Sql, "Titulos")

            For Each row As DataRow In ds.Tables(0).Rows
                Dim Tit As New Novo.TituloNovo(row("Titulo_Id"))
                Tit.NotaTitulo = New NotaFiscalXTitulo(pNF, Tit)
                Me.Add(Tit)
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
                      "                       and Cliente_Id      ='" & pNF.CodigoCliente & "'" & vbCrLf & _
                      "                       and EndCliente_Id   = " & pNF.EnderecoCliente & vbCrLf & _
                      "                       and EntradaSaida_Id ='" & IIf(pNF.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "'" & vbCrLf & _
                      "                       and Serie_Id        ='" & pNF.Serie & "'" & vbCrLf & _
                      "                       and Nota_Id         = " & pNF.Codigo & vbCrLf & _
                      "                    )" & vbCrLf & _
                      "  and Titulos.Provisao         = 2" & vbCrLf & _
                      "  and Titulos.UsuarioLiberacao = ''" & vbCrLf

                ds = Banco.ConsultaDataSet(Sql, "Titulos")

                For Each row As DataRow In ds.Tables(0).Rows
                    Dim Tit As New Novo.TituloNovo(row("Titulo_Id"))
                    Me.Add(Tit)
                Next
            End If
        End Sub

        'Lista Fixacao
        Sub New(ByRef pFX As [Lib].Negocio.Fixacao)
            Dim Banco As New AcessaBanco
            Dim ds As New DataSet
            Dim Sql As String = ""

            Sql = "SELECT Titulo_Id" & vbCrLf & _
                  "  FROM Titulos" & vbCrLf & _
                  " Where Empresa    ='" & pFX.ItemPedido.Pedido.CodigoEmpresa & "'" & vbCrLf & _
                  "   and EndEmpresa = " & pFX.ItemPedido.Pedido.EnderecoEmpresa & vbCrLf & _
                  "   and Pedido     ='" & pFX.ItemPedido.Pedido.Codigo & "'" & vbCrLf & _
                  "   and Fixacao    = " & pFX.Codigo & vbCrLf & _
                  "   and Situacao   = 1 " & vbCrLf

            ds = Banco.ConsultaDataSet(Sql, "Titulos")

            For Each row As DataRow In ds.Tables(0).Rows
                Dim Tit As New Novo.TituloNovo(row("Titulo_Id"))
                Me.Add(Tit)
            Next
        End Sub

        'Lista Frete
        Sub New(ByRef Fatura As FaturaDeFrete)
            Dim Banco As New AcessaBanco
            Dim ds As New DataSet
            Dim Sql As String = ""

            Sql = "SELECT Titulo_Id" & vbCrLf & _
                  "  FROM FaturaDeFreteXTitulo " & vbCrLf & _
                  " Where Empresa_Id         ='" & Fatura.CodigoEmpresa & "'" & vbCrLf & _
                  "   and EndEmpresa_Id      = " & Fatura.EnderecoEmpresa & vbCrLf & _
                  "   and Conveniado_Id      ='" & Fatura.CodigoConveniado & "'" & vbCrLf & _
                  "   and EndConveniado_Id   = " & Fatura.EnderecoConveniado & vbCrLf & _
                  "   and Fatura_Id          = " & Fatura.CodigoFatura

            ds = Banco.ConsultaDataSet(Sql, "Titulos")

            For Each row As DataRow In ds.Tables(0).Rows
                Dim Tit As New Novo.TituloNovo(row("Titulo_Id"))
                Me.Add(Tit)
            Next
        End Sub
#End Region

#Region "Fields"
        Private _NF As NotaFiscal
        Private _PD As Pedido
        Private _FX As Fixacao
        Private _FRETE As FaturaDeFrete

        Private _Titulo As Novo.TituloNovo
        Private _AdiantamentosAbertos As Novo.ListAdiantamentoNovo
        Private _ReajFinanceiro As ReajusteFinanceiro

        Private _msg As String
#End Region

#Region "Property"
        Public ReadOnly Property FinanceiroNovo As Boolean
            Get
                Return CBool(ConfigurationManager.AppSettings("financeiroNovo"))
            End Get
        End Property

        Public Property NF() As NotaFiscal
            Get
                Return _NF
            End Get
            Set(ByVal value As NotaFiscal)
                _NF = value
                'CarregarTitulos()
            End Set
        End Property

        Public Property PD As Pedido
            Get
                Return _PD
            End Get
            Set(ByVal value As Pedido)
                _PD = value
                'CarregarTitulos()
            End Set
        End Property

        Public Property FX As Fixacao
            Get
                Return _FX
            End Get
            Set(ByVal value As Fixacao)
                _FX = value
                'CarregarTitulos()
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

        Public Property Mensagem() As String
            Get
                Return _msg
            End Get
            Set(ByVal value As String)
                _msg = value
            End Set
        End Property

        'Adiantamentos abertos - vc esta na Lista de Titulos
        Public Property AdiantamentosAbertos() As Novo.ListAdiantamentoNovo
            Get
                If _AdiantamentosAbertos Is Nothing Then
                    Dim Par As New Hashtable

                    If Not PD Is Nothing Then
                        Par.Add("CodigoCliente", PD.Cliente.Codigo)
                        Par.Add("EndCliente", PD.Cliente.CodigoEndereco)
                        Par.Add("ConsolidarCliente", True)
                        Par.Add("SomenteComSaldo", True)

                        Par.Add("CodigoEmpresa", PD.CodigoEmpresa)
                        Par.Add("EndEmpresa", PD.EnderecoEmpresa)
                        Par.Add("CodigoPedido", PD.Codigo)

                        Par.Add("isTroca", False)
                        Par.Add("ContaAdiantamento", PD.SubOperacao.CodigoContaAdiantamento)

                        If PD.Troca Then
                            Par("isTroca") = True

                            Par.Add("CodigoEmpresaTroca", PD.CodigoEmpresaTroca)
                            Par.Add("EndEmpresaTroca", PD.EnderecoEmpresaTroca)
                            Par.Add("CodigoPedidoTroca", PD.CodigoPedidoTroca)
                        End If

                        _AdiantamentosAbertos = New Novo.ListAdiantamentoNovo(Par)

                    ElseIf Not NF Is Nothing Then
                        Par.Add("CodigoCliente", NF.Cliente.Codigo)
                        Par.Add("EndCliente", NF.Cliente.CodigoEndereco)
                        Par.Add("ConsolidarCliente", True)
                        Par.Add("SomenteComSaldo", True)

                        Par.Add("isTroca", False)
                        Par.Add("ContaAdiantamento", String.Empty)

                        If NF.Pedido IsNot Nothing Then
                            Par.Add("CodigoEmpresa", NF.Pedido.CodigoEmpresa)
                            Par.Add("EndEmpresa", NF.Pedido.EnderecoEmpresa)
                            Par.Add("CodigoPedido", NF.CodigoPedido)

                            Par("ContaAdiantamento") = NF.Pedido.SubOperacao.CodigoContaAdiantamento

                            If NF.Pedido.Troca Then
                                Par("isTroca") = True

                                Par.Add("CodigoEmpresaTroca", NF.Pedido.CodigoEmpresaTroca)
                                Par.Add("EndEmpresaTroca", NF.Pedido.EnderecoEmpresaTroca)
                                Par.Add("CodigoPedidoTroca", NF.Pedido.CodigoPedidoTroca)
                            End If
                        Else
                            Par.Add("CodigoEmpresa", NF.CodigoEmpresa)
                            Par.Add("EndEmpresa", NF.EnderecoEmpresa)
                            Par("ContaAdiantamento") = NF.SubOperacao.CodigoContaAdiantamento
                        End If

                        _AdiantamentosAbertos = New Novo.ListAdiantamentoNovo(Par)

                    ElseIf Not FX Is Nothing Then
                        Par.Add("CodigoCliente", FX.ItemPedido.Pedido.Cliente.Codigo)
                        Par.Add("EndCliente", FX.ItemPedido.Pedido.Cliente.CodigoEndereco)
                        Par.Add("ConsolidarCliente", True)
                        Par.Add("SomenteComSaldo", True)
                        Par.Add("isTroca", False)
                        Par.Add("ContaAdiantamento", FX.ItemPedido.Pedido.SubOperacao.CodigoContaAdiantamento)
                        _AdiantamentosAbertos = New Novo.ListAdiantamentoNovo(Par)
                    ElseIf Not _FRETE Is Nothing Then
                        Par.Add("CodigoEmpresa", _FRETE.CodigoEmpresa)
                        Par.Add("EndEmpresa", _FRETE.EnderecoEmpresa)
                        Par.Add("CodigoCliente", _FRETE.CodigoConveniado)
                        Par.Add("EndCliente", _FRETE.EnderecoConveniado)
                        Par.Add("ConsolidarCliente", True)
                        Par.Add("SomenteComSaldo", True)
                        Par.Add("isTroca", False)

                        _AdiantamentosAbertos = New Novo.ListAdiantamentoNovo(Par)
                    End If
                End If

                Return _AdiantamentosAbertos
            End Get
            Set(ByVal value As Novo.ListAdiantamentoNovo)
                _AdiantamentosAbertos = value
            End Set
        End Property

        Public ReadOnly Property Valor() As Decimal
            Get
                Return Me.Where(Function(s) s.CodigoSituacao = eSituacao.Normal And s.IUD <> "D").Sum(Function(s)
                                                                                                          Dim ret As Decimal
                                                                                                          If (s.Moeda.Classificacao = eTiposMoeda.Oficial) Then
                                                                                                              ret += s.Valores.EncargoValorDocumento.ValorOficial
                                                                                                          Else
                                                                                                              ret += s.Valores.EncargoValorDocumento.ValorMoeda
                                                                                                          End If
                                                                                                          Return ret
                                                                                                      End Function)
            End Get
        End Property

        'Reajusta Financeiro
        Public Property ReajFinanceiro As ReajusteFinanceiro
            Get
                Return _ReajFinanceiro 'vai alonso
            End Get
            Set(value As ReajusteFinanceiro)
                _ReajFinanceiro = value
            End Set
        End Property

#End Region

#Region "Métodos"
        Public Function Salvar(Optional ByVal UsaNumerador As Boolean = True) As Boolean
            Dim Banco As New AcessaBanco
            Dim sqls As New ArrayList

            sqls.Clear()
            Me.SalvarSql(sqls)

            If sqls.Count = 0 OrElse Banco.GravaBanco(sqls) Then
                Return True
            Else
                Return False
            End If
        End Function

        Public Sub SalvarSql(ByRef Sqls As ArrayList, Optional ByVal UsaNumerador As Boolean = True, Optional GravarComoEsta As Boolean = False)
            Dim i As Integer
            Dim NumeroTitulo As Integer
            Dim Num As New Numerador(1)

            If FinanceiroNovo AndAlso ReajFinanceiro IsNot Nothing AndAlso ReajFinanceiro.TitulosNew.Count > 0 Then
                NumeroTitulo = Num.Sequencia + 1
                For Each tit In ReajFinanceiro.TitulosNew
                    If tit.IUD = "I" Then
                        tit.Codigo = NumeroTitulo + i
                        i += 1
                    End If
                    tit.SalvarSql(Sqls, False, True)
                Next
                If i > 0 Then Sqls.Add(Num.IncrementarNumeradorSql(True, i))
            Else
                If Not UsaNumerador And Not GravarComoEsta Then
                    NumeroTitulo = Num.Sequencia
                End If

                For Each Tit As Novo.TituloNovo In Me
                    If Not Tit.IUD = Nothing Then
                        If Tit.IUD = "I" And Not UsaNumerador And Not GravarComoEsta Then
                            i += 1
                            Tit.Codigo = NumeroTitulo + i
                        End If
                        Tit.SalvarSql(Sqls, UsaNumerador, GravarComoEsta)
                    End If
                Next

                'Ajusta os titulos provisionados do pedido e cria o titulo em compensacao da nota
                If NF IsNot Nothing AndAlso Not NF.NFG AndAlso Not NF.SubOperacao.Classe = eClassesOperacoes.COMPLEMENTACOES Then
                    SalvarSqlAjustaTitulosEmPrevisaoEFazCompensacao(Sqls, NumeroTitulo, i)
                End If

                'Cria os titulos para baixar os adiantamentos
                If Me.AdiantamentosAbertos IsNot Nothing Then
                    CriaTitulosParaBaixaDeAdiantamento(Sqls, NumeroTitulo, i)
                End If

                If Not UsaNumerador And i > 0 Then Sqls.Add(Num.IncrementarNumeradorSql(True, i))
            End If
        End Sub

        Public Function Reprogramar(pData As DateTime) As Boolean
            Dim Codigos As String = ""
            For Each row In Me
                Codigos &= "," & row.Codigo
            Next
            Codigos = Codigos.Substring(1, Codigos.Length - 1)

            Dim sql As String = ""
            sql = "Update Titulos set " & vbCrLf & _
                  " Reprogramacao ='" & pData.ToString("yyyy-MM-dd") & "'" & vbCrLf & _
                  " Where Titulo_id in (" & Codigos & ")"
            Dim banco As New AcessaBanco
            Return banco.GravaBanco(sql)
        End Function

        Public Sub Parcelar(ByRef pFormaPagamento As Integer, ByVal pOrigem As String, Optional pTipoDePagamento As String = "", Optional valorParcela As Decimal = 0, Optional pQuantidade As Decimal = 0)
            Dim Banco As New AcessaBanco
            Dim ds As New DataSet
            Dim Sql As String = ""

            Dim ES As Negocio.eEntradaSaida
            Dim SO As New SubOperacao
            Dim Ped As Pedido = Nothing

            Dim pData As Date
            Dim pIndice As Decimal

            Dim IPI_ICMSST As Decimal = 0
            Dim ValorAParcelar As Decimal
            Dim eClasse As eClassesOperacoes
            Dim eMoeda As eTiposMoeda
            Dim CodigoIndexador As Integer

            '***********************************************************************************************************************************************
            '********************** Define o Valor do Parcelamento na Nota Fiscal / Pedido / FIXACAO em moeda Oficial ou Extrangeira ***********************
            '***********************************************************************************************************************************************
            Select Case pOrigem
                Case "NF"
                    Ped = NF.Pedido
                    ES = NF.SubOperacao.EntradaSaida
                    SO = NF.SubOperacao

                    pData = NF.Movimento
                    pIndice = NF.IndiceNota
                    For Each item As NotaFiscalXItem In NF.Itens
                        For Each Enc As NotaFiscalXItemXEncargo In item.Encargos
                            If Enc.Codigo = "IPI" Or Enc.Codigo = "ICMS-ST" Then IPI_ICMSST += Enc.Valor
                        Next
                    Next

                    If NF.Pedido.Moeda Is Nothing OrElse NF.Pedido.Moeda.Classificacao = eTiposMoeda.Oficial Then
                        ValorAParcelar = NF.TotalNota - NF.Titulos.AdiantamentosAbertos.Sum(Function(s) s.VlrBaixa)
                        eMoeda = eTiposMoeda.Oficial
                    Else
                        ValorAParcelar = Math.Round(NF.TotalNota / NF.IndiceNota, 2) - NF.Titulos.AdiantamentosAbertos.Sum(Function(s) s.VlrBaixa)
                        IPI_ICMSST = Math.Round(IPI_ICMSST / NF.IndiceNota, 2)
                        eMoeda = eTiposMoeda.MoedaEstrangeira
                    End If

                    If NF.Pedido.Moeda IsNot Nothing Then
                        CodigoIndexador = NF.Pedido.CodigoIndexador
                    Else
                        CodigoIndexador = 0 '2
                    End If
                    eClasse = NF.SubOperacao.Classe

                Case "PD"
                    Ped = PD
                    ES = PD.SubOperacao.EntradaSaida
                    SO = PD.SubOperacao

                    pData = PD.DataPedido
                    pIndice = PD.IndiceFixado

                    If PD.Moeda.Classificacao = eTiposMoeda.Oficial Then
                        ValorAParcelar += PD.Itens.LiquidoOficial
                    Else
                        ValorAParcelar += PD.Itens.LiquidoMoeda
                    End If

                    eClasse = PD.SubOperacao.Classe
                    eMoeda = PD.Moeda.Classificacao
                    CodigoIndexador = PD.CodigoIndexador
                    'If PD.MomentoFinanceiro <> eTipoFaturamento.Pedido AndAlso _
                    '    PD.MomentoFinanceiro <> eTipoFaturamento.Lote AndAlso _
                    '    PD.MomentoFinanceiro <> eTipoFaturamento.Peridiocidade Then
                    '    PD.CodigoCondicaoPagamento = pFormaPagamento
                    'End If

                Case "FX"
                    Ped = FX.ItemPedido.Pedido
                    ES = FX.SubOperacao.EntradaSaida
                    SO = FX.SubOperacao

                    pData = FX.Movimento
                    pIndice = IIf(FX.IndiceFixado = 0, FX.ItemPedido.Pedido.IndiceFixado, FX.IndiceFixado)

                    For Each encfx In FX.Encargos
                        If encfx.CodigoEncargo = "LIQUIDO" Then
                            If FX.ItemPedido.Pedido.Moeda.Classificacao = eTiposMoeda.Oficial Then
                                ValorAParcelar += encfx.ValorOficial
                            Else
                                ValorAParcelar += encfx.ValorMoeda
                            End If
                        End If
                    Next
                    eClasse = FX.ItemPedido.Pedido.SubOperacao.Classe
                    eMoeda = FX.ItemPedido.Pedido.Moeda.Classificacao
                    CodigoIndexador = FX.ItemPedido.Pedido.CodigoIndexador
            End Select

            'Faturamento por Lote
            If valorParcela > 0 Then ValorAParcelar = valorParcela

            If ValorAParcelar <= 0 Then
                Select Case pOrigem
                    Case "NF" : _msg = "Valor da Nota está zerado"
                    Case "PD" : _msg = "Valor do Pedido está zerado"
                    Case "FX" : _msg = "Valor da Fixacao está zerado"
                End Select
                Exit Sub
            End If

            If pIndice = 0 Then
                Dim c As New Cotacao(CodigoIndexador, pData)
                pIndice = c.Indice
            End If

            '***********************************************************************************************************************
            '****************************************** SQL que Realiza o Parcelamento *********************************************
            '***********************************************************************************************************************
            Sql &= "	DECLARE " & vbCrLf & _
                   "	@Diferenca numeric(18,2)," & vbCrLf & _
                   "	@ValorIPI numeric(18,2)," & vbCrLf & _
                   "	@ValorTotal numeric(18,2)," & vbCrLf & _
                   "	@Data varchar(10)," & vbCrLf & _
                   "	@FPagto int" & vbCrLf

            '--Informa o valor do IPI a ser cobrado"
            '-- Na Compra dilui em todas as parcelas nas outras cobra na primeira parcela
            If eClasse = eClassesOperacoes.COMPRAS Then
                Sql &= "	set @ValorIPI   =  0 " & vbCrLf & _
                       "	set @ValorTotal =    " & Str(ValorAParcelar)
            Else
                Sql &= "	set @ValorIPI   =  " & Str(IPI_ICMSST) & vbCrLf & _
                       "	set @ValorTotal =  " & Str(ValorAParcelar - (IPI_ICMSST)) & vbCrLf
            End If

            Sql &= "	set @Data       = '" & pData.ToSqlDate() & "'" & vbCrLf & _
                   "	set @FPagto     =  " & pFormaPagamento.ToString & vbCrLf

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
            Sql &= "	update #temp1 set" & vbCrLf & _
                   "	ValorParcela = ValorParcela + @ValorIPI + @Diferenca" & vbCrLf & _
                   "	where #temp1.Sequencia_Id = 1" & vbCrLf & _
                   "	select Pagamento_id, Descricao, Sequencia_id, Parcelas, Dias, Vencimento, ValorParcela, ValorTotal + @ValorIPI as ValorTotal from #temp1" & vbCrLf

            ds = Banco.ConsultaDataSet(Sql, "Titulos")

            Dim n As Numerador = New Numerador(1)
            Dim NumeroTitulo As Integer = n.Sequencia + IIf(valorParcela > 0, Me.Count + 1, 1)
            Dim parcela As Integer

            If Not valorParcela > 0 Then Me.Clear()

            For Each Row As DataRow In ds.Tables("Titulos").Rows
                Dim T As New TituloNovo
                T.DesligarControles()

                T.IUD = "I"
                T.ReceberPagar = IIf(ES = eEntradaSaida.Entrada, "P", "R")

                T.Codigo = NumeroTitulo
                NumeroTitulo += 1
                parcela += 1
                If Not String.IsNullOrWhiteSpace(pTipoDePagamento) Then
                    T.CodigoTipoPgto = pTipoDePagamento
                Else
                    T.CodigoTipoPgto = 0
                End If

                If Ped.Moeda IsNot Nothing Then
                    T.CodigoUnidadeDeNegocio = Ped.CodigoUnidadeNegocio
                    T.CodigoEmpresa = Ped.CodigoEmpresa
                    T.EnderecoEmpresa = Ped.EnderecoEmpresa
                    T.Pedido = Ped
                    'T.CodigoPedido = Ped.Codigo
                Else
                    T.CodigoUnidadeDeNegocio = NF.CodigoUnidadeDeNegocio
                    T.CodigoEmpresa = NF.CodigoEmpresa
                    T.EnderecoEmpresa = NF.EnderecoEmpresa
                    'T.Pedido = Ped
                End If


                If pOrigem = "FX" Then T.CodigoFixacao = FX.Codigo
                If pOrigem = "NF" Then
                    T.CodigoFixacao = NF.Itens(0).CodigoFixacao 'Pode ser zero se a nota nao for proviniente de uma fixacao
                    T.NotaTitulo = New NotaFiscalXTitulo(NF, T)
                End If

                T.CodigoProvisao = IIf(pOrigem = "PD" Or pOrigem = "FX", 2, 3) 'Se o Parcelamento for proveniente de um parcelamento de pedido ou de uma fixacao entao ele é criado como previsao

                If Ped.Moeda IsNot Nothing Then
                    T.CodigoMoeda = Ped.CodigoMoeda
                    T.CodigoIndexador = Ped.CodigoIndexador
                Else
                    T.CodigoMoeda = 1
                    T.CodigoIndexador = 2
                End If

                T.CodigoSituacao = 1

                T.Movimento = pData

                If Ped.Moeda IsNot Nothing AndAlso Ped.MomentoFinanceiro = eTipoFaturamento.Lote Then
                    'Pedido com faturamento em lote.
                    T.Vencimento = Funcoes.ValidaDataUtil(Ped.CodigoEmpresa, Ped.EnderecoEmpresa, Ped.DataVencimentoPedido)
                    T.Quantidade = pQuantidade
                ElseIf Ped.Moeda IsNot Nothing AndAlso Not Ped.Troca Then
                    T.Vencimento = Funcoes.ValidaDataUtil(Ped.CodigoEmpresa, Ped.EnderecoEmpresa, CDate(Row("Vencimento")))
                ElseIf NF IsNot Nothing AndAlso Not NF.Pedido.Troca Then
                    T.Vencimento = Funcoes.ValidaDataUtil(NF.CodigoEmpresa, NF.EnderecoEmpresa, CDate(Row("Vencimento")))
                Else
                    T.Vencimento = T.Movimento
                End If

                T.Reprogramacao = T.Vencimento
                T.DataMoeda = pData
                T.DataBaixa = T.Vencimento

                If Ped.Moeda IsNot Nothing Then
                    T.CodigoCliFor = Ped.CodigoCliente
                    T.EnderecoCliFor = Ped.EnderecoCliente

                    T.CodigoUnidadeDeNegocioRecPag = Ped.CodigoUnidadeNegocio
                    T.CodigoEmpresaRecPag = Ped.CodigoEmpresa
                    T.EndEmpresaRecPag = Ped.EnderecoEmpresa

                    T.CodigoContaContabilCliFor = SO.CodigoGrupoContas
                    T.CodigoContaContabilRecPag = Ped.Empresa.Empresa.CodigoContaGrupoBanco
                Else
                    T.CodigoCliFor = NF.CodigoCliente
                    T.EnderecoCliFor = NF.EnderecoCliente

                    T.CodigoUnidadeDeNegocioRecPag = NF.CodigoUnidadeDeNegocio
                    T.CodigoEmpresaRecPag = NF.CodigoEmpresa
                    T.EndEmpresaRecPag = NF.EnderecoEmpresa

                    T.CodigoContaContabilCliFor = SO.CodigoGrupoContas
                    T.CodigoContaContabilRecPag = NF.Empresa.Empresa.CodigoContaGrupoBanco
                End If


                T.IndiceTitulo = pIndice

                If eMoeda = eTiposMoeda.Oficial Then
                    T.Valores.EncargoValorDocumento.ValorOficial = Row("ValorParcela")
                Else
                    T.Valores.EncargoValorDocumento.ValorMoeda = Row("ValorParcela")
                End If

                T.ValorPrevisao = Row("ValorParcela")

                Select Case pOrigem
                    Case "NF"
                        T.Historico = "REF. NF " & NF.Codigo & "-" & NF.Serie & "-" & IIf(ES = eEntradaSaida.Saida, "Saida", "Entrada") & ", Parcela " & parcela & "/" & Row("Parcelas") & ", Pedido: " & NF.CodigoPedido & " / " & NF.Cliente.Nome
                        If Ped.Titulos.Where(Function(s) s.CodigoProvisao = 2).Count > 0 Then T.CodigoTituloOrigem = Ped.Titulos.Where(Function(s) s.CodigoProvisao = 2).First.Codigo
                    Case "PD" : T.Historico = "REF. Pedido: " & PD.Codigo & " / " & PD.Cliente.Nome & "-" & IIf(PD.SubOperacao.EntradaSaida = eEntradaSaida.Saida, "Saida", "Entrada") & ", Parcela " & parcela & "/" & Row("Parcelas") & " "
                    Case "FX" : T.Historico = "REF. A Fixacao " & FX.Codigo & " do Pedido: " & Ped.Codigo & " / " & Ped.Cliente.Nome & ", Parcela " & parcela & "/" & Row("Parcelas") & " "
                End Select

                T.LigarControles()
                Me.Add(T)
            Next
            'Para
            If eMoeda = eTiposMoeda.MoedaEstrangeira Then
                Dim vlrFinanceiro As Decimal = Me.Sum(Function(s) s.Valores.EncargoValorDocumento.ValorOficial)
                Dim vlrPedido As Decimal = PD.Itens.Sum(Function(s) s.Encargos.Where(Function(j) j.CodigoEncargo = "LIQUIDO").Sum(Function(i) i.ValorOficial))
                Dim vlrDiferenca = vlrFinanceiro - vlrPedido

                If vlrDiferenca > 0 Then
                    Me(0).DesligarControles()
                    Me(0).Valores.EncargoValorDocumento.ValorOficial -= vlrDiferenca
                    Me(0).Valores.EncargoValorLiquido.ValorOficial -= vlrDiferenca
                    Me(0).LigarControles()
                End If
            End If
        End Sub

        Public Function AJustaParcelasDaLista(ByVal Apartir As Integer, ByVal ValorOriginal As Decimal, ByVal ValorNovo As Decimal) As String
            Dim _tipoMoeda As eTiposMoeda

            _tipoMoeda = Me(Apartir).Pedido.Moeda.Classificacao
            'ULTIMA LINHA

            If (Apartir = Me.Count - 1) Then
                If _tipoMoeda = eTiposMoeda.Oficial Then
                    Me(Apartir).Valores.EncargoValorDocumento.ValorOficial = ValorOriginal
                Else
                    Me(Apartir).Valores.EncargoValorDocumento.ValorMoeda = ValorOriginal
                End If
                Return "Parcela Unica ou Ultima Parcela nao pode ser alterada"
            End If

            Dim saldo As Decimal = ValorOriginal - ValorNovo
            Dim numParcelas As Integer = (Me.Count - 1 - Apartir)
            Dim parcelas As Decimal = Math.Round(saldo / numParcelas, 2)
            Dim diferenca As Decimal = saldo - (parcelas * numParcelas)

            saldo = 0
            If (Apartir < Me.Count - 1) And ValorNovo > ValorOriginal Then

                For i As Integer = Apartir + 1 To Me.Count - 1
                    saldo += IIf(_tipoMoeda = eTiposMoeda.Oficial, Me(i).Valores.EncargoValorDocumento.ValorOficial, Me(i).Valores.EncargoValorDocumento.ValorMoeda)
                Next

                If saldo <= ValorNovo - ValorOriginal Then
                    If _tipoMoeda = eTiposMoeda.Oficial Then
                        Me(Apartir).Valores.EncargoValorDocumento.ValorOficial = ValorOriginal
                    Else
                        Me(Apartir).Valores.EncargoValorDocumento.ValorMoeda = ValorOriginal
                    End If
                    Return "Valor Informado ultrapassa o Valor da transacao"
                End If
            End If

            For i As Integer = Apartir + 1 To Me.Count - 1
                If _tipoMoeda = eTiposMoeda.Oficial Then
                    If i = Me.Count - 1 Then
                        Me(i).Valores.EncargoValorDocumento.ValorOficial += parcelas + diferenca
                    Else
                        Me(i).Valores.EncargoValorDocumento.ValorOficial += parcelas
                    End If
                Else
                    If i = Me.Count - 1 Then
                        Me(i).Valores.EncargoValorDocumento.ValorMoeda += parcelas + diferenca
                    Else
                        Me(i).Valores.EncargoValorDocumento.ValorMoeda += parcelas
                    End If
                End If
            Next
            Return ""
        End Function

        Public Sub ParcelarNotaProvisao(Optional pCodigoFixacao As Integer = 0) 'Somente Para quanto a Origem for Nota Fiscal
            If NF.TotalNota = 0 Then Exit Sub
            Dim SaldoAProvisionar As Decimal
            Dim eMoeda As eTiposMoeda
            Dim eClasse As eClassesOperacoes
            Dim CodigoIndexador As Integer

            If NF.Pedido Is Nothing OrElse NF.Pedido.Moeda.Classificacao = eTiposMoeda.Oficial Then
                SaldoAProvisionar = NF.TotalNota - NF.Titulos.AdiantamentosAbertos.Sum(Function(s) s.VlrBaixa)
                eMoeda = eTiposMoeda.Oficial
            Else
                SaldoAProvisionar = Math.Round(NF.TotalNota / NF.IndiceNota, 2) - NF.Titulos.AdiantamentosAbertos.Sum(Function(s) s.VlrBaixa)
                eMoeda = eTiposMoeda.MoedaEstrangeira
            End If

            CodigoIndexador = NF.Pedido.CodigoIndexador
            eClasse = NF.SubOperacao.Classe

            If SaldoAProvisionar <= 0 Then
                _msg = "Valor da Nota está zerado"
                Exit Sub
            End If

            Dim n As Numerador = New Numerador(1)
            Dim NumeroTitulo As Integer = n.Sequencia

            Me.Clear()
            For Each Row As Novo.TituloNovo In NF.Pedido.Titulos.Where(Function(s) s.CodigoProvisao = 2 And s.CodigoFixacao = pCodigoFixacao).OrderBy(Function(s) s.Reprogramacao).ThenBy(Function(s) s.Codigo)
                Dim T As New TituloNovo(Row.Codigo)
                If T.Valores.Count = 0 Then 'usei só pra instanciar a lista
                    _msg = "Titulos nao contem valores"
                    Exit Sub
                End If
                T.DesligarControles()

                T.IUD = "I"
                NumeroTitulo += 1
                T.Codigo = NumeroTitulo
                T.TituloOrigem = Row

                T.NotaTitulo = New NotaFiscalXTitulo(NF, T)
                T.CodigoProvisao = 3

                T.Movimento = NF.Movimento
                T.Pedido = NF.Pedido

                If T.Reprogramacao < NF.Movimento Then
                    T.Reprogramacao = NF.Movimento
                End If

                T.IndiceTitulo = Row.IndiceTitulo

                If Not T.CodigoIndexador = 99 Then
                    Dim c As New Cotacao(CodigoIndexador, T.Movimento)
                    T.IndiceTitulo = c.Indice
                End If

                T.LigarControles()

                If SaldoAProvisionar >= Row.Valores.EncargoValorDocumento.Valor Then
                    SaldoAProvisionar -= Row.Valores.EncargoValorDocumento.Valor
                Else
                    T.Valores.EncargoValorDocumento.Valor = IIf((T.Valores.EncargoValorDocumento.Valor - SaldoAProvisionar = 0.01), (SaldoAProvisionar + 0.01), SaldoAProvisionar)
                    SaldoAProvisionar = 0
                End If

                'faturamento por lote
                If NF.Pedido.MomentoFinanceiro = eTipoFaturamento.Lote Then
                    T.Reprogramacao = Funcoes.ValidaDataUtil(NF.CodigoEmpresa, NF.EnderecoEmpresa, NF.Movimento.AddDays(NF.Pedido.CondicaoPagamento.Periodo(0)))
                    T.Quantidade = Math.Round(T.Valores.EncargoValorDocumento.Valor * NF.TotalQuantidadeFiscal / NF.TotalNota, 1)
                    'Faturamento por Periodicidade
                ElseIf NF.Pedido.MomentoFinanceiro = eTipoFaturamento.Peridiocidade Then
                    Dim dataFaturamento As DateTime = NF.Movimento
                    'Realiza o faturamento por peridiocidade.
                    Select Case NF.Pedido.Peridiocidade
                        Case ePeriodicidade.Semanal
                            If dataFaturamento.DayOfWeek = DayOfWeek.Monday Then dataFaturamento = dataFaturamento.AddDays(1)
                            While dataFaturamento.DayOfWeek <> DayOfWeek.Monday
                                dataFaturamento = dataFaturamento.AddDays(1)
                            End While
                            'Aplica a condição de pagamento.
                        Case ePeriodicidade.Quinzenal
                            'para dia 1 ao 15 do mês fatura no dia 16
                            While dataFaturamento.Day < 16
                                dataFaturamento = dataFaturamento.AddDays(1)
                            End While
                            'para dia 16 ao 30/31 do mês fatura no próximo mês
                            While dataFaturamento.Day > 1
                                dataFaturamento = dataFaturamento.AddDays(1)
                            End While
                        Case ePeriodicidade.Mensal
                            dataFaturamento = dataFaturamento.AddMonths(1)
                            dataFaturamento = Convert.ToDateTime("01/" & dataFaturamento.Month & "/" & dataFaturamento.Year)
                        Case Else
                            Throw New Exception("Erro no faturamento por periodicidade! :(")
                    End Select

                    T.Reprogramacao = Funcoes.ValidaDataUtil(NF.CodigoEmpresa, NF.EnderecoEmpresa, dataFaturamento.AddDays(NF.Pedido.CondicaoPagamento.Periodo(0) - 1))

                End If

                T.DataMoeda = T.Reprogramacao
                T.DataBaixa = T.Reprogramacao

                T.Historico = "REF. NF " & NF.Codigo & "-" & NF.Serie & " de " & IIf(NF.EntradaSaida = eEntradaSaida.Saida, "Saida", "Entrada") & " do Pedido: " & NF.CodigoPedido & " / " & NF.Cliente.Nome & IIf(T.CodigoFixacao > 0, " - Fixacao Num. " & T.CodigoFixacao, "")
                Me.Add(T)

                If SaldoAProvisionar = 0 Then Exit Sub
            Next
        End Sub

        'Ajustas titulo apos Complementos e estorno de pedido
        Public Function AjustaParcelasComplementoEstornoPedido(ByVal pValor As Decimal, CE As eTiposLancamentosPedidos) As Boolean
            Try
                Dim codigoProvisao As Integer = IIf(Not PD.FinanceiroNovo AndAlso PD.MomentoFinanceiro = eTipoFaturamento.Pedido, 3, 2)
                Dim NumParcelas As Integer = Me.Where(Function(s) s.CodigoProvisao = codigoProvisao).ToList().Count

                'Valor positivo complemento 
                'Valor negativo estorno
                If CE = eTiposLancamentosPedidos.Complemento Then
                    'Se for complemento e nao existir parcelas, o sistema recupera uma das canceladas ou cria uma nova
                    If NumParcelas = 0 Then
                        'se o numero de parcelas for igual a 0 entao é pq nao existe nenhum titulo em previsao ativo com a situacao = 1 normal
                        'se o pedido for do financeiro novo havera um titulo em previsao com certeza porem cancelado, vamos reativalo
                        'se nao existir ai vamos criar um novo titulo
                        Dim where As String = "Empresa ='" & PD.CodigoEmpresa & "' and EndEmpresa = " & PD.EnderecoEmpresa & " and Pedido = " & PD.Codigo & " and provisao = " & codigoProvisao
                        Dim titulosEmPrevisaoCanceladosDoPedido As New Novo.ListTituloNovo(where)

                        If titulosEmPrevisaoCanceladosDoPedido.Count > 0 Then
                            Dim tit As Novo.TituloNovo = titulosEmPrevisaoCanceladosDoPedido.OrderByDescending(Function(s) s.Reprogramacao).ThenByDescending(Function(S) S.Codigo)(0)
                            tit.DesligarControles()
                            tit.IUD = "U"
                            tit.CodigoSituacao = 1
                            tit.Valores.EncargoValorDocumento.Valor = pValor
                            Me.Add(tit)
                            Return True
                        Else
                            'Cria Um Novo Titulo
                            Dim T As New Novo.TituloNovo
                            T.DesligarControles()
                            T.IUD = "I"
                            T.ReceberPagar = IIf(PD.SubOperacao.EntradaSaida = eEntradaSaida.Entrada, "P", "R")
                            T.Codigo = 0
                            T.CodigoUnidadeDeNegocio = PD.CodigoUnidadeNegocio
                            T.CodigoEmpresa = PD.CodigoEmpresa
                            T.EnderecoEmpresa = PD.EnderecoEmpresa
                            T.CodigoPedido = PD.Codigo
                            T.CodigoProvisao = IIf(PD.FinanceiroNovo, 2, 3)
                            T.CodigoMoeda = PD.CodigoMoeda
                            T.CodigoIndexador = PD.CodigoIndexador
                            T.CodigoSituacao = 1

                            T.Movimento = Date.Now
                            T.Vencimento = Funcoes.ValidaDataUtil(PD.CodigoEmpresa, PD.EnderecoEmpresa, CDate(Date.Now))
                            T.Reprogramacao = T.Vencimento
                            T.DataMoeda = Date.Now
                            T.DataBaixa = T.Vencimento

                            T.CodigoCliFor = PD.CodigoCliente
                            T.EnderecoCliFor = PD.EnderecoCliente

                            T.CodigoEmpresaRecPag = PD.CodigoEmpresa
                            T.EndEmpresaRecPag = PD.EnderecoEmpresa

                            T.CodigoContaContabilCliFor = PD.SubOperacao.CodigoGrupoContas
                            T.CodigoContaContabilRecPag = PD.Empresa.Empresa.CodigoContaGrupoBanco

                            T.IndiceTitulo = PD.IndiceFixado
                            T.Valores.EncargoValorDocumento.Valor = pValor

                            T.Historico = "REF. Pedido: " & PD.Codigo & " / " & PD.Cliente.Nome & "-" & IIf(PD.SubOperacao.EntradaSaida = eEntradaSaida.Saida, "Saida", "Entrada") & ",  Proviniente de complemento feito em " & Date.Now.ToString("dd-MM-yyyy")
                            T.LigarControles()
                            Me.Add(T)
                            Return True
                        End If
                    End If
                End If


                Dim parcelas As Decimal = Math.Round(pValor / NumParcelas, 2)
                Dim diferenca As Decimal = pValor - (parcelas * NumParcelas)
                Dim ParcelaSendoAlterada As Integer = 0

                If CE = eTiposLancamentosPedidos.Complemento Then
                    For Each row In Me.Where(Function(s) s.CodigoProvisao = codigoProvisao).OrderBy(Function(s) s.Reprogramacao).ThenBy(Function(s) s.Codigo)
                        ParcelaSendoAlterada += 1
                        row.IUD = "U"
                        If ParcelaSendoAlterada = NumParcelas Then
                            row.Valores.EncargoValorDocumento.Valor += parcelas + diferenca
                        Else
                            row.Valores.EncargoValorDocumento.Valor += parcelas
                        End If
                    Next
                End If

                ParcelaSendoAlterada = 0
                If CE = eTiposLancamentosPedidos.Estorno Then
                    Dim saldoADescontar As Decimal = pValor
                    While saldoADescontar > 0
                        For Each row In Me.Where(Function(s) s.CodigoProvisao = codigoProvisao).OrderBy(Function(s) s.Reprogramacao).ThenBy(Function(s) s.Codigo)
                            ParcelaSendoAlterada += 1
                            If row.Valores.EncargoValorDocumento.Valor < parcelas + diferenca Then
                                saldoADescontar -= row.Valores.EncargoValorDocumento.Valor

                                If NumParcelas <> ParcelaSendoAlterada Then
                                    parcelas = Math.Round(saldoADescontar / (NumParcelas - ParcelaSendoAlterada), 2)
                                    diferenca = saldoADescontar - (parcelas * (NumParcelas - ParcelaSendoAlterada))
                                End If

                                row.IUD = "U"
                                row.CodigoSituacao = 3
                                row.Valores.EncargoValorDocumento.Valor = 0
                            ElseIf row.Valores.EncargoValorDocumento.Valor = parcelas + diferenca Then
                                saldoADescontar -= row.Valores.EncargoValorDocumento.Valor
                                row.IUD = "U"
                                row.CodigoSituacao = 3
                                row.Valores.EncargoValorDocumento.Valor = 0
                            Else
                                row.IUD = "U"
                                row.Valores.EncargoValorDocumento.Valor -= parcelas + diferenca
                                saldoADescontar -= parcelas + diferenca
                                diferenca = 0
                            End If
                        Next
                    End While


                End If


                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function

        '*******  Compensacao Troca  **********
        Public Sub CompensacaoTroca(ByRef pValor As Decimal, ByRef NumeroTitulo As Integer, ByRef i As Integer)
            Dim DataFechada As Boolean = False
            Dim vlrCompTroca As Decimal = pValor
            For Each Item As NotaFiscalXItem In NF.Itens
                If pValor = 0 Then Exit For
                For Each NFDev As NotaFiscalDevolucaoXNotaFiscal In Item.NotasDevolucao
                    If pValor = 0 Then Exit For
                    'Verifica se a NF que está sendo devolvida está com a data fechada
                    If Not Funcoes.VerificaAcesso(NFDev.Nota.CodigoEmpresa, NFDev.Nota.EnderecoEmpresa, NFDev.Nota.Movimento, "NOTAS FISCAIS") AndAlso _
                           Not Funcoes.VerificaAcesso(NFDev.Nota.CodigoEmpresa, NFDev.Nota.EnderecoEmpresa, NFDev.Nota.Movimento, "CONTABIL") Then
                        DataFechada = True
                    End If
                    'Titulos
                    For Each tit As Novo.TituloNovo In NFDev.Nota.Titulos
                        'Data Fechada (AUDITADO)
                        If DataFechada Then 'And ou else

                            Dim titComp As New Novo.TituloNovo(NF.Pedido.Titulos.FirstOrDefault().Codigo)
                            'Pedido
                            titComp.IUD = "I"
                            i = i + 1
                            titComp.Codigo = NumeroTitulo + i
                            titComp.CodigoProvisao = eProvisao.Baixa
                            If NF.Pedido.SubOperacao.EntradaSaida = eEntradaSaida.Entrada Then
                                titComp.ReceberPagar = "R"
                                titComp.CodigoContaContabilCliFor = NF.Empresa.Empresa.CodigoContaCaixaCompensacao
                                titComp.CodigoContaContabilRecPag = NF.Empresa.Empresa.CodigoContaVariacaoMonetariaFornecedor
                            Else
                                titComp.ReceberPagar = "P"
                                titComp.CodigoContaContabilCliFor = NF.Empresa.Empresa.CodigoContaVariacaoMonetariaFornecedor
                                titComp.CodigoContaContabilRecPag = NF.Empresa.Empresa.CodigoContaCaixaCompensacao
                            End If

                            titComp.Valores.EncargoValorDocumento.Valor = pValor
                            titComp.CodigoTituloCompensacao = 0
                            NF.TitulosPedido.Add(titComp)
                            'Pedido Troca
                            Dim titTroca As New Novo.TituloNovo(NF.Pedido.PedidoTroca.Titulos.FirstOrDefault().Codigo)
                            titTroca.IUD = "I"
                            i = i + 1
                            titTroca.Codigo = NumeroTitulo + i
                            titTroca.CodigoProvisao = eProvisao.Baixa

                            If NF.Pedido.PedidoTroca.SubOperacao.EntradaSaida = eEntradaSaida.Entrada Then
                                titTroca.ReceberPagar = "R"
                                titTroca.CodigoContaContabilCliFor = NF.Empresa.Empresa.CodigoContaCaixaCompensacao
                                titTroca.CodigoContaContabilRecPag = NF.Empresa.Empresa.CodigoContaVariacaoMonetariaCliente
                            Else
                                titTroca.ReceberPagar = "P"
                                titTroca.CodigoContaContabilCliFor = NF.Empresa.Empresa.CodigoContaVariacaoMonetariaCliente
                                titTroca.CodigoContaContabilRecPag = NF.Empresa.Empresa.CodigoContaCaixaCompensacao
                            End If
                            titTroca.Valores.EncargoValorDocumento.Valor = pValor
                            titTroca.CodigoTituloCompensacao = 0
                            NF.TitulosPedido.Add(titTroca)

                            'Provisão no pedido de Troca
                            Dim titTrocaProvisao As New Novo.TituloNovo(NF.Pedido.PedidoTroca.Titulos.FirstOrDefault().Codigo)
                            titTrocaProvisao.IUD = "I"
                            i = i + 1
                            titTrocaProvisao.Codigo = NumeroTitulo + i
                            titTrocaProvisao.CodigoProvisao = eProvisao.Provisao
                            titTrocaProvisao.Valores.EncargoValorDocumento.Valor = pValor
                            titTrocaProvisao.CodigoTituloCompensacao = 0
                            NF.TitulosPedido.Add(titTrocaProvisao)

                            'Previsao  no Pedido
                            Dim origem As Integer = tit.CodigoTituloOrigem
                            If origem > 0 Then
                                Dim titprev = NF.TitulosPedido.Where(Function(s) s.Codigo = origem).FirstOrDefault
                                If titprev IsNot Nothing Then
                                    titprev.IUD = "U"
                                    titprev.Valores.EncargoValorDocumento.Valor += vlrCompTroca
                                Else
                                    Dim where As String = "Titulo_Id = " & origem
                                    Dim list = New Novo.ListTituloNovo(where)
                                    titprev = list(0)
                                    titprev.IUD = "U"
                                    titprev.CodigoSituacao = 1
                                    titprev.CodigoTituloCompensacao = 0
                                    titprev.Valores.EncargoValorDocumento.Valor = vlrCompTroca
                                    NF.TitulosPedido.Add(titprev)
                                End If
                            End If
                        Else
                            'Data Não Fechada
                            tit.IUD = "U"
                            tit.CodigoProvisao = eProvisao.Compensado

                            'Pedido de Troca
                            Dim CodigoTituloComp As Integer = tit.CodigoTituloCompensacao
                            Dim CodigoTitulo As Integer = tit.Codigo
                            For Each titTroca In NF.Pedido.PedidoTroca.Titulos.Where(Function(s) s.CodigoTituloCompensacao = CodigoTitulo OrElse s.Codigo = CodigoTituloComp)
                                If pValor >= titTroca.Valores.EncargoValorDocumento.Valor Then
                                    titTroca.IUD = "U"
                                    titTroca.CodigoProvisao = eProvisao.Provisao
                                    titTroca.CodigoTituloCompensacao = 0
                                    NF.TitulosPedido.Add(titTroca)
                                    pValor -= titTroca.Valores.EncargoValorDocumento.Valor
                                Else
                                    'Cria o Título.
                                    Dim TituloTroca As New Novo.TituloNovo(titTroca.Codigo)
                                    TituloTroca.IUD = "I"
                                    i += 1
                                    TituloTroca.Codigo = NumeroTitulo + i
                                    TituloTroca.Valores.EncargoValorDocumento.Valor = titTroca.Valores.EncargoValorDocumento.Valor - pValor
                                    TituloTroca.CodigoProvisao = eProvisao.Baixa
                                    TituloTroca.Sequencia = TituloTroca.Codigo
                                    TituloTroca.NotaTitulo.IUD = "I"
                                    NF.TitulosPedido.Add(TituloTroca)
                                    'Volta para provisão o valor restante
                                    titTroca.IUD = "U"
                                    titTroca.CodigoProvisao = eProvisao.Provisao
                                    titTroca.CodigoTituloCompensacao = 0
                                    titTroca.Valores.EncargoValorDocumento.Valor = pValor
                                    NF.TitulosPedido.Add(titTroca)

                                    If tit.Valores.EncargoValorDocumento.Valor > pValor Then
                                        Dim Titulo As New Novo.TituloNovo(tit.Codigo)
                                        Titulo.IUD = "I"
                                        i += 1
                                        Titulo.Codigo = NumeroTitulo + i
                                        Titulo.Valores.EncargoValorDocumento.Valor = tit.Valores.EncargoValorDocumento.Valor - pValor
                                        Titulo.CodigoProvisao = eProvisao.Baixa
                                        Titulo.Sequencia = Titulo.Codigo
                                        Titulo.NotaTitulo.IUD = "I"
                                        NF.TitulosPedido.Add(Titulo)
                                        tit.Valores.EncargoValorDocumento.Valor = vlrCompTroca
                                    End If
                                    pValor = 0
                                End If
                            Next

                            'Aumenta a Previsão do pedido
                            NF.TitulosPedido.Add(tit)
                            Dim origem As Integer = tit.CodigoTituloOrigem
                            If origem > 0 Then
                                Dim titprev = NF.TitulosPedido.Where(Function(s) s.Codigo = origem).FirstOrDefault
                                If titprev IsNot Nothing Then
                                    titprev.IUD = "U"
                                    titprev.Valores.EncargoValorDocumento.Valor += vlrCompTroca
                                Else
                                    Dim where As String = "Titulo_Id = " & origem
                                    Dim list = New Novo.ListTituloNovo(where)
                                    titprev = list(0)
                                    titprev.IUD = "U"
                                    titprev.CodigoSituacao = 1
                                    titprev.Valores.EncargoValorDocumento.Valor = vlrCompTroca
                                    NF.TitulosPedido.Add(titprev)
                                End If
                            End If

                            If pValor = 0 Then Exit For
                        End If
                    Next
                Next
            Next
        End Sub

        Private Sub AdiantamentoTroca(isBaixa As Boolean, ByVal pValor As Decimal, PagarReceber As String, ContaContabilProduto As String, ContaContabilAdiantamento As String, objResumo As ResumoFinanceiro, EfetuaTransferenciaFinanceira As Boolean)
            Dim T As New Novo.TituloNovo
            T.DesligarControles()
            T.IUD = "I"
            T.PedidoTroca = True
            T.ReceberPagar = PagarReceber
            'T.Codigo = 0
            T.CodigoUnidadeDeNegocio = NF.Pedido.CodigoUnidadeNegocio
            T.CodigoEmpresa = NF.CodigoEmpresa
            T.EnderecoEmpresa = NF.EnderecoEmpresa
            T.CodigoPedido = NF.CodigoPedido
            T.CodigoProvisao = eProvisao.Provisao
            T.CodigoMoeda = NF.Pedido.CodigoMoeda
            T.CodigoIndexador = NF.Pedido.CodigoIndexador
            T.CodigoSituacao = 1

            T.NotaTitulo = New Novo.NotaFiscalXTitulo(NF, T)

            T.Movimento = Date.Now
            T.Vencimento = Funcoes.ValidaDataUtil(NF.CodigoEmpresa, NF.EnderecoEmpresa, CDate(Date.Now))
            T.Reprogramacao = T.Vencimento
            T.DataMoeda = Date.Now
            T.DataBaixa = T.Vencimento

            T.CodigoCliFor = NF.CodigoCliente
            T.EnderecoCliFor = NF.EnderecoCliente

            If EfetuaTransferenciaFinanceira Then
                T.CodigoUnidadeDeNegocioRecPag = objResumo.ResumoTroca.CodigoUnidadeDeNegocio
                T.CodigoEmpresaRecPag = objResumo.ResumoTroca.CodigoEmpresa
                T.EndEmpresaRecPag = objResumo.ResumoTroca.CodigoEndEmpresa
            Else
                T.CodigoUnidadeDeNegocioRecPag = NF.CodigoUnidadeDeNegocio
                T.CodigoEmpresaRecPag = NF.CodigoEmpresa
                T.EndEmpresaRecPag = NF.EnderecoEmpresa
            End If

            T.CodigoContaContabilCliFor = ContaContabilProduto
            T.CodigoContaContabilRecPag = NF.Pedido.Empresa.Empresa.CodigoContaGrupoBanco

            T.IndiceTitulo = NF.Pedido.IndiceFixado

            If NF.Pedido.Moeda.Classificacao = eTiposMoeda.Oficial Then
                T.Valores.EncargoValorDocumento.ValorOficial = pValor
            Else
                T.Valores.EncargoValorDocumento.ValorMoeda = pValor
            End If

            T.Historico = IIf(isBaixa, "BAIXA DE ADIANTAMENTO", "ADIANTAMENTO") & " REF. NF: " & NF.Codigo & " / " & NF.Cliente.Nome & "-" & IIf(NF.SubOperacao.EntradaSaida = eEntradaSaida.Saida, "Saida", "Entrada") & ",  Proviniente de nota de troca feito em " & NF.Movimento.ToString("dd-MM-yyyy")

            If isBaixa Then
                T.AdiantamentosAbertos.DistribuirValorParaBaixarAdiantamentos(pValor, NF.Pedido.Moeda.Classificacao, False)
            End If

            T.LigarControles()

            Me.Add(T)
        End Sub

        Protected Sub AjustaTitulosEmPrevisaoPorAdiantamentoTroca(ByRef Sqls As ArrayList, ByRef NumeroTitulo As Integer, ByRef i As Integer)
            'Diminui a previsão do pedido
            Dim ValorASerRetiradoDaPrevisao As Decimal = NF.TotalNota
            Dim vlrCompensado As Decimal = NF.Pedido.PedidoTroca.Titulos.Where(Function(s) s.CodigoProvisao = eProvisao.Provisao).Sum(Function(s) s.Valores.EncargoValorDocumento.Valor)
            Dim CodigoProvisao As Integer = 0

            If Not vlrCompensado = 0 Then
                Dim tit As New TituloNovo(NF.Pedido.Titulos.Where(Function(s) s.CodigoProvisao = eProvisao.Previsao).FirstOrDefault.Codigo)

                If ValorASerRetiradoDaPrevisao > vlrCompensado Then
                    i += 1
                    tit.Codigo = NumeroTitulo + i
                    tit.IUD = "I"
                    tit.CodigoProvisao = eProvisao.Baixa
                    tit.CodigoContaContabilRecPag = NF.Pedido.Empresa.Empresa.CodigoContaCaixaCompensacao
                    tit.Valores.EncargoValorDocumento.Valor = vlrCompensado
                    tit.SalvarSql(Sqls, False, True)

                    Me(0).Valores.EncargoValorDocumento.Valor = ValorASerRetiradoDaPrevisao - vlrCompensado
                    CodigoProvisao = Me(0).Codigo
                    Me(0).CodigoContaContabilRecPag = NF.Pedido.Empresa.Empresa.CodigoContaCaixaCompensacao
                    Me(0).IUD = "U"
                    Me(0).NotaTitulo.IUD = "U"
                    Me(0).SalvarSql(Sqls)
                Else
                    Me(0).CodigoContaContabilRecPag = NF.Pedido.Empresa.Empresa.CodigoContaCaixaCompensacao
                    Me(0).IUD = "U"
                    Me(0).NotaTitulo.IUD = "U"
                    Me(0).CodigoProvisao = eProvisao.Baixa
                    Me(0).SalvarSql(Sqls)
                End If

                'Compensa a provisão do pedido de troca.
                For Each titTroca In NF.Pedido.PedidoTroca.Titulos.Where(Function(s) s.CodigoProvisao = eProvisao.Provisao)
                    titTroca.IUD = "U"
                    titTroca.CodigoProvisao = eProvisao.Baixa
                    titTroca.TituloOriginal = titTroca
                    If ValorASerRetiradoDaPrevisao > titTroca.Valores.EncargoValorDocumento.Valor Then
                        ValorASerRetiradoDaPrevisao -= titTroca.Valores.EncargoValorDocumento.Valor
                    Else
                        tit = New TituloNovo(NF.Pedido.PedidoTroca.Titulos.OrderBy(Function(s) s.CodigoProvisao).First.Codigo)
                        i += 1
                        tit.Codigo = NumeroTitulo + i
                        tit.IUD = "I"
                        tit.CodigoProvisao = eProvisao.Provisao
                        tit.CodigoContaContabilRecPag = NF.Pedido.Empresa.Empresa.CodigoContaCaixaCompensacao
                        tit.Valores.EncargoValorDocumento.Valor = titTroca.Valores.EncargoValorDocumento.Valor - ValorASerRetiradoDaPrevisao
                        tit.SalvarSql(Sqls, False, True)

                        titTroca.Valores.EncargoValorDocumento.Valor = ValorASerRetiradoDaPrevisao
                        ValorASerRetiradoDaPrevisao = 0
                    End If
                    titTroca.CodigoContaContabilRecPag = NF.Pedido.Empresa.Empresa.CodigoContaCaixaCompensacao 'Conta Caixa de Compensação
                    titTroca.Historico = "Título Compensado Pela NF: " & NF.Codigo & "-" & NF.Serie & " Pedido: " & NF.Pedido.Codigo & " Título: " & CodigoProvisao
                    titTroca.CodigoTituloCompensacao = CodigoProvisao
                    titTroca.SalvarSql(Sqls)
                    If ValorASerRetiradoDaPrevisao = 0 Then Exit For
                Next
            End If
        End Sub

        '******* Salvar Nota *************
        Private Sub SalvarSqlAjustaTitulosEmPrevisaoEFazCompensacao(ByRef Sqls As ArrayList, ByRef NumeroTitulo As Integer, ByRef i As Integer)
            'A Nota esta sendo Inserida e Nao é uma Devolucao
            If NF.IUD = "I" And Not NF.SubOperacao.Devolucao Then
                'Retira o valor a ser provisionado dos titulos em previsao do mais antigo para o com o vencimento mais recente
                Dim ValorASerRetiradoDaPrevisao As Decimal = NF.Titulos.Where(Function(s) s.CodigoProvisao = 3).Sum(Function(s) s.Valores.EncargoValorDocumento.Valor)

                For Each row In NF.TitulosPedido.Where(Function(s) s.CodigoProvisao = 2).OrderBy(Function(s) s.Reprogramacao).ThenBy(Function(s) s.Codigo)
                    If row.Pedido.Moeda.Classificacao = eTiposMoeda.Oficial Then
                        If row.Valores.EncargoValorDocumento.ValorOficial > ValorASerRetiradoDaPrevisao Then
                            row.IUD = "U"
                            row.Valores.EncargoValorDocumento.ValorOficial -= ValorASerRetiradoDaPrevisao
                            Exit For
                        Else
                            row.IUD = "C"
                            ValorASerRetiradoDaPrevisao -= row.Valores.EncargoValorDocumento.ValorOficial
                            row.Valores.EncargoValorDocumento.ValorOficial = 0
                        End If
                    Else
                        If row.Valores.EncargoValorDocumento.ValorMoeda > ValorASerRetiradoDaPrevisao Then
                            row.IUD = "U"
                            row.Valores.EncargoValorDocumento.ValorMoeda -= ValorASerRetiradoDaPrevisao
                            Exit For
                        Else
                            row.IUD = "C"
                            ValorASerRetiradoDaPrevisao -= row.Valores.EncargoValorDocumento.ValorMoeda
                            row.Valores.EncargoValorDocumento.ValorMoeda = 0
                        End If
                    End If

                    If NF.Pedido IsNot Nothing AndAlso NF.Pedido.MomentoFinanceiro = eTipoFaturamento.Lote AndAlso row.Valores.EncargoValorDocumento.Valor = 0 Then
                        Dim condicaoPagamento As New CondicaoPagamento(NF.Pedido.CodigoCondicaoPagamento)
                        'Pedido com Faturamento em Lote(quando acabar o previsão ajustar o vencimento de todas as provisões criadas apartir dela para a data da NF aplicando a condição de pagamento)
                        Dim codigoOrigem = row.Codigo
                        For Each tit In NF.TitulosPedido.Where(Function(s) s.CodigoTituloOrigem = codigoOrigem AndAlso s.CodigoSituacao = eSituacao.Normal AndAlso s.CodigoProvisao = eProvisao.Provisao)
                            tit.IUD = "U"
                            tit.Vencimento = Funcoes.ValidaDataUtil(NF.CodigoEmpresa, NF.EnderecoEmpresa, NF.Movimento.AddDays(condicaoPagamento.Periodo(0)))
                            tit.Reprogramacao = tit.Vencimento
                        Next
                    End If
                    If ValorASerRetiradoDaPrevisao = 0 Then Exit For
                Next
                NF.TitulosPedido.SalvarSql(Sqls)

                'A Nota esta sendo Inserida é uma Devolucao
            ElseIf NF.IUD = "I" And NF.SubOperacao.Devolucao Then
                '*******************************************************************
                '*** Compensa os titulos em Provisao *******************************
                '*******************************************************************
                Dim valor As Decimal = Me.Where(Function(s) s.CodigoProvisao = eProvisao.Compensado).Sum(Function(s) s.Valores.EncargoValorDocumento.Valor)
                Dim listOrigem = New Novo.ListTituloNovo()
                'NF´s que estão sendo devolvidas.
                If valor > 0 Then
                    Dim Liquidou As Boolean = False
                    Dim ValorAVoltarParaOTituloEmPrevisao As Decimal = 0
                    For Each Item As NotaFiscalXItem In NF.Itens
                        For Each itemNotaDevolucao As NotaFiscalDevolucaoXNotaFiscal In Item.NotasDevolucao.Where(Function(s) s.ValorDevolucao > 0)
                            For Each TitNotaDevolucao As TituloNovo In itemNotaDevolucao.Nota.Titulos.Where(Function(s) s.CodigoProvisao = eProvisao.Provisao).OrderBy(Function(s) s.Reprogramacao).ThenByDescending(Function(s) s.Codigo)
                                ValorAVoltarParaOTituloEmPrevisao = 0
                                If TitNotaDevolucao.Valores.EncargoValorDocumento.Valor > valor Then
                                    TitNotaDevolucao.IUD = "U"
                                    TitNotaDevolucao.Valores.EncargoValorDocumento.Valor -= valor
                                    If NF.Pedido.MomentoFinanceiro = eTipoFaturamento.Lote Then
                                        TitNotaDevolucao.Quantidade -= Math.Round(valor * NF.TotalQuantidadeFiscal / NF.TotalNota, 2)
                                    End If

                                    Dim NovoTitCompensar As New Novo.TituloNovo(TitNotaDevolucao.Codigo)
                                    NovoTitCompensar.IUD = "I"
                                    i += 1
                                    NovoTitCompensar.Codigo = NumeroTitulo + i
                                    NovoTitCompensar.TituloCompensacao = Me.Where(Function(s) s.CodigoProvisao = eProvisao.Compensado).First
                                    NovoTitCompensar.CodigoProvisao = eProvisao.Compensado
                                    NovoTitCompensar.Quantidade = 0
                                    NovoTitCompensar.Valores.EncargoValorDocumento.Valor = valor
                                    NovoTitCompensar.NotaTitulo = New Novo.NotaFiscalXTitulo(TitNotaDevolucao.NotaTitulo.NotaFiscal, NovoTitCompensar)
                                    NovoTitCompensar.NotaTitulo.IUD = "I"
                                    If valor > 0 Then NF.TitulosPedido.Add(NovoTitCompensar)
                                    Liquidou = True
                                    ValorAVoltarParaOTituloEmPrevisao = valor
                                    valor = 0
                                Else
                                    TitNotaDevolucao.IUD = "U"
                                    TitNotaDevolucao.CodigoProvisao = eProvisao.Compensado
                                    TitNotaDevolucao.Quantidade = 0
                                    TitNotaDevolucao.TituloCompensacao = Me.Where(Function(s) s.CodigoProvisao = eProvisao.Compensado).First
                                    valor -= TitNotaDevolucao.Valores.EncargoValorDocumento.Valor
                                    ValorAVoltarParaOTituloEmPrevisao = TitNotaDevolucao.Valores.EncargoValorDocumento.Valor
                                End If
                                Dim Codigo As Integer = TitNotaDevolucao.Codigo
                                NF.TitulosPedido.RemoveAll(Function(s) s.Codigo = Codigo)
                                NF.TitulosPedido.Add(TitNotaDevolucao)
                                Dim origem As Integer = TitNotaDevolucao.CodigoTituloOrigem

                                If origem > 0 Then
                                    Dim titprev = NF.TitulosPedido.Where(Function(s) s.Codigo = origem).FirstOrDefault
                                    If titprev IsNot Nothing Then
                                        titprev.IUD = "U"
                                        titprev.Valores.EncargoValorDocumento.Valor += ValorAVoltarParaOTituloEmPrevisao
                                    Else
                                        Dim where As String = "Titulo_Id = " & origem
                                        Dim list = New Novo.ListTituloNovo(where)
                                        titprev = list(0)
                                        titprev.IUD = "U"
                                        titprev.CodigoSituacao = 1
                                        titprev.Valores.EncargoValorDocumento.Valor = ValorAVoltarParaOTituloEmPrevisao
                                        NF.TitulosPedido.Add(titprev)
                                    End If
                                    listOrigem.Add(titprev)
                                End If
                                If Liquidou Then Exit For
                            Next
                        Next
                    Next
                End If
                'Titulos do Pedido.
                If valor > 0 Then
                    Dim Liquidou As Boolean = False
                    Dim TituloCompensacao As Integer = Me.Where(Function(s) s.CodigoProvisao = eProvisao.Compensado).FirstOrDefault.Codigo
                    Dim ValorAVoltarParaOTituloEmPrevisao As Decimal

                    For Each tit As Novo.TituloNovo In NF.TitulosPedido.Where(Function(s) s.CodigoProvisao = eProvisao.Provisao).OrderByDescending(Function(s) s.Reprogramacao).ThenByDescending(Function(s) s.Codigo)
                        ValorAVoltarParaOTituloEmPrevisao = 0
                        If valor = 0 Then Exit For
                        If tit.Valores.EncargoValorDocumento.Valor > valor Then
                            tit.IUD = "U"
                            tit.Valores.EncargoValorDocumento.Valor -= valor
                            If NF.Pedido.MomentoFinanceiro = eTipoFaturamento.Lote Then
                                tit.Quantidade -= Math.Round(valor * NF.TotalQuantidadeFiscal / NF.TotalNota, 2)
                            End If

                            Dim NovoTitCompensar As New Novo.TituloNovo(tit.Codigo)
                            NovoTitCompensar.IUD = "I"
                            i += 1
                            NovoTitCompensar.Codigo = NumeroTitulo + i
                            NovoTitCompensar.TituloCompensacao = Me.Where(Function(s) s.CodigoProvisao = eProvisao.Compensado).First
                            NovoTitCompensar.CodigoProvisao = eProvisao.Compensado
                            NovoTitCompensar.Quantidade = 0
                            NovoTitCompensar.Valores.EncargoValorDocumento.Valor = valor
                            NovoTitCompensar.NotaTitulo = New Novo.NotaFiscalXTitulo(tit.NotaTitulo.NotaFiscal, NovoTitCompensar)
                            NovoTitCompensar.NotaTitulo.IUD = "I"
                            If valor > 0 Then NF.TitulosPedido.Add(NovoTitCompensar)
                            Liquidou = True
                            ValorAVoltarParaOTituloEmPrevisao = valor
                        Else
                            tit.IUD = "U"
                            tit.CodigoProvisao = eProvisao.Compensado
                            tit.Quantidade = 0
                            tit.TituloCompensacao = Me.Where(Function(s) s.CodigoProvisao = eProvisao.Compensado).First
                            valor -= tit.Valores.EncargoValorDocumento.Valor
                            ValorAVoltarParaOTituloEmPrevisao = tit.Valores.EncargoValorDocumento.Valor
                        End If

                        Dim origem As Integer = tit.CodigoTituloOrigem
                        If origem > 0 Then
                            Dim titprev = NF.TitulosPedido.Where(Function(s) s.Codigo = origem).FirstOrDefault
                            If titprev IsNot Nothing Then
                                titprev.IUD = "U"
                                titprev.Valores.EncargoValorDocumento.Valor += ValorAVoltarParaOTituloEmPrevisao
                            Else
                                Dim where As String = "Titulo_Id = " & origem
                                Dim list = New Novo.ListTituloNovo(where)
                                titprev = list(0)
                                titprev.IUD = "U"
                                titprev.CodigoSituacao = 1
                                titprev.Valores.EncargoValorDocumento.Valor = ValorAVoltarParaOTituloEmPrevisao
                                NF.TitulosPedido.Add(titprev)
                            End If
                            listOrigem.Add(titprev)
                        End If
                        If Liquidou Then Exit For
                    Next
                End If
                'Altera a data das provisões conforme o pedido.
                If NF.Pedido.MomentoFinanceiro = eTipoFaturamento.Pedido Then
                    ReprogramarFaturamento(listOrigem, NumeroTitulo, i)
                End If
                'Realiza a compensação conforme pedido de troca.
                If NF.Pedido.Troca AndAlso valor > 0 Then CompensacaoTroca(valor, NumeroTitulo, i)
                'Salva os titulos alterados.
                NF.TitulosPedido.SalvarSql(Sqls, False, True)
            End If
        End Sub

        Private Sub ReprogramarFaturamento(ByVal pListaOrigem As ListTituloNovo, ByRef NumeroTitulo As Integer, ByRef i As Integer)
            If Not NF.TitulosPedido.Where(Function(s) s.CodigoProvisao = eProvisao.Provisao AndAlso s.Reprogramacao > pListaOrigem(0).Reprogramacao).Count > 0 Then Exit Sub
            Dim valor As Decimal = 0
            Dim where As String = "Pedido = " & NF.Pedido.Codigo & " and TituloOrigem = 0 and Provisao = 2 "
            Dim list = New Novo.ListTituloNovo(where)

            For Each Tit In pListaOrigem
                Dim codigo As Integer = Tit.Codigo
                list.RemoveAll(Function(s) s.Codigo = codigo)
                list.Add(Tit)
            Next

            For Each TitOrigem In list.OrderBy(Function(s) s.Codigo)
                Dim reprogramacao As Date = TitOrigem.Reprogramacao
                valor = TitOrigem.Valores.EncargoValorDocumento.Valor
                If valor > 0 Then
                    For Each Tit In NF.TitulosPedido.Where(Function(s) s.CodigoProvisao = eProvisao.Provisao AndAlso s.Reprogramacao > reprogramacao).OrderBy(Function(s) s.Reprogramacao).ThenBy(Function(s) s.Codigo)
                        If valor = 0 Then Exit For
                        Dim codigoTituloOrigem As Integer = Tit.CodigoTituloOrigem
                        Dim index As Integer = list.FindIndex(Function(s) s.Codigo = codigoTituloOrigem)
                        If Tit.Valores.EncargoValorDocumento.Valor <= valor Then
                            valor -= Tit.Valores.EncargoValorDocumento.Valor
                            list(index).Valores.EncargoValorDocumento.Valor += Tit.Valores.EncargoValorDocumento.Valor
                        Else
                            Dim NovoTit As New Novo.TituloNovo(Tit.Codigo)
                            NovoTit.IUD = "I"
                            i += 1
                            NovoTit.Codigo = NumeroTitulo + i
                            NovoTit.Valores.EncargoValorDocumento.Valor = Tit.Valores.EncargoValorDocumento.Valor - valor
                            NF.TitulosPedido.Add(NovoTit)
                            'Altera o valor do Título                            
                            Tit.Valores.EncargoValorDocumento.Valor = valor
                            list(index).Valores.EncargoValorDocumento.Valor += valor
                            valor = 0
                        End If

                        'Altera os dados de origem.
                        If Not Tit.IUD = "I" Then Tit.IUD = "U"
                        Tit.CodigoTituloOrigem = TitOrigem.Codigo
                        Tit.Reprogramacao = TitOrigem.Reprogramacao
                        If Not Tit.IUD = "I" Then
                            Tit.IUD = "U"
                            NF.TitulosPedido.Add(Tit)
                        End If
                    Next

                    If valor = 0 Then
                        TitOrigem.IUD = "U"
                        TitOrigem.CodigoSituacao = eSituacao.Excluido
                        TitOrigem.Valores.EncargoValorDocumento.Valor = 0
                    Else
                        TitOrigem.IUD = "U"
                        TitOrigem.CodigoSituacao = eSituacao.Normal
                        TitOrigem.Valores.EncargoValorDocumento.Valor = valor
                    End If

                    NF.TitulosPedido.Add(TitOrigem)
                End If
            Next
        End Sub

        Private Sub CriaTitulosParaBaixaDeAdiantamento(ByRef Sqls As ArrayList, ByRef NumeroTitulo As Integer, ByRef i As Integer)
            For Each adi In Me.AdiantamentosAbertos.Where(Function(s) s.VlrBaixa > 0)
                Dim tituloBxAd As New Novo.TituloNovo(adi.CodigoTitulo)
                tituloBxAd.IUD = "I"

                i += 1
                tituloBxAd.Codigo = NumeroTitulo + i

                tituloBxAd.Adiantamento = Nothing
                If FRETE Is Nothing Then
                    tituloBxAd.DataBaixa = NF.Movimento
                    tituloBxAd.Vencimento = NF.Movimento
                    tituloBxAd.Reprogramacao = NF.Movimento
                    tituloBxAd.DataMoeda = NF.Movimento
                    tituloBxAd.CodigoContaContabilRecPag = NF.Pedido.SubOperacao.CodigoGrupoContas
                Else
                    tituloBxAd.DataBaixa = FRETE.Movimento
                    tituloBxAd.Vencimento = FRETE.Movimento
                    tituloBxAd.Reprogramacao = FRETE.Movimento
                    tituloBxAd.DataMoeda = FRETE.Movimento
                    tituloBxAd.CodigoContaContabilRecPag = FRETE.ListTituloFatura(0).TituloNovo.Valores.EncargoValorDocumento.CodigoContaEncargo
                End If

                tituloBxAd.CodigoTipoPgto = 1
                tituloBxAd.CodigoCarteiraDoTitulo = 1


                If NF IsNot Nothing Then
                    tituloBxAd.NotaTitulo = New Novo.NotaFiscalXTitulo(NF, tituloBxAd)
                    If NF.NFG Then
                        tituloBxAd.CodigoPedido = NF.CodigoPedido
                    End If
                End If

                tituloBxAd.ReceberPagar = IIf(tituloBxAd.ReceberPagar = "R", "P", "R")
                tituloBxAd.AdiantamentosAbertos.Clear()
                tituloBxAd.AdiantamentosAbertos.Add(adi)

                For Each vlr In tituloBxAd.Valores
                    If Not vlr.Equals(tituloBxAd.Valores.EncargoValorDocumento) And Not vlr.Equals(tituloBxAd.Valores.EncargoValorLiquido) Then
                        tituloBxAd.Valores.Remove(vlr)
                    End If
                Next

                tituloBxAd.Valores.EncargoValorDocumento.DC = IIf(tituloBxAd.Valores.EncargoValorDocumento.DC = "C", "D", "C")
                tituloBxAd.Valores.EncargoValorLiquido.DC = IIf(tituloBxAd.Valores.EncargoValorDocumento.DC = "C", "D", "C")

                If NF IsNot Nothing Then
                    If NF.Pedido.Moeda.Classificacao = eTiposMoeda.Oficial Then
                        tituloBxAd.Valores.EncargoValorDocumento.ValorOficial = adi.VlrBaixa
                        tituloBxAd.Valores.EncargoValorDocumento.ValorMoeda = Funcoes.ConverteParaMoedaExtrangeira(adi.VlrBaixa, 2, NF.Movimento) 'Math.Round(adi.VlrBaixa / NF.IndiceNota, 2)

                        tituloBxAd.Valores.EncargoValorLiquido.ValorOficial = adi.VlrBaixa
                        tituloBxAd.Valores.EncargoValorLiquido.ValorMoeda = Funcoes.ConverteParaMoedaExtrangeira(adi.VlrBaixa, 2, NF.Movimento) 'Math.Round(adi.VlrBaixa / NF.IndiceNota, 2)
                    Else
                        tituloBxAd.Valores.EncargoValorDocumento.ValorOficial = Funcoes.ConverteParaMoedaExtrangeira(adi.VlrBaixa, 1, NF.Movimento) 'Math.Round(adi.VlrBaixa * NF.IndiceNota, 2)
                        tituloBxAd.Valores.EncargoValorDocumento.ValorMoeda = adi.VlrBaixa

                        tituloBxAd.Valores.EncargoValorLiquido.ValorOficial = Funcoes.ConverteParaMoedaExtrangeira(adi.VlrBaixa, 1, NF.Movimento) 'Math.Round(adi.VlrBaixa * NF.IndiceNota, 2)
                        tituloBxAd.Valores.EncargoValorLiquido.ValorMoeda = adi.VlrBaixa
                    End If
                Else
                    tituloBxAd.Valores.EncargoValorDocumento.ValorOficial = adi.VlrBaixa
                    tituloBxAd.Valores.EncargoValorDocumento.ValorMoeda = Funcoes.ConverteParaMoedaExtrangeira(adi.VlrBaixa, 3, FRETE.Movimento)
                End If

                tituloBxAd.Historico = "Baixa do Adiantamento " & adi.CodigoTitulo & " do Pedido " & tituloBxAd.CodigoPedido
                tituloBxAd.Observacoes = ""

                Me.Add(tituloBxAd)
                tituloBxAd.SalvarSql(Sqls, True, True)
            Next
        End Sub

#End Region

    End Class

    '***********************************************************************************************************************************************************
    '**************************************************************** CLASSE BASE TITULO ***********************************************************************
    '***********************************************************************************************************************************************************
    <Serializable()> _
    Public Class TituloNovo
        Implements IBaseEntity

#Region "Uteis"
        Private _Controlando As Boolean = True
        Private _MsgControle As String = ""
#End Region

#Region "Contrutor"
        Public Sub New()

        End Sub

        Public Sub New(ByVal pTitulo As Integer, Optional pCodigoEmpresa As String = "", Optional pCodigoBordero As Integer = 0)
            If pTitulo + pCodigoBordero = 0 Then Exit Sub

            Dim sql As String
            sql = "SELECT T.Titulo_Id, T.UnidadeDeNegocio, T.Empresa, T.EndEmpresa, T.RecPag, isnull(T.RegistroMestre,0) as RegistroMestre, " & vbCrLf & _
                  "       isnull(T.Pedido,0) as Pedido," & vbCrLf & _
                  "       T.Provisao, T.Moeda, T.Indexador, isnull(T.indice,0) as indice, T.TipoPagto, T.Situacao, T.Movimento, T.Vencimento, T.Reprogramacao, T.DataMoeda, T.DataBaixa, " & vbCrLf & _
                  "       T.CliFor, T.EnderecoCliFor, T.BancoCliFor, T.AgenciaCliFor, T.DigitoAgenciaCliFor, T.ContaCliFor, T.DigitoContaCliFor, T.ContaContabilCliFor," & vbCrLf & _
                  "       T.EmpresaRecPag, T.EndEmpresaRecPag, T.ContaContabilRecPag, " & vbCrLf & _
                  "       T.Cheque, isnull(T.NumeroDoCheque,0) as NumeroDoCheque, T.Slips, T.Recibo, T.CodigoDeBarra, isnull(T.CodigoDeBarraDigitado,0) as CodigoDeBarraDigitado, isnull(T.CodigoDeBarraPreImpresso,0) as CodigoDeBarraPreImpresso," & vbCrLf & _
                  "       T.Historico, T.Observacoes, isnull(T.CarteiraDoTitulo,0) as CarteiraDoTitulo, isnull(T.SituacaoBancaria,0) as SituacaoBancaria," & vbCrLf & _
                  "       isnull(T.Produto,'') as Produto, isnull(T.Quantidade,0) as Quantidade, isnull(T.ClienteRecPag,'') as ClienteRecPag, isnull(T.EndClienteRecPag,0) as EndClienteRecPag," & vbCrLf & _
                  "       isnull(T.PedidoRecPag,0) as PedidoRecPag, T.UnidadeDeNegocioRecPag," & vbCrLf & _
                  "       isnull(P.Troca,0) as PedidoTroca, " & vbCrLf & _
                  "       isnull(TituloOrigem,0) as TituloOrigem, isnull(TituloCompensacao,0) as TituloCompensacao, Fixacao, valorPrevisao " & vbCrLf & _
                  "  FROM Titulos T" & vbCrLf & _
                  "  Left Join Pedidos P" & vbCrLf & _
                  "    on P.Empresa_id    = T.Empresa" & vbCrLf & _
                  "   and P.EndEmpresa_id = T.EndEmpresa" & vbCrLf & _
                  "   and P.Pedido_id     = T.Pedido" & vbCrLf & _
                  "  Left Join SubOperacoes SO" & vbCrLf & _
                  "    on SO.Operacao_id     = P.Operacao" & vbCrLf & _
                  "   and SO.Suboperacoes_id = P.Suboperacao" & vbCrLf

            If pTitulo > 0 Then
                sql &= " Where T.Titulo_id = " & pTitulo '& " and T.Situacao = 1 "
            ElseIf pCodigoBordero > 0 Then
                sql &= " Where T.Titulo_id = (Select TituloBordero From Bordero where Empresa_Id ='" & pCodigoEmpresa & "' and Bordero_Id = " & pCodigoBordero & ")"
            Else
                'sql &= "Where x = y"
            End If

            Dim Banco As New AcessaBanco
            Dim ds As DataSet = Banco.ConsultaDataSet(sql, "Titulo")

            If ds.Tables.Count = 0 OrElse ds.Tables(0).Rows.Count = 0 Then Exit Sub
            Dim row As DataRow = ds.Tables(0).Rows(0)

            DesligarControles()

            '_IUD = "U"
            _Codigo = row("Titulo_Id")
            _CodigoUnidadeDeNegocio = row("UnidadeDeNegocio")

            _CodigoEmpresa = row("Empresa")
            _EndEmpresa = row("EndEmpresa")

            _ReceberPagar = row("RecPag")

            _RegistroMestre = row("RegistroMestre")

            _CodigoPedido = row("Pedido")
            _PedidoTroca = row("PedidoTroca")

            _CodigoProvisao = row("Provisao")
            _CodigoMoeda = row("Moeda")
            _CodigoIndexador = row("Indexador")
            _IndiceTitulo = row("Indice")

            _CodigoTipoPagto = row("TipoPagto")
            _CodigoSituacao = row("Situacao")

            _Movimento = row("Movimento")
            _Vencimento = row("Vencimento")
            _Reprogramacao = row("Reprogramacao")
            _DataMoeda = row("DataMoeda")
            If Not IsDBNull(row("DataBaixa")) Then _DataBaixa = row("DataBaixa")

            _CodigoCliFor = row("CliFor")
            _EnderecoCliFor = row("EnderecoCliFor")
            _CodigoBancoCliFor = row("BancoCliFor")
            _CodigoAgenciaCliFor = row("AgenciaCliFor")
            _DigitoAgenciaCliFor = row("DigitoAgenciaCliFor")
            _ContaCliFor = row("ContaCliFor")
            _DigitoContaCliFor = row("DigitoContaCliFor")
            _CodigoContaContabilCliFor = row("ContaContabilCliFor")

            _CodigoUnidadeDeNegocioRecPag = row("UnidadeDeNegocioRecPag").ToString()
            _CodigoEmpresaRecPag = row("EmpresaRecPag")
            _EndEmpresaRecPag = row("EndEmpresaRecPag")
            _CodigoContaContabilRecPag = row("ContaContabilRecPag")
            _Cheque = row("Cheque")
            _NumeroDoCheque = row("NumeroDoCheque")
            _Slips = row("Slips")
            _Recibo = row("Recibo")
            _CodigoFixacao = row("Fixacao")

            _CodigoDeBarras = row("CodigoDeBarra")
            _CodigoDeBarrasDigitado = row("CodigoDeBarraDigitado")
            _CodigoDeBarrasPreImpresso = row("CodigoDeBarraPreImpresso")
            _Historico = row("Historico")
            _Observacoes = row("Observacoes")
            _CodigoCarteiraDoTitulo = row("CarteiraDoTitulo")
            _SituacaoBancaria = row("SituacaoBancaria")

            _CodigoProduto = row("Produto")
            _Quantidade = row("Quantidade")
            _CodigoClienteRecPag = row("ClienteRecPag")
            _EndClienteRecPag = row("EndClienteRecPag")
            _CodigoPedidoRecPag = row("PedidoRecPag")

            _CodigoTituloOrigem = row("TituloOrigem")
            _CodigoTituloCompensacao = row("TituloCompensacao")
            '_ValorPrevisao = row("ValorPrevisao")

            '_Bordero = row("Bordero")
            '_BorderoMestre = row("BorderoMestre")
            '_CodigoTituloRecompra = row("TituloRecompra")

            If Not _CodigoIndexador = 99 Then VerificarIndice()
            LigarControles()
        End Sub
#End Region

#Region "Fields"
        '** Controle ***************************************************************
        Private _IUD As String = ""
        Private _TituloOriginal As TituloNovo
        Private _CriarAdiantamento As Boolean = True

        '*** Titulo ****************************************************************
        Private _ReceberPagar As String = ""
        Private _Codigo As Integer 'Titulo_Id

        '** Unidade de Negocio *****************************************************
        Private _CodigoUnidadeDeNegocio As String = ""
        Private _UnidadeDeNegocio As Cliente

        '** Empresa ****************************************************************
        Private _CodigoEmpresa As String = ""
        Private _EndEmpresa As Integer 'EndEmpresa
        Private _Empresa As Cliente 'Empresa

        '** Agrupamento ************************************************************
        Private _RegistroMestre As Integer
        Private _TituloMestre As TituloNovo
        Private _TitulosAgrupados As ListTituloNovo

        '** Pedido *****************************************************************
        Private _CodigoPedido As Integer
        Private _Pedido As Pedido
        Private _PedidoTroca As Boolean

        '** Fixacao *****************************************************************
        Private _CodigoFixacao As Integer

        '** Tipo Provisao **********************************************************
        Private _CodigoProvisao As Integer 'Provisao
        Private _Provisao As TipoProvisao
        Private _DescricaoProvisao As String = ""

        '** Moeda ******************************************************************
        Private _CodigoMoeda As Integer 'Moeda
        Private _Moeda As Moeda
        Private _DescricaoMoeda As String = ""

        '** Indexador **************************************************************
        Private _CodigoIndexador As Integer
        Private _Indexador As Indexador

        '** Tipo Pagamento *********************************************************
        Private _CodigoTipoPagto As Integer
        Private _TipoPagto As TipoDePagamento

        '** Situacao ***************************************************************
        Private _CodigoSituacao As Integer 'Situacao
        Private _Situacao As Situacao

        '** Datas ******************************************************************
        Private _Movimento As Date 'Movimento
        Private _Vencimento As Date 'Vencimento
        Private _Reprogramacao As Date 'Prorrogacao
        Private _DataMoeda As Date 'DataMoeda
        Private _DataBaixa As Date 'Baixa

        '** Cliente / Fornecedor ***************************************************
        Private _CodigoCliFor As String = "" 'Cliente
        Private _EnderecoCliFor As Integer 'EndCliente
        Private _CliFor As Cliente 'Cliente

        '** Banco Cliente / Fornecedor *********************************************
        Private _CodigoBancoCliFor As Integer
        Private _BancoCliFor As Banco
        Private _CodigoAgenciaCliFor As String = ""
        Private _DigitoAgenciaCliFor As String = ""
        Private _AgenciaCliFor As Agencia
        Private _ContaCliFor As String = ""
        Private _DigitoContaCliFor As String = ""
        Private _CodigoContaContabilCliFor As String = ""
        Private _ContaContabilCliFor As PlanoDeConta

        '** Deposito Cliente / Fornecedor ******************************************
        Private _CodigoDepositoCliFor As String = String.Empty
        Private _EndDepositoCliFor As Integer
        Private _DepositoCliFor As Cliente

        '** Unidade de Negocio *****************************************************
        Private _CodigoUnidadeDeNegocioRecPag As String = ""
        Private _UnidadeDeNegocioRecPag As Cliente

        '** Empresa Recebedora / Pagadora  *****************************************
        Private _CodigoEmpresaRecPag As String = ""
        Private _EndEmpresaRecPag As Integer
        Private _EmpresaRecPag As Cliente

        '** Conta Contabil Empresa Recebedora / Pagadora ***************************
        Private _CodigoContaContabilRecPag As String = ""
        Private _ContaContabilRecPag As PlanoDeConta

        '** Depósito Recebedora / Pagadora  *****************************************
        Private _CodigoDepositoRecPag As String = ""
        Private _EndDepositoRecPag As Integer
        Private _DepositoRecPag As Cliente

        '** Cheque - Slips - Recibo - Boleto ***************************************
        Private _Cheque As Boolean 'Cheque
        Private _NumeroDoCheque As Integer
        Private _Slips As Boolean 'Slips
        Private _Recibo As Boolean 'Recibo

        Private _CodigoDeBarras As String
        Private _CodigoDeBarrasDigitado As Boolean
        Private _CodigoDeBarrasPreImpresso As Boolean

        '** Historico - Observacao *************************************************
        Private _Historico As String = ""
        Private _Observacoes As String = ""

        '** Carteira ***************************************************************
        Private _CodigoCarteiraDoTitulo As Integer
        Private _CarteiraDoTitulo As CarteiraDoTitulo 'Carteira do titulo

        '** Situacao Bancaria ******************************************************
        Private _SituacaoBancaria As Integer

        '** Desdobrar Fornecedor ***************************************************
        Private _DesdobraFornecedor As TituloXDesdobrarFornecedor

        '** Indice *****************************************************************
        Private _IndiceTitulo As Decimal

        '** Contabilizacoes ********************************************************
        Private _Contabilizacoes As ListRazao
        Private _Lote As Integer
        Private _Sequencia As Integer

        '** Adiantamentos  *********************************************************
        Private _Adiantamento As Novo.AdiantamentoNovo
        Private _Baixas_AdiantamentoEfetuadas As Novo.ListAdiantamentoBaixaNovo
        Private _AdiantamentosAbertos As Novo.ListAdiantamentoNovo

        '** Lancamentos Contabeis **************************************************
        Private _CodigoProduto As String = ""
        Private _Produto As Produto
        Private _Quantidade As Decimal
        Private _CodigoClienteRecPag As String = ""
        Private _EndClienteRecPag As Integer
        Private _ClienteRecPag As Cliente
        Private _CodigoPedidoRecPag As Integer
        Private _PedidoRecPag As Pedido

        '** Bordero ****************************************************************
        Private _Bordero As Novo.Bordero 'Só vai ser instanciado com informacoes qdo for Duplicatas Descontadas e o titulo for o mestre
        Private _TituloBorderos As Novo.ListBorderoXTitulo 'tras todos os borderos no qual aquele titulo participou
        Private _UltimoBordero As Novo.BorderoXTitulo 'tras o bordero atual que aquele titulo esta vinculado

        '** Valores **************************************************************** 
        Private _Valores As Novo.ListTituloXContaContabil

        '** Destinacao - Procuracao ************************************************ 
        Private _TituloDestinacao As Novo.TituloXDestinacao

        '**** Nota Fiscal **********************************************************
        Private _NotaTitulo As Novo.NotaFiscalXTitulo

        '**** Titulo Origem ********************************************************
        Private _CodigoTituloOrigem As Integer
        Private _TituloOrigem As TituloNovo
        Private _CodigoTituloCompensacao As Integer
        Private _TituloCompensacao As TituloNovo
        Private _TituloHistorico As ListTituloXHistorico
        Private _ValorPrevisao As Decimal
#End Region

#Region "Property"
        Public ReadOnly Property FinanceiroNovo As Boolean
            Get
                Return CBool(ConfigurationManager.AppSettings("financeiroNovo"))
            End Get
        End Property

        '** Controle ***************************************************************
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
        Public Property IUD() As String
            Get
                Return _IUD
            End Get
            Set(ByVal value As String)
                _IUD = value
            End Set
        End Property
        Public Property TituloOriginal() As Novo.TituloNovo
            Get
                Return _TituloOriginal
            End Get
            Set(ByVal value As Novo.TituloNovo)
                _TituloOriginal = value
            End Set
        End Property

        Public ReadOnly Property isAdiantamento As Boolean
            Get
                If ContaContabilCliFor.Adiantamento Then
                    If Me.ReceberPagar = "R" And CodigoContaContabilCliFor.Substring(0, 1) = "2" Then Return True
                    If ReceberPagar = "P" And CodigoContaContabilCliFor.Substring(0, 1) = "1" Then Return True
                End If
                Return False
            End Get
        End Property
        Public ReadOnly Property isBaixaAdiantamento As Boolean
            Get
                If ContaContabilCliFor.Adiantamento OrElse (TituloOriginal IsNot Nothing AndAlso TituloOriginal.CodigoProvisao = 3) Then
                    If ReceberPagar = "R" And CodigoContaContabilCliFor.Substring(0, 1) = "1" Then Return True
                    If ReceberPagar = "P" And CodigoContaContabilCliFor.Substring(0, 1) = "2" Then Return True
                End If
                If ContaContabilRecPag.Adiantamento Then Return True
                Return False
            End Get
        End Property

        '*** Titulo ****************************************************************
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

        '** Unidade de Negocio *****************************************************
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

        '** Empresa ****************************************************************
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

        '** Agrupamento ************************************************************
        Public Property RegistroMestre() As Integer
            Get
                Return _RegistroMestre
            End Get
            Set(ByVal value As Integer)
                _RegistroMestre = value
                _TituloMestre = Nothing
            End Set
        End Property
        Public Property TituloMestre() As Novo.TituloNovo
            Get
                If _RegistroMestre = _Codigo Then Return Me

                If _TituloMestre Is Nothing And _RegistroMestre > 0 Then _TituloMestre = New Novo.TituloNovo(_RegistroMestre)
                Return _TituloMestre
            End Get
            Set(ByVal value As Novo.TituloNovo)
                _TituloMestre = value
            End Set
        End Property
        Public Property TitulosAgrupados() As Novo.ListTituloNovo
            Get
                If _TitulosAgrupados Is Nothing Then
                    If _Codigo = _RegistroMestre And _Codigo > 0 Then
                        _TitulosAgrupados = New Novo.ListTituloNovo("RegistroMestre = " & _Codigo & " and Titulo_id <> " & _Codigo)
                    Else
                        _TitulosAgrupados = New Novo.ListTituloNovo
                    End If
                End If

                Return _TitulosAgrupados
            End Get
            Set(ByVal value As Novo.ListTituloNovo)
                _TitulosAgrupados = value
            End Set
        End Property

        '** Pedido *****************************************************************
        Public Property CodigoPedido() As Integer
            Get
                If _CodigoPedido = 0 AndAlso Not Me.Pedido Is Nothing Then
                    _CodigoPedido = Me.Pedido.Codigo
                End If
                Return _CodigoPedido
            End Get
            Set(ByVal value As Integer)
                _CodigoPedido = value
                _Pedido = Nothing
                If Me.Controlando Then VerificarIndice()
            End Set
        End Property
        Public Property Pedido() As Pedido
            Get
                If _Pedido Is Nothing AndAlso _CodigoPedido > 0 AndAlso Me.CodigoEmpresa.Length > 0 Then
                    _Pedido = New Pedido(_CodigoEmpresa, _EndEmpresa, _CodigoPedido)
                End If
                Return _Pedido
            End Get
            Set(ByVal value As Pedido)
                _Pedido = value
                _CodigoPedido = value.Codigo
            End Set
        End Property
        Public Property PedidoTroca As Boolean
            Get
                Return _PedidoTroca
            End Get
            Set(value As Boolean)
                _PedidoTroca = value
            End Set
        End Property

        '** Fixacao ****************************************************************
        Public Property CodigoFixacao As Integer
            Get
                Return _CodigoFixacao
            End Get
            Set(ByVal value As Integer)
                _CodigoFixacao = value
            End Set
        End Property
        Public ReadOnly Property Fixacao As Fixacao
            Get
                If _CodigoFixacao > 0 Then
                    Return Me.Pedido.Itens(0).Fixacoes.Where(Function(s) s.Codigo = _CodigoFixacao).FirstOrDefault()
                End If
                Return Nothing
            End Get
        End Property

        '** Tipo Provisao **********************************************************
        Public Property CodigoProvisao() As Integer
            Get
                Return _CodigoProvisao
            End Get
            Set(ByVal value As Integer)
                _CodigoProvisao = value
                _Provisao = Nothing
                If Me.Controlando Then VerificarIndice()
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

        '** Moeda ******************************************************************
        Public Property CodigoMoeda() As Integer
            Get
                Return _CodigoMoeda
            End Get
            Set(ByVal value As Integer)
                _CodigoMoeda = value
                _Moeda = Nothing
                _DescricaoMoeda = ""
                If Me.Controlando Then VerificarIndice()
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
                If _CodigoMoeda > 0 And _DescricaoMoeda = "" Then _DescricaoMoeda = Moeda.Descricao
                Return _DescricaoMoeda
            End Get
        End Property

        '** Indexador **************************************************************
        Public Property CodigoIndexador() As Integer
            Get
                Return _CodigoIndexador
            End Get
            Set(ByVal value As Integer)
                _CodigoIndexador = value
                If Me.Controlando Then VerificarIndice()
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

        '** Tipo Pagamento *********************************************************
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

        '** Situacao ***************************************************************
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

        '** Datas ******************************************************************
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
                If Me.Controlando AndAlso Valores IsNot Nothing AndAlso Valores.Count > 0 Then
                    Valores.AtualizaValores()
                End If
            End Set
        End Property
        Public Property Reprogramacao() As Date
            Get
                Return _Reprogramacao
            End Get
            Set(ByVal value As Date)
                _Reprogramacao = value
                If Me.Controlando Then VerificarIndice()
            End Set
        End Property
        Public Property DataMoeda() As Date
            Get
                Return _DataMoeda
            End Get
            Set(ByVal value As Date)
                _DataMoeda = value
            End Set
        End Property
        Public Property DataBaixa() As Date
            Get
                Return _DataBaixa
            End Get
            Set(ByVal value As Date)
                _DataBaixa = value
                If Me.Controlando AndAlso Not Me.CodigoIndexador = 99 Then VerificarIndice()
            End Set
        End Property

        '** Cliente / Fornecedor ***************************************************
        Public Property CodigoCliFor() As String
            Get
                Return _CodigoCliFor
            End Get
            Set(ByVal value As String)
                _CodigoCliFor = value
                _CliFor = Nothing
            End Set
        End Property
        Public Property EnderecoCliFor() As Integer
            Get
                Return _EnderecoCliFor
            End Get
            Set(ByVal value As Integer)
                _EnderecoCliFor = value
            End Set
        End Property
        Public Property CliFor() As Cliente
            Get
                If _CliFor Is Nothing And _CodigoCliFor.Length > 0 Then _CliFor = New Cliente(_CodigoCliFor, _EnderecoCliFor)
                Return _CliFor
            End Get
            Set(ByVal value As Cliente)
                _CliFor = value
                _CodigoCliFor = value.Codigo
                _EnderecoCliFor = value.CodigoEndereco
            End Set
        End Property

        '** Banco Cliente / Fornecedor *********************************************
        Public Property CodigoBancoCliFor() As Integer
            Get
                Return _CodigoBancoCliFor
            End Get
            Set(ByVal value As Integer)
                _CodigoBancoCliFor = value
                _BancoCliFor = Nothing
                _AgenciaCliFor = Nothing
            End Set
        End Property
        Public ReadOnly Property BancoCliFor() As Banco
            Get
                If _BancoCliFor Is Nothing And _CodigoBancoCliFor > 0 Then _BancoCliFor = New Banco(_CodigoBancoCliFor)
                Return _BancoCliFor
            End Get
        End Property
        Public Property CodigoAgenciaCliFor() As String
            Get
                Return _CodigoAgenciaCliFor
            End Get
            Set(ByVal value As String)
                _CodigoAgenciaCliFor = value
                _AgenciaCliFor = Nothing
            End Set
        End Property
        Public Property DigitoAgenciaCliFor() As String
            Get
                Return _DigitoAgenciaCliFor
            End Get
            Set(ByVal value As String)
                _DigitoAgenciaCliFor = value
                _AgenciaCliFor = Nothing
            End Set
        End Property
        Public ReadOnly Property AgenciaCliFor() As Agencia
            Get
                If _CodigoBancoCliFor > 0 And _CodigoAgenciaCliFor.Length > 0 And _AgenciaCliFor Is Nothing Then _AgenciaCliFor = New Agencia(_CodigoBancoCliFor, _CodigoAgenciaCliFor, _DigitoAgenciaCliFor)
                Return _AgenciaCliFor
            End Get
        End Property
        Public Property ContaCliFor() As String
            Get
                Return _ContaCliFor
            End Get
            Set(ByVal value As String)
                _ContaCliFor = value
            End Set
        End Property
        Public Property DigitoContaCliFor() As String
            Get
                Return _DigitoContaCliFor
            End Get
            Set(ByVal value As String)
                _DigitoContaCliFor = value
            End Set
        End Property

        '** Deposito Cliente / Fornecedor *******************************************
        Public Property CodigoDepositoCliFor() As String
            Get
                Return _CodigoDepositoCliFor
            End Get
            Set(ByVal value As String)
                _CodigoDepositoCliFor = value
                _DepositoCliFor = Nothing
            End Set
        End Property
        Public Property EndDepositoCliFor() As Integer
            Get
                Return _EndDepositoCliFor
            End Get
            Set(ByVal value As Integer)
                _EndDepositoCliFor = value
            End Set
        End Property
        Public Property DepositoCliFor() As Cliente
            Get
                If _DepositoCliFor Is Nothing And _CodigoDepositoCliFor.Length > 0 Then _CliFor = New Cliente(_CodigoDepositoCliFor, _EndDepositoCliFor)
                Return _DepositoCliFor
            End Get
            Set(ByVal value As Cliente)
                _DepositoCliFor = value
                _CodigoDepositoCliFor = value.Codigo
                _EndDepositoCliFor = value.CodigoEndereco
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

        '** Empresa Recebedora / Pagadora  *****************************************
        Public Property CodigoEmpresaRecPag() As String
            Get
                Return _CodigoEmpresaRecPag
            End Get
            Set(ByVal value As String)
                _CodigoEmpresaRecPag = value
                _EmpresaRecPag = Nothing
            End Set
        End Property
        Public Property EndEmpresaRecPag() As Integer
            Get
                Return _EndEmpresaRecPag
            End Get
            Set(ByVal value As Integer)
                _EndEmpresaRecPag = value
                _EmpresaRecPag = Nothing
            End Set
        End Property
        Public Property EmpresaRecPag() As Cliente
            Get
                If _EmpresaRecPag Is Nothing And _CodigoEmpresaRecPag.Length > 0 Then _EmpresaRecPag = New Cliente(_CodigoEmpresaRecPag, _EndEmpresaRecPag)
                Return _EmpresaRecPag
            End Get
            Set(ByVal value As Cliente)
                _EmpresaRecPag = value
            End Set
        End Property

        '** Unidade de Negocio Recebedora / Pagadora *******************************
        Public Property CodigoUnidadeDeNegocioRecPag() As String
            Get
                Return _CodigoUnidadeDeNegocioRecPag
            End Get
            Set(ByVal value As String)
                _CodigoUnidadeDeNegocioRecPag = value
                _UnidadeDeNegocioRecPag = Nothing
            End Set
        End Property
        Public Property UnidadeDeNegocioRecPag() As Cliente
            Get
                If _UnidadeDeNegocioRecPag Is Nothing And _CodigoUnidadeDeNegocioRecPag > 0 Then _UnidadeDeNegocioRecPag = New Cliente(_CodigoUnidadeDeNegocioRecPag, 0)
                Return _UnidadeDeNegocioRecPag
            End Get
            Set(ByVal value As Cliente)
                _UnidadeDeNegocioRecPag = value
            End Set
        End Property

        '** Depósito Recebedora / Pagadora *****************************************
        Public Property CodigoDepositoRecPag() As String
            Get
                Return _CodigoDepositoRecPag
            End Get
            Set(ByVal value As String)
                _CodigoDepositoRecPag = value
                _DepositoRecPag = Nothing
            End Set
        End Property
        Public Property EndDepositoRecPag() As Integer
            Get
                Return _EndDepositoRecPag
            End Get
            Set(ByVal value As Integer)
                _EndDepositoRecPag = value
            End Set
        End Property
        Public Property DepositoRecPag() As Cliente
            Get
                If _DepositoRecPag Is Nothing And _CodigoDepositoRecPag.Length > 0 Then _CliFor = New Cliente(_CodigoDepositoRecPag, _EndDepositoRecPag)
                Return _DepositoRecPag
            End Get
            Set(ByVal value As Cliente)
                _DepositoRecPag = value
                _CodigoDepositoRecPag = value.Codigo
                _EndDepositoRecPag = value.CodigoEndereco
            End Set
        End Property

        '** Conta Contabil Empresa Recebedora / Pagadora ***************************
        Public Property CodigoContaContabilRecPag() As String
            Get
                Return _CodigoContaContabilRecPag
            End Get
            Set(ByVal value As String)
                _CodigoContaContabilRecPag = value
                _ContaContabilRecPag = Nothing

                If Me.Controlando Then
                    Dim liq As New Novo.TituloXContaContabil(Me)
                    liq.CodigoContaEncargo = ContaContabilRecPag.Conta
                    liq.ContaEncargo = ContaContabilRecPag
                    liq.Descricao = ContaContabilRecPag.Titulo
                    liq.DC = IIf(Me.Valores.EncargoValorDocumento.DC = "D", "C", "D")

                    Me.Valores.Remove(Me.Valores.EncargoValorLiquido)
                    Me.Valores.Add(liq)
                    Me.Valores.AtualizaLiquido()
                End If
            End Set
        End Property
        Public Property ContaContabilRecPag() As PlanoDeConta
            Get
                If _ContaContabilRecPag Is Nothing And _CodigoContaContabilRecPag.Length > 0 Then _ContaContabilRecPag = New PlanoDeConta("", 0, _CodigoContaContabilRecPag)
                Return _ContaContabilRecPag
            End Get
            Set(ByVal value As PlanoDeConta)
                _ContaContabilRecPag = value
            End Set
        End Property

        '** Cheque - Slips - Recibo - Boleto ***************************************
        Public Property Cheque() As Boolean
            Get
                Return _Cheque
            End Get
            Set(ByVal value As Boolean)
                _Cheque = value
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
        Public Property CodigoDeBarras() As String
            Get
                Return _CodigoDeBarras
            End Get
            Set(ByVal value As String)
                _CodigoDeBarras = value
            End Set
        End Property
        Public Property CodigoDeBarrasDigitado() As Boolean
            Get
                Return _CodigoDeBarrasDigitado
            End Get
            Set(ByVal value As Boolean)
                _CodigoDeBarrasDigitado = value
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

        '** Historico - Observacao *************************************************
        Public Property Historico() As String
            Get
                Return _Historico
            End Get
            Set(ByVal value As String)
                _Historico = value
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

        '** Carteira ***************************************************************
        Public Property CodigoCarteiraDoTitulo() As Integer
            Get
                Return _CodigoCarteiraDoTitulo
            End Get
            Set(ByVal value As Integer)
                _CodigoCarteiraDoTitulo = value
                _CarteiraDoTitulo = Nothing
            End Set
        End Property
        Public Property CarteiraDoTitulo() As CarteiraDoTitulo 'Carteira do titulo
            Get
                If _CarteiraDoTitulo Is Nothing Then _CarteiraDoTitulo = New CarteiraDoTitulo(_CodigoCarteiraDoTitulo)
                Return _CarteiraDoTitulo
            End Get
            Set(ByVal value As CarteiraDoTitulo)
                _CarteiraDoTitulo = value
            End Set
        End Property

        '** Situacao Bancaria ******************************************************
        Public Property SituacaoBancaria() As Integer
            Get
                Return _SituacaoBancaria
            End Get
            Set(ByVal value As Integer)
                _SituacaoBancaria = value
            End Set
        End Property

        '** Desdobrar Fornecedor ***************************************************
        Public Property DesdobrarFornecedor() As TituloXDesdobrarFornecedor
            Get
                If _DesdobraFornecedor Is Nothing Then
                    _DesdobraFornecedor = New TituloXDesdobrarFornecedor(Me)
                    If _DesdobraFornecedor.CodigoCliente.Length = 0 Then _DesdobraFornecedor = Nothing
                End If
                Return _DesdobraFornecedor
            End Get
            Set(ByVal value As TituloXDesdobrarFornecedor)
                _DesdobraFornecedor = value
            End Set
        End Property

        '** Indice *****************************************************************
        Public Property IndiceTitulo() As Decimal
            Get
                Return _IndiceTitulo
            End Get
            Set(ByVal value As Decimal)
                If Not Me.Controlando Then
                    _IndiceTitulo = value
                Else
                    If _IndiceTitulo <> value Then
                        _IndiceTitulo = value
                        If Valores IsNot Nothing AndAlso Valores.Count > 0 Then
                            Valores.AtualizaValores()
                        End If
                    End If
                End If
            End Set
        End Property

        '** Contabilizacoes ********************************************************
        Public Property Contabilizacoes() As ListRazao
            Get
                If _Contabilizacoes Is Nothing Then _Contabilizacoes = New ListRazao(Me)
                Return _Contabilizacoes
            End Get
            Set(ByVal value As ListRazao)
                _Contabilizacoes = value
            End Set
        End Property

        Public Property Lote() As Integer
            Get
                If _Lote = 0 Then _Lote = 70
                Return _Lote
            End Get
            Set(ByVal value As Integer)
                _Lote = value
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

        '** Adiantamentos **********************************************************
        Public Property Adiantamento() As Novo.AdiantamentoNovo
            Get
                If _Adiantamento Is Nothing Then _Adiantamento = New Novo.AdiantamentoNovo(Me)
                Return _Adiantamento
            End Get
            Set(ByVal value As Novo.AdiantamentoNovo)
                _Adiantamento = value
            End Set
        End Property
        Public Property Baixas_AdiantamentoEfetuadas() As Novo.ListAdiantamentoBaixaNovo
            Get
                If _Baixas_AdiantamentoEfetuadas Is Nothing And Me.Codigo > 0 Then
                    _Baixas_AdiantamentoEfetuadas = New Novo.ListAdiantamentoBaixaNovo(Me)
                End If
                Return _Baixas_AdiantamentoEfetuadas
            End Get
            Set(ByVal value As Novo.ListAdiantamentoBaixaNovo)
                _Baixas_AdiantamentoEfetuadas = value
            End Set
        End Property

        'Adiantamentos abertos - vc esta no Titulo
        Public Property AdiantamentosAbertos() As Novo.ListAdiantamentoNovo
            Get
                If _AdiantamentosAbertos Is Nothing Then
                    Dim Par As New Hashtable
                    Par.Add("CodigoCliente", Me.CodigoCliFor)
                    Par.Add("EndCliente", Me.EnderecoCliFor)
                    Par.Add("ConsolidarCliente", True)
                    Par.Add("SomenteComSaldo", True)

                    Par.Add("isTroca", False)
                    Par.Add("ContaAdiantamento", "")

                    If Me.ContaContabilCliFor.Adiantamento Then Par("ContaAdiantamento") = Me.ContaContabilCliFor.Conta

                    If Me.Pedido IsNot Nothing AndAlso Me.Pedido.Itens.Count > 0 Then
                        Par.Add("CodigoEmpresa", Me.Pedido.CodigoEmpresa)
                        Par.Add("EndEmpresa", Me.Pedido.EnderecoEmpresa)
                        Par.Add("CodigoPedido", Me.CodigoPedido)

                        If Me.Pedido.Troca Then
                            Par("isTroca") = True

                            If Me.Pedido.ResumoFinanceiro.SaldoAdiantamento = 0 Then
                                Par("ContaAdiantamento") = Me.Pedido.ResumoFinanceiro.ResumoTroca.ContaContabilAdiantamento
                            End If

                            Par.Add("CodigoEmpresaTroca", Me.Pedido.CodigoEmpresaTroca)
                            Par.Add("EndEmpresaTroca", Me.Pedido.EnderecoEmpresaTroca)
                            Par.Add("CodigoPedidoTroca", Me.Pedido.CodigoPedidoTroca)
                        End If
                    End If

                    _AdiantamentosAbertos = New Novo.ListAdiantamentoNovo(Par)
                End If
                Return _AdiantamentosAbertos
            End Get
            Set(ByVal value As Novo.ListAdiantamentoNovo)
                _AdiantamentosAbertos = value
            End Set
        End Property

        '** Lancamentos Contabeis **************************************************
        Public Property CodigoProduto() As String
            Get
                Return _CodigoProduto
            End Get
            Set(ByVal value As String)
                _CodigoProduto = value
                _Produto = Nothing
            End Set
        End Property
        Public ReadOnly Property Produto() As Produto
            Get
                If _Produto Is Nothing And _CodigoProduto.Length > 0 Then _Produto = New Produto(_CodigoProduto)
                Return _Produto
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
        Public Property CodigoClienteRecPag() As String
            Get
                Return _CodigoClienteRecPag
            End Get
            Set(ByVal value As String)
                _CodigoClienteRecPag = value
                _ClienteRecPag = Nothing
            End Set
        End Property
        Public Property EndClienteRecPag() As Integer
            Get
                Return _EndClienteRecPag
            End Get
            Set(ByVal value As Integer)
                _EndClienteRecPag = value
                _ClienteRecPag = Nothing
            End Set
        End Property
        Public Property ClienteRecPag() As Cliente
            Get
                If _ClienteRecPag Is Nothing And _CodigoClienteRecPag.Length > 0 Then _ClienteRecPag = New Cliente(Me.CodigoClienteRecPag, Me.EndClienteRecPag)
                Return _ClienteRecPag
            End Get
            Set(ByVal value As Cliente)
                _ClienteRecPag = value
                _CodigoClienteRecPag = value.Codigo
                _EndClienteRecPag = value.CodigoEndereco
            End Set
        End Property
        Public Property CodigoPedidoRecPag() As Integer
            Get
                Return _CodigoPedidoRecPag
            End Get
            Set(ByVal value As Integer)
                _CodigoPedidoRecPag = value
                _PedidoRecPag = Nothing
            End Set
        End Property
        Public Property PedidoRecPag() As Pedido
            Get
                If _PedidoRecPag Is Nothing And _CodigoPedidoRecPag > 0 Then _PedidoRecPag = New Pedido(_CodigoEmpresaRecPag, _EndEmpresaRecPag, _CodigoPedidoRecPag)
                Return _PedidoRecPag
            End Get
            Set(ByVal value As Pedido)
                _PedidoRecPag = value
            End Set
        End Property

        '** Bordero ****************************************************************
        Public Property Bordero As Novo.Bordero
            Get
                If _Bordero Is Nothing And Me.Codigo > 0 Then
                    _Bordero = New Novo.Bordero(Me.Codigo)
                ElseIf _Bordero Is Nothing And Me.Codigo = 0 Then
                    _Bordero = New Novo.Bordero()
                End If
                Return _Bordero
            End Get
            Set(value As Novo.Bordero)
                _Bordero = value
            End Set
        End Property
        Public Property TituloBorderos As Novo.ListBorderoXTitulo 'tras todos os borderos no qual aquele titulo participu 
            Get
                If _TituloBorderos Is Nothing Then _TituloBorderos = New Novo.ListBorderoXTitulo(Me.Codigo, False)
                Return _TituloBorderos
            End Get
            Set(value As Novo.ListBorderoXTitulo)
                _TituloBorderos = value
            End Set
        End Property
        Public Property UltimoBordero As Novo.BorderoXTitulo 'tras o bordero atual que aquele titulo esta vinculado
            Get
                If _UltimoBordero Is Nothing And TituloBorderos.Count > 0 Then _UltimoBordero = TituloBorderos.Where(Function(s) s.Bordero.CodigoSituacao = 1 And s.CodigoBordero = TituloBorderos.Max(Function(c) c.CodigoBordero)).FirstOrDefault
                Return _UltimoBordero
            End Get
            Set(value As Novo.BorderoXTitulo)
                _UltimoBordero = value
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
        Public Property TituloDestinacao() As Novo.TituloXDestinacao
            Get
                If _TituloDestinacao Is Nothing Then _TituloDestinacao = New Novo.TituloXDestinacao(Me)
                Return _TituloDestinacao
            End Get
            Set(ByVal value As Novo.TituloXDestinacao)
                _TituloDestinacao = value
            End Set
        End Property

        '** Nota Fiscal ************************************************************
        Public Property NotaTitulo As Novo.NotaFiscalXTitulo
            Get
                If _NotaTitulo Is Nothing And _Codigo > 0 Then
                    _NotaTitulo = New Novo.NotaFiscalXTitulo(Me)
                End If
                Return _NotaTitulo
            End Get
            Set(value As Novo.NotaFiscalXTitulo)
                _NotaTitulo = value
            End Set
        End Property

        '*** Titulo Origem *********************************************************
        Public Property CodigoTituloOrigem As Integer
            Get
                Return _CodigoTituloOrigem
            End Get
            Set(value As Integer)
                _CodigoTituloOrigem = value
            End Set
        End Property

        Public Property TituloOrigem As TituloNovo
            Get
                If _TituloOrigem Is Nothing And _CodigoTituloOrigem > 0 Then _TituloOrigem = New Novo.TituloNovo(_CodigoTituloOrigem)
                Return _TituloOrigem
            End Get
            Set(value As TituloNovo)
                _TituloOrigem = value
                If value IsNot Nothing Then _CodigoTituloOrigem = value.Codigo
            End Set
        End Property

        '*** Titulo Compensacao ****************************************************
        Public Property CodigoTituloCompensacao As Integer
            Get
                Return _CodigoTituloCompensacao
            End Get
            Set(value As Integer)
                _CodigoTituloCompensacao = value
            End Set
        End Property

        Public Property TituloCompensacao As TituloNovo
            Get
                If _TituloCompensacao Is Nothing And _CodigoTituloCompensacao > 0 Then _TituloCompensacao = New Novo.TituloNovo(_CodigoTituloCompensacao)
                Return _TituloCompensacao
            End Get
            Set(value As TituloNovo)
                _TituloCompensacao = value
                If value IsNot Nothing Then _CodigoTituloCompensacao = value.Codigo
            End Set
        End Property
        '*** TituloXHistorico ******************************************************
        Public ReadOnly Property TituloHistoricoAcao() As ListTituloXHistorico
            Get
                If _TituloHistorico Is Nothing And _Codigo > 0 Then _TituloHistorico = New Novo.ListTituloXHistorico(Me)
                Return _TituloHistorico
            End Get
        End Property

        Public ReadOnly Property Bloqueio() As Boolean
            Get
                If TituloHistoricoAcao IsNot Nothing AndAlso TituloHistoricoAcao.Count > 0 AndAlso TituloHistoricoAcao.First.Acao = "LIBERAR" Then
                    Return True
                End If
                Return False
            End Get
        End Property

        '******* Valor Previsão ****************************************************
        Public Property ValorPrevisao As Decimal
            Get
                Return _ValorPrevisao
            End Get
            Set(value As Decimal)
                _ValorPrevisao = value
            End Set
        End Property

#End Region

#Region "Methods"
        '********* Salvar *****************************
        Public Function Salvar() As Boolean
            Dim Banco As New AcessaBanco
            If FinanceiroNovo AndAlso Not String.IsNullOrWhiteSpace(Empresa.Empresa.Servidor) AndAlso UsuarioServidor.NomeServidor <> Empresa.Empresa.Servidor Then
                Banco = New AcessaBanco(2, Empresa.Empresa.Servidor)
            End If
            Dim Sqls As New ArrayList
            SalvarSql(Sqls)
            If FinanceiroNovo Then getSqlException(Sqls)
            Return Banco.GravaBanco(Sqls)
        End Function

        Public Function SalvarSql(ByRef sqls As ArrayList, Optional ByVal UsaNumerador As Boolean = True, Optional GravarComoEsta As Boolean = False) As ArrayList
            Dim strSQL As String
            Dim ObjBanco As New AcessaBanco

            If Me.Valores.EncargoValorDocumento.Valor = 0 AndAlso Me.IUD = "I" Then Return sqls
            'Baixa de Adiantamento por Provisão
            If Me.TituloOriginal IsNot Nothing AndAlso Me.TituloOriginal.CodigoProvisao = eProvisao.Provisao AndAlso Me.AdiantamentosAbertos.ValorTotalInformadoParaBaixa > 0 Then
                CriarTituloBaixaAdiantamento(sqls)
            End If

            'Recupera o Numero do Titulo e atualiza o mesmo na tabela Numerador
            Dim numTitulo As New Numerador(1)
            If Me.IUD = "I" And UsaNumerador And Not GravarComoEsta Then
                _Codigo = numTitulo.Sequencia + 1
                sqls.Add(numTitulo.IncrementarNumeradorSql())
                If TitulosAgrupados.Count > 0 Then _RegistroMestre = _Codigo
            End If

            If Me.ContaContabilCliFor.Adiantamento AndAlso Me.TituloOriginal IsNot Nothing AndAlso Me.Pedido IsNot Nothing Then
                If Not Me.TituloOriginal.ContaContabilCliFor.Adiantamento Then
                    CriarTituloAdiantamento(sqls)
                    _CriarAdiantamento = False
                Else
                    If Me.IUD.Equals("C") Then
                        AjustaPrevisaoCancelamentoAdiantamento(sqls)
                    Else
                        AjustarPrevisaoPedidoPorAdiantamento(sqls)
                    End If

                End If
            End If

            Select Case Me.IUD
                Case "I"
                    strSQL = "INSERT INTO Titulos " & vbCrLf & _
                             "  (Titulo_Id, UnidadeDeNegocio, Empresa, EndEmpresa, RecPag, RegistroMestre, Pedido, Fixacao," & vbCrLf & _
                             "   Provisao, Moeda, Indexador, Indice, TipoPagto, Situacao, Movimento, Vencimento, Reprogramacao, DataMoeda, DataBaixa, HoraBaixa, " & vbCrLf & _
                             "   CliFor, EnderecoCliFor, BancoCliFor, AgenciaCliFor, DigitoAgenciaCliFor, ContaCliFor, DigitoContaCliFor, ContaContabilCliFor, EmpresaRecPag, EndEmpresaRecPag, ContaContabilRecPag," & vbCrLf & _
                             "   Cheque, NumeroDoCheque, Slips, Recibo, CodigoDeBarra, CodigoDeBarraDigitado, CodigoDeBarraPreImpresso," & vbCrLf & _
                             "   Historico, Observacoes, CarteiraDoTitulo, SituacaoBancaria, Produto, Quantidade," & vbCrLf & _
                             "   ClienteRecPag, EndClienteRecPag, PedidoRecPag, DepositoCliFor, EndDepositoCliFor, DepositoRecPag, EndDepositoRecPag, UnidadeDeNegocioRecPag,TituloOrigem,TituloCompensacao, ValorPrevisao) " & vbCrLf & _
                             " VALUES (" & Me.Codigo & ",'" & Me.CodigoUnidadeDeNegocio & "','" & Me.CodigoEmpresa & "'," & Me.EnderecoEmpresa & ",'" & Me.ReceberPagar & "'," & IIf(Me.RegistroMestre > 0, Me.RegistroMestre, "NULL") & "," & Me.CodigoPedido.ToSqlNULL & "," & Me.CodigoFixacao.ToSqlNULL & "," & vbCrLf & _
                             Me.CodigoProvisao & "," & Me.CodigoMoeda & "," & Me.CodigoIndexador & "," & Str(Me.IndiceTitulo) & "," & Me.CodigoTipoPgto & "," & Me.CodigoSituacao & ",'" & Me.Movimento.ToString("yyyy-MM-dd") & "','" & Me.Vencimento.ToString("yyyy-MM-dd") & "','" & Me.Reprogramacao.ToString("yyyy-MM-dd") & "','" & Me.DataMoeda.ToString("yyyy-MM-dd") & "','" & Me.DataBaixa.ToString("yyyy-MM-dd") & "','" & Me.DataBaixa.ToString("yyyy-MM-dd") & " " & Date.Now.ToString("HH:mm:ss") & "'," & vbCrLf & _
                             "'" & Me.CodigoCliFor & "'," & Me.EnderecoCliFor & "," & Me.CodigoBancoCliFor & ",'" & CodigoAgenciaCliFor & "','" & Me.DigitoAgenciaCliFor & "','" & Me.ContaCliFor & "','" & Me.DigitoContaCliFor & "','" & Me.CodigoContaContabilCliFor & "','" & Me.CodigoEmpresaRecPag & "'," & Me.EndEmpresaRecPag & ",'" & Me.CodigoContaContabilRecPag & "'," & vbCrLf & _
                             IIf(Me.Cheque, 1, 0) & "," & Me.NumeroDoCheque & "," & IIf(Me.Slips, 1, 0) & "," & IIf(Me.Recibo, 1, 0) & ",'" & Me.CodigoDeBarras & "'," & IIf(Me.CodigoDeBarrasDigitado, 1, 0) & "," & IIf(Me.CodigoDeBarrasPreImpresso, 1, 0) & "," & vbCrLf & _
                             "'" & Me.Historico & "','" & Me.Observacoes & "'," & Me.CodigoCarteiraDoTitulo & "," & Me.CodigoCarteiraDoTitulo & ",'" & Me.CodigoProduto & "'," & Str(Me.Quantidade) & "," & vbCrLf & _
                             "'" & Me.CodigoClienteRecPag & "'," & Me.EndClienteRecPag & "," & Me.CodigoPedidoRecPag & "," & IIf(Not String.IsNullOrEmpty(Me.CodigoDepositoCliFor), "'" & Me.CodigoDepositoCliFor & "'", "NULL") & "," & IIf(Not String.IsNullOrEmpty(Me.CodigoDepositoCliFor), Me.EndDepositoCliFor, "NULL") & "," & IIf(Not String.IsNullOrEmpty(Me.CodigoDepositoRecPag), "'" & Me.CodigoDepositoRecPag & "'", "NULL") & "," & IIf(Not String.IsNullOrEmpty(Me.CodigoDepositoRecPag), Me.EndDepositoRecPag, "NULL") & ",'" & Me.CodigoUnidadeDeNegocioRecPag & "'," & vbCrLf
                    If Me.CodigoTituloOrigem > 0 Then
                        strSQL &= Me.TituloOrigem.Codigo & ","
                    Else
                        strSQL &= "0,"
                    End If
                    If Me.CodigoTituloCompensacao > 0 Then
                        strSQL &= Me.TituloCompensacao.Codigo & ","
                    Else
                        strSQL &= "0,"
                    End If

                    strSQL &= IIf(Me.CodigoProvisao = eProvisao.Previsao, Str(Math.Round(Me.ValorPrevisao, 2)), 0) & ")"

                    sqls.Add(strSQL)
                    SalvarTabelasRelacionadasSql(sqls)
                Case "U"
                    strSQL = " UPDATE Titulos Set " & vbCrLf & _
                             "   UnidadeDeNegocio         ='" & Me.CodigoUnidadeDeNegocio & "'" & vbCrLf & _
                             "  ,Empresa                  ='" & Me.CodigoEmpresa & "'" & vbCrLf & _
                             "  ,EndEmpresa               = " & Me.EnderecoEmpresa & vbCrLf & _
                             "  ,RecPag                   ='" & Me.ReceberPagar & "'" & vbCrLf & _
                             "  ,RegistroMestre           = " & IIf(Me.RegistroMestre > 0, Me.RegistroMestre, "NULL") & vbCrLf & _
                             "  ,Pedido                   = " & Me.CodigoPedido.ToSqlNULL & vbCrLf & _
                             "  ,Fixacao                  = " & Me.CodigoFixacao.ToSqlNULL & vbCrLf & _
                             "  ,Provisao                 = " & Me.CodigoProvisao & vbCrLf & _
                             "  ,Moeda                    = " & Me.CodigoMoeda & vbCrLf & _
                             "  ,Indexador                = " & Me.CodigoIndexador & vbCrLf & _
                             "  ,Indice                   = " & Str(Me.IndiceTitulo) & vbCrLf & _
                             "  ,TipoPagto                = " & Me.CodigoTipoPgto & vbCrLf & _
                             "  ,Situacao                 = " & Me.CodigoSituacao & vbCrLf & _
                             "  ,Movimento                ='" & Me.Movimento.ToString("yyyy-MM-dd") & "'" & vbCrLf & _
                             "  ,Vencimento               ='" & Me.Vencimento.ToString("yyyy-MM-dd") & "'" & vbCrLf & _
                             "  ,Reprogramacao            ='" & Me.Reprogramacao.ToString("yyyy-MM-dd") & "'" & vbCrLf & _
                             "  ,DataMoeda                ='" & Me.DataMoeda.ToString("yyyy-MM-dd") & "'" & vbCrLf & _
                             "  ,DataBaixa                ='" & Me.DataBaixa.ToString("yyyy-MM-dd") & "'" & vbCrLf

                    If Me.TituloOrigem IsNot Nothing AndAlso Me.TituloOriginal IsNot Nothing AndAlso Me.TituloOriginal.CodigoProvisao <> 1 AndAlso Me.CodigoProvisao = 1 Then
                        strSQL &= "  ,HoraBaixa                ='" & Me.DataBaixa.ToString("yyyy-MM-dd") & " " & Date.Now.ToString("HH:mm:ss") & "'" & vbCrLf
                    Else
                        strSQL &= "  ,HoraBaixa                ='" & Me.DataBaixa.ToString("yyyy-MM-dd HH:mm:ss") & "'" & vbCrLf
                    End If

                    strSQL &= "  ,CliFor                   ='" & Me.CodigoCliFor & "'" & vbCrLf & _
                              "  ,EnderecoCliFor           = " & Me.EnderecoCliFor & vbCrLf & _
                              "  ,BancoCliFor              = " & Me.CodigoBancoCliFor & vbCrLf & _
                              "  ,AgenciaCliFor            ='" & CodigoAgenciaCliFor & "'" & vbCrLf & _
                              "  ,DigitoAgenciaCliFor      ='" & Me.DigitoAgenciaCliFor & "'" & vbCrLf & _
                              "  ,ContaCliFor              ='" & Me.ContaCliFor & "'" & vbCrLf & _
                              "  ,DigitoContaCliFor        ='" & Me.DigitoContaCliFor & "'" & vbCrLf & _
                              "  ,ContaContabilCliFor      ='" & Me.CodigoContaContabilCliFor & "'" & vbCrLf & _
                              "  ,DepositoCliFor           = " & IIf(Not String.IsNullOrEmpty(Me.CodigoDepositoCliFor), "'" & Me.CodigoDepositoCliFor & "'", "NULL") & vbCrLf & _
                              "  ,EndDepositoCliFor        = " & IIf(Not String.IsNullOrEmpty(Me.CodigoDepositoCliFor), Me.EndDepositoCliFor, "NULL") & vbCrLf & _
                              "  ,UnidadeDeNegocioRecPag   ='" & Me.CodigoUnidadeDeNegocioRecPag & "'" & vbCrLf & _
                              "  ,EmpresaRecPag            ='" & Me.CodigoEmpresaRecPag & "'" & vbCrLf & _
                              "  ,EndEmpresaRecPag         = " & Me.EndEmpresaRecPag & vbCrLf & _
                              "  ,ContaContabilRecPag      ='" & Me.CodigoContaContabilRecPag & "'" & vbCrLf & _
                              "  ,DepositoRecPag           = " & IIf(Not String.IsNullOrEmpty(Me.CodigoDepositoRecPag), "'" & Me.CodigoDepositoRecPag & "'", "NULL") & vbCrLf & _
                              "  ,EndDepositoRecPag        = " & IIf(Not String.IsNullOrEmpty(Me.CodigoDepositoRecPag), Me.EndDepositoRecPag, "NULL") & vbCrLf & _
                              "  ,Cheque                   = " & CByte(Me.Cheque) & vbCrLf & _
                              "  ,NumeroDoCheque           = " & Me.NumeroDoCheque & vbCrLf & _
                              "  ,Slips                    = " & IIf(Me.Slips, 1, 0) & vbCrLf & _
                              "  ,Recibo                   = " & IIf(Me.Recibo, 1, 0) & vbCrLf & _
                              "  ,CodigoDeBarra            ='" & Me.CodigoDeBarras & "'" & vbCrLf & _
                              "  ,CodigoDeBarraDigitado    = " & IIf(Me.CodigoDeBarrasDigitado, 1, 0) & vbCrLf & _
                              "  ,CodigoDeBarraPreImpresso = " & IIf(Me.CodigoDeBarrasPreImpresso, 1, 0) & vbCrLf & _
                              "  ,Historico                ='" & Me.Historico & "'" & vbCrLf & _
                              "  ,Observacoes              ='" & Me.Observacoes & "'" & vbCrLf & _
                              "  ,CarteiraDoTitulo         = " & Me.CodigoCarteiraDoTitulo & vbCrLf & _
                              "  ,SituacaoBancaria         = " & Me.CodigoCarteiraDoTitulo & vbCrLf & _
                              "  ,Produto                  ='" & Me.CodigoProduto & "'" & vbCrLf & _
                              "  ,Quantidade               = " & Str(Me.Quantidade) & vbCrLf & _
                              "  ,ClienteRecPag            ='" & Me.CodigoClienteRecPag & "'" & vbCrLf & _
                              "  ,EndClienteRecPag         = " & Me.EndClienteRecPag & vbCrLf & _
                              "  ,PedidoRecPag             = " & Me.CodigoPedidoRecPag & vbCrLf & _
                              "  ,TituloOrigem             = " & Me.CodigoTituloOrigem & vbCrLf & _
                              "  ,TituloCompensacao        = " & Me.CodigoTituloCompensacao & vbCrLf & _
                              " WHERE Titulo_Id = " & _Codigo & vbCrLf
                    sqls.Add(strSQL)
                    SalvarTabelasRelacionadasSql(sqls)
                Case "D"
                    SalvarTabelasRelacionadasSql(sqls)
                    strSQL = "Delete Titulos WHERE Titulo_Id = " & _Codigo & ""
                    sqls.Add(strSQL)
                Case "C"
                    SalvarTabelasRelacionadasSql(sqls)
                    strSQL = "UPDATE Titulos Set Situacao = 3 WHERE Titulo_Id = " & _Codigo & ""
                    sqls.Add(strSQL)
                Case Else
                    SalvarTabelasRelacionadasSql(sqls)
            End Select


            Return sqls
        End Function

        Private Sub SalvarTabelasRelacionadasSql(ByRef Sqls As ArrayList)
            '************************************************************************************************************************************
            '********* AJUSTA O MESTRE E OS FILHOS **********************************************************************************************
            '************************************************************************************************************************************
            AjustaTitulosAgrupados(Sqls)

            '************************************************************************************************************************************
            '********************** GRAVA RECOMPRA **********************************************************************************************
            '************************************************************************************************************************************
            If Me.IUD = "I" AndAlso Not Me.TituloOriginal Is Nothing AndAlso Me.TituloOriginal.UltimoBordero IsNot Nothing Then
                Me.TituloOriginal.IUD = "U"
                Me.TituloOriginal.UltimoBordero.CodigoTituloRecompra = Me.Codigo
                Me.TituloOriginal.CodigoContaContabilRecPag = Me.Empresa.Empresa.CodigoContaGrupoBanco
                Me.TituloOriginal.CodigoCarteiraDoTitulo = 0
                Me.TituloOriginal.SalvarSql(Sqls)
            End If

            '************************************************************************************************************************************
            '********* AJUSTA O BORDERO E DUPLICATAS ********************************************************************************************
            '************************************************************************************************************************************
            AjustaTitulosBorderoDuplicatas(Sqls)

            '***************************************************************************************************************************************************************************
            '********* SALVA O VINCULO DO TITULO COM A NOTA  / USADO QUANDO UM TITULO E PAGO PARCIALMENTE E UM NOVO TITULO COM O RESTANTE DO VALOR É GERADO ****************************
            '***************************************************************************************************************************************************************************
            If NotaTitulo.Existe AndAlso NotaTitulo.NotaFiscal.Codigo > 0 Then
                If String.IsNullOrWhiteSpace(NotaTitulo.IUD) Then NotaTitulo.IUD = Me.IUD
                NotaTitulo.SalvarSql(Sqls)
            End If

            '************************************************************************************************************************************
            '********* SALVA VALORES DOS ENCARGOS  **********************************************************************************************
            '************************************************************************************************************************************
            Valores.SalvarSQL(Sqls)

            '****************************************************************
            '****** Salva Titulo x Destinacao "PROCURACAO" ******************
            '****************************************************************
            If Me.IUD = "I" And TituloDestinacao.CodigoProcuracao > 0 Then
                TituloDestinacao.IUD = "I"
                TituloDestinacao.SalvarSql(Sqls)
            ElseIf Me.IUD = "U" And TituloDestinacao.CodigoProcuracao = 0 Then
                TituloDestinacao.IUD = "D"
                TituloDestinacao.SalvarSql(Sqls)
            ElseIf Me.IUD = "D" Then
                TituloDestinacao.IUD = "D"
                TituloDestinacao.SalvarSql(Sqls)
            ElseIf TituloDestinacao.CodigoDestinatario.Length > 0 Then
                TituloDestinacao.IUD = Me.IUD
                TituloDestinacao.SalvarSql(Sqls)
            End If

            '****************************************************************
            '*********************   ADIANTAMENTOS   ************************ 
            '****************************************************************
            CriarSqlAdiantamento(Sqls)
            If Me.TituloOriginal Is Nothing AndAlso Me.CodigoProvisao <> eProvisao.Provisao Then
                CriarSqlBaixaAdiantamento(Sqls)
            End If

            If Not Me.DesdobrarFornecedor Is Nothing Then
                If Me.CodigoProvisao <> 1 Then Me.DesdobrarFornecedor.IUD = "D"
                Me.DesdobrarFornecedor.SalvarSql(Sqls)
            End If

            If Codigo = RegistroMestre Or RegistroMestre = 0 Then
                If Me.CodigoProvisao = 1 Then
                    Me.Contabilizacoes.ApurarContabilizacaoTitulo()
                    Me.Contabilizacoes.SalvarSqlTitulo(Sqls)
                ElseIf Me.Contabilizacoes.Count > 0 Then
                    Me.Contabilizacoes.SalvarSqlTitulo(Sqls)
                End If
            End If

            'Histórico do Título
            HistoricoTitulo(Sqls)

            '*******************************************************************************************************************************************
            '********* CRIA UM NOVO TITULO CASO SEJA SALVO O TITULO ORIGINAL PARCIALMENTE **************************************************************
            '********* (CodigoProvisao = 1 AndAlso Not TitulosAgrupados.Count > 0) Insere um titulo parcial quando for Baixa e não for agrupamento *****
            '*******************************************************************************************************************************************
            If ((CodigoPedido > 0) OrElse (CodigoProvisao = 1 AndAlso Not TitulosAgrupados.Count > 0)) AndAlso Not TituloOriginal Is Nothing AndAlso _CriarAdiantamento Then
                If Moeda.Classificacao = eTiposMoeda.Oficial Then
                    If Me.Valores.EncargoValorDocumento.ValorOficial < TituloOriginal.Valores.EncargoValorDocumento.ValorOficial Then CriarNovoTitulo(Sqls)
                Else
                    If Me.Valores.EncargoValorDocumento.ValorMoeda < TituloOriginal.Valores.EncargoValorDocumento.ValorMoeda Then CriarNovoTitulo(Sqls)
                End If
            End If
        End Sub

        '******** Auxiliares ************************
        Public Sub VerificarIndice()
            If Me.CodigoIndexador = 99 Then Exit Sub

            If Me.CodigoProvisao = eProvisao.Baixa Then
                _DataMoeda = Me.DataBaixa
            Else
                _DataMoeda = Me.Reprogramacao
            End If

            If Me.CodigoIndexador > 0 Then
                Dim objCotacao As New Cotacao(Me.CodigoIndexador, Me.DataMoeda)
                If objCotacao IsNot Nothing Then
                    IndiceTitulo = objCotacao.Indice
                End If
            End If

            If CodigoFixacao > 0 AndAlso Fixacao IsNot Nothing AndAlso Fixacao.IndiceFixado > 0 Then
                IndiceTitulo = Fixacao.IndiceFixado
                _DataMoeda = Fixacao.Movimento
            End If

        End Sub

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
                it = IIf(Adiantamento.Taxa > 0, Adiantamento.Taxa, IIf(Pedido.Taxa > 0, Pedido.Taxa, New Safra(Pedido.CodigoSafra).Taxa))
            Else
                it = Adiantamento.Taxa
            End If

            If it = 0 Then Exit Sub
            q = 1
            t = 30
            iq = (((1 + (it / 100)) ^ (q / t)) - 1) * 100
            'Vf = vi(1 + i) ^ n
            '? kitio
            If Moeda.Classificacao = eTiposMoeda.Oficial Then
                'Juros = Math.Abs(_ValorLiquido - (_ValorLiquido * (1 + iq / 100) ^ dias))
            Else
                'MoedaJuros = Math.Abs(_MoedaValorLiquido - (_MoedaValorLiquido * (1 + iq / 100) ^ dias))
            End If
        End Sub

        '*********************************************
        '******** Cria Titulos ***********************
        '*********************************************
        'Cria um Titulo quando esse titulo é de um Pedido e o Valor do Documento do mesmo é alterado para um Valor Menor
        Public Sub CriarNovoTitulo(ByRef Sqls As ArrayList)
            Dim CodigoTituloOriginal As String = Me.TituloOriginal.Codigo

            Me.IUD = "I"
            Me.Codigo = 0
            Me.RegistroMestre = 0

            Me.CodigoProvisao = Me.TituloOriginal.CodigoProvisao

            Me.Cheque = False
            Me.NumeroDoCheque = 0
            Me.Slips = False
            Me.Recibo = False

            Me.CodigoDeBarras = ""
            Me.CodigoDeBarrasDigitado = False
            Me.CodigoDeBarrasPreImpresso = False

            Me.Historico = "Titulo Criado Pela Alteracao do Capital do Titulo " & CodigoTituloOriginal
            Me.Observacoes = ""

            'Caso o Titulo Orginal tenha pedido, a conta contabil tem q ser a conta contabil do produto
            If CodigoPedido > 0 Then
                Me.CodigoContaContabilCliFor = Pedido.SubOperacao.CodigoGrupoContas
                Me.Valores.EncargoValorDocumento.CodigoContaEncargo = Pedido.SubOperacao.CodigoGrupoContas
            End If

            If Me.TituloOriginal.NotaTitulo.Existe Then
                Me.NotaTitulo.IUD = "I"
            End If

            If Me.Moeda.Classificacao = eTiposMoeda.Oficial Then
                Me.Valores.EncargoValorDocumento.ValorOficial = Me.TituloOriginal.Valores.EncargoValorDocumento.ValorOficial - Me.Valores.EncargoValorDocumento.ValorOficial
            Else
                Me.Valores.EncargoValorDocumento.ValorOficial = Me.TituloOriginal.Valores.EncargoValorDocumento.ValorOficial - Me.Valores.EncargoValorDocumento.ValorOficial
            End If
            Me.Valores.AtualizaValores()

            For Each vlr In Me.Valores
                If Not vlr.Equals(Me.Valores.EncargoValorDocumento) And Not vlr.Equals(Me.Valores.EncargoValorLiquido) Then
                    vlr.ValorOficial = 0
                    vlr.ValorMoeda = 0
                End If
            Next
            Me.Valores.AtualizaLiquido()
            Me.TituloOriginal = Nothing
            Me.SalvarSql(Sqls)
            Me.MsgControle = "Titulo Numero " & Me.Codigo & " Criado pela Alteracao de Capital do Titulo " & CodigoTituloOriginal
        End Sub

        Public Sub CriarTituloAdiantamento(ByRef Sqls As ArrayList)
            If TituloOriginal IsNot Nothing Then
                Dim CodigoTituloOriginal As String = Me.TituloOriginal.Codigo
                Dim TitAdiantamento As New TituloNovo(Me.Codigo)

                TitAdiantamento.IUD = "I"
                TitAdiantamento.Codigo = 0
                TitAdiantamento.RegistroMestre = 0

                TitAdiantamento.CodigoProvisao = Me.CodigoProvisao

                TitAdiantamento.Cheque = False
                TitAdiantamento.NumeroDoCheque = 0
                TitAdiantamento.Slips = False
                TitAdiantamento.Recibo = False

                TitAdiantamento.CodigoDeBarras = ""
                TitAdiantamento.CodigoDeBarrasDigitado = False
                TitAdiantamento.CodigoDeBarrasPreImpresso = False

                TitAdiantamento.Historico = "Adiantamento Criado Pela Alteracao do Capital do Titulo " & CodigoTituloOriginal & " - " & Me.Historico
                TitAdiantamento.Observacoes = Me.Observacoes

                TitAdiantamento.CodigoContaContabilCliFor = Me.CodigoContaContabilCliFor
                TitAdiantamento.CodigoContaContabilRecPag = Me.CodigoContaContabilRecPag
                TitAdiantamento.CodigoTipoPgto = Me.CodigoTipoPgto
                'Banco
                TitAdiantamento.CodigoBancoCliFor = Me.CodigoBancoCliFor
                TitAdiantamento.CodigoAgenciaCliFor = Me.CodigoAgenciaCliFor
                TitAdiantamento.DigitoAgenciaCliFor = Me.DigitoAgenciaCliFor
                TitAdiantamento.ContaCliFor = Me.ContaCliFor
                TitAdiantamento.DigitoContaCliFor = Me.DigitoContaCliFor
                TitAdiantamento.CodigoTituloOrigem = Me.Codigo
                'Empresa Pagadora/Recebedora
                TitAdiantamento.CodigoUnidadeDeNegocioRecPag = Me.CodigoUnidadeDeNegocioRecPag
                TitAdiantamento.CodigoEmpresaRecPag = Me.CodigoEmpresaRecPag

                Dim valorAdiantamento As Decimal = 0

                If Moeda.Classificacao = eTiposMoeda.Oficial Then
                    If Me.Valores.EncargoValorDocumento.ValorOficial = TituloOriginal.Valores.EncargoValorDocumento.ValorOficial Then
                        TitAdiantamento.Valores.EncargoValorDocumento.ValorOficial = Me.TituloOriginal.Valores.EncargoValorDocumento.ValorOficial
                        Me.IUD = "C"
                    Else
                        valorAdiantamento = Me.Valores.EncargoValorDocumento.ValorOficial
                        TitAdiantamento.Valores.EncargoValorDocumento.ValorOficial = valorAdiantamento
                        Me.Valores.EncargoValorDocumento.ValorOficial = TituloOriginal.Valores.EncargoValorDocumento.ValorOficial - valorAdiantamento
                    End If
                Else
                    If Me.Valores.EncargoValorDocumento.ValorMoeda = TituloOriginal.Valores.EncargoValorDocumento.ValorMoeda Then
                        TitAdiantamento.Valores.EncargoValorDocumento.ValorMoeda = Me.TituloOriginal.Valores.EncargoValorDocumento.ValorMoeda
                        Me.IUD = "C"
                    Else
                        valorAdiantamento = Me.Valores.EncargoValorDocumento.ValorMoeda
                        TitAdiantamento.Valores.EncargoValorDocumento.ValorMoeda = valorAdiantamento
                        Me.Valores.EncargoValorDocumento.ValorMoeda = TituloOriginal.Valores.EncargoValorDocumento.ValorMoeda - valorAdiantamento
                    End If
                End If

                TitAdiantamento.Valores.AtualizaValores()

                For Each vlr In TitAdiantamento.Valores
                    If Not vlr.Equals(TitAdiantamento.Valores.EncargoValorDocumento) And Not vlr.Equals(TitAdiantamento.Valores.EncargoValorLiquido) Then
                        vlr.ValorOficial = 0
                        vlr.ValorMoeda = 0
                    End If
                Next
                TitAdiantamento.Valores.AtualizaLiquido()
                TitAdiantamento.TituloOriginal = Nothing

                TitAdiantamento.Adiantamento.Taxa = Me.Adiantamento.Taxa
                TitAdiantamento.Adiantamento.Vencimento = Me.Adiantamento.Vencimento

                TitAdiantamento.SalvarSql(Sqls)


                Me.MsgControle = "Adiantamento " & TitAdiantamento.Codigo & " Criado pela Alteracao de Capital do Titulo " & CodigoTituloOriginal
                'Previsão do Pedido
                Me.CodigoContaContabilCliFor = _TituloOriginal.CodigoContaContabilCliFor
                Me.CodigoProvisao = _TituloOriginal.CodigoProvisao
                'banco
                Me.CodigoBancoCliFor = _TituloOriginal.CodigoBancoCliFor
                Me.CodigoAgenciaCliFor = _TituloOriginal.CodigoAgenciaCliFor
                Me.DigitoAgenciaCliFor = _TituloOriginal.DigitoAgenciaCliFor
                Me.ContaCliFor = _TituloOriginal.ContaCliFor
                Me.DigitoContaCliFor = _TituloOriginal.DigitoContaCliFor
                Me.CodigoContaContabilRecPag = _TituloOriginal.CodigoContaContabilRecPag
                'Empresa
                Me.CodigoEmpresaRecPag = _TituloOriginal.CodigoEmpresaRecPag
                Me.CodigoUnidadeDeNegocioRecPag = _TituloOriginal.CodigoUnidadeDeNegocioRecPag
                Me.CodigoTipoPgto = _TituloOriginal.CodigoTipoPgto
            ElseIf _CodigoPedido = 0 Then
                'Me.Adiantamento.IUD = "I"
                'Me.Adiantamento.CodigoTitulo = Me.Codigo
                'Me.Adiantamento.Atualizacao = Me.Movimento
                'Me.Adiantamento.Vencimento = Me.Vencimento
                'Me.Adiantamento.SalvarSql(Sqls)
            End If
        End Sub

        Public Sub CriarNovoTituloBaixaAdiantamento(ByRef Sqls As ArrayList, ByVal Titulo As Novo.TituloNovo, ByVal valor As Decimal, NumeroDoAdiantamento As Integer, Optional ByVal Financeiro As Boolean = False)
            If valor = 0 Then Exit Sub
            Dim CodigoTituloOriginal As String = Titulo.Codigo

            Me.IUD = "I"
            Me.Codigo = 0
            Me.RegistroMestre = 0

            Me.ReceberPagar = Titulo.ReceberPagar ' IIf(Titulo.ReceberPagar = "R", "P", "R")
            Me.CodigoProvisao = 2 'Previsão

            Me.Cheque = False
            Me.NumeroDoCheque = 0
            Me.Slips = False
            Me.Recibo = False

            Me.CodigoDeBarras = ""
            Me.CodigoDeBarrasDigitado = False
            Me.CodigoDeBarrasPreImpresso = False

            Me.Historico = "Titulo criado pela baixa do adiantamento " & CodigoTituloOriginal
            Me.Observacoes = ""

            'Caso o Titulo Orginal tenha pedido, a conta contabil tem q ser a conta contabil do produto
            If Titulo.CodigoPedido > 0 Then
                Me.CodigoPedido = Titulo.CodigoPedido
                Me.CodigoContaContabilCliFor = Titulo.Pedido.SubOperacao.CodigoGrupoContas
                Me.Valores.EncargoValorDocumento.CodigoContaEncargo = Titulo.Pedido.SubOperacao.CodigoGrupoContas
            End If

            If Me.Moeda.Classificacao = eTiposMoeda.Oficial Then
                Me.Valores.EncargoValorDocumento.ValorOficial = valor
            Else
                Me.Valores.EncargoValorDocumento.ValorOficial = valor
            End If
            Me.Valores.AtualizaValores()

            For Each vlr In Me.Valores
                If Not vlr.Equals(Me.Valores.EncargoValorDocumento) And Not vlr.Equals(Me.Valores.EncargoValorLiquido) Then
                    vlr.ValorOficial = 0
                    vlr.ValorMoeda = 0
                End If
            Next
            Me.Valores.AtualizaLiquido()
            Me.TituloOriginal = Nothing

            _Codigo = NumeroDoAdiantamento

            Me.SalvarSql(Sqls, False)
            Me.MsgControle = "Titulo Numero " & Me.Codigo & " Criado pela baixa do adiantamento " & CodigoTituloOriginal
        End Sub

        '****  Duplicatas BORDERO *********
        Private Sub AjustaTitulosBorderoDuplicatas(ByRef Sqls As ArrayList)
            If Me.Bordero.TitulosDoBordero.Count <= 0 Then Exit Sub

            Me.Bordero.IUD = Me.IUD
            Me.Bordero.CodigoTituloBordero = Me.Codigo
            Me.Bordero.TituloDoBordero = Me
            Me.Bordero.CodigoCarteiraDoTitulo = Me.CodigoCarteiraDoTitulo
            Me.Bordero.SalvarSql(Sqls)

            'Titulos que compoem o Bordero
            'If Me.Codigo = Me.Bordero.CodigoTituloBordero Then
            '    For Each BxT In Me.Bordero.TitulosDoBordero
            '        If Me.IUD = "D" Then
            '            BxT.IUD = "D"
            '        Else
            '            BxT.CodigoBordero = Me.Bordero.CodigoBordero
            '            BxT.Titulo.IUD = "U"
            '            BxT.Titulo.CodigoEmpresaRecPag = Me.CodigoEmpresaRecPag
            '            BxT.Titulo.EndEmpresaRecPag = Me.EndEmpresaRecPag
            '            BxT.Titulo.CodigoTipoPgto = Me.CodigoTipoPgto
            '            BxT.Titulo.CodigoContaContabilRecPag = Me.CodigoContaContabilCliFor
            '            BxT.Titulo.CodigoCarteiraDoTitulo = Me.CodigoCarteiraDoTitulo
            '        End If
            '    Next
            '    'If Me.IUD = "I" Then Me.Bordero.IUD = "I"
            '    Me.Bordero.SalvarSql(Sqls) ' TitulosDoBordero.SalvarSQL(Sqls)
            'End If
        End Sub

        Public Sub AtualizaValoresBordero()
            If Bordero.TitulosDoBordero.Count = 0 Then Exit Sub

            Dim vlrDocumentoOficial As Decimal
            Dim vlrDocumentoMoeda As Decimal

            Dim vlrDocumentoOficialRecompra As Decimal
            Dim vlrDocumentoMoedaRecompra As Decimal

            For Each BxT In Me.Bordero.TitulosDoBordero
                vlrDocumentoOficial += BxT.Titulo.Valores.EncargoValorLiquido.ValorOficial
                vlrDocumentoMoeda += BxT.Titulo.Valores.EncargoValorLiquido.ValorMoeda
            Next

            If Me.Bordero.TitulosRecomprados.Count > 0 Then
                For Each BxT2 In Me.Bordero.TitulosRecomprados
                    vlrDocumentoOficialRecompra += BxT2.Titulo.Valores.EncargoValorLiquido.ValorOficial
                    vlrDocumentoMoedaRecompra += BxT2.Titulo.Valores.EncargoValorLiquido.ValorMoeda
                Next

                Dim TxC As TituloXContaContabil = Me.Valores.Find(Function(s) s.CodigoContaEncargo = Me.Valores.EncargoValorDocumento.CodigoContaEncargo And s.DC = "D")

                If Not IsNothing(TxC) Then
                    TxC.ValorOficial = vlrDocumentoOficialRecompra
                    TxC.ValorMoeda = vlrDocumentoMoedaRecompra
                Else
                    TxC = New TituloXContaContabil(Me)
                    TxC.CodigoContaEncargo = Me.Valores.EncargoValorDocumento.CodigoContaEncargo
                    TxC.DC = "D"
                    TxC.Descricao = Me.Valores.EncargoValorDocumento.Descricao
                    TxC.ValorOficial = vlrDocumentoOficialRecompra
                    TxC.ValorMoeda = vlrDocumentoMoedaRecompra
                    Me.Valores.Insert(1, TxC)
                End If

            End If


            Me.Valores.EncargoValorDocumento.ValorOficial = vlrDocumentoOficial
            Me.Valores.EncargoValorDocumento.ValorMoeda = vlrDocumentoMoeda
            Valores.AtualizaLiquido()
        End Sub

        '****  Titulos Agrupados  *********
        Private Sub AjustaTitulosAgrupados(ByRef Sqls As ArrayList)
            If Me.TitulosAgrupados.Count > 0 Then
                For Each TitFilho As TituloNovo In Me.TitulosAgrupados
                    TitFilho.IUD = "U"
                    If Me.IUD = "D" Then
                        TitFilho.RegistroMestre = 0
                    Else
                        TitFilho.RegistroMestre = Me.Codigo

                        TitFilho.CodigoUnidadeDeNegocioRecPag = Me.CodigoUnidadeDeNegocioRecPag
                        TitFilho.CodigoEmpresaRecPag = Me.CodigoEmpresaRecPag
                        TitFilho.EndEmpresaRecPag = Me.EndEmpresaRecPag
                        TitFilho.CodigoTipoPgto = Me.CodigoTipoPgto
                        TitFilho.CodigoContaContabilRecPag = Me.CodigoContaContabilRecPag
                        TitFilho.CodigoProvisao = Me.CodigoProvisao
                        TitFilho.Reprogramacao = Me.Reprogramacao
                        TitFilho.DataBaixa = Me.DataBaixa
                        If TitFilho.CodigoIndexador <> 99 Then
                            TitFilho.DataMoeda = IIf(Me.CodigoProvisao = 1, Me.DataBaixa, Me.Reprogramacao)
                        End If
                    End If
                Next
                Me.TitulosAgrupados.SalvarSql(Sqls)
            End If
        End Sub

        Public Sub AtualizaValoresAgrupados()
            If TitulosAgrupados.Count = 0 Then Exit Sub

            Dim vlrDocumentoOficial As Decimal
            Dim vlrDocumentoMoeda As Decimal
            For Each Tit As Novo.TituloNovo In _TitulosAgrupados.Where(Function(s) Not s.IUD = "D")
                vlrDocumentoOficial += Tit.Valores.EncargoValorLiquido.ValorOficial
                vlrDocumentoMoeda += Tit.Valores.EncargoValorLiquido.ValorMoeda
            Next
            Me.Valores.EncargoValorDocumento.ValorOficial = vlrDocumentoOficial
            Me.Valores.EncargoValorDocumento.ValorMoeda = vlrDocumentoMoeda
            Valores.AtualizaLiquido()
        End Sub

        '*******  Adiantamentos  **********
        Private Sub CriarSqlAdiantamento(ByRef Sqls As ArrayList)
            If Me.Codigo = Me.RegistroMestre Then Exit Sub
            If Adiantamento.CodigoTitulo = 0 And Me.isAdiantamento And Me.CodigoProvisao = 1 Then
                Me.Adiantamento.IUD = "I"
                Me.Adiantamento.CodigoTitulo = Me.Codigo
                Me.Adiantamento.Atualizacao = Me.Movimento
                Me.Adiantamento.Vencimento = Me.Vencimento
                Me.Adiantamento.SalvarSql(Sqls)
            ElseIf Adiantamento.CodigoTitulo > 0 And Me.CodigoProvisao = 1 Then
                Me.Adiantamento.IUD = "U"
                Me.Adiantamento.SalvarSql(Sqls)
            ElseIf Me.Adiantamento.CodigoTitulo > 0 And Me.CodigoProvisao <> 1 AndAlso Me.Adiantamento.CodigoTitulo = _Codigo Then
                Me.Adiantamento.IUD = "D"
                Me.Adiantamento.SalvarSql(Sqls)
            End If
        End Sub

        Private Sub CriarSqlBaixaAdiantamento(ByRef Sqls As ArrayList)
            If (Me.CodigoProvisao <> 1 OrElse Me.IUD = "D" OrElse Me.IUD = "C") Then
                If Me.Baixas_AdiantamentoEfetuadas.Count > 0 Then
                    For Each row As Novo.AdiantamentoBaixaNovo In Me.Baixas_AdiantamentoEfetuadas
                        row.IUD = "D"
                        row.SalvarSql(Sqls)
                    Next
                End If
            ElseIf Me.CodigoProvisao = 1 AndAlso Me.isBaixaAdiantamento AndAlso Me.Baixas_AdiantamentoEfetuadas.Count > 0 Then
                For Each row As Novo.AdiantamentoBaixaNovo In Me.Baixas_AdiantamentoEfetuadas
                    If row.IUD <> "" Then
                        row.SalvarSql(Sqls)
                    End If
                Next
            ElseIf Me.CodigoProvisao = 1 Then
                If isBaixaAdiantamento AndAlso Me.AdiantamentosAbertos.Count > 0 Then
                    Dim numAdiant As Integer = 1
                    Dim numTitulo As New Numerador(1)
                    For Each row As Novo.AdiantamentoNovo In Me.AdiantamentosAbertos
                        If row.VlrBaixa > 0 Then
                            Dim BxAd As New Novo.AdiantamentoBaixaNovo(row)
                            BxAd.IUD = "I"
                            BxAd.Sequencia = BxAd.Adiantamento.Baixas.ProximaSequenciaLivre
                            BxAd.CodigoTituloBaixa = Me.Codigo
                            BxAd.Movimento = Me.Movimento
                            BxAd.Lancamento = "B"
                            BxAd.TipoLancamento = ""
                            If Me.Moeda.Classificacao = eTiposMoeda.Oficial Then
                                BxAd.ValorOficial = row.VlrBaixa
                                BxAd.ValorMoeda = Funcoes.ConverteParaMoedaExtrangeira(row.VlrBaixa, 2, row.Titulo.Reprogramacao) 'row.VlrBaixa / IndiceTitulo
                            Else
                                BxAd.ValorMoeda = row.VlrBaixa
                                BxAd.ValorOficial = Funcoes.ConverteParaMoedaExtrangeira(row.VlrBaixa, 1, row.Titulo.Reprogramacao) 'row.VlrBaixa * IndiceTitulo
                            End If

                            Dim valor As Decimal = Math.Abs(row.VlrBaixa)
                            'Adiciona o valor da baixa de adto em previsão para adiantamentos pagos em dinheiro e nao em produtos atraves de nota
                            If row.Titulo.Pedido IsNot Nothing AndAlso Not NotaTitulo.Existe Then
                                Dim tit As New Novo.TituloNovo(row.Titulo.CodigoTituloOrigem)
                                tit.IUD = "U"
                                If tit.CodigoSituacao = eSituacao.Normal Then
                                    tit.Valores.EncargoValorDocumento.Valor += valor
                                Else
                                    tit.CodigoSituacao = eSituacao.Normal
                                    tit.Valores.EncargoValorDocumento.Valor = valor
                                End If
                                valor = 0
                                tit.SalvarSql(Sqls)
                                'For Each tit As Novo.Titulo In Me.Pedido.Titulos.Where(Function(s) s.CodigoProvisao = 2)
                                '    If valor > 0 Then
                                '        Dim tit As New Novo.Titulo(row.Titulo.CodigoTituloOrigem)


                                '        If tit.Pedido.Moeda.Classificacao = eTiposMoeda.Oficial Then
                                '            If tit.Valores.EncargoValorDocumento.ValorOficial > valor Then
                                '                tit.IUD = "U"
                                '                tit.Valores.EncargoValorDocumento.ValorOficial += valor
                                '                valor -= tit.Valores.EncargoValorDocumento.ValorOficial
                                '            End If
                                '        Else
                                '            If tit.Valores.EncargoValorDocumento.ValorMoeda > valor Then
                                '                tit.IUD = "U"
                                '                tit.Valores.EncargoValorDocumento.ValorMoeda += valor
                                '            End If
                                '        End If
                                '        tit.SalvarSql(Sqls)
                                '    Else
                                '        Exit For
                                '    End If
                                'Next
                            End If

                            BxAd.SalvarSql(Sqls)
                            Me.Baixas_AdiantamentoEfetuadas.Add(BxAd)

                            'If row.Titulo.NotaTitulo.Existe AndAlso Not Me.PedidoTroca Then
                            'se nao tiver nota é pq foi criado pelo financeiro
                            If valor > 0 AndAlso Me.NotaTitulo.NotaFiscal IsNot Nothing AndAlso Not Me.NotaTitulo.NotaFiscal.NFG AndAlso Not Me.isBaixaAdiantamento Then CriarNovoTituloBaixaAdiantamento(Sqls, row.Titulo, valor, numTitulo.Sequencia + numAdiant, True)
                            numAdiant += 1
                            'End If
                        End If
                    Next

                    Sqls.Add(numTitulo.IncrementarNumeradorSql(True, numAdiant + 1))
                    Me.AdiantamentosAbertos = Nothing
                End If

            End If

        End Sub

        'Cria titulo para baixar o adiantamento
        Public Sub CriarTituloBaixaAdiantamento(ByRef Sqls As ArrayList)
            Dim numTitulo As New Numerador(1)
            Dim numBaixa As Integer = 1
            For Each tit In Me.AdiantamentosAbertos
                'cria o titulo para baixa
                Dim titBaixa As New Novo.TituloNovo(Me.Codigo)

                titBaixa.IUD = "I"
                titBaixa.Codigo = numTitulo.Sequencia + numBaixa
                titBaixa.Valores.EncargoValorDocumento.Valor = tit.VlrBaixa
                titBaixa.CodigoContaContabilRecPag = tit.Titulo.CodigoContaContabilCliFor
                titBaixa.CodigoProvisao = eProvisao.Baixa
                titBaixa.Historico = "Titulo criado pela baixa do adiantamento " & tit.Titulo.Codigo
                titBaixa.Observacoes = String.Empty
                titBaixa.ReceberPagar = IIf(Me.ReceberPagar = "R", "P", "R")
                titBaixa.AdiantamentosAbertos.Clear()
                titBaixa.SalvarSql(Sqls, False)
                'cria baixa
                Dim BxAd As New Novo.AdiantamentoBaixaNovo(tit)
                BxAd.IUD = "I"
                BxAd.Sequencia = BxAd.Adiantamento.Baixas.ProximaSequenciaLivre
                BxAd.CodigoTituloBaixa = titBaixa.Codigo
                BxAd.Movimento = Me.Movimento
                BxAd.Lancamento = "B"
                BxAd.TipoLancamento = ""
                If Me.Moeda.Classificacao = eTiposMoeda.Oficial Then
                    BxAd.ValorOficial = tit.VlrBaixa
                    'BxAd.ValorMoeda = tit.VlrBaixa / IndiceTitulo
                Else
                    BxAd.ValorMoeda = tit.VlrBaixa
                    BxAd.ValorOficial = tit.VlrBaixa * IndiceTitulo
                End If

                BxAd.SalvarSql(Sqls)
                numBaixa += 1
            Next

            Me._CriarAdiantamento = False
            If Me.Valores.EncargoValorDocumento.Valor = Me.AdiantamentosAbertos.ValorTotalInformadoParaBaixa Then
                Me.IUD = "D"
            Else
                Me.Valores.EncargoValorDocumento.Valor -= Me.AdiantamentosAbertos.ValorTotalInformadoParaBaixa
            End If
            Sqls.Add(numTitulo.IncrementarNumeradorSql(True, numBaixa + 1))
        End Sub

        Private Sub AjustarPrevisaoPedidoPorAdiantamento(ByRef Sqls As ArrayList) 'Ajusta a previsão do pedido caso um adiantamento avulso seja associado
            If Me.TituloOriginal IsNot Nothing AndAlso Me.TituloOriginal.ContaContabilCliFor.Adiantamento AndAlso Me.TituloOriginal.CodigoPedido <> Me.CodigoPedido Then
                Dim valorAdiantamento = Me.Valores.EncargoValorDocumento.Valor
                For Each tit In Me.Pedido.Titulos.Where(Function(t) t.CodigoProvisao = 2)
                    valorAdiantamento -= tit.Valores.EncargoValorDocumento.Valor
                    If valorAdiantamento < 0 Then
                        tit.Valores.EncargoValorDocumento.Valor = valorAdiantamento * (-1)
                        tit.IUD = "U"
                    Else
                        tit.IUD = "C"
                    End If
                    tit.SalvarSql(Sqls)
                    If valorAdiantamento <= 0 Then
                        Me.CodigoTituloOrigem = tit.Codigo
                        Exit Sub
                    End If
                Next
            End If
        End Sub

        Private Sub AjustaPrevisaoCancelamentoAdiantamento(ByRef Sqls As ArrayList)
            Dim tituloOrigem As New TituloNovo(Me.CodigoTituloOrigem)
            tituloOrigem.IUD = "U"
            If tituloOrigem.CodigoSituacao = 1 Then
                tituloOrigem.Valores.EncargoValorDocumento.Valor += Me.Valores.EncargoValorDocumento.Valor
            Else
                tituloOrigem.CodigoSituacao = 1
                tituloOrigem.Valores.EncargoValorDocumento.Valor = Me.Valores.EncargoValorDocumento.Valor
            End If
            tituloOrigem.SalvarSql(Sqls)
        End Sub

        '********** Controles das Propriedades *****************
        Public Sub DesligarControles()
            _Controlando = False
        End Sub

        Public Sub LigarControles()
            _Controlando = True
        End Sub

        Public Sub getSqlException(ByRef Sqls As ArrayList)
            If Me.Pedido IsNot Nothing Then
                Dim sql As String = "BEGIN TRY " & vbCrLf & _
                   "DECLARE @HORA_BLOQUEIO AS DATETIME = DATEADD(MINUTE, 3, (SELECT VERSAOHORARIOBLOQUEIO FROM PEDIDOS WHERE PEDIDO_ID = '" & Me.Pedido.Codigo & "' AND EMPRESA_ID = '" & Me.Pedido.CodigoEmpresa & "' AND ENDEMPRESA_ID = '" & Me.Pedido.EnderecoEmpresa & "')) " & vbCrLf & _
                   "PRINT 'HORA_ATUAL: ' + CAST(GETDATE() AS VARCHAR); " & vbCrLf & _
                   "PRINT 'HORA_BLOQUEIO: ' + CAST(@HORA_BLOQUEIO AS VARCHAR); " & vbCrLf & _
                   "IF (GETDATE() > @HORA_BLOQUEIO) " & vbCrLf & _
                   "BEGIN " & vbCrLf & _
                   "RAISERROR ('POR FAVOR, ATUALIZE O SALDO FINANCEIRO PARA REALIZAR ESTA AÇÃO!', " & vbCrLf & _
                   "16, " & vbCrLf & _
                   "1); " & vbCrLf & _
                   "END " & vbCrLf & _
                   "UPDATE PEDIDOS SET VERSAOHORARIOBLOQUEIO = NULL WHERE PEDIDO_ID = '" & Me.Pedido.Codigo & "' AND EMPRESA_ID = '" & Me.Pedido.CodigoEmpresa & "' AND ENDEMPRESA_ID = '" & Me.Pedido.EnderecoEmpresa & "'; " & vbCrLf & _
                   "END TRY " & vbCrLf & _
                   "BEGIN CATCH " & vbCrLf & _
                   "DECLARE @ErrorMessage NVARCHAR(4000); " & vbCrLf & _
                   "DECLARE @ErrorSeverity INT; " & vbCrLf & _
                   "DECLARE @ErrorState INT; " & vbCrLf & _
                   "SELECT " & vbCrLf & _
                   "@ErrorMessage = ERROR_MESSAGE(), " & vbCrLf & _
                   "@ErrorSeverity = ERROR_SEVERITY(), " & vbCrLf & _
                   "@ErrorState = ERROR_STATE(); " & vbCrLf & _
                   "RAISERROR (@ErrorMessage, " & vbCrLf & _
                   "@ErrorSeverity, " & vbCrLf & _
                   "@ErrorState); " & vbCrLf & _
                   "END CATCH;"
                Sqls.Add(sql)
            End If
        End Sub

        '*************** Histórico******************************
        Public Sub HistoricoTitulo(ByRef Sqls As ArrayList)
            Dim Historico As New TituloXHistorico()
            Select Case Me.IUD
                Case "I"
                    Historico.Acao = "INCLUIR"
                Case "U"
                    Historico.Acao = IIf(Me.CodigoProvisao = eProvisao.Baixa, "BAIXAR", "ALTERAR")
                Case "D"
                    Historico.Acao = "EXCLUIR"
                Case "C"
                    Historico.Acao = "CANCELAR"
            End Select
            Historico.IUD = Me.IUD
            Historico.CodigoTitulo = Me.Codigo
            Historico.Data = DateTime.Now
            Historico.Usuario = UsuarioServidor.NomeUsuario
            Historico.SalvarSql(Sqls)
        End Sub
#End Region

    End Class

End Namespace