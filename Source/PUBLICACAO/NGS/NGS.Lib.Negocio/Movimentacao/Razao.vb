Imports System.Web
Imports System.Data
Imports System.Collections.Generic
Imports Microsoft.VisualBasic
Imports NGS.Lib.Uteis

<Serializable()> _
Public Class ListRazao
    Inherits List(Of Razao)

#Region "Contrutor"
    Public Sub New()

    End Sub

    Public Sub New(ByVal pNF As NotaFiscal, Optional NotaEFinanceiro As Boolean = False)
        _NF = pNF
        CarregarContalizacao(NotaEFinanceiro)
    End Sub

    Public Sub New(ByVal pNFxI As NotaFiscalXItem)
        _NFxI = pNFxI
        _NF = pNFxI.NotaFiscal
        CarregarContalizacao()
    End Sub

    Public Sub New(ByVal pNFxE As NotaFiscalXItemXEncargo)
        _NFxE = pNFxE
        _NFxI = pNFxE.ItemNotaFiscal
        _NF = pNFxE.ItemNotaFiscal.NotaFiscal
        CarregarContalizacao()
    End Sub

    Public Sub New(ByVal pTitulo As Novo.TituloNovo)
        _Titulo = pTitulo
        CarregarContalizacao()
    End Sub

    Public Sub New(ByVal pTituloAtual As Titulo)
        _TituloAtual = pTituloAtual
        CarregarContalizacao()
    End Sub
#End Region

#Region "Fields"
    Private _NF As NotaFiscal
    Private _NFxI As NotaFiscalXItem
    Private _NFxE As NotaFiscalXItemXEncargo
    Private _Titulo As Novo.TituloNovo
    Private _TituloAtual As Titulo
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

    Public Property NFxI() As NotaFiscalXItem
        Get
            Return _NFxI
        End Get
        Set(ByVal value As NotaFiscalXItem)
            _NFxI = value
        End Set
    End Property

    Public Property NFxE() As NotaFiscalXItemXEncargo
        Get
            Return _NFxE
        End Get
        Set(ByVal value As NotaFiscalXItemXEncargo)
            _NFxE = value
        End Set
    End Property

    Public Property Titulo() As Novo.TituloNovo
        Get
            Return _Titulo
        End Get
        Set(ByVal value As Novo.TituloNovo)
            _Titulo = value
        End Set
    End Property

    Public Property TituloAtual() As Titulo
        Get
            Return _TituloAtual
        End Get
        Set(ByVal value As Titulo)
            _TituloAtual = value
        End Set
    End Property

    Public Sub CalcularSaldo()
        Dim SaldoRazao As Decimal = 0
        For Each Razao In Me
            SaldoRazao = SaldoRazao + (Razao.DebitoOficial - Razao.CreditoOficial)
            Razao.Saldo = SaldoRazao
        Next
    End Sub

#End Region

#Region "Methods"
    Public Function Salvar() As Boolean
        Dim Banco As New AcessaBanco
        Dim sqls As New ArrayList

        If Not Titulo Is Nothing Then

        Else
            Me.SalvarSqlNota(sqls)
        End If


        If sqls.Count = 0 OrElse Banco.GravaBanco(sqls) Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Sub SalvarSqlNota(ByRef Sqls As ArrayList)
        Dim Proc As String
        Proc = "Begin Transaction " & vbCrLf & _
               " DELETE RAZAO" & vbCrLf & _
               "  WHERE Empresa_Id      ='" & NF.CodigoEmpresa & "'" & vbCrLf & _
               "    And EndEmpresa_Id   = " & NF.EnderecoEmpresa & vbCrLf & _
               "    And Cliente_NF      ='" & NF.CodigoCliente & "'" & vbCrLf & _
               "    And EndCliente_NF   = " & NF.EnderecoCliente & vbCrLf & _
               "    And EntradaSaida_NF ='" & IIf(NF.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "'" & vbCrLf & _
               "    And Serie_NF        ='" & NF.Serie & "'" & vbCrLf & _
               "    And Numero_NF       = " & NF.Numero & vbCrLf

        If Not NFxI Is Nothing Then
            Proc &= "    And Produto_NF   ='" & NFxI.CodigoProduto & "'" & vbCrLf & _
                    "    And CFOP_NF      = " & NFxI.CFOP & vbCrLf & _
                    "    And Sequencia_NF = " & NFxI.Sequencia & vbCrLf
        End If

        If Not NFxE Is Nothing Then
            Proc &= "    And Encargo_NF   ='" & NFxE.Codigo & "'" & vbCrLf
        End If

        Proc &= " declare" & vbCrLf & _
                " @Sequencia int" & vbCrLf & _
                " Set @sequencia = (Select Sequencia" & vbCrLf & _
                "                     from numerador" & vbCrLf & _
                "                    where UPPER(Empresa_id) = '" & HttpContext.Current.Session("ssNomeServidor").ToUpper() & "'" & vbCrLf & _
                "                      and Numerador_id      = 60)" & vbCrLf & _
                " select @Sequencia = isnull(Max(Sequencia_Id),isnull(@sequencia,0))" & vbCrLf & _
                "   from razao (nolock)   " & vbCrLf & _
                "  Where lote_Id      = " & Me(0).Lote & vbCrLf & _
                "    and Movimento_Id ='" & CDate(NF.Movimento).ToString("yyyy/MM/dd") & "'" & vbCrLf & _
                "    and Sequencia_id between isnull(@sequencia,0) and isnull(@sequencia,0) + 99999" & vbCrLf

        Dim sql As String
        Dim i As Integer
        For Each row As Razao In Me
            i += 1
            row.Sequencia = -1
            sql = row.SalvarSql
            sql.Replace("-1", "@Sequencia + " & i)
            Proc &= sql & vbCrLf
        Next
        Proc &= " Commit Transaction "
        Sqls.Add(Proc)
    End Sub

    Public Sub SalvarSqlTitulo(ByRef Sqls As ArrayList)
        Dim sql As String
        sql = " DELETE RAZAO" & vbCrLf & _
              "  WHERE Titulo  in (select Titulo_Id from Titulos where RegistroMestre = " & Titulo.Codigo & " or Titulo_Id = " & Titulo.Codigo & ")"


        If (Titulo.CodigoProvisao = eProvisao.Provisao OrElse Titulo.CodigoProvisao = eProvisao.Compensado) OrElse ((Titulo.IUD = "D" OrElse Titulo.IUD = "C") OrElse Not (Titulo.CodigoProvisao = eProvisao.Baixa) AndAlso (Titulo.ContaContabilCliFor.Adiantamento AndAlso Titulo.IUD <> "I")) Then
            Sqls.Add(sql)
            Exit Sub
        End If
        'Titulo contábil deleta o razão sempre na alterção
        If Titulo.IUD = "U" Then 'AndAlso Titulo.ReceberPagar = "C") OrElse (Titulo.ContaContabilCliFor.Adiantamento AndAlso Not Titulo.Equals("I")) Then
            Sqls.Add(sql)
        End If

        Dim seq As Integer
        For Each row As Razao In Me
            Sqls.Add(row.SalvarSql(seq))
            seq += 1
        Next
    End Sub

    Private Sub CarregarContalizacao(Optional NotaEFinanceiro As Boolean = False)
        Dim sql As String
        sql = "SELECT Empresa_Id, EndEmpresa_Id, Conta_Id, Cliente_Id, EndCliente_Id, Movimento_Id, Lote_Id, Sequencia_Id," & vbCrLf & _
              "       UnidadeDeNegocio, isnull(Titulo,0) as Titulo, isnull(Pedido,0) as Pedido, isnull(PedidoFixacao,0) as PedidoFixacao, " & vbCrLf & _
              "       isnull(Produto,'') as Produto, Custo, Indexador, DataMoeda, DebitoOficial, CreditoOficial, DebitoMoeda, CreditoMoeda, Historico, PrevistoRealizado," & vbCrLf & _
              "       isnull(Cliente_Nf,'') as Cliente_nf, isnull(EndCliente_Nf,0) as EndCliente_Nf, isnull(EntradaSaida_Nf,'') as EntradaSaida_Nf, isnull(Serie_Nf,'') as Serie_Nf, isnull(Numero_Nf,0) as Numero_Nf, isnull(Produto_NF,'') as Produto_NF, isnull(Cfop_NF,0) as Cfop_NF, isnull(Sequencia_NF,0) as Sequencia_NF, isnull(Encargo_NF,'') as Encargo_NF," & vbCrLf & _
              "       isnull(ChequeEntregue,'') as ChequeEntregue,  isnull(PagamentoAutorizado,'') as PagamentoAutorizado, isnull(DataDaBaixa,movimento_id) as DataDaBaixa, isnull(Conciliacao,'') as Conciliacao,  isnull(Deposito,'') as Deposito, isnull(EndDeposito,0) as EndDeposito, isnull(Rateado,0) as Rateado, isnull(Processo,'') as Processo," & vbCrLf & _
              "       isnull(DebitoQuantidade,0) as DebitoQuantidade, isnull(CreditoQuantidade,0) as CreditoQuantidade, isnull(Situacao,1) as Situacao" & vbCrLf & _
              "  FROM Razao" & vbCrLf
        If Not NFxE Is Nothing Then
            sql &= "  Where Empresa_id             ='" & _NFxE.ItemNotaFiscal.NotaFiscal.CodigoEmpresa & "'" & vbCrLf & _
                   "    and EndEmpresa_id          = " & _NFxE.ItemNotaFiscal.NotaFiscal.EnderecoEmpresa & vbCrLf & _
                   "    and Cliente_NF             ='" & _NFxE.ItemNotaFiscal.NotaFiscal.CodigoCliente & "'" & vbCrLf & _
                   "    and EndCliente_NF          = " & _NFxE.ItemNotaFiscal.NotaFiscal.EnderecoCliente & vbCrLf & _
                   "    and EntradaSaida_NF        ='" & IIf(_NFxE.ItemNotaFiscal.NotaFiscal.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "'" & vbCrLf & _
                   "    and Serie_NF               ='" & _NFxE.ItemNotaFiscal.NotaFiscal.Serie & "'" & vbCrLf & _
                   "    and Numero_NF              = " & _NFxE.ItemNotaFiscal.NotaFiscal.Codigo & _
                   "    and Produto_NF             ='" & _NFxE.ItemNotaFiscal.CodigoProduto & "'" & vbCrLf & _
                   "    and isnull(Sequencia_NF,0) = " & _NFxE.ItemNotaFiscal.Sequencia & vbCrLf & _
                   "    and Encargo_NF             ='" & _NFxE.Codigo & "'"
        ElseIf Not NFxI Is Nothing Then
            sql &= "  Where Empresa_id             ='" & _NFxI.NotaFiscal.CodigoEmpresa & "'" & vbCrLf & _
                   "    and EndEmpresa_id          = " & _NFxI.NotaFiscal.EnderecoEmpresa & vbCrLf & _
                   "    and Cliente_NF             ='" & _NFxI.NotaFiscal.CodigoCliente & "'" & vbCrLf & _
                   "    and EndCliente_NF          = " & _NFxI.NotaFiscal.EnderecoCliente & vbCrLf & _
                   "    and EntradaSaida_NF        ='" & IIf(_NFxI.NotaFiscal.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "'" & vbCrLf & _
                   "    and Serie_NF               ='" & _NFxI.NotaFiscal.Serie & "'" & vbCrLf & _
                   "    and Numero_NF              = " & _NFxI.NotaFiscal.Codigo & _
                   "    and Produto_NF             ='" & _NFxI.CodigoProduto & "'" & vbCrLf & _
                   "    and isnull(Sequencia_NF,0) = " & _NFxI.Sequencia & vbCrLf
        ElseIf Not NF Is Nothing Then
            sql &= "  Where (" & vbCrLf & _
                   "         Empresa_id      ='" & _NF.CodigoEmpresa & "'" & vbCrLf & _
                   "     and EndEmpresa_id   = " & _NF.EnderecoEmpresa & vbCrLf & _
                   "     and Cliente_NF      ='" & _NF.CodigoCliente & "'" & vbCrLf & _
                   "     and EndCliente_NF   = " & _NF.EnderecoCliente & vbCrLf & _
                   "     and EntradaSaida_NF ='" & IIf(_NF.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "'" & vbCrLf & _
                   "     and Serie_NF        ='" & _NF.Serie & "'" & vbCrLf & _
                   "     and Numero_NF       = " & _NF.Codigo & vbCrLf & _
                   "        )" & vbCrLf
            If NotaEFinanceiro Then
                sql &= "      or titulo in (SELECT Titulo_Id" & vbCrLf & _
                       "                     FROM NotaFiscalXTitulo" & vbCrLf & _
                       "                    where Empresa_Id       ='" & _NF.CodigoEmpresa & "'" & vbCrLf & _
                       "                      and EndEmpresa_Id    = " & _NF.EnderecoEmpresa & vbCrLf & _
                       "                      and Cliente_Id       ='" & _NF.CodigoCliente & "'" & vbCrLf & _
                       "                      and EndCliente_Id    = " & _NF.EnderecoCliente & vbCrLf & _
                       "                      and EntradaSaida_Id  ='" & IIf(_NF.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "'" & vbCrLf & _
                       "                      and Serie_Id         ='" & _NF.Serie & "'" & vbCrLf & _
                       "                      and Nota_Id          = " & _NF.Codigo & vbCrLf & _
                       "                     )"
            End If

        ElseIf Not TituloAtual Is Nothing Then
            sql &= "  Where Titulo = " & TituloAtual.Codigo & vbCrLf
        ElseIf Not Titulo Is Nothing Then
            sql &= "  Where Empresa_id             ='" & _Titulo.CodigoEmpresa & "'" & vbCrLf & _
                   "    and EndEmpresa_id          = " & _Titulo.EnderecoEmpresa & vbCrLf & _
                   "    and Titulo                 <> 0" & vbCrLf
            If Titulo.RegistroMestre = 0 Then
                sql &= "    and Titulo                 = " & _Titulo.Codigo & vbCrLf
            ElseIf Titulo.Codigo = Titulo.RegistroMestre Then
                sql &= "    and (Titulo = " & _Titulo.Codigo & vbCrLf & _
                       "             OR" & vbCrLf & _
                       "         Titulo in (Select titulo_Id from Titulos Where RegistroMestre = " & _Titulo.Codigo & ")" & vbCrLf & _
                       "        )"
            ElseIf Titulo.RegistroMestre > 0 Then
                sql &= "    and (Titulo = " & _Titulo.RegistroMestre & vbCrLf & _
                       "             OR" & vbCrLf & _
                       "         Titulo in (Select Titulo_Id from Titulos Where RegistroMestre = " & _Titulo.RegistroMestre & ")" & vbCrLf & _
                       "        )"
            End If
        End If

        sql &= "  Order by Razao.Lote_Id," & vbCrLf & _
               "           case when isnull(Encargo_NF,'') = 'PRODUTO' then 1" & vbCrLf & _
               "                when isnull(Encargo_NF,'') = 'LIQUIDO' then 3" & vbCrLf & _
               "                else 2" & vbCrLf & _
               "           end," & vbCrLf & _
               "           Razao.Conta_Id " & vbCrLf

        Dim Banco As New AcessaBanco
        Dim ds As DataSet = Banco.ConsultaDataSet(sql, "Contabilizacao")

        For Each row As DataRow In ds.Tables(0).Rows
            Dim R As Razao = Nothing

            If Not _NFxE Is Nothing Then
                R = New Razao(_NFxE)
            ElseIf Not _NFxI Is Nothing Then
                R = New Razao(_NFxI)
            ElseIf Not _NF Is Nothing Then
                R = New Razao(_NF)
            ElseIf Not _TituloAtual Is Nothing Then
                R = New Razao(_TituloAtual)
            ElseIf Not _Titulo Is Nothing Then
                R = New Razao(_Titulo)
            End If

            R.IUD = "U"
            R.CodigoEmpresa = row("Empresa_Id")
            R.EndEmpresa = row("EndEmpresa_Id")
            R.CodigoConta = row("Conta_Id")
            R.CodigoCliente = row("Cliente_Id")
            R.EnderecoCliente = row("EndCliente_Id")
            R.Movimento = row("Movimento_Id")
            R.Lote = row("Lote_Id")
            R.Sequencia = row("Sequencia_Id")
            R.CodigoUnidadeDeNegocio = row("UnidadeDeNegocio")
            R.CodigoTitulo = row("Titulo")
            R.CodigoPedido = row("Pedido")
            R.codigoPedidoFixacao = row("PedidoFixacao")
            R.CodigoProduto = row("Produto")
            R.CodigoCusto = row("Custo")
            R.CodigoIndexador = row("Indexador")
            R.DataMoeda = row("DataMoeda")
            R.DebitoOficial = row("DebitoOficial")
            R.CreditoOficial = row("CreditoOficial")
            R.DebitoMoeda = row("DebitoMoeda")
            R.CreditoMoeda = row("CreditoMoeda")
            R.Historico = row("Historico")
            R.PrevistoRealizado = row("PrevistoRealizado")
            R.CodigoClienteNF = row("Cliente_NF")
            R.EnderecoClienteNF = row("EndCliente_NF")
            R.EntradaSaidaNF = row("EntradaSaida_NF")
            R.SerieNF = row("Serie_NF")
            R.NumeroNF = row("Numero_NF")
            R.CodigoProdutoNF = row("Produto_NF")
            R.CodigoCFOPNF = row("Cfop_NF")
            R.SequenciaNF = row("Sequencia_NF")
            R.CodigoEncargoNF = row("Encargo_NF")
            R.ChequeEntregue = row("ChequeEntregue")
            R.PagamentoAutorizado = row("PagamentoAutorizado")
            R.DataBaixa = row("DataDaBaixa")
            R.Conciliacao = row("Conciliacao")
            R.CodigoDeposito = row("Deposito")
            R.EndDeposito = row("EndDeposito")
            R.Rateado = row("Rateado")
            R.Processo = row("Processo")
            R.DebitoQuantidade = row("DebitoQuantidade")
            R.CreditoQuantidade = row("CreditoQuantidade")
            R.CodigoSituacao = row("Situacao")
            Me.Add(R)
        Next
    End Sub

    Public Sub ApurarContabilizacaoNota(Optional ByVal pLote As Integer = 0)
        If Not NF Is Nothing AndAlso Not NF.SubOperacao.Contabil Then Exit Sub

        Dim Indice As Decimal
        Dim DataIndice As Date
        Dim CodIndexador As Integer

        If NF.Itens(0).CodigoFixacao > 0 Then
            For Each f As Fixacao In NF.Pedido.Itens(0).Fixacoes
                If f.Codigo = NF.Itens(0).CodigoFixacao Then
                    Indice = f.IndiceFixado
                    DataIndice = f.Movimento
                    CodIndexador = 99
                    Exit For
                End If
            Next
        ElseIf NF.Pedido.IndiceFixado > 0 Then
            Indice = NF.Pedido.IndiceFixado
            DataIndice = NF.Pedido.DataPedido
            CodIndexador = 99
        Else
            Indice = New Cotacao(NF.Pedido.CodigoIndexador, NF.Movimento).Indice
            DataIndice = NF.Movimento
            CodIndexador = NF.Pedido.CodigoIndexador
        End If


        For Each Item As NotaFiscalXItem In NF.Itens
            For Each Enc As NotaFiscalXItemXEncargo In Item.Encargos
                '********************************************************************************
                '*******************************  LIQUIDO ***************************************
                '********************************************************************************
                If Enc.Codigo = "LIQUIDO" _
                And Enc.OperacaoEncargo.CodigoDebitaConta.Trim = "" _
                And Enc.OperacaoEncargo.CodigoCreditaConta.Trim = "" _
                And Enc.Valor > 0 Then
                    Dim R As New Razao(Enc)

                    R.CodigoConta = NF.SubOperacao.CodigoGrupoContas
                    R.CodigoCliente = NF.CodigoCliente
                    R.EnderecoCliente = NF.EnderecoCliente
                    R.Movimento = NF.Movimento

                    If pLote > 0 Then
                        R.Lote = pLote
                    ElseIf Item.CodigoOperacao > 68 Then
                        R.Lote = 9
                    Else
                        R.Lote = 10
                    End If

                    R.CodigoUnidadeDeNegocio = NF.Pedido.CodigoUnidadeNegocio
                    R.CodigoIndexador = CodIndexador
                    R.DataMoeda = DataIndice

                    If NF.EntradaSaida.ToString.Substring(0, 1) = "S" Then
                        R.DebitoOficial = Enc.Valor
                        R.DebitoMoeda = IIf(Indice < 1, Enc.Valor * Indice, Enc.Valor / Indice)
                    Else
                        R.CreditoOficial = Enc.Valor
                        R.CreditoMoeda = IIf(Indice < 1, Enc.Valor * Indice, Enc.Valor / Indice)
                    End If

                    R.Historico = "NF " & NF.Codigo & "-" & NF.Serie & " OP " & NF.CodigoOperacao & "-" & NF.CodigoSubOperacao & " Pedido " & NF.CodigoPedido & " " & NF.Cliente.Nome
                    R.CodigoPedido = NF.CodigoPedido
                    R.codigoPedidoFixacao = Enc.ItemNotaFiscal.CodigoFixacao
                    R.PrevistoRealizado = "R"
                    R.CodigoCusto = Enc.CentroDeCusto
                    If Not Item.SubOperacao.GrupoDeConta Is Nothing And Item.SubOperacao.GrupoDeConta.TemProduto Then
                        R.CodigoProduto = Item.CodigoProduto
                    End If
                    R.Rateado = Item.Rateado
                    Me.Add(R)
                Else
                    '********************************************************************************
                    '*******************************  DEBITO  ***************************************
                    '********************************************************************************
                    If Enc.OperacaoEncargo.CodigoDebitaConta.Trim <> "" _
                    And Enc.Valor > 0 Then
                        Dim R As New Razao(Enc)
                        R.CodigoConta = Enc.OperacaoEncargo.CodigoDebitaConta
                        R.CodigoCliente = IIf(R.Conta.TemCliente, NF.CodigoCliente, "")
                        R.EnderecoCliente = IIf(R.Conta.TemCliente, NF.EnderecoCliente, 0)
                        R.Movimento = NF.Movimento

                        If pLote > 0 Then
                            R.Lote = pLote
                        ElseIf Item.CodigoOperacao > 68 Then
                            R.Lote = 9
                        Else
                            R.Lote = 10
                        End If

                        R.CodigoUnidadeDeNegocio = NF.Pedido.CodigoUnidadeNegocio
                        R.CodigoIndexador = CodIndexador
                        R.DataMoeda = DataIndice

                        R.DebitoOficial = Enc.Valor
                        R.DebitoMoeda = IIf(Indice < 1, Enc.Valor * Indice, Enc.Valor / Indice)

                        R.Historico = Enc.Codigo & " OP " & NF.CodigoOperacao & "-" & NF.CodigoSubOperacao & " NF " & NF.Codigo & "-" & NF.Serie & " Pedido " & NF.CodigoPedido & " " & NF.Cliente.Nome
                        R.CodigoPedido = NF.CodigoPedido
                        R.PrevistoRealizado = "R"
                        R.CodigoCusto = Enc.CentroDeCusto
                        If Enc.OperacaoEncargo.CodigoDebitaConta.Length > 0 AndAlso Enc.OperacaoEncargo.DebitaConta.TemProduto Then
                            R.CodigoProduto = Item.CodigoProduto
                        End If
                        R.Rateado = Item.Rateado
                        Me.Add(R)
                    End If

                    '*********************************************************************************
                    '*******************************  CREDITO  ***************************************
                    '*********************************************************************************
                    If Enc.OperacaoEncargo.CodigoCreditaConta.Trim <> "" And Enc.Valor > 0 Then
                        Dim R As New Razao(Enc)
                        R.CodigoConta = Enc.OperacaoEncargo.CodigoCreditaConta
                        R.CodigoCliente = IIf(R.Conta.TemCliente, NF.CodigoCliente, "")
                        R.EnderecoCliente = IIf(R.Conta.TemCliente, NF.EnderecoCliente, 0)
                        R.Movimento = NF.Movimento

                        If pLote > 0 Then
                            R.Lote = pLote
                        ElseIf Item.CodigoOperacao > 68 Then
                            R.Lote = 9
                        Else
                            R.Lote = 10
                        End If

                        R.CodigoUnidadeDeNegocio = NF.Pedido.CodigoUnidadeNegocio
                        R.CodigoIndexador = CodIndexador
                        R.DataMoeda = DataIndice

                        R.CreditoOficial = Enc.Valor
                        R.CreditoMoeda = IIf(Indice < 1, Enc.Valor * Indice, Enc.Valor / Indice)

                        R.Historico = Enc.Codigo & " OP " & NF.CodigoOperacao & "-" & NF.CodigoSubOperacao & " NF " & NF.Codigo & "-" & NF.Serie & " Pedido " & NF.CodigoPedido & " - " & NF.Cliente.Nome
                        R.CodigoPedido = NF.CodigoPedido
                        R.PrevistoRealizado = "R"                                                          'Previsto/Realizado
                        R.CodigoCusto = Enc.CentroDeCusto
                        If Enc.OperacaoEncargo.CodigoCreditaConta.Length > 0 AndAlso Enc.OperacaoEncargo.CreditaConta.TemProduto Then
                            R.CodigoProduto = Item.CodigoProduto
                        End If
                        R.Rateado = Item.Rateado
                        Me.Add(R)
                    End If
                End If
            Next
        Next

    End Sub

    Public Sub ApurarContabilizacaoTitulo()
        If Titulo.RegistroMestre = 0 Then
            ApuraContabilizacaoTitulosSemAgrupamento()
        Else
            ApuraContabilizacaoTitulosComAgrupamento()
        End If
        ApurarTransferenciaContabilizacaoTitulo()
    End Sub

    Private Sub ApuraContabilizacaoTitulosSemAgrupamento()
        Me.Clear()

        For Each Enc As Novo.TituloXContaContabil In Titulo.Valores
            Dim R As New Razao(Titulo)
            R.IUD = "I"
            R.CodigoTitulo = Titulo.Codigo

            R.CodigoUnidadeDeNegocio = Titulo.CodigoUnidadeDeNegocio
            If Titulo.Valores.EncargoValorLiquido.CodigoContaEncargo = Enc.CodigoContaEncargo Then
                R.CodigoEmpresa = Titulo.CodigoEmpresaRecPag
                R.EndEmpresa = Titulo.EndEmpresaRecPag
            Else
                R.CodigoEmpresa = Titulo.CodigoEmpresa
                R.EndEmpresa = Titulo.EnderecoEmpresa
            End If

            R.CodigoConta = Enc.CodigoContaEncargo

            If Titulo.ReceberPagar = "C" Then
                If Titulo.Valores.EncargoValorDocumento.CodigoContaEncargo = Enc.CodigoContaEncargo Then
                    R.CodigoCliente = IIf(Enc.ContaEncargo.TemCliente, Titulo.CodigoCliFor, "")
                    R.EnderecoCliente = IIf(Enc.ContaEncargo.TemCliente, Titulo.EnderecoCliFor, 0)

                    R.CodigoPedido = Titulo.CodigoPedido
                    R.CodigoProduto = Titulo.CodigoProduto
                    R.DebitoQuantidade = Str(Titulo.Quantidade)

                ElseIf Titulo.Valores.EncargoValorLiquido.CodigoContaEncargo = Enc.CodigoContaEncargo Then
                    R.CodigoCliente = IIf(Enc.ContaEncargo.TemCliente, Titulo.CodigoClienteRecPag, "")
                    R.EnderecoCliente = IIf(Enc.ContaEncargo.TemCliente, Titulo.EndClienteRecPag, 0)

                    R.CodigoPedido = Titulo.CodigoPedidoRecPag
                    R.CodigoProduto = Titulo.CodigoProduto
                    R.CreditoQuantidade = Str(Titulo.Quantidade)
                Else

                    If Enc.DC = "D" Then
                        R.CodigoPedido = Titulo.CodigoPedido
                        R.DebitoQuantidade = Titulo.Quantidade
                    Else
                        R.CodigoPedido = Titulo.CodigoPedidoRecPag
                        R.CreditoQuantidade = Str(Titulo.Quantidade)
                    End If
                End If
            Else
                R.CodigoCliente = IIf(Enc.ContaEncargo.TemCliente, Titulo.CodigoCliFor, "")
                R.EnderecoCliente = IIf(Enc.ContaEncargo.TemCliente, Titulo.EnderecoCliFor, 0)
            End If

            R.Movimento = Titulo.DataBaixa
            R.Lote = Titulo.Lote
            If Titulo.Lote = 70 Then Titulo.Sequencia = Titulo.Codigo
            R.Sequencia = Titulo.Sequencia

            R.CodigoIndexador = Titulo.CodigoIndexador
            R.DataMoeda = Titulo.DataMoeda
            R.DataBaixa = Titulo.DataBaixa

            R.Processo = "BAIXA TITULO"
            R.Historico = Titulo.Historico

            R.PrevistoRealizado = "R"

            If Enc.DC = "C" Then
                R.CreditoOficial = Enc.ValorOficial
                R.CreditoMoeda = Enc.ValorMoeda
                R.DebitoOficial = 0
                R.DebitoMoeda = 0
            ElseIf Enc.DC = "D" Then
                R.CreditoOficial = 0
                R.CreditoMoeda = 0
                R.DebitoOficial = Enc.ValorOficial
                R.DebitoMoeda = Enc.ValorMoeda
            End If

            If Enc.ValorOficial > 0 Or Enc.ValorMoeda > 0 Then Me.Add(R)
        Next
    End Sub

    Private Sub ApuraContabilizacaoTitulosComAgrupamento()
        Me.Clear()
        '********************************************************************************************
        '*********************** Faz a Parte do Mestre ContaContabilRecPag **************************
        '********************************************************************************************
        Dim Enc As Novo.TituloXContaContabil = Titulo.Valores.EncargoValorLiquido
        Dim R As New Razao(Titulo)
        R.IUD = "I"
        R.CodigoTitulo = Titulo.Codigo

        R.CodigoUnidadeDeNegocio = Titulo.CodigoUnidadeDeNegocio
        R.CodigoEmpresa = Titulo.CodigoEmpresaRecPag
        R.EndEmpresa = Titulo.EndEmpresaRecPag

        R.CodigoConta = Enc.CodigoContaEncargo

        R.CodigoCliente = IIf(Enc.ContaEncargo.TemCliente, Titulo.CodigoCliFor, "")
        R.EnderecoCliente = IIf(Enc.ContaEncargo.TemCliente, Titulo.EnderecoCliFor, 0)

        R.Movimento = Titulo.Movimento
        R.Lote = 70
        R.Sequencia = Titulo.Codigo

        R.CodigoIndexador = Titulo.CodigoIndexador
        R.DataMoeda = Titulo.DataMoeda
        R.DataBaixa = Titulo.DataBaixa

        R.Processo = "BAIXA TITULO"
        R.Historico = Titulo.Historico

        R.PrevistoRealizado = "R"

        If Enc.DC = "C" Then
            R.CreditoOficial = Enc.ValorOficial
            R.CreditoMoeda = Enc.ValorMoeda
            R.DebitoOficial = 0
            R.DebitoMoeda = 0
        ElseIf Enc.DC = "D" Then
            R.CreditoOficial = 0
            R.CreditoMoeda = 0
            R.DebitoOficial = Enc.ValorOficial
            R.DebitoMoeda = Enc.ValorMoeda
        End If

        Me.Add(R)

        '***********************************************************************************************************************************
        '*********************** Faz a Parte dos filhos ContaContabilCliFor e demais contas - ContaContabilRecPag **************************
        '***********************************************************************************************************************************
        For Each Filho In Titulo.TitulosAgrupados
            For Each EncFilho As Novo.TituloXContaContabil In Filho.Valores
                If Not EncFilho.Equals(Filho.Valores.EncargoValorLiquido) And (EncFilho.ValorOficial > 0 Or EncFilho.ValorMoeda > 0) Then
                    Dim RFilho As New Razao(Filho)
                    RFilho.IUD = "I"
                    RFilho.CodigoTitulo = Filho.Codigo

                    RFilho.CodigoUnidadeDeNegocio = Filho.CodigoUnidadeDeNegocio

                    RFilho.CodigoEmpresa = Filho.CodigoEmpresa
                    RFilho.EndEmpresa = Filho.EnderecoEmpresa

                    RFilho.CodigoConta = EncFilho.CodigoContaEncargo

                    RFilho.CodigoCliente = IIf(EncFilho.ContaEncargo.TemCliente, Filho.CodigoCliFor, "")
                    RFilho.EnderecoCliente = IIf(EncFilho.ContaEncargo.TemCliente, Filho.EnderecoCliFor, 0)

                    RFilho.Movimento = Filho.DataBaixa
                    RFilho.Lote = 70
                    RFilho.Sequencia = Filho.Codigo

                    RFilho.CodigoIndexador = Filho.CodigoIndexador
                    RFilho.DataMoeda = Filho.DataMoeda
                    RFilho.DataBaixa = Filho.DataBaixa

                    RFilho.Processo = "BAIXA TITULO"
                    RFilho.Historico = Filho.Historico

                    RFilho.PrevistoRealizado = "R"

                    If EncFilho.DC = "C" Then
                        RFilho.CreditoOficial = EncFilho.ValorOficial
                        RFilho.CreditoMoeda = EncFilho.ValorMoeda
                        RFilho.DebitoOficial = 0
                        RFilho.DebitoMoeda = 0
                    ElseIf EncFilho.DC = "D" Then
                        RFilho.CreditoOficial = 0
                        RFilho.CreditoMoeda = 0
                        RFilho.DebitoOficial = EncFilho.ValorOficial
                        RFilho.DebitoMoeda = EncFilho.ValorMoeda
                    End If

                    Me.Add(RFilho)
                End If
            Next
        Next
    End Sub

    Private Sub ApurarTransferenciaContabilizacaoTitulo()
        '***************************************************
        '*****  Transferencia Financeira *******************
        '***************************************************
        Dim ListTF As ListTransferenciaFinanceira

        If Titulo.ReceberPagar = "P" Then
            ListTF = New ListTransferenciaFinanceira(Titulo.Empresa, Titulo.EmpresaRecPag) 'Debito/Credito
        Else
            'ListTF = New ListTransferenciaFinanceira(Titulo.EmpresaRecPag, Titulo.Empresa) 'Debito/Credito
            ListTF = New ListTransferenciaFinanceira(Titulo.Empresa, Titulo.EmpresaRecPag) 'Debito/Credito
        End If

        For Each TF As TransferenciaFinanceira In ListTF
            Dim R As New Razao(Titulo)
            R.IUD = "I"
            R.CodigoTitulo = Titulo.Codigo

            R.CodigoEmpresa = TF.CodigoEmpresaContabil
            R.EndEmpresa = TF.EndEmpresaContabil

            R.CodigoConta = TF.ContaContabil

            R.CodigoCliente = TF.CodigoClienteContabil
            R.EnderecoCliente = TF.EndClienteContabil

            R.Movimento = Titulo.DataBaixa
            R.Lote = 70
            R.Sequencia = Titulo.Codigo

            R.CodigoUnidadeDeNegocio = Titulo.CodigoUnidadeDeNegocio

            R.CodigoIndexador = Titulo.CodigoIndexador
            R.DataMoeda = Titulo.DataMoeda
            R.DataBaixa = Titulo.DataBaixa

            R.Processo = "BAIXA TITULO TRANSF. FINANC."
            R.Historico = Titulo.Historico

            R.PrevistoRealizado = "R"

            If (Titulo.ReceberPagar = "P" And TF.DebitoCredito = "C") Or (Titulo.ReceberPagar = "R" And TF.DebitoCredito = "D") Then
                R.CreditoOficial = Titulo.Valores.EncargoValorLiquido.ValorOficial
                R.CreditoMoeda = Titulo.Valores.EncargoValorLiquido.ValorMoeda
                R.DebitoOficial = 0
                R.DebitoMoeda = 0
            ElseIf (Titulo.ReceberPagar = "P" And TF.DebitoCredito = "D") Or (Titulo.ReceberPagar = "R" And TF.DebitoCredito = "C") Then
                R.CreditoOficial = 0
                R.CreditoMoeda = 0
                R.DebitoOficial = Titulo.Valores.EncargoValorLiquido.ValorOficial
                R.DebitoMoeda = Titulo.Valores.EncargoValorLiquido.ValorMoeda
            End If

            Me.Add(R)
        Next
    End Sub

#End Region

End Class

<Serializable()> _
Public Class Razao
    Implements IBaseEntity

#Region "Construtor"
    Public Sub New()

    End Sub

    Public Sub New(ByRef Titulo As Negocio.Titulo)
        _CodigoEmpresa = Titulo.CodigoEmpresa
        _EndEmpresa = Titulo.EnderecoEmpresa
        _CodigoTitulo = Titulo.Codigo
        _Titulo = Titulo
    End Sub

    Public Sub New(ByRef Titulo As Novo.TituloNovo)
        _CodigoEmpresa = Titulo.CodigoEmpresa
        _EndEmpresa = Titulo.EnderecoEmpresa
        _CodigoTituloNovo = Titulo.Codigo
        _TituloNovo = Titulo
    End Sub

    Public Sub New(ByRef Nota As NotaFiscal)
        _NotaFiscalRazao = Nota
        _CodigoEmpresa = Nota.CodigoEmpresa
        _EndEmpresa = Nota.EnderecoEmpresa
        _CodigoClienteNF = Nota.CodigoCliente
        _EnderecoClienteNF = Nota.EnderecoCliente
        _NumeroNF = Nota.Codigo
        _SerieNF = Nota.Serie
        _EntradaSaidaNF = Nota.EntradaSaida.ToString.Substring(0, 1)
        _Pedido = Nota.Pedido
    End Sub

    Public Sub New(ByRef NxI As NotaFiscalXItem)
        _NotaFiscalRazao = NxI.NotaFiscal
        _CodigoEmpresa = NxI.NotaFiscal.CodigoEmpresa
        _EndEmpresa = NxI.NotaFiscal.EnderecoEmpresa
        _CodigoClienteNF = NxI.NotaFiscal.CodigoCliente
        _EnderecoClienteNF = NxI.NotaFiscal.EnderecoCliente
        _NumeroNF = NxI.NotaFiscal.Codigo
        _SerieNF = NxI.NotaFiscal.Serie
        _EntradaSaidaNF = NxI.NotaFiscal.EntradaSaida.ToString.Substring(0, 1)
        _CodigoProdutoNF = NxI.CodigoProduto
        _CodigoCFOPNF = NxI.CFOP
        _SequenciaNF = NxI.Sequencia
    End Sub

    Public Sub New(ByRef NxIxE As NotaFiscalXItemXEncargo)
        _NotaFiscalRazao = NxIxE.ItemNotaFiscal.NotaFiscal
        _CodigoEmpresa = NxIxE.ItemNotaFiscal.NotaFiscal.CodigoEmpresa
        _EndEmpresa = NxIxE.ItemNotaFiscal.NotaFiscal.EnderecoEmpresa
        _CodigoClienteNF = NxIxE.ItemNotaFiscal.NotaFiscal.CodigoCliente
        _EnderecoClienteNF = NxIxE.ItemNotaFiscal.NotaFiscal.EnderecoCliente
        _NumeroNF = NxIxE.ItemNotaFiscal.NotaFiscal.Codigo
        _SerieNF = NxIxE.ItemNotaFiscal.NotaFiscal.Serie
        _EntradaSaidaNF = NxIxE.ItemNotaFiscal.NotaFiscal.EntradaSaida.ToString.Substring(0, 1)
        _CodigoProdutoNF = NxIxE.ItemNotaFiscal.CodigoProduto
        _CodigoCFOPNF = NxIxE.ItemNotaFiscal.CFOP
        _SequenciaNF = NxIxE.ItemNotaFiscal.Sequencia
        _CodigoEncargoNF = NxIxE.Codigo
    End Sub
#End Region

#Region "Fields"
    Private _IUD As String

    Private _CodigoEmpresa As String = ""
    Private _EndEmpresa As Integer
    Private _Empresa As Cliente

    Private _CodigoConta As String = ""
    Private _Conta As PlanoDeConta

    Private _CodigoCliente As String = ""
    Private _EnderecoCliente As Integer
    Private _Cliente As Cliente

    Private _Movimento As Date
    Private _Lote As Integer
    Private _Sequencia As Integer

    Private _CodigoUnidadeDeNegocio As String
    Private _UnidadeDeNegocio As Cliente

    Private _CodigoTitulo As Integer
    Private _Titulo As Negocio.Titulo

    Private _CodigoTituloNovo As Integer
    Private _TituloNovo As Novo.TituloNovo

    Private _CodigoPedido As Integer
    Private _Pedido As Pedido
    Private _CodigoPedidoFixacao As Integer ' Em caso de preco A Fixar é o numero da Fixacao

    Private _CodigoProduto As String = ""
    Private _Produto As Produto

    Private _CodigoCusto As String = ""
    Private _CentroDeCusto As CentroDeCusto

    Private _CodigoIndexador As Integer
    Private _Indexador As Indexador

    Private _DataMoeda As Date

    Private _DebitoOficial As Double
    Private _CreditoOficial As Double
    Private _DebitoMoeda As Double
    Private _CreditoMoeda As Double

    Private _Historico As String
    Private _PrevistoRealizado As String

    '*******************************************************
    '******* Nota Fiscal  **********************************
    '*******************************************************
    Private _ClienteNF As Cliente
    Private _CodigoClienteNF As String = ""
    Private _EnderecoClienteNF As Integer
    Private _EntradaSaidaNF As String
    Private _SerieNF As String = ""
    Private _NumeroNF As Integer
    Private _CodigoProdutoNF As String
    Private _ProdutoNF As Produto
    Private _CodigoCFOPNF As Integer
    Private _CFOPNF As CFOP
    Private _SequenciaNF As Integer
    Private _CodigoEncargoNF As String
    Private _EncargoNF As Encargo
    Private _NotaFiscalRazao As NotaFiscal
    Private _NotaFiscalXItemRazao As NotaFiscalXItem
    Private _NotaFiscalXItemXEncargoRazao As NotaFiscalXItemXEncargo
    '*******************************************************

    Private _ChequeEntregue As String = ""
    Private _PagamentoAutorizado As String = ""
    Private _DataBaixa As Date
    Private _Conciliacao As String = ""

    Private _CodigoDeposito As String = ""
    Private _EndDeposito As Integer
    Private _Deposito As Cliente

    Private _Rateado As Boolean
    Private _Processo As String
    Private _DebitoQuantidade As Decimal
    Private _CreditoQuantidade As Decimal
    Private _CodigoSituacao As Integer
    Private _Situacao As Situacao
    Private _UsuarioInclusao As String
    Private _UsuarioDataInclusao As Date

    Private _Saldo As Decimal

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

    '*************  EMPRESA   ***********************
    Public Property CodigoEmpresa() As String
        Get
            Return _CodigoEmpresa
        End Get
        Set(ByVal value As String)
            _CodigoEmpresa = value
        End Set
    End Property

    Public Property EndEmpresa() As Integer
        Get
            Return _EndEmpresa
        End Get
        Set(ByVal value As Integer)
            _EndEmpresa = value
        End Set
    End Property

    Public Property Empresa() As Cliente
        Get
            If _Empresa Is Nothing And _CodigoEmpresa > 0 Then _Empresa = New Negocio.Cliente(_CodigoEmpresa, _EndEmpresa)
            Return _Empresa
        End Get
        Set(ByVal value As Cliente)
            _Empresa = value
        End Set
    End Property

    '*************  CONTA CONTABIL   ***********************
    Public Property CodigoConta() As String
        Get
            Return _CodigoConta
        End Get
        Set(ByVal value As String)
            _CodigoConta = value
            _Conta = Nothing
        End Set
    End Property

    Public Property Conta() As PlanoDeConta
        Get
            If _Conta Is Nothing And _CodigoConta.Length > 0 Then _Conta = New PlanoDeConta("", 0, _CodigoConta)
            Return _Conta
        End Get
        Set(ByVal value As PlanoDeConta)
            _Conta = value
        End Set
    End Property

    '*************  CLIENTE   ***********************
    Public Property CodigoCliente() As String
        Get
            Return _CodigoCliente
        End Get
        Set(ByVal value As String)
            _CodigoCliente = value
            Cliente = Nothing
        End Set
    End Property

    Public Property EnderecoCliente() As Integer
        Get
            Return _EnderecoCliente
        End Get
        Set(ByVal value As Integer)
            _EnderecoCliente = value
            Cliente = Nothing
        End Set
    End Property

    Public Property Cliente() As Cliente
        Get
            If _Cliente Is Nothing And _CodigoCliente.Length > 0 Then _Cliente = New Cliente(_CodigoCliente, _EnderecoCliente)
            Return _Cliente
        End Get
        Set(ByVal value As Cliente)
            _Cliente = value
        End Set
    End Property

    '*************  MOV/LOTE/SEQ.   ***********************
    Public Property Movimento() As Date
        Get
            Return _Movimento
        End Get
        Set(ByVal value As Date)
            _Movimento = value
        End Set
    End Property

    Public Property Lote() As Integer
        Get
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

    '*************  UNIDADE DE NEGOCIO   ***********************
    Public Property CodigoUnidadeDeNegocio() As String
        Get
            Return _CodigoUnidadeDeNegocio
        End Get
        Set(ByVal value As String)
            _CodigoUnidadeDeNegocio = value
        End Set
    End Property

    Public ReadOnly Property UnidadeDeNegocio() As Negocio.Cliente
        Get
            If _UnidadeDeNegocio Is Nothing And _CodigoUnidadeDeNegocio.Length > 0 Then _UnidadeDeNegocio = New Cliente(_CodigoUnidadeDeNegocio, 0)
            Return _UnidadeDeNegocio
        End Get
    End Property

    '*************  TITULO   ***********************
    Public Property CodigoTitulo() As Integer
        Get
            Return _CodigoTitulo
        End Get
        Set(ByVal value As Integer)
            _CodigoTitulo = value
        End Set
    End Property

    Public Property Titulo() As Negocio.Titulo
        Get
            Return _Titulo
        End Get
        Set(ByVal value As Negocio.Titulo)
            _Titulo = value
        End Set
    End Property

    Public Property CodigoTituloNovo() As Integer
        Get
            Return _CodigoTituloNovo
        End Get
        Set(ByVal value As Integer)
            _CodigoTituloNovo = value
        End Set
    End Property

    Public Property TituloNovo() As Novo.TituloNovo
        Get
            Return _TituloNovo
        End Get
        Set(ByVal value As Novo.TituloNovo)
            _TituloNovo = value
        End Set
    End Property

    '*************  PEDIDO   ***********************
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
            If _Pedido Is Nothing And _CodigoEmpresa.Length > 0 And _CodigoPedido > 0 Then _Pedido = New Pedido(_CodigoEmpresa, _EndEmpresa, _CodigoPedido)
            Return _Pedido
        End Get
        Set(ByVal value As Pedido)
            _Pedido = value
        End Set
    End Property

    Public Property codigoPedidoFixacao() As Integer
        Get
            Return _CodigoPedidoFixacao
        End Get
        Set(ByVal value As Integer)
            _CodigoPedidoFixacao = value
        End Set
    End Property

    '*************  Produto   ***********************
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
            If _Produto Is Nothing And _CodigoProduto.Length > 0 Then _Produto = New Produto(_CodigoProduto)
            Return _Produto
        End Get
        Set(ByVal value As Produto)
            _Produto = value
        End Set
    End Property

    '*************  Centro De Custo   ***************
    Public Property CodigoCusto() As String
        Get
            Return _CodigoCusto
        End Get
        Set(ByVal value As String)
            _CodigoCusto = value
            _CentroDeCusto = Nothing
        End Set
    End Property

    Public Property CentroDeCusto() As CentroDeCusto
        Get
            If _CentroDeCusto Is Nothing And _CodigoCusto.Length > 0 Then _CentroDeCusto = New CentroDeCusto(_CodigoCusto)
            Return _CentroDeCusto
        End Get
        Set(ByVal value As CentroDeCusto)
            _CentroDeCusto = value
        End Set
    End Property

    '*************  Indexador   ***********************
    Public Property CodigoIndexador() As Integer
        Get
            Return _CodigoIndexador
        End Get
        Set(ByVal value As Integer)
            _CodigoIndexador = value
            _Indexador = Nothing
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

    '*************  Data Moeda   ***********************
    Public Property DataMoeda() As Date
        Get
            Return _DataMoeda
        End Get
        Set(ByVal value As Date)
            _DataMoeda = value
        End Set
    End Property

    '*************  Valores   ***********************
    Public Property DebitoOficial() As Double
        Get
            Return _DebitoOficial
        End Get
        Set(ByVal value As Double)
            _DebitoOficial = value
        End Set
    End Property

    Public Property CreditoOficial() As Double
        Get
            Return _CreditoOficial
        End Get
        Set(ByVal value As Double)
            _CreditoOficial = value
        End Set
    End Property

    Public Property DebitoMoeda() As Double
        Get
            Return _DebitoMoeda
        End Get
        Set(ByVal value As Double)
            _DebitoMoeda = value
        End Set
    End Property

    Public Property CreditoMoeda() As Double
        Get
            Return _CreditoMoeda
        End Get
        Set(ByVal value As Double)
            _CreditoMoeda = value
        End Set
    End Property

    '************* Historico   ***********************
    Public Property Historico() As String
        Get
            Return _Historico
        End Get
        Set(ByVal value As String)
            _Historico = value
        End Set
    End Property

    Public Property PrevistoRealizado() As String
        Get
            Return _PrevistoRealizado
        End Get
        Set(ByVal value As String)
            _PrevistoRealizado = value
        End Set
    End Property

    '************************************************************
    '**********  Notas Fiscais  *********************************
    '************************************************************
    Public Property CodigoClienteNF() As String
        Get
            Return _CodigoClienteNF
        End Get
        Set(ByVal value As String)
            _CodigoClienteNF = value
            _ClienteNF = Nothing
            _NotaFiscalRazao = Nothing
        End Set
    End Property

    Public Property EnderecoClienteNF() As Integer
        Get
            Return _EnderecoClienteNF
        End Get
        Set(ByVal value As Integer)
            _EnderecoClienteNF = value
            _ClienteNF = Nothing
            _NotaFiscalRazao = Nothing
        End Set
    End Property

    Public Property ClienteNF() As Cliente
        Get
            If _ClienteNF Is Nothing And _CodigoClienteNF > 0 Then _ClienteNF = New Cliente(_CodigoClienteNF, _EnderecoClienteNF)
            Return _ClienteNF
        End Get
        Set(ByVal value As Cliente)
            _ClienteNF = value
        End Set
    End Property

    Public Property EntradaSaidaNF() As String
        Get
            Return _EntradaSaidaNF
        End Get
        Set(ByVal value As String)
            _EntradaSaidaNF = value
            _NotaFiscalRazao = Nothing
        End Set
    End Property

    Public Property SerieNF() As String
        Get
            Return _SerieNF
        End Get
        Set(ByVal value As String)
            _SerieNF = value
            _NotaFiscalRazao = Nothing
        End Set
    End Property

    Public Property NumeroNF() As Integer
        Get
            Return _NumeroNF
        End Get
        Set(ByVal value As Integer)
            _NumeroNF = value
            _NotaFiscalRazao = Nothing
        End Set
    End Property

    Public Property CodigoProdutoNF() As String
        Get
            Return _CodigoProdutoNF
        End Get
        Set(ByVal value As String)
            _CodigoProdutoNF = value
            _ProdutoNF = Nothing
        End Set
    End Property

    Public Property ProdutoNF() As Produto
        Get
            If _ProdutoNF Is Nothing And _CodigoProdutoNF.Length > 0 Then _ProdutoNF = New Produto(_CodigoProduto)
            Return _ProdutoNF
        End Get
        Set(ByVal value As Produto)
            _ProdutoNF = value
        End Set
    End Property

    Public Property CodigoCFOPNF() As Integer
        Get
            Return _CodigoCFOPNF
        End Get
        Set(ByVal value As Integer)
            _CodigoCFOPNF = value
            _CFOPNF = Nothing
        End Set
    End Property

    Public Property CFOPNF() As CFOP
        Get
            If _CFOPNF Is Nothing And _CodigoCFOPNF > 0 Then _CFOPNF = New CFOP(_CodigoCFOPNF)
            Return _CFOPNF
        End Get
        Set(ByVal value As CFOP)
            _CFOPNF = value
        End Set
    End Property

    Public Property SequenciaNF() As Integer
        Get
            Return _SequenciaNF
        End Get
        Set(ByVal value As Integer)
            _SequenciaNF = value
        End Set
    End Property

    Public Property CodigoEncargoNF() As String
        Get
            Return _CodigoEncargoNF
        End Get
        Set(ByVal value As String)
            _CodigoEncargoNF = value
            _EncargoNF = Nothing
        End Set
    End Property

    Public Property EncargoNF() As Encargo
        Get
            If _EncargoNF Is Nothing And _CodigoEncargoNF.Length > 0 Then _EncargoNF = New Encargo(_CodigoEncargoNF)
            Return _EncargoNF
        End Get
        Set(ByVal value As Encargo)
            _EncargoNF = value
        End Set
    End Property

    Public Property NotaFiscalRazao() As NotaFiscal
        Get
            If _NotaFiscalRazao Is Nothing And _NumeroNF > 0 Then
                Dim cons As New NotaFiscal
                cons.CodigoEmpresa = _CodigoEmpresa
                cons.EnderecoEmpresa = _EndEmpresa
                cons.CodigoCliente = _CodigoClienteNF
                cons.EnderecoCliente = _EnderecoClienteNF
                cons.Codigo = _NumeroNF
                cons.Serie = _SerieNF
                cons.EntradaSaida = IIf(_EntradaSaidaNF = "S", eEntradaSaida.Saida, eEntradaSaida.Entrada)
                _NotaFiscalRazao = New NotaFiscal(cons)
            End If
            Return _NotaFiscalRazao
        End Get
        Set(ByVal value As NotaFiscal)
            _NotaFiscalRazao = value
        End Set
    End Property

    Public Property NotaFiscalXItemRazao() As NotaFiscalXItem
        Get
            Return _NotaFiscalXItemRazao
        End Get
        Set(ByVal value As NotaFiscalXItem)
            _NotaFiscalXItemRazao = value
        End Set
    End Property

    Public Property NotaFiscalXItemXEncargoRazao() As NotaFiscalXItemXEncargo
        Get
            Return _NotaFiscalXItemXEncargoRazao
        End Get
        Set(ByVal value As NotaFiscalXItemXEncargo)
            _NotaFiscalXItemXEncargoRazao = value
        End Set
    End Property
    '************************************************************

    Public Property ChequeEntregue() As String
        Get
            Return _ChequeEntregue
        End Get
        Set(ByVal value As String)
            _ChequeEntregue = value
        End Set
    End Property

    Public Property PagamentoAutorizado() As String
        Get
            Return _PagamentoAutorizado
        End Get
        Set(ByVal value As String)
            _PagamentoAutorizado = value
        End Set
    End Property

    Public Property DataBaixa() As Date
        Get
            Return _DataBaixa
        End Get
        Set(ByVal value As Date)
            _DataBaixa = value
        End Set
    End Property

    Public Property Conciliacao() As String
        Get
            Return _Conciliacao
        End Get
        Set(ByVal value As String)
            _Conciliacao = value
        End Set
    End Property

    '*******************  Deposito  *********************
    Public Property CodigoDeposito() As String
        Get
            Return _CodigoDeposito
        End Get
        Set(ByVal value As String)
            _CodigoDeposito = value
        End Set
    End Property

    Public Property EndDeposito() As Integer
        Get
            Return _EndDeposito
        End Get
        Set(ByVal value As Integer)
            _EndDeposito = value
        End Set
    End Property

    Public ReadOnly Property Deposito() As Negocio.Cliente
        Get
            If _Deposito Is Nothing And _CodigoDeposito.Length > 0 Then _Deposito = New Negocio.Cliente(_CodigoDeposito, _EndDeposito)
            Return _Deposito
        End Get
    End Property

    '****************************************************
    Public Property Rateado() As Boolean
        Get
            Return _Rateado
        End Get
        Set(ByVal value As Boolean)
            _Rateado = value
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

    Public Property DebitoQuantidade() As Decimal
        Get
            Return _DebitoQuantidade
        End Get
        Set(ByVal value As Decimal)
            _DebitoQuantidade = value
        End Set
    End Property

    Public Property CreditoQuantidade() As Decimal
        Get
            Return _CreditoQuantidade
        End Get
        Set(ByVal value As Decimal)
            _CreditoQuantidade = value
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

    Public Property UsuarioInclusao() As String
        Get
            Return _UsuarioInclusao
        End Get
        Set(ByVal value As String)
            _UsuarioInclusao = value
        End Set
    End Property

    Public Property UsuarioDataInclusao() As Date
        Get
            Return _UsuarioDataInclusao
        End Get
        Set(ByVal value As Date)
            _UsuarioDataInclusao = value
        End Set
    End Property
    Public Property Saldo As Decimal
        Get
            Return _Saldo
        End Get
        Set(ByVal value As Decimal)
            _Saldo = value
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
        Return Banco.GravaBanco(Sqls)
    End Function

    Public Sub SalvarSql(ByRef Sqls As ArrayList)
        Sqls.Add(SalvarSql())
    End Sub

    Public Function SalvarSql(Optional ByVal pSequencia As Integer = 0) As String
        Dim Sql As String = ""
        Select Case Me.IUD
            Case "I"
                If _Sequencia = -1 Then
                    Sql = " declare" & vbCrLf & _
                          " @Sequencia int" & vbCrLf & _
                          " Set @sequencia = (Select Sequencia" & vbCrLf & _
                          "                     from numerador" & vbCrLf & _
                          "                    where UPPER(Empresa_id) = '" & HttpContext.Current.Session("ssNomeServidor").ToUpper() & "'" & vbCrLf & _
                          "                      and Numerador_id      = 60)" & vbCrLf & _
                          " select @Sequencia = isnull(Max(Sequencia_Id),isnull(@sequencia,0))" & vbCrLf & _
                          "   from razao    " & vbCrLf & _
                          "  Where lote_Id      = " & _Lote & vbCrLf & _
                          "    and Movimento_Id ='" & CDate(Me.Movimento).ToString("yyyy/MM/dd") & "'" & vbCrLf & _
                          "    and Sequencia_id between isnull(@sequencia,0) and isnull(@sequencia,0) + 99999" & vbCrLf
                End If
                Sql &= " INSERT RAZAO (Empresa_Id, EndEmpresa_Id, Conta_Id, Cliente_Id, EndCliente_Id, Movimento_Id, Lote_Id, Sequencia_Id, UnidadeDeNegocio, Titulo, Pedido, PedidoFixacao," & vbCrLf & _
                      "       Produto, Custo, Indexador, DataMoeda, DebitoOficial, CreditoOficial, DebitoMoeda, CreditoMoeda, Historico, PrevistoRealizado, Cliente_Nf, EndCliente_Nf," & vbCrLf & _
                      "       EntradaSaida_Nf, Serie_Nf, Numero_Nf, ChequeEntregue, PagamentoAutorizado, DataDaBaixa, Conciliacao, Deposito, EndDeposito, Rateado," & vbCrLf & _
                      "       Processo, DebitoQuantidade, CreditoQuantidade, Situacao, Produto_NF, Cfop_NF, Sequencia_NF, Encargo_NF)" & vbCrLf & _
                      " values ('" & _CodigoEmpresa & "'," & _EndEmpresa & ",'" & _CodigoConta & "','" & _CodigoCliente & "'," & _EnderecoCliente & ",'" & _Movimento.ToString("yyyy-MM-dd") & "'," & _Lote & "," & IIf(_Sequencia = -1, "@Sequencia + " & pSequencia, _Sequencia) & "," & IIf(_CodigoUnidadeDeNegocio.Length > 0, "'" & _CodigoUnidadeDeNegocio & "'", "NULL") & "," & IIf(_CodigoTitulo > 0, _CodigoTitulo, "NULL") & "," & IIf(_CodigoPedido > 0, _CodigoPedido, "NULL") & "," & IIf(_CodigoPedidoFixacao > 0, _CodigoPedidoFixacao, "NULL") & "," & vbCrLf & _
                      "          " & IIf(_CodigoProduto.Length > 0, "'" & _CodigoProduto & "'", "NULL") & ",'" & _CodigoCusto & "'," & _CodigoIndexador & ",'" & DataMoeda.ToString("yyyy-MM-dd") & "'," & Str(Math.Round(_DebitoOficial, 2)) & "," & Str(Math.Round(_CreditoOficial, 2)) & "," & Str(Math.Round(_DebitoMoeda, 2)) & "," & Str(Math.Round(_CreditoMoeda, 2)) & ",'" & _Historico & "','" & _PrevistoRealizado & "'," & IIf(_CodigoClienteNF.Length > 0, "'" & _CodigoClienteNF & "'", "NULL") & "," & IIf(_CodigoClienteNF.Length > 0, _EnderecoClienteNF, "NULL") & "," & vbCrLf & _
                      "          " & IIf(_CodigoClienteNF.Length > 0, "'" & _EntradaSaidaNF & "'", "NULL") & "," & IIf(_CodigoClienteNF.Length > 0, "'" & _SerieNF & "'", "NULL") & "," & IIf(_CodigoClienteNF.Length > 0, _NumeroNF, "NULL") & "," & IIf(_ChequeEntregue.Length > 0, "'" & _ChequeEntregue & "'", "NULL") & "," & IIf(_PagamentoAutorizado.Length > 0, "'" & _PagamentoAutorizado & "'", "NULL") & ",'" & _DataBaixa.ToString("yyyy-MM-dd") & "'," & IIf(_Conciliacao.Length > 0, "'" & _Conciliacao & "'", "NULL") & "," & IIf(_CodigoDeposito.Length > 0, "'" & _CodigoDeposito & "'", "NULL") & "," & IIf(_CodigoDeposito.Length > 0, _EndDeposito, "NULL") & "," & CByte(_Rateado) & "," & _
                      "         '" & _Processo & "'," & Str(_DebitoQuantidade) & "," & Str(_CreditoQuantidade) & "," & IIf(_CodigoSituacao = 0, 1, _CodigoSituacao) & "," & IIf(_CodigoClienteNF.Length > 0, "'" & _CodigoProdutoNF & "'", "NULL") & "," & IIf(_CodigoClienteNF.Length > 0, _CodigoCFOPNF, "NULL") & "," & IIf(_CodigoClienteNF.Length > 0, _SequenciaNF, "NULL") & "," & IIf(_CodigoClienteNF.Length > 0, "'" & _CodigoEncargoNF & "'", "NULL") & ")"
                Return Sql
            Case "U"
                Sql = " UPDATE RAZAO SET" & vbCrLf & _
                      "    UnidadeDeNegocio    = " & IIf(_CodigoUnidadeDeNegocio.Length > 0, "'" & _CodigoUnidadeDeNegocio & "'", "NULL") & vbCrLf & _
                      "   ,Titulo              = " & IIf(_CodigoTitulo > 0, _CodigoTitulo, "NULL") & vbCrLf & _
                      "   ,Pedido              = " & IIf(_CodigoPedido > 0, _CodigoPedido, "NULL") & vbCrLf & _
                      "   ,PedidoFixacao       = " & IIf(_CodigoPedidoFixacao > 0, _CodigoPedidoFixacao, "NULL") & vbCrLf & _
                      "   ,Produto             = " & IIf(_CodigoProduto.Length > 0, "'" & _CodigoProduto & "'", "NULL") & vbCrLf & _
                      "   ,Custo               = " & IIf(_CodigoCusto.Length > 0, "'" & _CodigoCusto & "'", "NULL") & vbCrLf & _
                      "   ,Indexador           = " & _CodigoIndexador & vbCrLf & _
                      "   ,DataMoeda           ='" & _DataMoeda.ToString("yyyy-MM-dd") & "'" & vbCrLf & _
                      "   ,DebitoOficial       = " & Str(Math.Round(_DebitoOficial, 2)) & vbCrLf & _
                      "   ,CreditoOficial      = " & Str(Math.Round(_CreditoOficial, 2)) & vbCrLf & _
                      "   ,DebitoMoeda         = " & Str(Math.Round(_DebitoMoeda, 2)) & vbCrLf & _
                      "   ,CreditoMoeda        = " & Str(Math.Round(_CreditoMoeda, 2)) & vbCrLf & _
                      "   ,Historico           ='" & _Historico & "'" & vbCrLf & _
                      "   ,PrevistoRealizado   ='" & _PrevistoRealizado & "'" & vbCrLf & _
                      "   ,Cliente_Nf          = " & IIf(_CodigoClienteNF.Length > 0, "'" & _CodigoClienteNF & "'", "NULL") & vbCrLf & _
                      "   ,EndCliente_Nf       = " & IIf(_CodigoClienteNF.Length > 0, _EnderecoClienteNF, "NULL") & vbCrLf & _
                      "   ,EntradaSaida_Nf     = " & IIf(_CodigoClienteNF.Length > 0, "'" & _EntradaSaidaNF & "'", "NULL") & vbCrLf & _
                      "   ,Serie_Nf            = " & IIf(_CodigoClienteNF.Length > 0, "'" & _SerieNF & "'", "NULL") & vbCrLf & _
                      "   ,Numero_Nf           = " & IIf(_CodigoClienteNF.Length > 0, _NumeroNF, "NULL") & vbCrLf & _
                      "   ,ChequeEntregue      = " & IIf(_ChequeEntregue.Length > 0, "'" & _ChequeEntregue & "'", "NULL") & vbCrLf & _
                      "   ,PagamentoAutorizado = " & IIf(_PagamentoAutorizado.Length > 0, "'" & _PagamentoAutorizado & "'", "NULL") & vbCrLf & _
                      "   ,DataDaBaixa         ='" & _DataBaixa.ToString("yyyy-MM-dd") & "'" & vbCrLf & _
                      "   ,Conciliacao         = " & IIf(_Conciliacao.Length > 0, "'" & _Conciliacao & "'", "NULL") & vbCrLf & _
                      "   ,Deposito            = " & IIf(_CodigoDeposito.Length > 0, "'" & _CodigoDeposito & "'", "NULL") & vbCrLf & _
                      "   ,EndDeposito         = " & IIf(_CodigoDeposito.Length > 0, _EndDeposito, "NULL") & vbCrLf & _
                      "   ,Rateado             = " & CByte(_Rateado) & vbCrLf & _
                      "   ,Processo            ='" & _Processo & "'" & vbCrLf & _
                      "   ,DebitoQuantidade    = " & Str(_DebitoQuantidade) & vbCrLf & _
                      "   ,CreditoQuantidade   = " & Str(_CreditoQuantidade) & vbCrLf & _
                      "   ,Situacao            = " & IIf(_CodigoSituacao = 0, 1, _CodigoSituacao) & vbCrLf & _
                      "   ,Produto_NF          = " & IIf(_CodigoClienteNF.Length > 0, "'" & _CodigoProdutoNF & "'", "NULL") & vbCrLf & _
                      "   ,Cfop_NF             = " & IIf(_CodigoClienteNF.Length > 0, _CodigoCFOPNF, "NULL") & vbCrLf & _
                      "   ,Sequencia_NF        = " & IIf(_CodigoClienteNF.Length > 0, _SequenciaNF, "NULL") & vbCrLf & _
                      "   ,Encargo_NF          = " & IIf(_CodigoClienteNF.Length > 0, "'" & _CodigoEncargoNF & "'", "NULL") & vbCrLf & _
                      "  WHERE Empresa_Id    ='" & _CodigoEmpresa & "'" & vbCrLf & _
                      "    And EndEmpresa_Id = " & _EndEmpresa & vbCrLf & _
                      "    And Conta_Id      ='" & _CodigoConta & "'" & vbCrLf & _
                      "    And Cliente_Id    ='" & _CodigoCliente & "'" & vbCrLf & _
                      "    And EndCliente_Id = " & _EnderecoCliente & vbCrLf & _
                      "    And Movimento_Id  ='" & _Movimento.ToString("yyyy-MM-dd") & "'" & vbCrLf & _
                      "    And Lote_Id       = " & _Lote & vbCrLf & _
                      "    And Sequencia_Id  = " & _Sequencia & vbCrLf
                Return Sql
            Case "D"
                Sql = " DELETE RAZAO" & vbCrLf & _
                      "  WHERE Empresa_Id    ='" & _CodigoEmpresa & "'" & vbCrLf & _
                      "    And EndEmpresa_Id = " & _EndEmpresa & vbCrLf & _
                      "    And Conta_Id      ='" & _CodigoConta & "'" & vbCrLf & _
                      "    And Cliente_Id    ='" & _CodigoCliente & "'" & vbCrLf & _
                      "    And EndCliente_Id = " & _EnderecoCliente & vbCrLf & _
                      "    And Movimento_Id  ='" & _Movimento.ToString("yyyy-MM-dd") & "'" & vbCrLf & _
                      "    And Lote_Id       = " & _Lote & vbCrLf & _
                      "    And Sequencia_Id  = " & _Sequencia & vbCrLf
                Return Sql
        End Select
        Return Sql
    End Function
#End Region

#Region "Contabiliza Nota/Titulo"
    Public Sub ContabilizarPAMCARD(ByVal obj As Titulo, ByVal codigoConta As String, ByVal debCredito As String, ByVal vlrDocumento As Decimal, ByVal vlrMoeda As Decimal, ByVal historico As String, ByRef Sqls As ArrayList)
        Dim sql As String = ""
        sql = "INSERT INTO Razao "
        sql &= " (Empresa_Id, "
        sql &= " EndEmpresa_Id, "
        sql &= " Conta_Id, "
        sql &= " Cliente_Id, "
        sql &= " EndCliente_Id, "
        sql &= " Movimento_Id, "
        sql &= " Lote_Id, "
        sql &= " Sequencia_Id, "
        sql &= " Titulo, "
        sql &= " UnidadeDeNegocio, "
        sql &= " Indexador, "
        sql &= " DataMoeda, "
        sql &= " DebitoOficial, "
        sql &= " CreditoOficial, "
        sql &= " DebitoMoeda, "
        sql &= " CreditoMoeda, "
        sql &= " Conciliacao, "
        sql &= " DataDaBaixa, "
        sql &= " Historico, "
        sql &= " PrevistoRealizado,"
        sql &= " Processo,"
        sql &= " UsuarioInclusao,"
        sql &= " UsuarioInclusaoData)"
        sql &= " VALUES ("

        sql &= "'" & obj.CodigoEmpresa & "'"            'Empresa
        sql &= ", " & obj.EnderecoEmpresa               'Endereco Empresa 
        sql &= ", '" & codigoConta & "'"                'Conta

        If Len(codigoConta) = 7 Then
            sql &= ", '" & obj.CodigoCliente & "'"      'Cliente
            sql &= ", " & obj.EndCliente                'Endereco do Cliente
        Else
            sql &= ", ''"                               'Cliente
            sql &= ", 0"                                'Endereco do Cliente
        End If

        sql &= ", '" & CDate(obj.Movimento).ToString("yyyy/MM/dd") & "'"     'Data de Movimento
        sql &= ", 0072"
        sql &= ", " & obj.Codigo                                            'Sequencia no Razao = Registro do Titulo
        sql &= ", " & IIf(obj.Codigo > 0, obj.Codigo, "NULL")                 'Numero do Titulo
        sql &= ", '" & obj.CodigoUnidadeDeNegocio & "'"                     'Unidade de Negócio
        sql &= ", " & obj.CodigoIndexador                                   'Indexador
        sql &= ", '" & CDate(obj.Movimento).ToString("yyyy/MM/dd") & "'"     'Data da Moeda

        If debCredito = "D" Then
            sql &= ", " & Replace(vlrDocumento, ",", ".")       'Valor Débito
            sql &= ", 0.0"                                      'Valor Crédito
            sql &= ", " & Replace(vlrMoeda, ",", ".")           'Valor Débito
            sql &= ", 0.0"                                      'Valor Crédito
        Else
            sql &= ", 0.0"                                      'Valor Debito
            sql &= ", " & Replace(vlrDocumento, ",", ".")       'Valor Crédito
            sql &= ", 0.0"                                      'Valor Debito
            sql &= ", " & Replace(vlrMoeda, ",", ".")           'Valor Crédito
        End If

        sql &= ", NULL "                                                            'Conciliação
        sql &= ", NULL "                                                            'Data Conciliação
        sql &= ", '" & obj.Historico & "'"                                          'Histórico
        sql &= ", 'P'"                                                              'Previsto/Realizado
        sql &= ", 'CONHECIMENTODETRANSPORTE'"                                       'Processo
        sql &= ", '" & HttpContext.Current.Session("ssNomeUsuario") & "'"           'Usuario que Baixou
        sql &= ", '" & CDate(Today).ToString("yyyy/MM/dd") & "')"                    'Data da Baixa

        Sqls.Add(sql)
    End Sub

    Public Sub ContabilizarPAMCARD(ByVal obj As Novo.TituloNovo, ByVal codigoConta As String, ByVal debCredito As String, ByVal vlrDocumento As Decimal, ByVal vlrMoeda As Decimal, ByVal historico As String, ByRef Sqls As ArrayList)
        Dim sql As String = ""
        sql = "INSERT INTO Razao "
        sql &= " (Empresa_Id, "
        sql &= " EndEmpresa_Id, "
        sql &= " Conta_Id, "
        sql &= " Cliente_Id, "
        sql &= " EndCliente_Id, "
        sql &= " Movimento_Id, "
        sql &= " Lote_Id, "
        sql &= " Sequencia_Id, "
        sql &= " Titulo, "
        sql &= " UnidadeDeNegocio, "
        sql &= " Indexador, "
        sql &= " DataMoeda, "
        sql &= " DebitoOficial, "
        sql &= " CreditoOficial, "
        sql &= " DebitoMoeda, "
        sql &= " CreditoMoeda, "
        sql &= " Conciliacao, "
        sql &= " DataDaBaixa, "
        sql &= " Historico, "
        sql &= " PrevistoRealizado,"
        sql &= " Processo,"
        sql &= " UsuarioInclusao,"
        sql &= " UsuarioInclusaoData)"
        sql &= " VALUES ("

        sql &= "'" & obj.CodigoEmpresa & "'"            'Empresa
        sql &= ", " & obj.EnderecoEmpresa               'Endereco Empresa 
        sql &= ", '" & codigoConta & "'"                'Conta

        If Len(codigoConta) = 7 Then
            sql &= ", '" & obj.CodigoCliFor & "'"      'Cliente
            sql &= ", " & obj.EnderecoCliFor           'Endereco do Cliente
        Else
            sql &= ", ''"                               'Cliente
            sql &= ", 0"                                'Endereco do Cliente
        End If

        sql &= ", '" & CDate(obj.Movimento).ToString("yyyy/MM/dd") & "'"     'Data de Movimento
        sql &= ", 0072"
        sql &= ", " & obj.Codigo                                            'Sequencia no Razao = Registro do Titulo
        sql &= ", " & IIf(obj.Codigo > 0, obj.Codigo, "NULL")               'Numero do Titulo
        sql &= ", '" & obj.CodigoUnidadeDeNegocio & "'"                     'Unidade de Negócio
        sql &= ", " & obj.CodigoIndexador                                   'Indexador
        sql &= ", '" & CDate(obj.Movimento).ToString("yyyy/MM/dd") & "'"     'Data da Moeda

        If debCredito = "D" Then
            sql &= ", " & Replace(vlrDocumento, ",", ".")       'Valor Débito
            sql &= ", 0.0"                                      'Valor Crédito
            sql &= ", " & Replace(vlrMoeda, ",", ".")           'Valor Débito
            sql &= ", 0.0"                                      'Valor Crédito
        Else
            sql &= ", 0.0"                                      'Valor Debito
            sql &= ", " & Replace(vlrDocumento, ",", ".")       'Valor Crédito
            sql &= ", 0.0"                                      'Valor Debito
            sql &= ", " & Replace(vlrMoeda, ",", ".")           'Valor Crédito
        End If

        sql &= ", NULL "                                                            'Conciliação
        sql &= ", NULL "                                                            'Data Conciliação
        sql &= ", '" & obj.Historico & "'"                                          'Histórico
        sql &= ", 'P'"                                                              'Previsto/Realizado
        sql &= ", 'CONHECIMENTODETRANSPORTE'"                                       'Processo
        sql &= ", '" & HttpContext.Current.Session("ssNomeUsuario") & "'"           'Usuario que Baixou
        sql &= ", '" & CDate(Today).ToString("yyyy/MM/dd") & "')"                    'Data da Baixa

        Sqls.Add(sql)
    End Sub

    Public Function ContabilizarNota(Optional ByVal pLote As Integer = 0) As Boolean
        Dim Sqls As New ArrayList
        Dim Banco As New AcessaBanco
        ContabilizarNotaSql(Sqls, pLote)
        Return Banco.GravaBanco(Sqls)
    End Function

    Public Sub ContabilizarNotaSql(ByRef sqls As ArrayList, Optional ByVal pLote As Integer = 0)
        Contabilizador(sqls, "", pLote)
    End Sub

    Public Function VisualizarContabilizacao(Optional ByVal preview As Boolean = False) As DataSet
        Dim Banco As New AcessaBanco
        Dim sql As String = ""
        If preview Then Contabilizador(New ArrayList, sql, True)

        sql = "select R.Conta_Id as conta, R.Cliente_Id +'-'+ convert(nvarchar, R.EndCliente_Id) as Cliente, PlanoDeContas.Titulo, R.Movimento_Id as Movimento, R.Lote_Id as Lote, " & vbCrLf & _
              "       isnull(R.Produto, '') AS Produto, R.Custo, R.Historico, R.DebitoOficial as Debito, R.CreditoOficial as Credito, R.DebitoOficial -  R.CreditoOficial as Saldo " & vbCrLf & _
              "  from " & IIf(preview, "#", "") & "Razao R" & vbCrLf & _
              " INNER JOIN PlanoDeContas ON R.Conta_Id = PlanoDeContas.Conta_Id " & vbCrLf


        If Not NotaFiscalRazao Is Nothing Then
            sql &= " where R.Empresa_id    = '" & _CodigoEmpresa & "'" & vbCrLf & _
                   "   and R.EndEmpresa_Id = " & _EndEmpresa & vbCrLf & _
                   "   and (R.Pedido        = " & _CodigoPedido & vbCrLf & _
                   "    and R.Lote_Id       in (9,10) " & vbCrLf & _
                   "    and R.Serie_Nf      = '" & _SerieNF & "'" & vbCrLf & _
                   "    and R.Numero_Nf     = " & _NumeroNF & ")" & vbCrLf & _
                   "   or  (R.Pedido = " & _CodigoPedido & " and R.Lote_Id = 70)"
        End If


        If Not Titulo Is Nothing Then
            sql &= " Where R.Empresa_id    = '" & _CodigoEmpresa & "'" & vbCrLf & _
                   "   and R.EndEmpresa_Id = " & _EndEmpresa & vbCrLf & _
                   "   and R.Titulo        = " & _CodigoTitulo
        End If

        sql &= "      ORDER BY R.Lote_Id"

        Return Banco.ConsultaDataSet(sql, "PreView")
    End Function

    Public Sub ExcluiContabilizacaoNotaSQL(ByRef sqls As ArrayList, Optional ByVal plote As Integer = 0)
        Dim Sql As String
        Sql = " DELETE From Razao" & vbCrLf & _
            " WHERE Lote_Id "
        If plote > 0 Then
            Sql &= " = " & plote & vbCrLf
        ElseIf _NotaFiscalRazao.CodigoTipoDeDocumento = CInt(eTipoDeDocumento.CTRC) Or _
            _NotaFiscalRazao.CodigoTipoDeDocumento = CInt(eTipoDeDocumento.ReciboDeFrete) Or _
            _NotaFiscalRazao.CodigoTipoDeDocumento = CInt(eTipoDeDocumento.CTRC_SEM_NF) Or _
            _NotaFiscalRazao.CodigoTipoDeDocumento = CInt(eTipoDeDocumento.CT_E_TOM) Or _
            _NotaFiscalRazao.CodigoTipoDeDocumento = CInt(eTipoDeDocumento.Estadia) Or _
            _NotaFiscalRazao.CodigoTipoDeDocumento = CInt(eTipoDeDocumento.CT_E) Or _
            _NotaFiscalRazao.CodigoTipoDeDocumento = CInt(eTipoDeDocumento.CT_E_OUT) Or _
            _NotaFiscalRazao.CodigoTipoDeDocumento = CInt(eTipoDeDocumento.Anulacao) Then
            Sql &= " = 21 " & vbCrLf
        Else
            Sql &= " in (9,10,11) " & vbCrLf
        End If

        Sql &= "    And Empresa_Id      ='" & _NotaFiscalRazao.CodigoEmpresa & "'" & vbCrLf & _
              "    And EndEmpresa_Id   = " & _NotaFiscalRazao.EnderecoEmpresa & vbCrLf & _
              "    And Cliente_nf      ='" & _NotaFiscalRazao.CodigoCliente & "'" & vbCrLf & _
              "    And EndCliente_nf   = " & _NotaFiscalRazao.EnderecoCliente & vbCrLf & _
              "    And Numero_nf       = " & _NotaFiscalRazao.Codigo & vbCrLf & _
              "    And Serie_nf        ='" & _NotaFiscalRazao.Serie & "'" & vbCrLf & _
              "    And EntradaSaida_nf ='" & _NotaFiscalRazao.EntradaSaida.ToString.Substring(0, 1) & "'"
        sqls.Add(Sql)
    End Sub

    Public Sub ExcluiContabilizacaoTitulo(ByRef sqls As ArrayList)
        Dim Sql As String
        Sql = " DELETE From Razao" & vbCrLf & _
              "  WHERE Lote_Id         = 70" & vbCrLf & _
              "    And Empresa_Id      ='" & _CodigoEmpresa & "'" & vbCrLf & _
              "    And EndEmpresa_Id   = " & _EndEmpresa & vbCrLf & _
              "    And Titulo          = " & _CodigoTitulo
        sqls.Add(Sql)
    End Sub

    Public Sub Contabilizador(ByVal Sqls As ArrayList, ByRef Sql As String, Optional ByVal pLote As Integer = 0, Optional ByVal pEmpresa As String = "", Optional ByVal pEndereco As String = "")
        If Not NotaFiscalRazao.SubOperacao.Contabil Then Exit Sub

        If NotaFiscalRazao.CodigoSituacao = eSituacao.Cancelado OrElse _
            NotaFiscalRazao.CodigoSituacao = eSituacao.Inutilizada OrElse _
            NotaFiscalRazao.CodigoSituacao = eSituacao.Denegado Then Exit Sub

        Dim Banco As New AcessaBanco

        Dim Transferecia As Boolean = False
        Dim SeqTransf As Integer = 0

        'Dim EncLiq As Negocio.NotaFiscalXItemXEncargo = Nothing
        'Dim LiqDC As String = ""
        'Dim LiqConta As String = ""
        'Dim LiqValorConvertido As Decimal

        Dim EncTaxaPedagio As Negocio.NotaFiscalXItemXEncargo = Nothing
        Dim EncTarifaPedagio As Negocio.NotaFiscalXItemXEncargo = Nothing
        Dim LiqValorConvertidoTaxa As Decimal = Decimal.Zero
        Dim LiqValorConvertidoTarifa As Decimal = Decimal.Zero

        Sql = "Declare" & vbCrLf & _
              " @SequenciaOriginal int" & vbCrLf & _
              ",@SequenciaServidor int" & vbCrLf

        If (Me.NotaFiscalRazao.CodigoTipoDeDocumento = CInt(eTipoDeDocumento.Nota) OrElse Me.NotaFiscalRazao.TipoDeDocumentoFrete = eTipoDeDocumentoFrete.Circulacao) _
            AndAlso Not NotaFiscalRazao.Pedido Is Nothing AndAlso (NotaFiscalRazao.CodigoCliente <> NotaFiscalRazao.Pedido.CodigoCliente OrElse NotaFiscalRazao.EnderecoCliente <> NotaFiscalRazao.Pedido.EnderecoCliente) Then

            Sql &= ",@SequenciaTransferenciaCliente int" & vbCrLf &
                   " Set @SequenciaOriginal = (Select Sequencia" & vbCrLf &
                   "                             from numerador" & vbCrLf &
                   "                            where UPPER(Empresa_id) = '" & HttpContext.Current.Session("ssNomeServidor").ToUpper() & "'" & vbCrLf &
                   "                              and Numerador_id      = 60)" & vbCrLf &
                   " Set @SequenciaServidor = (select isnull(Max(Sequencia_Id),isnull(@SequenciaOriginal,0))" & vbCrLf &
                   "                             from razao (nolock) " & vbCrLf &
                   "                            Where lote_Id      = " & IIf(pLote > 0, pLote, IIf(_NotaFiscalRazao.CodigoOperacao > 68, "9", "10")) & vbCrLf &
                   "                              and Movimento_Id = '" & CDate(Me.NotaFiscalRazao.Movimento).ToString("yyyy/MM/dd") & "'" & vbCrLf &
                   "                              and Sequencia_id between isnull(@SequenciaOriginal,0) and isnull(@SequenciaOriginal,0) + 99999)" & vbCrLf &
                   " Set @SequenciaTransferenciaCliente = (Select isnull(Max(Sequencia_Id),isnull(@SequenciaOriginal,0))" & vbCrLf &
                   "                                         from razao (nolock) " & vbCrLf &
                   "                                        Where lote_Id      = 11" & vbCrLf &
                   "                                          and Movimento_Id = '" & CDate(Me.NotaFiscalRazao.Movimento).ToString("yyyy/MM/dd") & "'" & vbCrLf &
                   "                                          and Sequencia_id between isnull(@SequenciaOriginal,0) and isnull(@SequenciaOriginal,0) + 99999)" & vbCrLf
            Transferecia = True
        Else
            Sql &= " Set @SequenciaOriginal = (Select Sequencia" & vbCrLf & _
                   "                             from numerador" & vbCrLf & _
                   "                            where UPPER(Empresa_id) = '" & HttpContext.Current.Session("ssNomeServidor").ToUpper() & "'" & vbCrLf & _
                   "                              and Numerador_id      = 60)" & vbCrLf & _
                   " set @SequenciaServidor = (select isnull(Max(Sequencia_Id),isnull(@SequenciaOriginal,0))" & vbCrLf & _
                   "                             from razao  (nolock) " & vbCrLf


            If pLote > 0 Then
                Sql &= "                            Where lote_Id = " & pLote & vbCrLf
            ElseIf _NotaFiscalRazao.CodigoOperacao > 68 Then

                If Me.NotaFiscalRazao.CodigoTipoDeDocumento = 1 Then
                    Sql &= "                            Where lote_Id = 10" & vbCrLf
                Else
                    Sql &= "                            Where lote_Id = 9" & vbCrLf
                End If

            Else
                Sql &= "                            Where lote_Id = 10" & vbCrLf
            End If

            Sql &= "                              and Movimento_Id = '" & CDate(Me.NotaFiscalRazao.Movimento).ToString("yyyy/MM/dd") & "'" & vbCrLf & _
                   "                              and Sequencia_id between isnull(@SequenciaOriginal,0) and isnull(@SequenciaOriginal,0) + 99999)" & vbCrLf
        End If

        Dim i As Integer = 0
        For Each Item As Negocio.NotaFiscalXItem In NotaFiscalRazao.Itens.Where(Function(s) s.IUD <> "D")
            Dim lstEncargos As New List(Of NotaFiscalXItemXEncargo)

            If Me.NotaFiscalRazao.CodigoTipoDeDocumento = CInt(eTipoDeDocumento.CTRC) OrElse Me.NotaFiscalRazao.CodigoTipoDeDocumento = CInt(eTipoDeDocumento.CT_E) Then
                lstEncargos = Item.Encargos.Where(Function(s) Not s.Codigo.Trim.Contains("ADIANTAMENTO") AndAlso Not s.Codigo.Trim.Contains("ADTODEFRETE")).ToList()
            Else
                lstEncargos = Item.Encargos
            End If

            For Each Enc As Negocio.NotaFiscalXItemXEncargo In lstEncargos
                '****************************************************************
                '**** CASO TENHA ENCARGO NO VALOR NOVO ASSUMIR O MESMO **********
                '****************************************************************
                If Enc.ValorNovo > 0 Then Enc.Valor = Enc.ValorNovo

                '****************************************************************
                '**** CONVERTE O VALOR DO ENCARGO EM MOEDA EXTRANGEIRO **********
                '****************************************************************
                Dim VlrConvertido As Decimal = 0
                If NotaFiscalRazao.CodigoPedido > 0 Then
                    VlrConvertido = Funcoes.ConverteParaMoedaExtrangeira(Enc.Valor, 2, Me.NotaFiscalRazao.DataNota, True, False, 2)
                Else
                    VlrConvertido = 0
                End If

                '****************************************************
                '**** PEGA A REFERENCIA DO ENCARGO LIQUIDO **********
                '****************************************************
                'If Enc.Codigo = "LIQUIDO" Then
                '    EncLiq = Enc
                '    LiqValorConvertido = VlrConvertido
                'End If

                If Enc.Codigo = "TAXA PEDAGIO" Then
                    EncTaxaPedagio = Enc
                    LiqValorConvertidoTaxa = VlrConvertido
                End If

                If Enc.Codigo = "TARIFA PEDAGIO" Then
                    EncTarifaPedagio = Enc
                    LiqValorConvertidoTarifa = VlrConvertido
                End If

                Dim LoteSeq As String = ""
                '****************************************************************************************************************************
                '**** CASO O ENCARGO LIQUIDO ELE VAI CONTABILIZAR PELA CONTA DA OPERACAO ****************************************************
                '****************************************************************************************************************************
                If Enc.Codigo = "LIQUIDO" And Enc.Valor > 0 Then
                    Dim prod As NotaFiscalXItemXEncargo = Item.Encargos.Where(Function(s) s.Codigo.Trim.Contains("PRODUTO")).FirstOrDefault()

                    i += 1
                    If pLote > 0 Then
                        LoteSeq = pLote & ", @SequenciaServidor + " & i
                    ElseIf Item.CodigoOperacao > 68 Then

                        If Me.NotaFiscalRazao.CodigoTipoDeDocumento = 1 Then
                            LoteSeq = "0010, @SequenciaServidor + " & i  'LOTE / SEQUENCIA
                        Else
                            LoteSeq = "0009, @SequenciaServidor + " & i  'LOTE / SEQUENCIA
                        End If

                    Else
                        LoteSeq = "0010, @SequenciaServidor + " & i  'LOTE / SEQUENCIA
                    End If

                    Sql &= " INSERT INTO Razao (Empresa_Id, EndEmpresa_Id, Conta_Id, Cliente_Id, EndCliente_Id, Movimento_Id, Lote_Id, Sequencia_Id, Titulo, UnidadeDeNegocio, Indexador, DataMoeda, " & vbCrLf & _
                           "                   DebitoOficial, CreditoOficial, DebitoMoeda, CreditoMoeda," & vbCrLf & _
                           "                   Historico, Cliente_Nf, EndCliente_Nf, EntradaSaida_Nf, Serie_Nf, Numero_Nf, Produto_NF, CFOP_NF, Sequencia_NF, Encargo_NF," & vbCrLf & _
                           "                   Pedido, PrevistoRealizado, Custo, Produto, Rateado, Deposito, EndDeposito, UsuarioInclusao, UsuarioInclusaoData)" & vbCrLf & _
                           "Values ('" & IIf(String.IsNullOrWhiteSpace(pEmpresa), Me.NotaFiscalRazao.CodigoEmpresa, pEmpresa) & "'" & vbCrLf & _
                           ", " & IIf(String.IsNullOrWhiteSpace(pEndereco), Me.NotaFiscalRazao.EnderecoEmpresa, pEndereco) & vbCrLf & _
                           ",'" & Item.SubOperacao.CodigoGrupoContas & "'" & vbCrLf

                    Dim contaLiquida = New Negocio.PlanoDeConta("", 0, Item.SubOperacao.CodigoGrupoContas)

                    If contaLiquida.TemCliente Then
                        If Me.NotaFiscalRazao.SubOperacao.ProprietarioDaMercadoria Then
                            Sql &= ",'" & Me.NotaFiscalRazao.CodigoProprietarioDaMercadoria & "'" & vbCrLf & _
                                   ", " & Me.NotaFiscalRazao.EnderecoProprietarioDaMercadoria & vbCrLf
                        Else
                            Sql &= ",'" & Me.NotaFiscalRazao.CodigoCliente & "'" & vbCrLf & _
                                   ", " & Me.NotaFiscalRazao.EnderecoCliente & vbCrLf
                        End If
                    Else
                        Sql &= ",''" & vbCrLf & _
                                ",0" & vbCrLf
                    End If

                    Sql &= ",'" & CDate(Me.NotaFiscalRazao.Movimento).ToString("yyyy/MM/dd") & "'" & vbCrLf & _
                           ", " & LoteSeq & vbCrLf & _
                           ", NULL" & vbCrLf

                    If Me.NotaFiscalRazao.CodigoPedido = 0 Then
                        Sql &= ", '" & Me.NotaFiscalRazao.CodigoUnidadeDeNegocio & "' , 2 "
                    Else
                        Sql &= ", '" & Me.NotaFiscalRazao.Pedido.CodigoUnidadeNegocio & "'" & "," & Me.NotaFiscalRazao.Pedido.CodigoIndexador
                    End If

                    Sql &= ",'" & CDate(Me.NotaFiscalRazao.Movimento).ToString("yyyy/MM/dd") & "'" & vbCrLf

                    If prod IsNot Nothing AndAlso String.IsNullOrWhiteSpace(prod.ContaDeCredito) Then
                        'VALOR DÉBITO OFICIAL, CREDITO OFICIAL, DEBITO MOEDA, CREDITO MOEDA
                        Sql &= ",0.0" & ", " & Str(Enc.Valor) & ", 0.0" & ", " & Str(VlrConvertido)
                    Else
                        Sql &= "," & Str(Enc.Valor) & ",0.0 ," & Str(VlrConvertido) & ", 0.0"
                    End If

                    Dim Historico As String = String.Empty

                    If Me.NotaFiscalRazao.TipoDeDocumento.Historico = "NF" Then
                        Historico = "NF "
                    ElseIf Me.NotaFiscalRazao.TipoDeDocumento.Historico = "CTE" Then
                        Historico = "CTE "
                    ElseIf Me.NotaFiscalRazao.TipoDeDocumento.Historico = "CUPOM" Then
                        Historico = "CUPOM "
                    ElseIf Me.NotaFiscalRazao.TipoDeDocumento.Historico = "RECIBO" Then
                        Historico = "RECIBO "
                    ElseIf Me.NotaFiscalRazao.TipoDeDocumento.Historico = "FATURA" Then
                        Historico = "FATURA "
                    ElseIf Me.NotaFiscalRazao.TipoDeDocumento.Historico = "ESTADIA" Then
                        Historico = "ESTADIA "
                    ElseIf Me.NotaFiscalRazao.TipoDeDocumento.Historico = "RPA" Then
                        Historico = "RPA "
                    ElseIf Me.NotaFiscalRazao.TipoDeDocumento.Historico = "REC" Then
                        Historico = "REC "
                    ElseIf Me.NotaFiscalRazao.TipoDeDocumento.Historico = "CTR" Then
                        Historico = "CTR "
                    ElseIf Me.NotaFiscalRazao.TipoDeDocumento.Historico = "NFC" Then
                        Historico = "NFC "
                    Else
                        Historico = "DOCUMENTO "
                    End If

                    Historico &= Me.NotaFiscalRazao.Codigo & "-" & Me.NotaFiscalRazao.Serie & " OP " & Item.CodigoOperacao & "-" & Item.CodigoSubOperacao &
                        IIf(Me.NotaFiscalRazao.CodigoPedido > 0, " Pedido " & Me.NotaFiscalRazao.CodigoPedido, " ") & " " & Me.NotaFiscalRazao.Cliente.Nome

                    '06-09-2013 - FURLAN - CASO APAREÇA PROBLEMAS PELO HISTÓRICO NÃO ESTAR NA CONTABILIDADE POR OPERAÇÃO 69 E 70, REVERMOS.
                    'If Item.CodigoOperacao > 68 AndAlso Me.NotaFiscalRazao.Observacoes.Length > 0 Then Historico &= ". " & Me.NotaFiscalRazao.Observacoes

                    If Me.NotaFiscalRazao.CodigoTipoDeDocumento = 57 OrElse Me.NotaFiscalRazao.CodigoTipoDeDocumento = 58 Then
                        'NAO FAZ NADA
                    Else
                        If Item.CodigoOperacao > 70 AndAlso Me.NotaFiscalRazao.Observacoes.Length > 0 Then Historico &= ". " & Me.NotaFiscalRazao.Observacoes
                    End If

                    If Me.NotaFiscalRazao.NFG AndAlso Not String.IsNullOrWhiteSpace(Item.ObservacoesDoProduto) Then
                        Historico &= ". " & Item.ObservacoesDoProduto
                    End If

                    Sql &= ",'" & Historico & "'" & vbCrLf & _
                           ",'" & Me.NotaFiscalRazao.CodigoCliente & "'" & vbCrLf & _
                           ", " & Me.NotaFiscalRazao.EnderecoCliente & vbCrLf & _
                           ",'" & Me.NotaFiscalRazao.EntradaSaida.ToString.Substring(0, 1) & "'" & vbCrLf & _
                           ",'" & Me.NotaFiscalRazao.Serie & "'" & vbCrLf & _
                           ", " & Me.NotaFiscalRazao.Codigo & vbCrLf & _
                           ",'" & Item.CodigoProduto & "'" & vbCrLf & _
                           ", " & Item.CFOP & vbCrLf & _
                           ", " & Item.Sequencia & vbCrLf & _
                           ",'" & Enc.Codigo & "'" & vbCrLf & _
                           ", " & Me.NotaFiscalRazao.CodigoPedido.ToSqlNULL & vbCrLf & _
                           ",'P'" & vbCrLf & _
                           ",'" & Enc.CentroDeCusto & "'" & vbCrLf & _
                           ", " & IIf(New PlanoDeConta("", 0, Item.SubOperacao.CodigoGrupoContas).TemProduto, Item.CodigoProduto, "NULL") & vbCrLf & _
                           ", " & IIf(Item.Rateado, "1", "0") & vbCrLf

                    If Item.SubOperacao.EntradaSaida = eEntradaSaida.Saida And Item.SubOperacao.Deposito Then
                        Sql &= ",'" & Me.NotaFiscalRazao.CodigoDestino & "'," & Me.NotaFiscalRazao.EnderecoDestino & vbCrLf
                    Else
                        Sql &= ",'" & Me.NotaFiscalRazao.CodigoDeposito & "'," & _NotaFiscalRazao.EnderecoDeposito & vbCrLf
                    End If

                    Sql &= ",'" & HttpContext.Current.Session("ssNomeUsuario") & "', Convert(varchar, CURRENT_TIMESTAMP, 112))" & vbCrLf
                Else
                    '****************************************************************************************************************************
                    '**** CONTABILIZA TODOS OS ENCARGOS QUE TEM CONTAS INFORMADAS NA OPERACAOXENCARGOS ******************************************
                    '****************************************************************************************************************************
                    '*************************
                    '******** DEBITO *********
                    '*************************
                    If Not String.IsNullOrEmpty(Enc.OperacaoEncargo.CodigoDebitaConta.Trim) And Enc.Valor > 0 Then
                        i += 1
                        If pLote > 0 Then
                            LoteSeq = "," & pLote & ", @SequenciaServidor + " & i
                        ElseIf Item.CodigoOperacao > 68 Then

                            If Me.NotaFiscalRazao.CodigoTipoDeDocumento = 1 Then
                                LoteSeq = ", 0010, @SequenciaServidor + " & i  'LOTE / SEQUENCIA
                            Else
                                LoteSeq = ", 0009, @SequenciaServidor + " & i  'LOTE / SEQUENCIA
                            End If

                        Else
                            LoteSeq = ", 0010, @SequenciaServidor + " & i  'LOTE / SEQUENCIA
                        End If

                        Sql &= " INSERT INTO Razao (Empresa_Id, EndEmpresa_Id, Conta_Id, Cliente_Id, EndCliente_Id, Movimento_Id, Lote_Id, Sequencia_Id, Titulo, UnidadeDeNegocio, Indexador, DataMoeda, " & vbCrLf & _
                               "                   DebitoOficial, CreditoOficial, DebitoMoeda, CreditoMoeda," & vbCrLf & _
                               "                   Historico, Cliente_Nf, EndCliente_Nf, EntradaSaida_Nf, Serie_Nf, Numero_Nf, Produto_NF, CFOP_NF, Sequencia_NF, Encargo_NF," & vbCrLf & _
                               "                   Pedido, PrevistoRealizado, Custo, Produto, Rateado, Deposito, EndDeposito, UsuarioInclusao, UsuarioInclusaoData)" & vbCrLf & _
                               " values('" & IIf(String.IsNullOrWhiteSpace(pEmpresa), Me.NotaFiscalRazao.CodigoEmpresa, pEmpresa) & "'" & vbCrLf & _
                               ", " & IIf(String.IsNullOrWhiteSpace(pEndereco), Me.NotaFiscalRazao.EnderecoEmpresa, pEndereco) & vbCrLf & _
                               ",'" & Enc.OperacaoEncargo.CodigoDebitaConta & "'" & vbCrLf

                        If Enc.OperacaoEncargo.CodigoDebitaConta.Length = 7 Then
                            If Me.NotaFiscalRazao.SubOperacao.ProprietarioDaMercadoria Then
                                Sql &= ",'" & Me.NotaFiscalRazao.CodigoProprietarioDaMercadoria & "'" & vbCrLf & _
                                          ", " & Me.NotaFiscalRazao.EnderecoProprietarioDaMercadoria & vbCrLf
                            Else
                                Sql &= ",'" & Me.NotaFiscalRazao.CodigoCliente & "'" & vbCrLf & _
                                          ", " & Me.NotaFiscalRazao.EnderecoCliente & vbCrLf
                            End If
                        Else
                            Sql &= ",''" & vbCrLf & _
                                      ",0 " & vbCrLf
                        End If

                        Sql &= ",'" & CDate(Me.NotaFiscalRazao.Movimento).ToString("yyyy/MM/dd") & "'" & vbCrLf & _
                                  LoteSeq & vbCrLf & _
                                  ", NULL" & vbCrLf
                        If Me.NotaFiscalRazao.CodigoPedido = 0 Then
                            Sql &= ", '" & Me.NotaFiscalRazao.CodigoUnidadeDeNegocio & "' , 2 "
                        Else
                            Sql &= ", '" & Me.NotaFiscalRazao.Pedido.CodigoUnidadeNegocio & "'" & "," & Me.NotaFiscalRazao.Pedido.CodigoIndexador
                        End If
                        Sql &= ",'" & CDate(Me.NotaFiscalRazao.Movimento).ToString("yyyy/MM/dd") & "'" & vbCrLf & _
                                  ", " & IIf(Enc.ValorNovo > 0, Str(Enc.ValorNovo), Str(Enc.Valor)) & ",0.0 ," & Str(VlrConvertido) & ", 0.0"

                        Dim Historico As String = Enc.Codigo & " OP " & Item.CodigoOperacao & "-" & Item.CodigoSubOperacao

                        If Me.NotaFiscalRazao.TipoDeDocumento.Historico = "NF" Then
                            Historico &= " NF "
                        ElseIf Me.NotaFiscalRazao.TipoDeDocumento.Historico = "CTE" Then
                            Historico &= " CTE "
                        ElseIf Me.NotaFiscalRazao.TipoDeDocumento.Historico = "CUPOM" Then
                            Historico &= " CUPOM "
                        ElseIf Me.NotaFiscalRazao.TipoDeDocumento.Historico = "RECIBO" Then
                            Historico &= " RECIBO "
                        ElseIf Me.NotaFiscalRazao.TipoDeDocumento.Historico = "FATURA" Then
                            Historico &= " FATURA "
                        ElseIf Me.NotaFiscalRazao.TipoDeDocumento.Historico = "ESTADIA" Then
                            Historico &= " ESTADIA "
                        ElseIf Me.NotaFiscalRazao.TipoDeDocumento.Historico = "RPA" Then
                            Historico &= " RPA "
                        ElseIf Me.NotaFiscalRazao.TipoDeDocumento.Historico = "REC" Then
                            Historico &= " REC "
                        ElseIf Me.NotaFiscalRazao.TipoDeDocumento.Historico = "CTR" Then
                            Historico &= " CTR "
                        ElseIf Me.NotaFiscalRazao.TipoDeDocumento.Historico = "NFC" Then
                            Historico &= " NFC "
                        Else
                            Historico &= " DOCUMENTO "
                        End If

                        Historico &= Me.NotaFiscalRazao.Codigo & "-" & Me.NotaFiscalRazao.Serie &
                            IIf(Me.NotaFiscalRazao.CodigoPedido > 0, " Pedido " & Me.NotaFiscalRazao.CodigoPedido, " ") & " " & Me.NotaFiscalRazao.Cliente.Nome

                        If Me.NotaFiscalRazao.CodigoTipoDeDocumento = 57 OrElse Me.NotaFiscalRazao.CodigoTipoDeDocumento = 58 Then
                            'NAO FAZ NADA
                        Else
                            If (Enc.Codigo = "PRODUTO" Or Enc.Codigo = "LIQUIDO") _
                                AndAlso Item.CodigoOperacao > 68 _
                                AndAlso Me.NotaFiscalRazao.Observacoes.Length > 0 Then
                                Historico &= ". " & Me.NotaFiscalRazao.Observacoes
                            End If
                        End If

                        If Me.NotaFiscalRazao.NFG AndAlso Not String.IsNullOrWhiteSpace(Item.ObservacoesDoProduto) Then
                            Historico &= ". " & Item.ObservacoesDoProduto
                        End If

                        Sql &= ",'" & Historico & " '" & vbCrLf & _
                                  ",'" & Me.NotaFiscalRazao.CodigoCliente & "'" & vbCrLf & _
                                  ", " & Me.NotaFiscalRazao.EnderecoCliente & vbCrLf & _
                                  ",'" & Me.NotaFiscalRazao.EntradaSaida.ToString.Substring(0, 1) & "'" & vbCrLf & _
                                  ",'" & Me.NotaFiscalRazao.Serie & "'" & vbCrLf & _
                                  ", " & Me.NotaFiscalRazao.Codigo & vbCrLf & _
                                  ",'" & Item.CodigoProduto & "'" & vbCrLf & _
                                  ", " & Item.CFOP & vbCrLf & _
                                  ", " & Item.Sequencia & vbCrLf & _
                                  ",'" & Enc.Codigo & "'" & vbCrLf & _
                                  ", " & Me.NotaFiscalRazao.CodigoPedido.ToSqlNULL & vbCrLf & _
                                  ",'P'" & vbCrLf & _
                                  ",'" & Enc.CentroDeCusto & "'" & vbCrLf

                        If Enc.Codigo = "PRODUTO" AndAlso Item.Encargos.EncProduto.GrupoDeContaDebito.ProdutoParaCusto AndAlso Not Item.CodigoProdutoCusto Is Nothing AndAlso Item.CodigoProdutoCusto.Length > 0 Then
                            Sql &= ",'" & Item.CodigoProdutoCusto & "'" & vbCrLf
                        Else
                            Sql &= "," & IIf(New Negocio.PlanoDeConta("", 0, Enc.OperacaoEncargo.CodigoDebitaConta).TemProduto, Item.CodigoProduto, "NULL") & vbCrLf
                        End If

                        Sql &= "," & IIf(Item.Rateado, "1", "0") & vbCrLf

                        If Item.SubOperacao.EntradaSaida = eEntradaSaida.Saida And Item.SubOperacao.Deposito Then
                            Sql &= ",'" & Me.NotaFiscalRazao.CodigoDestino & "'," & Me.NotaFiscalRazao.EnderecoDestino & vbCrLf
                        Else
                            Sql &= ",'" & Me.NotaFiscalRazao.CodigoDeposito & "'," & Me.NotaFiscalRazao.EnderecoDeposito & vbCrLf
                        End If

                        Sql &= ",'" & HttpContext.Current.Session("ssNomeUsuario") & "', Convert(varchar, CURRENT_TIMESTAMP, 112))" & vbCrLf
                    End If

                    '*************************
                    '******** CREDITO ********
                    '*************************
                    If Not String.IsNullOrEmpty(Enc.OperacaoEncargo.CodigoCreditaConta.Trim) And Enc.Valor > 0 Then
                        i += 1
                        If pLote > 0 Then
                            LoteSeq = "," & pLote & ", @SequenciaServidor + " & i
                        ElseIf Item.CodigoOperacao > 68 Then

                            If Me.NotaFiscalRazao.CodigoTipoDeDocumento = 1 Then
                                LoteSeq = ", 0010, @SequenciaServidor + " & i  'LOTE / SEQUENCIA
                            Else
                                LoteSeq = ", 0009, @SequenciaServidor + " & i  'LOTE / SEQUENCIA
                            End If

                        Else
                            LoteSeq = ", 0010, @SequenciaServidor + " & i  'LOTE / SEQUENCIA
                        End If

                        Sql &= "INSERT INTO Razao (Empresa_Id, EndEmpresa_Id, Conta_Id, Cliente_Id, EndCliente_Id, Movimento_Id, Lote_Id, Sequencia_Id, Titulo, UnidadeDeNegocio, Indexador, DataMoeda, " & vbCrLf & _
                                 "                   DebitoOficial, CreditoOficial, DebitoMoeda, CreditoMoeda," & vbCrLf & _
                                 "                   Historico, Cliente_Nf, EndCliente_Nf, EntradaSaida_Nf, Serie_Nf, Numero_Nf, Produto_NF, CFOP_NF, Sequencia_NF, Encargo_NF," & vbCrLf & _
                                 "                   Pedido, PrevistoRealizado, Custo, Produto, Rateado, Deposito, EndDeposito, UsuarioInclusao, UsuarioInclusaoData)" & vbCrLf & _
                                 " values('" & IIf(String.IsNullOrWhiteSpace(pEmpresa), Me.NotaFiscalRazao.CodigoEmpresa, pEmpresa) & "'" & vbCrLf & _
                                 ", " & IIf(String.IsNullOrWhiteSpace(pEndereco), Me.NotaFiscalRazao.EnderecoEmpresa, pEndereco) & vbCrLf & _
                                 ",'" & Enc.OperacaoEncargo.CodigoCreditaConta & "'" & vbCrLf

                        If Enc.OperacaoEncargo.CodigoCreditaConta.Length = 7 Then
                            If Me.NotaFiscalRazao.SubOperacao.ProprietarioDaMercadoria Then
                                Sql &= ",'" & Me.NotaFiscalRazao.CodigoProprietarioDaMercadoria & "'" & vbCrLf & _
                                          ", " & Me.NotaFiscalRazao.EnderecoProprietarioDaMercadoria & vbCrLf
                            Else
                                Sql &= ",'" & Me.NotaFiscalRazao.CodigoCliente & "'" & vbCrLf & _
                                          ", " & Me.NotaFiscalRazao.EnderecoCliente & vbCrLf
                            End If
                        Else
                            Sql &= ",''" & vbCrLf & _
                                      ",0 " & vbCrLf
                        End If

                        Sql &= ",'" & CDate(Me.NotaFiscalRazao.Movimento).ToString("yyyy/MM/dd") & "'" & vbCrLf & _
                                  LoteSeq & vbCrLf & _
                                  ",NULL" & vbCrLf
                        If Me.NotaFiscalRazao.CodigoPedido = 0 Then
                            Sql &= ", '" & Me.NotaFiscalRazao.CodigoUnidadeDeNegocio & "' , 2 "
                        Else
                            Sql &= ", '" & Me.NotaFiscalRazao.Pedido.CodigoUnidadeNegocio & "'" & "," & Me.NotaFiscalRazao.Pedido.CodigoIndexador
                        End If

                        Sql &= ",'" & CDate(Me.NotaFiscalRazao.Movimento).ToString("yyyy/MM/dd") & "'" & vbCrLf &
                               ",0.0, " & IIf(Enc.ValorNovo > 0, Str(Enc.ValorNovo), Str(Enc.Valor)) & ", 0.0," & Str(VlrConvertido) & vbCrLf

                        Dim Historico As String = Enc.Codigo & " OP " & Item.CodigoOperacao & "-" & Item.CodigoSubOperacao

                        If Me.NotaFiscalRazao.TipoDeDocumento.Historico = "NF" Then
                            Historico &= " NF "
                        ElseIf Me.NotaFiscalRazao.TipoDeDocumento.Historico = "CTE" Then
                            Historico &= " CTE "
                        ElseIf Me.NotaFiscalRazao.TipoDeDocumento.Historico = "CUPOM" Then
                            Historico &= " CUPOM "
                        ElseIf Me.NotaFiscalRazao.TipoDeDocumento.Historico = "RECIBO" Then
                            Historico &= " RECIBO "
                        ElseIf Me.NotaFiscalRazao.TipoDeDocumento.Historico = "FATURA" Then
                            Historico &= " FATURA "
                        ElseIf Me.NotaFiscalRazao.TipoDeDocumento.Historico = "ESTADIA" Then
                            Historico &= " ESTADIA "
                        ElseIf Me.NotaFiscalRazao.TipoDeDocumento.Historico = "RPA" Then
                            Historico &= " RPA "
                        ElseIf Me.NotaFiscalRazao.TipoDeDocumento.Historico = "REC" Then
                            Historico &= " REC "
                        ElseIf Me.NotaFiscalRazao.TipoDeDocumento.Historico = "CTR" Then
                            Historico &= " CTR "
                        ElseIf Me.NotaFiscalRazao.TipoDeDocumento.Historico = "NFC" Then
                            Historico &= " NFC "
                        Else
                            Historico &= " DOCUMENTO "
                        End If

                        Historico &= Me.NotaFiscalRazao.Codigo & "-" & Me.NotaFiscalRazao.Serie &
                            IIf(Me.NotaFiscalRazao.CodigoPedido > 0, " Pedido " & Me.NotaFiscalRazao.CodigoPedido, "") & " - " & Me.NotaFiscalRazao.Cliente.Nome

                        If Me.NotaFiscalRazao.CodigoTipoDeDocumento = 57 OrElse Me.NotaFiscalRazao.CodigoTipoDeDocumento = 58 Then
                            'NAO FAZ NADA
                        Else
                            If Enc.Codigo = "PRODUTO" _
                                AndAlso Item.CodigoOperacao > 68 _
                                AndAlso Me.NotaFiscalRazao.Observacoes.Length > 0 Then
                                Historico &= ". " & Me.NotaFiscalRazao.Observacoes
                            End If
                        End If

                        If Me.NotaFiscalRazao.NFG AndAlso Not String.IsNullOrWhiteSpace(Item.ObservacoesDoProduto) Then
                            Historico &= ". " & Item.ObservacoesDoProduto
                        End If

                        Sql &= ",'" & Historico & " '" & vbCrLf & _
                                  ",'" & Me.NotaFiscalRazao.CodigoCliente & "'" & vbCrLf & _
                                  ", " & Me.NotaFiscalRazao.EnderecoCliente & vbCrLf & _
                                  ",'" & Me.NotaFiscalRazao.EntradaSaida.ToString.Substring(0, 1) & "'" & vbCrLf & _
                                  ",'" & Me.NotaFiscalRazao.Serie & "'" & vbCrLf & _
                                  ", " & Me.NotaFiscalRazao.Codigo & vbCrLf & _
                                  ",'" & Item.CodigoProduto & "'" & vbCrLf & _
                                  ", " & Item.CFOP & vbCrLf & _
                                  ", " & Item.Sequencia & vbCrLf & _
                                  ",'" & Enc.Codigo & "'" & vbCrLf & _
                                  ", " & Me.NotaFiscalRazao.CodigoPedido.ToSqlNULL & vbCrLf & _
                                  ",'P'" & vbCrLf & _
                                  ",'" & Enc.CentroDeCusto & "'" & vbCrLf

                        If Enc.Codigo = "PRODUTO" AndAlso Item.Encargos.EncProduto.GrupoDeContaCredito.ProdutoParaCusto AndAlso Not Item.CodigoProdutoCusto Is Nothing AndAlso Item.CodigoProdutoCusto.Length > 0 Then
                            Sql &= ",'" & Item.CodigoProdutoCusto & "'" & vbCrLf
                        Else
                            Sql &= "," & IIf(New Negocio.PlanoDeConta("", 0, Enc.OperacaoEncargo.CodigoCreditaConta).TemProduto, Item.CodigoProduto, "NULL") & vbCrLf
                        End If

                        Sql &= "," & IIf(Item.Rateado, "1", "0") & vbCrLf

                        If Item.SubOperacao.EntradaSaida = eEntradaSaida.Saida And Item.SubOperacao.Deposito Then
                            Sql &= ",'" & Me.NotaFiscalRazao.CodigoDestino & "'," & Me.NotaFiscalRazao.EnderecoDestino & vbCrLf
                        Else
                            Sql &= ",'" & Me.NotaFiscalRazao.CodigoDeposito & "'," & Me.NotaFiscalRazao.EnderecoDeposito & vbCrLf
                        End If

                        Sql &= ",'" & HttpContext.Current.Session("ssNomeUsuario") & "', Convert(varchar, CURRENT_TIMESTAMP, 112))" & vbCrLf
                    End If
                End If
            Next

            If NotaFiscalRazao.TipoDeDocumentoFrete = eTipoDeDocumentoFrete.Circulacao Then
                Dim index As Integer = 0
                If LiqValorConvertidoTaxa > 0 Then TransfereciaEntreContasPedagio(Item, EncTaxaPedagio, "D", EncTaxaPedagio.OperacaoEncargo.CodigoDebitaConta, LiqValorConvertidoTaxa, Sql, index)
                If LiqValorConvertidoTarifa > 0 Then TransfereciaEntreContasPedagio(Item, EncTarifaPedagio, "D", EncTarifaPedagio.OperacaoEncargo.CodigoDebitaConta, LiqValorConvertidoTarifa, Sql, index)
            Else
                If Transferecia Then TransfereciaEntreContasCliFor(Item, Sql, SeqTransf)
            End If
        Next

        Sqls.Add(Sql)
    End Sub

    Private Sub TransfereciaEntreContasCliFor(ByVal Item As Negocio.NotaFiscalXItem, ByRef sql As String, ByRef i As Integer)

        If Item.NotaFiscal.SubOperacao.Classe = eClassesOperacoes.GLOBAL Then
            TransfereciaEncargo(Item.Encargos.EncProduto, sql, i)
        End If

        TransfereciaEncargo(Item.Encargos.EncLiquido, sql, i)
    End Sub

    Private Sub TransfereciaEncargo(ByVal Encargo As Negocio.NotaFiscalXItemXEncargo, ByRef sql As String, ByRef i As Integer)
        Dim DC As String = ""

        If Encargo.SubOperacao.EntradaSaida = eEntradaSaida.Entrada Then
            DC = IIf(Encargo.Codigo = "PRODUTO", "D", "C")
        Else
            DC = IIf(Encargo.Codigo = "PRODUTO", "C", "D")
        End If

        Dim Conta As String = ""
        Conta = IIf(Encargo.Codigo = "PRODUTO", IIf(DC = "C", Encargo.ContaDeCredito, Encargo.ContaDeDebito), Encargo.SubOperacao.CodigoGrupoContas)


        '**************************************************
        '************* ESTORNA EMPRESA DA NOTA ************
        '**************************************************
        i += 1
        sql &= "INSERT INTO Razao (Empresa_Id, EndEmpresa_Id, Conta_Id, Cliente_Id, EndCliente_Id, Movimento_Id, Lote_Id, Sequencia_Id, Titulo, UnidadeDeNegocio, Indexador, DataMoeda, " & vbCrLf & _
               "                   DebitoOficial, CreditoOficial, DebitoMoeda, CreditoMoeda," & vbCrLf & _
               "                   Historico, Cliente_Nf, EndCliente_Nf, EntradaSaida_Nf, Serie_Nf, Numero_Nf, Produto_NF, CFOP_NF, Sequencia_NF, Encargo_NF," & vbCrLf & _
               "                   Pedido, PrevistoRealizado, Custo, Produto, Rateado, Deposito, EndDeposito, UsuarioInclusao, UsuarioInclusaoData)" & vbCrLf & _
               " values('" & Me.NotaFiscalRazao.CodigoEmpresa & "'" & vbCrLf & _
               ", " & Me.NotaFiscalRazao.EnderecoEmpresa & vbCrLf & _
               ",'" & Conta & "'" & vbCrLf & _
               ",'" & Me.NotaFiscalRazao.CodigoCliente & "'" & vbCrLf & _
               ", " & Me.NotaFiscalRazao.EnderecoCliente & vbCrLf & _
               ",'" & CDate(Me.NotaFiscalRazao.Movimento).ToString("yyyy/MM/dd") & "'" & vbCrLf & _
               ",11" & vbCrLf & _
               ", @SequenciaTransferenciaCliente + " & i & vbCrLf & _
               ",NULL" & vbCrLf
        If Me.NotaFiscalRazao.CodigoPedido = 0 Then
            sql &= ", '" & Me.NotaFiscalRazao.CodigoUnidadeDeNegocio & "' , 2 "
        Else
            sql &= ", '" & Me.NotaFiscalRazao.Pedido.CodigoUnidadeNegocio & "'" & "," & Me.NotaFiscalRazao.Pedido.CodigoIndexador
        End If
        sql &= ",'" & CDate(Me.NotaFiscalRazao.Movimento).ToString("yyyy/MM/dd") & "'" & vbCrLf

        If DC = "D" Then
            sql &= ",0 ," & Str(Encargo.Valor) & ", 0, 0" & vbCrLf
        Else
            sql &= "," & Str(Encargo.Valor) & ", 0, 0, 0" & vbCrLf
        End If

        Dim Historico As String = "ESTORNO P/ TRANSF. ENTRE CONTAS CLIFOR - " & Encargo.Codigo & " OP " & Encargo.ItemNotaFiscal.CodigoOperacao & "-" & Encargo.ItemNotaFiscal.CodigoSubOperacao

        If Me.NotaFiscalRazao.TipoDeDocumento.Historico = "NF" Then
            Historico &= " NF "
        ElseIf Me.NotaFiscalRazao.TipoDeDocumento.Historico = "CTE" Then
            Historico &= " CTE "
        ElseIf Me.NotaFiscalRazao.TipoDeDocumento.Historico = "CUPOM" Then
            Historico &= " CUPOM "
        ElseIf Me.NotaFiscalRazao.TipoDeDocumento.Historico = "RECIBO" Then
            Historico &= " RECIBO "
        ElseIf Me.NotaFiscalRazao.TipoDeDocumento.Historico = "FATURA" Then
            Historico &= " FATURA "
        ElseIf Me.NotaFiscalRazao.TipoDeDocumento.Historico = "ESTADIA" Then
            Historico &= " ESTADIA "
        ElseIf Me.NotaFiscalRazao.TipoDeDocumento.Historico = "RPA" Then
            Historico &= " RPA "
        ElseIf Me.NotaFiscalRazao.TipoDeDocumento.Historico = "REC" Then
            Historico &= " REC "
        ElseIf Me.NotaFiscalRazao.TipoDeDocumento.Historico = "CTR" Then
            Historico &= " CTR "
        ElseIf Me.NotaFiscalRazao.TipoDeDocumento.Historico = "NFC" Then
            Historico &= " NFC "
        Else
            Historico &= " DOCUMENTO "
        End If

        Historico &= Me.NotaFiscalRazao.Codigo & "-" & Me.NotaFiscalRazao.Serie & " Pedido " & Me.NotaFiscalRazao.CodigoPedido & " - " & Me.NotaFiscalRazao.Cliente.Nome

        If Me.NotaFiscalRazao.CodigoTipoDeDocumento = 57 OrElse Me.NotaFiscalRazao.CodigoTipoDeDocumento = 58 Then
            'NAO FAZ NADA
        Else
            If Encargo.ItemNotaFiscal.CodigoOperacao > 68 AndAlso Me.NotaFiscalRazao.Observacoes.Length > 0 Then
                Historico &= ". " & Me.NotaFiscalRazao.Observacoes
            End If
        End If

        If Me.NotaFiscalRazao.NFG AndAlso Not String.IsNullOrWhiteSpace(Encargo.ItemNotaFiscal.ObservacoesDoProduto) Then
            Historico &= ". " & Encargo.ItemNotaFiscal.ObservacoesDoProduto
        End If

        sql &= ",'" & Historico & " '" & vbCrLf & _
                  ",'" & Me.NotaFiscalRazao.CodigoCliente & "'" & vbCrLf & _
                  ", " & Me.NotaFiscalRazao.EnderecoCliente & vbCrLf & _
                  ",'" & Me.NotaFiscalRazao.EntradaSaida.ToString.Substring(0, 1) & "'" & vbCrLf & _
                  ",'" & Me.NotaFiscalRazao.Serie & "'" & vbCrLf & _
                  ", " & Me.NotaFiscalRazao.Codigo & vbCrLf & _
                  ",'" & Encargo.ItemNotaFiscal.CodigoProduto & "'" & vbCrLf & _
                  ", " & Encargo.ItemNotaFiscal.CFOP & vbCrLf & _
                  ", " & Encargo.ItemNotaFiscal.Sequencia & vbCrLf & _
                  ",'" & Encargo.Codigo & "'" & vbCrLf & _
                  ", " & Me.NotaFiscalRazao.CodigoPedido.ToSqlNULL & vbCrLf & _
                  ",'P'" & vbCrLf & _
                  ",'" & Encargo.CentroDeCusto & "'," & vbCrLf & _
                  IIf(New Negocio.PlanoDeConta("", 0, Conta).TemProduto, Encargo.ItemNotaFiscal.CodigoProduto, "NULL") & "," & vbCrLf & _
                  IIf(Encargo.ItemNotaFiscal.Rateado, "1", "0") & vbCrLf

        If Encargo.ItemNotaFiscal.SubOperacao.EntradaSaida = eEntradaSaida.Saida And Encargo.ItemNotaFiscal.SubOperacao.Deposito Then
            sql &= ",'" & Me.NotaFiscalRazao.CodigoDestino & "'," & Me.NotaFiscalRazao.EnderecoDestino & vbCrLf
        Else
            sql &= ",'" & Me.NotaFiscalRazao.CodigoDeposito & "'," & Me.NotaFiscalRazao.EnderecoDeposito & vbCrLf
        End If

        sql &= ",'" & HttpContext.Current.Session("ssNomeUsuario") & "', Convert(varchar, CURRENT_TIMESTAMP, 112))" & vbCrLf

        '**************************************************
        '******** Transfere para Cliente do Pedido ********
        '**************************************************
        i += 1
        sql &= "INSERT INTO Razao (Empresa_Id, EndEmpresa_Id, Conta_Id, Cliente_Id, EndCliente_Id, Movimento_Id, Lote_Id, Sequencia_Id, Titulo, UnidadeDeNegocio, Indexador, DataMoeda, " & vbCrLf & _
               "                   DebitoOficial, CreditoOficial, DebitoMoeda, CreditoMoeda," & vbCrLf & _
               "                   Historico, Cliente_Nf, EndCliente_Nf, EntradaSaida_Nf, Serie_Nf, Numero_Nf, Produto_NF, CFOP_NF, Sequencia_NF, Encargo_NF," & vbCrLf & _
               "                   Pedido, PrevistoRealizado, Custo, Produto, Rateado, Deposito, EndDeposito, UsuarioInclusao, UsuarioInclusaoData)" & vbCrLf & _
               " values('" & Me.NotaFiscalRazao.CodigoEmpresa & "'" & vbCrLf & _
               ", " & Me.NotaFiscalRazao.EnderecoEmpresa & vbCrLf & _
               ",'" & Conta & "'" & vbCrLf & _
               ",'" & Me.NotaFiscalRazao.Pedido.CodigoCliente & "'" & vbCrLf & _
               ", " & Me.NotaFiscalRazao.Pedido.EnderecoCliente & vbCrLf & _
               ",'" & CDate(Me.NotaFiscalRazao.Movimento).ToString("yyyy/MM/dd") & "'" & vbCrLf & _
               ",11" & vbCrLf & _
               ",  @SequenciaTransferenciaCliente + " & i & vbCrLf & _
               ",NULL" & vbCrLf
        If Me.NotaFiscalRazao.CodigoPedido = 0 Then
            sql &= ", '" & Me.NotaFiscalRazao.CodigoUnidadeDeNegocio & "' , 2 "
        Else
            sql &= ", '" & Me.NotaFiscalRazao.Pedido.CodigoUnidadeNegocio & "'" & "," & Me.NotaFiscalRazao.Pedido.CodigoIndexador
        End If

        sql &= ",'" & CDate(Me.NotaFiscalRazao.Movimento).ToString("yyyy/MM/dd") & "'" & vbCrLf

        If DC = "D" Then
            sql &= "," & Str(Encargo.Valor) & ",0 ,0 ,0" & vbCrLf
        Else
            sql &= ",0 , " & Str(Encargo.Valor) & ",0 ,0" & vbCrLf
        End If

        Historico = "LANCAM. DE TRANSF. ENTRE CONTAS CLIFOR - " & Encargo.Codigo & " OP " & Encargo.ItemNotaFiscal.CodigoOperacao & "-" & Encargo.ItemNotaFiscal.CodigoSubOperacao

        If Me.NotaFiscalRazao.TipoDeDocumento.Historico = "NF" Then
            Historico &= " NF "
        ElseIf Me.NotaFiscalRazao.TipoDeDocumento.Historico = "CTE" Then
            Historico &= " CTE "
        ElseIf Me.NotaFiscalRazao.TipoDeDocumento.Historico = "CUPOM" Then
            Historico &= " CUPOM "
        ElseIf Me.NotaFiscalRazao.TipoDeDocumento.Historico = "RECIBO" Then
            Historico &= " RECIBO "
        ElseIf Me.NotaFiscalRazao.TipoDeDocumento.Historico = "FATURA" Then
            Historico &= " FATURA "
        ElseIf Me.NotaFiscalRazao.TipoDeDocumento.Historico = "ESTADIA" Then
            Historico &= " ESTADIA "
        ElseIf Me.NotaFiscalRazao.TipoDeDocumento.Historico = "RPA" Then
            Historico &= " RPA "
        ElseIf Me.NotaFiscalRazao.TipoDeDocumento.Historico = "REC" Then
            Historico &= " REC "
        ElseIf Me.NotaFiscalRazao.TipoDeDocumento.Historico = "CTR" Then
            Historico &= " CTR "
        ElseIf Me.NotaFiscalRazao.TipoDeDocumento.Historico = "NFC" Then
            Historico &= " NFC "
        Else
            Historico &= " DOCUMENTO "
        End If

        Historico &= Me.NotaFiscalRazao.Codigo & "-" & Me.NotaFiscalRazao.Serie & " Pedido " & Me.NotaFiscalRazao.CodigoPedido & " - " & Me.NotaFiscalRazao.Cliente.Nome

        If Me.NotaFiscalRazao.CodigoTipoDeDocumento = 57 OrElse Me.NotaFiscalRazao.CodigoTipoDeDocumento = 58 Then
            'NAO FAZ NADA
        Else
            If Encargo.ItemNotaFiscal.CodigoOperacao > 68 AndAlso Me.NotaFiscalRazao.Observacoes.Length > 0 Then
                Historico &= ". " & Me.NotaFiscalRazao.Observacoes
            End If
        End If

        If Me.NotaFiscalRazao.NFG AndAlso Not String.IsNullOrWhiteSpace(Encargo.ItemNotaFiscal.ObservacoesDoProduto) Then
            Historico &= ". " & Encargo.ItemNotaFiscal.ObservacoesDoProduto
        End If

        sql &= ",'" & Historico & " '" & vbCrLf & _
               ",'" & Me.NotaFiscalRazao.CodigoCliente & "'" & vbCrLf & _
               ", " & Me.NotaFiscalRazao.EnderecoCliente & vbCrLf & _
               ",'" & Me.NotaFiscalRazao.EntradaSaida.ToString.Substring(0, 1) & "'" & vbCrLf & _
               ",'" & Me.NotaFiscalRazao.Serie & "'" & vbCrLf & _
               ", " & Me.NotaFiscalRazao.Codigo & vbCrLf & _
               ",'" & Encargo.ItemNotaFiscal.CodigoProduto & "'" & vbCrLf & _
               ", " & Encargo.ItemNotaFiscal.CFOP & vbCrLf & _
               ", " & Encargo.ItemNotaFiscal.Sequencia & vbCrLf & _
               ",'" & Encargo.Codigo & "'" & vbCrLf & _
               ", " & Me.NotaFiscalRazao.CodigoPedido.ToSqlNULL & vbCrLf & _
               ",'P'" & vbCrLf & _
               ",'" & Encargo.CentroDeCusto & "'," & vbCrLf & _
               IIf(New Negocio.PlanoDeConta("", 0, Conta).TemProduto, Encargo.ItemNotaFiscal.CodigoProduto, "NULL") & "," & vbCrLf & _
               IIf(Encargo.ItemNotaFiscal.Rateado, "1", "0") & vbCrLf

        If Encargo.ItemNotaFiscal.SubOperacao.EntradaSaida = eEntradaSaida.Saida And Encargo.ItemNotaFiscal.SubOperacao.Deposito Then
            sql &= ",'" & Me.NotaFiscalRazao.CodigoDestino & "'," & Me.NotaFiscalRazao.EnderecoDestino & vbCrLf
        Else
            sql &= ",'" & Me.NotaFiscalRazao.CodigoDeposito & "'," & Me.NotaFiscalRazao.EnderecoDeposito & vbCrLf
        End If

        sql &= ",'" & HttpContext.Current.Session("ssNomeUsuario") & "', Convert(varchar, CURRENT_TIMESTAMP, 112))" & vbCrLf
    End Sub

    Private Sub TransfereciaEntreContasPedagio(ByVal Item As Negocio.NotaFiscalXItem, ByVal Liquido As Negocio.NotaFiscalXItemXEncargo, ByVal LiqDC As String, ByVal LiqConta As String, ByVal LiqValorConvertido As Decimal, ByRef sql As String, ByRef i As Integer)
        '**************************************************
        '************* PRIMEIRA PERNA ************
        '**************************************************
        i += 1
        sql &= "INSERT INTO Razao (Empresa_Id, EndEmpresa_Id, Conta_Id, Cliente_Id, EndCliente_Id, Movimento_Id, Lote_Id, Sequencia_Id, Titulo, UnidadeDeNegocio, Indexador, DataMoeda, " & vbCrLf & _
               "                   DebitoOficial, CreditoOficial, DebitoMoeda, CreditoMoeda," & vbCrLf & _
               "                   Historico, Cliente_Nf, EndCliente_Nf, EntradaSaida_Nf, Serie_Nf, Numero_Nf, Produto_NF, CFOP_NF, Sequencia_NF, Encargo_NF," & vbCrLf & _
               "                   Pedido, PrevistoRealizado, Custo, Produto, Rateado, Deposito, EndDeposito, UsuarioInclusao, UsuarioInclusaoData)" & vbCrLf & _
               " values('" & Me.NotaFiscalRazao.CodigoEmpresa & "'" & vbCrLf & _
               ", " & Me.NotaFiscalRazao.EnderecoEmpresa & vbCrLf & _
               ",'" & "1030101" & "'" & vbCrLf & _
               ",'" & Me.NotaFiscalRazao.CodigoDeposito & "'" & vbCrLf & _
               ", " & Me.NotaFiscalRazao.EnderecoDeposito & vbCrLf & _
               ",'" & CDate(Me.NotaFiscalRazao.Movimento).ToString("yyyy/MM/dd") & "'" & vbCrLf & _
               ",11" & vbCrLf & _
               ", @SequenciaTransferenciaCliente + " & i & vbCrLf & _
               ",NULL" & vbCrLf
        If Me.NotaFiscalRazao.CodigoPedido = 0 Then
            sql &= ", '" & Me.NotaFiscalRazao.CodigoUnidadeDeNegocio & "' , 2 "
        Else
            sql &= ", '" & Me.NotaFiscalRazao.Pedido.CodigoUnidadeNegocio & "'" & "," & Me.NotaFiscalRazao.Pedido.CodigoIndexador
        End If
        'IIf(Me.NotaFiscalRazao.CodigoPedido = 0, ", ''", ", '" & Me.NotaFiscalRazao.Pedido.CodigoUnidadeNegocio & "'") & vbCrLf & _
        'IIf(Me.NotaFiscalRazao.CodigoPedido = 0, ", 2", "," & Me.NotaFiscalRazao.Pedido.CodigoIndexador) & vbCrLf & _
        sql &= ",'" & CDate(Me.NotaFiscalRazao.Movimento).ToString("yyyy/MM/dd") & "'" & vbCrLf

        If LiqDC = "D" Then
            sql &= ",0.0, " & Str(Liquido.Valor) & ", 0.0," & Str(LiqValorConvertido) & vbCrLf
        Else
            sql &= "," & Str(Liquido.Valor) & ", 0.0," & Str(LiqValorConvertido) & ",0.0" & vbCrLf
        End If

        Dim Historico As String = "ESTORNO P/ TRANSF. ENTRE CONTAS CLIFOR - " & Liquido.Codigo & " OP " & Item.CodigoOperacao & "-" & Item.CodigoSubOperacao & " NF " & Me.NotaFiscalRazao.Codigo & "-" & Me.NotaFiscalRazao.Serie & " Pedido " & Me.NotaFiscalRazao.CodigoPedido & " - " & Me.NotaFiscalRazao.Cliente.Nome

        If Me.NotaFiscalRazao.CodigoTipoDeDocumento = 57 OrElse Me.NotaFiscalRazao.CodigoTipoDeDocumento = 58 Then
            'NAO FAZ NADA
        Else
            If Item.CodigoOperacao > 68 AndAlso Me.NotaFiscalRazao.Observacoes.Length > 0 Then
                Historico &= ". " & Me.NotaFiscalRazao.Observacoes
            End If
        End If

        sql &= ",'" & Historico & " '" & vbCrLf & _
                  ",'" & Me.NotaFiscalRazao.CodigoCliente & "'" & vbCrLf & _
                  ", " & Me.NotaFiscalRazao.EnderecoCliente & vbCrLf & _
                  ",'" & Me.NotaFiscalRazao.EntradaSaida.ToString.Substring(0, 1) & "'" & vbCrLf & _
                  ",'" & Me.NotaFiscalRazao.Serie & "'" & vbCrLf & _
                  ", " & Me.NotaFiscalRazao.Codigo & vbCrLf & _
                  ",'" & Item.CodigoProduto & "'" & vbCrLf & _
                  ", " & Item.CFOP & vbCrLf & _
                  ", " & Item.Sequencia & vbCrLf & _
                  ",'" & Liquido.Codigo & "'" & vbCrLf & _
                  ", " & Me.NotaFiscalRazao.CodigoPedido.ToSqlNULL & vbCrLf & _
                  ",'P'" & vbCrLf & _
                  ",'" & Liquido.CentroDeCusto & "'," & vbCrLf & _
                  IIf(New Negocio.PlanoDeConta("", 0, Liquido.OperacaoEncargo.CodigoCreditaConta).TemProduto, Item.CodigoProduto, "NULL") & "," & vbCrLf & _
                  IIf(Item.Rateado, "1", "0") & vbCrLf

        If Item.SubOperacao.EntradaSaida = eEntradaSaida.Saida And Item.SubOperacao.Deposito Then
            sql &= ",'" & Me.NotaFiscalRazao.CodigoDestino & "'," & Me.NotaFiscalRazao.EnderecoDestino & vbCrLf
        Else
            sql &= ",'" & Me.NotaFiscalRazao.CodigoDeposito & "'," & Me.NotaFiscalRazao.EnderecoDeposito & vbCrLf
        End If

        sql &= ",'" & HttpContext.Current.Session("ssNomeUsuario") & "', Convert(varchar, CURRENT_TIMESTAMP, 112))" & vbCrLf

        '**************************************************
        '******** SEGUNDA PERNA ********
        '**************************************************
        i += 1
        sql &= "INSERT INTO Razao (Empresa_Id, EndEmpresa_Id, Conta_Id, Cliente_Id, EndCliente_Id, Movimento_Id, Lote_Id, Sequencia_Id, Titulo, UnidadeDeNegocio, Indexador, DataMoeda, " & vbCrLf & _
               "                   DebitoOficial, CreditoOficial, DebitoMoeda, CreditoMoeda," & vbCrLf & _
               "                   Historico, Cliente_Nf, EndCliente_Nf, EntradaSaida_Nf, Serie_Nf, Numero_Nf, Produto_NF, CFOP_NF, Sequencia_NF, Encargo_NF," & vbCrLf & _
               "                   Pedido, PrevistoRealizado, Custo, Produto, Rateado, Deposito, EndDeposito, UsuarioInclusao, UsuarioInclusaoData)" & vbCrLf & _
               " values('" & Me.NotaFiscalRazao.CodigoDeposito & "'" & vbCrLf & _
               ", " & Me.NotaFiscalRazao.EnderecoDeposito & vbCrLf & _
               ",'" & "1030101" & "'" & vbCrLf & _
               ",'" & Me.NotaFiscalRazao.CodigoEmpresa & "'" & vbCrLf & _
               ", " & Me.NotaFiscalRazao.EnderecoEmpresa & vbCrLf & _
               ",'" & CDate(Me.NotaFiscalRazao.Movimento).ToString("yyyy/MM/dd") & "'" & vbCrLf & _
               ",11" & vbCrLf & _
               ",  @SequenciaTransferenciaCliente + " & i & vbCrLf & _
               ",NULL" & vbCrLf
        If Me.NotaFiscalRazao.CodigoPedido = 0 Then
            sql &= ", '" & Me.NotaFiscalRazao.CodigoUnidadeDeNegocio & "' , 2 "
        Else
            sql &= ", '" & Me.NotaFiscalRazao.Pedido.CodigoUnidadeNegocio & "'" & "," & Me.NotaFiscalRazao.Pedido.CodigoIndexador
        End If
        'IIf(Me.NotaFiscalRazao.CodigoPedido = 0, ", ''", ", '" & Me.NotaFiscalRazao.Pedido.CodigoUnidadeNegocio & "'") & vbCrLf & _
        'IIf(Me.NotaFiscalRazao.CodigoPedido = 0, ", 2", "," & Me.NotaFiscalRazao.Pedido.CodigoIndexador) & vbCrLf & _
        sql &= ",'" & CDate(Me.NotaFiscalRazao.Movimento).ToString("yyyy/MM/dd") & "'" & vbCrLf

        If LiqDC = "D" Then
            sql &= "," & Str(Liquido.Valor) & ", 0.0," & Str(LiqValorConvertido) & ",0.0" & vbCrLf
        Else
            sql &= ",0.0, " & Str(Liquido.Valor) & ", 0.0," & Str(LiqValorConvertido) & vbCrLf
        End If

        Historico = "LANCAM. DE TRANSF. ENTRE CONTAS CLIFOR - " & Liquido.Codigo & " OP " & Item.CodigoOperacao & "-" & Item.CodigoSubOperacao & " NF " & Me.NotaFiscalRazao.Codigo & "-" & Me.NotaFiscalRazao.Serie & " Pedido " & Me.NotaFiscalRazao.CodigoPedido & " - " & Me.NotaFiscalRazao.Cliente.Nome

        If Me.NotaFiscalRazao.CodigoTipoDeDocumento = 57 OrElse Me.NotaFiscalRazao.CodigoTipoDeDocumento = 58 Then
            'NAO FAZ NADA
        Else
            If Item.CodigoOperacao > 68 AndAlso Me.NotaFiscalRazao.Observacoes.Length > 0 Then
                Historico &= ". " & Me.NotaFiscalRazao.Observacoes
            End If
        End If

        sql &= ",'" & Historico & " '" & vbCrLf & _
               ",'" & Me.NotaFiscalRazao.CodigoCliente & "'" & vbCrLf & _
               ", " & Me.NotaFiscalRazao.EnderecoCliente & vbCrLf & _
               ",'" & Me.NotaFiscalRazao.EntradaSaida.ToString.Substring(0, 1) & "'" & vbCrLf & _
               ",'" & Me.NotaFiscalRazao.Serie & "'" & vbCrLf & _
               ", " & Me.NotaFiscalRazao.Codigo & vbCrLf & _
               ",'" & Item.CodigoProduto & "'" & vbCrLf & _
               ", " & Item.CFOP & vbCrLf & _
               ", " & Item.Sequencia & vbCrLf & _
               ",'" & Liquido.Codigo & "'" & vbCrLf & _
               ", " & Me.NotaFiscalRazao.CodigoPedido.ToSqlNULL & vbCrLf & _
               ",'P'" & vbCrLf & _
               ",'" & Liquido.CentroDeCusto & "'," & vbCrLf & _
               IIf(New Negocio.PlanoDeConta("", 0, Liquido.OperacaoEncargo.CodigoCreditaConta).TemProduto, Item.CodigoProduto, "NULL") & "," & vbCrLf & _
               IIf(Item.Rateado, "1", "0") & vbCrLf

        If Item.SubOperacao.EntradaSaida = eEntradaSaida.Saida And Item.SubOperacao.Deposito Then
            sql &= ",'" & Me.NotaFiscalRazao.CodigoDestino & "'," & Me.NotaFiscalRazao.EnderecoDestino & vbCrLf
        Else
            sql &= ",'" & Me.NotaFiscalRazao.CodigoDeposito & "'," & Me.NotaFiscalRazao.EnderecoDeposito & vbCrLf
        End If

        sql &= ",'" & HttpContext.Current.Session("ssNomeUsuario") & "', Convert(varchar, CURRENT_TIMESTAMP, 112))" & vbCrLf

        '**************************************************
        '******** TERCEIRA PERNA ********
        '**************************************************
        i += 1
        sql &= "INSERT INTO Razao (Empresa_Id, EndEmpresa_Id, Conta_Id, Cliente_Id, EndCliente_Id, Movimento_Id, Lote_Id, Sequencia_Id, Titulo, UnidadeDeNegocio, Indexador, DataMoeda, " & vbCrLf & _
               "                   DebitoOficial, CreditoOficial, DebitoMoeda, CreditoMoeda," & vbCrLf & _
               "                   Historico, Cliente_Nf, EndCliente_Nf, EntradaSaida_Nf, Serie_Nf, Numero_Nf, Produto_NF, CFOP_NF, Sequencia_NF, Encargo_NF," & vbCrLf & _
               "                   Pedido, PrevistoRealizado, Custo, Produto, Rateado, Deposito, EndDeposito, UsuarioInclusao, UsuarioInclusaoData)" & vbCrLf & _
               " values('" & Me.NotaFiscalRazao.CodigoDeposito & "'" & vbCrLf & _
               ", " & Me.NotaFiscalRazao.EnderecoDeposito & vbCrLf & _
               ",'" & Me.NotaFiscalRazao.Empresa.Empresa.CodigoContaGrupoBanco & "'" & vbCrLf & _
               ",'' " & vbCrLf & _
               ",0 " & vbCrLf & _
               ",'" & CDate(Me.NotaFiscalRazao.Movimento).ToString("yyyy/MM/dd") & "'" & vbCrLf & _
               ",11" & vbCrLf & _
               ",  @SequenciaTransferenciaCliente + " & i & vbCrLf & _
               ",NULL" & vbCrLf
        If Me.NotaFiscalRazao.CodigoPedido = 0 Then
            sql &= ", '" & Me.NotaFiscalRazao.CodigoUnidadeDeNegocio & "' , 2 "
        Else
            sql &= ", '" & Me.NotaFiscalRazao.Pedido.CodigoUnidadeNegocio & "'" & "," & Me.NotaFiscalRazao.Pedido.CodigoIndexador
        End If
        'IIf(Me.NotaFiscalRazao.CodigoPedido = 0, ", ''", ", '" & Me.NotaFiscalRazao.Pedido.CodigoUnidadeNegocio & "'") & vbCrLf & _
        'IIf(Me.NotaFiscalRazao.CodigoPedido = 0, ", 2", "," & Me.NotaFiscalRazao.Pedido.CodigoIndexador) & vbCrLf & _
        sql &= ",'" & CDate(Me.NotaFiscalRazao.Movimento).ToString("yyyy/MM/dd") & "'" & vbCrLf & _
               "," & " 0.0, " & Str(Liquido.Valor) & ", 0.0, " & Str(LiqValorConvertido) & vbCrLf

        Historico = "LANCAM. DE TRANSF. ENTRE CONTAS CLIFOR - " & Liquido.Codigo & " OP " & Item.CodigoOperacao & "-" & Item.CodigoSubOperacao & " NF " & Me.NotaFiscalRazao.Codigo & "-" & Me.NotaFiscalRazao.Serie & " Pedido " & Me.NotaFiscalRazao.CodigoPedido & " - " & Me.NotaFiscalRazao.Cliente.Nome

        If Me.NotaFiscalRazao.CodigoTipoDeDocumento = 57 OrElse Me.NotaFiscalRazao.CodigoTipoDeDocumento = 58 Then
            'NAO FAZ NADA
        Else
            If Item.CodigoOperacao > 68 AndAlso Me.NotaFiscalRazao.Observacoes.Length > 0 Then
                Historico &= ". " & Me.NotaFiscalRazao.Observacoes
            End If
        End If

        sql &= ",'" & Historico & " '" & vbCrLf & _
               ",'" & Me.NotaFiscalRazao.CodigoCliente & "'" & vbCrLf & _
               ", " & Me.NotaFiscalRazao.EnderecoCliente & vbCrLf & _
               ",'" & Me.NotaFiscalRazao.EntradaSaida.ToString.Substring(0, 1) & "'" & vbCrLf & _
               ",'" & Me.NotaFiscalRazao.Serie & "'" & vbCrLf & _
               ", " & Me.NotaFiscalRazao.Codigo & vbCrLf & _
               ",'" & Item.CodigoProduto & "'" & vbCrLf & _
               ", " & Item.CFOP & vbCrLf & _
               ", " & Item.Sequencia & vbCrLf & _
               ",'" & Liquido.Codigo & "'" & vbCrLf & _
               ", " & Me.NotaFiscalRazao.CodigoPedido.ToSqlNULL & vbCrLf & _
               ",'P'" & vbCrLf & _
               ",'" & Liquido.CentroDeCusto & "'," & vbCrLf & _
               IIf(New Negocio.PlanoDeConta("", 0, Liquido.OperacaoEncargo.CodigoCreditaConta).TemProduto, Item.CodigoProduto, "NULL") & "," & vbCrLf & _
               IIf(Item.Rateado, "1", "0") & vbCrLf

        If Item.SubOperacao.EntradaSaida = eEntradaSaida.Saida And Item.SubOperacao.Deposito Then
            sql &= ",'" & Me.NotaFiscalRazao.CodigoDestino & "'," & Me.NotaFiscalRazao.EnderecoDestino & vbCrLf
        Else
            sql &= ",'" & Me.NotaFiscalRazao.CodigoDeposito & "'," & Me.NotaFiscalRazao.EnderecoDeposito & vbCrLf
        End If

        sql &= ",'" & HttpContext.Current.Session("ssNomeUsuario") & "', Convert(varchar, CURRENT_TIMESTAMP, 112))" & vbCrLf
    End Sub
#End Region

#Region "Contabiliza Conhecimento de Terceiro X Conhecimento Próprio - Para crédito de PIS/COFINS "
    Public Sub ContabilizaCTExCTE(ByVal Sqls As ArrayList, ByVal Deposito As String, ByVal EndDeposito As String, Optional ByVal pLote As Integer = 0)

        Dim SqlAux As String
        Dim Banco As New AcessaBanco

        Dim SqlSeq As String
        SqlSeq = " declare" & vbCrLf & _
                 " @Sequencia int" & vbCrLf & _
                 " Set @sequencia = (Select Sequencia" & vbCrLf & _
                 "                     from numerador" & vbCrLf & _
                 "                    where UPPER(Empresa_id) = '" & HttpContext.Current.Session("ssNomeServidor").ToUpper() & "'" & vbCrLf & _
                 "                      and Numerador_id      = 60)" & vbCrLf & _
                 " select isnull(Max(Sequencia_Id),isnull(@sequencia,0))" & vbCrLf & _
                 "   from razao    " & vbCrLf & _
                 "  Where lote_Id      = " & IIf(pLote > 0, pLote, IIf(_NotaFiscalRazao.CodigoOperacao > 68, "9", "10")) & vbCrLf & _
                 "    and Movimento_Id = '" & CDate(Me.NotaFiscalRazao.Movimento).ToString("yyyy/MM/dd") & "'" & vbCrLf & _
                 "    and Sequencia_id between isnull(@sequencia,0) and isnull(@sequencia,0) + 99999" & vbCrLf

        Dim ds As DataSet = Banco.ConsultaDataSet(SqlSeq, "Sequencia")

        Dim i As Integer = ds.Tables(0).Rows(0)(0)

        For Each Item As Negocio.NotaFiscalXItem In NotaFiscalRazao.Itens
            For Each Enc As Negocio.NotaFiscalXItemXEncargo In Item.Encargos

                Dim LoteSeq As String = ""

                If (Enc.Codigo = "PIS" Or Enc.Codigo = "COFINS") AndAlso Enc.OperacaoEncargo.CodigoDebitaConta.Trim <> "" AndAlso Enc.Valor > 0 Then
                    i += 1
                    If pLote > 0 Then
                        LoteSeq = "," & pLote & "," & i
                    ElseIf Item.CodigoOperacao > 68 Then
                        LoteSeq = ", 0009, " & i  'Lote / Sequencia
                    Else
                        LoteSeq = ", 0010, " & i  'Lote / Sequencia
                    End If

                    SqlAux = "INSERT INTO Razao (Empresa_Id, EndEmpresa_Id, Conta_Id, Cliente_Id, EndCliente_Id, Movimento_Id, Lote_Id, Sequencia_Id, Titulo, UnidadeDeNegocio, Indexador, " & vbCrLf & _
                             "       DataMoeda, DebitoOficial, CreditoOficial, DebitoMoeda, CreditoMoeda, Historico, Cliente_Nf, EndCliente_Nf, EntradaSaida_Nf, Serie_Nf, Numero_Nf, Produto_NF, CFOP_NF, Sequencia_NF, Encargo_NF, Pedido, PrevistoRealizado, Custo, Produto, Rateado, Deposito, EndDeposito, UsuarioInclusao, UsuarioInclusaoData)" & vbCrLf & _
                             " values( "
                    SqlAux &= "'" & Deposito & "'"                                                 'Empresa
                    SqlAux &= ", " & EndDeposito                                                   'Endereco Empresa 
                    SqlAux &= ", '" & Enc.OperacaoEncargo.CodigoDebitaConta & "'"                 'Conta

                    If Enc.OperacaoEncargo.CodigoDebitaConta.Length = 7 Then
                        SqlAux &= ", '" & _NotaFiscalRazao.CodigoCliente & "'"                     'Cliente
                        SqlAux &= ", " & _NotaFiscalRazao.EnderecoCliente                          'Endereco do Cliente
                    Else
                        SqlAux &= ",''"                     'Cliente
                        SqlAux &= ",0 "                      'Endereco do Cliente
                    End If

                    SqlAux &= ", '" & CDate(Me.NotaFiscalRazao.Movimento).ToString("yyyy/MM/dd") & "'" 'Data de Movimento
                    SqlAux &= LoteSeq                                                        'Lote Sequencia
                    SqlAux &= ", NULL"                                                       'Numero do Titulo

                    If Me.NotaFiscalRazao.CodigoPedido = 0 Then
                        SqlAux &= ", '" & Me.NotaFiscalRazao.CodigoUnidadeDeNegocio & "'"     'Unidade de Negócio
                    Else
                        SqlAux &= ", '" & Me.NotaFiscalRazao.Pedido.CodigoUnidadeNegocio & "'"     'Unidade de Negócio
                    End If

                    SqlAux &= ", 3"                                                          'Indexador
                    SqlAux &= ", '" & CDate(Me.NotaFiscalRazao.Movimento).ToString("yyyy/MM/dd") & "'" 'Data da Moeda
                    SqlAux &= ", " & Str(Enc.Valor)                                          'Valor Débito
                    SqlAux &= ", 0.0"                                                        'Valor Crédito
                    SqlAux &= ", 0.0"                                                        'Valor Débito
                    SqlAux &= ", 0.0"                                                        'Valor Crédito

                    Dim Historico As String = Enc.Codigo & " OP " & _NotaFiscalRazao.CodigoOperacao & "-" & _NotaFiscalRazao.CodigoSubOperacao & " NF " & _NotaFiscalRazao.Codigo & "-" & _NotaFiscalRazao.Serie & " Pedido " & _NotaFiscalRazao.CodigoPedido & " " & _NotaFiscalRazao.Cliente.Nome

                    If Me.NotaFiscalRazao.CodigoTipoDeDocumento = 57 OrElse Me.NotaFiscalRazao.CodigoTipoDeDocumento = 58 Then
                        'NAO FAZ NADA
                    Else
                        If (Enc.Codigo = "PRODUTO" Or Enc.Codigo = "LIQUIDO") AndAlso Item.CodigoOperacao > 68 AndAlso _NotaFiscalRazao.Observacoes.Length > 0 Then Historico &= ". " & _NotaFiscalRazao.Observacoes
                    End If

                    SqlAux &= ", '" & Historico & " '"                                         'Histórico

                    SqlAux &= ", '" & _NotaFiscalRazao.CodigoCliente & "'"                         'Cliente
                    SqlAux &= ", " & _NotaFiscalRazao.EnderecoCliente                              'Endereco do Cliente
                    SqlAux &= ", '" & _NotaFiscalRazao.EntradaSaida.ToString.Substring(0, 1) & "'" 'E/S
                    SqlAux &= ", '" & _NotaFiscalRazao.Serie & "'"                                 'Serie
                    SqlAux &= ", " & _NotaFiscalRazao.Codigo                                       'Numero 
                    SqlAux &= ",'" & Item.CodigoProduto & "'"
                    SqlAux &= "," & Item.CFOP
                    SqlAux &= "," & Item.Sequencia
                    SqlAux &= ",'" & Enc.Codigo & "'"
                    SqlAux &= ", " & _NotaFiscalRazao.CodigoPedido.ToSqlNULL                                 'Pedido 
                    SqlAux &= ", 'P',"                                                           'Previsto/Realizado  
                    SqlAux &= "'" & Enc.CentroDeCusto & "',"   'Custo
                    SqlAux &= IIf(New Negocio.PlanoDeConta("", 0, Enc.OperacaoEncargo.CodigoDebitaConta).TemProduto, Item.CodigoProduto, "NULL") & ","
                    SqlAux &= IIf(Item.Rateado, "1", "0")

                    If _NotaFiscalRazao.SubOperacao.EntradaSaida = eEntradaSaida.Saida And _NotaFiscalRazao.SubOperacao.Deposito Then
                        SqlAux &= ",'" & _NotaFiscalRazao.CodigoDestino & "'," & _NotaFiscalRazao.EnderecoDestino
                    Else
                        SqlAux &= ",'" & _NotaFiscalRazao.CodigoDeposito & "'," & _NotaFiscalRazao.EnderecoDeposito
                    End If

                    SqlAux &= ",'" & HttpContext.Current.Session("ssNomeUsuario") & "', Convert(varchar, CURRENT_TIMESTAMP, 112))"

                    Sqls.Add(SqlAux)
                End If

                If (Enc.Codigo = "PIS" Or Enc.Codigo = "COFINS") AndAlso Enc.OperacaoEncargo.CodigoCreditaConta.Trim <> "" And Enc.Valor > 0 Then
                    i += 1
                    If pLote > 0 Then
                        LoteSeq = "," & pLote & "," & i
                    ElseIf Item.CodigoOperacao > 68 Then
                        LoteSeq = ", 0009, " & i  'Lote / Sequencia
                    Else
                        LoteSeq = ", 0010, " & i  'Lote / Sequencia
                    End If

                    SqlAux = "INSERT INTO Razao (Empresa_Id, EndEmpresa_Id, Conta_Id, Cliente_Id, EndCliente_Id, Movimento_Id, Lote_Id, Sequencia_Id, Titulo, UnidadeDeNegocio, Indexador, DataMoeda, " & vbCrLf & _
                             "       DebitoOficial, CreditoOficial, DebitoMoeda, CreditoMoeda, Historico, Cliente_Nf, EndCliente_Nf, EntradaSaida_Nf, Serie_Nf, Numero_Nf, Produto_NF, CFOP_NF, Sequencia_NF, Encargo_NF, Pedido, PrevistoRealizado, Custo, Produto, Rateado, Deposito, EndDeposito, UsuarioInclusao, UsuarioInclusaoData)" & vbCrLf & _
                             " values("
                    SqlAux &= "'" & Deposito & "'"                                                 'Empresa
                    SqlAux &= ", " & EndDeposito                                                   'Endereco Empresa 
                    SqlAux &= ",'" & Enc.OperacaoEncargo.CodigoCreditaConta & "'"                 'Conta

                    If Enc.OperacaoEncargo.CodigoCreditaConta.Length = 7 Then
                        SqlAux &= ", '" & _NotaFiscalRazao.CodigoCliente & "'"                     'Cliente
                        SqlAux &= ", " & _NotaFiscalRazao.EnderecoCliente                          'Endereco do Cliente
                    Else
                        SqlAux &= ",''"                     'Cliente
                        SqlAux &= ",0 "                     'Endereco do Cliente
                    End If

                    SqlAux &= ", '" & CDate(Me.NotaFiscalRazao.Movimento).ToString("yyyy/MM/dd") & "'" 'Data de Movimento

                    SqlAux &= LoteSeq                                                        'Lote Sequencia

                    SqlAux &= ", NULL"                                                       'Numero do Titulo

                    If Me.NotaFiscalRazao.CodigoPedido = 0 Then
                        SqlAux &= ", '" & Me.NotaFiscalRazao.CodigoUnidadeDeNegocio & "'"     'Unidade de Negócio
                    Else
                        SqlAux &= ", '" & Me.NotaFiscalRazao.Pedido.CodigoUnidadeNegocio & "'"     'Unidade de Negócio
                    End If

                    SqlAux &= ", 3"                                                          'Indexador
                    SqlAux &= ", '" & CDate(Me.NotaFiscalRazao.Movimento).ToString("yyyy/MM/dd") & "'" 'Data da Moeda

                    SqlAux &= ", 0.0"                                                      'Valor Debito
                    SqlAux &= ", " & Str(Enc.Valor)                                        'Valor Crédito
                    SqlAux &= ", 0.0"                                                      'Valor Debito
                    SqlAux &= ", 0.0"                                                      'Valor Crédito

                    Dim Historico As String = Enc.Codigo & " OP " & _NotaFiscalRazao.CodigoOperacao & "-" & _NotaFiscalRazao.CodigoSubOperacao & " NF " & _NotaFiscalRazao.Codigo & "-" & _NotaFiscalRazao.Serie & " Pedido " & _NotaFiscalRazao.CodigoPedido & " - " & _NotaFiscalRazao.Cliente.Nome

                    If Me.NotaFiscalRazao.CodigoTipoDeDocumento = 57 OrElse Me.NotaFiscalRazao.CodigoTipoDeDocumento = 58 Then
                        'NAO FAZ NADA
                    Else
                        If (Enc.Codigo = "PRODUTO" Or Enc.Codigo = "LIQUIDO") AndAlso Item.CodigoOperacao > 68 AndAlso _NotaFiscalRazao.Observacoes.Length > 0 Then Historico &= ". " & _NotaFiscalRazao.Observacoes
                    End If

                    SqlAux &= ", '" & Historico & " '"                                         'Histórico

                    SqlAux &= ", '" & _NotaFiscalRazao.CodigoCliente & "'"                         'Cliente
                    SqlAux &= ", " & _NotaFiscalRazao.EnderecoCliente                              'Endereco do Cliente
                    SqlAux &= ", '" & _NotaFiscalRazao.EntradaSaida.ToString.Substring(0, 1) & "'" 'E/S
                    SqlAux &= ", '" & _NotaFiscalRazao.Serie & "'"                                 'Serie
                    SqlAux &= ", " & _NotaFiscalRazao.Codigo                                       'Numero 
                    SqlAux &= ",'" & Item.CodigoProduto & "'"
                    SqlAux &= "," & Item.CFOP
                    SqlAux &= "," & Item.Sequencia
                    SqlAux &= ",'" & Enc.Codigo & "'"
                    SqlAux &= ", " & _NotaFiscalRazao.CodigoPedido.ToSqlNULL                                 'Numero 
                    SqlAux &= ", 'P',"                                                           'Previsto/Realizado"
                    SqlAux &= "'" & Enc.CentroDeCusto & "'," 'Custo
                    SqlAux &= IIf(New Negocio.PlanoDeConta("", 0, Enc.OperacaoEncargo.CodigoCreditaConta).TemProduto, Item.CodigoProduto, "NULL") & ","
                    SqlAux &= IIf(Item.Rateado, "1", "0")

                    If _NotaFiscalRazao.SubOperacao.EntradaSaida = eEntradaSaida.Saida And _NotaFiscalRazao.SubOperacao.Deposito Then
                        SqlAux &= ",'" & _NotaFiscalRazao.CodigoDestino & "'," & _NotaFiscalRazao.EnderecoDestino
                    Else
                        SqlAux &= ",'" & _NotaFiscalRazao.CodigoDeposito & "'," & _NotaFiscalRazao.EnderecoDeposito
                    End If

                    SqlAux &= ",'" & HttpContext.Current.Session("ssNomeUsuario") & "', Convert(varchar, CURRENT_TIMESTAMP, 112))"

                    Sqls.Add(SqlAux)
                End If
            Next
        Next
    End Sub
#End Region

#Region "Contabiliza Contrato de Frete"
    Public Sub ContabilizaContratoDeFrete(ByVal Sqls As ArrayList)
        If Not NotaFiscalRazao.SubOperacao.Contabil Then Exit Sub

        Dim SqlAux As String
        Dim Banco As New AcessaBanco

        Dim SqlSeq As String
        SqlSeq = " declare" & vbCrLf & _
                 " @Sequencia int" & vbCrLf & _
                 " Set @sequencia = (Select Sequencia" & vbCrLf & _
                 "                     from numerador" & vbCrLf & _
                 "                    where UPPER(Empresa_id) = '" & HttpContext.Current.Session("ssNomeServidor").ToUpper() & "'" & vbCrLf & _
                 "                      and Numerador_id      = 60)" & vbCrLf & _
                 " select isnull(Max(Sequencia_Id),isnull(@sequencia,0)) as Sequencia" & vbCrLf & _
                 "   from razao    " & vbCrLf & _
                 "  Where lote_Id      = 72 " & vbCrLf & _
                 "    and Movimento_Id = '" & CDate(Me.NotaFiscalRazao.Movimento).ToString("yyyy/MM/dd") & "'" & vbCrLf & _
                 "    and Sequencia_id between isnull(@sequencia,0) and isnull(@sequencia,0) + 99999" & vbCrLf

        Dim ds As DataSet = Banco.ConsultaDataSet(SqlSeq, "Sequencia")

        Dim i As Integer = ds.Tables(0).Rows(0)(0)

        If NotaFiscalRazao.SubOperacao.Contabil Then
            For Each Item As Negocio.NotaFiscalXItem In NotaFiscalRazao.Itens
                For Each Enc As Negocio.NotaFiscalXItemXEncargo In Item.Encargos

                    Dim LoteSeq As String = ""

                    If (Enc.Codigo = "ADTODEFRETE" Or Enc.Codigo = "PEDAGIO" Or Enc.Codigo = "CADASTRO PANCAR" Or Enc.Codigo = "SEGURO") AndAlso _
                        Enc.OperacaoEncargo.CodigoDebitaConta.Trim <> "" And Enc.Valor > 0 Then

                        i += 1
                        LoteSeq = ", 0072, " & i  'Lote / Sequencia

                        SqlAux = "INSERT INTO Razao (Empresa_Id, EndEmpresa_Id, Conta_Id, Cliente_Id, EndCliente_Id, Movimento_Id, Lote_Id, Sequencia_Id, Titulo, UnidadeDeNegocio, Indexador, " & vbCrLf & _
                                 "       DataMoeda, DebitoOficial, CreditoOficial, DebitoMoeda, CreditoMoeda, Historico, Cliente_Nf, EndCliente_Nf, EntradaSaida_Nf, Serie_Nf, Numero_Nf, Pedido, PrevistoRealizado, Custo, Produto, Rateado, UsuarioInclusao, UsuarioInclusaoData)" & vbCrLf & _
                                 " values( "
                        SqlAux &= "'" & _NotaFiscalRazao.NotaTrocaOrigem.NotaTrocaOrigem.CodigoEmpresa & "'" 'Empresa
                        SqlAux &= ", " & _NotaFiscalRazao.NotaTrocaOrigem.NotaTrocaOrigem.EnderecoEmpresa    'Endereco Empresa 
                        SqlAux &= ", '" & Enc.OperacaoEncargo.CodigoDebitaConta & "'"  'Conta

                        If Enc.OperacaoEncargo.CodigoDebitaConta.Length = 7 Then
                            SqlAux &= ", '" & _NotaFiscalRazao.CodigoCliente & "'"                'Cliente
                            SqlAux &= ", " & _NotaFiscalRazao.EnderecoCliente                     'Endereco do Cliente
                        Else
                            SqlAux &= ",''"                                                       'Cliente
                            SqlAux &= ",0 "                                                       'Endereco do Cliente
                        End If

                        SqlAux &= ", '" & CDate(Me.NotaFiscalRazao.Movimento).ToString("yyyy/MM/dd") & "'" 'Data de Movimento
                        SqlAux &= LoteSeq                                                        'Lote Sequencia
                        SqlAux &= ", NULL"                                                       'Numero do Titulo
                        SqlAux &= ", '" & _NotaFiscalRazao.NotaTrocaOrigem.NotaTrocaOrigem.Pedido.CodigoUnidadeNegocio & "'" 'Unidade de Negócio
                        SqlAux &= ", 3"                                                          'Indexador
                        SqlAux &= ", '" & CDate(Me.NotaFiscalRazao.Movimento).ToString("yyyy/MM/dd") & "'" 'Data da Moeda
                        SqlAux &= ", " & Str(Enc.Valor)                                          'Valor Débito
                        SqlAux &= ", 0.0"                                                        'Valor Crédito
                        SqlAux &= ", 0.0"                                                        'Valor Débito
                        SqlAux &= ", 0.0"                                                        'Valor Crédito

                        SqlAux &= ", '" & Enc.Codigo & " OP " & _NotaFiscalRazao.CodigoOperacao & "-" & _NotaFiscalRazao.CodigoSubOperacao & " NF " & _NotaFiscalRazao.Codigo & "-" & _NotaFiscalRazao.Serie & " Pedido " & _NotaFiscalRazao.CodigoPedido & " " & _NotaFiscalRazao.Cliente.Nome & "'"   'Histórico

                        SqlAux &= ", '" & _NotaFiscalRazao.CodigoCliente & "'"                         'Cliente
                        SqlAux &= ", " & _NotaFiscalRazao.EnderecoCliente                              'Endereco do Cliente
                        SqlAux &= ", '" & _NotaFiscalRazao.EntradaSaida.ToString.Substring(0, 1) & "'" 'E/S
                        SqlAux &= ", '" & _NotaFiscalRazao.Serie & "'"                                 'Serie
                        SqlAux &= ", " & _NotaFiscalRazao.Codigo                                       'Numero 
                        SqlAux &= ", " & _NotaFiscalRazao.CodigoPedido.ToSqlNULL                                 'Pedido 
                        SqlAux &= ", 'R',"                                                           'Previsto/Realizado  
                        SqlAux &= "'" & Enc.CentroDeCusto & "',"   'Custo
                        SqlAux &= IIf(New Negocio.PlanoDeConta("", 0, Enc.OperacaoEncargo.CodigoDebitaConta).TemProduto, Item.CodigoProduto, "NULL") & ","
                        SqlAux &= IIf(Item.Rateado, "1", "0")
                        SqlAux &= ",'" & HttpContext.Current.Session("ssNomeUsuario") & "', Convert(varchar, CURRENT_TIMESTAMP, 112))"

                        Sqls.Add(SqlAux)

                        i += 1
                        LoteSeq = ", 0072, " & i  'Lote / Sequencia

                        SqlAux = "INSERT INTO Razao (Empresa_Id, EndEmpresa_Id, Conta_Id, Cliente_Id, EndCliente_Id, Movimento_Id, Lote_Id, Sequencia_Id, Titulo, UnidadeDeNegocio, Indexador, DataMoeda, " & vbCrLf & _
                                 "       DebitoOficial, CreditoOficial, DebitoMoeda, CreditoMoeda, Historico, Cliente_Nf, EndCliente_Nf, EntradaSaida_Nf, Serie_Nf, Numero_Nf, Pedido, PrevistoRealizado, Custo, Produto, Rateado, UsuarioInclusao, UsuarioInclusaoData)" & vbCrLf & _
                                 " values("
                        SqlAux &= "'" & _NotaFiscalRazao.NotaTrocaOrigem.NotaTrocaOrigem.CodigoEmpresa & "'"  'Empresa
                        SqlAux &= ", " & _NotaFiscalRazao.NotaTrocaOrigem.NotaTrocaOrigem.EnderecoEmpresa     'Endereco Empresa 

                        If (Enc.Codigo = "ADTODEFRETE" Or Enc.Codigo = "PEDAGIO") Then
                            SqlAux &= ",'" & _NotaFiscalRazao.Empresa.Empresa.CodigoContaGrupoBanco & "'" 'Conta
                        Else
                            SqlAux &= ", '" & Enc.OperacaoEncargo.CodigoCreditaConta & "'"               'Conta
                        End If

                        If _NotaFiscalRazao.Empresa.Empresa.CodigoContaGrupoBanco.Length = 7 Then
                            SqlAux &= ", '" & _NotaFiscalRazao.CodigoCliente & "'"                'Cliente
                            SqlAux &= ", " & _NotaFiscalRazao.EnderecoCliente                     'Endereco do Cliente
                        Else
                            SqlAux &= ",''"                                                       'Cliente
                            SqlAux &= ",0 "                                                       'Endereco do Cliente
                        End If

                        SqlAux &= ", '" & CDate(Me.NotaFiscalRazao.Movimento).ToString("yyyy/MM/dd") & "'" 'Data de Movimento

                        SqlAux &= LoteSeq                                                        'Lote Sequencia

                        SqlAux &= ", NULL"                                                          'Numero do Titulo
                        SqlAux &= ", '" & _NotaFiscalRazao.NotaTrocaOrigem.NotaTrocaOrigem.Pedido.CodigoUnidadeNegocio & "'"  'Unidade de Negócio
                        SqlAux &= ", 3"                                                          'Indexador
                        SqlAux &= ", '" & CDate(Me.NotaFiscalRazao.Movimento).ToString("yyyy/MM/dd") & "'" 'Data da Moeda

                        SqlAux &= ", 0.0"                                                      'Valor Debito
                        SqlAux &= ", " & Str(Enc.Valor)                                        'Valor Crédito
                        SqlAux &= ", 0.0"                                                      'Valor Debito
                        SqlAux &= ", 0.0"                                                      'Valor Crédito

                        SqlAux &= ", '" & Enc.Codigo & " OP " & _NotaFiscalRazao.CodigoOperacao & "-" & _NotaFiscalRazao.CodigoSubOperacao & " NF " & _NotaFiscalRazao.Codigo & "-" & _NotaFiscalRazao.Serie & IIf(_NotaFiscalRazao.CodigoPedido > 0, " Pedido " & _NotaFiscalRazao.CodigoPedido, " ") & " - " & _NotaFiscalRazao.Cliente.Nome & "'"   'Histórico

                        SqlAux &= ", '" & _NotaFiscalRazao.CodigoCliente & "'"                         'Cliente
                        SqlAux &= ", " & _NotaFiscalRazao.EnderecoCliente                              'Endereco do Cliente
                        SqlAux &= ", '" & _NotaFiscalRazao.EntradaSaida.ToString.Substring(0, 1) & "'" 'E/S
                        SqlAux &= ", '" & _NotaFiscalRazao.Serie & "'"                                 'Serie
                        SqlAux &= ", " & _NotaFiscalRazao.Codigo                                       'Numero 
                        SqlAux &= ", " & _NotaFiscalRazao.CodigoPedido.ToSqlNULL                                 'Numero 
                        SqlAux &= ", 'R',"                                                           'Previsto/Realizado"
                        SqlAux &= "'" & Enc.CentroDeCusto & "'," 'Custo
                        SqlAux &= IIf(New Negocio.PlanoDeConta("", 0, Enc.OperacaoEncargo.CodigoCreditaConta).TemProduto, Item.CodigoProduto, "NULL") & ","
                        SqlAux &= IIf(Item.Rateado, "1", "0")
                        SqlAux &= ",'" & HttpContext.Current.Session("ssNomeUsuario") & "', Convert(varchar, CURRENT_TIMESTAMP, 112))"

                        Sqls.Add(SqlAux)
                    End If
                Next
            Next
        End If
    End Sub

    Public Sub ExcluiContabilizacaoContratoDeFrete(ByVal Sqls As ArrayList)
        Dim Sql As String
        Sql = " DELETE From Razao" & vbCrLf & _
              "  WHERE Lote_Id         = 72 " & vbCrLf & _
              "    And Empresa_Id      ='" & _NotaFiscalRazao.NotaTrocaOrigem.NotaTrocaOrigem.CodigoEmpresa & "'" & vbCrLf & _
              "    And EndEmpresa_Id   = " & _NotaFiscalRazao.NotaTrocaOrigem.NotaTrocaOrigem.EnderecoEmpresa & vbCrLf & _
              "    And Cliente_nf      ='" & _NotaFiscalRazao.CodigoCliente & "'" & vbCrLf & _
              "    And EndCliente_nf   = " & _NotaFiscalRazao.EnderecoCliente & vbCrLf & _
              "    And Numero_nf       = " & _NotaFiscalRazao.Codigo & vbCrLf & _
              "    And Serie_nf        ='" & _NotaFiscalRazao.Serie & "'" & vbCrLf & _
              "    And EntradaSaida_nf ='" & _NotaFiscalRazao.EntradaSaida.ToString.Substring(0, 1) & "'"
        Sqls.Add(Sql)
    End Sub
#End Region

#Region "Contabilizar Fretes"
    Public Function ExcluirFretesRazao(ByVal pNota As Negocio.NotaFiscal) As Boolean
        Dim sqls As New ArrayList
        sqls.Add(ExcluirFretesRazaoSql(pNota.Movimento, pNota.Movimento, "", 0, pNota))
        Dim banco As New AcessaBanco
        Return banco.GravaBanco(sqls)
    End Function

    Public Function ExcluirFretesRazao(ByVal pEmpresa As String, ByVal pEndEmpresa As Integer, ByVal DataInicial As Date, ByVal dataFinal As Date, Optional ByVal pNota As Negocio.NotaFiscal = Nothing) As Boolean
        Dim sqls As New ArrayList
        sqls.Add(ExcluirFretesRazaoSql(DataInicial, dataFinal, pEmpresa, pEndEmpresa, pNota))
        Dim banco As New AcessaBanco

        Return banco.GravaBanco(sqls)
    End Function

    Public Function ExcluirFretesRazaoSql(ByVal DataInicial As Date, ByVal dataFinal As Date, Optional ByVal pEmpresa As String = "", Optional ByVal pEndEmpresa As Integer = 0, Optional ByVal pNota As Negocio.NotaFiscal = Nothing) As String
        Dim sql As String
        sql = "Delete razao" & vbCrLf & _
              " Where Movimento_id  between '" & CDate(DataInicial).ToString("yyyy-MM-dd") & "' and '" & CDate(dataFinal).ToString("yyyy-MM-dd") & "'" & vbCrLf & _
              "   and Lote_id       in (21,22)" & vbCrLf
        If pNota Is Nothing Then
            sql &= IIf(pEmpresa.Length > 0, " and Empresa_Id ='" & pEmpresa & "' and EndEmpresa_id = " & pEndEmpresa, "")
        Else
            sql &= " And Empresa_Id      ='" & pNota.CodigoDeposito & "'" & vbCrLf & _
                   " And EndEmpresa_Id   = " & pNota.EnderecoDeposito & vbCrLf & _
                   " And Cliente_Nf      ='" & pNota.CodigoEmpresa & "'" & vbCrLf & _
                   " And EndCliente_nf   = " & pNota.EnderecoEmpresa & vbCrLf & _
                   " And EntradaSaida_Nf ='" & IIf(pNota.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "'" & vbCrLf & _
                   " And Serie_Nf        ='" & pNota.Serie & "'" & vbCrLf & _
                   " And Numero_nf       = " & pNota.Codigo
        End If

        Return sql
    End Function

    Public Function ContabilizarFretesNoRazao(ByVal pNota As Negocio.NotaFiscal) As Boolean
        Return ContabilizarFretesNoRazao(pNota.CodigoEmpresa, pNota.EnderecoEmpresa, pNota.Movimento, pNota.Movimento, pNota.CodigoDeposito, pNota.EnderecoDeposito, pNota)
    End Function

    Public Function ContabilizarFretesNoRazao(ByVal pEmpresaTransportadora As String, ByVal pEndEmpresaTransportadora As Integer, ByVal pDataInicial As Date, ByVal pDataFinal As Date, Optional ByVal pEmpresaContratante As String = "", Optional ByVal pEndEmpresaContratante As Integer = 0, Optional ByVal pNota As Negocio.NotaFiscal = Nothing) As Boolean
        Dim ds As DataSet
        Dim Banco As New AcessaBanco
        Dim Sqls As New ArrayList
        Dim sql As String

        Try
            Sqls.Add(ExcluirFretesRazaoSql(pDataInicial, pDataFinal, pEmpresaContratante, pEndEmpresaContratante, pNota))

            sql = "Select NF.Empresa_Id, NF.EndEmpresa_Id, NF.Cliente_Id, NF.EndCliente_Id, NF.EntradaSaida_Id, NF.Serie_Id, NF.Nota_Id, NF.Movimento, " & vbCrLf & _
                  "       isnull(NN2.OrigemEmpresa_Id,NN1.OrigemEmpresa_Id) as OrigemEmpresa_Id, " & vbCrLf & _
                  "       isnull(NN2.OrigemEndEmpresa_Id,NN1.OrigemEndEmpresa_Id) as OrigemEndEmpresa_Id, " & vbCrLf & _
                  "       isnull(NN2.OrigemCliente_Id,NN1.OrigemCliente_Id) as OrigemCliente_Id, " & vbCrLf & _
                  "       isnull(NN2.OrigemEndCliente_Id,NN1.OrigemEndCliente_Id) as OrigemEndCliente_Id, " & vbCrLf & _
                  "       isnull(NN2.OrigemEntradaSaida_Id,NN1.OrigemEntradaSaida_Id) as OrigemEntradaSaida_Id, " & vbCrLf & _
                  "       isnull(NN2.OrigemNota_id,NN1.OrigemNota_id) as OrigemNota_id, " & vbCrLf & _
                  "       isnull(NN2.OrigemSerie_Id,NN1.OrigemSerie_Id) as OrigemSerie_Id, " & vbCrLf & _
                  "       NFxT.Proprietario as Transportador, " & vbCrLf & _
                  "       NFxT.EndProprietario as EndTransportador, " & vbCrLf & _
                  "       C.Nome, " & vbCrLf & _
                  "       P.UnidadedeNegocio,  " & vbCrLf & _
                  "       NFOr.pedido, " & vbCrLf & _
                  "       (Select Top(1) Produto_Id " & vbCrLf & _
                  "          from Notasfiscaisxitens NfxI " & vbCrLf & _
                  "         Where NfxI.Empresa_Id      = NFOr.Empresa_Id " & vbCrLf & _
                  "		   and NfxI.EndEmpresa_Id   = NFOr.EndEmpresa_Id " & vbCrLf & _
                  "		   and NfxI.Cliente_Id      = NFOr.Cliente_Id " & vbCrLf & _
                  "		   and NfxI.EndCliente_Id   = NFOr.EndCliente_Id " & vbCrLf & _
                  "		   and NfxI.EntradaSaida_Id = NFOr.EntradaSaida_Id " & vbCrLf & _
                  "		   and NfxI.Serie_Id        = NFOr.Serie_Id " & vbCrLf & _
                  "		   and NfxI.Nota_Id         = NFOr.Nota_Id " & vbCrLf & _
                  "        ) as OrigemProduto_Id " & vbCrLf & _
                  "  into #temp " & vbCrLf & _
                  "  From NotasFiscais NF " & vbCrLf & _
                  " INNER JOIN NotasFiscaisXTransportadores NFxT " & vbCrLf & _
                  "    ON NF.Empresa_Id      = NFxT.Empresa_Id  " & vbCrLf & _
                  "   AND NF.EndEmpresa_Id   = NFxT.EndEmpresa_Id  " & vbCrLf & _
                  "   AND NF.Cliente_Id      = NFxT.Cliente_Id  " & vbCrLf & _
                  "   AND NF.EndCliente_Id   = NFxT.EndCliente_Id  " & vbCrLf & _
                  "   AND NF.EntradaSaida_Id = NFxT.EntradaSaida_Id  " & vbCrLf & _
                  "   AND NF.Serie_Id        = NFxT.Serie_Id  " & vbCrLf & _
                  "   AND NF.Nota_Id         = NFxT.Nota_Id  " & vbCrLf & _
                  " inner join NotasxNotas NN1 " & vbCrLf & _
                  "    on NN1.Empresa_Id      = NF.Empresa_Id " & vbCrLf & _
                  "   and NN1.EndEmpresa_Id   = NF.EndEmpresa_Id " & vbCrLf & _
                  "   and NN1.Cliente_Id      = NF.Cliente_Id " & vbCrLf & _
                  "   and NN1.EndCliente_Id   = NF.EndCliente_Id " & vbCrLf & _
                  "   and NN1.EntradaSaida_Id = NF.EntradaSaida_Id " & vbCrLf & _
                  "   and NN1.Serie_Id        = NF.Serie_Id " & vbCrLf & _
                  "   and NN1.Nota_Id         = NF.Nota_Id " & vbCrLf & _
                  "  Left join NotasxNotas NN2 " & vbCrLf & _
                  "    on NN2.Empresa_Id      = NN1.OrigemEmpresa_Id " & vbCrLf & _
                  "   and NN2.EndEmpresa_Id   = NN1.OrigemEndEmpresa_Id " & vbCrLf & _
                  "   and NN2.Cliente_Id      = NN1.OrigemCliente_Id " & vbCrLf & _
                  "   and NN2.EndCliente_Id   = NN1.OrigemEndCliente_Id " & vbCrLf & _
                  "   and NN2.EntradaSaida_Id = NN1.OrigemEntradaSaida_Id " & vbCrLf & _
                  "   and NN2.Serie_Id        = NN1.OrigemSerie_Id " & vbCrLf & _
                  "   and NN2.Nota_Id         = NN1.OrigemNota_Id " & vbCrLf & _
                  " inner join NotasFiscais NFOr " & vbCrLf & _
                  "    on NFOr.Empresa_Id      = isnull(NN2.OrigemEmpresa_Id,NN1.OrigemEmpresa_Id) " & vbCrLf & _
                  "   and NFOr.EndEmpresa_Id   = isnull(NN2.OrigemEndEmpresa_Id,NN1.OrigemEndEmpresa_Id) " & vbCrLf & _
                  "   and NFOr.Cliente_Id      = isnull(NN2.OrigemCliente_Id,NN1.OrigemCliente_Id) " & vbCrLf & _
                  "   and NFOr.EndCliente_Id   = isnull(NN2.OrigemEndCliente_Id,NN1.OrigemEndCliente_Id) " & vbCrLf & _
                  "   and NFOr.EntradaSaida_Id = isnull(NN2.OrigemEntradaSaida_Id,NN1.OrigemEntradaSaida_Id) " & vbCrLf & _
                  "   and NFOr.Serie_Id        = isnull(NN2.OrigemSerie_Id,NN1.OrigemSerie_Id) " & vbCrLf & _
                  "   and NFOr.Nota_Id         = isnull(NN2.OrigemNota_id,NN1.OrigemNota_id) " & vbCrLf & _
                  " inner Join Pedidos P " & vbCrLf & _
                  "    on P.Pedido_id = NFOr.Pedido " & vbCrLf & _
                  " Inner Join Clientes C " & vbCrLf & _
                  "    on C.Cliente_id  = NFOr.cliente_id " & vbCrLf & _
                  "   and C.Endereco_id = NFOr.Endcliente_id " & vbCrLf

            If pNota Is Nothing Then
                sql &= " WHERE NF.serie_Id          ='REC' " & vbCrLf & _
                       "   and NF.EntradaSaida_Id   ='S' " & vbCrLf & _
                       "   and NF.Movimento between  '" & CDate(pDataInicial).ToString("yyyy-MM-dd") & "' and '" & CDate(pDataFinal).ToString("yyyy-MM-dd") & "'" & vbCrLf & _
                       "   And NF.Empresa_Id        ='" & pEmpresaTransportadora & "'" & vbCrLf & _
                       "   And NF.EndEmpresa_Id     = " & pEndEmpresaTransportadora & vbCrLf & _
                       "   And NF.Situacao          = 1 " & vbCrLf
            Else
                sql &= " WHERE NF.serie_Id            ='" & pNota.Serie & "'" & vbCrLf & _
                         "   and NF.EntradaSaida_Id   ='" & IIf(pNota.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "' " & vbCrLf & _
                         "   and NF.Movimento         ='" & CDate(pNota.Movimento).ToString("yyyy-MM-dd") & "'" & vbCrLf & _
                         "   And NF.Empresa_Id        ='" & pNota.CodigoEmpresa & "'" & vbCrLf & _
                         "   And NF.EndEmpresa_Id     = " & pNota.EnderecoEmpresa & vbCrLf & _
                         "   And NF.Situacao          = 1 " & vbCrLf
            End If


            sql &= " SELECT distinct T.Empresa_Id, T.EndEmpresa_Id, T.Cliente_Id, T.EndCliente_Id, T.EntradaSaida_Id, T.Serie_Id, T.Nota_Id, T.Movimento,  " & vbCrLf & _
                  "        T.Transportador, T.EndTransportador, " & vbCrLf & _
                  "        T.OrigemEmpresa_Id, T.OrigemEndEmpresa_Id," & vbCrLf & _
                  "        convert(nvarchar(1000),'') as historico, " & vbCrLf & _
                  "        convert(nvarchar(18),'') as UnidadeDeNegocio, " & vbCrLf & _
                  "        convert(nvarchar(30),'') as OrigemProduto_Id " & vbCrLf & _
                  "  into #Temp2 " & vbCrLf & _
                  "  FROM #Temp T " & vbCrLf & _
                  IIf(pEmpresaContratante.Length > 0, " Where T.OrigemEmpresa_Id    ='" & pEmpresaContratante & "' and T.OrigemEndEmpresa_Id = " & pEndEmpresaContratante, "") & vbCrLf & _
                  " DECLARE CurItens " & vbCrLf & _
                  " CURSOR FOR  " & vbCrLf & _
                  " SELECT Nota_id  " & vbCrLf & _
                  "   FROM #Temp2 " & vbCrLf & _
                  " DECLARE @Nota INTEGER " & vbCrLf & _
                  " DECLARE @HISTORICO NVARCHAR(1000) " & vbCrLf & _
                  " DECLARE @DESCRICAO AS NVARCHAR(1000) " & vbCrLf & _
                  " Declare @Pedido as integer " & vbCrLf & _
                  " Declare @PedidoTemp as integer " & vbCrLf & _
                  " Declare @Cliente as nvarchar(50) " & vbCrLf & _
                  " OPEN CurItens " & vbCrLf & _
                  " FETCH NEXT FROM CurItens INTO @Nota " & vbCrLf & _
                  " WHILE @@FETCH_STATUS = 0 " & vbCrLf & _
                  " BEGIN " & vbCrLf & _
                  "    set @DESCRICAO = '' " & vbCrLf & _
                  "	DECLARE Cur2  " & vbCrLf & _
                  "	CURSOR FOR  " & vbCrLf & _
                  "	SELECT Pedido, ' Cliente: ' + OrigemCliente_Id + '-' + convert(nvarchar,OrigemEndCliente_Id)+' ' + Nome as Cliente,  'NF' + convert(nvarchar,OrigemNota_Id) + '-' + OrigemSerie_id as Historico " & vbCrLf & _
                  "	  FROM #Temp " & vbCrLf & _
                  "     Where Nota_id = @Nota " & vbCrLf & _
                  "     Order by Empresa_Id, EndEmpresa_Id, Cliente_Id, EndCliente_Id, EntradaSaida_Id, Pedido, Nota_Id, Serie_Id " & vbCrLf & _
                  "    OPEN Cur2 " & vbCrLf & _
                  "	FETCH NEXT FROM Cur2 INTO @pedido, @Cliente, @HISTORICO " & vbCrLf & _
                  "    Set @PedidoTemp = 0  " & vbCrLf & _
                  "	WHILE @@FETCH_STATUS = 0  " & vbCrLf & _
                  "	BEGIN  " & vbCrLf & _
                  "      if @PedidoTemp <> @pedido  " & vbCrLf & _
                  "        begin  " & vbCrLf & _
                  "         set @PedidoTemp =  @pedido  " & vbCrLf & _
                  "         set @DESCRICAO = @DESCRICAO + ' Pedido:' + convert(nvarchar,@pedido) + @Cliente + ' NF:' +  @HISTORICO + ', ' " & vbCrLf & _
                  "        end " & vbCrLf & _
                  "      else " & vbCrLf & _
                  "        SET @DESCRICAO = @DESCRICAO + @HISTORICO + ', '" & vbCrLf & _
                  "	FETCH NEXT FROM Cur2 INTO @pedido, @Cliente, @HISTORICO " & vbCrLf & _
                  "	END " & vbCrLf & _
                  "	CLOSE Cur2 " & vbCrLf & _
                  "	DEALLOCATE Cur2 " & vbCrLf & _
                  "    UPDATE #Temp2 set " & vbCrLf & _
                  "        Historico        = @DESCRICAO " & vbCrLf & _
                  "       ,UnidadeDeNegocio = (Select Top 1 UnidadeDeNegocio from #temp Where Nota_id = @Nota) " & vbCrLf & _
                  "       ,OrigemProduto_Id = (Select Top 1 OrigemProduto_Id from #temp Where Nota_id = @Nota) " & vbCrLf & _
                  "    Where Nota_id = @Nota " & vbCrLf & _
                  " FETCH NEXT FROM CurItens INTO  @Nota " & vbCrLf & _
                  " END " & vbCrLf & _
                  " CLOSE CurItens " & vbCrLf & _
                  " DEALLOCATE CurItens " & vbCrLf & _
                  " select * " & vbCrLf & _
                  "   from #temp2" & vbCrLf & _
                  " Order by OrigemEmpresa_Id, OrigemEndEmpresa_Id, movimento, Transportador, EndTransportador " & vbCrLf & _
                  " drop table #temp " & vbCrLf & _
                  " drop table #temp2 " & vbCrLf


            ds = Banco.ConsultaDataSet(sql, "Fretes")
            Dim i21 As Integer
            Dim i22 As Integer

            Dim Empresa As String
            Dim EndEmpresa As Integer
            Dim data As Date
            Empresa = ds.Tables(0).Rows(0)("OrigemEmpresa_Id")
            EndEmpresa = ds.Tables(0).Rows(0)("OrigemEndEmpresa_Id")
            data = ds.Tables(0).Rows(0)("Movimento")

            For Each row As DataRow In ds.Tables(0).Rows
                Dim cons As New Negocio.NotaFiscal
                cons.CodigoEmpresa = row("Empresa_Id")
                cons.EnderecoEmpresa = row("EndEmpresa_Id")
                cons.CodigoCliente = row("Cliente_Id")
                cons.EnderecoCliente = row("EndCliente_Id")
                cons.Codigo = row("Nota_Id")
                cons.Serie = row("Serie_Id")
                cons.EntradaSaida = IIf(row("EntradaSaida_Id") = "S", Negocio.eEntradaSaida.Saida, Negocio.eEntradaSaida.Entrada)
                _NotaFiscalRazao = New Negocio.NotaFiscal(cons)


                If Empresa <> row("OrigemEmpresa_Id") Or EndEmpresa <> row("OrigemEndEmpresa_Id") Or data <> row("Movimento") Then
                    i21 = 0
                    i22 = 0
                    Empresa = row("OrigemEmpresa_Id")
                    EndEmpresa = row("OrigemEndEmpresa_Id")
                    data = row("Movimento")
                End If

                'If NotaFiscalRazao.SubOperacao.Contabil Then
                For Each Item As Negocio.NotaFiscalXItem In NotaFiscalRazao.Itens
                    Item.Encargos = New Negocio.ListNotaFiscalXItemXEncargo(Item, New List(Of eEtapaEncago) From {eEtapaEncago.Todos})
                    For Each Enc As Negocio.NotaFiscalXItemXEncargo In Item.Encargos
                        'If Enc.Codigo <> "ADTODEFRETE" Then
                        If Enc.Codigo = "LIQUIDO" And Enc.OperacaoEncargo.CodigoDebitaConta = "" And Enc.OperacaoEncargo.CodigoCreditaConta = "" And Enc.Valor > 0 Then
                            sql = "INSERT INTO Razao(Empresa_Id, EndEmpresa_Id, Conta_Id, Cliente_Id, EndCliente_Id, Movimento_Id, Lote_Id, Sequencia_Id, Titulo, UnidadeDeNegocio, Indexador, " & vbCrLf & _
                                  "  DataMoeda, DebitoOficial, CreditoOficial, DebitoMoeda, CreditoMoeda, Historico, Cliente_Nf, EndCliente_Nf, EntradaSaida_Nf, Serie_Nf, Numero_Nf, Pedido, PrevistoRealizado, Custo, Produto, UsuarioInclusao, UsuarioInclusaoData)" & vbCrLf & _
                                  " VALUES ("

                            sql &= "'" & row("OrigemEmpresa_Id") & "'"                              'Empresa Contratante do serviço de transporte
                            sql &= ", " & row("OrigemEndEmpresa_Id")                                'Endereco Contratante do serviço de transporte

                            sql &= ", '" & _NotaFiscalRazao.SubOperacao.CodigoGrupoContas & "'"           'Conta


                            sql &= ",'" & row("Transportador") & "'"                    'Cliente
                            sql &= ", " & row("EndTransportador")                       'Endereco do Cliente


                            sql &= ", '" & CDate(Me.NotaFiscalRazao.Movimento).ToString("yyyy/MM/dd") & "'"   'Data de Movimento

                            If Enc.Encargo.Etapa > 0 Then
                                i22 += 1
                                sql &= ", 0022, " & i22                                         'Lote 22 Referente a complementacao do servico de Transporte 
                            Else
                                i21 += 1
                                sql &= ", 0021, " & i21                                         'Lote 21 Referente a servico de Transporte 
                            End If

                            sql &= ", NULL"                                                            'Numero do Titulo
                            sql &= ", '" & row("UnidadeDeNegocio") & "'"                            'Unidade de Negócio
                            sql &= ", 3"                                                            'Indexador
                            sql &= ", '" & CDate(Me.NotaFiscalRazao.Movimento).ToString("yyyy/MM/dd") & "'"   'Data da Moeda

                            If _NotaFiscalRazao.EntradaSaida.ToString.Substring(0, 1) = "S" Then
                                sql &= ", " & Str(Enc.Valor)                                        'Valor Débito
                                sql &= ", 0.0"                                                      'Valor Crédito
                                sql &= ", 0.0"                                                      'Valor Débito
                                sql &= ", 0.0"                                                      'Valor Crédito
                            Else
                                sql &= ", 0.0"                                                      'Valor Debito
                                sql &= ", " & Str(Enc.Valor)                                        'Valor Crédito
                                sql &= ", 0.0"                                                      'Valor Debito
                                sql &= ", 0.0"                                                      'Valor Crédito
                            End If

                            sql &= ", '" & IIf(_NotaFiscalRazao.Serie = "REC", "RECIBO ", "CTRC ") & _NotaFiscalRazao.Codigo & "-" & _NotaFiscalRazao.Serie & " OP " & _NotaFiscalRazao.CodigoOperacao & "-" & _NotaFiscalRazao.CodigoSubOperacao & " REFERENTE A : " & row("Historico") & "'"   'Histórico & row("OrigemNota_id") & "-" & row("OrigemSerie_Id") & " PEDIDO  " & row("Pedido") & " DE " & row("Nome") &

                            sql &= ",'" & _NotaFiscalRazao.CodigoEmpresa & "'"                     'Cliente  = Empresa que efetuou o Serviço de Transporte
                            sql &= ", " & _NotaFiscalRazao.EnderecoEmpresa                          'Endereco = Empresa que efetuou o Serviço de Transporte
                            sql &= ",'" & _NotaFiscalRazao.EntradaSaida.ToString.Substring(0, 1) & "'" 'E/S
                            sql &= ",'" & _NotaFiscalRazao.Serie & "'"                             'Serie
                            sql &= ", " & _NotaFiscalRazao.Codigo                                   'Numero 
                            sql &= ", " & _NotaFiscalRazao.CodigoPedido.ToSqlNULL                             'Numero 
                            sql &= ",'P',0,"      'Previsto/Realizado
                            sql &= IIf(New Negocio.PlanoDeConta("", 0, Item.SubOperacao.CodigoGrupoContas).TemProduto, row("OrigemProduto_Id"), "NULL")  'Produto Transportado Milho, Soja
                            sql &= ",'" & HttpContext.Current.Session("ssNomeUsuario") & "', Convert(varchar, CURRENT_TIMESTAMP, 112));"

                            Sqls.Add(sql)
                        Else
                            If Enc.OperacaoEncargo.CodigoDebitaConta <> "" And Enc.Valor > 0 Then
                                sql = "INSERT INTO Razao (Empresa_Id, EndEmpresa_Id, Conta_Id, Cliente_Id, EndCliente_Id, Movimento_Id, Lote_Id, Sequencia_Id, Titulo, UnidadeDeNegocio, Indexador," & vbCrLf & _
                                         " DataMoeda, DebitoOficial, CreditoOficial, DebitoMoeda, CreditoMoeda, Historico, Cliente_Nf, EndCliente_Nf, EntradaSaida_Nf, Serie_Nf, Numero_Nf, Pedido, PrevistoRealizado, Custo, Produto, UsuarioInclusao, UsuarioInclusaoData)" & vbCrLf & _
                                         " VALUES ("
                                sql &= "'" & row("OrigemEmpresa_Id") & "'"                            'Empresa Contratante do serviço de transporte
                                sql &= ", " & row("OrigemEndEmpresa_Id")                              'Endereco Contratante do serviço de transporte



                                sql &= ", '" & Enc.OperacaoEncargo.CodigoDebitaConta & "'"                 'Conta

                                If Enc.OperacaoEncargo.CodigoDebitaConta.Length = 7 Then
                                    sql &= ",'" & row("Transportador") & "'"                    'Cliente
                                    sql &= ", " & row("EndTransportador")                       'Endereco do Cliente
                                Else
                                    sql &= ",''"                     'Cliente
                                    sql &= ",0 "                      'Endereco do Cliente
                                End If

                                sql &= ", '" & CDate(Me.NotaFiscalRazao.Movimento).ToString("yyyy/MM/dd") & "'" 'Data de Movimento

                                If Enc.Encargo.Etapa > 0 Then
                                    i22 += 1
                                    sql &= ", 0022, " & i22                                         'Lote 22 Referente a complementacao do servico de Transporte 
                                Else
                                    i21 += 1
                                    sql &= ", 0021, " & i21                                         'Lote 21 Referente a servico de Transporte 
                                End If

                                sql &= ", NULL"                                                          'Numero do Titulo
                                sql &= ", '" & row("UnidadeDeNegocio") & "'"     'Unidade de Negócio
                                sql &= ", 3"                                                          'Indexador
                                sql &= ", '" & CDate(Me.NotaFiscalRazao.Movimento).ToString("yyyy/MM/dd") & "'" 'Data da Moeda
                                sql &= ", " & Str(Enc.Valor)                                          'Valor Débito
                                sql &= ", 0.0"                                                        'Valor Crédito
                                sql &= ", 0.0"                                                        'Valor Débito
                                sql &= ", 0.0"                                                        'Valor Crédito
                                sql &= ", '" & Enc.Codigo & " OP " & _NotaFiscalRazao.CodigoOperacao & "-" & _NotaFiscalRazao.CodigoSubOperacao & " " & IIf(_NotaFiscalRazao.Serie = "REC", "RECIBO ", "CTRC ") & _NotaFiscalRazao.Codigo & "-" & _NotaFiscalRazao.Serie & " REFERENTE A: " & row("Historico") & "'" 'Histórico row("OrigemNota_id") & "-" & row("OrigemSerie_Id") & " PEDIDO  " & row("Pedido") & " DE " & row("Nome") &
                                sql &= ", '" & _NotaFiscalRazao.CodigoEmpresa & "'"                         'Cliente  = Empresa que efetuou o Serviço de Transporte
                                sql &= ", " & _NotaFiscalRazao.EnderecoEmpresa                              'Endereco = Empresa que efetuou o Serviço de Transporte 
                                sql &= ", '" & _NotaFiscalRazao.EntradaSaida.ToString.Substring(0, 1) & "'" 'E/S
                                sql &= ", '" & _NotaFiscalRazao.Serie & "'"                                 'Serie
                                sql &= ", " & _NotaFiscalRazao.Codigo                                       'Numero 
                                sql &= ", " & _NotaFiscalRazao.CodigoPedido.ToSqlNULL                                 'Pedido 
                                sql &= ", 'P',0,"                                                           'Previsto/Realizado  
                                sql &= IIf(New Negocio.PlanoDeConta("", 0, Enc.OperacaoEncargo.CodigoDebitaConta).TemProduto, row("OrigemProduto_Id"), "NULL")  'Produto Transportado Milho, Soja
                                sql &= ",'" & HttpContext.Current.Session("ssNomeUsuario") & "', Convert(varchar, CURRENT_TIMESTAMP, 112));" & vbCrLf

                                Sqls.Add(sql)
                            End If

                            If Enc.OperacaoEncargo.CodigoCreditaConta <> "" And Enc.Valor > 0 Then
                                sql = "INSERT INTO Razao(Empresa_Id, EndEmpresa_Id, Conta_Id, Cliente_Id, EndCliente_Id, Movimento_Id, Lote_Id, Sequencia_Id, Titulo, UnidadeDeNegocio, Indexador, DataMoeda," & vbCrLf & _
                                         " DebitoOficial, CreditoOficial, DebitoMoeda, CreditoMoeda, Historico, Cliente_Nf, EndCliente_Nf, EntradaSaida_Nf, Serie_Nf, Numero_Nf, Pedido, PrevistoRealizado, Custo, Produto, UsuarioInclusao, UsuarioInclusaoData)" & vbCrLf & _
                                         " VALUES ("
                                sql &= "'" & row("OrigemEmpresa_Id") & "'"                            'Empresa Contratante do serviço de transporte
                                sql &= ", " & row("OrigemEndEmpresa_Id")                              'Endereco Empresa Contratante do serviço de transporte


                                sql &= ",'" & Enc.OperacaoEncargo.CodigoCreditaConta & "'"  'Conta
                                If Enc.OperacaoEncargo.CodigoCreditaConta.Length = 7 Then
                                    sql &= ",'" & row("Transportador") & "'"                     'Cliente
                                    sql &= ", " & row("EndTransportador")                        'Endereco do Cliente
                                Else
                                    sql &= ",''"                     'Cliente
                                    sql &= ",0 "                      'Endereco do Cliente
                                End If

                                sql &= ", '" & CDate(Me.NotaFiscalRazao.Movimento).ToString("yyyy/MM/dd") & "'" 'Data de Movimento

                                If Enc.Encargo.Etapa > 0 Then
                                    i22 += 1
                                    sql &= ", 0022, " & i22                                         'Lote 22 Referente a complementacao do servico de Transporte 
                                Else
                                    i21 += 1
                                    sql &= ", 0021, " & i21                                         'Lote 21 Referente a servico de Transporte 
                                End If

                                sql &= ", NULL"                                                     'Numero do Titulo
                                sql &= ", '" & row("UnidadeDeNegocio") & "'"     'Unidade de Negócio
                                sql &= ", 3"                                                          'Indexador
                                sql &= ", '" & CDate(Me.NotaFiscalRazao.Movimento).ToString("yyyy/MM/dd") & "'" 'Data da Moeda

                                sql &= ", 0.0"                                                      'Valor Debito
                                sql &= ", " & Str(Enc.Valor)                                        'Valor Crédito
                                sql &= ", 0.0"                                                      'Valor Debito
                                sql &= ", 0.0"                                                      'Valor Crédito
                                sql &= ", '" & Enc.Codigo & " OP " & _NotaFiscalRazao.CodigoOperacao & "-" & _NotaFiscalRazao.CodigoSubOperacao & " " & IIf(_NotaFiscalRazao.Serie = "REC", "RECIBO ", "CTRC ") & _NotaFiscalRazao.Codigo & "-" & _NotaFiscalRazao.Serie & " REFERENTE A: " & row("Historico") & "'" 'Histórico & row("OrigemNota_id") & "-" & row("OrigemSerie_Id") & " PEDIDO  " & row("Pedido") & " DE " & row("Nome")
                                sql &= ", '" & _NotaFiscalRazao.CodigoEmpresa & "'"                         'Cliente = Empresa que efetuou o Serviço de Transporte
                                sql &= ", " & _NotaFiscalRazao.EnderecoEmpresa                              'Endereco = Empresa que efetuou o Serviço de Transporte 
                                sql &= ", '" & _NotaFiscalRazao.EntradaSaida.ToString.Substring(0, 1) & "'" 'E/S
                                sql &= ", '" & _NotaFiscalRazao.Serie & "'"                                 'Serie
                                sql &= ", " & _NotaFiscalRazao.Codigo                                       'Numero 
                                sql &= ", " & _NotaFiscalRazao.CodigoPedido.ToSqlNULL                                 'Numero 
                                sql &= ", 'P',0,"                                                           'Previsto/Realizado"
                                sql &= IIf(New Negocio.PlanoDeConta("", 0, Enc.OperacaoEncargo.CodigoCreditaConta).TemProduto, row("OrigemProduto_Id"), "NULL")  'Produto Transportado Milho, Soja
                                sql &= ",'" & HttpContext.Current.Session("ssNomeUsuario") & "', Convert(varchar, CURRENT_TIMESTAMP, 112));" & vbCrLf

                                Sqls.Add(sql)
                            End If
                        End If
                        'End If
                    Next
                Next
                'End If
            Next
        Catch ex As Exception
            Return False
        End Try

        Return Banco.GravaBanco(Sqls)
    End Function
#End Region

#Region "Transferir Fretes para o Razão - Custo Neri"
    Public Function TransferirFretesNoRazao(ByVal pEmpresaTransportadora As String, ByVal pEndEmpresaTransportadora As Integer, ByVal pDataInicial As Date, ByVal pDataFinal As Date, Optional ByVal pEmpresaContratante As String = "", Optional ByVal pEndEmpresaContratante As Integer = 0, Optional ByVal pNota As Negocio.NotaFiscal = Nothing) As Boolean
        Dim ds As DataSet
        Dim Banco As New AcessaBanco
        Dim Sqls As New ArrayList
        Dim sql As String

        Try
            sql = "Select NF.Empresa_Id, NF.EndEmpresa_Id, NF.Cliente_Id, NF.EndCliente_Id, NF.EntradaSaida_Id, NF.Serie_Id, NF.Nota_Id, NF.Movimento, " & vbCrLf & _
                  "       isnull(NN2.OrigemEmpresa_Id,NN1.OrigemEmpresa_Id) as OrigemEmpresa_Id, " & vbCrLf & _
                  "       isnull(NN2.OrigemEndEmpresa_Id,NN1.OrigemEndEmpresa_Id) as OrigemEndEmpresa_Id, " & vbCrLf & _
                  "       isnull(NN2.OrigemCliente_Id,NN1.OrigemCliente_Id) as OrigemCliente_Id, " & vbCrLf & _
                  "       isnull(NN2.OrigemEndCliente_Id,NN1.OrigemEndCliente_Id) as OrigemEndCliente_Id, " & vbCrLf & _
                  "       isnull(NN2.OrigemEntradaSaida_Id,NN1.OrigemEntradaSaida_Id) as OrigemEntradaSaida_Id, " & vbCrLf & _
                  "       isnull(NN2.OrigemNota_id,NN1.OrigemNota_id) as OrigemNota_id, " & vbCrLf & _
                  "       isnull(NN2.OrigemSerie_Id,NN1.OrigemSerie_Id) as OrigemSerie_Id, " & vbCrLf & _
                  "       NFxT.Proprietario as Transportador, " & vbCrLf & _
                  "       NFxT.EndProprietario as EndTransportador, " & vbCrLf & _
                  "       C.Nome, " & vbCrLf & _
                  "       P.UnidadedeNegocio,  " & vbCrLf & _
                  "       NFOr.pedido, " & vbCrLf & _
                  "       (Select Top(1) Produto_Id " & vbCrLf & _
                  "          from Notasfiscaisxitens NfxI " & vbCrLf & _
                  "         Where NfxI.Empresa_Id      = NFOr.Empresa_Id " & vbCrLf & _
                  "		   and NfxI.EndEmpresa_Id   = NFOr.EndEmpresa_Id " & vbCrLf & _
                  "		   and NfxI.Cliente_Id      = NFOr.Cliente_Id " & vbCrLf & _
                  "		   and NfxI.EndCliente_Id   = NFOr.EndCliente_Id " & vbCrLf & _
                  "		   and NfxI.EntradaSaida_Id = NFOr.EntradaSaida_Id " & vbCrLf & _
                  "		   and NfxI.Serie_Id        = NFOr.Serie_Id " & vbCrLf & _
                  "		   and NfxI.Nota_Id         = NFOr.Nota_Id " & vbCrLf & _
                  "        ) as OrigemProduto_Id " & vbCrLf & _
                  "  into #temp " & vbCrLf & _
                  "  From NotasFiscais NF " & vbCrLf & _
                  " INNER JOIN NotasFiscaisXTransportadores NFxT " & vbCrLf & _
                  "    ON NF.Empresa_Id      = NFxT.Empresa_Id  " & vbCrLf & _
                  "   AND NF.EndEmpresa_Id   = NFxT.EndEmpresa_Id  " & vbCrLf & _
                  "   AND NF.Cliente_Id      = NFxT.Cliente_Id  " & vbCrLf & _
                  "   AND NF.EndCliente_Id   = NFxT.EndCliente_Id  " & vbCrLf & _
                  "   AND NF.EntradaSaida_Id = NFxT.EntradaSaida_Id  " & vbCrLf & _
                  "   AND NF.Serie_Id        = NFxT.Serie_Id  " & vbCrLf & _
                  "   AND NF.Nota_Id         = NFxT.Nota_Id  " & vbCrLf & _
                  " inner join NotasxNotas NN1 " & vbCrLf & _
                  "    on NN1.Empresa_Id      = NF.Empresa_Id " & vbCrLf & _
                  "   and NN1.EndEmpresa_Id   = NF.EndEmpresa_Id " & vbCrLf & _
                  "   and NN1.Cliente_Id      = NF.Cliente_Id " & vbCrLf & _
                  "   and NN1.EndCliente_Id   = NF.EndCliente_Id " & vbCrLf & _
                  "   and NN1.EntradaSaida_Id = NF.EntradaSaida_Id " & vbCrLf & _
                  "   and NN1.Serie_Id        = NF.Serie_Id " & vbCrLf & _
                  "   and NN1.Nota_Id         = NF.Nota_Id " & vbCrLf & _
                  "  Left join NotasxNotas NN2 " & vbCrLf & _
                  "    on NN2.Empresa_Id      = NN1.OrigemEmpresa_Id " & vbCrLf & _
                  "   and NN2.EndEmpresa_Id   = NN1.OrigemEndEmpresa_Id " & vbCrLf & _
                  "   and NN2.Cliente_Id      = NN1.OrigemCliente_Id " & vbCrLf & _
                  "   and NN2.EndCliente_Id   = NN1.OrigemEndCliente_Id " & vbCrLf & _
                  "   and NN2.EntradaSaida_Id = NN1.OrigemEntradaSaida_Id " & vbCrLf & _
                  "   and NN2.Serie_Id        = NN1.OrigemSerie_Id " & vbCrLf & _
                  "   and NN2.Nota_Id         = NN1.OrigemNota_Id " & vbCrLf & _
                  " inner join NotasFiscais NFOr " & vbCrLf & _
                  "    on NFOr.Empresa_Id      = isnull(NN2.OrigemEmpresa_Id,NN1.OrigemEmpresa_Id) " & vbCrLf & _
                  "   and NFOr.EndEmpresa_Id   = isnull(NN2.OrigemEndEmpresa_Id,NN1.OrigemEndEmpresa_Id) " & vbCrLf & _
                  "   and NFOr.Cliente_Id      = isnull(NN2.OrigemCliente_Id,NN1.OrigemCliente_Id) " & vbCrLf & _
                  "   and NFOr.EndCliente_Id   = isnull(NN2.OrigemEndCliente_Id,NN1.OrigemEndCliente_Id) " & vbCrLf & _
                  "   and NFOr.EntradaSaida_Id = isnull(NN2.OrigemEntradaSaida_Id,NN1.OrigemEntradaSaida_Id) " & vbCrLf & _
                  "   and NFOr.Serie_Id        = isnull(NN2.OrigemSerie_Id,NN1.OrigemSerie_Id) " & vbCrLf & _
                  "   and NFOr.Nota_Id         = isnull(NN2.OrigemNota_id,NN1.OrigemNota_id) " & vbCrLf & _
                  " inner Join Pedidos P " & vbCrLf & _
                  "    on P.Pedido_id = NFOr.Pedido " & vbCrLf & _
                  " Inner Join Clientes C " & vbCrLf & _
                  "    on C.Cliente_id  = NFOr.cliente_id " & vbCrLf & _
                  "   and C.Endereco_id = NFOr.Endcliente_id " & vbCrLf

            sql &= " WHERE NF.serie_Id            ='" & pNota.Serie & "'" & vbCrLf & _
                     "   and NF.EntradaSaida_Id   ='" & IIf(pNota.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "' " & vbCrLf & _
                     "   and NF.Movimento between  '" & CDate(pDataInicial).ToString("yyyy-MM-dd") & "' and '" & CDate(pDataFinal).ToString("yyyy-MM-dd") & "'" & vbCrLf & _
                     "   And NF.Empresa_Id        ='" & pEmpresaTransportadora & "'" & vbCrLf & _
                     "   And NF.EndEmpresa_Id     = " & pEndEmpresaTransportadora & vbCrLf & _
                     "   And NF.Situacao          = 1 " & vbCrLf

            sql &= " SELECT distinct T.Empresa_Id, T.EndEmpresa_Id, T.Cliente_Id, T.EndCliente_Id, T.EntradaSaida_Id, T.Serie_Id, T.Nota_Id, T.Movimento,  " & vbCrLf & _
                  "        T.Transportador, T.EndTransportador, " & vbCrLf & _
                  "        T.OrigemEmpresa_Id, T.OrigemEndEmpresa_Id," & vbCrLf & _
                  "        convert(nvarchar(1000),'') as historico, " & vbCrLf & _
                  "        convert(nvarchar(18),'') as UnidadeDeNegocio, " & vbCrLf & _
                  "        convert(nvarchar(30),'') as OrigemProduto_Id " & vbCrLf & _
                  "  into #Temp2 " & vbCrLf & _
                  "  FROM #Temp T " & vbCrLf & _
                  IIf(pEmpresaContratante.Length > 0, " Where T.OrigemEmpresa_Id    ='" & pEmpresaContratante & "' and T.OrigemEndEmpresa_Id = " & pEndEmpresaContratante, "") & vbCrLf & _
                  " DECLARE CurItens " & vbCrLf & _
                  " CURSOR FOR  " & vbCrLf & _
                  " SELECT Nota_id  " & vbCrLf & _
                  "   FROM #Temp2 " & vbCrLf & _
                  " DECLARE @Nota INTEGER " & vbCrLf & _
                  " DECLARE @HISTORICO NVARCHAR(1000) " & vbCrLf & _
                  " DECLARE @DESCRICAO AS NVARCHAR(1000) " & vbCrLf & _
                  " Declare @Pedido as integer " & vbCrLf & _
                  " Declare @PedidoTemp as integer " & vbCrLf & _
                  " Declare @Cliente as nvarchar(50) " & vbCrLf & _
                  " OPEN CurItens " & vbCrLf & _
                  " FETCH NEXT FROM CurItens INTO @Nota " & vbCrLf & _
                  " WHILE @@FETCH_STATUS = 0 " & vbCrLf & _
                  " BEGIN " & vbCrLf & _
                  "    set @DESCRICAO = '' " & vbCrLf & _
                  "	DECLARE Cur2  " & vbCrLf & _
                  "	CURSOR FOR  " & vbCrLf & _
                  "	SELECT Pedido, ' Cliente: ' + OrigemCliente_Id + '-' + convert(nvarchar,OrigemEndCliente_Id)+' ' + Nome as Cliente,  'NF' + convert(nvarchar,OrigemNota_Id) + '-' + OrigemSerie_id as Historico " & vbCrLf & _
                  "	  FROM #Temp " & vbCrLf & _
                  "     Where Nota_id = @Nota " & vbCrLf & _
                  "     Order by Empresa_Id, EndEmpresa_Id, Cliente_Id, EndCliente_Id, EntradaSaida_Id, Pedido, Nota_Id, Serie_Id " & vbCrLf & _
                  "    OPEN Cur2 " & vbCrLf & _
                  "	FETCH NEXT FROM Cur2 INTO @pedido, @Cliente, @HISTORICO " & vbCrLf & _
                  "    Set @PedidoTemp = 0  " & vbCrLf & _
                  "	WHILE @@FETCH_STATUS = 0  " & vbCrLf & _
                  "	BEGIN  " & vbCrLf & _
                  "      if @PedidoTemp <> @pedido  " & vbCrLf & _
                  "        begin  " & vbCrLf & _
                  "         set @PedidoTemp =  @pedido  " & vbCrLf & _
                  "         set @DESCRICAO = @DESCRICAO + ' Pedido:' + convert(nvarchar,@pedido) + @Cliente + ' NF:' +  @HISTORICO + ', ' " & vbCrLf & _
                  "        end " & vbCrLf & _
                  "      else " & vbCrLf & _
                  "        SET @DESCRICAO = @DESCRICAO + @HISTORICO + ', '" & vbCrLf & _
                  "	FETCH NEXT FROM Cur2 INTO @pedido, @Cliente, @HISTORICO " & vbCrLf & _
                  "	END " & vbCrLf & _
                  "	CLOSE Cur2 " & vbCrLf & _
                  "	DEALLOCATE Cur2 " & vbCrLf & _
                  "    UPDATE #Temp2 set " & vbCrLf & _
                  "        Historico        = @DESCRICAO " & vbCrLf & _
                  "       ,UnidadeDeNegocio = (Select Top 1 UnidadeDeNegocio from #temp Where Nota_id = @Nota) " & vbCrLf & _
                  "       ,OrigemProduto_Id = (Select Top 1 OrigemProduto_Id from #temp Where Nota_id = @Nota) " & vbCrLf & _
                  "    Where Nota_id = @Nota " & vbCrLf & _
                  " FETCH NEXT FROM CurItens INTO  @Nota " & vbCrLf & _
                  " END " & vbCrLf & _
                  " CLOSE CurItens " & vbCrLf & _
                  " DEALLOCATE CurItens " & vbCrLf & _
                  " select * " & vbCrLf & _
                  "   from #temp2" & vbCrLf & _
                  " Order by OrigemEmpresa_Id, OrigemEndEmpresa_Id, movimento, Transportador, EndTransportador " & vbCrLf & _
                  " drop table #temp " & vbCrLf & _
                  " drop table #temp2 " & vbCrLf


            ds = Banco.ConsultaDataSet(sql, "Fretes")
            Dim i21 As Integer = 2000
            Dim i22 As Integer = 2000

            Dim Empresa As String
            Dim EndEmpresa As Integer
            Dim data As Date
            Empresa = ds.Tables(0).Rows(0)("OrigemEmpresa_Id")
            EndEmpresa = ds.Tables(0).Rows(0)("OrigemEndEmpresa_Id")
            data = ds.Tables(0).Rows(0)("Movimento")

            For Each row As DataRow In ds.Tables(0).Rows
                Dim cons As New Negocio.NotaFiscal
                cons.CodigoEmpresa = row("Empresa_Id")
                cons.EnderecoEmpresa = row("EndEmpresa_Id")
                cons.CodigoCliente = row("Cliente_Id")
                cons.EnderecoCliente = row("EndCliente_Id")
                cons.Codigo = row("Nota_Id")
                cons.Serie = row("Serie_Id")
                cons.EntradaSaida = IIf(row("EntradaSaida_Id") = "S", Negocio.eEntradaSaida.Saida, Negocio.eEntradaSaida.Entrada)
                _NotaFiscalRazao = New Negocio.NotaFiscal(cons)

                sql = "Delete razao" & vbCrLf & _
                          " Where Movimento_id  between '" & CDate(pDataInicial).ToString("yyyy-MM-dd") & "' and '" & CDate(pDataFinal).ToString("yyyy-MM-dd") & "'" & vbCrLf & _
                          "   and Lote_id       in (21,22)" & vbCrLf & _
                          " And Empresa_Id      ='" & row("OrigemEmpresa_Id") & "'" & vbCrLf & _
                          " And EndEmpresa_Id   = " & row("OrigemEndEmpresa_Id") & vbCrLf & _
                          " And Cliente_Nf      ='" & row("Empresa_Id") & "'" & vbCrLf & _
                          " And EndCliente_nf   = " & row("EndEmpresa_Id") & vbCrLf & _
                          " And EntradaSaida_Nf ='" & row("EntradaSaida_Id") & "'" & vbCrLf & _
                          " And Serie_Nf        ='" & row("Serie_Id") & "'" & vbCrLf & _
                          " And Numero_nf       = " & row("Nota_Id")

                Sqls.Add(sql)

                If Empresa <> row("OrigemEmpresa_Id") Or EndEmpresa <> row("OrigemEndEmpresa_Id") Or data <> row("Movimento") Then
                    i21 = 2000
                    i22 = 2000
                    Empresa = row("OrigemEmpresa_Id")
                    EndEmpresa = row("OrigemEndEmpresa_Id")
                    data = row("Movimento")
                End If

                'If NotaFiscalRazao.SubOperacao.Contabil Then
                For Each Item As Negocio.NotaFiscalXItem In NotaFiscalRazao.Itens
                    Item.Encargos = New Negocio.ListNotaFiscalXItemXEncargo(Item, New List(Of eEtapaEncago) From {eEtapaEncago.Todos})
                    For Each Enc As Negocio.NotaFiscalXItemXEncargo In Item.Encargos
                        'If Enc.Codigo <> "ADTODEFRETE" Then
                        If Enc.Codigo = "LIQUIDO" And Enc.OperacaoEncargo.CodigoDebitaConta = "" And Enc.OperacaoEncargo.CodigoCreditaConta = "" And Enc.Valor > 0 Then
                            sql = "INSERT INTO Razao(Empresa_Id, EndEmpresa_Id, Conta_Id, Cliente_Id, EndCliente_Id, Movimento_Id, Lote_Id, Sequencia_Id, Titulo, UnidadeDeNegocio, Indexador, " & vbCrLf & _
                                  "  DataMoeda, DebitoOficial, CreditoOficial, DebitoMoeda, CreditoMoeda, Historico, Cliente_Nf, EndCliente_Nf, EntradaSaida_Nf, Serie_Nf, Numero_Nf, Pedido, PrevistoRealizado, Custo, Produto, UsuarioInclusao, UsuarioInclusaoData)" & vbCrLf & _
                                  " VALUES ("

                            sql &= "'" & row("OrigemEmpresa_Id") & "'"                              'Empresa Contratante do serviço de transporte
                            sql &= ", " & row("OrigemEndEmpresa_Id")                                'Endereco Contratante do serviço de transporte

                            sql &= ", '" & _NotaFiscalRazao.SubOperacao.CodigoGrupoContas & "'"           'Conta


                            sql &= ",'" & row("Transportador") & "'"                    'Cliente
                            sql &= ", " & row("EndTransportador")                       'Endereco do Cliente


                            sql &= ", '" & CDate(Me.NotaFiscalRazao.Movimento).ToString("yyyy/MM/dd") & "'"   'Data de Movimento

                            If Enc.Encargo.Etapa > 0 Then
                                i22 += 1
                                sql &= ", 0022, " & i22                                         'Lote 22 Referente a complementacao do servico de Transporte 
                            Else
                                i21 += 1
                                sql &= ", 0021, " & i21                                         'Lote 21 Referente a servico de Transporte 
                            End If

                            sql &= ", NULL"                                                            'Numero do Titulo
                            sql &= ", '" & row("UnidadeDeNegocio") & "'"                            'Unidade de Negócio
                            sql &= ", 3"                                                            'Indexador
                            sql &= ", '" & CDate(Me.NotaFiscalRazao.Movimento).ToString("yyyy/MM/dd") & "'"   'Data da Moeda

                            If _NotaFiscalRazao.EntradaSaida.ToString.Substring(0, 1) = "S" Then
                                sql &= ", " & Str(Enc.Valor)                                        'Valor Débito
                                sql &= ", 0.0"                                                      'Valor Crédito
                                sql &= ", 0.0"                                                      'Valor Débito
                                sql &= ", 0.0"                                                      'Valor Crédito
                            Else
                                sql &= ", 0.0"                                                      'Valor Debito
                                sql &= ", " & Str(Enc.Valor)                                        'Valor Crédito
                                sql &= ", 0.0"                                                      'Valor Debito
                                sql &= ", 0.0"                                                      'Valor Crédito
                            End If

                            sql &= ", '" & IIf(_NotaFiscalRazao.Serie = "REC", "RECIBO ", "CTRC ") & _NotaFiscalRazao.Codigo & "-" & _NotaFiscalRazao.Serie & " OP " & _NotaFiscalRazao.CodigoOperacao & "-" & _NotaFiscalRazao.CodigoSubOperacao & " REFERENTE A : " & row("Historico") & "'"   'Histórico & row("OrigemNota_id") & "-" & row("OrigemSerie_Id") & " PEDIDO  " & row("Pedido") & " DE " & row("Nome") &

                            sql &= ",'" & _NotaFiscalRazao.CodigoEmpresa & "'"                     'Cliente  = Empresa que efetuou o Serviço de Transporte
                            sql &= ", " & _NotaFiscalRazao.EnderecoEmpresa                          'Endereco = Empresa que efetuou o Serviço de Transporte
                            sql &= ",'" & _NotaFiscalRazao.EntradaSaida.ToString.Substring(0, 1) & "'" 'E/S
                            sql &= ",'" & _NotaFiscalRazao.Serie & "'"                             'Serie
                            sql &= ", " & _NotaFiscalRazao.Codigo                                   'Numero 
                            sql &= ", " & _NotaFiscalRazao.CodigoPedido.ToSqlNULL                             'Numero 
                            sql &= ",'P',0,"      'Previsto/Realizado
                            sql &= IIf(New Negocio.PlanoDeConta("", 0, Item.SubOperacao.CodigoGrupoContas).TemProduto, row("OrigemProduto_Id"), "NULL")  'Produto Transportado Milho, Soja
                            sql &= ",'" & HttpContext.Current.Session("ssNomeUsuario") & "', Convert(varchar, CURRENT_TIMESTAMP, 112));"

                            Sqls.Add(sql)
                        Else
                            If Enc.OperacaoEncargo.CodigoDebitaConta <> "" And Enc.Valor > 0 Then
                                sql = "INSERT INTO Razao (Empresa_Id, EndEmpresa_Id, Conta_Id, Cliente_Id, EndCliente_Id, Movimento_Id, Lote_Id, Sequencia_Id, Titulo, UnidadeDeNegocio, Indexador," & vbCrLf & _
                                         " DataMoeda, DebitoOficial, CreditoOficial, DebitoMoeda, CreditoMoeda, Historico, Cliente_Nf, EndCliente_Nf, EntradaSaida_Nf, Serie_Nf, Numero_Nf, Pedido, PrevistoRealizado, Custo, Produto, UsuarioInclusao, UsuarioInclusaoData)" & vbCrLf & _
                                         " VALUES ("
                                sql &= "'" & row("OrigemEmpresa_Id") & "'"                            'Empresa Contratante do serviço de transporte
                                sql &= ", " & row("OrigemEndEmpresa_Id")                              'Endereco Contratante do serviço de transporte



                                sql &= ", '" & Enc.OperacaoEncargo.CodigoDebitaConta & "'"                 'Conta

                                If Enc.OperacaoEncargo.CodigoDebitaConta.Length = 7 Then
                                    sql &= ",'" & row("Transportador") & "'"                    'Cliente
                                    sql &= ", " & row("EndTransportador")                       'Endereco do Cliente
                                Else
                                    sql &= ",''"                     'Cliente
                                    sql &= ",0 "                      'Endereco do Cliente
                                End If

                                sql &= ", '" & CDate(Me.NotaFiscalRazao.Movimento).ToString("yyyy/MM/dd") & "'" 'Data de Movimento

                                If Enc.Encargo.Etapa > 0 Then
                                    i22 += 1
                                    sql &= ", 0022, " & i22                                         'Lote 22 Referente a complementacao do servico de Transporte 
                                Else
                                    i21 += 1
                                    sql &= ", 0021, " & i21                                         'Lote 21 Referente a servico de Transporte 
                                End If

                                sql &= ", NULL"                                                          'Numero do Titulo
                                sql &= ", '" & row("UnidadeDeNegocio") & "'"     'Unidade de Negócio
                                sql &= ", 3"                                                          'Indexador
                                sql &= ", '" & CDate(Me.NotaFiscalRazao.Movimento).ToString("yyyy/MM/dd") & "'" 'Data da Moeda
                                sql &= ", " & Str(Enc.Valor)                                          'Valor Débito
                                sql &= ", 0.0"                                                        'Valor Crédito
                                sql &= ", 0.0"                                                        'Valor Débito
                                sql &= ", 0.0"                                                        'Valor Crédito
                                sql &= ", '" & Enc.Codigo & " OP " & _NotaFiscalRazao.CodigoOperacao & "-" & _NotaFiscalRazao.CodigoSubOperacao & " " & IIf(_NotaFiscalRazao.Serie = "REC", "RECIBO ", "CTRC ") & _NotaFiscalRazao.Codigo & "-" & _NotaFiscalRazao.Serie & " REFERENTE A: " & row("Historico") & "'" 'Histórico row("OrigemNota_id") & "-" & row("OrigemSerie_Id") & " PEDIDO  " & row("Pedido") & " DE " & row("Nome") &
                                sql &= ", '" & _NotaFiscalRazao.CodigoEmpresa & "'"                         'Cliente  = Empresa que efetuou o Serviço de Transporte
                                sql &= ", " & _NotaFiscalRazao.EnderecoEmpresa                              'Endereco = Empresa que efetuou o Serviço de Transporte 
                                sql &= ", '" & _NotaFiscalRazao.EntradaSaida.ToString.Substring(0, 1) & "'" 'E/S
                                sql &= ", '" & _NotaFiscalRazao.Serie & "'"                                 'Serie
                                sql &= ", " & _NotaFiscalRazao.Codigo                                       'Numero 
                                sql &= ", " & _NotaFiscalRazao.CodigoPedido.ToSqlNULL                                 'Pedido 
                                sql &= ", 'P',0,"                                                           'Previsto/Realizado  
                                sql &= IIf(New Negocio.PlanoDeConta("", 0, Enc.OperacaoEncargo.CodigoDebitaConta).TemProduto, row("OrigemProduto_Id"), "NULL")  'Produto Transportado Milho, Soja
                                sql &= ",'" & HttpContext.Current.Session("ssNomeUsuario") & "', Convert(varchar, CURRENT_TIMESTAMP, 112));" & vbCrLf

                                Sqls.Add(sql)
                            End If

                            If Enc.OperacaoEncargo.CodigoCreditaConta <> "" And Enc.Valor > 0 Then
                                sql = "INSERT INTO Razao(Empresa_Id, EndEmpresa_Id, Conta_Id, Cliente_Id, EndCliente_Id, Movimento_Id, Lote_Id, Sequencia_Id, Titulo, UnidadeDeNegocio, Indexador, DataMoeda," & vbCrLf & _
                                         " DebitoOficial, CreditoOficial, DebitoMoeda, CreditoMoeda, Historico, Cliente_Nf, EndCliente_Nf, EntradaSaida_Nf, Serie_Nf, Numero_Nf, Pedido, PrevistoRealizado, Custo, Produto, UsuarioInclusao, UsuarioInclusaoData)" & vbCrLf & _
                                         " VALUES ("
                                sql &= "'" & row("OrigemEmpresa_Id") & "'"                            'Empresa Contratante do serviço de transporte
                                sql &= ", " & row("OrigemEndEmpresa_Id")                              'Endereco Empresa Contratante do serviço de transporte


                                sql &= ",'" & Enc.OperacaoEncargo.CodigoCreditaConta & "'"  'Conta
                                If Enc.OperacaoEncargo.CodigoCreditaConta.Length = 7 Then
                                    sql &= ",'" & row("Transportador") & "'"                     'Cliente
                                    sql &= ", " & row("EndTransportador")                        'Endereco do Cliente
                                Else
                                    sql &= ",''"                     'Cliente
                                    sql &= ",0 "                      'Endereco do Cliente
                                End If

                                sql &= ", '" & CDate(Me.NotaFiscalRazao.Movimento).ToString("yyyy/MM/dd") & "'" 'Data de Movimento

                                If Enc.Encargo.Etapa > 0 Then
                                    i22 += 1
                                    sql &= ", 0022, " & i22                                         'Lote 22 Referente a complementacao do servico de Transporte 
                                Else
                                    i21 += 1
                                    sql &= ", 0021, " & i21                                         'Lote 21 Referente a servico de Transporte 
                                End If

                                sql &= ", NULL"                                                          'Numero do Titulo
                                sql &= ", '" & row("UnidadeDeNegocio") & "'"     'Unidade de Negócio
                                sql &= ", 3"                                                          'Indexador
                                sql &= ", '" & CDate(Me.NotaFiscalRazao.Movimento).ToString("yyyy/MM/dd") & "'" 'Data da Moeda

                                sql &= ", 0.0"                                                      'Valor Debito
                                sql &= ", " & Str(Enc.Valor)                                        'Valor Crédito
                                sql &= ", 0.0"                                                      'Valor Debito
                                sql &= ", 0.0"                                                      'Valor Crédito
                                sql &= ", '" & Enc.Codigo & " OP " & _NotaFiscalRazao.CodigoOperacao & "-" & _NotaFiscalRazao.CodigoSubOperacao & " " & IIf(_NotaFiscalRazao.Serie = "REC", "RECIBO ", "CTRC ") & _NotaFiscalRazao.Codigo & "-" & _NotaFiscalRazao.Serie & " REFERENTE A: " & row("Historico") & "'" 'Histórico & row("OrigemNota_id") & "-" & row("OrigemSerie_Id") & " PEDIDO  " & row("Pedido") & " DE " & row("Nome")
                                sql &= ", '" & _NotaFiscalRazao.CodigoEmpresa & "'"                         'Cliente = Empresa que efetuou o Serviço de Transporte
                                sql &= ", " & _NotaFiscalRazao.EnderecoEmpresa                              'Endereco = Empresa que efetuou o Serviço de Transporte 
                                sql &= ", '" & _NotaFiscalRazao.EntradaSaida.ToString.Substring(0, 1) & "'" 'E/S
                                sql &= ", '" & _NotaFiscalRazao.Serie & "'"                                 'Serie
                                sql &= ", " & _NotaFiscalRazao.Codigo                                       'Numero 
                                sql &= ", " & _NotaFiscalRazao.CodigoPedido.ToSqlNULL                                 'Numero 
                                sql &= ", 'P',0,"                                                           'Previsto/Realizado"
                                sql &= IIf(New Negocio.PlanoDeConta("", 0, Enc.OperacaoEncargo.CodigoCreditaConta).TemProduto, row("OrigemProduto_Id"), "NULL")  'Produto Transportado Milho, Soja
                                sql &= ",'" & HttpContext.Current.Session("ssNomeUsuario") & "', Convert(varchar, CURRENT_TIMESTAMP, 112));" & vbCrLf

                                Sqls.Add(sql)
                            End If
                        End If
                        'End If
                    Next
                Next
                'End If
            Next
        Catch ex As Exception
            Return False
        End Try

        Return Banco.GravaBanco(Sqls)
    End Function
#End Region

End Class