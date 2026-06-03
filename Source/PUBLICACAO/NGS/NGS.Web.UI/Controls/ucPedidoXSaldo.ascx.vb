Imports System.Data
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class ucPedidoXSaldo
    Inherits BaseUserControl

    Private objNotaFiscal As NotaFiscal
    'Private objSaldoPedidos As ListSaldoPedido
    Private objListSaldoPedidos As ListSaldoPedido2015
    Private objSaldoPedidoSelecionado As SaldoPedido2015


    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If ddlSafra.Items.Count = 0 Then ddl.Carregar(ddlSafra, CarregarDDL.Tabela.Safra, "")
        If Not IsPostBack And IsConnect Then
            txtDtNota.Text = Date.Now.ToString("dd/MM/yyyy")
            txtDtMovimento.Text = Date.Now.ToString("dd/MM/yyyy")
            txtDataNotaFix.Text = Date.Now.ToString("dd/MM/yyyy")
            txtDataMovimentoFix.Text = Date.Now.ToString("dd/MM/yyyy")
            txtDataNotaDevNN.Text = Date.Now.ToString("dd/MM/yyyy")
            txtDataMovimentoDevNN.Text = Date.Now.ToString("dd/MM/yyyy")

            If Not Session("ssTipoRetorno") Is Nothing AndAlso Session("ssTipoRetorno").ToString.Contains("objItensPedidoSelecionadosNXI") Then
                CarregarDadosNotaFiscalXItem()
            End If
        End If
    End Sub

    Public Overrides Sub SetarHID(ByVal NewGuid As String)
        HID.Value = NewGuid.ToString
    End Sub

#Region "Session"
    Private Sub SessaoSalvaNotaFiscal()
        Session("objNotaFiscal" & HID.Value) = objNotaFiscal
    End Sub

    Private Sub SessaoRecuperaNotaFiscal()
        objNotaFiscal = CType(Session("objNotaFiscal" & HID.Value), NotaFiscal)
    End Sub

    'Saldo dos Pedidos de Deposito
    Private Sub SessaoSalvaSaldoPedidos()
        Session("objListSaldoPedidos" & HID.Value) = objListSaldoPedidos
    End Sub

    Private Sub SessaoRecuperaSaldoPedidos()
        objListSaldoPedidos = CType(Session("objListSaldoPedidos" & HID.Value), ListSaldoPedido2015)
    End Sub

    'Pedidos Selecionado
    Private Sub SessaoSalvaSaldoPedidoSelecionado()
        Session("objSaldoPedidoSelecionado" & HID.Value) = objSaldoPedidoSelecionado
    End Sub

    Private Sub SessaoRecuperaSaldoPedidoSelecionado()
        objSaldoPedidoSelecionado = CType(Session("objSaldoPedidoSelecionado" & HID.Value), SaldoPedido2015)
    End Sub

    Private ReadOnly Property SessaoDsXML() As DataSet
        Get
            Return CType(Session("dsXML" & HID.Value), DataSet)
        End Get
    End Property

#End Region

    Protected Sub btnEmpresa_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Session("ssCampo") = "Livre"
        Popup.ConsultaDeEmpresas(Me.Page, "Saldo")
    End Sub

    Public Sub CarregarDadosNotaFiscalXItem()
        SessaoRecuperaNotaFiscal()

        Dim Empresa As New Cliente(objNotaFiscal.CodigoEmpresa, objNotaFiscal.EnderecoEmpresa)
        txtCodigoEmpresa.Value = Empresa.Codigo & "-" & Empresa.CodigoEndereco
        txtNomeEmpresa.Text = Funcoes.AlinharEsquerda(Empresa.Nome, 28, ".") &
                              " - " & Funcoes.AlinharEsquerda(Empresa.Cidade, 20, ".") & " " & Empresa.CodigoEstado &
                              " " & Funcoes.AlinharEsquerda(Funcoes.FormatarCpfCnpj(Empresa.Codigo), 18, ".") &
                              "-" & Empresa.CodigoEndereco.ToString() & "-" & Empresa.Reduzido
        txtNomeEmpresa.ReadOnly = True
        btnEmpresa.Enabled = False

        Dim Clie As New Cliente(objNotaFiscal.CodigoCliente, objNotaFiscal.EnderecoCliente)
        txtCodCliente.Value = Clie.Codigo & "-" & Clie.CodigoEndereco
        txtNomeCliente.Text = Funcoes.AlinharEsquerda(Clie.Nome, 28, ".") &
                              " - " & Funcoes.AlinharEsquerda(Clie.Cidade, 20, ".") & " " & Clie.CodigoEstado &
                              " " & Funcoes.AlinharEsquerda(Funcoes.FormatarCpfCnpj(Clie.Codigo), 18, ".") &
                              "-" & Clie.CodigoEndereco.ToString() & "-" & Clie.Reduzido
        txtNomeCliente.Enabled = False
        btnCliente.Enabled = False

        'Importação do Xml para notas de entrada
        If objNotaFiscal.Codigo > 0 AndAlso Not SessaoDsXML Is Nothing Then
            'PEDIDOS
            txtNota.Text = objNotaFiscal.Codigo
            txtSerie.Text = objNotaFiscal.Serie
            txtDtNota.Text = objNotaFiscal.DataNota

            txtNota.Enabled = False
            txtSerie.Enabled = False
            txtDtNota.Enabled = False

            'FIXAÇÃO
            txtNotaFix.Text = objNotaFiscal.Codigo
            txtSerieFix.Text = objNotaFiscal.Serie
            txtDataNotaFix.Text = objNotaFiscal.DataNota

            txtNotaFix.Enabled = False
            txtSerieFix.Enabled = False
            txtDataNotaFix.Enabled = False


            chkNossaEmissao.Checked = False
            chkNossaEmissao.Enabled = False

            primeiraVez.Value = "sim"

            Consultar(False)

        End If

        If objNotaFiscal.CodigoPedido > 0 Then
            txtPedido.Text = objNotaFiscal.CodigoPedido
            txtPedido.Enabled = False
        Else
            txtDataInicial.Text = Date.Now.AddMonths(-12).ToString("dd/MM/yyyy")
            txtDataFinal.Text = Date.Now.ToString("dd/MM/yyyy")
            txtPedido.Enabled = True
        End If
    End Sub

    Public Sub CarregarDados(ByVal parameters As Dictionary(Of String, Object))
        ddl.Carregar(ddlSafra, CarregarDDL.Tabela.Safra, "")

        If Not parameters("emp") Is Nothing Then
            Dim Empresa As New [Lib].Negocio.Cliente(parameters("emp"), parameters("ende"))
            txtCodigoEmpresa.Value = Empresa.Codigo & "-" & Empresa.CodigoEndereco

            txtNomeEmpresa.Text = Funcoes.AlinharEsquerda(Empresa.Nome, 28, ".") &
                                  " - " & Funcoes.AlinharEsquerda(Empresa.Cidade, 20, ".") & " " & Empresa.CodigoEstado &
                                  " " & Funcoes.AlinharEsquerda(Funcoes.FormatarCpfCnpj(Empresa.Codigo), 18, ".") &
                                  "-" & Empresa.CodigoEndereco.ToString() & "-" & Empresa.Reduzido
            txtNomeEmpresa.ReadOnly = True
            btnEmpresa.Enabled = False
        End If

        If Not parameters("cli") Is Nothing Then
            Dim Clie As New [Lib].Negocio.Cliente(parameters("cli"), parameters("endc"))
            txtCodCliente.Value = Clie.Codigo & "-" & Clie.CodigoEndereco

            txtNomeCliente.Text = Funcoes.AlinharEsquerda(Clie.Nome, 28, ".") &
                                  " - " & Funcoes.AlinharEsquerda(Clie.Cidade, 20, ".") & " " & Clie.CodigoEstado &
                                  " " & Funcoes.AlinharEsquerda(Funcoes.FormatarCpfCnpj(Clie.Codigo), 18, ".") &
                                  "-" & Clie.CodigoEndereco.ToString() & "-" & Clie.Reduzido
            txtNomeCliente.Enabled = False
            btnCliente.Enabled = False
        End If

        If Not parameters("id") Is Nothing Then
            txtPedido.Text = parameters("id")
            txtPedido.Enabled = False
        Else
            txtDataInicial.Text = Date.Now.AddMonths(-12).ToString("dd/MM/yyyy")
            txtDataFinal.Text = Date.Now.ToString("dd/MM/yyyy")
        End If

        If Not parameters("subope") Is Nothing AndAlso parameters("subope") = "N" Then
            ddlOperacao.Enabled = False
        End If

        txtDtNota.Text = Date.Now.ToString("dd/MM/yyyy")
        txtDtMovimento.Text = Date.Now.ToString("dd/MM/yyyy")
        txtDataNotaFix.Text = Date.Now.ToString("dd/MM/yyyy")
        txtDataMovimentoFix.Text = Date.Now.ToString("dd/MM/yyyy")
        txtDataNotaDevNN.Text = Date.Now.ToString("dd/MM/yyyy")
        txtDataMovimentoDevNN.Text = Date.Now.ToString("dd/MM/yyyy")
    End Sub

    Protected Sub btnCliente_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim ucConsultaClientes As ucConsultaClientes = DirectCast(Me.NamingContainer.FindControl("ucConsultaClientes"), ucConsultaClientes)
        If ucConsultaClientes IsNot Nothing Then
            ucConsultaClientes.Limpar()
        End If
        Popup.ConsultaDeClientes(Me, "objClietePXS", "txtNome")
    End Sub

    Function BuscaFixacoes() As DataSet
        Dim Emp() As String = txtCodigoEmpresa.Value.Split("-")
        Dim Cli() As String = txtCodCliente.Value.Split("-")
        Dim sql As String = ""
        sql = "SELECT P.Cliente," & vbCrLf &
              "       P.EndCliente," & vbCrLf &
              "       C.Nome," & vbCrLf &
              "       PxF.Pedido_Id as Pedido," & vbCrLf &
              "       M.Descricao as NomeMoeda," & vbCrLf &
              "       PxF.Movimento," & vbCrLf &
              "	      PxF.fixacao_Id as Fixacao," & vbCrLf &
              "	      PxF.Produto_Id + '-' + prd.Nome AS Produto," & vbCrLf &
              "       convert(varchar,PxF.Operacao) + '-' + convert(varchar,PxF.SubOperacao) + ' - ' + SO.Descricao as Operacao, " & vbCrLf &
              "       PxF.Quantidade," & vbCrLf &
              "       case" & vbCrLf &
              "          when M.Classificacao = 'O'" & vbCrLf &
              "            then UnitarioOficial" & vbCrLf &
              "            else UnitarioMoeda" & vbCrLf &
              "       end as Unitario," & vbCrLf &
              "       case" & vbCrLf &
              "          when M.Classificacao = 'O'" & vbCrLf &
              "            then TotalOficial" & vbCrLf &
              "            else TotalMoeda" & vbCrLf &
              "       end as Valor, " & vbCrLf &
              "       NotasXFixadas.Valor AS ValorNF," & vbCrLf &
              "       case" & vbCrLf &
              "          when M.Classificacao = 'O'" & vbCrLf &
              "            then TotalOficial" & vbCrLf &
              "            else TotalMoeda" & vbCrLf &
              "       end - NotasXFixadas.Valor AS Saldo" & vbCrLf &
              "" & vbCrLf &
              "  FROM VW_PedidosXItensXFixacoes PxF" & vbCrLf &
              "  LEFT JOIN (SELECT NFI.Empresa_Id, NFI.EndEmpresa_Id, NFI.Fixacao, NFI.Pedido, NFI.Produto_Id" & vbCrLf &
              "             FROM NotasFiscaisxItens NFI " & vbCrLf &
              "             INNER JOIN NotasFiscais NF " & vbCrLf &
              "                     ON NF.Empresa_Id    = NFI.Empresa_Id " & vbCrLf &
              "                    AND NF.EndEmpresa_Id = NFI.EndEmpresa_Id " & vbCrLf &
              "                    AND NF.Cliente_Id    = NFI.Cliente_Id " & vbCrLf &
              "                    AND NF.EndCliente_Id = NFI.EndCliente_Id " & vbCrLf &
              "                    AND NF.EntradaSaida_Id = NFI.EntradaSaida_Id " & vbCrLf &
              "                    AND NF.Serie_Id = NFI.Serie_Id " & vbCrLf &
              "                    AND NF.Nota_Id = NFI.Nota_Id " & vbCrLf &
              "             WHERE NF.Situacao in (1,4)" & vbCrLf &
              "               AND NF.Empresa_Id     ='" & Emp(0) & "'" & vbCrLf &
              "               AND NF.EndEmpresa_Id  = " & Emp(1) & vbCrLf

        If Not String.IsNullOrWhiteSpace(txtPedido.Text) AndAlso CInt(txtPedido.Text) > 0 Then
            sql &= "               AND NF.Pedido    = " & Trim(txtPedido.Text) & vbCrLf
        End If

        If chkConsolidarCliente.Checked Then
            sql &= "               AND left(NF.Cliente_Id,8)  ='" & Left(Cli(0), 8) & "') NotaXFixacao " & vbCrLf
        Else
            sql &= "               AND NF.Cliente_Id    ='" & Cli(0) & "'" & vbCrLf &
                   "               AND NF.EndCliente_Id = " & Cli(1) & ") NotaXFixacao " & vbCrLf
        End If

        sql &= "    ON PxF.Empresa_Id    = NotaXFixacao.Empresa_Id" & vbCrLf &
              "   AND PxF.EndEmpresa_Id = NotaXFixacao.EndEmpresa_Id" & vbCrLf &
              "   AND PxF.Pedido_Id     = NotaXFixacao.Pedido" & vbCrLf &
              "   AND PxF.Produto_Id    = NotaXFixacao.Produto_Id" & vbCrLf &
              "   and PxF.Fixacao_id    = NotaXFixacao.Fixacao" & vbCrLf &
              "  LEFT JOIN (SELECT FXN.Empresa_Id, FXN.EndEmpresa_Id, FXN.Pedido_Id, FXN.Fixacao_Id, " & vbCrLf &
              "                    FXN.Produto_Id, sum(isnull(FXN.Valor,0)) AS Valor " & vbCrLf &
              "             FROM FixacaoXNotaFiscal FXN " & vbCrLf &
              "             WHERE FXN.Empresa_Id     ='" & Emp(0) & "'" & vbCrLf &
              "               AND FXN.EndEmpresa_Id  = " & Emp(1) & vbCrLf

        If Not String.IsNullOrWhiteSpace(txtPedido.Text) AndAlso CInt(txtPedido.Text) > 0 Then
            sql &= "               AND FXN.Pedido_Id    = " & Trim(txtPedido.Text) & vbCrLf
        End If

        If chkConsolidarCliente.Checked Then
            sql &= "               AND left(FXN.Cliente_Id,8)  ='" & Left(Cli(0), 8) & "'" & vbCrLf
        Else
            sql &= "               AND FXN.Cliente_Id    ='" & Cli(0) & "'" & vbCrLf &
                   "               AND FXN.EndCliente_Id = " & Cli(1) & vbCrLf
        End If

        sql &= "             GROUP BY FXN.Empresa_Id, FXN.EndEmpresa_Id, FXN.Pedido_Id, FXN.Fixacao_Id, FXN.Produto_Id) NotasXFixadas " & vbCrLf &
              "    ON PxF.Empresa_Id    = NotasXFixadas.Empresa_Id" & vbCrLf &
              "   AND PxF.EndEmpresa_Id = NotasXFixadas.EndEmpresa_Id" & vbCrLf &
              "   AND PxF.Pedido_Id     = NotasXFixadas.Pedido_Id" & vbCrLf &
              "   and PxF.Fixacao_id    = NotasXFixadas.Fixacao_Id" & vbCrLf &
              "   AND PxF.Produto_Id    = NotasXFixadas.Produto_Id" & vbCrLf &
              " INNER JOIN SubOperacoes SO" & vbCrLf &
              "    ON PxF.Operacao    = SO.Operacao_Id" & vbCrLf &
              "   AND PxF.SubOperacao = SO.SubOperacoes_Id" & vbCrLf &
              "   AND SO.Classe       = '" & eClassesOperacoes.COMPLEMENTACOES.ToString & "'" & vbCrLf &
              " INNER JOIN PEDIDOS P" & vbCrLf &
              "    ON P.Empresa_Id     = PxF.Empresa_Id " & vbCrLf &
              "   AND P.EndEmpresa_Id  = PxF.EndEmpresa_Id " & vbCrLf &
              "   AND P.Pedido_Id      = PxF.Pedido_Id " & vbCrLf &
              " INNER JOIN Clientes C" & vbCrLf &
              "    ON C.Cliente_Id  = P.Cliente" & vbCrLf &
              "   AND C.Endereco_Id = P.EndCliente" & vbCrLf &
              " Inner Join Moedas M" & vbCrLf &
              "    ON M.Moeda_Id = P.Moeda" & vbCrLf &
              " Inner Join Produtos Prd" & vbCrLf &
              "    ON prd.Produto_Id =  PxF.Produto_Id" & vbCrLf &
              " WHERE isnull(NotaXFixacao.Fixacao,0) = 0" & vbCrLf &
              "   AND PxF.Empresa_Id     ='" & Emp(0) & "'" & vbCrLf &
              "   AND PxF.EndEmpresa_Id  = " & Emp(1) & vbCrLf &
              "   AND isnull(P.PedidoBloqueado,0) = 0 " & vbCrLf

        If Not String.IsNullOrWhiteSpace(txtPedido.Text) AndAlso CInt(txtPedido.Text) > 0 Then
            sql &= "  AND PxF.Pedido_Id    = " & Trim(txtPedido.Text) & vbCrLf
        End If

        If chkConsolidarCliente.Checked Then
            sql &= "   AND left(P.Cliente,8)  ='" & Left(Cli(0), 8) & "'" & vbCrLf
        Else
            sql &= "   AND P.Cliente          ='" & Cli(0) & "'" & vbCrLf &
                   "   AND P.EndCliente       = " & Cli(1) & vbCrLf
        End If

        sql &= "   AND ((case" & vbCrLf &
                "          when M.Classificacao = 'O'" & vbCrLf &
                "            then TotalOficial" & vbCrLf &
                "            else TotalMoeda" & vbCrLf &
                "        end < NotasXFixadas.Valor) or " & vbCrLf &
                "       (case" & vbCrLf &
                "          when M.Classificacao = 'O'" & vbCrLf &
                "            then TotalOficial" & vbCrLf &
                "            else TotalMoeda" & vbCrLf &
                "        end > NotasXFixadas.Valor))" & vbCrLf &
                " Order by PxF.Movimento, PxF.fixacao_Id"

        Dim ds As DataSet
        ds = Banco.ConsultaDataSet(sql, "Fixacoes")

        Return ds
    End Function

    Private Sub BuscaPedidos()
        objListSaldoPedidos = New ListSaldoPedido2015(ChecarParametros)

        If rdComSaldo.Checked Then

            gridGlobalDireta.DataSource = objListSaldoPedidos.Where(Function(s) s.Tipo = 1 And (s.SaldoQtdeDiretoFiscal > 0 Or s.SaldoQtdeGlobalFiscal > 0 Or s.SaldoQtdeRemessaFiscal > 0 Or s.SaldoValorOficialGlobalDireto > 0))
            GridAFixar.DataSource = objListSaldoPedidos.Where(Function(s) s.Tipo = 2 And s.SaldoQtdeAFixar > 0)
            gridDeposito.DataSource = objListSaldoPedidos.Where(Function(s) s.Tipo = 3 And s.QtdeEntregueFiscalDeposito > 0)

        ElseIf rdSemSaldo.Checked Then

            gridGlobalDireta.DataSource = objListSaldoPedidos.Where(Function(s) s.Tipo = 1 And s.SaldoQtdeDiretoFiscal = 0 And s.SaldoQtdeGlobalFiscal = 0 And s.SaldoQtdeRemessaFiscal = 0)
            GridAFixar.DataSource = objListSaldoPedidos.Where(Function(s) s.Tipo = 2 And s.SaldoQtdeAFixar = 0)
            gridDeposito.DataSource = objListSaldoPedidos.Where(Function(s) s.Tipo = 3 And s.QtdeEntregueFiscalDeposito = 0)

        Else

            gridGlobalDireta.DataSource = objListSaldoPedidos.Where(Function(s) s.Tipo = 1)
            GridAFixar.DataSource = objListSaldoPedidos.Where(Function(s) s.Tipo = 2)
            gridDeposito.DataSource = objListSaldoPedidos.Where(Function(s) s.Tipo = 3)

        End If

        gridGlobalDireta.DataBind()
        GridAFixar.DataBind()
        gridDeposito.DataBind()

        Dim i As Integer = 0
        While i < gridGlobalDireta.Rows.Count
            If CType(gridGlobalDireta.Rows(i).FindControl("hidFiscalAberto"), HiddenField).Value Then 'Value corresponde a fiscal aberto sim ou nao
                CType(gridGlobalDireta.Rows(i).FindControl("imgAbertoFechado"), ImageButton).ImageUrl = "~/Images/certo.jpg"
                CType(gridGlobalDireta.Rows(i).FindControl("imgAbertoFechado"), ImageButton).ToolTip = "Fiscal Aberto"
            Else
                CType(gridGlobalDireta.Rows(i).FindControl("imgAbertoFechado"), ImageButton).ImageUrl = "~/Images/erro.jpg"
                CType(gridGlobalDireta.Rows(i).FindControl("imgAbertoFechado"), ImageButton).ToolTip = "Fiscal Fechado"
            End If
            i += 1
        End While

        i = 0
        While i < GridAFixar.Rows.Count
            If CType(GridAFixar.Rows(i).FindControl("hidFiscalAberto"), HiddenField).Value Then 'Value corresponde a fiscal aberto sim ou nao
                CType(GridAFixar.Rows(i).FindControl("imgAbertoFechado"), ImageButton).ImageUrl = "~/Images/certo.jpg"
                CType(GridAFixar.Rows(i).FindControl("imgAbertoFechado"), ImageButton).ToolTip = "Fiscal Aberto"
            Else
                CType(GridAFixar.Rows(i).FindControl("imgAbertoFechado"), ImageButton).ImageUrl = "~/Images/erro.jpg"
                CType(GridAFixar.Rows(i).FindControl("imgAbertoFechado"), ImageButton).ToolTip = "Fiscal Fechado"
            End If
            i += 1
        End While

        i = 0
        While i < gridDeposito.Rows.Count
            If CType(gridDeposito.Rows(i).FindControl("hidFiscalAberto"), HiddenField).Value Then 'Value corresponde a fiscal aberto sim ou nao
                CType(gridDeposito.Rows(i).FindControl("imgAbertoFechado"), ImageButton).ImageUrl = "~/Images/certo.jpg"
                CType(gridDeposito.Rows(i).FindControl("imgAbertoFechado"), ImageButton).ToolTip = "Fiscal Aberto"
            Else
                CType(gridDeposito.Rows(i).FindControl("imgAbertoFechado"), ImageButton).ImageUrl = "~/Images/erro.jpg"
                CType(gridDeposito.Rows(i).FindControl("imgAbertoFechado"), ImageButton).ToolTip = "Fiscal Fechado"
            End If
            i += 1
        End While
        SessaoSalvaSaldoPedidos()
    End Sub

    Public Function ChecarParametros(Optional pTipoPedido As String = "") As Hashtable
        '*******************************
        '***** PARAMETROS TRATATOS *****
        '*******************************
        'TipoPedido        nvarchar(5)  = NULL   1-Global/Direto 2-Afixar 3-Deposito
        'TipoApuracao      bit          = 0,     0 Sintetico - 1 Analitico        
        'Empresa           nvarchar(18) = NULL,
        'EndEmpresa        int          = NULL, 
        'Cliente           nvarchar(18) = NULL,        
        'EndCliente        int          = NULL, 
        'FilialDev         nvarchar(18) = NULL, 
        'EndFilialDev      int          = NULL, 
        'DataReferencia    datetime     = NULL, 
        'Safra             nvarchar(50) = NULL, 
        'Pedido            int          = NULL, 
        'Saldo             int          = NULL, -- 0 Sem Saldo, 1 Com Saldo, NULL Todos 
        'Fiscal            int          = NULL, -- 0 Fechado  , 1 Aberto   , NULL Todos
        'PeriodoInicial    datetime     = NULL,
        'PeriodoFinal      datetime     = NULL,
        'Operacao          int          = NULL,
        'SubOperacao       int          = NULL,
        'Produto           int          = NULL


        SessaoRecuperaNotaFiscal()

        Dim parametros As New Hashtable

        If pTipoPedido.Length > 0 Then parametros.Add("TipoPedido", pTipoPedido)

        If txtNomeEmpresa.Text.Length > 0 Then
            Dim x() As String = txtCodigoEmpresa.Value.Split("-")
            parametros.Add("Empresa", x(0))
            parametros.Add("EndEmpresa", x(1))
        End If

        If txtNomeCliente.Text.Length > 0 Then
            Dim x() As String = txtCodCliente.Value.Split("-")
            If chkConsolidarCliente.Checked Then
                parametros.Add("Cliente", Left(x(0), 8))
            Else
                parametros.Add("Cliente", x(0))
                parametros.Add("EndCliente", x(1))
            End If
        End If

        'DataInicial DataFinal As String - Trata-se da data em que o pedido foi realizado pode informar uma ou a outra caso informe as duas se torna um intervalo LOGICO
        If chkPeriodo.Checked Then
            If txtDataInicial.Text.Length > 0 Then parametros.Add("PeriodoInicial", CDate(txtDataInicial.Text).ToString("yyyy-MM-dd"))
            If txtDataFinal.Text.Length > 0 Then parametros.Add("PeriodoFinal", CDate(txtDataFinal.Text).ToString("yyyy-MM-dd"))
        End If

        If rdComSaldo.Checked Then
            parametros.Add("Saldo", 1)
        ElseIf rdSemSaldo.Checked Then
            parametros.Add("Saldo", 0)
        End If

        If rdFAberto.Checked Then
            parametros.Add("Fiscal", 1)
        ElseIf rdFFechado.Checked Then
            parametros.Add("Fiscal", 0)
        End If

        If ddlSafra.SelectedIndex > 0 Then
            parametros.Add("Safra", ddlSafra.SelectedValue)
        End If

        SessaoRecuperaNotaFiscal()

        If txtPedido.Text.Length > 0 Then
            parametros.Add("Pedido", txtPedido.Text)
        ElseIf Not SessaoDsXML Is Nothing AndAlso Not objNotaFiscal Is Nothing AndAlso Not objNotaFiscal.NCMXML Is Nothing AndAlso objNotaFiscal.NCMXML.Length > 0 Then
            If primeiraVez.Value = "sim" Then
                parametros.Add("NCM", objNotaFiscal.NCMXML)
                parametros.Add("ParteNomeDoProduto", objNotaFiscal.ParteNomeProdNCMXML)
                primeiraVez.Value = ""
            Else
                parametros.Add("NCM", "")
                parametros.Add("ParteNomeDoProduto", "")
            End If
        End If

        Return parametros
    End Function

    Protected Sub ddlOperacao_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ddlOperacao.SelectedIndexChanged
        SessaoRecuperaSaldoPedidoSelecionado()

        Dim sb As String() = ddlOperacao.SelectedValue.Split("-")
        Dim subope As New SubOperacao(sb(0), sb(1))

        If objSaldoPedidoSelecionado.Empresa.Empresa.PedidoBloqueado Then
            Dim opPed As New SubOperacao(objSaldoPedidoSelecionado.CodigoOperacao, objSaldoPedidoSelecionado.CodigoSubOperacao)

            If subope.Pedido AndAlso subope.Codigo <> opPed.Codigo AndAlso subope.Classe <> eClassesOperacoes.GLOBAL Then
                ddlOperacao.SelectedValue = opPed.CodigoOperacao & "-" & opPed.Codigo
                subope = New SubOperacao(sb(0), sb(1))
                ddlFilial.SelectedValue = objSaldoPedidoSelecionado.CodigoCliente & "-" & objSaldoPedidoSelecionado.EndCliente
                ddlFilial.Enabled = False
                MsgBox(Me.Page, "Operação selecionada é de Pedido, portando não pode ser diferente da lançada no Pedido.")
            End If
        End If

        If subope.Devolucao And ddlFilial.Items.Count > 1 Then
            ddlFilial.Enabled = True
            MsgBox(Me.Page, "Este Pedido fez transações com várias filiais. Selecione a filial para devolução e os saldos serão apurados de acordo com a mesma.")
        End If

        Dim fil As String() = ddlFilial.SelectedValue.Split("-")

        objSaldoPedidoSelecionado.RecarregarItens(subope.Classe.ToString, subope.Devolucao, Not subope.QuantidadeFiscal, 0, "", IIf(subope.Devolucao, fil(0), ""), IIf(subope.Devolucao, fil(1), 0))

        If gridGlobalDiretaItens.Visible Then
            gridGlobalDiretaItens.DataSource = objSaldoPedidoSelecionado.Itens
            gridGlobalDiretaItens.DataBind()
        End If

        If gridDepositoItens.Visible Then
            gridDepositoItens.DataSource = objSaldoPedidoSelecionado.Itens
            gridDepositoItens.DataBind()
        End If

        If GridAFixar.Visible Then
            GridAFixar.DataSource = objSaldoPedidoSelecionado.Itens
            GridAFixar.DataBind()
        End If

        SessaoSalvaSaldoPedidoSelecionado()
        MudancaDeOperacao()
    End Sub

    Private Sub MudancaDeOperacao()
        Dim op() As String = ddlOperacao.SelectedValue.Split("-")
        Dim subope As New SubOperacao(op(0), op(1))
        chkNossaEmissao.Enabled = True

        SessaoRecuperaNotaFiscal()

        If Not objNotaFiscal Is Nothing Then
            chkNossaEmissao.Checked = True
            If subope.EntradaSaida = eEntradaSaida.Saida Then
                If objNotaFiscal.Empresa.Empresa.NossaEmissao Then chkNossaEmissao.Enabled = False
            Else
                chkNossaEmissao.Checked = False
            End If
            chkNossaEmissao_CheckedChanged(Nothing, Nothing)
            SessaoSalvaNotaFiscal()
        End If

        If subope.Devolucao And gridGlobalDiretaItens.Visible Then
            gridGlobalDiretaItens.Columns(11).Visible = False
            gridGlobalDiretaItens.Columns(10).Visible = False
        Else
            gridGlobalDiretaItens.Columns(11).Visible = True
            gridGlobalDiretaItens.Columns(10).Visible = True
        End If
    End Sub

    Protected Sub ddlSafra_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        If ddlSafra.SelectedIndex = 0 Then
            chkPeriodo.Checked = True
        Else
            chkPeriodo.Checked = False
        End If
        pnlPeriodo.Visible = chkPeriodo.Checked
    End Sub

    Protected Sub chkNossaEmissao_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkNossaEmissao.CheckedChanged
        If SessaoDsXML IsNot Nothing Then
            chkNossaEmissao.Checked = False
            chkNossaEmissao.Enabled = False
            Exit Sub
        End If

        pnlNossaEmissao.Visible = Not chkNossaEmissao.Checked

        Dim retroAtiva As Boolean = False

        If Funcoes.VerificaPermissao("NotaFiscalRetroativa", "Gravar") Then
            chkNotaRetroativa.Visible = chkNossaEmissao.Checked
            chkNotaRetroativa.Checked = False

            retroAtiva = True
        Else
            chkNotaRetroativa.Visible = False
            chkNotaRetroativa.Checked = False
        End If

        If chkNossaEmissao.Checked Then
            txtNota.Text = String.Empty
            txtSerie.Text = String.Empty

            If Not retroAtiva Then
                txtDtNota.Text = Now().Date()
                txtDtMovimento.Text = Now().Date()

                txtDtNota.Enabled = False
                txtDtMovimento.Enabled = False
            Else
                txtDtNota.Enabled = True
                txtDtMovimento.Enabled = True

                If CDate(txtDtMovimento.Text) < Now().AddDays(-10) Then
                    txtDtNota.Text = Now().Date()
                    txtDtMovimento.Text = Now().Date()
                    MsgBox(Me.Page, "Data do Movimento não pode ser inferior à 10 dias.")
                ElseIf CDate(txtDtMovimento.Text) < Now().Date() Then
                    MsgBox(Me.Page, "Data do Movimento inferior à data Atual, caso tenha certeza pode continuar.")
                End If
            End If
        Else
            txtDtNota.Enabled = True
            txtDtMovimento.Enabled = True
        End If
    End Sub

    Protected Sub chkNossaEmissaoFix_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkNossaEmissaoFix.CheckedChanged
        If chkNossaEmissaoFix.Checked Then
            divNossaEmissao.Visible = False

            If CDate(txtDataMovimentoFix.Text) < Now().AddDays(-10) Then
                txtDataNotaFix.Text = Now().Date()
                txtDataMovimentoFix.Text = Now().Date()
                MsgBox(Me.Page, "Data do Movimento não pode ser inferior à 10 dias.")
            ElseIf CDate(txtDataMovimentoFix.Text) < Now().Date() Then
                MsgBox(Me.Page, "Data do Movimento inferior à data Atual, caso tenha certeza pode continuar.")
            End If
        Else
            divNossaEmissao.Visible = True
        End If
    End Sub

    Protected Sub chkNossaEmissaoDevNN_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkNossaEmissaoDevNN.CheckedChanged
        If chkNossaEmissaoDevNN.Checked Then
            divNossaEmissaoDevNN.Visible = False

            If CDate(txtDataMovimentoDevNN.Text) < Now().AddDays(-10) Then
                txtDataNotaDevNN.Text = Now().Date()
                txtDataMovimentoDevNN.Text = Now().Date()
                MsgBox(Me.Page, "Data do Movimento não pode ser inferior à 10 dias.")
            ElseIf CDate(txtDataMovimentoDevNN.Text) < Now().Date() Then
                MsgBox(Me.Page, "Data do Movimento inferior à data Atual, caso tenha certeza pode continuar.")
            End If
        Else
            divNossaEmissaoDevNN.Visible = True
        End If
    End Sub

    Protected Sub btnConfirmar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnConfirmar.Click

        SessaoRecuperaNotaFiscal()

        'Validação para gerar notas fiscais retroativas'
        If chkNotaRetroativa.Checked AndAlso Not Funcoes.VerificaPermissao("NotaFiscalRetroativa", "Gravar") Then
            MsgBox(Me.Page, "Usuário sem permissão para gerar Nota Fiscal Retroativa - (Processo: NotaFiscalRetroativa, Permissão: Gravar)")
            Exit Sub
        End If

        If CDate(txtDtMovimento.Text) < Now().Date() Then
            MsgBox(Me.Page, "Data do movimento inferior à data atual, caso tenha certeza pode continuar!")
        End If

        SessaoRecuperaSaldoPedidoSelecionado()

        If Not objSaldoPedidoSelecionado Is Nothing _
        AndAlso objSaldoPedidoSelecionado.Pedido.Troca _
        AndAlso objSaldoPedidoSelecionado.Pedido.CodigoPedidoTroca = 0 Then
            MsgBox(Me.Page, "Este pedido está marcado como troca porém não está vinculado a outro pedido. Caso este não seja mais de troca então desmarque a opção troca.")
        End If

        If Not objSaldoPedidoSelecionado Is Nothing _
        AndAlso objSaldoPedidoSelecionado.Pedido.Troca _
        AndAlso objSaldoPedidoSelecionado.Pedido.CodigoPedidoTroca > 0 _
        AndAlso ((objSaldoPedidoSelecionado.Pedido.CodigoMoeda = eTiposMoeda.Oficial AndAlso Math.Abs(objSaldoPedidoSelecionado.Pedido.Itens.LiquidoOficial - objSaldoPedidoSelecionado.Pedido.PedidoTroca.Itens.LiquidoOficial) > 0.01) _
                  OrElse (objSaldoPedidoSelecionado.Pedido.CodigoMoeda = eTiposMoeda.MoedaEstrangeira AndAlso Math.Abs(objSaldoPedidoSelecionado.Pedido.Itens.LiquidoOficial - objSaldoPedidoSelecionado.Pedido.PedidoTroca.Itens.LiquidoOficial) > 0.1)) Then

            MsgBox(Me.Page, String.Format("Quando o pedido é de Troca o Valor Líquido total dos itens dos pedidos devem ser iguais. Pedido: {0} Valor = {1} e Pedido de Troca: {2} Valor = {3}",
                          objSaldoPedidoSelecionado.Pedido.Codigo, objSaldoPedidoSelecionado.Pedido.Itens.LiquidoOficial, objSaldoPedidoSelecionado.Pedido.PedidoTroca.Codigo, objSaldoPedidoSelecionado.Pedido.PedidoTroca.Itens.LiquidoOficial))

            Exit Sub
        End If

        Dim tipo As Integer = 0
        If gridGlobalDiretaItens.Visible Then

            tipo = 1

            If SessaoDsXML IsNot Nothing Then

                Dim iCountProdutoSelecionadoGlobal As Integer = 0
                For Each row As GridViewRow In gridGlobalDiretaItens.Rows
                    Dim chk As CheckBox = CType(row.FindControl("chkSelecionado"), CheckBox)
                    If chk IsNot Nothing AndAlso chk.Checked Then
                        iCountProdutoSelecionadoGlobal += 1
                    End If
                Next

                If iCountProdutoSelecionadoGlobal > objNotaFiscal.Itens.Count() Then
                    MsgBox(Me.Page, "A quantidade de itens selecionado do pedido é maior que os itens importados do XML!")
                    Exit Sub
                End If

            End If

            For Each row As GridViewRow In gridGlobalDiretaItens.Rows
                    Dim chk As CheckBox = CType(row.FindControl("chkSelecionado"), CheckBox)
                    If chk IsNot Nothing Then
                        objSaldoPedidoSelecionado.Itens(row.RowIndex).Selecionado = chk.Checked

                        'Importação do xml para notas de entrada
                        If PreencherDadosDoXml(row, chk) = False Then
                            Exit Sub
                        End If

                        '------
                    End If
                Next
            End If

            If gridDepositoItens.Visible Then

            tipo = 2

            If SessaoDsXML IsNot Nothing Then

                Dim iCountProdutoSelecionadoDeposito As Integer = 0
                For Each row As GridViewRow In gridDepositoItens.Rows
                    Dim chk As CheckBox = CType(row.FindControl("chkSelecionado"), CheckBox)
                    If chk IsNot Nothing AndAlso chk.Checked Then
                        iCountProdutoSelecionadoDeposito += 1
                    End If
                Next

                If iCountProdutoSelecionadoDeposito > objNotaFiscal.Itens.Count() Then
                    MsgBox(Me.Page, "A quantidade de itens selecionado do pedido é maior que os itens importados do XML!")
                    Exit Sub
                End If

            End If

            For Each row As GridViewRow In gridDepositoItens.Rows
                Dim chk As CheckBox = CType(row.FindControl("chkSelecionado"), CheckBox)
                If chk IsNot Nothing Then
                    objSaldoPedidoSelecionado.Itens(row.RowIndex).Selecionado = chk.Checked
                    'Importação do xml para notas de entrada
                    If PreencherDadosDoXml(row, chk) = False Then
                        Exit Sub
                    End If
                    '-----
                End If
            Next
        End If

        If GridAFixarItens.Visible Then

            tipo = 3

            If SessaoDsXML IsNot Nothing Then

                Dim iCountProdutoSelecionadoFixar As Integer = 0
                For Each row As GridViewRow In GridAFixarItens.Rows
                    Dim chk As CheckBox = CType(row.FindControl("chkSelecionado"), CheckBox)
                    If chk IsNot Nothing AndAlso chk.Checked Then
                        iCountProdutoSelecionadoFixar += 1
                    End If
                Next

                If iCountProdutoSelecionadoFixar > objNotaFiscal.Itens.Count() Then
                    MsgBox(Me.Page, "A quantidade de itens selecionado do pedido é maior que os itens importados do XML!")
                    Exit Sub
                End If

            End If

            For Each row As GridViewRow In GridAFixarItens.Rows
                Dim chk As CheckBox = CType(row.FindControl("chkSelecionado"), CheckBox)
                If chk IsNot Nothing Then
                    objSaldoPedidoSelecionado.Itens(row.RowIndex).Selecionado = chk.Checked
                    'Importação do xml para notas de entrada
                    If PreencherDadosDoXml(row, chk) = False Then
                        Exit Sub
                    End If
                    '-----
                End If
            Next
        End If


        If chkTodasOperacoes.Checked Then
            Dim sb As String() = ddlOperacao.SelectedValue.Split("-")
            Dim Parametros As New OperacaoXEstado

            Parametros.Empresa = Left(objSaldoPedidoSelecionado.CodigoEmpresa, 8)
            Parametros.CodigoGrupoProduto = objSaldoPedidoSelecionado.Itens(0).Produto.CodigoGrupo
            Parametros.CodigoProduto = objSaldoPedidoSelecionado.Itens(0).CodigoProduto
            Parametros.CodigoOperacao = sb(0)
            Parametros.CodigoSubOperacao = sb(1)
            Parametros.EstadoOrigem = objSaldoPedidoSelecionado.Empresa.Municipio.CodigoEstado
            Parametros.EstadoDestino = objSaldoPedidoSelecionado.Cliente.Municipio.CodigoEstado

            Dim OxE = New OperacaoXEstado(Parametros)
            PreencherXmlCFOPSaidaXCFOPEntrada(objSaldoPedidoSelecionado.Itens(0).XmlCFOP, OxE.CodigoFiscal)
        End If


        SessaoSalvaSaldoPedidoSelecionado()

        ''*****************************************************************************
        ''*******************************  FIXACAO   **********************************
        ''*****************************************************************************
        If divFixacoes.Visible Then
            If gridFixacao.SelectedIndex < 0 Then
                MsgBox(Me.Page, "Selecione Uma Fixação para Continuar")
                Exit Sub
            End If

            If Not IsDate(txtDataNotaFix.Text) Or Not IsDate(txtDataMovimentoFix.Text) Then
                MsgBox(Me.Page, "Informe A Data da Nota de Fixação e a data de Entrada Saída da Nota de Fixação.")
                Exit Sub
            ElseIf CDate(txtDataMovimentoFix.Text) > Now() Then
                MsgBox(Me.Page, "Data do Movimento não pode ser superior à data atual.")
                Exit Sub
            ElseIf CDate(txtDataNotaFix.Text) > Now Then
                MsgBox(Me.Page, "Data da Nota não pode ser superior à data atual.")
                Exit Sub
            ElseIf Not chkNossaEmissaoFix.Checked AndAlso String.IsNullOrWhiteSpace(txtNotaFix.Text) Then
                MsgBox(Me.Page, "Nota Fiscal não foi informada.")
                Exit Sub
            ElseIf Not chkNossaEmissaoFix.Checked AndAlso Not CInt(txtNotaFix.Text) > 0 Then
                MsgBox(Me.Page, "Número da Nota Fiscal não pode ser 0.")
                Exit Sub
            ElseIf String.IsNullOrWhiteSpace(txtSerieFix.Text) AndAlso txtSerieFix.Visible Then
                MsgBox(Me.Page, "Série da Nota Fiscal não foi informada.")
                Exit Sub
            ElseIf ddlOperacaoFixacao.SelectedIndex = 0 Then
                MsgBox(Me.Page, "Operação da Fixação não foi selecionada.")
                Exit Sub
            End If

            Dim op() As String
            op = ddlOperacaoFixacao.SelectedValue.ToString.Split("-")


            If objNotaFiscal IsNot Nothing Then
                objNotaFiscal.CodigoPedido = gridFixacao.SelectedRow.Cells(2).Text
                objNotaFiscal.CodigoOperacao = op(0)
                objNotaFiscal.CodigoSubOperacao = op(1)

                objNotaFiscal.DataNota = CDate(txtDataNotaFix.Text)
                objNotaFiscal.Movimento = CDate(txtDataMovimentoFix.Text)

                objNotaFiscal.Operacao = Nothing
                objNotaFiscal.SubOperacao = Nothing

                If objNotaFiscal.SubOperacao.EntradaSaida = eEntradaSaida.Entrada Then
                    If chkNossaEmissaoFix.Checked Then
                        objNotaFiscal.NossaEmissao = True
                    Else
                        If txtNotaFix.Text.Trim.Length = 0 Or txtSerieFix.Text.Trim.Length = 0 Then
                            MsgBox(Me.Page, "Informe o número e série da nota de entrada da fixação!")
                            Exit Sub
                        End If
                        objNotaFiscal.NotaProdutor = txtNotaFix.Text
                        objNotaFiscal.SerieNotaProdutor = Trim(txtSerieFix.Text)
                    End If
                End If
                SessaoSalvaNotaFiscal()
            End If

            ' Temporário, até fazer o tombo em todos os clientes para virtual - Furlan - 21/06/2016
            If FinanceiroVirtual AndAlso objNotaFiscal.SubOperacao.Financeiro AndAlso objNotaFiscal.Pedido.Parcelas.Count = 0 Then
                MsgBox(Me.Page, "Pedido deve ser atualizado para Financeiro Virtual, qualquer dúvida entre em contato com o suporte!")
                Exit Sub
            End If

            If Not Session("ssTipoRetorno") Is Nothing AndAlso Session("ssTipoRetorno").ToString.Contains("objItensPedidoSelecionadosNXI") Then
                Session("objFixacao" & HID.Value) = gridFixacao.SelectedRow.Cells(2).Text & ";" & gridFixacao.SelectedRow.Cells(5).Text & ";" & gridFixacao.SelectedRow.Cells(6).Text & ";"
                If TypeOf Me.Page Is NotaFiscalXItens Then
                    CType(Me.Page, NotaFiscalXItens).CarregarFixacao()
                End If
            End If

            Session.Remove("ProdutosPedido" & HID.Value)
            Popup.CloseDialog(Me.Page, "divPedidoXSaldo")
            Exit Sub
        End If

        ''*****************************************************************************
        ''*******************************  PEDIDOS   **********************************
        ''*****************************************************************************

        If DadosPedidoXItens(tipo) Then
            ' Temporário, até fazer o tombo em todos os clientes para virtual - Furlan - 21/06/2016
            If FinanceiroVirtual Then
                SessaoRecuperaNotaFiscal()

                If objNotaFiscal.SubOperacao.Financeiro AndAlso objNotaFiscal.Pedido.Parcelas.Count = 0 Then
                    MsgBox(Me.Page, "Pedido deve ser atualizado para Financeiro Virtual, qualquer dúvida entre em contato com o suporte!")
                    Exit Sub
                End If
            End If

            Selecionar(objSaldoPedidoSelecionado)
        End If
    End Sub

    Private Function PreencherDadosDoXml(ByVal row As GridViewRow, ByVal chk As CheckBox) As Boolean

        'Importação do xml para notas de entrada

        If SessaoDsXML IsNot Nothing Then
            If chk.Checked AndAlso chk.Enabled Then
                If CType(row.FindControl("ddlXmlProdutoXDePara"), DropDownList).SelectedIndex > 0 Then
                    Dim ProdutoDe = CType(row.FindControl("ddlXmlProdutoXDePara"), DropDownList).SelectedValue
                    PreencherXmlProdutoXDePara(row.RowIndex, ProdutoDe)

                    Dim iRowIndex As Integer = 0
                    For Each itemNF In objNotaFiscal.Itens
                        If itemNF.ProdutoXML = ProdutoDe Then
                            Exit For
                        End If
                        iRowIndex += 1
                    Next

                    If PreencherCamposXmlDoItemSelecionado(iRowIndex) = False Then
                        Return False
                    End If

                Else
                    MsgBox(Me.Page, "Existe(m) produtos da NF importada que não foram vínculados com seus correspondentes NGS.")
                    Return False

                End If
            ElseIf chk.Checked Then
                If PreencherCamposXmlDoItemSelecionado(row.RowIndex) = False Then
                    Return False
                End If
            End If
        End If

        Return True

    End Function

    Private Function DadosPedidoXItens(tipo As Integer) As Boolean
        Try

            If CDate(txtDtMovimento.Text) > Now() Then
                MsgBox(Me.Page, "Data do Movimento não pode ser superior à data atual.")
                Return False
            ElseIf CDate(txtDtNota.Text) > Now Then
                MsgBox(Me.Page, "Data da Nota não pode ser superior à data atual.")
                Return False
            ElseIf Not chkNossaEmissao.Checked AndAlso String.IsNullOrWhiteSpace(txtNota.Text) Then
                MsgBox(Me.Page, "Nota Fiscal não foi informada.")
                Return False
            ElseIf Not chkNossaEmissao.Checked AndAlso Not CInt(txtNota.Text) > 0 Then
                MsgBox(Me.Page, "Número da Nota Fiscal não pode ser 0.")
                Return False
            ElseIf Not chkNossaEmissao.Checked AndAlso String.IsNullOrWhiteSpace(txtSerie.Text) Then
                MsgBox(Me.Page, "Série da Nota Fiscal não foi informada.")
                Return False
            ElseIf gridGlobalDiretaItens.Visible AndAlso gridGlobalDireta.SelectedRow Is Nothing Then
                MsgBox(Me.Page, "É necessário preencher os campos obrigatórios!")
                Return False
            ElseIf gridGlobalDiretaItens.Visible AndAlso gridGlobalDireta.SelectedRow IsNot Nothing AndAlso CType(gridGlobalDireta.Rows(gridGlobalDireta.SelectedRow.RowIndex).FindControl("hidFiscalAberto"), HiddenField).Value = "False" Then
                MsgBox(Me.Page, "Pedido com o Fiscal Fechado não pode ser usado para emissão de Nota Fiscal!")
                Return False

            ElseIf gridDepositoItens.Visible AndAlso gridDeposito.SelectedRow Is Nothing Then
                MsgBox(Me.Page, "É necessário preencher os campos obrigatórios!")
                Return False
            ElseIf gridDepositoItens.Visible AndAlso gridDeposito.SelectedRow IsNot Nothing AndAlso CType(gridDeposito.Rows(gridDeposito.SelectedRow.RowIndex).FindControl("hidFiscalAberto"), HiddenField).Value = "False" Then
                MsgBox(Me.Page, "Pedido com o Fiscal Fechado não pode ser usado para emissão de Nota Fiscal!")
                Return False

            ElseIf gridDepositoItens.Visible AndAlso gridDeposito.SelectedRow Is Nothing Then
                MsgBox(Me.Page, "É necessário preencher os campos obrigatórios!")
                Return False
            ElseIf gridDepositoItens.Visible AndAlso gridDeposito.SelectedRow IsNot Nothing AndAlso CType(gridDeposito.Rows(gridDeposito.SelectedRow.RowIndex).FindControl("hidFiscalAberto"), HiddenField).Value = "False" Then
                MsgBox(Me.Page, "Pedido com o Fiscal Fechado não pode ser usado para emissão de Nota Fiscal!")
                Return False
            End If

            Dim op() As String
            op = ddlOperacao.SelectedValue.Split("-")

            If Not IsDate(txtDtNota.Text) Or Not IsDate(txtDtMovimento.Text) Then
                MsgBox(Me.Page, "Informe A Data da Nota e a data de Entrada Saida da Nota!")
                Return False
            End If

            SessaoRecuperaNotaFiscal()
            If objNotaFiscal IsNot Nothing Then
                objNotaFiscal.CodigoPedido = CInt(lblPedidoSelecionado.Text)
                objNotaFiscal.DataNota = CDate(txtDtNota.Text)
                objNotaFiscal.Movimento = CDate(txtDtMovimento.Text)
                objNotaFiscal.CodigoOperacao = op(0)
                objNotaFiscal.CodigoSubOperacao = op(1)
                objNotaFiscal.Operacao = Nothing
                objNotaFiscal.SubOperacao = Nothing

                If objNotaFiscal IsNot Nothing AndAlso objNotaFiscal.SubOperacao.EntradaSaida = eEntradaSaida.Entrada Then
                    If chkNossaEmissao.Checked Then
                        objNotaFiscal.Retroativa = chkNotaRetroativa.Checked
                        objNotaFiscal.NossaEmissao = True
                    Else
                        If txtNota.Text.Trim.Length = 0 Or txtSerie.Text.Trim.Length = 0 Then
                            MsgBox(Me.Page, "Informe o número e série da nota de entrada!")
                            Return False
                        End If
                        objNotaFiscal.NotaProdutor = txtNota.Text
                        objNotaFiscal.SerieNotaProdutor = Trim(txtSerie.Text)
                    End If
                Else
                    objNotaFiscal.Retroativa = chkNotaRetroativa.Checked
                    objNotaFiscal.NossaEmissao = chkNossaEmissao.Checked
                End If

                'Nota Fiscal Retroativa'
                If objNotaFiscal.Retroativa Then
                    Dim resp As String = objNotaFiscal.ValidaNotaRetroativa()
                    If Not String.IsNullOrWhiteSpace(resp) Then
                        MsgBox(Me.Page, resp)
                        Return False
                    End If
                End If
            End If

            SessaoRecuperaSaldoPedidos()

            If objSaldoPedidoSelecionado.Itens.Where(Function(s) s.Selecionado).Count > 0 Then
                objSaldoPedidoSelecionado.CodigoOperacao = op(0)
                objSaldoPedidoSelecionado.CodigoSubOperacao = op(1)
                Session(Session("ssTipoRetorno")) = objSaldoPedidoSelecionado
            Else
                MsgBox(Me.Page, "Produto não foi selecionado!")
                Return False
            End If

            If gridGlobalDiretaItens.Visible Then
                Dim objSubOperacao As SubOperacao = New SubOperacao(op(0), op(1))
                If objSubOperacao.Devolucao AndAlso Not objSubOperacao.QuantidadeFiscal Then
                    For Each sel In objSaldoPedidoSelecionado.Itens.Where(Function(s) s.Selecionado)
                        'If sel.VlrNotaOficialGlobalLiquido > 0 AndAlso sel.SaldoValorOficialGlobalDireto = 0 Then
                        '    MsgBox(Me.Page, "O produto " & sel.NomeProduto & " da filial selecionada não têm saldo para devolução!" & " Saldo: " & sel.SaldoValorOficialGlobalDireto.ToString("N2"))
                        '    Return False
                        'ElseIf sel.VlrNotaOficialDiretaLiquido > 0 AndAlso sel.SaldoValorOficialGlobalDireto = 0 Then
                        '    MsgBox(Me.Page, "O produto " & sel.NomeProduto & " da filial selecionada não têm saldo para devolução!" & " Saldo: " & sel.SaldoValorOficialGlobalDireto.ToString("N2"))
                        '    Return False
                        'End If
                        If Not sel.VlrNotaOficialDiretaLiquido > 0 Then
                            MsgBox(Me.Page, "O produto " & sel.NomeProduto & " da filial selecionada não têm saldo para devolução!" & " Saldo: " & sel.SaldoValorOficialGlobalDireto.ToString("N2"))
                            Return False
                            'ElseIf Not sel.VlrNotaOficialGlobalLiquido > 0 Then
                            '    MsgBox(Me.Page, "O produto " & sel.NomeProduto & " da filial selecionada não têm saldo para devolução!" & " Saldo: " & sel.SaldoValorOficialGlobalDireto.ToString("N2"))
                            '    Return False
                        End If
                    Next
                End If
            End If

            Dim Filial As String() = ddlFilial.SelectedValue.Split("-")
            objNotaFiscal.CodigoCliente = Filial(0)
            objNotaFiscal.EnderecoCliente = Filial(1)

            'Verificar se a Nota já foi lançada
            If Not objNotaFiscal.NossaEmissao Then
                Dim Mensagem As String = objNotaFiscal.NotaFiscalJaLancada()
                If Not String.IsNullOrWhiteSpace(Mensagem) Then
                    MsgBox(Me.Page, Mensagem)
                    Return False
                End If
            End If

            SessaoSalvaNotaFiscal()

            Return True

        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Function

    Private Function NotaFiscalJaLancada(ByVal objNF As NotaFiscal) As String
        SessaoRecuperaNotaFiscal()
        Dim sql As String = String.Empty

        sql = "SELECT ISNULL(NFG,0) NFG " & vbCrLf &
              "  FROM NOTASFISCAIS" & vbCrLf &
              " WHERE Cliente_Id = '" & objNF.CodigoCliente & "'" & vbCrLf &
              "   AND EndCliente_Id = " & objNF.EnderecoCliente & vbCrLf &
              "   AND EntradaSaida_Id = '" & IIf(objNF.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "' " & vbCrLf &
              "   AND Serie_Id= '" & objNF.SerieNotaProdutor & "'" & vbCrLf &
              "   AND Nota_Id= " & objNF.NotaProdutor & vbCrLf
        If objNF.NossaEmissao Then
            sql &= "   AND Empresa_Id = '" & objNF.CodigoEmpresa & "'" & vbCrLf &
                   "   AND EndEmpresa_Id = " & objNF.EnderecoEmpresa & vbCrLf
        End If
        Dim Ds As DataSet = Banco.ConsultaDataSet(sql, "NF")

        If Ds.Tables(0).Rows.Count > 0 Then
            If Ds.Tables(0).Rows(0)("NFG") Then
                Return "Existe uma Nota Fiscal Geral já emitida com esse Número"
            Else
                Return "Existe uma Nota Fiscal já emitida com esse Número"
            End If
        Else
            Return ""
        End If
    End Function

    Public Sub LimparCampos()
        txtNomeEmpresa.Text = ""
        txtNomeCliente.Text = ""
        txtPedido.Text = ""

        lblSafraSelecionada.Text = String.Empty
        lblPedidoSelecionado.Text = String.Empty
        divProdutosDoPedido.Visible = False
        btnConfirmar.Visible = False

        txtDtNota.Text = Date.Now.ToString("dd/MM/yyyy")
        txtDtMovimento.Text = Date.Now.ToString("dd/MM/yyyy")
        txtDataNotaFix.Text = Date.Now.ToString("dd/MM/yyyy")
        txtDataMovimentoFix.Text = Date.Now.ToString("dd/MM/yyyy")
        rdComSaldo.Checked = True
        rdFAberto.Checked = True
        chkConsolidarCliente.Enabled = True
        chkConsolidarCliente.Checked = False
        divFixacoes.Visible = False
        divDevolucaoVNotas.Visible = False
        ddlSubOperacaoDevNN.Enabled = False
        PnlSelecao.Visible = False

        gridGlobalDireta.DataSource = Nothing
        gridGlobalDireta.DataBind()
        gridDeposito.DataSource = Nothing
        gridDeposito.DataBind()
        GridAFixar.DataSource = Nothing
        GridAFixar.DataBind()
        gridFixacao.DataSource = Nothing
        gridFixacao.DataBind()

        gridGlobalDiretaItens.Visible = False
        gridGlobalDiretaItens.DataSource = Nothing
        gridGlobalDiretaItens.DataBind()
        gridDepositoItens.Visible = False
        gridDepositoItens.DataSource = Nothing
        gridDepositoItens.DataBind()
        GridAFixarItens.Visible = False
        GridAFixarItens.DataSource = Nothing
        GridAFixarItens.DataBind()

        txtNota.Text = String.Empty
        txtNotaFix.Text = String.Empty
        txtSerie.Text = String.Empty
        txtSerieFix.Text = String.Empty
        txtPedido.Enabled = True
    End Sub

    Protected Sub chkPeriodo_CheckedChanged(ByVal sender As Object, ByVal e As EventArgs) Handles chkPeriodo.CheckedChanged
        pnlPeriodo.Visible = chkPeriodo.Checked
    End Sub

    Public Sub Consultar(ByVal Fixacoes As Boolean)
        PnlSelecao.Visible = False
        gridGlobalDireta.DataSource = Nothing
        gridGlobalDireta.DataBind()
        gridDeposito.DataSource = Nothing
        gridDeposito.DataBind()
        GridAFixar.DataSource = Nothing
        GridAFixar.DataBind()

        gridGlobalDiretaItens.DataSource = Nothing
        gridGlobalDiretaItens.DataBind()
        gridDepositoItens.DataSource = Nothing
        gridDepositoItens.DataBind()
        GridAFixarItens.DataSource = Nothing
        GridAFixarItens.DataBind()

        ddlOperacao.Items.Clear()
        ddlFilial.Items.Clear()

        If Fixacoes Then
            Dim dsAFixar As New DataSet
            dsAFixar = BuscaFixacoes()

            If dsAFixar.Tables(0).Rows.Count > 0 Then
                divFixacoes.Visible = True
                gridFixacao.DataSource = dsAFixar
                gridFixacao.DataBind()

                For i As Integer = 0 To gridFixacao.Rows.Count - 1
                    If CDec(gridFixacao.Rows(i).Cells(10).Text) < CDec(gridFixacao.Rows(i).Cells(11).Text) Then
                        gridFixacao.Rows(i).Cells(12).ForeColor = Drawing.Color.Red
                    End If
                Next
            Else
                MsgBox(Me.Page, "Nenhum Registro Encontrado")
            End If
        Else
            divFixacoes.Visible = False

            BuscaPedidos()
            If gridGlobalDireta.Rows.Count + gridDeposito.Rows.Count + GridAFixar.Rows.Count > 0 Then
                PnlSelecao.Visible = True

                divPedidosGlobalDireta.Visible = gridGlobalDireta.Rows.Count > 0
                divPedidosDeposito.Visible = gridDeposito.Rows.Count > 0
                divPedidosAFixar.Visible = GridAFixar.Rows.Count > 0
            Else
                MsgBox(Me.Page, "Nenhum Registro Encontrado")
            End If
        End If

        'VerificaTabsPedido()
        Popup.CenterDialog(Me.Page, "divPedidoXSaldo")
        If gridGlobalDireta.Rows.Count > 0 Then
            Funcoes.AbrirAccordion(Me.Page, "divPedidosGlobalDireta", True, True)
        ElseIf gridDeposito.Rows.Count > 0 Then
            Funcoes.AbrirAccordion(Me.Page, "divPedidosDeposito", True, True)
        ElseIf GridAFixar.Rows.Count > 0 Then
            Funcoes.AbrirAccordion(Me.Page, "divPedidosAFixar", True, True)
        End If
    End Sub

    Protected Sub lnkConsultarPedidos_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkConsultarPedidos.Click
        Try
            Consultar(False)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkFechar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkFechar.Click
        Popup.CloseDialog(Me.Page, "divPedidoXSaldo")
    End Sub

    Protected Sub lnkConsultarFixacoes_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkConsultarFixacoes.Click
        Try
            Consultar(True)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkDevolucaoVNotas_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkDevolucaoVNotas.Click
        Try
            divDevolucaoVNotas.Visible = True

            ddlOperacaoDevNN.Items.Clear()
            ddlSubOperacaoDevNN.Items.Clear()

            ddl.Carregar(ddlOperacaoDevNN, CarregarDDL.Tabela.Operacao, "", True)
            ddl.Carregar(ddlGrupoProdutoDevNN, CarregarDDL.Tabela.GrupoProduto)


        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub ddlFilial_TextChanged(ByVal sender As Object, ByVal e As EventArgs) Handles ddlFilial.TextChanged
        SessaoRecuperaSaldoPedidoSelecionado()
        SessaoRecuperaNotaFiscal()

        Dim fil As String()
        fil = ddlFilial.SelectedValue.Split("-")
        objNotaFiscal.CodigoCliente = fil(0)
        objNotaFiscal.CodigoCliente = fil(1)

        Dim sb As String() = ddlOperacao.SelectedValue.Split("-")
        Dim subope As New SubOperacao(sb(0), sb(1))

        objSaldoPedidoSelecionado.RecarregarItens(subope.Classe.ToString(), subope.Devolucao, Not subope.QuantidadeFiscal, 0, "", IIf(subope.Devolucao, fil(0), ""), IIf(subope.Devolucao, fil(1), 0))

        SessaoSalvaSaldoPedidoSelecionado()
        SessaoRecuperaNotaFiscal()

        If gridGlobalDiretaItens.Visible Then
            gridGlobalDiretaItens.DataSource = objSaldoPedidoSelecionado.Itens
            gridGlobalDiretaItens.DataBind()
        End If

        If gridDepositoItens.Visible Then
            gridDepositoItens.DataSource = objSaldoPedidoSelecionado.Itens
            gridDepositoItens.DataBind()
        End If

        If GridAFixarItens.Visible Then
            GridAFixarItens.DataSource = objSaldoPedidoSelecionado.Itens
            GridAFixarItens.DataBind()
        End If
    End Sub

    Protected Overrides Sub Selecionar(ByVal obj As IBaseEntity)
        Try
            Session(Session("ssTipoRetorno")) = obj
            If Session("ssTipoRetorno") IsNot Nothing Then
                If MainUserControl IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(MainUserControl.ClientID) Then
                    Dim uc As UserControl = CType(WebHelpers.FindControlRecursive(Me.Page, MainUserControl.ClientID.Split("_")(1)), UserControl)
                    CType(uc, IBaseUserControl).Carregar(obj)
                Else
                    CType(Me.Page, IBasePage).Carregar(obj)
                End If
                Popup.CloseDialog(Me.Page, "divPedidoXSaldo")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub gridFixacao_SelectedIndexChanged(sender As Object, e As EventArgs) Handles gridFixacao.SelectedIndexChanged
        ddlOperacaoFixacao.Items.Clear()
        Dim op() As String = gridFixacao.Rows(gridFixacao.SelectedIndex).Cells(7).Text.Split("-")

        If CDec(gridFixacao.Rows(gridFixacao.SelectedIndex).Cells(10).Text) < CDec(gridFixacao.Rows(gridFixacao.SelectedIndex).Cells(11).Text) Then
            ddl.Carregar(ddlOperacaoFixacao, CarregarDDL.Tabela.OperacaoSubOperacao, "So.Classe = '" & eClassesOperacoes.AFIXAR.ToString & "' AND So.QuantidadeFiscal = 'N' AND So.DEVOLUCAO = 'S' AND So.Operacao_id = " & op(0), True)
        Else
            ddl.Carregar(ddlOperacaoFixacao, CarregarDDL.Tabela.OperacaoSubOperacao, "So.Operacao_id = " & op(0) & " AND So.SubOperacoes_Id = " & op(1), True)
        End If

        Dim intIndice As Integer = ddlOperacaoFixacao.Items.IndexOf(ddlOperacaoFixacao.Items.FindByValue(Trim(op(0)) & "-" & Trim(op(1))))
        If Not intIndice = -1 Then
            ddlOperacaoFixacao.SelectedIndex = intIndice
        End If

        If ddlOperacaoFixacao.SelectedIndex > 0 Then
            Dim sb As String() = ddlOperacaoFixacao.SelectedValue.Split("-")
            Dim subope As New SubOperacao(sb(0), sb(1))
            If subope.EntradaSaida = eEntradaSaida.Saida Then
                chkNossaEmissaoFix.Checked = True
            Else
                chkNossaEmissaoFix.Checked = False
            End If
            chkNossaEmissaoFix_CheckedChanged(chkNossaEmissaoFix, Nothing)
        End If
        btnConfirmar.Visible = True
    End Sub

    Protected Sub ddlOperacaoFixacao_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddlOperacaoFixacao.SelectedIndexChanged
        If ddlOperacaoFixacao.SelectedIndex > 0 Then
            Dim sb As String() = ddlOperacaoFixacao.SelectedValue.Split("-")
            Dim subope As New SubOperacao(sb(0), sb(1))
            If subope.EntradaSaida = eEntradaSaida.Saida Then
                chkNossaEmissaoFix.Checked = True
            Else
                chkNossaEmissaoFix.Checked = False
            End If
            chkNossaEmissaoFix_CheckedChanged(chkNossaEmissaoFix, Nothing)
        End If
    End Sub

    Protected Sub ddlOperacaoDevNN_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddlOperacaoDevNN.SelectedIndexChanged
        If ddlOperacaoDevNN.SelectedIndex > 0 Then
            ddlSubOperacaoDevNN.Enabled = True

            ddlSubOperacaoDevNN.Items.Clear()

            ddl.Carregar(ddlSubOperacaoDevNN, CarregarDDL.Tabela.OperacaoSubOperacao, "So.Operacao_id = " & ddlOperacaoDevNN.SelectedValue & " and So.Devolucao = 'S'", True)
        End If
    End Sub

    Protected Sub ddlSubOperacaoDevNN_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddlSubOperacaoDevNN.SelectedIndexChanged
        If ddlSubOperacaoDevNN.SelectedIndex > 0 Then
            Dim sb As String() = ddlSubOperacaoDevNN.SelectedValue.Split("-")
            Dim subope As New SubOperacao(sb(0), sb(1))

            If subope.EntradaSaida = eEntradaSaida.Saida Then
                chkNossaEmissaoDevNN.Checked = True
            Else
                chkNossaEmissaoDevNN.Checked = False
            End If

            chkNossaEmissaoDevNN_CheckedChanged(chkNossaEmissaoDevNN, Nothing)
        End If
    End Sub

    Protected Sub ddlGrupoProdutoDevNN_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddlGrupoProdutoDevNN.SelectedIndexChanged
        Try
            If String.IsNullOrEmpty(ddlSubOperacaoDevNN.SelectedValue) Then
                ddlGrupoProdutoDevNN.SelectedIndex = 0
                MsgBox(Me.Page, "Operação não foi selecionada.")
                Exit Sub
            End If

            ddl.Carregar(ddlProdutosDevNN, CarregarDDL.Tabela.Produto, " Situacao = 1 and Grupo ='" & ddlGrupoProdutoDevNN.SelectedValue & "'")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkBuscaProdutoDevNN_Click(sender As Object, e As EventArgs) Handles lnkBuscaProdutoDevNN.Click
        Try
            Dim ucConsultaProduto As ucConsultaProduto = CType(Me.Page.FindControlRecursive("ucConsultaProduto"), ucConsultaProduto)
            If ucConsultaProduto IsNot Nothing Then
                ucConsultaProduto.Limpar()
                ucConsultaProduto.MainUserControl = Me
                Dim txtNome As TextBox = CType(ucConsultaProduto.FindControlRecursive("txtNome"), TextBox)
                Popup.ConsultaDeProduto(Me.Page, "objProdutoDevNN" & HID.Value, txtNome.ClientID, True)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Public Overrides Sub Carregar(obj As IBaseEntity)
        If Session("objProdutoDevNN" & HID.Value) IsNot Nothing Then

            Dim objProduto As [Lib].Negocio.Produto = CType(obj, [Lib].Negocio.Produto)
            ddlGrupoProdutoDevNN.SelectedValue = objProduto.CodigoGrupo

            ddlProdutosDevNN.Items.Clear()
            ddl.Carregar(ddlProdutosDevNN, CarregarDDL.Tabela.Produto, " Situacao = 1")
            ddlProdutosDevNN.SelectedValue = objProduto.Codigo.Trim()


            Session.Remove("objProdutoDevNN" & HID.Value)
        End If
    End Sub

    Protected Sub gridGlobalDireta_RowCommand(sender As Object, e As System.Web.UI.WebControls.GridViewCommandEventArgs) Handles gridGlobalDireta.RowCommand
        If e.CommandName = "Select" Then
            CarregarGridProdutos(CInt(gridGlobalDireta.Rows(CInt(e.CommandArgument)).Cells(2).Text), 1)
        End If
    End Sub

    Protected Sub gridDeposito_RowCommand(sender As Object, e As System.Web.UI.WebControls.GridViewCommandEventArgs) Handles gridDeposito.RowCommand
        If e.CommandName = "Select" Then
            CarregarGridProdutos(CInt(gridDeposito.Rows(CInt(e.CommandArgument)).Cells(2).Text), 2)
        End If
    End Sub

    Protected Sub GridAFixar_RowCommand(sender As Object, e As System.Web.UI.WebControls.GridViewCommandEventArgs) Handles GridAFixar.RowCommand
        If e.CommandName = "Select" Then
            CarregarGridProdutos(CInt(GridAFixar.Rows(CInt(e.CommandArgument)).Cells(2).Text), 3)
        End If
    End Sub

    Public Sub CarregarGridProdutos(pCodigoPedido As Integer, Tipo As Integer)
        'Tipo
        '1 - GlobalDireta
        '2 - Deposito
        '3 - Afixar
        SessaoRecuperaSaldoPedidos()
        SessaoRecuperaNotaFiscal()
        objSaldoPedidoSelecionado = objListSaldoPedidos.Where(Function(s) s.CodigoPedido = pCodigoPedido).FirstOrDefault

        Dim Parametros As New Hashtable

        If Not objSaldoPedidoSelecionado.FiscalAberto Then
            Select Case Tipo
                Case 1
                    gridGlobalDiretaItens.DataSource = Nothing
                    gridGlobalDiretaItens.DataBind()
                Case 2
                    gridDepositoItens.DataSource = Nothing
                    gridDepositoItens.DataBind()
                Case 3
                    GridAFixarItens.DataSource = Nothing
                    GridAFixarItens.DataBind()
            End Select
            MsgBox(Me.Page, "Pedido bloqueado não pode ser selecionado.")
            Exit Sub
        End If

        If objNotaFiscal.VerificarLiberaCarregamento(pCodigoPedido) = False Then
            MsgBox(Me.Page, "Pedido aguardando Liberação para Carregamento.")
            Exit Sub
        End If

        Dim objCliente As New Cliente(objNotaFiscal.CodigoCliente, objNotaFiscal.EnderecoCliente)
        If objCliente.CodigoSituacao = 50 And objSaldoPedidoSelecionado.SubOperacao.Classe = eClassesOperacoes.VENDAS Then
            ddlOperacao.Items.Clear()
            Select Case Tipo
                Case 1
                    gridGlobalDiretaItens.DataSource = Nothing
                    gridGlobalDiretaItens.DataBind()
                Case 2
                    gridDepositoItens.DataSource = Nothing
                    gridDepositoItens.DataBind()
                Case 3
                    GridAFixarItens.DataSource = Nothing
                    GridAFixarItens.DataBind()
            End Select
            MsgBox(Me.Page, "Cliente com Documentação pendente, Faturamento Bloqueado.")
            Exit Sub
        End If

        If objSaldoPedidoSelecionado.Pedido.Troca Then
            If objSaldoPedidoSelecionado.Pedido.PedidoTroca.PedidoTroca Is Nothing Then
                MsgBox(Me.Page, "Este pedido está vinculado ao pedido " & objSaldoPedidoSelecionado.Pedido.PedidoTroca.Codigo & " e esse não está marcado como troca.")
                Exit Sub
            End If
        End If

        If objSaldoPedidoSelecionado.SubOperacao.Classe = eClassesOperacoes.GLOBAL OrElse
            objSaldoPedidoSelecionado.SubOperacao.Classe = eClassesOperacoes.REMESSAS Then
            objNotaFiscal.CodigoPedido = objSaldoPedidoSelecionado.CodigoPedido
            objNotaFiscal.CodigoCliente = objNotaFiscal.Pedido.CodigoCliente
            objNotaFiscal.EnderecoCliente = objNotaFiscal.Pedido.EnderecoCliente
            chkConsolidarCliente.Checked = False
            chkConsolidarCliente.Enabled = False

            txtCodCliente.Value = objCliente.Codigo & "-" & objCliente.CodigoEndereco
            txtNomeCliente.Text = Funcoes.AlinharEsquerda(objCliente.Nome, 28, ".") &
                                  " - " & Funcoes.AlinharEsquerda(objCliente.Cidade, 20, ".") & " " & objCliente.CodigoEstado &
                                  " " & Funcoes.AlinharEsquerda(Funcoes.FormatarCpfCnpj(objCliente.Codigo), 18, ".") &
                                  "-" & objCliente.CodigoEndereco.ToString() & "-" & objCliente.Reduzido
        End If

        If objSaldoPedidoSelecionado.SubOperacao.EntradaSaida = eEntradaSaida.Saida Then
            ddlFilial.Items.Clear()
            Funcoes.AdicionarClienteAoDDL(ddlFilial, objSaldoPedidoSelecionado.Cliente)
            ddlFilial.SelectedValue = objSaldoPedidoSelecionado.CodigoCliente & "-" & objSaldoPedidoSelecionado.EndCliente
            ddlFilial.Enabled = False
        Else
            Parametros.Add("Empresa_Id", objSaldoPedidoSelecionado.CodigoEmpresa)
            Parametros.Add("EndEmpresa_Id", objSaldoPedidoSelecionado.EndEmpresa)
            Parametros.Add("Cliente_NF", objNotaFiscal.CodigoCliente)
            Parametros.Add("EndCliente_NF", objNotaFiscal.EnderecoCliente)
            Parametros.Add("Pedido", objSaldoPedidoSelecionado.CodigoPedido)

            ddl.Carregar(ddlFilial, CarregarDDL.Tabela.FiliaisPedido, "", False, Parametros)

            ddlFilial.SelectedValue = objNotaFiscal.CodigoCliente & "-" & objNotaFiscal.EnderecoCliente
            ddlFilial.Enabled = False
        End If

        'Importação do Xml para notas de entrada
        SessaoSalvaSaldoPedidoSelecionado()

        'Caso não carregue nenhuma operação com os critério do CFOP ele irá carregar novamente as operações sem filtro de CFOP
        If SessaoDsXML IsNot Nothing Then
            chkTodasOperacoes.Visible = True
            If Not CarregarOperacoes(False) Then
                chkTodasOperacoes.Checked = True
                chkTodasOperacoes_CheckedChanged(chkTodasOperacoes, New EventArgs())
                chkTodasOperacoes.Enabled = False
            End If
        Else
            chkTodasOperacoes.Visible = False
            CarregarOperacoes(True)
        End If

        'Deve selecionar a operação do Pedido independente se for carga do XML ou não - Furlan - 19/02/2024
        'If SessaoDsXML IsNot Nothing Then
        '    ddlOperacao.SelectedIndex = 0
        'Else
        ddlOperacao.SelectedValue = CStr(objSaldoPedidoSelecionado.CodigoOperacao) & "-" & CStr(objSaldoPedidoSelecionado.CodigoSubOperacao)
        'End If

        If Tipo = 1 Then
            gridGlobalDiretaItens.Visible = True
            gridGlobalDiretaItens.DataSource = objSaldoPedidoSelecionado.Itens.ToArray
            gridGlobalDiretaItens.DataBind()

        Else
            gridGlobalDiretaItens.Visible = False
            gridGlobalDiretaItens.DataSource = Nothing
            gridGlobalDiretaItens.DataBind()
        End If

        If Tipo = 2 Then
            gridDepositoItens.Visible = True
            gridDepositoItens.DataSource = objSaldoPedidoSelecionado.Itens.ToArray
            gridDepositoItens.DataBind()
        Else
            gridDepositoItens.Visible = False
            gridDepositoItens.DataSource = Nothing
            gridDepositoItens.DataBind()
        End If

        If Tipo = 3 Then
            GridAFixarItens.Visible = True
            GridAFixarItens.DataSource = objSaldoPedidoSelecionado.Itens.ToArray
            GridAFixarItens.DataBind()
        Else
            GridAFixarItens.Visible = False
            GridAFixarItens.DataSource = Nothing
            GridAFixarItens.DataBind()
        End If

        lblPedidoSelecionado.Text = pCodigoPedido
        lblSafraSelecionada.Text = objSaldoPedidoSelecionado.CodigoSafra
        divProdutosDoPedido.Visible = True
        btnConfirmar.Visible = True

        SessaoSalvaNotaFiscal()
        SessaoSalvaSaldoPedidoSelecionado()

        MudancaDeOperacao()
    End Sub

    Protected Sub gridGlobalDireta_RowDataBound(sender As Object, e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles gridGlobalDireta.RowDataBound
        If e.Row.RowType = DataControlRowType.DataRow Then
            If e.Row.Cells(5).Text.ToUpper = "REAL" Then
                CType(e.Row.FindControl("L02"), Literal).Visible = False
                CType(e.Row.FindControl("L04"), Literal).Visible = False

                CType(e.Row.FindControl("L08"), Literal).Parent.Visible = False
                CType(e.Row.FindControl("L08CF"), Literal).Parent.Visible = False
                CType(e.Row.FindControl("L14"), Literal).Parent.Visible = False
                CType(e.Row.FindControl("L14CF"), Literal).Parent.Visible = False
            Else
                CType(e.Row.FindControl("L02"), Literal).Visible = True
                CType(e.Row.FindControl("L04"), Literal).Visible = True

                CType(e.Row.FindControl("L08"), Literal).Parent.Visible = True
                CType(e.Row.FindControl("L08CF"), Literal).Parent.Visible = True
                CType(e.Row.FindControl("L14"), Literal).Parent.Visible = True
                CType(e.Row.FindControl("L14CF"), Literal).Parent.Visible = True
            End If
        End If
    End Sub

    Protected Sub gridGlobalDiretaItens_RowDataBound(sender As Object, e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles gridGlobalDiretaItens.RowDataBound
        'Importação do Xml para notas de entrada
        MostrarOcultarColunaProdutoEntrada(e)

        If e.Row.RowType = DataControlRowType.DataRow Then
            If objSaldoPedidoSelecionado.DescricaoMoeda = "REAL" Then
                CType(e.Row.FindControl("L25"), Literal).Visible = False
                CType(e.Row.FindControl("L26"), Literal).Visible = False

                CType(e.Row.FindControl("L34CF"), Literal).Parent.Visible = False
                CType(e.Row.FindControl("L34"), Literal).Parent.Visible = False
                CType(e.Row.FindControl("L40CF"), Literal).Parent.Visible = False
                CType(e.Row.FindControl("L40"), Literal).Parent.Visible = False
            Else
                CType(e.Row.FindControl("L25"), Literal).Visible = True
                CType(e.Row.FindControl("L26"), Literal).Visible = True

                CType(e.Row.FindControl("L34CF"), Literal).Parent.Visible = True
                CType(e.Row.FindControl("L34"), Literal).Parent.Visible = True
                CType(e.Row.FindControl("L40CF"), Literal).Parent.Visible = True
                CType(e.Row.FindControl("L40"), Literal).Parent.Visible = True
            End If
        End If
    End Sub

    'Importação do Xml para notas de entrada
    Protected Sub gridDepositoItens_RowDataBound(sender As Object, e As GridViewRowEventArgs) Handles gridDepositoItens.RowDataBound
        MostrarOcultarColunaProdutoEntrada(e)
    End Sub

    'Importação do Xml para notas de entrada
    Protected Sub GridAFixarItens_RowDataBound(sender As Object, e As GridViewRowEventArgs) Handles GridAFixarItens.RowDataBound
        MostrarOcultarColunaProdutoEntrada(e)
    End Sub

    'Importação do Xml para notas de entrada
    Private Sub MostrarOcultarColunaProdutoEntrada(ByVal e As GridViewRowEventArgs)
        SessaoRecuperaNotaFiscal()
        If e.Row.RowType = DataControlRowType.Header Then
            If objNotaFiscal.Eletronica AndAlso SessaoDsXML IsNot Nothing Then
                e.Row.Cells(3).Visible = True
            Else
                e.Row.Cells(3).Visible = False
            End If
        ElseIf e.Row.RowType = DataControlRowType.DataRow Then
            Dim ddlXmlProdutoXDePara As DropDownList = CType(e.Row.FindControl("ddlXmlProdutoXDePara"), DropDownList)
            If objNotaFiscal.Eletronica AndAlso SessaoDsXML IsNot Nothing Then
                e.Row.Cells(3).Visible = True
                Dim ProdutoSelecionado As String = objSaldoPedidoSelecionado.Itens(e.Row.RowIndex).XmlProdutoXDePara.CodigoProdutoXML
                CarregarddlProdutosXml(ddlXmlProdutoXDePara, ProdutoSelecionado)
                If Not String.IsNullOrWhiteSpace(ProdutoSelecionado) And Not ddlXmlProdutoXDePara.Items.FindByValue(ProdutoSelecionado) Is Nothing Then
                    CType(e.Row.FindControl("chkSelecionado"), CheckBox).Checked = True
                    'CType(e.Row.FindControl("chkSelecionado"), CheckBox).Enabled = False
                    'ddlXmlProdutoXDePara.Enabled = False
                ElseIf objSaldoPedidoSelecionado.Itens(e.Row.RowIndex).SaldoQtdeDiretoFiscal = 0 Then
                    'ddlXmlProdutoXDePara.Enabled = False
                    'CType(e.Row.FindControl("chkSelecionado"), CheckBox).Enabled = False
                End If
            Else
                e.Row.Cells(3).Visible = False
            End If
            '---
        End If
    End Sub

    'Importação do Xml para notas de entrada
    'Carrega a lista de produtos do xml e caso o vínculo entre os produtos (NGS X Xml) já exista na tabela XmlProdutoXDePara o mesmo será setado
    'senão a lista ficará disponível para que o usuário escolha e ao ser confirmado o pedido os vínculos não existentes serão criados.
    Private Sub CarregarddlProdutosXml(ByVal ddlXmlProdutoXDePara As DropDownList, ByVal ProdutoSelecionado As String)
        SessaoRecuperaSaldoPedidoSelecionado()
        If Not SessaoDsXML Is Nothing Then
            For Each row As DataRow In SessaoDsXML.Tables("Prod").Rows
                ddlXmlProdutoXDePara.Items.Add(New ListItem(row("cProd") & " - " & row("xProd"), row("cProd")))
            Next

            ddlXmlProdutoXDePara.Items.Insert(0, New ListItem("Selecione"))
            If Not String.IsNullOrWhiteSpace(ProdutoSelecionado) Then
                If Not ddlXmlProdutoXDePara.Items.FindByValue(ProdutoSelecionado) Is Nothing Then
                    ddlXmlProdutoXDePara.Items.FindByValue(ProdutoSelecionado).Selected = True
                    SessaoDsXML.Tables("Prod").Select("CProd <> '" & ProdutoSelecionado & "'")
                End If
            End If
        End If
    End Sub

    'Na Classe SaldoPedido2015 foram criados os campos para receber as informações vindas do XML de importação das notas
    'Esses campos serão, caso tenham dados, inseridos em suas determinadas posições no momento da inserção dos ítens da nota.
    Private Function PreencherCamposXmlDoItemSelecionado(ByVal IndiceDoProduto As Integer) As Boolean
        Try
            SessaoRecuperaSaldoPedidoSelecionado()
            Dim t = objSaldoPedidoSelecionado.Itens(IndiceDoProduto)
            objSaldoPedidoSelecionado.Itens(IndiceDoProduto).XmlCFOP = DocumentoEletronico.CarregarDadosDoCampo(SessaoDsXML, "Prod", IndiceDoProduto, "CFOP")
            objSaldoPedidoSelecionado.Itens(IndiceDoProduto).XmluCom = DocumentoEletronico.CarregarDadosDoCampo(SessaoDsXML, "Prod", IndiceDoProduto, "uCom")
            objSaldoPedidoSelecionado.Itens(IndiceDoProduto).XmlqCom = DocumentoEletronico.CarregarDadosDoCampo(SessaoDsXML, "Prod", IndiceDoProduto, "qCom", True)
            objSaldoPedidoSelecionado.Itens(IndiceDoProduto).XmlvUnCom = DocumentoEletronico.CarregarDadosDoCampo(SessaoDsXML, "Prod", IndiceDoProduto, "vUnCom", True)
            objSaldoPedidoSelecionado.Itens(IndiceDoProduto).XmlvProd = DocumentoEletronico.CarregarDadosDoCampo(SessaoDsXML, "Prod", IndiceDoProduto, "vProd", True)
            objSaldoPedidoSelecionado.Itens(IndiceDoProduto).XmluTrib = DocumentoEletronico.CarregarDadosDoCampo(SessaoDsXML, "Prod", IndiceDoProduto, "uTrib")
            objSaldoPedidoSelecionado.Itens(IndiceDoProduto).XmlqTrib = DocumentoEletronico.CarregarDadosDoCampo(SessaoDsXML, "Prod", IndiceDoProduto, "qTrib", True)
            objSaldoPedidoSelecionado.Itens(IndiceDoProduto).XmlvUnTrib = DocumentoEletronico.CarregarDadosDoCampo(SessaoDsXML, "Prod", IndiceDoProduto, "vUnTrib", True)
            objSaldoPedidoSelecionado.Itens(IndiceDoProduto).XmlinfAdProd = DocumentoEletronico.CarregarDadosDoCampo(SessaoDsXML, "Det", IndiceDoProduto, "infAdProd")

            SessaoSalvaSaldoPedidoSelecionado()

            Return True
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
            Return False
        End Try
    End Function

    Private Sub PreencherXmlProdutoXDePara(ByVal IndiceDoProduto As Integer, ByVal ProdutoDe As String)
        Try

            SessaoRecuperaSaldoPedidoSelecionado()

            If objNotaFiscal Is Nothing Then
                SessaoRecuperaNotaFiscal()
            End If

            'objNotaFiscal.Itens(IndiceDoProduto).CodigoProduto = objSaldoPedidoSelecionado.Itens(IndiceDoProduto).CodigoProduto

            For Each itemNF In objNotaFiscal.Itens
                If itemNF.ProdutoXML = ProdutoDe Then
                    itemNF.CodigoProduto = objSaldoPedidoSelecionado.Itens(IndiceDoProduto).CodigoProduto
                End If
            Next

            Dim objItemNota As NotaFiscalXItem
            objItemNota = objNotaFiscal.Itens.Where(Function(x) x.CodigoProduto = objSaldoPedidoSelecionado.Itens(IndiceDoProduto).CodigoProduto).FirstOrDefault()

            Dim ProdutoPara As String = objSaldoPedidoSelecionado.Itens(IndiceDoProduto).CodigoProduto
            'Dim Sqls As New ArrayList
            Dim objProd As New XmlProdutoXDePara(objSaldoPedidoSelecionado.CodigoCliente, objSaldoPedidoSelecionado.EndCliente, ProdutoPara, ProdutoDe)
            objProd.IUD = "I"
            objProd.CodigoCliente = objSaldoPedidoSelecionado.CodigoCliente
            objProd.EndCliente = objSaldoPedidoSelecionado.EndCliente
            objProd.CodigoProdutoXML = ProdutoDe
            objProd.CodigoProduto = ProdutoPara

            If SessaoDsXML IsNot Nothing Then
                objProd.NCMProdutoXML = objItemNota.NCMProdutoXML
                objProd.NomeProdutoXML = objItemNota.NomeProdutoXML
                objProd.UnidadeProdutoXML = objItemNota.UnidadeProdutoXML
            End If

            objProd.Salvar()
            'If Not Banco.GravaBanco(Sqls) Then
            '    MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
            'End If
            objSaldoPedidoSelecionado.Itens(IndiceDoProduto).XmlProdutoXDePara.CodigoProdutoXML = ProdutoDe

            SessaoSalvaSaldoPedidoSelecionado()

            SessaoSalvaNotaFiscal()

        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    'Caso o produto que esteja no xml não existe na tabela XmlProdutoXDePara o mesmo será inserido
    Private Sub PreencherXmlCFOPSaidaXCFOPEntrada(ByVal CFOPSaida As Integer, ByVal CFOPEntrada As Integer)
        Dim ListaDeCFOPS As New ListXmlCFOPSaidaXCFOPEntrada(CFOPSaida)
        If ListaDeCFOPS.Count = 0 Then
            Dim ObjXmlCfop As New XmlCFOPSaidaXCFOPEntrada()
            ObjXmlCfop.IUD = "I"
            ObjXmlCfop.CodigoCFOPSaida = CFOPSaida
            ObjXmlCfop.CodigoCFOPEntrada = CFOPEntrada
            ObjXmlCfop.Salvar()
        End If

    End Sub

    Private Function CarregarOperacoes(ByVal ListarTodas As Boolean) As Boolean
        Try
            Dim Parametros As New Hashtable()
            SessaoRecuperaSaldoPedidoSelecionado()

            Parametros.Clear()
            Parametros.Add("Empresa", objSaldoPedidoSelecionado.CodigoEmpresa)
            Parametros.Add("EndEmpresa", objSaldoPedidoSelecionado.EndEmpresa)
            Parametros.Add("Operacao", objSaldoPedidoSelecionado.CodigoOperacao)
            Parametros.Add("SubOperacao", objSaldoPedidoSelecionado.CodigoSubOperacao)
            Parametros.Add("Pedido", objSaldoPedidoSelecionado.CodigoPedido)

            If SessaoDsXML IsNot Nothing AndAlso Not ListarTodas Then
                Parametros.Add("EstadoOrigem", objSaldoPedidoSelecionado.Empresa.Municipio.CodigoEstado)
                Parametros.Add("EstadoDestino", objSaldoPedidoSelecionado.Cliente.Municipio.CodigoEstado)
                Parametros.Add("RegiaoDestino", objSaldoPedidoSelecionado.Cliente.Estado.Regiao)
                Parametros.Add("GrupoDeProduto", objSaldoPedidoSelecionado.Itens(0).Produto.CodigoGrupo)
                Parametros.Add("Produto", objSaldoPedidoSelecionado.Itens(0).CodigoProduto)
                Parametros.Add("CFOPSaida", DocumentoEletronico.CarregarDadosDoCampo(SessaoDsXML, "Prod", 0, "CFOP"))
            End If

            'Usado para listar apenas suboperações com situação 1 - Normal.
            ddl.Carregar(ddlOperacao, CarregarDDL.Tabela.OperacaoSubOperacaoPermitidasNaNota, "Situacao = 1", False, Parametros)

            If ddlOperacao.Items.Count > 0 Then
                Return True
            Else
                Return False
            End If

        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
            Return False
        End Try
    End Function

    Protected Sub chkTodasOperacoes_CheckedChanged(sender As Object, e As EventArgs) Handles chkTodasOperacoes.CheckedChanged
        'Caso não carregue nenhuma operação com os critério do CFOP ele irá carregar novamente as operações sem filtro de CFOP
        If SessaoDsXML IsNot Nothing Then
            CarregarOperacoes(CType(sender, CheckBox).Checked)
        End If
    End Sub

End Class