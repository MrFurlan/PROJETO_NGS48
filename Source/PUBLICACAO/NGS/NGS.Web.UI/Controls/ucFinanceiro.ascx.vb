
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class ucFinanceiro
    Inherits BaseUserControl

#Region "Inicializar"
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not IsPostBack And IsConnect Then
            Dim Parametros As New Hashtable
            Parametros.Clear()
            Parametros.Add("listarTudo", "N")

            ddl.Carregar(ddlTipoDePagto, CarregarDDL.Tabela.TipoDePagamento, "", True, Parametros)
            ddl.Carregar(ddlFormaPagamentoGeral, CarregarDDL.Tabela.TipoDePagamento, "", True, Parametros)
            ddl.Carregar(ddlCondicaoPagamento, CarregarDDL.Tabela.CondicaoDePagamento, "")
            ddl.Carregar(lstCondicoesPgtoEntrega, CarregarDDL.Tabela.CondicaoDePagamento, "")
            ddl.Carregar(ddlTipoFaturamento, CarregarDDL.Tabela.TipoDeFaturamento, "TipoDeFaturamento_Id in (0,2,3,4,5)", False)
        End If
    End Sub

    Public Property SetarOrigem As String
        Get
            Return HOrigem.Value
        End Get
        Set(ByVal value As String)
            HOrigem.Value = value
        End Set
    End Property

    Public Overrides Sub SetarHID(ByVal guid As String)
        HID.Value = guid
        Limpar()
    End Sub
#End Region

#Region "Retorno UC"
    Public Overrides Sub Carregar(obj As [Lib].Negocio.IBaseEntity)
        If Session("EncargosPlanoDeContas" & HID.Value) IsNot Nothing Then
            Dim LContas As ListEncargosPlanoDeContas = CType(Session("EncargosPlanoDeContas" & HID.Value), ListEncargosPlanoDeContas)
            If LContas.Count = 0 Then Exit Sub
            Dim Msg As String = ""
            SessaoRecuperaTitulos()
            Dim i As Integer = gridVencimentos.SelectedIndex

            For Each row In LContas.Where(Function(s) s.Selecionado AndAlso Not objTitulos(i).Valores.Select(Function(x) x.CodigoContaEncargo).Contains(s.CodigoContaEncargo))
                Dim TxC As New Novo.TituloXContaContabil(objTitulos(i))
                TxC.IUD = "I"
                TxC.CodigoContaEncargo = row.CodigoContaEncargo
                TxC.Descricao = row.TituloEncargo
                TxC.DC = IIf(objTitulos(i).ReceberPagar = "C", "I", (IIf(objTitulos(i).ReceberPagar = "R", row.ContaEncargo.Receber, row.ContaEncargo.Pagar)))

                If String.IsNullOrWhiteSpace(TxC.DC) Then
                    Msg &= IIf(Not String.IsNullOrWhiteSpace(Msg), ", ", "") & row.CodigoContaEncargo & "-" & row.TituloEncargo
                Else
                    objTitulos(i).Valores.Insert(1, TxC)
                    gridValores.DataSource = objTitulos(i).Valores
                    gridValores.DataBind()
                End If
            Next
            objTitulos(i).Valores.AtualizaValores()
            SessaoSalvaTitulos()
            If Not String.IsNullOrWhiteSpace(Msg) Then MsgBox(Me.Page, "Verifique no cadastro de Encargos x Plano de contas o comportamento da conta selecionada no contas a Pagar/Receber: " & Msg)

        ElseIf Session("objBancoUcFinan" & HID.Value) IsNot Nothing Then
            Dim CxC As ClienteXContaBancaria = CType(Session("objBancoUcFinan" & HID.Value), [Lib].Negocio.ClienteXContaBancaria)
            lblBanco.Text = CxC.CodigoBanco & " | " & CxC.NomeBanco
            lblAgencia.Text = CxC.CodigoAgencia & "-" & CxC.DigitoAgencia & " | " & CxC.Praca
            lblContaCorrente.Text = CxC.ContaCorrente & "-" & CxC.DigitoConta

            SessaoRecuperaTitulos()
            If Not String.IsNullOrWhiteSpace(CxC.CodigoBanco) Then
                Dim i As Integer = gridVencimentos.SelectedIndex
                If objTitulos(i).IUD <> "I" Then objTitulos(i).IUD = "U"
                objTitulos(i).CodigoBancoCliFor = CxC.CodigoBanco
                objTitulos(i).CodigoAgenciaCliFor = CxC.CodigoAgencia
                objTitulos(i).DigitoAgenciaCliFor = CxC.DigitoAgencia
                objTitulos(i).ContaCliFor = CxC.ContaCorrente
                objTitulos(i).DigitoContaCliFor = CxC.DigitoConta
            End If
            SessaoSalvaTitulos()
            Session.Remove("objBancoUcFinan" & HID.Value)
        End If
    End Sub
#End Region

#Region "Session NotaFiscal / Pedido / Fixacao"
    Private objResumo As ResumoFinanceiro
    Private objNotaFiscal As [Lib].Negocio.NotaFiscal
    Private objPedido As [Lib].Negocio.Pedido
    Private objFixacao As [Lib].Negocio.Fixacao
    Private objTitulos As [Lib].Negocio.Novo.ListTituloNovo
    Private objTitulosFrete As [Lib].Negocio.ListTitulo
    Private objFrete As [Lib].Negocio.FaturaDeFrete

    Private Sub SessaoSalvaResumoFinanceiro()
        Session("Resumo" & HID.Value) = objResumo
    End Sub
    Private Sub SessaoRecuperaResumoFinanceiro()
        objResumo = CType(Session("Resumo" & HID.Value), ResumoFinanceiro)
        If objResumo Is Nothing Then CarregarResumo()
    End Sub

    Private Sub SessaoSalvaNotaFiscal()
        Session("objNotaFiscal" & HID.Value) = objNotaFiscal
    End Sub
    Private Sub SessaoRecuperaNotaFiscal()
        objNotaFiscal = CType(Session("objNotaFiscal" & HID.Value), [Lib].Negocio.NotaFiscal)
    End Sub

    Private Sub SessaoSalvaPedido()
        Session("objPedido" & HID.Value) = objPedido
    End Sub
    Private Sub SessaoRecuperaPedido()
        objPedido = CType(Session("objPedido" & HID.Value), [Lib].Negocio.Pedido)
    End Sub

    Private Sub SessaoSalvaFixacao()
        Session("objFixacao" & HID.Value) = objFixacao
    End Sub
    Private Sub SessaoRecuperaFixacao()
        objFixacao = CType(Session("objFixacao" & HID.Value), [Lib].Negocio.Fixacao)
    End Sub

    Private Sub SessaoSalvaFrete()
        Session("Fatura" & HID.Value) = objFrete
    End Sub
    Private Sub SessaoRecuperaFrete()
        objFrete = CType(Session("Fatura" & HID.Value), [Lib].Negocio.FaturaDeFrete)
    End Sub

    Private Sub SessaoSalvaTitulos()
        If SetarOrigem = "PD" Then
            SessaoSalvaPedido()
        ElseIf SetarOrigem = "NF" Then
            SessaoSalvaNotaFiscal()
        ElseIf SetarOrigem = "FX" Then
            SessaoSalvaFixacao()
        ElseIf SetarOrigem = "FRETE" Then
            SessaoSalvaFrete()
        Else
            Session("objTitulos" & HID.Value) = objTitulos
        End If
        SessaoSalvaResumoFinanceiro()
    End Sub

    Private Sub SessaoRecuperaTitulos()
        SessaoRecuperaResumoFinanceiro()

        If SetarOrigem = "PD" Then
            SessaoRecuperaPedido()
            objTitulos = objPedido.Titulos
            objTitulos.PD = objPedido
        ElseIf SetarOrigem = "NF" Then
            SessaoRecuperaNotaFiscal()
            objTitulos = objNotaFiscal.Titulos
            objTitulos.NF = objNotaFiscal
        ElseIf SetarOrigem = "FX" Then
            SessaoRecuperaFixacao()
            objTitulos = objFixacao.Titulos
            objTitulos.FX = objFixacao
        ElseIf SetarOrigem = "FRETE" Then
            SessaoRecuperaFrete()
            If FinanceiroNovo Then
                objTitulos = objFrete.ListTituloFatura.Select(Function(s) s.Titulo)
                objTitulos.FRETE = objFrete
            Else
                Dim titulos As New [Lib].Negocio.ListTitulo
                For Each titfat In objFrete.ListTituloFatura
                    titulos.Add(New [Lib].Negocio.Titulo(titfat.CodigoTitulo, titfat.Titulo.ReceberPagar))
                Next
                objTitulosFrete = titulos
                objTitulosFrete.FRETE = objFrete
            End If
        Else
            objTitulos = CType(Session("objTitulos" & HID.Value), [Lib].Negocio.Novo.ListTituloNovo)
        End If
    End Sub
#End Region

#Region "Mudança de Valores"
    Protected Sub ddlTipoDePagto_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        SessaoRecuperaTitulos()
        objTitulos(gridVencimentos.SelectedIndex).CodigoTipoPgto = ddlTipoDePagto.SelectedValue
        If objTitulos(gridVencimentos.SelectedIndex).CodigoTipoPgto = 3 OrElse
           objTitulos(gridVencimentos.SelectedIndex).CodigoTipoPgto = 6 OrElse
           objTitulos(gridVencimentos.SelectedIndex).CodigoTipoPgto = 7 OrElse
           objTitulos(gridVencimentos.SelectedIndex).CodigoTipoPgto = 11 Then
            divLinhaBancariaTitulo.Visible = True
            divLinhaBancariaDados.Visible = True
            BtnConta_Click(BtnConta, Nothing)
        Else
            divLinhaBancariaTitulo.Visible = False
            divLinhaBancariaDados.Visible = False
        End If
        SessaoSalvaTitulos()
    End Sub

    Protected Sub chkDigitado_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkDigitado.CheckedChanged
        SessaoRecuperaTitulos()
        objTitulos(CInt(IndiceGrid.Value)).CodigoDeBarrasDigitado = chkDigitado.Checked
        SessaoSalvaTitulos()
    End Sub

    Protected Sub chkPreImpresso_CheckedChanged(ByVal sender As Object, ByVal e As EventArgs) Handles chkPreImpresso.CheckedChanged
        SessaoRecuperaTitulos()
        objTitulos(CInt(IndiceGrid.Value)).CodigoDeBarrasPreImpresso = chkPreImpresso.Checked
        SessaoSalvaTitulos()
    End Sub

    ' ----- AJUSTE: trabalha com FinanceiroNovo/Legado e faz bind condicional do gridVencimentos
    Protected Sub txtValorOficial_TextChanged(ByVal sender As Object, ByVal e As EventArgs)
        Dim txtValorOficial As TextBox = CType(sender, TextBox)
        Dim row As GridViewRow = CType(txtValorOficial.NamingContainer, GridViewRow)

        Dim i As Integer = CInt(IndiceGrid.Value)
        If Not IsNumeric(txtValorOficial.Text) OrElse CDec(txtValorOficial.Text) < 0 Then txtValorOficial.Text = "0,00"

        SessaoRecuperaTitulos()

        If FinanceiroNovo Then
            If objTitulos(i).CodigoProvisao = 1 Then
                MsgBox(Me.Page, "Não é permitido alterar titulos baixados")
                AtualizaValores()
                Exit Sub
            End If

            If objTitulos(i).Valores(row.RowIndex).CodigoContaEncargo = objTitulos(i).Valores.EncargoValorDocumento.CodigoContaEncargo _
              AndAlso CDec(txtValorOficial.Text) > CDec(ValorDocumento.Value) Then
                MsgBox(Me.Page, "Não é Permitido aumentar o valor do capital de um titulo vinculado a um Pedido com a conta contabil do pedido")
                txtValorOficial.Text = objTitulos(i).Valores.EncargoValorDocumento.ValorOficial
                Exit Sub
            End If

            objTitulos(i).Valores(row.RowIndex).ValorOficial = CDec(txtValorOficial.Text)
            Salvar()

            gridValores.DataSource = objTitulos(i).Valores.ToArray
            gridValores.DataBind()
        Else
            If objTitulosFrete(i).CodigoProvisao = 1 Then
                MsgBox(Me.Page, "Não é permitido alterar titulos baixados")
                Exit Sub
            End If

            If objTitulosFrete(i).Valores(row.RowIndex).CodigoContaEncargo = objTitulosFrete(i).Valores.EncargoValorDocumento.CodigoContaEncargo _
              AndAlso CDec(txtValorOficial.Text) > CDec(ValorDocumento.Value) Then
                MsgBox(Me.Page, "Não é Permitido aumentar o valor do capital de um titulo vinculado a um Pedido com a conta contabil do pedido")
                txtValorOficial.Text = objTitulosFrete(i).Valores.EncargoValorDocumento.ValorOficial
                Exit Sub
            End If

            objTitulosFrete(i).Valores(row.RowIndex).ValorOficial = CDec(txtValorOficial.Text)
            Salvar()

            gridValores.DataSource = objTitulosFrete(i).Valores.ToArray
            gridValores.DataBind()
        End If

        ' BIND condicional do gridVencimentos
        If FinanceiroNovo Then
            gridVencimentos.DataSource = objTitulos
        Else
            gridVencimentos.DataSource = objTitulosFrete
        End If
        gridVencimentos.DataBind()

        SessaoSalvaTitulos()
    End Sub

    ' ----- AJUSTE: idem para ValorMoeda
    Protected Sub txtValorMoeda_TextChanged(sender As Object, e As EventArgs)
        Dim txtValorMoeda As TextBox = CType(sender, TextBox)
        Dim row As GridViewRow = CType(txtValorMoeda.NamingContainer, GridViewRow)

        Dim i As Integer = CInt(IndiceGrid.Value)
        If Not IsNumeric(txtValorMoeda.Text) OrElse CDec(txtValorMoeda.Text) < 0 Then txtValorMoeda.Text = "0,00"

        SessaoRecuperaTitulos()

        If FinanceiroNovo Then
            If objTitulos(i).CodigoProvisao = 1 Then
                MsgBox(Me.Page, "Não é permitido alterar titulos baixados")
                AtualizaValores()
                Exit Sub
            End If

            objTitulos(i).Valores(row.RowIndex).ValorMoeda = CDec(txtValorMoeda.Text)
            Salvar()

            gridValores.DataSource = objTitulos(i).Valores.ToArray
            gridValores.DataBind()
        Else
            If objTitulosFrete(i).CodigoProvisao = 1 Then
                MsgBox(Me.Page, "Não é permitido alterar titulos baixados")
                Exit Sub
            End If

            objTitulosFrete(i).Valores(row.RowIndex).ValorMoeda = CDec(txtValorMoeda.Text)
            Salvar()

            gridValores.DataSource = objTitulosFrete(i).Valores.ToArray
            gridValores.DataBind()
        End If

        If FinanceiroNovo Then
            gridVencimentos.DataSource = objTitulos
        Else
            gridVencimentos.DataSource = objTitulosFrete
        End If
        gridVencimentos.DataBind()

        SessaoSalvaTitulos()
    End Sub

    ' ----- AJUSTE: txtDataVencimento também precisa respeitar a fonte
    Protected Sub txtDataVencimento_TextChanged(ByVal sender As Object, ByVal e As EventArgs) Handles txtDataVencimento.TextChanged
        SessaoRecuperaTitulos()
        Dim idx As Integer = gridVencimentos.SelectedIndex

        objTitulos(idx).Vencimento = CDate(txtDataVencimento.Text)
        objTitulos(idx).Reprogramacao = CDate(txtDataVencimento.Text)
        If FinanceiroNovo Then
            gridVencimentos.DataSource = objTitulos
        Else
            gridVencimentos.DataSource = objTitulosFrete
        End If
        gridVencimentos.DataBind()
        SessaoSalvaTitulos()
    End Sub

    Public Sub AtualizaValores()
        Dim i As Integer = IndiceGrid.Value
        pnlTitulos.Visible = True

        ' No detalhe de título, continuamos usando objTitulos (fluxo novo). Para legado, esse painel pode ser oculto ou ajustado separadamente.
        txtMoeda.Text = objTitulos(i).DescricaoMoeda
        txtCodigoTitulo.Text = objTitulos(i).Codigo
        txtDataVencimento.Text = objTitulos(i).Reprogramacao.ToString("dd/MM/yyyy")
        chkDigitado.Checked = objTitulos(i).CodigoDeBarrasDigitado
        txtCodigoDeBarra.Text = objTitulos(i).CodigoDeBarras

        gridValores.DataSource = objTitulos(i).Valores.ToArray
        gridValores.DataBind()

        ddlTipoDePagto.SelectedValue = IIf(objTitulos(i).CodigoTipoPgto = 0, "", objTitulos(i).CodigoTipoPgto)

        If objTitulos(i).CodigoTipoPgto = 3 OrElse
            objTitulos(i).CodigoTipoPgto = 6 OrElse
            objTitulos(i).CodigoTipoPgto = 7 OrElse
            objTitulos(i).CodigoTipoPgto = 11 Then
            divLinhaBancariaTitulo.Visible = True
            divLinhaBancariaDados.Visible = True
        Else
            divLinhaBancariaTitulo.Visible = False
            divLinhaBancariaDados.Visible = False
        End If

        If objTitulos(i).CodigoBancoCliFor > 0 Then
            lblBanco.Text = objTitulos(i).CodigoBancoCliFor & " | " & objTitulos(i).BancoCliFor.Descricao
            lblAgencia.Text = objTitulos(i).CodigoAgenciaCliFor & "-" & objTitulos(i).DigitoAgenciaCliFor & " | " & objTitulos(i).AgenciaCliFor.Praca
            lblContaCorrente.Text = objTitulos(i).ContaCliFor & "-" & objTitulos(i).DigitoContaCliFor
        Else
            lblBanco.Text = ""
            lblAgencia.Text = ""
            lblContaCorrente.Text = ""
        End If

        pnlTitulos.Enabled = objTitulos(i).CodigoProvisao <> 1
    End Sub

    Protected Sub LimparValores()
        txtMoeda.Text = String.Empty
        txtCodigoTitulo.Text = String.Empty
        txtDataVencimento.Text = String.Empty
        chkDigitado.Checked = False
        txtCodigoDeBarra.Text = String.Empty
        gridValores.DataSource = Nothing
        gridValores.DataBind()
        ddlTipoDePagto.ClearSelection()
        lblBanco.Text = String.Empty
        lblAgencia.Text = String.Empty
        lblContaCorrente.Text = String.Empty
        pnlTitulos.Visible = False
        divLinhaBancariaTitulo.Visible = False
        divLinhaBancariaDados.Visible = False
    End Sub
#End Region

#Region "Methods"
    Public Sub AtualizaGridReajustaFinanceiro(TitulosAlterados As Novo.ListTituloNovo)
        gridBaixasAdiantamentos.DataSource = TitulosAlterados.SelectMany(Function(s) s.Baixas_AdiantamentoEfetuadas)
        gridBaixasAdiantamentos.DataBind()

        ' Condicional
        If FinanceiroNovo Then
            gridVencimentos.DataSource = objTitulos
        Else
            gridVencimentos.DataSource = objTitulosFrete
        End If
        gridVencimentos.DataBind()
    End Sub

    Public Sub ConfigurarFinanceiro(ByVal pOrigem As String)
        Select Case pOrigem
            Case "FX"
                divLinhaAdiantamento.Visible = False
                divLinhaTipoDoTitulo.Visible = False
                divLinhaCondicaoPagamento.Visible = True
                divFormaDePagamento.Visible = True
            Case "PD"
                divLinhaFaturamento.Visible = True
                divLinhaCondicaoPagamento.Visible = True
                divFormaDePagamento.Visible = True
                divBarraBotoes.Visible = True
            Case "FRETE"
                divBarraBotoes.Visible = False
        End Select
    End Sub

    Public Overrides Sub Limpar()
        txtPedido01.Text = "0,00"
        txtValorPedido01.Text = "0,00"
        txtValorTitulosPrevisao01.Text = "0,00"
        txtValorTitulosProvisao01.Text = "0,00"
        txtValorTitulosCompensados01.Text = "0,00"
        txtValorTitulosBaixados01.Text = "0,00"
        txtValorAdiantamentoOriginal01.Text = "0,00"
        txtValorAdiantamento01.Text = "0,00"
        txtAdiantamentoCompensado01.Text = "0,00"
        txtAdiantamentoPago01.Text = "0,00"
        txtSaldoAdiantamento01.Text = "0,00"
        txtSaldoBaixa01.Text = "0,00"
        txtValorPago01.Text = "0,00"

        txtPedido02.Text = "0,00"
        txtValorPedido02.Text = "0,00"
        txtValorTitulosPrevisao02.Text = "0,00"
        txtValorTitulosProvisao02.Text = "0,00"
        txtValorTitulosCompensados02.Text = "0,00"
        txtValorTitulosBaixados02.Text = "0,00"
        txtValorAdiantamentoOriginal02.Text = "0,00"
        txtValorAdiantamento02.Text = "0,00"
        txtAdiantamentoCompensado02.Text = "0,00"
        txtAdiantamentoPago02.Text = "0,00"
        txtSaldoAdiantamento02.Text = "0,00"
        txtSaldoBaixa02.Text = "0,00"
        txtValorPago02.Text = "0,00"

        ddlCondicaoPagamento.SelectedIndex = 0
        txtCodigoTitulo.Text = ""
        txtDataVencimento.Text = ""
        txtCodigoDeBarra.Text = ""
        chkDigitado.Checked = False
        chkPreImpresso.Checked = False
        txtMoeda.Text = ""
        ValorDocumento.Value = 0

        lblBanco.Text = ""
        lblAgencia.Text = ""
        lblContaCorrente.Text = ""
        SessaoRecuperaPedido()
        divEntrega.Visible = SetarOrigem = "PD"

        If SetarOrigem = "PD" Then
            ddlTipoFaturamento.SelectedValue = 2
            divLinhaCondicaoPagamento.Visible = True
        End If
        pnlFormaPagamento.Visible = True
        divLinhaQdteLote.Visible = False
        divLinhaPeridiocidade.Visible = False
        LnkReajuste.Parent.Visible = SetarOrigem = "NF"
        ddlTipoFaturamento.Enabled = True
        ddlCondicaoPagamento.Enabled = True
        ddlPeridiocidade.Enabled = True

        ddlCondicaoPagamento.Enabled = True
        ddlFormaPagamentoGeral.Enabled = True
        ddlFormaPagamentoGeral.ClearSelection()
        pnlTitulos.Visible = False

        If Not gridAdiantamentosDisponiveis.Rows.Count > 0 Then
            gridAdiantamentosDisponiveis.DataSource = Nothing
            gridAdiantamentosDisponiveis.DataBind()
        End If

        gridVencimentos.DataSource = Nothing
        gridVencimentos.DataBind()

        gridValores.DataSource = Nothing
        gridValores.DataBind()
    End Sub

    Public Sub CarregarControle()
        Limpar()
        gridVencimentos.Columns(10).Visible = False
        gridVencimentos.Columns(11).Visible = False
        Select Case SetarOrigem
            Case "NF"
                SessaoRecuperaNotaFiscal()
                objNotaFiscal.Titulos.NF = objNotaFiscal
                objTitulos = objNotaFiscal.Titulos
                If objNotaFiscal.Pedido IsNot Nothing AndAlso objNotaFiscal.Pedido.CodigoCondicaoPagamento <> 0 Then ddlCondicaoPagamento.SelectedValue = objNotaFiscal.Pedido.CodigoCondicaoPagamento
                divLinhaTipoDoTitulo.Visible = False

                If objNotaFiscal.CodigoSubOperacao > 0 AndAlso (Not objNotaFiscal.NFG AndAlso objNotaFiscal.Pedido.Troca) Then
                    divLinhaAdiantamento.Visible = False
                Else
                    divLinhaAdiantamento.Visible = True
                End If

                If objNotaFiscal.NFG OrElse (objNotaFiscal.CodigoPedido > 0 AndAlso objNotaFiscal.Pedido.MomentoFinanceiro = 3) Then
                    divLinhaCondicaoPagamento.Visible = True
                    divFormaDePagamento.Visible = True
                Else
                    divLinhaCondicaoPagamento.Visible = False
                    divFormaDePagamento.Visible = False
                End If

                If objNotaFiscal.IUD = "U" Then
                    gridAdiantamentosDisponiveis.Visible = False
                    gridBaixasAdiantamentos.Visible = True
                    gridBaixasAdiantamentos.DataSource = objTitulos.SelectMany(Function(s) s.Baixas_AdiantamentoEfetuadas)
                    gridBaixasAdiantamentos.DataBind()
                    divLinhaAdiantamento.Visible = gridBaixasAdiantamentos.Rows.Count > 0
                Else
                    gridAdiantamentosDisponiveis.Visible = True
                    gridBaixasAdiantamentos.Visible = False
                End If

                If objTitulos.Where(Function(s) s.CodigoProvisao = 1).Count > 0 Then
                    ddlCondicaoPagamento.Enabled = False
                End If
                If objTitulos.Count > 0 AndAlso objTitulos(0).CodigoTipoPgto > 0 Then
                    ddlFormaPagamentoGeral.SelectedValue = objTitulos(0).CodigoTipoPgto
                End If

            Case "PD"
                SessaoRecuperaPedido()
                objPedido.Titulos.PD = objPedido

                If objPedido.CodigoCondicaoPagamento = 0 Then
                    ddlCondicaoPagamento.SelectedIndex = 0
                Else
                    ddlCondicaoPagamento.SelectedValue = objPedido.CodigoCondicaoPagamento
                End If

                divLinhaAdiantamento.Visible = False

                objTitulos = objPedido.Titulos
                If objPedido.IUD = "I" Then
                    ddlCondicaoPagamento.Enabled = True
                    ddlTipoFaturamento.Enabled = True
                Else
                    ddlCondicaoPagamento.Enabled = False
                    ddlTipoFaturamento.Enabled = False
                End If

                ddlTipoFaturamento.SelectedValue = objPedido.MomentoFinanceiro
                If objPedido.MomentoFinanceiro = 2 Or objPedido.MomentoFinanceiro = 3 Then
                    divLinhaCondicaoPagamento.Visible = True
                ElseIf objPedido.MomentoFinanceiro = 0 Then
                    divLinhaCondicaoPagamento.Visible = False
                End If

                If objPedido.MomentoFinanceiro = eTipoFaturamento.Lote Then
                    gridVencimentos.Columns(10).Visible = True
                    gridVencimentos.Columns(11).Visible = True
                    For Each Tit In objTitulos
                        If Tit.CodigoTituloOrigem = 0 Then
                            Tit.CodigoTituloOrigem = Tit.Codigo
                        End If
                    Next
                ElseIf objPedido.MomentoFinanceiro = eTipoFaturamento.Peridiocidade Then
                    divLinhaPeridiocidade.Visible = True
                    ddlPeridiocidade.SelectedValue = objPedido.PeriodicidadeEntrega
                    ddlPeridiocidade.Enabled = False
                End If

                If objPedido.IUD = "I" Then Exit Sub

            Case "FX"
                SessaoRecuperaFixacao()
                SessaoRecuperaPedido()
                objFixacao.Titulos.FX = objFixacao
                objTitulos = objFixacao.Titulos

            Case "FRETE"
                SessaoRecuperaFrete()

                If FinanceiroNovo Then
                    objTitulos = objFrete.ListTituloFatura.Select(Function(s) s.TituloNovo)
                Else
                    Dim titulos As New [Lib].Negocio.ListTitulo
                    For Each titfat In objFrete.ListTituloFatura
                        titulos.Add(New [Lib].Negocio.Titulo(titfat.CodigoTitulo, titfat.Titulo.ReceberPagar))
                    Next
                    objTitulosFrete = titulos
                    objTitulosFrete.FRETE = objFrete
                End If

                If FinanceiroNovo Then
                    SessaoSalvaTitulos()
                    gridAdiantamentosDisponiveis.DataSource = objTitulos.AdiantamentosAbertos
                    gridAdiantamentosDisponiveis.DataBind()
                    divLinhaAdiantamento.Visible = gridAdiantamentosDisponiveis.Rows.Count > 0
                End If
        End Select

        CarregarResumo()

        ' ----- BIND condicional do gridVencimentos ao final do CarregarControle
        If FinanceiroNovo Then
            gridVencimentos.DataSource = objTitulos
        Else
            gridVencimentos.DataSource = objTitulosFrete
        End If
        gridVencimentos.DataBind()

        If FinanceiroNovo Then
            If objTitulos.Where(Function(s) s.CodigoProvisao = 1).Count > 0 Then
                divBarraBotoes.Visible = False
            Else
                divBarraBotoes.Visible = True
            End If
            SessaoSalvaTitulos()
        Else
            If objTitulosFrete IsNot Nothing AndAlso objTitulosFrete.Count > 0 AndAlso objTitulosFrete(0).CodigoTipoPgto > 0 Then
                ddlFormaPagamentoGeral.SelectedValue = objTitulosFrete(0).CodigoTipoPgto
            End If
            If objTitulosFrete IsNot Nothing AndAlso objTitulosFrete.Where(Function(s) s.CodigoProvisao = 1).Count > 0 Then
                divBarraBotoes.Visible = False
            Else
                divBarraBotoes.Visible = True
            End If
        End If
    End Sub

    Public Sub CarregarResumo()
        Select Case SetarOrigem
            Case "NF"
                SessaoRecuperaNotaFiscal()
                If objNotaFiscal Is Nothing OrElse objNotaFiscal.CodigoPedido = 0 Then Exit Sub
                objResumo = New ResumoFinanceiro(objNotaFiscal.CodigoEmpresa, objNotaFiscal.EnderecoEmpresa, objNotaFiscal.CodigoPedido)
            Case "PD"
                If objPedido Is Nothing Then SessaoRecuperaPedido()
                objResumo = New ResumoFinanceiro(objPedido.CodigoEmpresa, objPedido.EnderecoEmpresa, objPedido.Codigo)
                objPedido.ResumoFinanceiro = objResumo
            Case "FX"
                SessaoRecuperaFixacao()
                objResumo = New ResumoFinanceiro(objFixacao.ItemPedido.Pedido.CodigoEmpresa, objFixacao.ItemPedido.Pedido.EnderecoEmpresa, objFixacao.ItemPedido.Pedido.Codigo)
                SessaoSalvaFixacao()
        End Select
        SessaoSalvaResumoFinanceiro()
        AtualizaFormComAClasseResumoFinanceiro()
    End Sub

    Public Function AtualizaValorPedido() As Decimal
        SessaoRecuperaTitulos()
        If objPedido.Itens.Count = 0 Then Exit Function

        Dim vlr = (From p In objPedido.Itens.SelectMany(Function(s) s.Encargos)
                   Group By p.CodigoEncargo
                    Into ValorMoeda = Sum(p.ValorMoeda), ValorOficial = Sum(p.ValorOficial)
                   Where CodigoEncargo = "LIQUIDO"
                   Select ValorMoeda, ValorOficial).FirstOrDefault

        Dim ValorCE As Decimal = (objResumo.ValorPedido - IIf(objPedido.Moeda.Classificacao = eTiposMoeda.Oficial, vlr.ValorOficial, vlr.ValorMoeda)) * (-1)

        objResumo.ValorPedido = IIf(objPedido.Moeda.Classificacao = eTiposMoeda.Oficial, vlr.ValorOficial, vlr.ValorMoeda)

        If SetarOrigem = "PD" AndAlso Not objPedido.Troca Then
            ddlCondicaoPagamento.Enabled = objPedido.IUD = "I"

            If objPedido.IUD <> "I" Then
                objTitulos.AjustaParcelasComplementoEstornoPedido(Math.Abs(ValorCE), IIf(ValorCE > 0, eTiposLancamentosPedidos.Complemento, eTiposLancamentosPedidos.Estorno))

                If Not objPedido.FinanceiroNovo AndAlso objPedido.MomentoFinanceiro = eTipoFaturamento.Pedido Then
                    objResumo.ValorTitulosEmProvisao += ValorCE
                Else
                    objResumo.ValorTitulosEmPrevisao += ValorCE
                End If

                For Each lancamento In objPedido.Itens.SelectMany(Function(s) s.Lancamentos).Where(Function(x) x.CalculadoPedido = False)
                    lancamento.CalculadoPedido = True
                Next

                ' BIND condicional
                If FinanceiroNovo Then
                    gridVencimentos.DataSource = objTitulos
                Else
                    gridVencimentos.DataSource = objTitulosFrete
                End If
                gridVencimentos.DataBind()
                SessaoSalvaTitulos()
            End If
        End If

        SessaoSalvaResumoFinanceiro()
        AtualizaFormComAClasseResumoFinanceiro()
    End Function

    Public Function AtualizarValorNotaOuFixacaoOuTroca() As Decimal
        If objNotaFiscal.TotalNota = 0 Then Exit Function
        SessaoRecuperaTitulos()

        If SetarOrigem = "NF" AndAlso (objNotaFiscal.SubOperacao.Classe.Equals(eClassesOperacoes.COMPLEMENTACOES) OrElse (objNotaFiscal.SubOperacao.Devolucao AndAlso objNotaFiscal.SubOperacao.Classe.Equals(eClassesOperacoes.AFIXAR))) Then
            objTitulos.RemoveAll(Function(t) t.CodigoFixacao <> objNotaFiscal.Itens(0).CodigoFixacao OrElse t.CodigoProvisao = eProvisao.Baixa)
            objTitulos.AdiantamentosAbertos.RemoveAll(Function(s) s.Titulo.CodigoFixacao <> objNotaFiscal.Itens(0).CodigoFixacao)

            For Each Tit As Novo.TituloNovo In objTitulos.Where(Function(T) T.CodigoProvisao = 2)
                Tit.IUD = "U"
                Tit.CodigoProvisao = 3
                Tit.NotaTitulo = New Novo.NotaFiscalXTitulo(objNotaFiscal, Tit)
                Tit.NotaTitulo.IUD = "I"
            Next

            If objTitulos.AdiantamentosAbertos.Count > 0 Then
                objTitulos.AdiantamentosAbertos.DistribuirValorParaBaixarAdiantamentos(objTitulos.AdiantamentosAbertos.Sum(Function(t) t.Valor), eTiposMoeda.Oficial)
                gridAdiantamentosDisponiveis.DataSource = objTitulos.AdiantamentosAbertos
                gridAdiantamentosDisponiveis.DataBind()
            End If

            divLinhaAdiantamento.Visible = objTitulos.AdiantamentosAbertos.Count > 0
            divBarraBotoes.Visible = False

            ' BIND condicional
            If FinanceiroNovo Then
                gridVencimentos.DataSource = objTitulos
            Else
                gridVencimentos.DataSource = objTitulosFrete
            End If
            gridVencimentos.DataBind()

            SessaoSalvaTitulos()

        ElseIf SetarOrigem = "NF" AndAlso objNotaFiscal.IUD = "I" AndAlso objNotaFiscal.SubOperacao.Devolucao _
                               AndAlso objNotaFiscal.Pedido.FinanceiroNovo Then
            objTitulos.Clear()
            divLinhaAdiantamento.Visible = False

            If objResumo.ValorTitulosEmProvisao > 0 Then
                Dim TitPrev As New Novo.TituloNovo(objNotaFiscal.Pedido.Titulos.Where(Function(s) s.CodigoProvisao = 3).FirstOrDefault.Codigo)
                TitPrev.IUD = "I"
                TitPrev.CodigoProvisao = eProvisao.Compensado
                TitPrev.Quantidade = 0
                TitPrev.ReceberPagar = IIf(TitPrev.ReceberPagar = "R", "P", "R")
                TitPrev.Valores.EncargoValorDocumento.DC = IIf(TitPrev.ReceberPagar = "R", "C", "D")
                TitPrev.Valores.EncargoValorLiquido.DC = IIf(TitPrev.ReceberPagar = "R", "D", "C")

                If objNotaFiscal.Pedido.Moeda.Classificacao = eTiposMoeda.Oficial Then
                    TitPrev.Valores.EncargoValorDocumento.Valor = IIf(objResumo.ValorTitulosEmProvisao > objNotaFiscal.TotalNota, objNotaFiscal.TotalNota, objResumo.ValorTitulosEmProvisao)
                Else
                    TitPrev.Valores.EncargoValorDocumento.Valor = IIf(objResumo.ValorTitulosEmProvisao > Math.Round(objNotaFiscal.TotalNota / objNotaFiscal.IndiceNota, 2), Math.Round(objNotaFiscal.TotalNota / objNotaFiscal.IndiceNota, 2), objResumo.ValorTitulosEmProvisao)
                End If
                TitPrev.NotaTitulo = New Novo.NotaFiscalXTitulo(objNotaFiscal, TitPrev)
                objTitulos.Add(TitPrev)
            End If

            Dim vlrProvisaoPedido As Decimal = 0
            If objNotaFiscal.Pedido.Moeda.Classificacao = eTiposMoeda.Oficial Then
                vlrProvisaoPedido = objResumo.ValorTitulosEmProvisao
            Else
                vlrProvisaoPedido = Math.Round(objResumo.ValorTitulosEmProvisao * objNotaFiscal.IndiceNota, 2)
            End If

            If objNotaFiscal.TotalNota > vlrProvisaoPedido Then
                Dim TitAdiant As New Novo.TituloNovo(objNotaFiscal.Pedido.Titulos.OrderBy(Function(s) s.Reprogramacao).FirstOrDefault.Codigo)
                TitAdiant.IUD = "I"
                TitAdiant.ReceberPagar = IIf(objNotaFiscal.Pedido.SubOperacao.EntradaSaida = eEntradaSaida.Entrada, "P", "R")
                If objNotaFiscal.Pedido.Moeda.Classificacao = eTiposMoeda.Oficial Then
                    TitAdiant.Valores.EncargoValorDocumento.Valor = objNotaFiscal.TotalNota - objResumo.ValorTitulosEmProvisao
                Else
                    TitAdiant.Valores.EncargoValorDocumento.Valor = Math.Round(objNotaFiscal.TotalNota / objNotaFiscal.IndiceNota, 2) - objResumo.ValorTitulosEmProvisao
                End If

                If objNotaFiscal.Pedido.Troca Then
                    TitAdiant.CodigoProvisao = eProvisao.Compensado
                    TitAdiant.CodigoContaContabilCliFor = objNotaFiscal.Pedido.SubOperacao.CodigoGrupoContas
                    TitAdiant.CodigoContaContabilRecPag = objNotaFiscal.Empresa.Empresa.CodigoContaCaixaCompensacao
                    TitAdiant.ReceberPagar = IIf(TitAdiant.ReceberPagar = "R", "P", "R")
                Else
                    TitAdiant.CodigoProvisao = eProvisao.Baixa
                    If String.IsNullOrWhiteSpace(objNotaFiscal.Pedido.SubOperacao.CodigoContaAdiantamento) Then
                        MsgBox(Me.Page, "Para geração de adiantamento a fornecedor é necessário incluir a conta de adiantamento na operação: " & objNotaFiscal.Pedido.CodigoOperacao & "-" & objNotaFiscal.Pedido.CodigoSubOperacao & " " & objNotaFiscal.Pedido.SubOperacao.Descricao)
                        Return False
                    End If
                    TitAdiant.CodigoContaContabilCliFor = objNotaFiscal.Pedido.SubOperacao.CodigoContaAdiantamento
                    TitAdiant.CodigoContaContabilRecPag = TitAdiant.Pedido.SubOperacao.CodigoGrupoContas
                End If

                TitAdiant.NotaTitulo = New Novo.NotaFiscalXTitulo(objNotaFiscal, TitAdiant)
                objTitulos.Add(TitAdiant)
            End If

            If FinanceiroNovo Then
                gridVencimentos.DataSource = objTitulos
            Else
                gridVencimentos.DataSource = objTitulosFrete
            End If
            gridVencimentos.DataBind()

            SessaoSalvaTitulos()

        ElseIf ((SetarOrigem = "NF" AndAlso objNotaFiscal.NFG) OrElse (SetarOrigem = "FX")) And objTitulos.Count > 0 Then
            If objNotaFiscal.Titulos.Where(Function(s) s.CodigoProvisao = eProvisao.Baixa).Count > 0 Then Exit Function
            Dim quantParcela As Integer = objTitulos.Where(Function(s) s.CodigoProvisao <> 1).Count
            Dim contador As Integer = 1
            Dim SomaValor As Decimal = 0
            Dim condicaoPagamento As New CondicaoPagamento(ddlCondicaoPagamento.SelectedValue)

            If objTitulos.Count < condicaoPagamento.Parcelas Then
                Dim listaTitulo As New Novo.ListTituloNovo
                For Each tit As Novo.TituloNovo In objTitulos
                    listaTitulo.Add(tit)
                Next
                objTitulos.Clear()
                objTitulos.Parcelar(ddlCondicaoPagamento.SelectedValue, SetarOrigem, ddlFormaPagamentoGeral.SelectedValue)
                For i As Integer = objTitulos.Count - 1 To objTitulos.Count - listaTitulo.Count Step -1
                    objTitulos.RemoveAt(i)
                Next
                For Each tit As Novo.TituloNovo In listaTitulo
                    objTitulos.Add(tit)
                Next
                quantParcela = objTitulos.Count
            Else
                quantParcela = condicaoPagamento.Parcelas
            End If

            If objNotaFiscal IsNot Nothing Then objNotaFiscal.Pedido.CodigoCondicaoPagamento = ddlCondicaoPagamento.SelectedValue

            For Each tit In objTitulos.Where(Function(s) s.CodigoProvisao <> 1).OrderBy(Function(t) t.Codigo)
                If contador > quantParcela OrElse objNotaFiscal.TotalNota = objTitulos.Where(Function(s) s.CodigoProvisao = eProvisao.Baixa).Sum(Function(s) s.Valores.EncargoValorDocumento.Valor) Then
                    tit.IUD = "D"
                    tit.NotaTitulo.IUD = "D"
                Else
                    If tit.IUD <> "I" Then tit.IUD = "U"
                    tit.Vencimento = Funcoes.ValidaDataUtil(objNotaFiscal.CodigoEmpresa, objNotaFiscal.EnderecoEmpresa, objNotaFiscal.Movimento.AddDays(condicaoPagamento.Periodo(contador - 1)))
                    tit.Reprogramacao = tit.Vencimento
                    If String.IsNullOrWhiteSpace(tit.NotaTitulo.IUD) Then tit.NotaTitulo.IUD = "I"

                    Dim valorTotal As Decimal = 0
                    If objNotaFiscal IsNot Nothing Then
                        valorTotal = objNotaFiscal.TotalNota
                    ElseIf objFixacao IsNot Nothing Then
                        valorTotal = objFixacao.TotalOficial
                        If objTitulos.Where(Function(s) s.CodigoProvisao = 3).Count Then tit.CodigoProvisao = 3
                    End If

                    If contador = quantParcela Then
                        tit.Valores.EncargoValorDocumento.Valor = (valorTotal - SomaValor).ToString("N2")
                    Else
                        tit.Valores.EncargoValorDocumento.Valor = (valorTotal / quantParcela).ToString("N2")
                    End If
                    SomaValor += tit.Valores.EncargoValorDocumento.Valor
                    contador = contador + 1
                End If
            Next

            If FinanceiroNovo Then
                gridVencimentos.DataSource = objTitulos
            Else
                gridVencimentos.DataSource = objTitulosFrete
            End If
            gridVencimentos.DataBind()
            SessaoSalvaTitulos()

        ElseIf SetarOrigem = "FRETE" Then
            SessaoRecuperaFrete()
            Dim vlrBaixa As Decimal = objTitulos.Sum(Function(s) s.Valores.EncargoValorLiquido.Valor)
            gridAdiantamentosDisponiveis.DataSource = objTitulos.AdiantamentosAbertos
            gridAdiantamentosDisponiveis.DataBind()

            If FinanceiroNovo Then
                gridVencimentos.DataSource = objTitulos
            Else
                gridVencimentos.DataSource = objTitulosFrete
            End If
            gridVencimentos.DataBind()
        Else
            objTitulos.Clear()

            If Not objTitulos.AdiantamentosAbertos Is Nothing AndAlso objTitulos.AdiantamentosAbertos.Count > 0 Then
                If objNotaFiscal.Pedido Is Nothing OrElse objNotaFiscal.Pedido.Moeda Is Nothing Then
                    objTitulos.AdiantamentosAbertos.DistribuirValorParaBaixarAdiantamentos(objNotaFiscal.TotalNota, eTiposMoeda.Oficial, True)
                Else
                    objTitulos.AdiantamentosAbertos.DistribuirValorParaBaixarAdiantamentos(objNotaFiscal.TotalNota, objNotaFiscal.Pedido.Moeda.Classificacao, True)
                End If
            End If

            If SetarOrigem = "NF" AndAlso Not objNotaFiscal.NFG Then
                pnlFormaPagamento.Visible = False
                If objNotaFiscal.Pedido.MomentoFinanceiro = eTipoFaturamento.Pedido OrElse
                   objNotaFiscal.Pedido.MomentoFinanceiro = eTipoFaturamento.Lote OrElse
                   objNotaFiscal.Pedido.MomentoFinanceiro = eTipoFaturamento.Peridiocidade Then
                    objTitulos.ParcelarNotaProvisao()
                    Dim vlrProvisao As Decimal = objTitulos.Sum(Function(i) i.Valores.EncargoValorDocumento.ValorOficial) + objNotaFiscal.Titulos.AdiantamentosAbertos.Sum(Function(T) T.VlrBaixa)
                    If Not objTitulos.Sum(Function(i) i.Valores.EncargoValorDocumento.ValorOficial) + objNotaFiscal.Titulos.AdiantamentosAbertos.Sum(Function(T) T.VlrBaixa) = objNotaFiscal.TotalNota Then
                        Dim vlrAjusteEncargo As Decimal = (objTitulos.Sum(Function(i) i.Valores.EncargoValorDocumento.ValorOficial) + objNotaFiscal.Titulos.AdiantamentosAbertos.Sum(Function(T) T.VlrBaixa)) - objNotaFiscal.TotalNota
                        If TypeOf Me.Page Is NotaFiscalXItens Then
                            CType(Me.Page, NotaFiscalXItens).AjustarEncargos(vlrAjusteEncargo)
                        End If
                    End If
                ElseIf objNotaFiscal.Pedido.MomentoFinanceiro = eTipoFaturamento.NotaFiscal Then
                    ddlCondicaoPagamento.SelectedValue = objNotaFiscal.Pedido.CodigoCondicaoPagamento
                    objTitulos.Parcelar(ddlCondicaoPagamento.SelectedValue, SetarOrigem)
                End If

                If FinanceiroNovo Then
                    gridVencimentos.DataSource = objTitulos
                Else
                    gridVencimentos.DataSource = objTitulosFrete
                End If
                gridVencimentos.DataBind()

                gridAdiantamentosDisponiveis.DataSource = objTitulos.AdiantamentosAbertos.Where(Function(s) s.Titulo.CodigoPedido = objNotaFiscal.CodigoPedido)
            Else
                gridAdiantamentosDisponiveis.DataSource = objTitulos.AdiantamentosAbertos.OrderBy(Function(s) s.Titulo.Movimento)
            End If
            gridAdiantamentosDisponiveis.DataBind()

            divLinhaAdiantamento.Visible = gridAdiantamentosDisponiveis.Rows.Count > 0
            SessaoSalvaTitulos()
        End If
    End Function

    Public Sub AtualizaFaturaFrete(ByVal pVlrSaida As Decimal, ByVal pVlrChegada As Decimal, ByVal pData As Date, Optional ByVal pBaixaAdiantamento As Boolean = True)
        SessaoRecuperaTitulos()
        SessaoRecuperaFrete()

        If pBaixaAdiantamento Then
            objTitulos.AdiantamentosAbertos.DistribuirValorParaBaixarAdiantamentos(pVlrChegada, eTiposMoeda.Oficial, True)
        Else
            For Each ad In objTitulos.AdiantamentosAbertos
                ad.VlrBaixa = 0
            Next
        End If

        gridAdiantamentosDisponiveis.DataSource = objTitulos.AdiantamentosAbertos
        gridAdiantamentosDisponiveis.DataBind()

        For Each tit As [Lib].Negocio.Novo.TituloNovo In objTitulos
            tit.IUD = "U"
            tit.CodigoProvisao = eProvisao.Provisao
            tit.Reprogramacao = pData
            tit.Movimento = pData
            tit.Valores.EncargoValorDocumento.ValorOficial = pVlrSaida
            tit.Valores.EncargoValorDocumento.ValorMoeda = Funcoes.ConverteParaMoedaExtrangeira(pVlrChegada, 3, pData)

            Dim CodigoContaJuroPago As String = tit.Empresa.Empresa.CodigoContaJuroPago
            If (pVlrSaida - pVlrChegada) > 0 Then
                tit.Valores.RemoveAll(Function(s) s.CodigoContaEncargo = CodigoContaJuroPago)
                Dim enc As New [Lib].Negocio.Novo.TituloXContaContabil(tit)
                enc.CodigoContaEncargo = tit.Empresa.Empresa.CodigoContaDescontoObtido
                enc.ValorOficial = pVlrSaida - pVlrChegada
                enc.DC = "C"
                enc.ValorMoeda = Funcoes.ConverteParaMoedaExtrangeira(enc.ValorOficial, 3, pData)
                tit.Valores.Add(enc)
            Else
                tit.Valores.RemoveAll(Function(s) s.CodigoContaEncargo = CodigoContaJuroPago)
                Dim enc As New [Lib].Negocio.Novo.TituloXContaContabil(tit)
                enc.CodigoContaEncargo = tit.Empresa.Empresa.CodigoContaJuroPago
                enc.DC = "D"
                enc.ValorOficial = pVlrChegada - pVlrSaida
                enc.ValorMoeda = Funcoes.ConverteParaMoedaExtrangeira(enc.ValorOficial, 3, pData)
                tit.Valores.Add(enc)
            End If

            If objTitulos.AdiantamentosAbertos.Sum(Function(s) s.VlrBaixa) >= pVlrChegada Then
                tit.CodigoSituacao = eSituacao.Excluido
            Else
                tit.Valores.EncargoValorLiquido.ValorOficial = pVlrChegada - objTitulos.AdiantamentosAbertos.Sum(Function(s) s.VlrBaixa)
                tit.Valores.EncargoValorLiquido.ValorMoeda = Funcoes.ConverteParaMoedaExtrangeira(pVlrChegada - objTitulos.AdiantamentosAbertos.Sum(Function(s) s.VlrBaixa), 3, pData)
                tit.CodigoSituacao = eSituacao.Normal
            End If
        Next

        If FinanceiroNovo Then
            gridVencimentos.DataSource = objTitulos
        Else
            gridVencimentos.DataSource = objTitulosFrete
        End If
        gridVencimentos.DataBind()

        For Each tit In objTitulos
            objFrete.ListTituloFatura.Add(New FaturaDeFretexTitulo(objFrete) With {.CodigoTitulo = tit.Codigo, .TituloNovo = tit})
        Next

        SessaoSalvaTitulos()
        SessaoSalvaFrete()
    End Sub

    Public Sub AtualizaFormComAClasseResumoFinanceiro()
        If Not objResumo Is Nothing Then
            txtPedido01.Text = objResumo.CodigoPedido
            txtValorPedido01.Text = objResumo.ValorPedido.ToString("N2")

            txtValorTitulosPrevisao01.Text = objResumo.ValorTitulosEmPrevisao.ToString("N2")
            txtValorTitulosProvisao01.Text = objResumo.ValorTitulosEmProvisao.ToString("N2")
            txtValorTitulosBaixados01.Text = objResumo.ValorTitulosBaixado.ToString("N2")
            txtValorTitulosCompensados01.Text = objResumo.ValorTitulosCompensado.ToString("N2")

            txtValorAdiantamentoOriginal01.Text = objResumo.ValorAdiantamentoOriginal.ToString("N2")
            txtValorAdiantamento01.Text = objResumo.ValorAdiantamento.ToString("N2")
            txtAdiantamentoCompensado01.Text = objResumo.ValorAdiantamentoCompensado.ToString("N2")
            txtAdiantamentoPago01.Text = objResumo.ValorAdiantamentoPagoDireto.ToString("N2")

            txtAdiantamentoAmortizado01.Text = objResumo.ValorAdiantamentoAmortizado.ToString("N2")
            txtSaldoAdiantamento01.Text = objResumo.SaldoAdiantamento.ToString("N2")
            txtSaldoBaixa01.Text = objResumo.SaldoBaixa.ToString("N2")
            txtValorPago01.Text = objResumo.ValorPago.ToString("N2")

            If objResumo.ResumoTroca IsNot Nothing Then
                txtPedido02.Text = objResumo.ResumoTroca.CodigoPedido
                txtValorPedido02.Text = objResumo.ResumoTroca.ValorPedido.ToString("N2")

                txtValorTitulosPrevisao02.Text = objResumo.ResumoTroca.ValorTitulosEmPrevisao.ToString("N2")
                txtValorTitulosProvisao02.Text = objResumo.ResumoTroca.ValorTitulosEmProvisao.ToString("N2")
                txtValorTitulosBaixados02.Text = objResumo.ResumoTroca.ValorTitulosBaixado.ToString("N2")
                txtValorTitulosCompensados02.Text = objResumo.ResumoTroca.ValorTitulosCompensado.ToString("N2")

                txtValorAdiantamentoOriginal02.Text = objResumo.ResumoTroca.ValorAdiantamentoOriginal.ToString("N2")
                txtValorAdiantamento02.Text = objResumo.ResumoTroca.ValorAdiantamento.ToString("N2")
                txtAdiantamentoCompensado02.Text = objResumo.ResumoTroca.ValorAdiantamentoCompensado.ToString("N2")
                txtAdiantamentoPago02.Text = objResumo.ResumoTroca.ValorAdiantamentoPagoDireto.ToString("N2")

                txtAdiantamentoAmortizado02.Text = objResumo.ResumoTroca.ValorAdiantamentoAmortizado.ToString("N2")
                txtSaldoAdiantamento02.Text = objResumo.ResumoTroca.SaldoAdiantamento.ToString("N2")
                txtSaldoBaixa02.Text = objResumo.ResumoTroca.SaldoBaixa.ToString("N2")
                txtValorPago02.Text = objResumo.ResumoTroca.ValorPago.ToString("N2")
            End If
        End If
    End Sub

    Public Sub AtualizaAdiantamentos(ByVal pTitulos As Novo.ListTituloNovo)
        Dim vlrBaixa As Decimal = Decimal.Zero
        For Each row As GridViewRow In gridAdiantamentosDisponiveis.Rows
            vlrBaixa = CDec(CType(row.FindControl("txtVlrBaixa"), TextBox).Text)
            pTitulos.AdiantamentosAbertos(row.RowIndex).VlrBaixa = IIf(vlrBaixa > 0, vlrBaixa, 0)
        Next
    End Sub

    ' ----- AJUSTE: Salvar continua trabalhando com as duas fontes (ajuste legado limitado ao set do ValorDocumento)
    Public Sub Salvar()
        Dim msg As String = ""
        Dim i As Integer = CInt(IndiceGrid.Value)

        If FinanceiroNovo Then
            If objTitulos(i).IUD <> "I" Then objTitulos(i).IUD = "U"

            If objTitulos(i).Moeda.Classificacao = eTiposMoeda.Oficial Then
                If objTitulos(i).Valores.EncargoValorDocumento.ValorOficial <> CDec(ValorDocumento.Value) Then
                    msg = objTitulos.AJustaParcelasDaLista(i, ValorDocumento.Value, objTitulos(i).Valores.EncargoValorDocumento.ValorOficial)
                    If msg.Length > 0 Then
                        MsgBox(Me.Page, msg)
                    Else
                        ValorDocumento.Value = objTitulos(i).Valores.EncargoValorDocumento.ValorOficial
                    End If
                End If
            Else
                If objTitulos(i).Valores.EncargoValorDocumento.ValorMoeda <> CDec(ValorDocumento.Value) Then
                    msg = objTitulos.AJustaParcelasDaLista(i, ValorDocumento.Value, objTitulos(i).Valores.EncargoValorDocumento.ValorMoeda)
                    If msg.Length > 0 Then
                        MsgBox(Me.Page, msg)
                    Else
                        ValorDocumento.Value = objTitulos(i).Valores.EncargoValorDocumento.ValorMoeda
                    End If
                End If
            End If
        Else
            If objTitulosFrete(i).IUD <> "I" Then objTitulosFrete(i).IUD = "U"
            If objTitulosFrete(i).Moeda IsNot Nothing AndAlso objTitulosFrete(i).Moeda.Classificacao = eTiposMoeda.Oficial Then
                ValorDocumento.Value = objTitulosFrete(i).Valores.EncargoValorDocumento.ValorOficial
            Else
                ValorDocumento.Value = objTitulosFrete(i).Valores.EncargoValorDocumento.ValorMoeda
            End If
        End If
    End Sub

    Public Sub ParcelarTroca()
        SessaoRecuperaTitulos()
        SessaoRecuperaPedido()
        divLinhaPeridiocidade.Visible = False
        pnlFormaPagamento.Visible = False

        ddlFormaPagamentoGeral.SelectedValue = 1
        objPedido.MomentoFinanceiro = eTipoFaturamento.Pedido
        objPedido.PeriodicidadeEntrega = eTipoFaturamento.Pedido
        objPedido.CodigoCondicaoPagamento = 1
        objTitulos.Parcelar(1, SetarOrigem, ddlFormaPagamentoGeral.SelectedValue)

        objTitulos(0).Vencimento = objPedido.DataVencimentoPedido
        objTitulos(0).Reprogramacao = objTitulos(0).Vencimento
        objTitulos(0).DataMoeda = objTitulos(0).Vencimento
        objTitulos(0).DataBaixa = objTitulos(0).Vencimento

        If FinanceiroNovo Then
            gridVencimentos.DataSource = objTitulos
        Else
            gridVencimentos.DataSource = objTitulosFrete
        End If
        gridVencimentos.DataBind()

        SessaoSalvaTitulos()
        SessaoSalvaPedido()
    End Sub
#End Region

    Protected Sub LnkParcelamento_Click(sender As Object, e As EventArgs) Handles LnkParcelamento.Click
        gridVencimentos.Columns(10).Visible = False
        gridVencimentos.Columns(11).Visible = False
        If Not SetarOrigem = "FRETE" Then
            If divLinhaFaturamento.Visible AndAlso ddlTipoFaturamento.SelectedIndex = 0 Then
                MsgBox(Me.Page, "Faturamento não foi selecionado!")
                Exit Sub
            ElseIf ddlCondicaoPagamento.SelectedIndex = 0 Then
                MsgBox(Me.Page, "Condições de Pagamento não foi selecionado!")
                Exit Sub
            ElseIf ddlFormaPagamentoGeral.SelectedIndex = 0 Then
                MsgBox(Me.Page, "Forma de Pagamento não foi selecionada!")
                Exit Sub
            ElseIf ddlTipoFaturamento.SelectedValue = eTipoFaturamento.Lote Then
                If String.IsNullOrWhiteSpace(txtQtdeLote.Text) OrElse Not CDec(txtQtdeLote.Text) > 0 Then
                    MsgBox(Me.Page, "Campo Quantidade é de preenchimento obrigatório!")
                    Exit Sub
                End If
            End If
        End If

        IndiceGrid.Value = -1
        ValorDocumento.Value = 0
        Dim IUD As String = String.Empty

        If Not String.IsNullOrWhiteSpace(SetarOrigem) Then
            SessaoRecuperaTitulos()
            SessaoRecuperaPedido()

            Dim vlrBaixa As Decimal = Decimal.Zero
            Dim vlrAdiantamento As Decimal = Decimal.Zero
            For Each row As GridViewRow In gridAdiantamentosDisponiveis.Rows
                vlrBaixa += CDec(CType(row.FindControl("txtVlrBaixa"), TextBox).Text)
                vlrAdiantamento = CDec(CType(row.FindControl("txtVlrBaixa"), TextBox).Text)
                objTitulos.AdiantamentosAbertos(row.RowIndex).VlrBaixa = IIf(vlrAdiantamento > 0, vlrAdiantamento, 0)
            Next

            If SetarOrigem = "PD" Then
                objPedido.CodigoCondicaoPagamento = ddlCondicaoPagamento.SelectedValue
                IUD = objPedido.IUD
                If objPedido.MomentoFinanceiro = eTipoFaturamento.NotaFiscal OrElse objPedido.MomentoFinanceiro = eTipoFaturamento.Peridiocidade Then
                    objPedido.CodigoCondicaoPagamento = ddlCondicaoPagamento.SelectedValue
                    If objPedido.MomentoFinanceiro = eTipoFaturamento.Peridiocidade Then objPedido.PeriodicidadeEntrega = ddlPeridiocidade.SelectedValue
                    objTitulos.Parcelar(1, SetarOrigem, ddlFormaPagamentoGeral.SelectedValue)

                    objTitulos(0).Vencimento = objPedido.DataVencimentoPedido
                    objTitulos(0).Reprogramacao = objTitulos(0).Vencimento
                    objTitulos(0).DataMoeda = objTitulos(0).Vencimento
                    objTitulos(0).DataBaixa = objTitulos(0).Vencimento

                    If FinanceiroNovo Then
                        gridVencimentos.DataSource = objTitulos
                    Else
                        gridVencimentos.DataSource = objTitulosFrete
                    End If
                    gridVencimentos.DataBind()

                    SessaoSalvaTitulos()
                    SessaoSalvaPedido()
                    Exit Sub
                ElseIf objPedido.MomentoFinanceiro = eTipoFaturamento.Lote Then
                    Dim quantidade As Decimal = Convert.ToDecimal(txtQtdeLote.Text.Replace(".", ""))
                    Dim valor = (objPedido.Itens.LiquidoOficial * quantidade / objPedido.Itens.QuantidadeTotal).ToString("N2")

                    Dim total As Decimal = objTitulos.Sum(Function(s) s.Quantidade) + quantidade

                    If total = objPedido.Itens.QuantidadeTotal Then
                        valor = objPedido.Itens.LiquidoOficial - objTitulos.Sum(Function(s) s.Valores.EncargoValorDocumento.Valor)
                    End If

                    If valor > objPedido.Itens.LiquidoOficial - objTitulos.Sum(Function(s) s.Valores.EncargoValorDocumento.Valor) Then
                        MsgBox(Me.Page, "Não há saldo disponível no pedido para a quantidade informada no lote!")
                        Exit Sub
                    End If

                    If total > objPedido.Itens.QuantidadeTotal Then
                        MsgBox(Me.Page, "Não há saldo disponível no pedido para a quantidade informada no lote!")
                        Exit Sub
                    End If

                    objTitulos.Parcelar(1, SetarOrigem, ddlFormaPagamentoGeral.SelectedValue, valor, txtQtdeLote.Text)
                    gridVencimentos.Columns(10).Visible = True
                    gridVencimentos.Columns(11).Visible = True

                    If FinanceiroNovo Then
                        gridVencimentos.DataSource = objTitulos
                    Else
                        gridVencimentos.DataSource = objTitulosFrete
                    End If
                    gridVencimentos.DataBind()

                    ddlTipoFaturamento.Enabled = False
                    ddlCondicaoPagamento.Enabled = False
                    SessaoSalvaTitulos()
                    SessaoSalvaPedido()
                    txtQtdeLote.Text = String.Empty
                    Exit Sub
                End If

            ElseIf SetarOrigem = "NF" AndAlso objNotaFiscal.NFG Then
                IUD = objNotaFiscal.IUD
                If objNotaFiscal.Pedido Is Nothing Then
                    objNotaFiscal.Pedido = New Pedido()
                    objNotaFiscal.Pedido.CodigoCondicaoPagamento = ddlCondicaoPagamento.SelectedValue
                End If

            ElseIf SetarOrigem = "FX" Then
                IUD = objFixacao.IUD

            ElseIf SetarOrigem = "FRETE" Then
                SessaoRecuperaFrete()
                SessaoRecuperaTitulos()

                For Each tit As [Lib].Negocio.Novo.TituloNovo In objTitulos
                    If vlrBaixa >= tit.Valores.EncargoValorLiquido.Valor Then
                        tit.CodigoSituacao = eSituacao.Excluido
                    Else
                        tit.Valores.EncargoValorDocumento.ValorOficial = tit.Valores.EncargoValorDocumento.Valor - vlrBaixa
                        tit.Valores.EncargoValorDocumento.ValorMoeda = Funcoes.ConverteParaMoedaExtrangeira(tit.Valores.EncargoValorDocumento.Valor - objTitulos.AdiantamentosAbertos.Sum(Function(s) s.VlrBaixa), 3, tit.Movimento)
                        tit.CodigoSituacao = eSituacao.Normal
                        tit.CodigoProvisao = eProvisao.Provisao
                    End If
                Next

                objTitulos.AdiantamentosAbertos.DistribuirValorParaBaixarAdiantamentos(vlrBaixa, eTiposMoeda.Oficial, True)
                gridAdiantamentosDisponiveis.DataSource = objTitulos.AdiantamentosAbertos
                gridAdiantamentosDisponiveis.DataBind()

                If FinanceiroNovo Then
                    gridVencimentos.DataSource = objTitulos
                Else
                    gridVencimentos.DataSource = objTitulosFrete
                End If
                gridVencimentos.DataBind()

                For Each tit In objTitulos
                    objFrete.ListTituloFatura.Add(New FaturaDeFretexTitulo(objFrete) With {.CodigoTitulo = tit.Codigo, .TituloNovo = tit})
                Next

                SessaoSalvaTitulos()
                SessaoSalvaFrete()
            End If

            If IUD.Equals("I") Then
                objTitulos.Parcelar(ddlCondicaoPagamento.SelectedValue, SetarOrigem, ddlFormaPagamentoGeral.SelectedValue)
                If SetarOrigem = "FX" Then
                    SessaoRecuperaFixacao()
                    Dim vlrFixacao = objFixacao.NotasFixacao.Sum(Function(s) s.ValorAFixarLiberado)
                    If vlrFixacao = objFixacao.TotalOficial Then
                        For Each tit In objTitulos
                            tit.CodigoProvisao = eProvisao.Provisao
                        Next
                    End If
                End If
                If FinanceiroNovo Then
                    gridVencimentos.DataSource = objTitulos
                Else
                    gridVencimentos.DataSource = objTitulosFrete
                End If
                gridVencimentos.DataBind()
            ElseIf Not SetarOrigem = "FRETE" Then
                AtualizarValorNotaOuFixacaoOuTroca()
            End If
        End If
    End Sub

    Protected Sub BtnConta_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnConta.Click
        SessaoRecuperaTitulos()
        Dim i As Integer = gridVencimentos.SelectedIndex
        If objTitulos Is Nothing OrElse Not objTitulos.Count > 0 OrElse objTitulos(i).CodigoProvisao = 1 Then Exit Sub

        Dim cliente As [Lib].Negocio.Cliente = objTitulos(i).CliFor
        Dim ucConsultaDadosBancarios As ucConsultaDadosBancarios = CType(Me.Page.FindControlRecursive("ucConsultaDadosBancarios"), ucConsultaDadosBancarios)
        If cliente Is Nothing OrElse String.IsNullOrWhiteSpace(cliente.Codigo) Then
            MsgBox(Me.Page, "Selecione um cliente para continuar!")
            Exit Sub
        End If

        ucConsultaDadosBancarios.MainUserControl = Me
        ucConsultaDadosBancarios.CarregaGrid(cliente.Codigo, cliente.CodigoEndereco)
        Popup.ConsultaDeDadosBancarios(Me.Page, "objBancoUcFinan" & HID.Value)
    End Sub

    Protected Sub ddlCondicaoPagamento_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ddlCondicaoPagamento.SelectedIndexChanged
        ' fluxo original estava comentado
    End Sub

    Protected Sub gridVencimentos_SelectedIndexChanged(ByVal sender As Object, ByVal e As EventArgs) Handles gridVencimentos.SelectedIndexChanged
        SessaoRecuperaTitulos()

        Dim codigoSel As Integer = 0
        Integer.TryParse(gridVencimentos.SelectedRow.Cells(1).Text, codigoSel)

        If FinanceiroNovo Then
            Dim i As Integer = If(codigoSel = 0, gridVencimentos.SelectedIndex, objTitulos.FindIndex(Function(s) s.Codigo = codigoSel))
            IndiceGrid.Value = i

            Select Case SetarOrigem
                Case "NF"
                    ValorDocumento.Value = If(objTitulos(i).Moeda.Classificacao = eTiposMoeda.Oficial, objTitulos(i).Valores.EncargoValorDocumento.ValorOficial, objTitulos(i).Valores.EncargoValorDocumento.ValorMoeda)
                Case "PD"
                    ValorDocumento.Value = If(objTitulos.PD.Moeda.Classificacao = eTiposMoeda.Oficial, objTitulos(i).Valores.EncargoValorDocumento.ValorOficial, objTitulos(i).Valores.EncargoValorDocumento.ValorMoeda)
                Case "FX"
                    ValorDocumento.Value = If(objTitulos.FX.ItemPedido.Pedido.Moeda.Classificacao = eTiposMoeda.Oficial, objTitulos(i).Valores.EncargoValorDocumento.ValorOficial, objTitulos(i).Valores.EncargoValorDocumento.ValorMoeda)
            End Select

            AtualizaValores()
        Else
            Dim i As Integer = If(codigoSel = 0, gridVencimentos.SelectedIndex, objTitulosFrete.FindIndex(Function(s) s.Codigo = codigoSel))
            IndiceGrid.Value = i

            If objTitulosFrete(i).Moeda IsNot Nothing AndAlso objTitulosFrete(i).Moeda.Classificacao = eTiposMoeda.Oficial Then
                ValorDocumento.Value = objTitulosFrete(i).Valores.EncargoValorDocumento.ValorOficial
            Else
                ValorDocumento.Value = objTitulosFrete(i).Valores.EncargoValorDocumento.ValorMoeda
            End If
        End If
    End Sub

    Protected Sub BtnAdicionarConta_Click(sender As Object, e As EventArgs) Handles BtnAdicionarConta.Click
        SessaoRecuperaTitulos()
        If IndiceGrid.Value < 0 Then Exit Sub

        If (objTitulos(IndiceGrid.Value).ContaContabilCliFor IsNot Nothing AndAlso objTitulos(IndiceGrid.Value).ContaContabilCliFor.EncargosPlanoDeContas IsNot Nothing) Then
            Dim ucConsultaEncargosPlanoDeContas As ucConsultaEncargosPlanoDeContas = CType(Me.Page.FindControlRecursive("ucConsultaEncargosPlanoDeContas"), ucConsultaEncargosPlanoDeContas)
            ucConsultaEncargosPlanoDeContas.SetarHID(HID.Value)
            Session("EncargosPlanoDeContas" & HID.Value) = objTitulos(IndiceGrid.Value).ContaContabilCliFor.EncargosPlanoDeContas
            ucConsultaEncargosPlanoDeContas.BindGridView(objTitulos(IndiceGrid.Value).ContaContabilCliFor.EncargosPlanoDeContas)
            ucConsultaEncargosPlanoDeContas.MainUserControl = Me
            Popup.ConsultaDeEncargosPlanoDeContas(Me.Page, "EncargosPlanoDeContas" & HID.Value, "btnSelecionar")
        End If
    End Sub

    Protected Sub gridVencimentos_RowDataBound(sender As Object, e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles gridVencimentos.RowDataBound
        SessaoRecuperaTitulos()
        If e.Row.RowType = DataControlRowType.DataRow Then
            Select Case e.Row.Cells(2).Text
                Case "R" : e.Row.Cells(2).Text = "Receber"
                Case "P" : e.Row.Cells(2).Text = "Pagar"
                Case "C" : e.Row.Cells(2).Text = "Contabil"
            End Select

            If FinanceiroNovo Then
                If objTitulos.ReajFinanceiro IsNot Nothing AndAlso objTitulos.ReajFinanceiro.TitulosNew.Count > 0 Then
                    Dim index As Integer = objTitulos.ReajFinanceiro.TitulosNew.FindIndex(Function(s) s.Codigo = CInt(e.Row.Cells(1).Text))
                    If index >= 0 AndAlso (objTitulos.ReajFinanceiro.TitulosNew(index).isAdiantamento Or objTitulos.ReajFinanceiro.TitulosNew(index).isBaixaAdiantamento) Then
                        e.Row.BackColor = Drawing.Color.Yellow
                    End If
                    If objTitulos(index).Pedido IsNot Nothing AndAlso objTitulos(index).Pedido.MomentoFinanceiro = eTipoFaturamento.Lote AndAlso objTitulos(index).CodigoProvisao = eProvisao.Previsao Then
                        e.Row.Cells(0).Controls(0).Visible = False
                        e.Row.BackColor = Drawing.Color.Silver
                        e.Row.Cells(2).Text = "LOTE"
                        Dim quantidade As Decimal = objTitulos(index).Quantidade - objTitulos.Where(Function(s) s.CodigoTituloOrigem = objTitulos(index).Codigo AndAlso s.CodigoProvisao = eProvisao.Provisao).Sum(Function(s) s.Quantidade)
                        e.Row.Cells(11).Text = quantidade
                        e.Row.Cells(3).Text = If(quantidade > 0, "Em Andamento", "Encerrado")
                    End If
                End If
            Else
                If objTitulosFrete.ReajFinanceiro IsNot Nothing AndAlso objTitulosFrete.ReajFinanceiro.TitulosNew.Count > 0 Then
                    Dim index As Integer = objTitulosFrete.ReajFinanceiro.TitulosNew.FindIndex(Function(s) s.Codigo = CInt(e.Row.Cells(1).Text))
                    If index >= 0 AndAlso (objTitulosFrete.ReajFinanceiro.TitulosNew(index).isAdiantamento Or objTitulosFrete.ReajFinanceiro.TitulosNew(index).isBaixaAdiantamento) Then
                        e.Row.BackColor = Drawing.Color.Yellow
                    End If
                Else
                    Dim index As Integer = objTitulosFrete.FindIndex(Function(s) s.Codigo = CInt(e.Row.Cells(1).Text))
                    If objTitulosFrete(index).isAdiantamento Or objTitulosFrete(index).isBaixaAdiantamento Then
                        e.Row.BackColor = Drawing.Color.Yellow
                    End If
                    If Not objTitulosFrete(index).Pedido Is Nothing AndAlso objTitulosFrete(index).Pedido.MomentoFinanceiro = eTipoFaturamento.Lote AndAlso objTitulosFrete(index).CodigoProvisao = eProvisao.Previsao Then
                        e.Row.Cells(0).Controls(0).Visible = False
                        e.Row.BackColor = Drawing.Color.Silver
                        e.Row.Cells(2).Text = "LOTE"
                        Dim quantidade As Decimal = objTitulosFrete(index).Quantidade - objTitulosFrete.Where(Function(s) s.CodigoTituloOrigem = objTitulosFrete(index).Codigo AndAlso s.CodigoProvisao = eProvisao.Provisao).Sum(Function(s) s.Quantidade)
                        e.Row.Cells(11).Text = quantidade
                        e.Row.Cells(3).Text = If(quantidade > 0, "Em Andamento", "Encerrado")
                    End If
                End If
            End If
        End If
    End Sub

    Protected Sub gridValores_RowDataBound(sender As Object, e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles gridValores.RowDataBound
        If e.Row.RowType = DataControlRowType.DataRow Then
            Dim txtValor As TextBox
            Dim txtValor2 As TextBox

            SessaoRecuperaTitulos()

            If objTitulos(IndiceGrid.Value).Moeda.Classificacao = eTiposMoeda.Oficial Then
                txtValor = CType(e.Row.FindControl("txtValorOficial"), TextBox)
                txtValor2 = CType(e.Row.FindControl("txtValorMoeda"), TextBox)
            Else
                txtValor = CType(e.Row.FindControl("txtValorMoeda"), TextBox)
                txtValor2 = CType(e.Row.FindControl("txtValorOficial"), TextBox)
            End If

            txtValor.Enabled = True
            txtValor.BackColor = Drawing.Color.White

            txtValor2.Enabled = False
            txtValor2.BackColor = Drawing.Color.LightYellow
        End If
    End Sub

    Protected Sub txtVlrBaixa_TextChanged(sender As Object, e As EventArgs)
        Dim txt As TextBox = CType(sender, TextBox)
        Dim row As GridViewRow = CType(txt.NamingContainer, GridViewRow)

        SessaoRecuperaTitulos()
        If Not IsNumeric(txt.Text) Then txt.Text = 0
        objTitulos.AdiantamentosAbertos(row.RowIndex).VlrBaixa = Math.Abs(CDec(txt.Text))

        If objTitulos.Count > 0 AndAlso objTitulos.AdiantamentosAbertos(row.RowIndex).VlrBaixa > 0 Then
            LnkParcelamento_Click(LnkParcelamento, Nothing)
        End If
        SessaoSalvaTitulos()
    End Sub

    Protected Sub gridAdiantamentosDisponiveis_RowDataBound(sender As Object, e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles gridAdiantamentosDisponiveis.RowDataBound
        If e.Row.RowType = DataControlRowType.DataRow Then
            Dim index As Integer = 0

            Select Case SetarOrigem
                Case "NF"
                    SessaoRecuperaNotaFiscal()
                Case "PD"
                    SessaoRecuperaPedido()
                    index = objTitulos.AdiantamentosAbertos.FindIndex(Function(s) s.CodigoTitulo = CInt(e.Row.Cells(0).Text))
                    If objTitulos.AdiantamentosAbertos(index).Titulo.CodigoPedido <> objTitulos.NF.CodigoPedido And Not objTitulos.PD.Troca Then
                        If objTitulos.AdiantamentosAbertos(index).Titulo.PedidoTroca And objTitulos.AdiantamentosAbertos(index).Titulo.Pedido.FinanceiroAberto Then
                            e.Row.BackColor = Drawing.Color.Red
                        End If
                    End If
            End Select
        End If
    End Sub

    Private Function txtPracaAgencia() As Object
        Throw New NotImplementedException
    End Function

    Protected Sub lstCondicoesPgtoEntrega_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lstCondicoesPgtoEntrega.SelectedIndexChanged
        Try
            SessaoRecuperaPedido()
            If lstCondicoesPgtoEntrega.SelectedIndex > 0 Then
                objPedido.CodigoCondicaoPagamentoDaEntrega = lstCondicoesPgtoEntrega.SelectedValue
                objPedido.PeriodicidadeEntrega = ddlTipoFaturamento.SelectedValue
                SessaoSalvaPedido()
            End If
        Catch ex As Exception
            MsgBox(Me.Page, Funcoes.EliminarCaracteresEspeciais(ex.ToString))
        End Try
    End Sub

    Protected Sub txtQuotaDeEntrega_TextChanged(sender As Object, e As EventArgs) Handles txtQuotaDeEntrega.TextChanged
        If Not String.IsNullOrWhiteSpace(txtQuotaDeEntrega.Text) Then
            SessaoRecuperaPedido()
            objPedido.QuotaEntrega = txtQuotaDeEntrega.Text
            SessaoSalvaPedido()
        End If
    End Sub

    Protected Sub ddlTipoFaturamento_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddlTipoFaturamento.SelectedIndexChanged
        SessaoRecuperaPedido()
        SessaoRecuperaTitulos()

        objTitulos.Clear()

        If FinanceiroNovo Then
            gridVencimentos.DataSource = objTitulos
        Else
            gridVencimentos.DataSource = objTitulosFrete
        End If
        gridVencimentos.DataBind()

        objPedido.MomentoFinanceiro = ddlTipoFaturamento.SelectedValue
        objPedido.PeriodicidadeEntrega = ddlTipoFaturamento.SelectedValue
        divLinhaQdteLote.Visible = False
        divLinhaPeridiocidade.Visible = False

        If objPedido.MomentoFinanceiro = eTipoFaturamento.Pedido OrElse
           objPedido.MomentoFinanceiro = eTipoFaturamento.NotaFiscal Then
            divLinhaCondicaoPagamento.Visible = True
            divFormaDePagamento.Visible = True
        ElseIf objPedido.MomentoFinanceiro = eTipoFaturamento.Lote Then
            divLinhaCondicaoPagamento.Visible = True
            divFormaDePagamento.Visible = True
            divLinhaQdteLote.Visible = True
        ElseIf objPedido.MomentoFinanceiro = eTipoFaturamento.Nenhum Then
            divLinhaCondicaoPagamento.Visible = False
            divFormaDePagamento.Visible = False
        ElseIf objPedido.MomentoFinanceiro = eTipoFaturamento.Peridiocidade Then
            divLinhaCondicaoPagamento.Visible = True
            divFormaDePagamento.Visible = True
            divLinhaPeridiocidade.Visible = True
        End If

        ddlCondicaoPagamento.ClearSelection()
        SessaoSalvaTitulos()
        SessaoSalvaPedido()
    End Sub

    Protected Sub btnAtualizarTitulo_Click(sender As Object, e As EventArgs) Handles btnAtualizarTitulo.Click
        Try
            If Not String.IsNullOrWhiteSpace(txtCodigoTitulo.Text) Then
                SessaoRecuperaTitulos()
                If objTitulos(gridVencimentos.SelectedIndex).IUD <> "I" Then objTitulos(gridVencimentos.SelectedIndex).IUD = "U"
                objTitulos(gridVencimentos.SelectedIndex).Reprogramacao = Funcoes.ValidaDataUtil(objTitulos(gridVencimentos.SelectedIndex).CodigoEmpresa, objTitulos(gridVencimentos.SelectedIndex).EnderecoEmpresa, txtDataVencimento.Text)
                objTitulos(gridVencimentos.SelectedIndex).Vencimento = Funcoes.ValidaDataUtil(objTitulos(gridVencimentos.SelectedIndex).CodigoEmpresa, objTitulos(gridVencimentos.SelectedIndex).EnderecoEmpresa, txtDataVencimento.Text)
                If chkPreImpresso.Checked = False Then
                    If Not String.IsNullOrWhiteSpace(ddlTipoDePagto.SelectedValue) AndAlso ddlTipoDePagto.SelectedValue = 4 Then
                        If Not String.IsNullOrWhiteSpace(txtCodigoDeBarra.Text) Then
                            If Funcoes.ValidaCodigoBarras(txtCodigoDeBarra.Text, chkDigitado.Checked, txtDataVencimento.Text, objTitulos(gridVencimentos.SelectedIndex).Valores.EncargoValorDocumento.Valor, objTitulos(gridVencimentos.SelectedIndex).CodigoEmpresa, objTitulos(gridVencimentos.SelectedIndex).EnderecoEmpresa, Banco) Then
                                objTitulos(gridVencimentos.SelectedIndex).CodigoDeBarras = txtCodigoDeBarra.Text
                                If FinanceiroNovo Then
                                    gridVencimentos.DataSource = objTitulos
                                Else
                                    gridVencimentos.DataSource = objTitulosFrete
                                End If
                                gridVencimentos.DataBind()
                                If chkDigitado.Checked Then txtCodigoDeBarra.Text = Funcoes.FormataLinhaDigitavelOriginal(txtCodigoDeBarra.Text)
                            Else
                                MsgBox(Me.Page, Funcoes.EliminarCaracteresEspeciais(RTrim(HttpContext.Current.Session("ssMessage").ToString)))
                                Exit Sub
                            End If
                        End If
                    End If
                    LimparValores()
                    SessaoSalvaTitulos()
                Else
                    Throw New Exception("Nenhum registro selecionado!")
                End If
            End If

        Catch ex As Exception
            MsgBox(Me.Page, ex.Message.ToString())
        End Try
    End Sub

    Protected Sub gridBaixasAdiantamentos_SelectedIndexChanged(sender As Object, e As EventArgs)
    End Sub

    Protected Sub LnkReajuste_Click(sender As Object, e As EventArgs) Handles LnkReajuste.Click
        If SetarOrigem = "NF" Then
            SessaoRecuperaTitulos()
            If objNotaFiscal.IUD <> "U" Then
                MsgBox(Me.Page, "O Reajuste Financeiro somente é realizado em alteração de valores, da nota")
                Exit Sub
            ElseIf objNotaFiscal.TotalNota = objNotaFiscal.NotaFiscalOriginal.TotalNota Then
                objNotaFiscal.Titulos.ReajFinanceiro = Nothing
                gridBaixasAdiantamentos.DataSource = objTitulos.SelectMany(Function(s) s.Baixas_AdiantamentoEfetuadas)
                gridBaixasAdiantamentos.DataBind()

                If FinanceiroNovo Then
                    gridVencimentos.DataSource = objTitulos
                Else
                    gridVencimentos.DataSource = objTitulosFrete
                End If
                gridVencimentos.DataBind()
            ElseIf objNotaFiscal.IUD = "U" Then
                Dim resp As ArrayList = objNotaFiscal.Titulos.ReajFinanceiro.Reajusta

                If resp(0) Then
                    AtualizaGridReajustaFinanceiro(objNotaFiscal.Titulos.ReajFinanceiro.TitulosNotaModificada)
                Else
                    MsgBox(Me.Page, resp(1))
                End If
            End If
        End If
    End Sub

    Protected Sub LnkLimpar_Click(sender As Object, e As EventArgs) Handles lnkLimpar.Click
        Try
            Limpar()
            SessaoRecuperaTitulos()
            objTitulos.Clear()
            SessaoSalvaTitulos()
        Catch ex As Exception
            MsgBox(Me.Page, "Erro ao limpar o financeiro!")
        End Try
    End Sub

    ' ----- AJUSTE: filtro de situação com bind condicional
    Protected Sub ddlSituacao_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddlSituacao.SelectedIndexChanged
        SessaoRecuperaTitulos()

        Select Case ddlSituacao.SelectedValue
            Case "0" : If FinanceiroNovo Then
                    gridVencimentos.DataSource = objTitulos
                Else
                    gridVencimentos.DataSource = objTitulosFrete
                End If
        End Select
        gridVencimentos.DataBind()
    End Sub

End Class
