Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class ucVencimentos
    Inherits BaseUserControl

#Region "NotaFiscal"
    Private objNotaFiscal As [Lib].Negocio.NotaFiscal

    Private Sub SessaoSalvaNotaFiscal()
        Session("objNotaFiscal" & HID.Value) = objNotaFiscal
    End Sub

    Private Sub SessaoRecuperaNotaFiscal()
        objNotaFiscal = CType(Session("objNotaFiscal" & HID.Value), [Lib].Negocio.NotaFiscal)
    End Sub
#End Region

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If IsPostBack Then
            Return
        End If
    End Sub

    Public Overrides Sub SetarHID(ByVal NewGuid As String)
        HID.Value = NewGuid.ToString
    End Sub

    Public Overrides Sub Limpar()
        IndiceGrid.Value = "-1"
        Origem.Value = String.Empty
        txtCodigoTitulo.Text = String.Empty
        txtDataVencimento.Text = String.Empty
        chkDigitado.Checked = False
        txtCodigoDeBarra.Text = String.Empty
        txtMoeda.Text = String.Empty
        ValorDocumento.Value = String.Format("{0:N0}", Decimal.Zero)
        txtValorDocumento.Text = String.Empty
        txtDesconto.Text = String.Empty
        txtDeducao.Text = String.Empty
        txtMultaJuro.Text = String.Empty
        txtOutrosAcrescimos.Text = String.Empty
        txtValorLiquido.Text = String.Empty
        txtCodBanco.Text = String.Empty
        txtNomeBanco.Text = String.Empty
        txtCodAgencia.Text = String.Empty
        txtDigitoAgencia.Text = String.Empty
        txtPracaAgencia.Text = String.Empty
        txtConta.Text = String.Empty
        txtDigitoConta.Text = String.Empty
        gridVencimentos.SelectedIndex = -1
    End Sub

    Public Overrides Sub Carregar(ByVal obj As IBaseEntity)
        If Not Session("objContaBancariaVencimentos" & HID.Value) Is Nothing Then
            txtCodBanco.Text = CType(Session("objContaBancariaVencimentos" & HID.Value), [Lib].Negocio.ClienteXContaBancaria).CodigoBanco
            txtNomeBanco.Text = CType(Session("objContaBancariaVencimentos" & HID.Value), [Lib].Negocio.ClienteXContaBancaria).NomeBanco
            txtCodAgencia.Text = CType(Session("objContaBancariaVencimentos" & HID.Value), [Lib].Negocio.ClienteXContaBancaria).CodigoAgencia
            txtDigitoAgencia.Text = CType(Session("objContaBancariaVencimentos" & HID.Value), [Lib].Negocio.ClienteXContaBancaria).DigitoAgencia
            txtPracaAgencia.Text = CType(Session("objContaBancariaVencimentos" & HID.Value), [Lib].Negocio.ClienteXContaBancaria).Praca
            txtConta.Text = CType(Session("objContaBancariaVencimentos" & HID.Value), [Lib].Negocio.ClienteXContaBancaria).ContaCorrente
            txtDigitoConta.Text = CType(Session("objContaBancariaVencimentos" & HID.Value), [Lib].Negocio.ClienteXContaBancaria).DigitoConta
            Session.Remove("objContaBancariaVencimentos" & HID.Value)
            Session("ssTipoRetorno") = "objVencimentosNxI" & HID.Value
            SessaoRecuperaNotaFiscal()
            If txtCodBanco.Text.Length > 0 Then
                Dim i As Integer = gridVencimentos.SelectedIndex
                objNotaFiscal.VencimentosNota(i).CodigoBancoCliente = txtCodBanco.Text
                objNotaFiscal.VencimentosNota(i).CodigoAgenciaCliente = txtCodAgencia.Text
                objNotaFiscal.VencimentosNota(i).DigitoAgenciaCliente = txtDigitoAgencia.Text
                objNotaFiscal.VencimentosNota(i).ContaCliente = txtConta.Text
                objNotaFiscal.VencimentosNota(i).DigitoContaCliente = txtDigitoConta.Text
            End If
            SessaoSalvaNotaFiscal()
        End If
    End Sub

    Public Sub BindGridView(ByVal parameters As Dictionary(Of String, Object))

        SessaoRecuperaNotaFiscal()

        objNotaFiscal.CarregarVencimentosDaNota(Origem.Value, ddlCarteira, ddlCondicaoPagamento, ddlTipoDePagto, parameters)

        If Origem.Value = "NF" Then
            gridVencimentos.DataSource = objNotaFiscal.VencimentosNota
            gridVencimentos.DataBind()
            SessaoSalvaNotaFiscal()
        End If

    End Sub

    Protected Sub BtnConta_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        SessaoRecuperaNotaFiscal()
        If (gridVencimentos.SelectedIndex >= 0) Then
            Dim i As Integer = gridVencimentos.SelectedIndex
            If objNotaFiscal.VencimentosNota(i).CodigoProvisao = 1 Then Exit Sub
            Dim objCliente As [Lib].Negocio.Cliente = objNotaFiscal.Cliente
            Dim ucConsultaDadosBancarios As ucConsultaDadosBancarios = DirectCast(Me.NamingContainer.FindControl("ucConsultaDadosBancarios"), ucConsultaDadosBancarios)
            If ucConsultaDadosBancarios IsNot Nothing Then
                ucConsultaDadosBancarios.Limpar()
                ucConsultaDadosBancarios.MainUserControl = Me
                ucConsultaDadosBancarios.CarregaGrid(objCliente.Codigo, objCliente.CodigoEndereco)
                Popup.ConsultaDeDadosBancarios(Me.Page, "objContaBancariaVencimentos" & HID.Value)
            End If
        Else
            MsgBox(Me.Page, "É obrigatório selecionar o título!")
        End If
    End Sub

    Protected Sub btnConfirmar_Click(ByVal sender As Object, ByVal e As EventArgs)
        If ddlCarteira.SelectedIndex = 0 Then
            MsgBox(Me.Page, "É obrigatório selecionar a Carteira Financeira para o(s) título(s)!")
            Exit Sub
        End If

        Session("Vencimentos" & HID.Value) = True
        If (Session("ssTipoRetorno") IsNot Nothing AndAlso Session("ssTipoRetorno").ToString.Contains("objVencimentosNxI")) Then
            If TypeOf Me.Page Is NotaFiscalXItens Then
                CType(Me.Page, NotaFiscalXItens).CarregarVencimentos()
            ElseIf TypeOf Me.Page Is AlterarNotaFiscal Then
                CType(Me.Page, AlterarNotaFiscal).SalvarVencimentos()
            End If
        End If
        Popup.CloseDialog(Me.Page, "divConsultaVencimentos")
    End Sub

    Protected Sub btnSair_Click(ByVal sender As Object, ByVal e As EventArgs)
        Popup.CloseDialog(Me.Page, "divConsultaVencimentos")
    End Sub

    Protected Sub ddlCarteira_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        If ddlCarteira.SelectedIndex = 0 Then
            MsgBox(Me.Page, "É obrigatório selecionar a Carteira Financeira para o(s) título(s)!")
            Exit Sub
        End If

        SessaoRecuperaNotaFiscal()

        If objNotaFiscal.VencimentosNota IsNot Nothing AndAlso objNotaFiscal.VencimentosNota.Count > 0 Then
            For Each v In objNotaFiscal.VencimentosNota
                v.CodigoCarteira = ddlCarteira.SelectedValue
            Next

            SessaoSalvaNotaFiscal()
        End If
    End Sub
    Protected Sub ddlCondicaoPagamento_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        If ddlCarteira.SelectedIndex = 0 Then
            MsgBox(Me.Page, "É obrigatório selecionar a Carteira Financeira para o(s) título(s)!")
            Exit Sub
        End If

        IndiceGrid.Value = -1
        ValorDocumento.Value = 0
        SessaoRecuperaNotaFiscal()
        If objNotaFiscal.VencimentosNota Is Nothing Then objNotaFiscal.VencimentosNota = New [Lib].Negocio.ListTitulo(objNotaFiscal)
        objNotaFiscal.VencimentosNota.ParcelarNota(ddlCondicaoPagamento.SelectedValue)

        If objNotaFiscal.VencimentosNota IsNot Nothing AndAlso objNotaFiscal.VencimentosNota.Count > 0 Then
            For Each v In objNotaFiscal.VencimentosNota
                If ddlCarteira.SelectedIndex > 0 Then v.CodigoCarteira = ddlCarteira.SelectedValue
            Next
        End If

        gridVencimentos.DataSource = objNotaFiscal.VencimentosNota
        gridVencimentos.DataBind()
        SessaoSalvaNotaFiscal()
    End Sub

    Protected Sub gridVencimentos_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        SessaoRecuperaNotaFiscal()

        Dim Sql As String = "Select isnull(AVista,0) as Avista " & vbCrLf &
                                     "from Pagamentos " & vbCrLf &
                                     "where Pagamento_id = " & objNotaFiscal.Pedido.CodigoCondicaoPagamento

        Dim ds As DataSet = Banco.ConsultaDataSet(Sql, "Pagamentos")

        If ds IsNot Nothing AndAlso ds.Tables(0).Rows.Count > 0 Then
            Dim temAVista As Boolean = False

            For Each row As DataRow In ds.Tables(0).Rows
                If CBool(row("Avista")) AndAlso objNotaFiscal.Cliente.ApenasAVista Then
                    temAVista = True
                End If
            Next

            If temAVista Then
                MsgBox(Me.Page, "Condição para esse cliente é apenas à Vista e vencimento não pode ser modificado.", eTitulo.Info)
                Exit Sub
            End If
        End If

        If objNotaFiscal.Pedido.Moeda.Classificacao = [Lib].Negocio.eTiposMoeda.Oficial Then
            If CInt(IndiceGrid.Value) >= 0 AndAlso CDec(ValorDocumento.Value) <> objNotaFiscal.VencimentosNota(CInt(IndiceGrid.Value)).ValorDoDocumento Then
                objNotaFiscal.VencimentosNota(CInt(IndiceGrid.Value)).ValorDoDocumento = CDec(ValorDocumento.Value)
            End If
            ValorDocumento.Value = objNotaFiscal.VencimentosNota(gridVencimentos.SelectedIndex).ValorDoDocumento
        Else
            If CInt(IndiceGrid.Value) >= 0 AndAlso CDec(ValorDocumento.Value) <> objNotaFiscal.VencimentosNota(CInt(IndiceGrid.Value)).MoedaValorDoDocumento Then
                objNotaFiscal.VencimentosNota(CInt(IndiceGrid.Value)).MoedaValorDoDocumento = CDec(ValorDocumento.Value)
            End If
            ValorDocumento.Value = objNotaFiscal.VencimentosNota(gridVencimentos.SelectedIndex).MoedaValorDoDocumento
        End If
        IndiceGrid.Value = gridVencimentos.SelectedIndex
        AtualizaValores()
    End Sub

    Private Sub SelecionarIndiceCombo(ByVal Combo As DropDownList, ByVal Valor As String)
        If Valor = "0" Then
            Combo.SelectedIndex = 0
        Else
            Combo.SelectedIndex = Combo.Items.IndexOf(Combo.Items.FindByValue(Valor))
        End If
    End Sub

    Private Sub Salvar()
        If gridVencimentos.SelectedIndex = -1 Then
            MsgBox(Me.Page, "Selecione um título para alteração!")
            Exit Sub
        End If

        SessaoRecuperaNotaFiscal()
        If TituloBaixado() Then
            AtualizaValores()
            Exit Sub
        End If

        Dim msg As String = ""
        Dim i As Integer = gridVencimentos.SelectedIndex
        If objNotaFiscal.VencimentosNota(i).IUD <> "I" Then objNotaFiscal.VencimentosNota(i).IUD = "U"

        If objNotaFiscal.Pedido.Moeda.Classificacao = [Lib].Negocio.eTiposMoeda.Oficial Then
            If objNotaFiscal.VencimentosNota(i).ValorDoDocumento <> CDec(ValorDocumento.Value) Then
                msg = objNotaFiscal.VencimentosNota.AjustaParcelas(i, ValorDocumento.Value, objNotaFiscal.VencimentosNota(i).ValorDoDocumento)
                If msg.Length > 0 Then
                    MsgBox(Me.Page, msg)
                End If
            End If
        Else
            If objNotaFiscal.VencimentosNota(i).MoedaValorDoDocumento <> CDec(ValorDocumento.Value) Then
                msg = objNotaFiscal.VencimentosNota.AjustaParcelas(i, ValorDocumento.Value, objNotaFiscal.VencimentosNota(i).MoedaValorDoDocumento)
                If msg.Length > 0 Then
                    MsgBox(Me.Page, msg)
                End If
            End If
        End If

        gridVencimentos.DataSource = objNotaFiscal.VencimentosNota
        gridVencimentos.DataBind()
        SessaoSalvaNotaFiscal()
        If msg = "" Then
            ValorDocumento.Value = CDec(txtValorDocumento.Text)
        End If
        AtualizaValores()
    End Sub

    Protected Sub txtValorDocumento_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        SessaoRecuperaNotaFiscal()
        If Not gridVencimentos.Rows.Count > 0 Then
            Exit Sub
        End If

        Dim i As Integer = gridVencimentos.SelectedIndex
        If i > -1 Then
            If TituloBaixado() Then Exit Sub
            If objNotaFiscal.Pedido.Moeda.Classificacao = [Lib].Negocio.eTiposMoeda.Oficial Then
                objNotaFiscal.VencimentosNota(i).ValorDoDocumento = txtValorDocumento.Text
            Else
                objNotaFiscal.VencimentosNota(i).MoedaValorDoDocumento = txtValorDocumento.Text
            End If
            SessaoSalvaNotaFiscal()
            Salvar()
        End If
    End Sub

    Protected Sub txtDesconto_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        SessaoRecuperaNotaFiscal()
        If Not gridVencimentos.Rows.Count > 0 Then
            Exit Sub
        End If

        Dim i As Integer = gridVencimentos.SelectedIndex
        If i > -1 Then
            If TituloBaixado() Then Exit Sub
            If objNotaFiscal.Pedido.Moeda.Classificacao = [Lib].Negocio.eTiposMoeda.Oficial Then
                objNotaFiscal.VencimentosNota(i).Descontos = txtDesconto.Text
            Else
                objNotaFiscal.VencimentosNota(i).MoedaDescontos = txtDesconto.Text
            End If
            SessaoSalvaNotaFiscal()
            AtualizaValores()
        End If
    End Sub

    Protected Sub txtDeducao_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        SessaoRecuperaNotaFiscal()
        If Not gridVencimentos.Rows.Count > 0 Then
            Exit Sub
        End If

        Dim i As Integer = gridVencimentos.SelectedIndex
        If i > -1 Then
            If TituloBaixado() Then Exit Sub
            If objNotaFiscal.Pedido.Moeda.Classificacao = [Lib].Negocio.eTiposMoeda.Oficial Then
                objNotaFiscal.VencimentosNota(i).Deducoes = txtDeducao.Text
            Else
                objNotaFiscal.VencimentosNota(i).MoedaDeducoes = txtDeducao.Text
            End If
            SessaoSalvaNotaFiscal()
            AtualizaValores()
        End If
    End Sub

    Protected Sub txtMultaJuro_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        SessaoRecuperaNotaFiscal()
        If Not gridVencimentos.Rows.Count > 0 Then
            Exit Sub
        End If

        Dim i As Integer = gridVencimentos.SelectedIndex
        If i > -1 Then
            If TituloBaixado() Then Exit Sub
            If objNotaFiscal.Pedido.Moeda.Classificacao = [Lib].Negocio.eTiposMoeda.Oficial Then
                objNotaFiscal.VencimentosNota(i).Juros = txtMultaJuro.Text
            Else
                objNotaFiscal.VencimentosNota(i).MoedaJuros = txtMultaJuro.Text
            End If
            SessaoSalvaNotaFiscal()
            AtualizaValores()
        End If
    End Sub

    Protected Sub txtOutrosAcrescimos_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        SessaoRecuperaNotaFiscal()
        If Not gridVencimentos.Rows.Count > 0 Then
            Exit Sub
        End If

        Dim i As Integer = gridVencimentos.SelectedIndex
        If i > -1 Then
            If TituloBaixado() Then Exit Sub
            If objNotaFiscal.Pedido.Moeda.Classificacao = [Lib].Negocio.eTiposMoeda.Oficial Then
                objNotaFiscal.VencimentosNota(i).Acrescimos = txtOutrosAcrescimos.Text
            Else
                objNotaFiscal.VencimentosNota(i).MoedaAcrescimos = txtOutrosAcrescimos.Text
            End If
            SessaoSalvaNotaFiscal()
            AtualizaValores()
        End If
    End Sub

    Protected Sub ddlTipoDePagto_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        SessaoRecuperaNotaFiscal()
        If Not gridVencimentos.Rows.Count > 0 Then
            Exit Sub
        End If
        If gridVencimentos.SelectedIndex > -1 Then
            objNotaFiscal.VencimentosNota(gridVencimentos.SelectedIndex).CodigoTipoPgto = ddlTipoDePagto.SelectedValue
            SessaoSalvaNotaFiscal()
        End If
    End Sub

    Protected Sub chkDigitado_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        SessaoRecuperaNotaFiscal()
        If Not gridVencimentos.Rows.Count > 0 Then
            Exit Sub
        End If
        If gridVencimentos.SelectedIndex > -1 Then
            objNotaFiscal.VencimentosNota(gridVencimentos.SelectedIndex).CodigoDigitado = chkDigitado.Checked
            SessaoSalvaNotaFiscal()
        End If
    End Sub

    Protected Sub txtCodigoDeBarra_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        SessaoRecuperaNotaFiscal()
        If Not gridVencimentos.Rows.Count > 0 Then
            Exit Sub
        End If
        If gridVencimentos.SelectedIndex > -1 Then
            objNotaFiscal.VencimentosNota(gridVencimentos.SelectedIndex).CodigoDeBarras = txtCodigoDeBarra.Text
            SessaoSalvaNotaFiscal()
        End If
    End Sub

    Protected Sub txtDataVencimento_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDataVencimento.TextChanged
        If Not String.IsNullOrWhiteSpace(txtDataVencimento.Text) Then
            SessaoRecuperaNotaFiscal()
            If (Not gridVencimentos.Rows.Count > 0) OrElse (Not gridVencimentos.SelectedIndex > -1) Then
                Exit Sub
            End If
            If gridVencimentos.SelectedIndex > -1 Then
                objNotaFiscal.VencimentosNota(gridVencimentos.SelectedIndex).Vencimento = Funcoes.ValidaDataUtil(objNotaFiscal.Empresa.Codigo, objNotaFiscal.Empresa.CodigoEndereco, CDate(txtDataVencimento.Text))
                objNotaFiscal.VencimentosNota(gridVencimentos.SelectedIndex).Prorrogacao = Funcoes.ValidaDataUtil(objNotaFiscal.Empresa.Codigo, objNotaFiscal.Empresa.CodigoEndereco, CDate(txtDataVencimento.Text))
                txtDataVencimento.Text = String.Format("{0:dd/MM/yyyy}", Funcoes.ValidaDataUtil(objNotaFiscal.Empresa.Codigo, objNotaFiscal.Empresa.CodigoEndereco, CDate(txtDataVencimento.Text)))
                gridVencimentos.DataSource = objNotaFiscal.VencimentosNota.ToArray()
                gridVencimentos.DataBind()
                SessaoSalvaNotaFiscal()
            End If
        End If
    End Sub

    Private Function TituloBaixado() As Boolean
        If Not gridVencimentos.Rows.Count > 0 Then
            Return False
        End If

        Dim i As Integer = gridVencimentos.SelectedIndex
        If i > -1 Then
            If objNotaFiscal.VencimentosNota(i).CodigoProvisao = 1 Then
                AtualizaValores()
                Return True
            Else
                Return False
            End If
        End If
        Return False
    End Function

    Private Sub AtualizaValores()
        SessaoRecuperaNotaFiscal()
        If Not gridVencimentos.Rows.Count > 0 Then
            Exit Sub
        End If

        Dim i As Integer = gridVencimentos.SelectedIndex
        If i > -1 Then
            txtMoeda.Text = objNotaFiscal.VencimentosNota(i).Moeda.Descricao
            txtCodigoTitulo.Text = objNotaFiscal.VencimentosNota(i).Codigo
            txtDataVencimento.Text = objNotaFiscal.VencimentosNota(i).Vencimento.ToString("dd/MM/yyyy")
            chkDigitado.Checked = objNotaFiscal.VencimentosNota(i).CodigoDigitado
            txtCodigoDeBarra.Text = objNotaFiscal.VencimentosNota(i).CodigoDeBarras

            If objNotaFiscal.Pedido.Moeda.Classificacao = [Lib].Negocio.eTiposMoeda.Oficial Then
                txtValorDocumento.Text = objNotaFiscal.VencimentosNota(i).ValorDoDocumento.ToString("N2")
                txtDesconto.Text = objNotaFiscal.VencimentosNota(i).Descontos.ToString("N2")
                txtDeducao.Text = objNotaFiscal.VencimentosNota(i).Deducoes.ToString("N2")
                txtMultaJuro.Text = objNotaFiscal.VencimentosNota(i).Juros.ToString("N2")
                txtOutrosAcrescimos.Text = objNotaFiscal.VencimentosNota(i).Acrescimos.ToString("N2")
                txtValorLiquido.Text = objNotaFiscal.VencimentosNota(i).ValorLiquido.ToString("N2")
            Else
                txtValorDocumento.Text = objNotaFiscal.VencimentosNota(i).MoedaValorDoDocumento.ToString("N2")
                txtDesconto.Text = objNotaFiscal.VencimentosNota(i).MoedaDescontos.ToString("N2")
                txtDeducao.Text = objNotaFiscal.VencimentosNota(i).MoedaDeducoes.ToString("N2")
                txtMultaJuro.Text = objNotaFiscal.VencimentosNota(i).MoedaJuros.ToString("N2")
                txtOutrosAcrescimos.Text = objNotaFiscal.VencimentosNota(i).MoedaAcrescimos.ToString("N2")
                txtValorLiquido.Text = objNotaFiscal.VencimentosNota(i).MoedaValorLiquido.ToString("N2")
            End If

            SelecionarIndiceCombo(ddlTipoDePagto, objNotaFiscal.VencimentosNota(i).CodigoTipoPgto)

            If objNotaFiscal.VencimentosNota(i).CodigoBancoCliente > 0 Then
                txtCodBanco.Text = objNotaFiscal.VencimentosNota(i).CodigoBancoCliente
                txtNomeBanco.Text = objNotaFiscal.VencimentosNota(i).BancoCliente.Descricao
                txtCodAgencia.Text = objNotaFiscal.VencimentosNota(i).CodigoAgenciaCliente
                txtDigitoAgencia.Text = objNotaFiscal.VencimentosNota(i).DigitoAgenciaCliente
                txtPracaAgencia.Text = objNotaFiscal.VencimentosNota(i).AgenciaCliente.Praca
                txtConta.Text = objNotaFiscal.VencimentosNota(i).ContaCliente
                txtDigitoConta.Text = objNotaFiscal.VencimentosNota(i).DigitoContaCliente
            Else
                txtCodBanco.Text = ""
                txtNomeBanco.Text = ""
                txtCodAgencia.Text = ""
                txtDigitoAgencia.Text = ""
                txtPracaAgencia.Text = ""
                txtConta.Text = ""
                txtDigitoConta.Text = ""
            End If
        End If
    End Sub

End Class