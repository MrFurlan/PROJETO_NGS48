Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class ucAutorizacaoDeEmbarque
    Inherits BaseUserControl

#Region "Inicialização"
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not IsPostBack And IsConnect Then

        End If
    End Sub

    Public Overrides Sub SetarHID(ByVal guid As String)
        HID.Value = guid.ToString
    End Sub

    Public Sub SetarLocalEntregaProduto(ByVal linhaLocalEntrega As Integer, ByVal linhaProd As Integer)
        Limpar()
        hdnLinhaProd.Value = linhaProd
        hdnLinhaLocalEntrega.Value = linhaLocalEntrega
        SessaoRecuperaEmbarque()
        lblLocalEntrega.Text = objEmbarquePedido.LocaisDeEntrega(linhaLocalEntrega).CodigoClienteEntrega & " - " & objEmbarquePedido.LocaisDeEntrega(linhaLocalEntrega).EntregaFormatado
        lblProduto.Text = objEmbarquePedido.LocaisDeEntrega(linhaLocalEntrega).Produtos(linhaProd).CodigoProduto & " - " & objEmbarquePedido.LocaisDeEntrega(linhaLocalEntrega).Produtos(linhaProd).DescricaoProduto

        '*****************************************************PREENCHE O RESUMO POR PEDIDO******************************************************

        txtNormal.Text = objEmbarquePedido.LocaisDeEntrega(linhaLocalEntrega).Produtos.QtdeNormalPrd.ToString("N4")
        txtComplemento.Text = objEmbarquePedido.LocaisDeEntrega(linhaLocalEntrega).Produtos.QtdeComplementoPrd.ToString("N4")
        txtEstorno.Text = objEmbarquePedido.LocaisDeEntrega(linhaLocalEntrega).Produtos.QtdeEstornoPrd.ToString("N4")

        txtMaxAutorizar.Text = objEmbarquePedido.LocaisDeEntrega(linhaLocalEntrega).Produtos.MaximoAutorizarPrd.ToString("N4")
        txtMaxEstornar.Text = objEmbarquePedido.LocaisDeEntrega(linhaLocalEntrega).Produtos.MaximoEstornoPrd.ToString("N4")

        txtEmbarcado.Text = objEmbarquePedido.LocaisDeEntrega(linhaLocalEntrega).Produtos.QtdeEmbarcadoPrd.ToString("N4")
        txtQtdeDevolvido.Text = objEmbarquePedido.LocaisDeEntrega(linhaLocalEntrega).Produtos.QtdeDevolvidoPrd.ToString("N4")
        txtQtdeAutorizado.Text = objEmbarquePedido.LocaisDeEntrega(linhaLocalEntrega).Produtos.QtdeAutorizadoPrd.ToString("N4")

        '*****************************************************PREENCHE O RESUMO POR ENTREGA*****************************************************

        txtNormalLocal.Text = objEmbarquePedido.LocaisDeEntrega(linhaLocalEntrega).Produtos.QtdeNormalLE.ToString("N4")
        txtComplementoLocal.Text = objEmbarquePedido.LocaisDeEntrega(linhaLocalEntrega).Produtos.QtdeComplementoLE.ToString("N4")
        txtEstornoLocal.Text = objEmbarquePedido.LocaisDeEntrega(linhaLocalEntrega).Produtos.QtdeEstornoLE.ToString("N4")

        txtMaxExtornarLocal.Text = objEmbarquePedido.LocaisDeEntrega(linhaLocalEntrega).Produtos.MaximoEstornoLE.ToString("N4")

        txtEmbarcadoLocal.Text = objEmbarquePedido.LocaisDeEntrega(linhaLocalEntrega).Produtos.QtdeEmbarcadoLE.ToString("N4")
        txtQtdeDevolvidaLocal.Text = objEmbarquePedido.LocaisDeEntrega(linhaLocalEntrega).Produtos.QtdeDevolvidoLE.ToString("N4")
        txtQtdeAutorizadoLocal.Text = objEmbarquePedido.LocaisDeEntrega(linhaLocalEntrega).Produtos.QtdeAutorizadoLE.ToString("N4")

        BindGridView()
    End Sub
#End Region

#Region "Variaveis Locais"
    Private objEmbarquePedido As [Lib].Negocio.EmbarquePedido
#End Region

#Region "Session"
    Private Sub SessaoSalvarEmbarque()
        Session("ssEmbarquePedido" & HID.Value) = objEmbarquePedido
    End Sub

    Private Sub SessaoRecuperaEmbarque()
        objEmbarquePedido = CType(Session("ssEmbarquePedido" & HID.Value), [Lib].Negocio.EmbarquePedido)
    End Sub
#End Region

#Region "Methods"
    Private Sub PreencheComboTipoLancamento()
        cmbTipoLancamentoAut.Items.Clear()
        If grd Is Nothing OrElse grd.Rows.Count = 0 Then
            cmbTipoLancamentoAut.Items.Add(New ListItem("Normal", "N"))
        Else
            cmbTipoLancamentoAut.Items.Add(New ListItem("Complemento", "C"))
            cmbTipoLancamentoAut.Items.Add(New ListItem("Estorno", "E"))
        End If
    End Sub

    Public Overrides Sub Limpar()
        txtMovimentoItemAut.Text = Format(DateTime.Now, "dd/MM/yyyy")
        cmbTipoLancamentoAut.Items.Clear()
        txtQuantidadeAut.Text = String.Empty
    End Sub

    Public Sub BindGridView()
        SessaoRecuperaEmbarque()
        grd.DataSource = objEmbarquePedido.LocaisDeEntrega(hdnLinhaLocalEntrega.Value).Produtos(hdnLinhaProd.Value).Lancamentos.ToArray
        grd.DataBind()
        PreencheComboTipoLancamento()
    End Sub

    Protected Overrides Sub Selecionar()
        Popup.CloseDialog(Me.Page, "divAutorizacaoDeEmbarque")
    End Sub

    Private Function ValidarAutorizacao() As Boolean
        If String.IsNullOrWhiteSpace(txtMovimentoItemAut.Text) Then
            MsgBox(Me.Page, "É necessário informar a data de movimento!")
            txtMovimentoItemAut.Focus()
            Return False
        ElseIf Not IsDate(txtMovimentoItemAut.Text) Then
            MsgBox(Me.Page, "Data de movimento inválida!")
            txtMovimentoItemAut.Focus()
            Return False
        ElseIf CDate(txtMovimentoItemAut.Text).Date < DateTime.Now.Date Then
            MsgBox(Me.Page, "Data de movimento não pode ser menor que a data atual!")
            txtMovimentoItemAut.Focus()
            Return False
        ElseIf String.IsNullOrWhiteSpace(txtQuantidadeAut.Text) OrElse Not CDec(txtQuantidadeAut.Text) > 0 Then
            MsgBox(Me.Page, "É necessário informar a quantidade!")
            txtQuantidadeAut.Focus()
            Return False
        ElseIf cmbTipoLancamentoAut.SelectedValue = "C" OrElse cmbTipoLancamentoAut.SelectedValue = "N" Then
            If CDec(txtQuantidadeAut.Text) > CDec(txtMaxAutorizar.Text) Then
                SessaoRecuperaEmbarque()
                If objEmbarquePedido.Pedido.SubOperacao.QuantidadePedido Then
                    MsgBox(Me.Page, "Quantidade Informada é maior que a autorização prevista.")
                    Return False
                End If
            ElseIf txtMaxAutorizar.Text = 0 Then
                MsgBox(Me.Page, "Quantidade autorizada já atingiu o limite previsto.")
                Return False
            End If
        ElseIf cmbTipoLancamentoAut.SelectedValue = "E" Then
            If CDec(txtQuantidadeAut.Text) > CDec(txtMaxExtornarLocal.Text) Then
                MsgBox(Me.Page, "Quantidade Informada é maior que o quantidade de estorno válido.")
                Return False
            ElseIf CDec(txtMaxExtornarLocal.Text) = 0 Then
                MsgBox(Me.Page, "Quantidade de estorno já atingiu o limite cadastrado.")
                Return False
            End If
        End If
        Return True
    End Function
#End Region

#Region "Eventos"
    Protected Sub lnkNovoAut_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkNovoAut.Click
        If ValidarAutorizacao() Then
            SessaoRecuperaEmbarque()

            Dim objLancamento As New LancamentoProdutoEmbarque(objEmbarquePedido.LocaisDeEntrega(hdnLinhaLocalEntrega.Value).Produtos(hdnLinhaProd.Value))
            objLancamento.IUD = "I"
            If objEmbarquePedido.LocaisDeEntrega(hdnLinhaLocalEntrega.Value).Produtos(hdnLinhaProd.Value).Lancamentos IsNot Nothing AndAlso objEmbarquePedido.LocaisDeEntrega(hdnLinhaLocalEntrega.Value).Produtos(hdnLinhaProd.Value).Lancamentos.Count > 0 Then
                objLancamento.NrLancamento = objEmbarquePedido.LocaisDeEntrega(hdnLinhaLocalEntrega.Value).Produtos(hdnLinhaProd.Value).Lancamentos.Max(Function(s) s.NrLancamento) + 1
            Else
                objLancamento.NrLancamento = 1
            End If
            objLancamento.Movimento = CDate(txtMovimentoItemAut.Text)
            objLancamento.TipoDeLancamento = cmbTipoLancamentoAut.SelectedValue
            objLancamento.Quantidade = CDec(txtQuantidadeAut.Text)
            objLancamento.UsuarioInclusao = ActiveUser

            Dim Sqls As New ArrayList
            objLancamento.SalvarSql(Sqls)

            If Not Banco.GravaBanco(Sqls) Then
                MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                Exit Sub
            End If

            objEmbarquePedido.LocaisDeEntrega(hdnLinhaLocalEntrega.Value).Produtos(hdnLinhaProd.Value).Lancamentos.Add(objLancamento)

            BindGridView()
            MsgBox(Me.Page, "Autorização de embarque incluído com Sucesso.", eTitulo.Sucess)
            SessaoSalvarEmbarque()

            Popup.CloseDialog(Me.Page, "divAutorizacaoDeEmbarque")
            If TypeOf Me.Page Is AutorizacaoDeEmbarque Then
                CType(Me.Page, AutorizacaoDeEmbarque).CarregarResumoEmbarque()
            End If
        End If
    End Sub

    Protected Sub lnkLimparAut_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkLimparAut.Click
        Limpar()
    End Sub

    Protected Sub lnkFechar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkFechar.Click
        Popup.CloseDialog(Me.Page, "divAutorizacaoDeEmbarque")
    End Sub

    Protected Sub grd_RowDataBound(sender As Object, e As GridViewRowEventArgs) Handles grd.RowDataBound
        If e.Row.RowType = DataControlRowType.DataRow Then
            If e.Row.DataItem IsNot Nothing Then
                If e.Row.Cells(1).Text = "C" Then
                    e.Row.Cells(1).Text = "Complemento"
                ElseIf e.Row.Cells(1).Text = "N" Then
                    e.Row.Cells(1).Text = "Normal"
                ElseIf e.Row.Cells(1).Text = "E" Then
                    e.Row.Cells(1).Text = "Estorno"
                End If
            End If
        End If
    End Sub
#End Region

End Class