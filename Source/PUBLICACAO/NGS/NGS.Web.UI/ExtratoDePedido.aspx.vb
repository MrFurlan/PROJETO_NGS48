Imports System.Data
Imports System.IO
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class ExtratoDePedido
    Inherits BasePage

#Region "Estruturas"

    Private Structure structTotaisQuantidade
        Public QuantidadeContratada As Decimal
        Public QuantidadeFixada As Decimal
        Public QuantidadeEntregue As Decimal
        Public QuantidadeEntregueFisico As Decimal
        Public QuantidadeDevolvida As Decimal
        Public QuantidadeDevolvidaFisico As Decimal
        Public QuantidadeCessionario As Decimal
        Public QuantidadeCedente As Decimal
        Public ValorFixado As Decimal
        Public ValorPago As Decimal
        Public ValorNotaFiscal As Decimal
        Public tClasse As eClassesOperacoes
    End Structure

#End Region

#Region "Variáveis locais"

    Private objPedidos As New [Lib].Negocio.Pedidos()
    Private objQtdePedido As New structTotaisQuantidade()
    Private dateEmissao As DateTime = DateTime.Now
    Private strEmpresaQry, strClienteQry, strPedido, strGrupoProduto, strProduto, strNomeProduto, strDesPrd, strDataFim, strSafra As String
    Private blnEntrada As Boolean = False, blnSaida As Boolean = False
    Private intPagina As Integer = 0
    Private strNomePainel As String
    Private intLinhas As Integer = 0
    Private ProdutoAgrupado As Boolean = False

#End Region

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Comercial)
        If Not IsPostBack And IsConnect Then
            If Funcoes.VerificaPermissao("ExtratoDePedido", "ACESSAR") Then
                BuscaEmpresa()
                BuscarGruposProdutos()
                CarregarSafras()
                Limpar()
            Else
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Comercial.aspx")
                Exit Sub
            End If
        End If
    End Sub

    Public Overrides Sub Carregar(ByVal obj As [Lib].Negocio.IBaseEntity)
        Try
            If Session("objClienteExtrato" & HID.Value) IsNot Nothing Then
                Dim itemCliente As ListItem = Funcoes.FormatarListItemCliente(CType(Session("objClienteExtrato" & HID.Value), [Lib].Negocio.Cliente))
                txtClientes.Text = itemCliente.Text
                txtCodigoCliente.Value = itemCliente.Value
                Session.Remove("objClienteExtrato" & HID.Value)
            ElseIf Session("objPedidoExtrato" & HID.Value) IsNot Nothing Then
                Dim p As [Lib].Negocio.Pedido = CType(Session("objPedidoExtrato" & HID.Value), [Lib].Negocio.Pedido)
                txtPedido.Text = p.Codigo
                Session.Remove("objPedidoExtrato" & HID.Value)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub BuscaEmpresa()
        ddl.Carregar(cmbEmpresa, CarregarDDL.Tabela.ClientesXEmpresas, "", True)
    End Sub

    Private Sub BuscarGruposProdutos()
        ddl.Carregar(cmbGrupoProduto, CarregarDDL.Tabela.GrupoProduto, "", True)
    End Sub

    Private Sub BuscarProdutos()
        ddl.Carregar(cmbProduto, CarregarDDL.Tabela.Produto, " Grupo = '" & cmbGrupoProduto.SelectedValue & "'", True)
    End Sub

    Private Sub CarregarSafras()
        ddl.Carregar(ddlSafra, CarregarDDL.Tabela.Safra, "", True)
    End Sub

    Private Function ValidarCampos() As Boolean
        If cmbEmpresa.SelectedIndex = 0 Then
            MsgBox(Me.Page, "Empresa não foi selecionada.")
            Return False
        ElseIf txtClientes.Text.Length = 0 AndAlso txtPedido.Text.Length = 0 Then
            MsgBox(Me.Page, "Cliente não foi selecionado.")
            Return False
        ElseIf Not String.IsNullOrWhiteSpace(txtPedido.Text) Then
            If txtPedido.Text.Contains(",") OrElse txtPedido.Text.Contains(";") Then
                txtPedido.Text = txtPedido.Text.Replace(";", ",")
                For Each num As String In txtPedido.Text.Split(",")
                    If Not IsNumeric(num) Then
                        MsgBox(Me.Page, "Numero(s) de pedido inválido(s)")
                        Return False
                    End If
                Next
            Else
                If Not IsNumeric(txtPedido.Text) Then
                    MsgBox(Me.Page, "Numero de pedido inválido")
                    Return False
                End If
            End If
        ElseIf cmbGrupoProduto.SelectedIndex > 0 AndAlso cmbProduto.SelectedIndex = 0 Then
            MsgBox(Me.Page, "Produto não foi selecionado.")
            Return False
        End If

        Return True
    End Function

    Protected Sub cmdBuscaCliente_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdBuscaCliente.Click
        Try
            Dim strJScript As String = ""
            HttpContext.Current.Session("ssCampo") = "Livre"
            ucConsultaClientes.Limpar()
            Popup.ConsultaDeClientes(Me.Page, "objClienteExtrato" & HID.Value)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub cmbGrupoProduto_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbGrupoProduto.SelectedIndexChanged
        Try
            BuscarProdutos()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub Limpar()
        txtClientes.Text = ""
        txtCodigoCliente.Value = ""
        cmbGrupoProduto.SelectedIndex = 0
        ddlSafra.SelectedIndex = 0
        cmbProduto.Items.Clear()
        txtDataFinal.Text = DateTime.Now.ToString("dd/MM/yyyy")
        txtPedido.Text = ""
        chkEntrada.Checked = True
        chkSaida.Checked = True
        chkNomeProduto.Checked = False
        rdPdf.Checked = True
        Funcoes.VerificaEmpresa(cmbEmpresa)
        HID.Value = Guid.NewGuid().ToString
        ucConsultaClientes.SetarHID(HID.Value)
        ucConsultaPedidos.SetarHID(HID.Value)

        LiberaEmpresa()
    End Sub

    Private Sub LiberaEmpresa()
        If Not UsuarioServidor.LiberaEmpresa Then
            cmbEmpresa.Enabled = False
        End If
    End Sub

    Protected Sub cmdBuscaPedido_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdBuscaPedido.Click
        Try
            If String.IsNullOrWhiteSpace(cmbEmpresa.SelectedValue) OrElse String.IsNullOrWhiteSpace(txtCodigoCliente.Value) Then
                MsgBox(Me.Page, "É necessário informar a empresa e o cliente para buscar o número do pedido.")
                Exit Sub
            End If

            Dim strEmpresa As String() = cmbEmpresa.SelectedValue.Split("-")
            Dim strCliente As String() = txtCodigoCliente.Value.Split("-")
            HttpContext.Current.Session("ssCampo" & HID.Value) = "ApenasNumeroPedido"
            HttpContext.Current.Session("ssCnpjDaEmpresa") = strEmpresa(0)
            HttpContext.Current.Session("ssEndDaEmpresa") = strEmpresa(1)
            HttpContext.Current.Session("txtCnpjDoCliente") = strCliente(0)
            HttpContext.Current.Session("txtEndDoCliente") = strCliente(1)
            If ddlSafra.SelectedIndex > 0 Then HttpContext.Current.Session("codigoSafra") = ddlSafra.SelectedValue

            Dim parameters As New Dictionary(Of String, Object)
            parameters("unidade") = ""
            parameters("empresa") = strEmpresa(0)
            parameters("enderecoEmpresa") = strEmpresa(1)
            parameters("cliente") = IIf(strCliente(0) <> "", strCliente(0), "")
            parameters("enderecoCliente") = IIf(strCliente.Length > 1, strCliente(1), "")

            Popup.ConsultaDePedidos(Me.Page, "objPedidoExtrato" & HID.Value, "txtNome")
            ucConsultaPedidos.BindGridView(parameters)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub cmdBuscaPedidoEfetivo_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdBuscaPedidoEfetivo.Click
        Try
            Dim parameters As New Dictionary(Of String, Object)
            Dim strEmpresa As String() = cmbEmpresa.SelectedValue.Split("-")
            HttpContext.Current.Session("ssCampo" & HID.Value) = "ApenasNumeroPedido"
            HttpContext.Current.Session("ssCnpjDaEmpresa") = strEmpresa(0)
            HttpContext.Current.Session("ssEndDaEmpresa") = strEmpresa(1)

            parameters("empresa") = strEmpresa(0)
            parameters("enderecoEmpresa") = strEmpresa(1)
            parameters("pedidoefetivo") = txtPedidoEfetivo.Text

            Popup.ConsultaDePedidos(Me.Page, "objPedidoExtrato" & HID.Value, "txtNome")
            ucConsultaPedidos.BindGridView(parameters)
        Catch ex As Exception

        End Try
    End Sub

    Protected Sub lnkConsultar_Click(sender As Object, e As EventArgs) Handles lnkConsultar.Click
        Try
            If Funcoes.VerificaPermissao("ExtratoDePedido", "RELATORIO") Then
                If ValidarCampos() Then
                    Dim CodCliente As String = ""
                    Dim EndCliente As Integer

                    If Not String.IsNullOrWhiteSpace(txtCodigoCliente.Value) Then
                        CodCliente = txtCodigoCliente.Value.Split("-")(0)
                        EndCliente = txtCodigoCliente.Value.Split("-")(1)
                    End If

                    If rdPdf.Checked Then

                        Dim SituacaoPedido As String = String.Empty

                        If rdAberto.Checked Then
                            SituacaoPedido = "A"
                        ElseIf rdFechado.Checked Then
                            SituacaoPedido = "F"
                        Else
                            SituacaoPedido = "T"
                        End If

                        Extrato.Emitir(Me.Page, FinanceiroNovo, cmbEmpresa.SelectedValue.Split("-")(0), cmbEmpresa.SelectedValue.Split("-")(1), SituacaoPedido, txtPedido.Text, CodCliente,
                                       EndCliente, ddlSafra.SelectedValue, IIf(Not String.IsNullOrWhiteSpace(txtDataFinal.Text) AndAlso IsDate(txtDataFinal.Text), CDate(txtDataFinal.Text), Nothing),
                                       cmbProduto.SelectedValue, chkEntrada.Checked, chkSaida.Checked, chkNomeProduto.Checked, chkOcultarFinanceiro.Checked, chkOcultarLancamentoContabil.Checked, chkSintetico.Checked, chkOcultarFrete.Checked, chkPesagem.Checked)
                    Else
                        Dim strQueryString As String = "?fim=" & txtDataFinal.Text

                        If cmbEmpresa.SelectedIndex > -1 Then strQueryString &= "&empresa=" & cmbEmpresa.SelectedValue
                        If txtClientes.Text.Length > 0 Then strQueryString &= "&cliente=" & txtCodigoCliente.Value.Replace(";", "-")
                        If txtPedido.Text.Length > 0 Then strQueryString &= "&pedido=" & txtPedido.Text
                        If cmbGrupoProduto.SelectedIndex > 0 Then strQueryString &= "&grupoproduto=" & cmbGrupoProduto.SelectedValue
                        If cmbProduto.SelectedIndex > 0 Then
                            strQueryString &= "&produto=" & cmbProduto.SelectedValue & _
                                              "&nomeproduto=" & cmbProduto.SelectedItem.Text.Substring(0, cmbProduto.SelectedItem.Text.IndexOf(".."))
                        End If

                        strQueryString &= "&desprd=" & IIf(chkNomeProduto.Checked, "S", "N")

                        If ddlSafra.SelectedIndex > 0 Then strQueryString &= "&safra=" & ddlSafra.SelectedValue

                        If chkEntrada.Checked And Not chkSaida.Checked Then
                            strQueryString &= "&es=E"
                        ElseIf chkSaida.Checked And Not chkEntrada.Checked Then
                            strQueryString &= "&es=S"
                        Else
                            strQueryString &= "&es=ES"
                        End If

                        If rdAberto.Checked Then
                            strQueryString = "&tipo=A"
                        ElseIf rdFechado.Checked Then
                            strQueryString = "&tipo=F"
                        Else
                            strQueryString = "&tipo=T"
                        End If


                        ScriptManager.RegisterClientScriptBlock(Me.Page, Me.Page.GetType, Guid.NewGuid().ToString, "window.open(""ExtratoPedidoHTML.aspx" & strQueryString & """, ""relatorio"", ""status=no, toolbar=yes, location=no, menubar=yes, resizable=yes, height=600, width=800, scrollbars=yes"");", True)
                    End If
                End If
                Else
                    MsgBox(Me.Page, "Usuário sem permissão para emitir relatório.", eTitulo.Info)
                End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkLimpar_Click(sender As Object, e As EventArgs) Handles lnkLimpar.Click
        Try
            Limpar()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "ExtratoDePedido")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub

End Class