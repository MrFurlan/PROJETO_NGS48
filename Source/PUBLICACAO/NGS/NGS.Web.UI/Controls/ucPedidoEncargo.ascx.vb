Imports System.Data
Imports System.Collections.Generic
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class ucPedidoEncargo
    Inherits BaseUserControl

#Region "Inicializacao"
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

    End Sub

    Public Overrides Sub SetarHID(ByVal guid As String)
        HID.Value = guid.ToString
    End Sub

    Public Sub SetarParametros(LinhaProdutoPedido As Integer)
        HIDLinhaProduto.Value = LinhaProdutoPedido

        ddlProdutoEncargo.Items.Clear()
        For Each item In objPedido.Itens
            ddlProdutoEncargo.Items.Add(New ListItem(item.CodigoProduto & " - " & item.Descricao, item.CodigoProduto))
        Next

        objPedido.Itens(HIDLinhaProduto.Value).AjustarValorOficial = chkAjustarReais.Checked

        'Armazena a lista de encargos originais do item para que possam ser revertidos caso seja usado o limpar
        ListaDeEncargosOriginais = New [Lib].Negocio.ListPedidoXEncargo(objPedido.Itens(HIDLinhaProduto.Value))

        txtDataEncargos.Text = objPedido.Itens(HIDLinhaProduto.Value).Lancamentos.LancamentoNormal.Movimento
        If objPedido.IUD = "U" Then
            txtDataEncargos.Enabled = True
            lnkRecarregarEncargos.Visible = True
        Else
            txtDataEncargos.Enabled = False
            lnkRecarregarEncargos.Visible = False
        End If

        If objPedido.Moeda.Classificacao = [Lib].Negocio.eTiposMoeda.MoedaEstrangeira Then
            chkAjustarReais.Visible = True
        Else
            chkAjustarReais.Visible = False
        End If

        ddlProdutoEncargo.SelectedValue = objPedido.Itens(LinhaProdutoPedido).CodigoProduto
        lblConfigOperacao.Text = objPedido.Itens(LinhaProdutoPedido).CodigoOperacaoXEstado
        GridEncargos.DataSource = objPedido.Itens(LinhaProdutoPedido).Encargos.ToArray
        GridEncargos.DataBind()
    End Sub
#End Region

#Region "Session"
    Private Property objPedido() As [Lib].Negocio.Pedido
        Get
            If CType(Session("objPedido" & HID.Value), [Lib].Negocio.Pedido) Is Nothing Then
                Return New [Lib].Negocio.Pedido
            Else
                Return CType(Session("objPedido" & HID.Value), [Lib].Negocio.Pedido)
            End If
        End Get
        Set(ByVal value As [Lib].Negocio.Pedido)
            Session("objPedido" & HID.Value) = value
        End Set
    End Property

    Private Property ListaDeEncargosOriginais
        Get
            Return CType(Session("ListaDeEncargosOriginais" & HID.Value), [Lib].Negocio.ListPedidoXEncargo)
        End Get
        Set(value)
            Session("ListaDeEncargosOriginais" & HID.Value) = value
        End Set
    End Property
#End Region

    Public Overrides Sub Limpar()
        If Not String.IsNullOrWhiteSpace(HIDLinhaProduto.Value) AndAlso HIDLinhaProduto.Value > -1 Then
            ddlProdutoEncargo.SelectedValue = objPedido.Itens(HIDLinhaProduto.Value).CodigoProduto
            lblConfigOperacao.Text = objPedido.Itens(HIDLinhaProduto.Value).CodigoOperacaoXEstado
            objPedido.Itens(HIDLinhaProduto.Value).Encargos = ListaDeEncargosOriginais
            chkAjustarReais.Checked = False
            objPedido.Itens(HIDLinhaProduto.Value).AjustarValorOficial = chkAjustarReais.Checked
            GridEncargos.DataSource = objPedido.Itens(HIDLinhaProduto.Value).Encargos.ToArray
            GridEncargos.DataBind()
        End If
    End Sub

    Protected Sub ddlProdutoEncargo_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddlProdutoEncargo.SelectedIndexChanged
        Dim i As Integer = objPedido.Itens.FindIndex(Function(s) s.CodigoProduto = ddlProdutoEncargo.SelectedValue)
        HIDLinhaProduto.Value = i
        lblConfigOperacao.Text = objPedido.Itens(i).CodigoOperacaoXEstado
        GridEncargos.DataSource = objPedido.Itens(i).Encargos.ToArray
        GridEncargos.DataBind()
    End Sub

    Protected Sub GridEncargos_RowDataBound(sender As Object, e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridEncargos.RowDataBound
        If e.Row.RowType = DataControlRowType.DataRow Then
            Dim btn As Button = CType(e.Row.FindControl("btnAlterarEncargo"), Button)
            If objPedido.Itens(HIDLinhaProduto.Value).Encargos(e.Row.RowIndex).CodigoEncargo = "LIQUIDO" Or objPedido.Itens(HIDLinhaProduto.Value).Encargos(e.Row.RowIndex).CodigoEncargo = "PRODUTO" Then
                btn.Visible = False
            Else
                btn.Visible = True
            End If
        End If
    End Sub

    Protected Sub btnAlterarEncargo_Click(sender As Object, e As EventArgs)
        'SessaoRecuperaPedido()
        Dim btn As Button = CType(sender, Button)
        Dim txt As TextBox

        btn.Visible = False
        Dim linha As Integer = CType(btn.NamingContainer, GridViewRow).RowIndex

        btn = CType(GridEncargos.Rows(linha).FindControl("btnSalvarEncargo"), Button)
        btn.Visible = True

        btn = CType(GridEncargos.Rows(linha).FindControl("btnCancelarEncargo"), Button)
        btn.Visible = True


        Dim NomeEnc As String = objPedido.Itens(HIDLinhaProduto.Value).Encargos(linha).CodigoEncargo
        If NomeEnc = "FRETES" Or NomeEnc = "DESCONTOS" Or NomeEnc.Contains("FUNRURAL") Or NomeEnc.Contains("SENAR") Or NomeEnc.Contains("FACS") Or NomeEnc.Contains("FETHAB") Or NomeEnc.Contains("FUNDERSUL") Or NomeEnc.Contains("FUNDEMS") Or NomeEnc.Contains("ADUANEIRAS") Then
            txt = CType(GridEncargos.Rows(linha).FindControl("txtValor"), TextBox)
            txt.Enabled = True
        ElseIf NomeEnc.Contains("ICMS") Or NomeEnc.Contains("IPI") Or NomeEnc.Contains("PIS") Or NomeEnc.Contains("COFINS") Then
            txt = CType(GridEncargos.Rows(linha).FindControl("txtBase"), TextBox)
            txt.Enabled = True

            'txt = CType(GridEncargos.Rows(linha).FindControl("txtPercentual"), TextBox)
            'txt.Enabled = True

            txt = CType(GridEncargos.Rows(linha).FindControl("txtValor"), TextBox)
            txt.Enabled = True
        Else
            txt = CType(GridEncargos.Rows(linha).FindControl("txtValor"), TextBox)
            txt.Enabled = True
        End If
    End Sub

    Protected Sub btnCancelarEncargo_Click(sender As Object, e As EventArgs)
        'SessaoRecuperaPedido()
        lblConfigOperacao.Text = objPedido.Itens(HIDLinhaProduto.Value).CodigoOperacaoXEstado
        GridEncargos.DataSource = objPedido.Itens(HIDLinhaProduto.Value).Encargos.ToArray
        GridEncargos.DataBind()
    End Sub

    Protected Sub btnSalvarEncargo_Click(sender As Object, e As EventArgs)
        'SessaoRecuperaPedido()
        Dim btn As Button = CType(sender, Button)
        btn.Visible = False

        Dim linha As Integer = CType(btn.NamingContainer, GridViewRow).RowIndex

        btn = CType(GridEncargos.Rows(linha).FindControl("btnCancelarEncargo"), Button)
        btn.Visible = False

        btn = CType(GridEncargos.Rows(linha).FindControl("btnAlterarEncargo"), Button)
        btn.Visible = True

        '************************  VALOR  ************************
        Dim txt As TextBox
        txt = CType(GridEncargos.Rows(linha).FindControl("txtValor"), TextBox)
        txt.Enabled = False

        If objPedido.Moeda.Classificacao = [Lib].Negocio.eTiposMoeda.Oficial OrElse objPedido.Itens(HIDLinhaProduto.Value).AjustarValorOficial Then
            objPedido.Itens(HIDLinhaProduto.Value).Encargos(linha).ValorOficial = Math.Round(CDec(txt.Text), 2)
            objPedido.Itens(HIDLinhaProduto.Value).Encargos(linha).ValorMoeda = Funcoes.ConverteMoeda(objPedido.Itens(HIDLinhaProduto.Value).Encargos(linha).ValorOficial, objPedido.IndiceFixado, eTiposMoeda.MoedaEstrangeira, True, False, 2)
        Else
            objPedido.Itens(HIDLinhaProduto.Value).Encargos(linha).ValorMoeda = Math.Round(CDec(txt.Text), 2)
            objPedido.Itens(HIDLinhaProduto.Value).Encargos(linha).ValorOficial = Funcoes.ConverteMoeda(objPedido.Itens(HIDLinhaProduto.Value).Encargos(linha).ValorMoeda, objPedido.IndiceFixado, eTiposMoeda.Oficial, True, False, 2)
        End If


        '************************  VALOR  ************************
        txt = CType(GridEncargos.Rows(linha).FindControl("txtBase"), TextBox)
        txt.Enabled = False

        If objPedido.Moeda.Classificacao = [Lib].Negocio.eTiposMoeda.Oficial OrElse objPedido.Itens(HIDLinhaProduto.Value).AjustarValorOficial Then
            objPedido.Itens(HIDLinhaProduto.Value).Encargos(linha).BaseOficial = Math.Round(CDec(txt.Text), 2)
            objPedido.Itens(HIDLinhaProduto.Value).Encargos(linha).BaseMoeda = Funcoes.ConverteMoeda(objPedido.Itens(HIDLinhaProduto.Value).Encargos(linha).BaseOficial, objPedido.IndiceFixado, eTiposMoeda.MoedaEstrangeira, True, False, 2)
        Else
            objPedido.Itens(HIDLinhaProduto.Value).Encargos(linha).BaseMoeda = Math.Round(CDec(txt.Text), 2)
            objPedido.Itens(HIDLinhaProduto.Value).Encargos(linha).BaseOficial = Funcoes.ConverteMoeda(objPedido.Itens(HIDLinhaProduto.Value).Encargos(linha).BaseMoeda, objPedido.IndiceFixado, eTiposMoeda.Oficial, True, False, 2)
        End If

        objPedido.Itens(HIDLinhaProduto.Value).Encargos.AtualizaLiquido()
        objPedido.Itens(HIDLinhaProduto.Value).IUD = "U"

        'SessaoSalvaPedido()
        lblConfigOperacao.Text = objPedido.Itens(HIDLinhaProduto.Value).CodigoOperacaoXEstado
        GridEncargos.DataSource = objPedido.Itens(HIDLinhaProduto.Value).Encargos.ToArray
        GridEncargos.DataBind()
    End Sub

    Protected Sub chkAjustarReais_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        SetarParametros(HIDLinhaProduto.Value)
    End Sub

    Protected Sub lnkRecarregarEncargos_Click(sender As Object, e As EventArgs) Handles lnkRecarregarEncargos.Click

        If Funcoes.VerificaPermissao("PedidosXItens", "ALTERAR") Then
            If objPedido.IUD = "U" Then
                If Not objPedido.FiscalAberto Then
                    MsgBox(Me.Page, "Pedido com Fiscal Fechado não pode ajustado.")
                    Exit Sub
                ElseIf Not objPedido.FinanceiroAberto Then
                    MsgBox(Me.Page, "Pedido com Finaneiro Fechado não pode ser ajustado.")
                    Exit Sub
                End If

                Dim dataAnteriorItem As Date = objPedido.Itens(HIDLinhaProduto.Value).Lancamentos.LancamentoNormal.Movimento
                objPedido.Itens(HIDLinhaProduto.Value).IUD = "U"
                objPedido.Itens(HIDLinhaProduto.Value).Lancamentos.LancamentoNormal.Movimento = CDate(txtDataEncargos.Text).ToShortDateString()
                objPedido.Itens(HIDLinhaProduto.Value).AjustarValorOficial = chkAjustarReais.Checked

                If (Left(objPedido.CodigoEmpresa, 8) = "24450490" OrElse Left(objPedido.CodigoEmpresa, 8) = "44979506") AndAlso objPedido.Moeda.Classificacao = eTiposMoeda.MoedaEstrangeira Then
                    For Each item In objPedido.Itens
                        If item.QuantidadePedidoFaturamento > 0 Then
                            If objPedido.IndexadorFixo Then
                                For Each lcto In item.Lancamentos
                                    If lcto.QuantidadeFaturamento > 0 Then
                                        If objPedido.IndexadorFixo Then
                                            lcto.UnitarioOficial = Funcoes.ConverteMoeda(lcto.UnitarioMoeda, objPedido.IndiceFixado, eTiposMoeda.Oficial, True, False, 10)
                                            lcto.TotalOficial = Funcoes.ConverteMoeda(lcto.TotalMoeda, objPedido.IndiceFixado, eTiposMoeda.Oficial, True, False, 10)
                                        Else
                                            lcto.UnitarioOficial = Math.Round(Funcoes.ConverteParaMoedaOficial(lcto.UnitarioMoeda, objPedido.Indexador.Codigo, lcto.Movimento), 10, MidpointRounding.AwayFromZero)
                                            lcto.TotalOficial = Math.Round(Funcoes.ConverteParaMoedaOficial(lcto.TotalMoeda, objPedido.Indexador.Codigo, lcto.Movimento), 2, MidpointRounding.AwayFromZero)
                                        End If
                                    End If
                                Next
                            End If
                        End If
                    Next
                End If

                objPedido.Itens(HIDLinhaProduto.Value).Encargos = Nothing
                objPedido.Itens(HIDLinhaProduto.Value).CodigoOperacaoXEstado = 0

                objPedido.Itens(HIDLinhaProduto.Value).Encargos.CriaListar()

                objPedido.Itens(HIDLinhaProduto.Value).Encargos.AtualizaLiquido()
                objPedido.Itens(HIDLinhaProduto.Value).Lancamentos.LancamentoNormal.Movimento = dataAnteriorItem

                lblConfigOperacao.Text = objPedido.Itens(HIDLinhaProduto.Value).CodigoOperacaoXEstado
                GridEncargos.DataSource = objPedido.Itens(HIDLinhaProduto.Value).Encargos.ToArray
                GridEncargos.DataBind()
            Else
                MsgBox(Me.Page, "Usuario sem permissão para atualizar encargos")
            End If
        Else
            MsgBox(Me.Page, "Usuario sem permissão para atualizar encargos")
        End If

    End Sub

    Protected Sub lnkLimpar_Click(sender As Object, e As EventArgs) Handles lnkLimpar.Click
        Limpar()
    End Sub

    Protected Sub LnkFechar_Click(sender As Object, e As EventArgs) Handles LnkFechar.Click
        HIDLinhaProduto.Value = -1

        If TypeOf Me.Page Is PedidosXItens Then
            CType(Me.Page, PedidosXItens).AtualizarResumoItens()
        End If

        If TypeOf Me.Page Is AjustarPedido Then
            CType(Me.Page, AjustarPedido).AtualizarGridEncargos(True)
        End If

        Popup.CloseDialog(Me.Page, "divPedidoEncargo")
    End Sub
End Class