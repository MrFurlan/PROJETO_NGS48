Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class ucEncargosNFG
    Inherits BaseUserControl

#Region "Variáveis"
    Private objNotaFiscal As [Lib].Negocio.NotaFiscal
    Private dblEncargos As Double
    Private dblLiquido As Double
    Private dblTotal As Double
#End Region

#Region "Eventos"
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not IsPostBack Then
            CarregarProdutos()
        End If
    End Sub

    'Protected Sub lnkNovo_Click(sender As Object, e As EventArgs) Handles lnkNovo.Click
    '    For Each row As GridViewRow In DgEncargos.Rows
    '        AtualizarEncargos(row.RowIndex)
    '    Next

    '    If TypeOf Me.Page Is NotasFiscaisGeraisNova Then
    '        CType(Me.Page, NotasFiscaisGeraisNova).AtualizarItensNoGrid()
    '        CType(Me.Page, NotasFiscaisGeraisNova).AtualizaValorParcelamento()
    '    End If
    '    Popup.CloseDialog(Me.Page, "divEncargosNFG")
    'End Sub

    Protected Sub lnkRecarregar_Click(sender As Object, e As EventArgs) Handles lnkRecarregar.Click
        SessaoRecuperaNotaFiscal()

        If Not objNotaFiscal.Itens Is Nothing Then objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(hdfIndex.Value).IUD = "U"

        Dim fCliente As String = objNotaFiscal.CodigoCliente
        Dim fEndCliente As Integer = objNotaFiscal.EnderecoCliente

        If objNotaFiscal.CodigoTipoDeDocumento = 2 OrElse objNotaFiscal.CodigoTipoDeDocumento = 8 OrElse objNotaFiscal.CodigoTipoDeDocumento = 10 OrElse objNotaFiscal.CodigoTipoDeDocumento = 14 OrElse objNotaFiscal.CodigoTipoDeDocumento = 57 Then
            objNotaFiscal.TipoDeDocumentoFrete = eTipoDeDocumentoFrete.Comprovacao

            If Not objNotaFiscal.NotasTrocaOrigem Is Nothing AndAlso _
                objNotaFiscal.NotasTrocaOrigem.Count > 0 AndAlso _
                objNotaFiscal.Cliente.CodigoEstado <> objNotaFiscal.NotasTrocaOrigem(0).Cliente.CodigoEstado Then
                objNotaFiscal.CodigoCliente = objNotaFiscal.NotasTrocaOrigem(0).CodigoCliente
                objNotaFiscal.EnderecoCliente = objNotaFiscal.NotasTrocaOrigem(0).EnderecoCliente
            End If
        End If

        Dim nXe = New ListNotaFiscalXItemXEncargo(objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(hdfIndex.Value), New List(Of eEtapaEncago) From {eEtapaEncago.Normal})

        If nXe.Count > 0 Then
            objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(hdfIndex.Value).Encargos = Nothing
            objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(hdfIndex.Value).Encargos = nXe

            DgEncargos.DataSource = objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(hdfIndex.Value).Encargos
            DgEncargos.DataBind()
            HabilitarCampos()

            If objNotaFiscal.CodigoTipoDeDocumento = 2 OrElse objNotaFiscal.CodigoTipoDeDocumento = 8 OrElse objNotaFiscal.CodigoTipoDeDocumento = 10 OrElse objNotaFiscal.CodigoTipoDeDocumento = 14 OrElse objNotaFiscal.CodigoTipoDeDocumento = 57 Then
                objNotaFiscal.TipoDeDocumentoFrete = Nothing
                objNotaFiscal.CodigoCliente = fCliente
                objNotaFiscal.EnderecoCliente = fEndCliente
            End If

            SessaoSalvaNotaFiscal()
        Else
            MsgBox(Me.Page, "Encargos não foram encotrados/recarregados")
        End If
    End Sub

    Protected Sub lnkLimpar_Click(sender As Object, e As EventArgs) Handles lnkLimpar.Click
        Limpar()
    End Sub

    Protected Sub lnkFechar_Click(sender As Object, e As EventArgs) Handles lnkFechar.Click
        If TypeOf Me.Page Is NotasFiscaisGerais Then
            CType(Me.Page, NotasFiscaisGerais).AtualizarItensNoGrid()
            CType(Me.Page, NotasFiscaisGerais).AtualizarValorParcelamento()
        End If
        Popup.CloseDialog(Me.Page, "divEncargosNFG")
    End Sub

    Protected Sub ddlProdutoSel_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        SessaoRecuperaNotaFiscal()
        hdfIndex.Value = ddlProdutoSel.SelectedIndex
        DgEncargos.DataSource = objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(hdfIndex.Value).Encargos
        DgEncargos.DataBind()
        HabilitarCampos()
    End Sub

    Protected Sub DgEncargos_RowCreated(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs)
        If e.Row.RowType = DataControlRowType.DataRow Then
            Dim btnEncargoItem As Button = CType(e.Row.FindControl("btnEncargoItem"), Button)
            btnEncargoItem.CommandArgument = e.Row.RowIndex.ToString()
        End If
    End Sub

    Protected Sub DgEncargos_RowCommand(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewCommandEventArgs) Handles DgEncargos.RowCommand
        If e.CommandName = "OK" Then
            Dim index As Integer = Convert.ToInt32(e.CommandArgument)
            AtualizarEncargos(index)
        End If
    End Sub

#End Region

#Region "Métodos"
    Private Sub SessaoSalvaNotaFiscal()
        Session("objNotaFiscal" & HID.Value) = objNotaFiscal
    End Sub

    Private Sub SessaoRecuperaNotaFiscal()
        objNotaFiscal = CType(Session("objNotaFiscal" & HID.Value), [Lib].Negocio.NotaFiscal)
    End Sub

    Public Overrides Sub SetarHID(ByVal guid As String)
        HID.Value = guid.ToString
    End Sub

    Public Overrides Sub Limpar()
        hdfIndex.Value = ""
        'lnkNovo.Parent.Visible = False
        If ddlProdutoSel.Items.Count > 0 Then
            ddlProdutoSel.SelectedIndex = 0
            hdfIndex.Value = ddlProdutoSel.SelectedIndex
        End If
        SessaoRecuperaNotaFiscal()
        lnkRecarregar.Parent.Visible = objNotaFiscal.IUD = "U"
        DgEncargos.DataSource = objNotaFiscal.Itens.Where(Function(s) s.CodigoProduto = ddlProdutoSel.SelectedValue And s.IUD <> "D").SelectMany(Function(s) s.Encargos)
        DgEncargos.DataBind()
        HabilitarCampos()
    End Sub

    Public Sub Inicializar(ByVal posicao As Integer)
        CarregarProdutos()
        hdfIndex.Value = posicao
        SessaoRecuperaNotaFiscal()
        ddlProdutoSel.SelectedValue = objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(posicao).CodigoProduto
        DgEncargos.DataSource = objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(hdfIndex.Value).Encargos
        DgEncargos.DataBind()
        HabilitarCampos()
    End Sub

    Private Sub CarregarProdutos()
        SessaoRecuperaNotaFiscal()
        ddlProdutoSel.Items.Clear()
        If objNotaFiscal IsNot Nothing Then
            For Each item As NotaFiscalXItem In objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")
                Dim row As ListItem = New ListItem(String.Format("{0} - {1}", item.CodigoProduto, item.Produto.Nome), item.CodigoProduto)
                ddlProdutoSel.Items.Add(row)
            Next
        End If
    End Sub

    Private Sub AtualizarEncargos(ByVal index As Integer)
        SessaoRecuperaNotaFiscal()
        Dim base As Decimal = CType(DgEncargos.Rows(index).FindControl("txtBaseEncargoItem"), TextBox).Text
        Dim percentual As Decimal = CType(DgEncargos.Rows(index).FindControl("txtPercentualEncargoItem"), TextBox).Text
        Dim valor As Decimal = CType(DgEncargos.Rows(index).FindControl("txtValorEncargoItem"), TextBox).Text

        If objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(hdfIndex.Value).Encargos(index).Codigo = "FACS" OrElse _
           objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(hdfIndex.Value).Encargos(index).Codigo = "FETHAB" OrElse _
           objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(hdfIndex.Value).Encargos(index).Codigo = "IAGRO" OrElse _
           objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(hdfIndex.Value).Encargos(index).Codigo = "SENAR" OrElse _
           objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(hdfIndex.Value).Encargos(index).Codigo = "FUNRURAL JUDICIAL" OrElse _
           objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(hdfIndex.Value).Encargos(index).Codigo = "FUNRURAL" Then

            Dim vlrOriginal As Decimal = 0
            If objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(hdfIndex.Value).Encargos(index).ValorPeso = eValorPeso.Peso Then
                vlrOriginal = Math.Round(objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(hdfIndex.Value).Encargos(index).Base / 1000 * objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(hdfIndex.Value).Encargos(index).Percentual, 2, MidpointRounding.AwayFromZero)
            Else
                vlrOriginal = Math.Round(objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(hdfIndex.Value).Encargos(index).Base * objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(hdfIndex.Value).Encargos(index).Percentual / 100, 2, MidpointRounding.AwayFromZero)
            End If
            If valor > (vlrOriginal + 1) OrElse valor < (vlrOriginal - 1) Then
                valor = vlrOriginal
                CType(DgEncargos.Rows(index).FindControl("txtValorEncargoItem"), TextBox).Text = valor
                MsgBox(Me.Page, "Variação máxima do " & objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(hdfIndex.Value).Encargos(index).Codigo & " não pode ser superior/inferior à R$ 1,00. Qualquer dúvida entre em contato com o Suporte.")
                Exit Sub
            End If
        End If

        If objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(hdfIndex.Value).Encargos(index).Codigo = "FRETES" Then objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(hdfIndex.Value).Encargos(index).Encargo.Atualizacao = True

        Dim valorApuracao As Decimal = Math.Round(base * percentual / 100, 2, MidpointRounding.AwayFromZero)

        If valor > (valorApuracao + 1) OrElse _
            valor < (valorApuracao - 1) Then
            MsgBox(Me.Page, "Variação máxima do " & objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(hdfIndex.Value).Encargos(index).Codigo & " não pode ser superior/inferior à R$ 1,00. Qualquer dúvida entre em contato com o Suporte.")
            Exit Sub
        Else
            objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(hdfIndex.Value).Encargos(index).Base = base
            objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(hdfIndex.Value).Encargos(index).Percentual = percentual
            objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(hdfIndex.Value).Encargos(index).PercentualExibicao = percentual
            objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(hdfIndex.Value).Encargos(index).Valor = valor
        End If

        If objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(hdfIndex.Value).Encargos(index).ValorPeso = eValorPeso.Peso Then
            If Not objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(hdfIndex.Value).Encargos(index).Codigo = "FACS" And Not objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(hdfIndex.Value).Encargos(index).Codigo = "FETHAB" And Not objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(hdfIndex.Value).Encargos(index).Codigo = "IAGRO" Then
                objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(hdfIndex.Value).Encargos(index).Percentual = Math.Round((objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(hdfIndex.Value).Encargos(index).Valor * 1000) / objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(hdfIndex.Value).Encargos(index).Base, 6, MidpointRounding.AwayFromZero)
            End If
        Else
            If objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(hdfIndex.Value).Encargos(index).Base = 0 OrElse _
                objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(hdfIndex.Value).Encargos(index).Percentual = 0 OrElse _
                objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(hdfIndex.Value).Encargos(index).Valor = 0 Then
                objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(hdfIndex.Value).Encargos(index).Base = 0
                objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(hdfIndex.Value).Encargos(index).Percentual = 0
                objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(hdfIndex.Value).Encargos(index).Valor = 0
            Else
                '    If objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(hdfIndex.Value).Encargos(index).Percentual > 0 AndAlso Not objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(hdfIndex.Value).Encargos(index).Codigo = "FUNRURAL" Then
                '        'objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(hdfIndex.Value).Encargos(index).Percentual = Math.Round((objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(hdfIndex.Value).Encargos(index).Valor * 100) / objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(hdfIndex.Value).Encargos(index).Base, 6, MidpointRounding.AwayFromZero)
                '        objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(hdfIndex.Value).Encargos(index).Valor = Math.Round(objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(hdfIndex.Value).Encargos(index).Base * (Convert.ToDouble(objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(hdfIndex.Value).Encargos(index).Percentual) / 100), 2)
                '    End If
            End If
        End If

        If objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(hdfIndex.Value).Encargos(index).Codigo = "FRETES" Then objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(hdfIndex.Value).Encargos(index).Encargo.Atualizacao = False

        If Not objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(hdfIndex.Value).IUD = "I" Then objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(hdfIndex.Value).IUD = "U"

        DgEncargos.DataSource = objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(hdfIndex.Value).Encargos
        DgEncargos.DataBind()
        objNotaFiscal.AtualizaTotais()
        HabilitarCampos()
        SessaoSalvaNotaFiscal()
    End Sub

    Private Sub HabilitarCampos()
        Dim rowIndex As Integer
        For rowIndex = 0 To DgEncargos.Rows.Count - 1
            If DgEncargos.Rows(rowIndex).Cells(0).Text() = "LIQUIDO" OrElse _
               DgEncargos.Rows(rowIndex).Cells(0).Text().Contains("PRODUTO") Then
                CType(DgEncargos.Rows(rowIndex).FindControl("txtBaseEncargoItem"), TextBox).Enabled = False
                CType(DgEncargos.Rows(rowIndex).FindControl("txtPercentualEncargoItem"), TextBox).Enabled = False
                CType(DgEncargos.Rows(rowIndex).FindControl("txtValorEncargoItem"), TextBox).Enabled = False
                CType(DgEncargos.Rows(rowIndex).FindControl("btnEncargoItem"), Button).Enabled = False
            Else
                CType(DgEncargos.Rows(rowIndex).FindControl("txtBaseEncargoItem"), TextBox).Enabled = True
                CType(DgEncargos.Rows(rowIndex).FindControl("txtPercentualEncargoItem"), TextBox).Enabled = True
                CType(DgEncargos.Rows(rowIndex).FindControl("txtValorEncargoItem"), TextBox).Enabled = True
                CType(DgEncargos.Rows(rowIndex).FindControl("btnEncargoItem"), Button).Enabled = True
            End If
        Next
    End Sub
#End Region

End Class