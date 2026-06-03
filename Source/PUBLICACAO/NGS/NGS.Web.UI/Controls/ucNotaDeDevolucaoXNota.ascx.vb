Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class ucNotaDeDevolucaoXNota
    Inherits BaseUserControl

#Region "NotaFiscal"
    Private objNotaFiscal As [Lib].Negocio.NotaFiscal
    Private temTituloNoBanco As Boolean
    Private temTituloEndosso As Boolean
    Private valorOriginal As Decimal

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

    Public Sub BindGridView()
        CarregarLista(False)
        AtualizaQtdeVlr()
    End Sub

    Public Overrides Sub SetarHID(ByVal NewGuid As String)
        Session.Remove("ssValorOriginal" & HID.Value)

        HID.Value = NewGuid.ToString

        valorOriginal = 0
        Session("ssValorOriginal" & HID.Value) = valorOriginal
    End Sub

    Public Sub SetarIndice(ByVal pIndice As Integer)
        hdfIndice.Value = pIndice
    End Sub

    Private Sub CarregarLista(ByVal limparValores As Boolean)
        Dim i As Integer = hdfIndice.Value
        SessaoRecuperaNotaFiscal()

        If Not objNotaFiscal.SubOperacao.QuantidadeFiscal Then
            divReajusteUnitario.Style.Clear()
            txtUnitarioMedio.Text = objNotaFiscal.Pedido.Itens.Where(Function(s) s.CodigoProduto = objNotaFiscal.Itens(i).CodigoProduto).FirstOrDefault.UnitarioMedioFaturamento.ToString("N10")
        Else
            divReajusteUnitario.Style.Add("Display", "none")
            txtUnitarioMedio.Text = 0
        End If

        If objNotaFiscal.IUD = "I" Then
            For Each row As [Lib].Negocio.NotaFiscalDevolucaoXNotaFiscal In objNotaFiscal.Itens(i).NotasDevolucao

                row.ValorDevolucao = Math.Round(row.ValorDevolucao, 2, MidpointRounding.AwayFromZero)

                If row.Nota.VencimentosNota IsNot Nothing AndAlso row.Nota.VencimentosNota.Count > 0 AndAlso row.Nota.VencimentosNota.Any(Function(s) s.CodigoSituacao = eSituacao.Normal And s.CodigoProvisao = eProvisao.Baixa) Then
                    row.QuantidadeDevolucao = 0
                    row.ValorDevolucao = 0
                End If
            Next
        End If

        gridNotaDeDevolucaoXNota.DataSource = objNotaFiscal.Itens(i).NotasDevolucao
        gridNotaDeDevolucaoXNota.DataBind()

        For j As Integer = gridNotaDeDevolucaoXNota.Rows.Count - 1 To 0 Step -1

            For Each n In objNotaFiscal.Itens(i).NotasDevolucao
                If gridNotaDeDevolucaoXNota.Rows(j).Cells(0).Text = n.Nota.Codigo Then

                    CType(gridNotaDeDevolucaoXNota.Rows(j).FindControl("imgMenos"), ImageButton).Visible = False
                    CType(gridNotaDeDevolucaoXNota.Rows(j).FindControl("imgMais"), ImageButton).Visible = False

                    If Not n.Nota.ChaveNFE.Length = 44 Then

                        CType(gridNotaDeDevolucaoXNota.Rows(j).FindControl("txtQuantidade"), TextBox).Text = "0,0000"
                        CType(gridNotaDeDevolucaoXNota.Rows(j).FindControl("txtQuantidade"), TextBox).Enabled = False

                        CType(gridNotaDeDevolucaoXNota.Rows(j).FindControl("txtValor"), TextBox).Text = "0,00"
                        CType(gridNotaDeDevolucaoXNota.Rows(j).FindControl("txtValor"), TextBox).Enabled = False

                        CType(gridNotaDeDevolucaoXNota.Rows(j).FindControl("imgInfNota"), ImageButton).Visible = True
                        CType(gridNotaDeDevolucaoXNota.Rows(j).FindControl("imgInfNota"), ImageButton).ToolTip = "Nota Fiscal sem a CHAVE ELETRÔNICA não pode ser utilizada."

                    ElseIf n.Nota.SubOperacao.Financeiro AndAlso (n.Nota.VencimentosNota Is Nothing OrElse n.Nota.VencimentosNota.Count = 0) Then
                        CType(gridNotaDeDevolucaoXNota.Rows(j).FindControl("txtQuantidade"), TextBox).Text = "0,0000"
                        CType(gridNotaDeDevolucaoXNota.Rows(j).FindControl("txtQuantidade"), TextBox).Enabled = False

                        CType(gridNotaDeDevolucaoXNota.Rows(j).FindControl("txtValor"), TextBox).Text = "0,00"
                        CType(gridNotaDeDevolucaoXNota.Rows(j).FindControl("txtValor"), TextBox).Enabled = False

                        CType(gridNotaDeDevolucaoXNota.Rows(j).FindControl("imgInfNota"), ImageButton).Visible = True
                        CType(gridNotaDeDevolucaoXNota.Rows(j).FindControl("imgInfNota"), ImageButton).ToolTip = "Nota Fiscal sem Título vinculado não pode ser utilizado."

                    ElseIf n.Nota.VencimentosNota IsNot Nothing AndAlso n.Nota.VencimentosNota.Count > 0 AndAlso n.Nota.VencimentosNota.Any(Function(s) s.CodigoSituacao = eSituacao.Normal And s.CodigoProvisao = eProvisao.Baixa) Then

                        CType(gridNotaDeDevolucaoXNota.Rows(j).FindControl("txtQuantidade"), TextBox).Text = "0,0000"
                        CType(gridNotaDeDevolucaoXNota.Rows(j).FindControl("txtQuantidade"), TextBox).Enabled = False

                        CType(gridNotaDeDevolucaoXNota.Rows(j).FindControl("txtValor"), TextBox).Text = "0,00"
                        CType(gridNotaDeDevolucaoXNota.Rows(j).FindControl("txtValor"), TextBox).Enabled = False

                        CType(gridNotaDeDevolucaoXNota.Rows(j).FindControl("imgInfNota"), ImageButton).Visible = True
                        CType(gridNotaDeDevolucaoXNota.Rows(j).FindControl("imgInfNota"), ImageButton).ToolTip = "Nota Fiscal com Título baixado não pode ser utilizada."

                    ElseIf n.Nota.VencimentosNota IsNot Nothing AndAlso n.Nota.VencimentosNota.Count > 0 AndAlso n.Nota.VencimentosNota.Any(Function(s) s.CodigoSituacao = eSituacao.RemessaBancaria) Then
                        CType(gridNotaDeDevolucaoXNota.Rows(j).FindControl("txtQuantidade"), TextBox).Text = "0,0000"
                        CType(gridNotaDeDevolucaoXNota.Rows(j).FindControl("txtQuantidade"), TextBox).Enabled = False

                        CType(gridNotaDeDevolucaoXNota.Rows(j).FindControl("txtValor"), TextBox).Text = "0,00"
                        CType(gridNotaDeDevolucaoXNota.Rows(j).FindControl("txtValor"), TextBox).Enabled = False

                        CType(gridNotaDeDevolucaoXNota.Rows(j).FindControl("imgInfNota"), ImageButton).Visible = True
                        CType(gridNotaDeDevolucaoXNota.Rows(j).FindControl("imgInfNota"), ImageButton).ToolTip = "Nota Fiscal com Título em Cobrança Bancária não pode ser utilizada."

                    ElseIf n.Nota.VencimentosNota IsNot Nothing AndAlso n.Nota.VencimentosNota.Count > 0 AndAlso n.Nota.VencimentosNota.Any(Function(s) s.CodigoSituacao = eSituacao.EndossoTitulo) Then
                        CType(gridNotaDeDevolucaoXNota.Rows(j).FindControl("txtQuantidade"), TextBox).Text = "0,0000"
                        CType(gridNotaDeDevolucaoXNota.Rows(j).FindControl("txtQuantidade"), TextBox).Enabled = False

                        CType(gridNotaDeDevolucaoXNota.Rows(j).FindControl("txtValor"), TextBox).Text = "0,00"
                        CType(gridNotaDeDevolucaoXNota.Rows(j).FindControl("txtValor"), TextBox).Enabled = False

                        CType(gridNotaDeDevolucaoXNota.Rows(j).FindControl("imgInfNota"), ImageButton).Visible = True
                        CType(gridNotaDeDevolucaoXNota.Rows(j).FindControl("imgInfNota"), ImageButton).ToolTip = "Nota Fiscal com Título de Endosso não pode ser utilizada."

                    ElseIf n.Nota.VencimentosNota IsNot Nothing AndAlso n.Nota.VencimentosNota.Count > 0 AndAlso n.Nota.VencimentosNota.Any(Function(s) s.RegistroMestre > 0) Then
                        CType(gridNotaDeDevolucaoXNota.Rows(j).FindControl("txtQuantidade"), TextBox).Text = "0,0000"
                        CType(gridNotaDeDevolucaoXNota.Rows(j).FindControl("txtQuantidade"), TextBox).Enabled = False

                        CType(gridNotaDeDevolucaoXNota.Rows(j).FindControl("txtValor"), TextBox).Text = "0,00"
                        CType(gridNotaDeDevolucaoXNota.Rows(j).FindControl("txtValor"), TextBox).Enabled = False

                        CType(gridNotaDeDevolucaoXNota.Rows(j).FindControl("imgInfNota"), ImageButton).Visible = True
                        CType(gridNotaDeDevolucaoXNota.Rows(j).FindControl("imgInfNota"), ImageButton).ToolTip = "Nota Fiscal com Título em Agrupamento não pode ser utilizado."
                    ElseIf Not objNotaFiscal.NossaEmissao AndAlso gridNotaDeDevolucaoXNota.Rows.Count = 1 Then
                        CType(gridNotaDeDevolucaoXNota.Rows(j).FindControl("txtValor"), TextBox).Enabled = True

                        valorOriginal = CDec(CType(gridNotaDeDevolucaoXNota.Rows(j).FindControl("txtValor"), TextBox).Text)

                        Session("ssValorOriginal" & HID.Value) += valorOriginal

                        If Funcoes.VerificaPermissao("AjustarCentavoNotaFiscal", "ALTERAR") Then
                            CType(gridNotaDeDevolucaoXNota.Rows(j).FindControl("imgMenos"), ImageButton).Visible = True
                            CType(gridNotaDeDevolucaoXNota.Rows(j).FindControl("imgMais"), ImageButton).Visible = True
                        End If
                    Else
                        If Funcoes.VerificaPermissao("AjustarCentavoNotaFiscal", "ALTERAR") Then
                            CType(gridNotaDeDevolucaoXNota.Rows(j).FindControl("imgMenos"), ImageButton).Visible = True
                            CType(gridNotaDeDevolucaoXNota.Rows(j).FindControl("imgMais"), ImageButton).Visible = True
                        End If
                    End If

                    If Funcoes.VerificaPermissao("AjustarValorTotalNotaFiscal", "ALTERAR") Then
                        CType(gridNotaDeDevolucaoXNota.Rows(j).FindControl("txtValor"), TextBox).Enabled = True
                    End If

                End If
            Next

        Next j

        If limparValores Then
            lblQuantidade.Text = "0,0000"
            lblValor.Text = "0,00"
        Else
            lblQtdeNota.Text = objNotaFiscal.Itens(i).QuantidadeFiscal.ToString("N4")
            lblVlrNota.Text = objNotaFiscal.Itens(i).ValorTotal.ToString("N2")
        End If

        SessaoSalvaNotaFiscal()

    End Sub

    Protected Sub txtQuantidade_TextChanged(sender As Object, e As EventArgs)
        SessaoRecuperaNotaFiscal()
        Dim txtQuantidade As TextBox = CType(sender, TextBox)
        Dim row As GridViewRow = CType(txtQuantidade.NamingContainer, GridViewRow)
        Dim i As Integer = Convert.ToInt32(hdfIndice.Value)

        If objNotaFiscal.Itens(i).Produto.CodigoEmbalagem = 1 AndAlso CDec(txtQuantidade.Text) > 0 Then
            'ainda vou testar - Furlan - 23/08/2022
            Dim vString As String = txtQuantidade.Text
            Dim vDecimal As Decimal = CDec(vString.Split(",")(1))

            If vDecimal > 0 Then
                MsgBox(Me.Page, "Produto a GRANEL não pode ter casa decimal.", eTitulo.Info)

                txtQuantidade.Text = "0,0000"

                AtualizaQtdeVlr()

                Exit Sub
            End If
        End If

        If objNotaFiscal.IUD <> "I" Then
            txtQuantidade.Text = objNotaFiscal.Itens(i).NotasDevolucao(row.RowIndex).QuantidadeDevolucao
            Exit Sub
        End If

        If txtQuantidade.Text.Length > 0 AndAlso CDec(txtQuantidade.Text) > objNotaFiscal.Itens(i).NotasDevolucao(row.RowIndex).QuantidadeSaldo Then
            txtQuantidade.Text = objNotaFiscal.Itens(i).NotasDevolucao(row.RowIndex).QuantidadeSaldo

            AtualizaQtdeVlr()

            Exit Sub
        End If

        If txtQuantidade.Text.Length > 0 AndAlso CDec(txtQuantidade.Text) > 0 Then
            objNotaFiscal.Itens(i).NotasDevolucao(row.RowIndex).QuantidadeDevolucao = txtQuantidade.Text
            objNotaFiscal.Itens(i).NotasDevolucao(row.RowIndex).ValorDevolucao = Math.Round(objNotaFiscal.Itens(i).NotasDevolucao(row.RowIndex).QuantidadeDevolucao * objNotaFiscal.Itens(i).NotasDevolucao(row.RowIndex).UnitarioNota, 2, MidpointRounding.AwayFromZero)

            objNotaFiscal.Itens(i).NotasDevolucaoOriginal(row.RowIndex).QuantidadeDevolucao = objNotaFiscal.Itens(i).NotasDevolucao(row.RowIndex).QuantidadeDevolucao
            objNotaFiscal.Itens(i).NotasDevolucaoOriginal(row.RowIndex).ValorDevolucao = objNotaFiscal.Itens(i).NotasDevolucao(row.RowIndex).ValorDevolucao

            CType(gridNotaDeDevolucaoXNota.Rows(row.RowIndex).FindControl("txtValor"), TextBox).Text = objNotaFiscal.Itens(i).NotasDevolucao(row.RowIndex).ValorDevolucao.ToString("N2")
        Else
            objNotaFiscal.Itens(i).NotasDevolucao(row.RowIndex).QuantidadeDevolucao = 0
            objNotaFiscal.Itens(i).NotasDevolucao(row.RowIndex).ValorDevolucao = 0

            objNotaFiscal.Itens(i).NotasDevolucaoOriginal(row.RowIndex).QuantidadeDevolucao = 0
            objNotaFiscal.Itens(i).NotasDevolucaoOriginal(row.RowIndex).ValorDevolucao = 0

            CType(gridNotaDeDevolucaoXNota.Rows(row.RowIndex).FindControl("txtValor"), TextBox).Text = "0,00"
        End If

        SessaoSalvaNotaFiscal()
        AtualizaQtdeVlr()
    End Sub

    Protected Sub txtValor_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        SessaoRecuperaNotaFiscal()
        Dim txtValor As TextBox = CType(sender, TextBox)
        Dim row As GridViewRow = CType(txtValor.NamingContainer, GridViewRow)
        Dim i As Integer = Convert.ToInt32(hdfIndice.Value)

        If Not objNotaFiscal.NossaEmissao AndAlso objNotaFiscal.SubOperacao.QuantidadeFiscal AndAlso objNotaFiscal.Itens(i).NotasDevolucao.Count = 1 Then

            valorOriginal = CType(Session("ssValorOriginal" & HID.Value), Decimal)

            If (valorOriginal > (CDec(txtValor.Text) + 1) OrElse
                    valorOriginal < (CDec(txtValor.Text) - 1)) And Funcoes.VerificaPermissao("AjustarCentavoNotaFiscal", "ALTERAR") = False Then

                MsgBox(Me.Page, "Variação máxima do valor não pode ser superior/inferior à R$ 1,00. Qualquer dúvida entre em contato com o Suporte.")

                Exit Sub
            End If

        End If

        If objNotaFiscal.IUD <> "I" OrElse String.IsNullOrWhiteSpace(txtValor.Text) OrElse CDec(txtValor.Text) > objNotaFiscal.Itens(i).NotasDevolucao(row.RowIndex).ValorSaldo Then
            txtValor.Text = objNotaFiscal.Itens(i).NotasDevolucao(row.RowIndex).ValorDevolucao
            Exit Sub
        End If

        If txtValor.Text.Length > 0 AndAlso CDec(txtValor.Text) > 0 Then
            objNotaFiscal.Itens(i).NotasDevolucao(row.RowIndex).ValorDevolucao = txtValor.Text
            objNotaFiscal.Itens(i).NotasDevolucao(row.RowIndex).ValorLiquidoDevolucao = txtValor.Text
            SessaoSalvaNotaFiscal()
        ElseIf objNotaFiscal.Itens(i).NotasDevolucao(row.RowIndex).ValorDevolucao > 0 Then
            objNotaFiscal.Itens(i).NotasDevolucao(row.RowIndex).ValorDevolucao = 0
            objNotaFiscal.Itens(i).NotasDevolucao(row.RowIndex).ValorLiquidoDevolucao = 0
            SessaoSalvaNotaFiscal()
        End If
        AtualizaQtdeVlr()
    End Sub

    Protected Sub imgMenos_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)
        Try
            SessaoRecuperaNotaFiscal()

            Dim imgNotaFiscal As ImageButton = CType(sender, ImageButton)
            Dim row As GridViewRow = CType(imgNotaFiscal.NamingContainer, GridViewRow)
            Dim i As Integer = Convert.ToInt32(hdfIndice.Value)

            Dim valorNFAntes = Math.Round(objNotaFiscal.Itens(i).NotasDevolucao(row.RowIndex).ValorDevolucao, 2, MidpointRounding.AwayFromZero)

            Dim valorNF As Decimal = (Math.Round(objNotaFiscal.Itens(i).NotasDevolucao(row.RowIndex).ValorDevolucao, 2, MidpointRounding.AwayFromZero) - 0.01)

            Dim valorNFOriginal As Decimal = Math.Round(objNotaFiscal.Itens(i).NotasDevolucaoOriginal(row.RowIndex).ValorDevolucao, 2, MidpointRounding.AwayFromZero)

            If (valorNFOriginal - valorNF) > 1.5 And Funcoes.VerificaPermissao("AjustarCentavoNotaFiscal", "ALTERAR") = False Then
                MsgBox(Me.Page, "Variação máxima do valor não pode ser inferior à R$ 1,50. Qualquer dúvida entre em contato com o Suporte.")
                objNotaFiscal.Itens(i).NotasDevolucao(row.RowIndex).ValorDevolucao = valorNFAntes
            Else
                objNotaFiscal.Itens(i).NotasDevolucao(row.RowIndex).ValorDevolucao = valorNF
            End If

            CType(gridNotaDeDevolucaoXNota.Rows(row.RowIndex).FindControl("txtValor"), TextBox).Text = objNotaFiscal.Itens(i).NotasDevolucao(row.RowIndex).ValorDevolucao

            SessaoSalvaNotaFiscal()

            AtualizaQtdeVlr()

        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub imgMais_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)
        Try
            SessaoRecuperaNotaFiscal()

            Dim imgNotaFiscal As ImageButton = CType(sender, ImageButton)
            Dim row As GridViewRow = CType(imgNotaFiscal.NamingContainer, GridViewRow)
            Dim i As Integer = Convert.ToInt32(hdfIndice.Value)

            Dim valorNFAntes = Math.Round(objNotaFiscal.Itens(i).NotasDevolucao(row.RowIndex).ValorDevolucao, 2, MidpointRounding.AwayFromZero)

            Dim valorNF As Decimal = (Math.Round(objNotaFiscal.Itens(i).NotasDevolucao(row.RowIndex).ValorDevolucao, 2, MidpointRounding.AwayFromZero) + 0.01)

            Dim valorNFOriginal As Decimal = Math.Round(objNotaFiscal.Itens(i).NotasDevolucaoOriginal(row.RowIndex).ValorDevolucao, 2, MidpointRounding.AwayFromZero)

            If (valorNF - valorNFOriginal) > 1.5 And Funcoes.VerificaPermissao("AjustarCentavoNotaFiscal", "ALTERAR") = False Then
                MsgBox(Me.Page, "Variação máxima do valor não pode ser superior à R$ 1,50. Qualquer dúvida entre em contato com o Suporte.")
                objNotaFiscal.Itens(i).NotasDevolucao(row.RowIndex).ValorDevolucao = valorNFAntes
            Else
                objNotaFiscal.Itens(i).NotasDevolucao(row.RowIndex).ValorDevolucao = valorNF
            End If

            CType(gridNotaDeDevolucaoXNota.Rows(row.RowIndex).FindControl("txtValor"), TextBox).Text = objNotaFiscal.Itens(i).NotasDevolucao(row.RowIndex).ValorDevolucao

            SessaoSalvaNotaFiscal()

            AtualizaQtdeVlr()

        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Public Sub AtualizaQtdeVlr()
        SessaoRecuperaNotaFiscal()
        Dim i As Integer = Convert.ToInt32(hdfIndice.Value)
        lblQuantidade.Text = objNotaFiscal.Itens(i).NotasDevolucao.SomaQtde.ToString("N4")
        lblValor.Text = objNotaFiscal.Itens(i).NotasDevolucao.SomaVlr.ToString("N2")

        If CDec(lblQuantidade.Text) <> CDec(lblQtdeNota.Text) Then
            lblQuantidade.Style.Add("color", "red")
        Else
            lblQuantidade.Style.Remove("color")
        End If
        lblQuantidade.Text = (CDec(lblQtdeNota.Text) - CDec(lblQuantidade.Text)).ToString("N4")

        If CDec(lblValor.Text) <> CDec(lblVlrNota.Text) Then
            lblValor.Style.Add("color", "red")
        Else
            lblValor.Style.Remove("color")
        End If
        lblValor.Text = (CDec(lblVlrNota.Text) - CDec(lblValor.Text)).ToString("N2")

        grdResultDev.DataSource = From p In objNotaFiscal.Itens(i).NotasDevolucao
                                  Group By p.ItemNota.CodigoProduto
                                   Into Qtde = Sum(p.QuantidadeDevolucao), Valor = Sum(p.ValorDevolucao)
                                  Where Qtde > 0 Or Valor > 0
                                  Select New RetornoNotasDevolucao With {.indexItem = i,
                                                                         .Quantidade = Qtde,
                                                                         .Valor = Valor}

        'grdResultDev.DataSource = From p In objNotaFiscal.Itens(i).NotasDevolucao _
        '                         Group By p.UnitarioNota _
        '                          Into Qtde = Sum(p.QuantidadeDevolucao), Valor = Sum(p.ValorDevolucao) _
        '                         Order By UnitarioNota _
        '                        Select UnitarioNota, Qtde, Valor _
        '                        Where Qtde > 0 Or Valor > 0

        grdResultDev.DataBind()

        temTituloNoBanco = False
        temTituloEndosso = False

        For Each nt In objNotaFiscal.Itens(i).NotasDevolucao
            temTituloNoBanco = nt.Nota.VencimentosNota.Any(Function(x) x.CodigoProvisao = eProvisao.Previsao And x.CodigoSituacao = eSituacao.RemessaBancaria)

            temTituloEndosso = nt.Nota.VencimentosNota.Any(Function(x) x.CodigoProvisao = eProvisao.Previsao And x.CodigoSituacao = eSituacao.EndossoTitulo)

            If temTituloNoBanco Then Exit For

            If temTituloEndosso Then Exit For
        Next

        If temTituloNoBanco Then MsgBox(Me.Page, "Atenção... a(s) Nota(s) com Remessa Bancária não poderão ser utilizadas.", eTitulo.Info)

        If temTituloEndosso Then MsgBox(Me.Page, "Atenção... a(s) Nota(s) com Título de Endosso não poderão ser utilizadas.", eTitulo.Info)
    End Sub

    Protected Sub lnkConfirmar_Click(sender As Object, e As EventArgs) Handles lnkConfirmar.Click
        Dim i As Integer = Convert.ToInt32(hdfIndice.Value)
        SessaoRecuperaNotaFiscal()
        objNotaFiscal.Itens(i).NotasDevolucao.MsgDevolucao = ""

        Dim temDecimal As Boolean = False

        If objNotaFiscal.Itens(i).Produto.CodigoEmbalagem = 1 Then
            'ainda vou testar - Furlan - 23/08/2022
            For Each nDev In objNotaFiscal.Itens(i).NotasDevolucao

                If nDev.QuantidadeDevolucao > 0 Then
                    Dim vString As String = nDev.QuantidadeDevolucao.ToString()

                    If vString.Contains(",") Then
                        Dim vDecimal As Decimal = CDec(vString.Split(",")(1))

                        If vDecimal > 0 Then
                            MsgBox(Me.Page, "Produto a GRANEL não pode ter casa decimal.", eTitulo.Info)

                            nDev.QuantidadeDevolucao = 0

                            temDecimal = True
                        End If
                    End If
                End If
            Next
        End If

        If temDecimal Then
            SessaoSalvaNotaFiscal()

            lnkLimpar_Click(New Object, New EventArgs)

            Exit Sub
        End If

        If objNotaFiscal.Itens(i).QuantidadeFiscal > 0 AndAlso objNotaFiscal.Itens(i).NotasDevolucao.SomaQtde > 0 AndAlso objNotaFiscal.Itens(i).NotasDevolucao.SomaQtde <> objNotaFiscal.Itens(i).QuantidadeFiscal Then
            MsgBox(Me.Page, "A quantidade informada nos itens para devolução nao e a mesma do item da nota fiscal!")
        ElseIf objNotaFiscal.Itens(i).QuantidadeFiscal = 0 AndAlso objNotaFiscal.Itens(i).ValorTotal > 0 AndAlso Math.Round(objNotaFiscal.Itens(i).NotasDevolucao.SomaVlr, 2, MidpointRounding.AwayFromZero) <> Math.Round(objNotaFiscal.Itens(i).ValorTotal, 2, MidpointRounding.AwayFromZero) Then
            MsgBox(Me.Page, "O valor informado nos Itens para devolução nao e o mesmo do item da nota fiscal!")
        Else
            SessaoSalvaNotaFiscal()
            If (Session("ssTipoRetorno") IsNot Nothing AndAlso Session("ssTipoRetorno").ToString.Contains("objNotaDeDevolucaoXNota")) Then
                If TypeOf Me.Page Is NotaFiscalXItens Then
                    'Session("ssDevolucaoXNota" & HID.Value) = i
                    Dim retornoList = From p In objNotaFiscal.Itens(i).NotasDevolucao
                                      Group By p.ItemNota.CodigoProduto
                                       Into Qtde = Sum(p.QuantidadeDevolucao), Valor = Sum(p.ValorDevolucao)
                                      Where Qtde > 0 Or Valor > 0
                                      Select New RetornoNotasDevolucao With {.indexItem = i,
                                                                             .Quantidade = Qtde,
                                                                             .Valor = Valor}

                    CType(Me.Page, NotaFiscalXItens).CarregarNotasDevolucao(retornoList.ToList())
                End If

                Session.Remove("ssValorOriginal" & HID.Value)

                Popup.CloseDialog(Me.Page, "divNotaDeDevolucaoXNota")
            End If
        End If
    End Sub

    Protected Sub lnkRecarregar_Click(sender As Object, e As EventArgs) Handles lnkRecarregar.Click
        RecarregarNotas(txtUnitarioMedio.Text) 'chkReajusteUnitario.Checked)
    End Sub

    Protected Sub lnkLimpar_Click(sender As Object, e As EventArgs) Handles lnkLimpar.Click
        Dim i As Integer = Convert.ToInt32(hdfIndice.Value)

        SessaoRecuperaNotaFiscal()

        For Each row As [Lib].Negocio.NotaFiscalDevolucaoXNotaFiscal In objNotaFiscal.Itens(i).NotasDevolucao

            row.ValorDevolucao = Math.Round(row.ValorDevolucao, 2, MidpointRounding.AwayFromZero)

            row.QuantidadeDevolucao = 0
            row.ValorDevolucao = 0
        Next

        lblQuantidade.Text = "0,0000"
        lblValor.Text = "0,00"
        gridNotaDeDevolucaoXNota.DataSource = objNotaFiscal.Itens(i).NotasDevolucao.ToArray
        gridNotaDeDevolucaoXNota.DataBind()

        SessaoSalvaNotaFiscal()

        For j As Integer = gridNotaDeDevolucaoXNota.Rows.Count - 1 To 0 Step -1
            For Each n In objNotaFiscal.Itens(i).NotasDevolucao
                If gridNotaDeDevolucaoXNota.Rows(j).Cells(0).Text = n.Nota.Codigo Then

                    CType(gridNotaDeDevolucaoXNota.Rows(j).FindControl("imgMenos"), ImageButton).Visible = False
                    CType(gridNotaDeDevolucaoXNota.Rows(j).FindControl("imgMais"), ImageButton).Visible = False

                    If Not n.Nota.ChaveNFE.Length = 44 Then

                        CType(gridNotaDeDevolucaoXNota.Rows(j).FindControl("txtQuantidade"), TextBox).Text = "0,0000"
                        CType(gridNotaDeDevolucaoXNota.Rows(j).FindControl("txtQuantidade"), TextBox).Enabled = False

                        CType(gridNotaDeDevolucaoXNota.Rows(j).FindControl("txtValor"), TextBox).Text = "0,00"
                        CType(gridNotaDeDevolucaoXNota.Rows(j).FindControl("txtValor"), TextBox).Enabled = False

                        CType(gridNotaDeDevolucaoXNota.Rows(j).FindControl("imgInfNota"), ImageButton).Visible = True
                        CType(gridNotaDeDevolucaoXNota.Rows(j).FindControl("imgInfNota"), ImageButton).ToolTip = "Nota Fiscal sem a CHAVE ELETRÔNICA não pode ser utilizada."

                    ElseIf n.Nota.SubOperacao.Financeiro AndAlso (n.Nota.VencimentosNota Is Nothing OrElse n.Nota.VencimentosNota.Count = 0) Then
                        CType(gridNotaDeDevolucaoXNota.Rows(j).FindControl("txtQuantidade"), TextBox).Text = "0,0000"
                        CType(gridNotaDeDevolucaoXNota.Rows(j).FindControl("txtQuantidade"), TextBox).Enabled = False

                        CType(gridNotaDeDevolucaoXNota.Rows(j).FindControl("txtValor"), TextBox).Text = "0,00"
                        CType(gridNotaDeDevolucaoXNota.Rows(j).FindControl("txtValor"), TextBox).Enabled = False

                        CType(gridNotaDeDevolucaoXNota.Rows(j).FindControl("imgInfNota"), ImageButton).Visible = True
                        CType(gridNotaDeDevolucaoXNota.Rows(j).FindControl("imgInfNota"), ImageButton).ToolTip = "Nota Fiscal sem Título vinculado não pode ser utilizado."

                    ElseIf n.Nota.VencimentosNota IsNot Nothing AndAlso n.Nota.VencimentosNota.Count > 0 AndAlso n.Nota.VencimentosNota.Any(Function(s) s.RegistroMestre > 0) Then
                        CType(gridNotaDeDevolucaoXNota.Rows(j).FindControl("txtQuantidade"), TextBox).Text = "0,0000"
                        CType(gridNotaDeDevolucaoXNota.Rows(j).FindControl("txtQuantidade"), TextBox).Enabled = False

                        CType(gridNotaDeDevolucaoXNota.Rows(j).FindControl("txtValor"), TextBox).Text = "0,00"
                        CType(gridNotaDeDevolucaoXNota.Rows(j).FindControl("txtValor"), TextBox).Enabled = False

                        CType(gridNotaDeDevolucaoXNota.Rows(j).FindControl("imgInfNota"), ImageButton).Visible = True
                        CType(gridNotaDeDevolucaoXNota.Rows(j).FindControl("imgInfNota"), ImageButton).ToolTip = "Nota Fiscal com Título em Agrupamento não pode ser utilizado."

                    ElseIf n.Nota.VencimentosNota IsNot Nothing AndAlso n.Nota.VencimentosNota.Count > 0 Then
                        For Each t In n.Nota.VencimentosNota
                            If t.CodigoProvisao = eProvisao.Baixa OrElse t.CodigoSituacao = eSituacao.RemessaBancaria OrElse t.CodigoSituacao = eSituacao.EndossoTitulo Then

                                CType(gridNotaDeDevolucaoXNota.Rows(j).FindControl("txtQuantidade"), TextBox).Text = "0,0000"
                                CType(gridNotaDeDevolucaoXNota.Rows(j).FindControl("txtQuantidade"), TextBox).Enabled = False

                                CType(gridNotaDeDevolucaoXNota.Rows(j).FindControl("txtValor"), TextBox).Text = "0,00"
                                CType(gridNotaDeDevolucaoXNota.Rows(j).FindControl("txtValor"), TextBox).Enabled = False

                                CType(gridNotaDeDevolucaoXNota.Rows(j).FindControl("imgInfNota"), ImageButton).Visible = True

                                If t.CodigoSituacao = eSituacao.RemessaBancaria Then
                                    CType(gridNotaDeDevolucaoXNota.Rows(j).FindControl("imgInfNota"), ImageButton).ToolTip = "Nota Fiscal com Título em Cobrança Bancária não pode ser utilizada."
                                ElseIf t.CodigoSituacao = eSituacao.EndossoTitulo Then
                                    CType(gridNotaDeDevolucaoXNota.Rows(j).FindControl("imgInfNota"), ImageButton).ToolTip = "Nota Fiscal com Título de Endosso não pode ser utilizada."
                                Else
                                    CType(gridNotaDeDevolucaoXNota.Rows(j).FindControl("imgInfNota"), ImageButton).ToolTip = "Nota Fiscal com Título baixado não pode ser utilizada."
                                End If
                            Else
                                If Funcoes.VerificaPermissao("AjustarCentavoNotaFiscal", "ALTERAR") Then
                                    CType(gridNotaDeDevolucaoXNota.Rows(j).FindControl("imgMenos"), ImageButton).Visible = True
                                    CType(gridNotaDeDevolucaoXNota.Rows(j).FindControl("imgMais"), ImageButton).Visible = True
                                End If
                            End If
                        Next
                    Else
                        If Funcoes.VerificaPermissao("AjustarCentavoNotaFiscal", "ALTERAR") Then
                            CType(gridNotaDeDevolucaoXNota.Rows(j).FindControl("imgMenos"), ImageButton).Visible = True
                            CType(gridNotaDeDevolucaoXNota.Rows(j).FindControl("imgMais"), ImageButton).Visible = True
                        End If
                    End If
                End If
            Next
        Next j

        temTituloNoBanco = False
        temTituloEndosso = False

        For Each nt In objNotaFiscal.Itens(i).NotasDevolucao
            temTituloNoBanco = nt.Nota.VencimentosNota.Any(Function(x) x.CodigoProvisao = eProvisao.Previsao And x.CodigoSituacao = eSituacao.RemessaBancaria)

            temTituloEndosso = nt.Nota.VencimentosNota.Any(Function(x) x.CodigoProvisao = eProvisao.Previsao And x.CodigoSituacao = eSituacao.EndossoTitulo)

            If temTituloNoBanco Then Exit For

            If temTituloEndosso Then Exit For
        Next

        If temTituloNoBanco Then MsgBox(Me.Page, "Atenção... a(s) Nota(s) com Remessa Bancária não poderão ser utilizadas.", eTitulo.Info)

        If temTituloEndosso Then MsgBox(Me.Page, "Atenção... a(s) Nota(s) com Título de Endosso não poderão ser utilizadas.", eTitulo.Info)

    End Sub

    Protected Sub lnkRegravar_Click(sender As Object, e As EventArgs) Handles lnkRegravar.Click
        SessaoRecuperaNotaFiscal()
        Dim sqls As New ArrayList
        objNotaFiscal.Itens(hdfIndice.Value).IUD = "U"
        objNotaFiscal.Itens(hdfIndice.Value).NotasDevolucao.SalvarSql(sqls)

        Dim bd As New AcessaBanco
        bd.GravaBanco(sqls)
    End Sub

    Protected Sub lnkCancelar_Click(sender As Object, e As EventArgs) Handles lnkCancelar.Click
        Session.Remove("ssValorOriginal" & HID.Value)

        Popup.CloseDialog(Me.Page, "divNotaDeDevolucaoXNota")
    End Sub

    Protected Sub gridNotaDeDevolucaoXNota_RowDataBound(sender As Object, e As GridViewRowEventArgs) Handles gridNotaDeDevolucaoXNota.RowDataBound
        If e.Row.RowType = DataControlRowType.DataRow Then
            SessaoRecuperaNotaFiscal()

            If objNotaFiscal.Itens(hdfIndice.Value).QuantidadeFiscal > 0 Then
                CType(e.Row.FindControl("txtQuantidade"), TextBox).Enabled = objNotaFiscal.Itens(hdfIndice.Value).NotasDevolucao(e.Row.RowIndex).QuantidadeSaldo > 0
                CType(e.Row.FindControl("txtValor"), TextBox).Enabled = False
            Else
                CType(e.Row.FindControl("txtQuantidade"), TextBox).Enabled = False
                CType(e.Row.FindControl("txtValor"), TextBox).Enabled = objNotaFiscal.Itens(hdfIndice.Value).NotasDevolucao(e.Row.RowIndex).ValorSaldo > 0
            End If
        End If
    End Sub

    Protected Sub chkReajusteUnitario_CheckedChanged(sender As Object, e As EventArgs) Handles chkReajusteUnitario.CheckedChanged
        Dim i As Integer = hdfIndice.Value
        SessaoRecuperaNotaFiscal()

        If chkReajusteUnitario.Checked Then
            txtUnitarioMedio.Text = objNotaFiscal.Pedido.Itens.Where(Function(s) s.CodigoProduto = objNotaFiscal.Itens(i).CodigoProduto).FirstOrDefault.UnitarioMedioFaturamento.ToString("N10")
            divUnitarioMedio.Style.Clear()
        Else
            txtUnitarioMedio.Text = 0
            divUnitarioMedio.Style.Add("Display", "none")
        End If

        RecarregarNotas(txtUnitarioMedio.Text)
    End Sub

    Private Sub RecarregarNotas(Optional ByVal UnitarioMedio As Decimal = 0)
        Dim i As Integer = Convert.ToInt32(hdfIndice.Value)
        SessaoRecuperaNotaFiscal()
        objNotaFiscal.Itens(i).NotasDevolucao = New ListNotaFiscalDevolucaoXNotaFiscal(objNotaFiscal.Itens(i))
        objNotaFiscal.Itens(i).NotasDevolucao.CarregarNotasParaSelecao(UnitarioMedio)

        If objNotaFiscal.IUD = "I" Then
            For Each row As [Lib].Negocio.NotaFiscalDevolucaoXNotaFiscal In objNotaFiscal.Itens(i).NotasDevolucao

                row.QuantidadeDevolucao = 0
                row.ValorDevolucao = 0
                'AJUSTE - NUTRI PEDIDO 1372 - 2024-12-27
                'Não tenho a devolução lançada ainda, então minha devolução é zero
                'row.ValorDevolucao = Math.Round(row.ValorDevolucao, 2, MidpointRounding.AwayFromZero)
                'If row.Nota.VencimentosNota IsNot Nothing AndAlso row.Nota.VencimentosNota.Count > 0 AndAlso row.Nota.VencimentosNota.Any(Function(s) s.CodigoSituacao = eSituacao.Normal And s.CodigoProvisao = eProvisao.Baixa) Then
                '    row.QuantidadeDevolucao = 0
                '    row.ValorDevolucao = 0
                'End If

            Next
        End If

        SessaoSalvaNotaFiscal()

        gridNotaDeDevolucaoXNota.DataSource = objNotaFiscal.Itens(i).NotasDevolucao.ToArray
        gridNotaDeDevolucaoXNota.DataBind()

        For j As Integer = gridNotaDeDevolucaoXNota.Rows.Count - 1 To 0 Step -1
            For Each n In objNotaFiscal.Itens(i).NotasDevolucao
                If gridNotaDeDevolucaoXNota.Rows(j).Cells(0).Text = n.Nota.Codigo Then

                    CType(gridNotaDeDevolucaoXNota.Rows(j).FindControl("imgMenos"), ImageButton).Visible = False
                    CType(gridNotaDeDevolucaoXNota.Rows(j).FindControl("imgMais"), ImageButton).Visible = False

                    If Not n.Nota.ChaveNFE.Length = 44 Then

                        CType(gridNotaDeDevolucaoXNota.Rows(j).FindControl("txtQuantidade"), TextBox).Text = "0,0000"
                        CType(gridNotaDeDevolucaoXNota.Rows(j).FindControl("txtQuantidade"), TextBox).Enabled = False

                        CType(gridNotaDeDevolucaoXNota.Rows(j).FindControl("txtValor"), TextBox).Text = "0,00"
                        CType(gridNotaDeDevolucaoXNota.Rows(j).FindControl("txtValor"), TextBox).Enabled = False

                        CType(gridNotaDeDevolucaoXNota.Rows(j).FindControl("imgInfNota"), ImageButton).Visible = True
                        CType(gridNotaDeDevolucaoXNota.Rows(j).FindControl("imgInfNota"), ImageButton).ToolTip = "Nota Fiscal sem a CHAVE ELETRÔNICA não pode ser utilizada."

                    ElseIf n.Nota.SubOperacao.Financeiro AndAlso (n.Nota.VencimentosNota Is Nothing OrElse n.Nota.VencimentosNota.Count = 0) Then
                        CType(gridNotaDeDevolucaoXNota.Rows(j).FindControl("txtQuantidade"), TextBox).Text = "0,0000"
                        CType(gridNotaDeDevolucaoXNota.Rows(j).FindControl("txtQuantidade"), TextBox).Enabled = False

                        CType(gridNotaDeDevolucaoXNota.Rows(j).FindControl("txtValor"), TextBox).Text = "0,00"
                        CType(gridNotaDeDevolucaoXNota.Rows(j).FindControl("txtValor"), TextBox).Enabled = False

                        CType(gridNotaDeDevolucaoXNota.Rows(j).FindControl("imgInfNota"), ImageButton).Visible = True
                        CType(gridNotaDeDevolucaoXNota.Rows(j).FindControl("imgInfNota"), ImageButton).ToolTip = "Nota Fiscal sem Título vinculado não pode ser utilizado."

                    ElseIf n.Nota.VencimentosNota IsNot Nothing AndAlso n.Nota.VencimentosNota.Count > 0 AndAlso n.Nota.VencimentosNota.Any(Function(s) s.CodigoSituacao = eSituacao.Normal And s.CodigoProvisao = eProvisao.Baixa) Then

                        CType(gridNotaDeDevolucaoXNota.Rows(j).FindControl("txtQuantidade"), TextBox).Text = "0,0000"
                        CType(gridNotaDeDevolucaoXNota.Rows(j).FindControl("txtQuantidade"), TextBox).Enabled = False

                        CType(gridNotaDeDevolucaoXNota.Rows(j).FindControl("txtValor"), TextBox).Text = "0,00"
                        CType(gridNotaDeDevolucaoXNota.Rows(j).FindControl("txtValor"), TextBox).Enabled = False

                        CType(gridNotaDeDevolucaoXNota.Rows(j).FindControl("imgInfNota"), ImageButton).Visible = True
                        CType(gridNotaDeDevolucaoXNota.Rows(j).FindControl("imgInfNota"), ImageButton).ToolTip = "Nota Fiscal com Título baixado não pode ser utilizada."

                    ElseIf n.Nota.VencimentosNota IsNot Nothing AndAlso n.Nota.VencimentosNota.Count > 0 AndAlso n.Nota.VencimentosNota.Any(Function(s) s.RegistroMestre > 0) Then
                        CType(gridNotaDeDevolucaoXNota.Rows(j).FindControl("txtQuantidade"), TextBox).Text = "0,0000"
                        CType(gridNotaDeDevolucaoXNota.Rows(j).FindControl("txtQuantidade"), TextBox).Enabled = False

                        CType(gridNotaDeDevolucaoXNota.Rows(j).FindControl("txtValor"), TextBox).Text = "0,00"
                        CType(gridNotaDeDevolucaoXNota.Rows(j).FindControl("txtValor"), TextBox).Enabled = False

                        CType(gridNotaDeDevolucaoXNota.Rows(j).FindControl("imgInfNota"), ImageButton).Visible = True
                        CType(gridNotaDeDevolucaoXNota.Rows(j).FindControl("imgInfNota"), ImageButton).ToolTip = "Nota Fiscal com Título em Agrupamento não pode ser utilizado."

                    ElseIf n.Nota.VencimentosNota IsNot Nothing AndAlso n.Nota.VencimentosNota.Count > 0 AndAlso n.Nota.VencimentosNota.Any(Function(s) s.CodigoSituacao = eSituacao.RemessaBancaria) Then
                        CType(gridNotaDeDevolucaoXNota.Rows(j).FindControl("txtQuantidade"), TextBox).Text = "0,0000"
                        CType(gridNotaDeDevolucaoXNota.Rows(j).FindControl("txtQuantidade"), TextBox).Enabled = False

                        CType(gridNotaDeDevolucaoXNota.Rows(j).FindControl("txtValor"), TextBox).Text = "0,00"
                        CType(gridNotaDeDevolucaoXNota.Rows(j).FindControl("txtValor"), TextBox).Enabled = False

                        CType(gridNotaDeDevolucaoXNota.Rows(j).FindControl("imgInfNota"), ImageButton).Visible = True
                        CType(gridNotaDeDevolucaoXNota.Rows(j).FindControl("imgInfNota"), ImageButton).ToolTip = "Nota Fiscal com Título em Cobrança Bancária não pode ser utilizada."

                    ElseIf n.Nota.VencimentosNota IsNot Nothing AndAlso n.Nota.VencimentosNota.Count > 0 AndAlso n.Nota.VencimentosNota.Any(Function(s) s.CodigoSituacao = eSituacao.EndossoTitulo) Then
                        CType(gridNotaDeDevolucaoXNota.Rows(j).FindControl("txtQuantidade"), TextBox).Text = "0,0000"
                        CType(gridNotaDeDevolucaoXNota.Rows(j).FindControl("txtQuantidade"), TextBox).Enabled = False

                        CType(gridNotaDeDevolucaoXNota.Rows(j).FindControl("txtValor"), TextBox).Text = "0,00"
                        CType(gridNotaDeDevolucaoXNota.Rows(j).FindControl("txtValor"), TextBox).Enabled = False

                        CType(gridNotaDeDevolucaoXNota.Rows(j).FindControl("imgInfNota"), ImageButton).Visible = True
                        CType(gridNotaDeDevolucaoXNota.Rows(j).FindControl("imgInfNota"), ImageButton).ToolTip = "Nota Fiscal com Título de Endosso não pode ser utilizada."

                    ElseIf Not objNotaFiscal.NossaEmissao AndAlso gridNotaDeDevolucaoXNota.Rows.Count = 1 Then
                        CType(gridNotaDeDevolucaoXNota.Rows(j).FindControl("txtValor"), TextBox).Enabled = True

                        valorOriginal = CDec(CType(gridNotaDeDevolucaoXNota.Rows(j).FindControl("txtValor"), TextBox).Text)

                        Session("ssValorOriginal" & HID.Value) += valorOriginal
                        If Funcoes.VerificaPermissao("AjustarCentavoNotaFiscal", "ALTERAR") Then
                            CType(gridNotaDeDevolucaoXNota.Rows(j).FindControl("imgMenos"), ImageButton).Visible = True
                            CType(gridNotaDeDevolucaoXNota.Rows(j).FindControl("imgMais"), ImageButton).Visible = True
                        End If
                    Else
                        If Funcoes.VerificaPermissao("AjustarCentavoNotaFiscal", "ALTERAR") Then
                            CType(gridNotaDeDevolucaoXNota.Rows(j).FindControl("imgMenos"), ImageButton).Visible = True
                            CType(gridNotaDeDevolucaoXNota.Rows(j).FindControl("imgMais"), ImageButton).Visible = True
                        End If
                    End If

                End If
            Next
        Next j

        AtualizaQtdeVlr()
    End Sub

    Protected Sub txtUnitarioMedio_TextChanged(sender As Object, e As EventArgs) Handles txtUnitarioMedio.TextChanged
        If txtUnitarioMedio.Text > 0 Then
            RecarregarNotas(txtUnitarioMedio.Text)
        End If
    End Sub
End Class