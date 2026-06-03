Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class AnaliseDeCreditoParametros
    Inherits BasePage

#Region "Variáveis Locais"
    Dim ObjParametrosAno As [Lib].Negocio.ListParametrosAnaliseDeCredito

    Dim ObjParametros As [Lib].Negocio.ParametrosAnaliseDeCredito
    Private ListCulturaPortifolios As ListCulturaPortifolio
    Private objCulturaPortifolio As CulturaPortifolio
    Private objParametrosAnaliseDeCreditoCultura As ParametrosAnaliseDeCreditoCultura
#End Region

#Region "SESSAO"
    '******************************************************************
    '***************** LISTA DE DEFINICOES DO ANO  ********************
    '******************************************************************
    Private Sub SessaoSalvaParametrosAno()
        Session("ObjListParametros") = ObjParametrosAno
    End Sub

    Private Sub SessaoRecuperaParametrosAno()
        ObjParametrosAno = CType(Session("ObjListParametros"), [Lib].Negocio.ListParametrosAnaliseDeCredito)
    End Sub

    '******************************************************************
    Private Sub SessaoRecuperaParametros()
        SessaoRecuperaParametrosAno()
        If ObjParametrosAno.Count = 0 Then
            ObjParametrosAno = New [Lib].Negocio.ListParametrosAnaliseDeCredito(ddlAno.SelectedValue)
            ObjParametros = New [Lib].Negocio.ParametrosAnaliseDeCredito
            ObjParametros.DefinicaoAno = 0
            ObjParametros.DataDefinicao = Date.Now
            ObjParametrosAno.Add(ObjParametros)
            SessaoSalvaParametrosAno()
        ElseIf ObjParametrosAno.Count = 1 Then
            ObjParametros = ObjParametrosAno(0)
        Else
            ObjParametros = ObjParametrosAno(GridDefinicoes.SelectedIndex)
        End If
    End Sub

    '******************************************************************
    '***************** LISTA DE CULTURA PORTIFOLIO  *******************
    '******************************************************************
    Private Sub SessaoSalvaListaCulturaPortifolio()
        Session("ListCulturaPortifolios") = ListCulturaPortifolios
    End Sub

    Private Sub SessaoRecuperaListaCulturaPortifolio()
        ListCulturaPortifolios = Session("ListCulturaPortifolios")
    End Sub
    '******************************************************************

    Private Sub SessaoSalvaParametroAnaliseDeCreditoCultura()
        Session("objParametrosAnaliseDeCreditoCultura") = objParametrosAnaliseDeCreditoCultura
    End Sub

    Private Sub SessaoRecuperaParametroAnaliseDeCreditoCultura()
        objParametrosAnaliseDeCreditoCultura = Session("objParametrosAnaliseDeCreditoCultura")
    End Sub
#End Region

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Revenda)
        If Not IsPostBack And IsConnect Then
            If Funcoes.VerificaPermissao("AnaliseDeCreditoParametros", "ACESSAR") Then
                ddl.Carregar(ddlCultura, CarregarDDL.Tabela.Cultura, "")
                ucSelecaoProduto.ProdutosVisiveis = False
                ucSelecaoProduto.NomeUC = "Grupos de Produtos"
                TCItensParametro.ActiveTabIndex = 0
            Else
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página!", "~/Revenda.aspx")
                Exit Sub
            End If
        End If
    End Sub

    Protected Sub ddlAno_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ddlAno.SelectedIndexChanged
        'Carrega ddlSafra
        ddlSafra.Items.Clear()
        If ddlAno.SelectedValue = "" Then Exit Sub
        ddl.Carregar(ddlSafra, CarregarDDL.Tabela.Safra, "Year(Vencimento) =" & ddlAno.SelectedValue, False)
        '**************************

        Limpar()
        ObjParametrosAno = New [Lib].Negocio.ListParametrosAnaliseDeCredito(ddlAno.SelectedValue)

        If ObjParametrosAno.Count = 0 Then
            SessaoSalvaParametrosAno()
            SessaoRecuperaParametros()
        End If

        GridDefinicoes.DataSource = ObjParametrosAno.ToArray
        GridDefinicoes.DataBind()
        SessaoSalvaParametrosAno()

    End Sub

    Protected Sub GridDefinicoes_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridDefinicoes.SelectedIndexChanged
        SessaoRecuperaParametros()
        txtPercContasAbertas.Text = ObjParametros.PercContasAbertas.ToString("N2")
        chkPerguntas.Checked = ObjParametros.PerguntasComportamentais
        txtPercLimCredA.Text = ObjParametros.PercLimiteCreditoA.ToString("N2")
        txtPercLimCredB.Text = ObjParametros.PercLimiteCreditoB.ToString("N2")
        txtPercLimCredC.Text = ObjParametros.PercLimiteCreditoC.ToString("N2")

        txtRiscoAlto.Text = ObjParametros.PercReducaoRiscoAlto.ToString("N2")
        txtRiscoMedio.Text = ObjParametros.PercReducaoRiscoMedio.ToString("N2")
        txtRiscoBaixo.Text = ObjParametros.PercReducaoRiscoBaixo.ToString("N2")

        txtCustoArrendamentoha.Text = ObjParametros.CustoArrendamentoHa.ToString("N2")

        txtCotacaoDolar.Text = ObjParametros.CotacaoDolar.ToString("N6")

        gridCulturas.DataSource = ObjParametros.Culturas.ToArray
        gridCulturas.DataBind()

        gridPerguntas.DataSource = ObjParametros.Perguntas.ToArray
        gridPerguntas.DataBind()

    End Sub

    Protected Sub gridCulturas_RowEditing(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewEditEventArgs)
        gridCulturas.EditIndex = e.NewEditIndex
        SessaoRecuperaParametros()
        gridCulturas.DataSource = ObjParametros.Culturas.ToArray
        gridCulturas.DataBind()

    End Sub

    Protected Sub gridCulturas_RowCancelingEdit(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewCancelEditEventArgs)
        gridCulturas.EditIndex = e.Cancel
        SessaoRecuperaParametros()
        gridCulturas.DataSource = ObjParametros.Culturas.ToArray
        gridCulturas.DataBind()
    End Sub

    Protected Sub gridCulturas_RowDeleting(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewDeleteEventArgs)
        SessaoRecuperaParametros()
        ObjParametros.Culturas.RemoveAt(e.RowIndex)
        SessaoSalvaParametrosAno()
        gridCulturas.DataSource = ObjParametros.Culturas.ToArray
        gridCulturas.DataBind()
    End Sub

    Protected Sub gridCulturas_RowUpdating(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewUpdateEventArgs)
        SessaoRecuperaParametros()
        With gridCulturas.Rows(e.RowIndex)
            ObjParametros.Culturas(e.RowIndex).Produtividade = CType(.FindControl("txtProdutividade"), TextBox).Text
            ObjParametros.Culturas(e.RowIndex).PrecoSaco = CType(.FindControl("txtPrecoSaco"), TextBox).Text
            ObjParametros.Culturas(e.RowIndex).CustoTotalHa = CType(.FindControl("txtCustoTotalHa"), TextBox).Text
            ObjParametros.Culturas(e.RowIndex).CustoPortifolioHa = CType(.FindControl("txtCustoPortifolioHa"), TextBox).Text
        End With
        SessaoSalvaParametrosAno()
    End Sub

    Private Sub Limpar()
        If ddlAno.SelectedIndex = 0 Then ddlAno.SelectedValue = Date.Now.Year + 1

        ObjParametrosAno = New [Lib].Negocio.ListParametrosAnaliseDeCredito(ddlAno.SelectedValue)
        SessaoSalvaParametrosAno()

        txtPercContasAbertas.Text = "0,00"
        chkPerguntas.Checked = False
        txtPercLimCredA.Text = "0,00"
        txtPercLimCredB.Text = "0,00"
        txtPercLimCredC.Text = "0,00"
        txtPeso.Text = "0,00"
        txtRiscoAlto.Text = "0,00"
        txtRiscoMedio.Text = "0,00"
        txtRiscoBaixo.Text = "0,00"
        txtCustoArrendamentoha.Text = "0,00"
        txtCotacaoDolar.Text = "0,000000"

        ucSelecaoProduto.Limpar()

        gridCulturas.DataSource = Nothing
        gridCulturas.DataBind()

        gridPerguntas.DataSource = Nothing
        gridPerguntas.DataBind()

        GridCulturaPortifolio.DataSource = Nothing
        GridCulturaPortifolio.DataBind()

    End Sub

    Public Function DefinicaoValida() As Boolean
        For Each cult As [Lib].Negocio.ParametrosAnaliseDeCreditoCultura In ObjParametros.Culturas
            If cult.Produtividade = 0 Or cult.PrecoSaco = 0 Or cult.CustoPortifolioHa = 0 Or cult.CustoTotalHa = 0 Then Return False
        Next
        Return True
    End Function


    Protected Sub chkPerguntas_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        SessaoRecuperaParametros()

        If chkPerguntas.Checked Then
            Dim perc As Decimal
            For Each perg As [Lib].Negocio.ParametrosAnaliseDeCreditoPergunta In ObjParametros.Perguntas
                perc += perg.PercPeso
            Next
            If perc = 100 Then
                ObjParametros.PerguntasComportamentais = True
            Else
                chkPerguntas.Checked = False
                ObjParametros.PerguntasComportamentais = False
                MsgBox(Me.Page, "As Perguntas só podem ser usadas se a soma dos Pesos for igual a 100 %")
            End If
        Else
            ObjParametros.PerguntasComportamentais = False
        End If

        SessaoSalvaParametrosAno()
    End Sub

    '***************Carrega os registros do Grid de Cultura Portifólio**************************************'
    Protected Sub ImdPortifolio_Click(ByVal sender As Object, ByVal e As EventArgs)
        SessaoRecuperaParametros()
        Dim imd As LinkButton = CType(sender, LinkButton)
        Dim row As GridViewRow = CType(imd.NamingContainer, GridViewRow)
        Dim i As Integer = ObjParametros.Culturas.FindIndex(Function(s) s.CodigoSafra = row.Cells(1).Text And s.CodigoCultura = row.Cells(2).Text)
        HDSafraCultura.Value = i
        CarregaListaCulturaPortifoliosNoGrid()
        SessaoSalvaParametroAnaliseDeCreditoCultura()
    End Sub

    Private Sub CarregaListaCulturaPortifoliosNoGrid()
        If ObjParametros Is Nothing Then SessaoRecuperaParametros()

        lblSafra.Text = ObjParametros.Culturas(HDSafraCultura.Value).CodigoSafra
        lblCultura.Text = ObjParametros.Culturas(HDSafraCultura.Value).NomeCultura

        objParametrosAnaliseDeCreditoCultura = ObjParametros.Culturas(HDSafraCultura.Value)
        ListCulturaPortifolios = New ListCulturaPortifolio(objParametrosAnaliseDeCreditoCultura)
        '*************** A lista precisa ser salva neste ponto para que sejam totalizados os valores do grid.
        SessaoSalvaListaCulturaPortifolio()

        GridCulturaPortifolio.DataSource = ListCulturaPortifolios
        GridCulturaPortifolio.DataBind()
    End Sub

    '***************Inclui registros no Grid de Cultura Portifólio**************************************'
    Protected Sub lnkAdicionarGrupoProduto_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkAdicionarGrupoProduto.Click
        SessaoRecuperaParametroAnaliseDeCreditoCultura()
        SessaoRecuperaListaCulturaPortifolio()
        'Dim tam As String = ucSelecaoProduto.GetStringGrupoProdutoSelecionado.Split(";")(0).Split("|")(0)
        Dim ret As String() = ucSelecaoProduto.GetStringGrupoProdutoSelecionado.Split(";")(0).Split("|")(1).Split(",")

        For i As Integer = 0 To ret.Length - 1
            Dim pos As Integer = i
            If Not existe(ret(pos)) Then
                If IsNothing(ListCulturaPortifolios.Find(Function(s) s.CodigoGrupoProduto = ret(pos))) Then
                    objCulturaPortifolio = New CulturaPortifolio(objParametrosAnaliseDeCreditoCultura)
                    objCulturaPortifolio.IUD = "I"
                    objCulturaPortifolio.CodigoGrupoProduto = ret(i)
                    ListCulturaPortifolios.Add(objCulturaPortifolio)
                Else
                    MsgBox(Me.Page, "Os Grupos repetidos não foram Adicionados!!!")
                End If
            Else
                MsgBox(Me.Page, "Um Grupo de Nível Superior ou Inferior já foi adicionado!!!")
            End If
        Next

        ucSelecaoProduto.SelecionaTodosOsGrupos(False)
        GridCulturaPortifolio.DataSource = ListCulturaPortifolios
        GridCulturaPortifolio.DataBind()
        SessaoSalvaListaCulturaPortifolio()
    End Sub

    Private Function existe(ByVal pGrupo As String) As Boolean
        Dim retorno As Boolean = False
        Dim sg As String = ""

        For i As Integer = 5 To 1 Step -1
            sg = Left(pGrupo, IIf(pGrupo.Length - i < 0, 1, pGrupo.Length - i))
            If Not String.IsNullOrWhiteSpace(sg) Then
                retorno = Not IsNothing(ListCulturaPortifolios.Find(Function(s) s.CodigoGrupoProduto = sg))
                If retorno Then
                    Exit For
                End If
            End If
        Next

        If Not retorno Then
            For i As Integer = 1 To 4 Step 1
                Dim pos As Integer = i
                retorno = Not IsNothing(ListCulturaPortifolios.Find(Function(s) Left(s.CodigoGrupoProduto, IIf(s.CodigoGrupoProduto.Length - pos < 0, 0, s.CodigoGrupoProduto.Length - pos)) = pGrupo))
                If retorno Then
                    Exit For
                End If
            Next
        End If
        Return retorno
    End Function

    Protected Sub txtQuantidade_TextChanged(ByVal sender As Object, ByVal e As EventArgs)
        SessaoRecuperaListaCulturaPortifolio()

        Dim txt As TextBox = CType(sender, TextBox)
        Dim row As GridViewRow = CType(txt.NamingContainer, GridViewRow)

        Dim CodGrupoProduto As String = GridCulturaPortifolio.Rows(row.RowIndex).Cells(1).Text
        Dim objCulturaPortifolio As CulturaPortifolio
        objCulturaPortifolio = ListCulturaPortifolios.Find(Function(s) s.CodigoGrupoProduto = CodGrupoProduto)
        objCulturaPortifolio.Quantidade = txt.Text

        SessaoSalvaListaCulturaPortifolio()
    End Sub

    Protected Sub txtValor_TextChanged(ByVal sender As Object, ByVal e As EventArgs)
        SessaoRecuperaListaCulturaPortifolio()

        Dim txt As TextBox = CType(sender, TextBox)
        Dim row As GridViewRow = CType(txt.NamingContainer, GridViewRow)

        Dim CodGrupoProduto As String = GridCulturaPortifolio.Rows(row.RowIndex).Cells(1).Text
        Dim objCulturaPortifolio As CulturaPortifolio
        objCulturaPortifolio = ListCulturaPortifolios.Find(Function(s) s.CodigoGrupoProduto = CodGrupoProduto)
        objCulturaPortifolio.Valor = txt.Text
        TotalizaCustoPortifolioHa()
        SessaoSalvaListaCulturaPortifolio()

        '********Força a totalização dos Valores do grid **************************************************
        GridCulturaPortifolio_RowDataBound(Me, New GridViewRowEventArgs(GridCulturaPortifolio.FooterRow))
    End Sub

    Private Sub TotalizaCustoPortifolioHa()
        SessaoRecuperaParametroAnaliseDeCreditoCultura()
        SessaoRecuperaListaCulturaPortifolio()
        objParametrosAnaliseDeCreditoCultura.CustoPortifolioHa = ListCulturaPortifolios.CustoTotalPortifolio
    End Sub

    '***************Salva os registros do Grid de Cultura Portifólio**************************************'
    'Protected Sub lnkNovo_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkNovo.Click
    '    SessaoRecuperaListaCulturaPortifolio()
    '    If IsNothing(ListCulturaPortifolios.Find(Function(s) s.Valor = 0.0)) Then
    '        If ListCulturaPortifolios.Salvar() Then
    '            MsgBox(Me.Page, "Registros gravados com Sucesso.", eTitulo.Sucess)
    '            AtualizaTotalPortifolioGridCulturas()
    '        End If
    '    Else
    '        MsgBox(Me.Page, "Campo Valor é Obrigatório!")
    '    End If
    'End Sub

    Protected Sub GridCulturaPortifolio_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridCulturaPortifolio.RowDataBound
        If e.Row.RowType = DataControlRowType.Footer Then
            SessaoRecuperaListaCulturaPortifolio()
            If Not IsNothing(ListCulturaPortifolios) Then
                e.Row.Cells(4).Text = String.Format("{0:c}", ListCulturaPortifolios.CustoTotalPortifolio)
                e.Row.Cells(4).HorizontalAlign = HorizontalAlign.Right
            End If
        End If
    End Sub

    Protected Sub lnkLimpar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkLimpar.Click
        CarregaListaCulturaPortifoliosNoGrid()
    End Sub

    'Private Sub AtualizaTotalPortifolioGridCulturas()
    '    SessaoRecuperaParametroAnaliseDeCreditoCultura()
    '    SessaoRecuperaParametros()
    '    objParametrosAnaliseDeCreditoCultura.CustoPortifolioHa = ListCulturaPortifolios.CustoTotalPortifolio
    '    objParametrosAnaliseDeCreditoCultura.IUD = "U"
    '    objParametrosAnaliseDeCreditoCultura.Salvar()

    '    Dim index = gridCulturas.EditIndex
    '    gridCulturas.DataSource = ObjParametros.Culturas
    '    gridCulturas.DataBind()
    '    gridCulturas.SetEditRow(index)
    'End Sub

    Protected Sub lnkLimparB_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkLimparB.Click
        Limpar()
    End Sub

    Protected Sub lnkNovoB_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkNovoB.Click
        If CDec(txtPercLimCredA.Text) < 0 Or CDec(txtPercLimCredA.Text) > 100 _
        Or CDec(txtPercLimCredB.Text) < 0 Or CDec(txtPercLimCredB.Text) > 100 _
        Or CDec(txtPercLimCredC.Text) < 0 Or CDec(txtPercLimCredC.Text) > 100 Then
            MsgBox(Me.Page, "Os Percetuais de Limite de Credito devem estar numa faixa de 0 a 100")
            Exit Sub
        End If

        If CDec(txtRiscoAlto.Text) < 0 Or CDec(txtRiscoAlto.Text) > 100 _
        Or CDec(txtRiscoMedio.Text) < 0 Or CDec(txtRiscoMedio.Text) > 100 _
        Or CDec(txtRiscoBaixo.Text) < 0 Or CDec(txtRiscoBaixo.Text) > 100 Then
            MsgBox(Me.Page, "Os Percetuais de Risco devem estar numa faixa de 0 a 100")
            Exit Sub
        End If

        SessaoRecuperaParametros()
        If Not DefinicaoValida() Then
            MsgBox(Me.Page, "Os Parametros da Cultura nao podem ser iguais a zero")
            Exit Sub
        End If

        Dim def As Integer = 0
        For Each par As [Lib].Negocio.ParametrosAnaliseDeCredito In ObjParametrosAno
            If par.DefinicaoAno > def Then def = par.DefinicaoAno
        Next

        ObjParametros.IUD = "I"
        ObjParametros.Ano = ddlAno.SelectedValue
        ObjParametros.DefinicaoAno = def + 1
        ObjParametros.DataDefinicao = Date.Now
        ObjParametros.PercContasAbertas = txtPercContasAbertas.Text
        ObjParametros.PerguntasComportamentais = chkPerguntas.Checked
        ObjParametros.PercLimiteCreditoA = txtPercLimCredA.Text
        ObjParametros.PercLimiteCreditoB = txtPercLimCredB.Text
        ObjParametros.PercLimiteCreditoC = txtPercLimCredC.Text
        ObjParametros.PercReducaoRiscoAlto = txtRiscoAlto.Text
        ObjParametros.PercReducaoRiscoMedio = txtRiscoMedio.Text
        ObjParametros.PercReducaoRiscoBaixo = txtRiscoBaixo.Text
        ObjParametros.CustoArrendamentoHa = txtCustoArrendamentoha.Text
        ObjParametros.CotacaoDolar = txtCotacaoDolar.Text

        Dim i As Integer = 1
        For Each row In ObjParametros.Perguntas
            row.CodigoPergunta = i
            i += 1
        Next

        If ObjParametros.Salvar Then
            MsgBox(Me.Page, "Definicao salva com Sucesso.", eTitulo.Sucess)
            ddlAno_SelectedIndexChanged(ddlAno, e)
        Else
            MsgBox(Me.Page, "Erro ao Salvar a Definicao")
        End If
    End Sub

    Protected Sub lnkAjudaB_Click(sender As Object, e As EventArgs) Handles lnkAjudaB.Click
        Try
            Funcoes.Ajuda(Me.Page, "AnaliseDeCreditoParametros")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub

    Protected Sub lnkGravarPergunta_Click(sender As Object, e As EventArgs) Handles lnkGravarPergunta.Click
        If txtPergunta.Text.Trim.Length = 0 Then
            MsgBox(Me.Page, "Formule uma Pergunta.")
            Exit Sub
        End If

        If Not IsNumeric(txtPeso.Text) OrElse CDec(txtPeso.Text) <= 0 OrElse CDec(txtPeso.Text) >= 100 Then
            MsgBox(Me.Page, "Informe um Percentual de 0 a 100 no campo Peso da Pergunta.")
            Exit Sub
        End If

        SessaoRecuperaParametrosAno()
        If ObjParametrosAno Is Nothing Then
            ObjParametrosAno = New [Lib].Negocio.ListParametrosAnaliseDeCredito(ddlAno.SelectedValue)
            ObjParametros = New [Lib].Negocio.ParametrosAnaliseDeCredito
        Else
            SessaoRecuperaParametros()
        End If

        Dim Pergunta As [Lib].Negocio.ParametrosAnaliseDeCreditoPergunta
        Dim i As Integer = ObjParametros.Perguntas.FindIndex(Function(s) s.CodigoPergunta = txtId.Text)

        If i >= 0 Then
            Pergunta = ObjParametros.Perguntas(i)
        Else
            Dim CodPerg As Integer = ObjParametros.Perguntas.Max(Function(s) s.CodigoPergunta)
            Pergunta = New [Lib].Negocio.ParametrosAnaliseDeCreditoPergunta(ObjParametros)
            Pergunta.CodigoPergunta = CodPerg + 1
        End If

        Pergunta.Descricao = txtPergunta.Text.ToUpper
        Pergunta.PercPeso = txtPeso.Text
        If i < 0 Then ObjParametros.Perguntas.Add(Pergunta)

        ObjParametros.PerguntasComportamentais = False
        chkPerguntas.Checked = False

        gridPerguntas.DataSource = ObjParametros.Perguntas.OrderBy(Function(s) s.CodigoPergunta).ToArray
        gridPerguntas.DataBind()

        SessaoSalvaParametrosAno()
        LimparPerguntas()
    End Sub

    Protected Sub btnPergunta_Click(sender As Object, e As EventArgs)
        SessaoRecuperaParametros()
        Dim btn As Button = CType(sender, Button)
        Dim row As GridViewRow = CType(btn.NamingContainer, GridViewRow)
        Dim i As Integer = ObjParametros.Perguntas.FindIndex(Function(s) s.CodigoPergunta = row.Cells(1).Text)

        txtId.Text = ObjParametros.Perguntas(i).CodigoPergunta
        txtPergunta.Text = ObjParametros.Perguntas(i).Descricao
        txtPeso.Text = ObjParametros.Perguntas(i).PercPeso
    End Sub

    Protected Sub lnkLimparPergunta_Click(sender As Object, e As EventArgs) Handles lnkLimparPergunta.Click
        LimparPerguntas()
    End Sub

    Protected Sub LimparPerguntas()
        txtId.Text = String.Empty
        txtPergunta.Text = String.Empty
        txtPeso.Text = String.Empty
    End Sub

    Protected Sub imgExcluir_Click(sender As Object, e As System.Web.UI.ImageClickEventArgs)
        SessaoRecuperaParametros()
        Dim img As ImageButton = CType(sender, ImageButton)
        Dim row As GridViewRow = CType(img.NamingContainer, GridViewRow)

        Dim i As Integer = ObjParametros.Perguntas.FindIndex(Function(s) s.CodigoPergunta = row.Cells(1).Text)
        ObjParametros.Perguntas.RemoveAt(i)

        gridPerguntas.DataSource = ObjParametros.Perguntas.OrderBy(Function(s) s.CodigoPergunta).ToArray
        gridPerguntas.DataBind()

        SessaoSalvaParametrosAno()
    End Sub

    Protected Sub lnkGravarCultura_Click(sender As Object, e As EventArgs) Handles lnkGravarCultura.Click
        If ddlCultura.SelectedIndex = 0 Then
            MsgBox(Me.Page, "Selecione uma Cultura.")
            Exit Sub
        End If

        If Not IsNumeric(txtProdutividadeha.Text) OrElse CDec(txtProdutividadeha.Text) <= 0 Then
            MsgBox(Me.Page, "Informe um valor para Produtividade por Ha.")
            Exit Sub
        End If

        If Not IsNumeric(txtPrecoSaco.Text) OrElse CDec(txtPrecoSaco.Text) <= 0 Then
            MsgBox(Me.Page, "Informe um valor para o preço do Saco.")
            Exit Sub
        End If

        If Not IsNumeric(txtCustoTotalHa.Text) OrElse CDec(txtCustoTotalHa.Text) <= 0 Then
            MsgBox(Me.Page, "Informe um valor para o Custo Total por Ha.")
            Exit Sub
        End If

        If Not IsNumeric(txtCustoPortifolioHa.Text) OrElse CDec(txtCustoPortifolioHa.Text) <= 0 Then
            MsgBox(Me.Page, "Informe um valor para o Custo do Portifolio por Ha.")
            Exit Sub
        End If

        SessaoRecuperaParametrosAno()
        If ObjParametrosAno Is Nothing Then
            ObjParametrosAno = New [Lib].Negocio.ListParametrosAnaliseDeCredito(ddlAno.SelectedValue)
            ObjParametros = New [Lib].Negocio.ParametrosAnaliseDeCredito
        Else
            SessaoRecuperaParametros()
        End If

        Dim i As Integer = ObjParametros.Culturas.FindIndex(Function(s) s.CodigoSafra = ddlSafra.SelectedValue And s.CodigoCultura = ddlCultura.SelectedValue)

        Dim CultParam As [Lib].Negocio.ParametrosAnaliseDeCreditoCultura
        If i >= 0 Then
            CultParam = ObjParametros.Culturas(i)
        Else
            CultParam = New [Lib].Negocio.ParametrosAnaliseDeCreditoCultura(ObjParametros)
            CultParam.CodigoSafra = ddlSafra.SelectedValue
            CultParam.CodigoCultura = ddlCultura.SelectedValue
        End If

        CultParam.Produtividade = txtProdutividadeha.Text
        CultParam.PrecoSaco = txtPrecoSaco.Text
        CultParam.CustoTotalHa = txtCustoTotalHa.Text
        CultParam.CustoPortifolioHa = txtCustoPortifolioHa.Text

        If i < 0 Then ObjParametros.Culturas.Add(CultParam)

        gridCulturas.DataSource = ObjParametros.Culturas.ToArray
        gridCulturas.DataBind()

        SessaoSalvaParametrosAno()
        LimparCulturas()
    End Sub

    Protected Sub btnEditar_Click(sender As Object, e As EventArgs)
        SessaoRecuperaParametros()
        Dim btn As Button = CType(sender, Button)
        Dim row As GridViewRow = CType(btn.NamingContainer, GridViewRow)
        Dim i As Integer = ObjParametros.Culturas.FindIndex(Function(s) s.CodigoSafra = row.Cells(1).Text And s.CodigoCultura = row.Cells(2).Text)

        ddlSafra.SelectedValue = ObjParametros.Culturas(i).CodigoSafra
        ddlCultura.SelectedValue = ObjParametros.Culturas(i).CodigoCultura
        txtProdutividadeha.Text = ObjParametros.Culturas(i).Produtividade.ToString("N2")
        txtPrecoSaco.Text = ObjParametros.Culturas(i).PrecoSaco.ToString("N2")
        txtReceitaHa.Text = ObjParametros.Culturas(i).Receita.ToString("N2")
        txtCustoTotalHa.Text = ObjParametros.Culturas(i).CustoTotalHa.ToString("N2")
        txtCustoPortifolioHa.Text = ObjParametros.Culturas(i).CustoPortifolioHa.ToString("N2")
    End Sub

    Protected Sub LimparCulturas()
        ddlSafra.SelectedIndex = -1
        ddlCultura.SelectedIndex = 0

        txtProdutividadeha.Text = String.Empty
        txtPrecoSaco.Text = String.Empty
        txtReceitaHa.Text = String.Empty
        txtCustoTotalHa.Text = String.Empty
        txtCustoPortifolioHa.Text = String.Empty
    End Sub

    Protected Sub lnkLimparCultura_Click(sender As Object, e As EventArgs) Handles lnkLimparCultura.Click
        LimparCulturas()
    End Sub

    Protected Sub imgExcluirCultura_Click(sender As Object, e As System.Web.UI.ImageClickEventArgs)
        SessaoRecuperaParametros()
        Dim img As ImageButton = CType(sender, ImageButton)
        Dim row As GridViewRow = CType(img.NamingContainer, GridViewRow)

        Dim i As Integer = ObjParametros.Culturas.FindIndex(Function(s) s.CodigoSafra = gridCulturas.Rows(row.RowIndex).Cells(1).Text And s.CodigoCultura = gridCulturas.Rows(row.RowIndex).Cells(2).Text)
        ObjParametros.Culturas.RemoveAt(i)

        gridCulturas.DataSource = ObjParametros.Culturas.ToArray
        gridCulturas.DataBind()

        SessaoSalvaParametrosAno()
    End Sub

    Protected Sub imgExcluirPortifolio_Click(sender As Object, e As System.Web.UI.ImageClickEventArgs)
        SessaoRecuperaListaCulturaPortifolio()
        Dim img As ImageButton = CType(sender, ImageButton)
        Dim row As GridViewRow = CType(img.NamingContainer, GridViewRow)
        Dim i As Integer = ListCulturaPortifolios.FindIndex(Function(s) s.CodigoGrupoProduto = GridCulturaPortifolio.SelectedRow.Cells(1).Text)

        ListCulturaPortifolios.RemoveAt(i)

        SessaoSalvaListaCulturaPortifolio()

        GridCulturaPortifolio.DataSource = ListCulturaPortifolios
        GridCulturaPortifolio.DataBind()
    End Sub
End Class