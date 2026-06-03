Imports System.IO
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis
Imports OfficeOpenXml
Imports OfficeOpenXml.Style
Imports System.Drawing

Public Class OrdemDeProducao
    Inherits BasePage

    Dim ds As DataSet
    Dim objPrd As Produto
    Dim objOP As OrdemParaProducao
    Dim objOPXP As OrdemParaProducaoXProduto
    Dim objOPXC As OrdemParaProducaoXConsumo
    Dim objOPXI As OrdemParaProducaoXInsumo
    Dim objOPXCXLote As OrdemParaProducaoXConsumoXLote
    Dim objOPXE As OrdemParaProducaoXEspecificacao
    Dim objOPXEpi As OrdemParaProducaoXEPI
    Dim objOPXEmb As OrdemParaProducaoXEmbalagens
    Dim objOPXProcedimento As OrdemParaProducaoXProcedimento
    Dim bGerenciarOrdemDeProducao As Boolean
    Dim bAlterarOrdemDeProducao As Boolean
    Dim _codigoProdutoProducaoSelecionado As String



    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.ProducaoEstoque)
        If Not IsPostBack And IsConnect Then

            If Not UsuarioServidor.KeyCodeActive Then
                MsgBox(Me.Page, "Sistema com chave de licença expirada. Entre em contato com o suporte.", "~/Producao.aspx", eTitulo.Info)
                Exit Sub
            End If

            If Funcoes.VerificaPermissao("OrdemDeProducao", "ACESSAR") Then
                BuncarUnidadeDeNegocio()
                ddl.Carregar(ddlGrupoProdutoProducao, CarregarDDL.Tabela.GrupoProdutoXConsumo)
                ddl.Carregar(ddlEPI, CarregarDDL.Tabela.EPI, " ATIVO = 1")
                ddl.Carregar(ddlEmbalagem, CarregarDDL.Tabela.Embalagem, "", True)
                ddl.Carregar(ddlUnidadeDeComercializacao, CarregarDDL.Tabela.UnidadeDeMedida, "", True)
                ddl.Carregar(ddlProcedimento, CarregarDDL.Tabela.ProcedimentoParaProducao, " ATIVO = 1")

                ddl.Carregar(ddlGrupoProdutoConsumo, CarregarDDL.Tabela.GrupoProduto)
                ddl.Carregar(ddlGrupoProdutoInsumo, CarregarDDL.Tabela.GrupoProduto)

                Limpar()

            Else
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Producao.aspx")
                Exit Sub
            End If
        End If
    End Sub

    Private Sub SessaoSalvaOrdem()
        Session("objOrdemDeProducao" & HID.Value) = objOP
    End Sub

    Private Sub SessaoRecuperaOrdem()
        objOP = CType(Session("objOrdemDeProducao" & HID.Value), [Lib].Negocio.OrdemParaProducao)
    End Sub

    Public Overrides Sub Carregar(obj As IBaseEntity)
        If Session("objProdutoxPRD" & HID.Value) IsNot Nothing Then
            Dim objProduto As [Lib].Negocio.Produto = CType(obj, [Lib].Negocio.Produto)

            If Not objProduto.ControlarNumeroDoLote Then
                MsgBox(Me.Page, "Produto selecionado não está marcado Controlar Número do Lote.", eTitulo.Info)
                If gridProducao.Rows.Count = 0 Then
                    Limpar()
                End If
                Exit Sub
            End If

            If objProduto.ControlarEstoque Then
                ddlGrupoProdutoProducao.SelectedValue = objProduto.CodigoGrupo
                CarregarProdutoProducao()
                ddlProdutosProducao.SelectedValue = objProduto.Codigo
                Session.Remove("objProdutoxPRD" & HID.Value)
                UnidadeComercializacao()
            Else
                MsgBox(Me.Page, "Produto selecionado não está marcado para CONTROLAR ESTOQUE.", eTitulo.Info)
            End If
        ElseIf Session("objProdutoxCON" & HID.Value) IsNot Nothing Then

            Dim objProduto As [Lib].Negocio.Produto = CType(obj, [Lib].Negocio.Produto)

            ddlGrupoProdutoConsumo.SelectedValue = objProduto.CodigoGrupo
            ddl.Carregar(ddlProdutosConsumo, CarregarDDL.Tabela.Produto, " Situacao = 1 and Grupo ='" & ddlGrupoProdutoConsumo.SelectedValue & "'")

            With ddlProdutosConsumo
                ddlProdutosConsumo.SelectedIndex = .Items.IndexOf(.Items.FindByValue(objProduto.Codigo))
            End With

            Session.Remove("objProdutoxCON" & HID.Value)

        ElseIf Session("objProdutoxINSU" & HID.Value) IsNot Nothing Then

            Dim objProduto As [Lib].Negocio.Produto = CType(obj, [Lib].Negocio.Produto)

            ddlGrupoProdutoInsumo.SelectedValue = objProduto.CodigoGrupo
            ddl.Carregar(ddlProdutosInsumo, CarregarDDL.Tabela.Produto, " Situacao = 1 and Grupo ='" & ddlGrupoProdutoInsumo.SelectedValue & "'")

            With ddlProdutosInsumo
                ddlProdutosInsumo.SelectedIndex = .Items.IndexOf(.Items.FindByValue(objProduto.Codigo))
            End With

            If ddlProdutosInsumo.SelectedIndex > 0 Then
                ProdutosInsumo()
            End If

            Session.Remove("objProdutoxINSU" & HID.Value)

        ElseIf Session("objConsultaDeLote" & HID.Value) IsNot Nothing Then

            Dim objProduto As [Lib].Negocio.Produto = CType(obj, [Lib].Negocio.Produto)
            txtLoteAlt.Text = objProduto.NumeroDoLote
            txtValidadeAlt.Text = objProduto.ValidadeDoLote
            ProdutosConsumo()
            Session.Remove("objConsultaDeLote" & HID.Value)

        ElseIf Session("objConsultaXOrdem" & HID.Value) IsNot Nothing Then
            objOP = New OrdemParaProducao()
            objOP = CType(obj, [Lib].Negocio.OrdemParaProducao)

            Session.Remove("objConsultaXOrdem" & HID.Value)

            RecuperarOrdem()
        End If
    End Sub

    Private Sub BuncarUnidadeDeNegocio()
        ddl.Carregar(ddlUnidadeNegocio, CarregarDDL.Tabela.UnidadeDeNegocio)
        Funcoes.VerificaUnidadeDeNegocio(ddlUnidadeNegocio, ddlEmpresa)
    End Sub

    Protected Sub ddlUnidadeNegocio_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            If ddlUnidadeNegocio.SelectedIndex > 0 Then
                ddl.Carregar(ddlEmpresa, CarregarDDL.Tabela.Empresas, ddlUnidadeNegocio.SelectedValue.ToString)
            Else
                ddlEmpresa.Items.Clear()

                Limpar()
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Public Property CodigoProdutoProducaoSelecionado As String
        Get

            If String.IsNullOrEmpty(_codigoProdutoProducaoSelecionado) AndAlso gridProducao.SelectedRow IsNot Nothing Then
                _codigoProdutoProducaoSelecionado = gridProducao.SelectedRow.Cells(1).Text ' Ajuste o índice conforme necessário
            End If

            Return _codigoProdutoProducaoSelecionado
        End Get
        Set(value As String)
            _codigoProdutoProducaoSelecionado = value
        End Set
    End Property

    Protected Sub lnkBuscaProdutoProducao_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lnkBuscaProdutoProducao.Click
        Try
            If String.IsNullOrWhiteSpace(txtQuantidadeProducao.Text) OrElse CDec(txtQuantidadeProducao.Text) = 0 Then
                MsgBox(Me.Page, "Quantidade para Produção não foi informada.", eTitulo.Info)
                Exit Sub
            End If

            ucConsultaProduto.Limpar()
            Dim txtNome As TextBox = CType(ucConsultaProduto.FindControlRecursive("txtNome"), TextBox)
            Popup.ConsultaDeProduto(Me.Page, "objProdutoxPRD" & HID.Value, txtNome.ClientID, True)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub ddlGrupoProdutoProducao_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddlGrupoProdutoProducao.SelectedIndexChanged
        Try
            CarregarProdutoProducao()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub CarregarProdutoProducao()
        ddl.Carregar(ddlProdutosProducao, CarregarDDL.Tabela.ProdutoProducao, " P.Grupo = '" & ddlGrupoProdutoProducao.SelectedValue & "'")
    End Sub

    Protected Sub ddlProdutosProducao_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddlProdutosProducao.SelectedIndexChanged
        Try
            If ValidartProducao() Then

                UnidadeComercializacao()

            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub UnidadeComercializacao()

        Dim objProduto = New [Lib].Negocio.Produto(ddlProdutosProducao.SelectedValue)

        ddlUnidadeComercializacao.Items.Clear()

        ddlUnidadeComercializacao.Items.Add(New ListItem("", "0"))

        For Each un In objProduto.UnidadesDeComercializacao
            ddlUnidadeComercializacao.Items.Add(New ListItem(un.CodigoUnidade & " - " & un.FatorConversao.ToString("N4"), un.CodigoUnidade & " - " & un.FatorConversao.ToString("N4")))
        Next

        If ddlUnidadeComercializacao.Items.Count > 1 Then ddlUnidadeComercializacao.SelectedIndex = 1

        ddlUnidadeComercializacao.Enabled = True

    End Sub

    Private Function ValidartProducao() As Boolean


        If String.IsNullOrWhiteSpace(txtQuantidadeProducao.Text) OrElse CDec(txtQuantidadeProducao.Text) = 0 Then
            ddlProdutosProducao.SelectedIndex = 0
            MsgBox(Me.Page, "Quantidade para Produção não foi informada.", eTitulo.Info)
            Return False
        ElseIf ddlProdutosProducao.SelectedIndex = 0 Then
            ddlProdutosProducao.SelectedIndex = 0
            MsgBox(Me.Page, "Produto para Produção não foi informado.", eTitulo.Info)
            Return False
        ElseIf gridProducao.Rows.Count > 1 Then
            MsgBox(Me.Page, "Produto para Produção já foi informado.", eTitulo.Info)
            Return False
        ElseIf Not ddlUnidadeComercializacao.SelectedItem Is Nothing AndAlso ddlUnidadeComercializacao.SelectedItem.Text.Length = 0 Then
            MsgBox(Me.Page, "É necessário informar um peso para o produto!", eTitulo.Info)
            Return False
        Else
            Dim objProducao As DataTable = CType(Session("objProducao" & HID.Value), DataTable)

            If objProducao.Select($"Produto = '{ddlProdutosProducao.SelectedValue}'").Length > 0 Then
                MsgBox(Me.Page, "Produto já cadastrado! É necessário informar um outro produto ou remover o existente!", eTitulo.Info)
                Return False
            End If
        End If
        Return True
    End Function

    Protected Sub imgAdicionarDadosParaProducao_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)
        Try
            If ValidartProducao() Then

                Dim objProduto = New [Lib].Negocio.Produto(ddlProdutosProducao.SelectedValue)

                If Not objProduto.ControlarNumeroDoLote Then
                    MsgBox(Me.Page, "Produto selecionado não está marcado Controlar Número do Lote.", eTitulo.Info)
                    Limpar()
                    Exit Sub
                End If

                If objProduto.ControlarEstoque Then
                    ProdutoProducao()
                Else
                    MsgBox(Me.Page, "Produto selecionado não está marcado para CONTROLAR ESTOQUE.", eTitulo.Info)
                End If

            End If

        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub ProdutoProducao()

        Dim objProducao = New ListProdutoXConsumos()

        objProducao.CarregarProduto(ddlProdutosProducao.SelectedValue)

        If objProducao.Count = 0 Then
            ddlProdutosProducao.SelectedIndex = 0
            MsgBox(Me.Page, "Produto para Produção não foi encontrado.", eTitulo.Info)
            Exit Sub
        End If

        Dim produtoXLote = New ProdutoXSequenciaDeLote(objProducao(0).Produto.Codigo, CType(txtDataProducao.Text, DateTime).Year)

        If String.IsNullOrWhiteSpace(produtoXLote.SequenciaDoProduto) Then
            ddlProdutosProducao.SelectedIndex = 0
            MsgBox(Me.Page, "Sequência do Lote do Produto não foi encontrada.", eTitulo.Info)
            Exit Sub
        End If

        Dim CodigoUnidadeComercializacao As String = ""
        Dim FatorConversao As Decimal = 0

        If ddlUnidadeComercializacao.SelectedItem.Text.Length > 0 Then
            CodigoUnidadeComercializacao = ddlUnidadeComercializacao.SelectedItem.Text.Split("-")(0).Trim
            FatorConversao = ddlUnidadeComercializacao.SelectedItem.Text.Split("-")(1)
        End If

        Dim primeira = True

        For Each row In objProducao

            If primeira Then
                Dim drItem As DataRow = CType(Session("objProducao" & HID.Value), DataTable).NewRow()

                drItem("Produto") = row.Produto.Codigo
                drItem("NomeProduto") = row.Produto.Nome
                drItem("Quantidade") = txtQuantidadeProducao.Text
                drItem("CodigoUnidadeComercializacao") = CodigoUnidadeComercializacao
                drItem("FatorConversao") = FatorConversao
                drItem("AjusteProducao") = 0
                CType(Session("objProducao" & HID.Value), DataTable).Rows.Add(drItem)

                primeira = False
            End If

            Dim drItemC As DataRow = CType(Session("objConsumo" & HID.Value), DataTable).NewRow()


            drItemC("CodigoProdutoProducao") = row.Produto.Codigo
            drItemC("Produto") = row.ProdutoConsumo.Codigo
            drItemC("NomeProduto") = row.ProdutoConsumo.Nome

            If row.ProdutoConsumo.ControlarEstoque Then
                drItemC("NumeroDoLote") = "INFORMAR"
            Else
                drItemC("NumeroDoLote") = ""
            End If

            Dim qtde = Math.Round((row.Percentual * CDec(txtQuantidadeProducao.Text)) / 100, 4, MidpointRounding.AwayFromZero)

            If row.ProdutoConsumo.Unidade = "KG" AndAlso FatorConversao > 0 Then
                qtde = (qtde * FatorConversao)
            End If

            drItemC("Quantidade") = (qtde).ToString("N4")
            drItemC("Percentual") = row.Percentual.ToString("N4")
            drItemC("Sinal") = "+"
            drItemC("AjusteConsumo") = 0

            CType(Session("objConsumo" & HID.Value), DataTable).Rows.Add(drItemC)

        Next

        gridProducao.DataSource = CType(Session("objProducao" & HID.Value), DataTable)
        gridProducao.DataBind()

        If gridProducao.Rows.Count > 0 Then
            Dim ultimaLinha As Integer = gridProducao.Rows.Count - 1 ' Pega o índice da última linha
            gridProducao.SelectedIndex = ultimaLinha ' Seleciona a última linha
            gridProducao_SelectedIndexChanged(Nothing, Nothing) ' Dispara o evento manualmente
        End If

        gridProducao.Columns(4).Visible = True
        gridProducao.Columns(6).Visible = False

        divConsumo.Style.Value = "width: 98%;"

        gridConsumo.DataSource = dtgridConsumo()
        gridConsumo.DataBind()

        'gridConsumo.Columns(6).Visible = False
        'gridConsumo.Columns(7).Visible = False

        gridConsumo.Columns(6).Visible = True
        gridConsumo.Columns(7).Visible = True
        divConsumoLote.Visible = False

        DisableSelecionarLote()

        Dim objProducaoI = New ListProdutoXInsumos()

        objProducaoI.CarregarProduto(ddlProdutosProducao.SelectedValue)

        Dim dataInsumo As Date = Now()
        Dim dataInsumoAnt As Date = dataInsumo.AddDays(-1)

        For Each row In objProducaoI

            Dim drItemI As DataRow = CType(Session("objInsumo" & HID.Value), DataTable).NewRow()

            drItemI("CodigoProdutoProducao") = row.Produto.Codigo
            drItemI("Produto") = row.ProdutoInsumo.Codigo
            drItemI("NomeProduto") = row.ProdutoInsumo.Nome

            'drItemI("Base") = CInt(row.Base)
            drItemI("Base") = row.Base
            drItemI("Quantidade") = 0
            drItemI("Estoque") = CDec(0)

            Dim qInsumo = New OrdemParaProducaoXInsumo()

            drItemI("Estoque") = CDec(qInsumo.buscarEstoqueProduto(ddlEmpresa.SelectedValue.Split("-")(0), ddlEmpresa.SelectedValue.Split("-")(1), row.ProdutoInsumo.Codigo, dataInsumoAnt.ToString("yyyy-MM-dd"), dataInsumo.ToString("yyyy-MM-dd")))

            CType(Session("objInsumo" & HID.Value), DataTable).Rows.Add(drItemI)

        Next

        gridInsumo.DataSource = dtgridInsumo()
        gridInsumo.DataBind()

        Dim j As Integer = 0
        While j < gridInsumo.Rows.Count
            If CInt(gridInsumo.Rows(j).Cells(4).Text) > 0 Then
                CType(gridInsumo.Rows(j).FindControl("txtQuantidadeInsumo"), TextBox).Enabled = True
            Else
                CType(gridInsumo.Rows(j).FindControl("txtQuantidadeInsumo"), TextBox).Enabled = False
            End If

            j += 1
        End While

        For Each eP In objProducao(0).Produto.ProdutoXEspecificacao.Where(Function(s) s.Ativo = True)
            Dim drEDP As DataRow = CType(Session("objEspecificacaoDoProduto" & HID.Value), DataTable).NewRow()

            drEDP("CodigoProdutoProducao") = eP.CodigoProduto
            drEDP("Codigo") = eP.CodigoEspecificacao
            drEDP("Descricao") = eP.EspecificacaoDoProduto.Descricao
            drEDP("FaixaInicial") = eP.FaixaInicial
            drEDP("FaixaFinal") = eP.FaixaFinal
            drEDP("Resultado") = 0

            CType(Session("objEspecificacaoDoProduto" & HID.Value), DataTable).Rows.Add(drEDP)
        Next

        Dim dtgridEspecificacao As DataTable = dtgridEspecificacaoDoProduto()

        If dtgridEspecificacao.Rows.Count > 0 Then
            gridEspecificacaoDoProduto.DataSource = dtgridEspecificacao
            gridEspecificacaoDoProduto.DataBind()

            gridEspecificacaoDoProduto.Columns(4).Visible = False
        End If


        For Each ePI In objProducao(0).Produto.ProdutoXEPI.Where(Function(s) s.Ativo = True)
            Dim drEPI As DataRow = CType(Session("objEPI" & HID.Value), DataTable).NewRow()

            drEPI("Codigo") = ePI.CodigoEPI
            drEPI("Descricao") = ePI.Descricao

            CType(Session("objEPI" & HID.Value), DataTable).Rows.Add(drEPI)
        Next

        If CType(Session("objEPI" & HID.Value), DataTable).Rows.Count > 0 Then
            gridEPI.DataSource = CType(Session("objEPI" & HID.Value), DataTable)
            gridEPI.DataBind()
        End If


        For Each eProc In objProducao(0).Produto.ProdutoXProcedimento.Where(Function(s) s.Ativo = True)
            Dim drProc As DataRow = CType(Session("objProcedimento" & HID.Value), DataTable).NewRow()

            drProc("Codigo") = eProc.CodigoProcedimento
            drProc("Descricao") = eProc.Descricao

            CType(Session("objProcedimento" & HID.Value), DataTable).Rows.Add(drProc)
        Next

        If CType(Session("objProcedimento" & HID.Value), DataTable).Rows.Count > 0 Then
            gridProcedimento.DataSource = CType(Session("objProcedimento" & HID.Value), DataTable)
            gridProcedimento.DataBind()
        End If


        ddlGrupoProdutoProducao.SelectedIndex = 0
        ddlProdutosProducao.Items.Clear()
        txtQuantidadeProducao.Text = String.Empty

        ddlUnidadeComercializacao.Items.Clear()
        ddlUnidadeComercializacao.Enabled = False

        If CType(Session("objConsumo" & HID.Value), DataTable).Rows.Count > 0 And txtSequencia.Enabled = True Then
            lnkNovo.Parent.Visible = True
            lnkConsultar.Parent.Visible = False
            verNumerador()
        ElseIf CType(Session("objConsumo" & HID.Value), DataTable).Rows.Count = 0 Then
            lnkNovo.Parent.Visible = False
            lnkConsultar.Parent.Visible = True

            'txtSequencia.Enabled = True
            'txtSequencia.Text = String.Empty

            MsgBox(Me.Page, "Produto para Consumo não foi encontrado.", eTitulo.Info)

        End If

    End Sub

    Private Function dtgridConsumo() As DataTable

        Dim dtConsumo As DataTable = CType(Session("objConsumo" & HID.Value), DataTable)
        Dim dtConsumoFiltrado As DataTable = dtConsumo.Clone() ' Clonar estrutura da tabela

        For Each row As DataRow In dtConsumo.Select("CodigoProdutoProducao = '" & CodigoProdutoProducaoSelecionado & "'")
            dtConsumoFiltrado.ImportRow(row)
        Next

        Return dtConsumoFiltrado

    End Function

    Private Function dtgridInsumo() As DataTable

        Dim dtConsumo As DataTable = CType(Session("objInsumo" & HID.Value), DataTable)
        Dim dtInsumoFiltrado As DataTable = dtConsumo.Clone() ' Clonar estrutura da tabela

        For Each row As DataRow In dtConsumo.Select("CodigoProdutoProducao = '" & CodigoProdutoProducaoSelecionado & "'")
            dtInsumoFiltrado.ImportRow(row)
        Next

        Return dtInsumoFiltrado

    End Function

    Private Function dtgridEspecificacaoDoProduto() As DataTable

        Dim dtEspecificacaoDoProduto As DataTable = CType(Session("objEspecificacaoDoProduto" & HID.Value), DataTable)
        Dim dtEspecificacaoDoProdutoFiltrado As DataTable = dtEspecificacaoDoProduto.Clone() ' Clonar estrutura da tabela

        For Each row As DataRow In dtEspecificacaoDoProduto.Select("CodigoProdutoProducao = '" & CodigoProdutoProducaoSelecionado & "'")
            dtEspecificacaoDoProdutoFiltrado.ImportRow(row)
        Next

        Return dtEspecificacaoDoProdutoFiltrado

    End Function

    Private Function dtgridLoteDeFornecedor() As DataTable

        Dim dtLoteDeFornecedor As DataTable = CType(Session("objLoteFornecedor" & HID.Value), DataTable)
        Dim dtdtLoteDeFornecedorFiltrado As DataTable = dtLoteDeFornecedor.Clone() ' Clonar estrutura da tabela

        For Each row As DataRow In dtLoteDeFornecedor.Select("CodigoProdutoProducao = '" & CodigoProdutoProducaoSelecionado & "'")
            dtdtLoteDeFornecedorFiltrado.ImportRow(row)
        Next

        Return dtdtLoteDeFornecedorFiltrado

    End Function

    Private Function dtgridConsumoAlt() As DataTable

        Dim dtConsumo As DataTable = CType(Session("objConsumoAlt" & HID.Value), DataTable)
        Dim dtConsumoFiltrado As DataTable = dtConsumo.Clone() ' Clonar estrutura da tabela

        For Each row As DataRow In dtConsumo.Select("CodigoProdutoProducao = '" & CodigoProdutoProducaoSelecionado & "'")
            dtConsumoFiltrado.ImportRow(row)
        Next

        Return dtConsumoFiltrado

    End Function

    Private Function dtgridInsumoAlt() As DataTable

        Dim dtInsumo As DataTable = CType(Session("objInsumoAlt" & HID.Value), DataTable)
        Dim dtInsumoFiltrado As DataTable = dtInsumo.Clone() ' Clonar estrutura da tabela

        For Each row As DataRow In dtInsumo.Select("CodigoProdutoProducao = '" & CodigoProdutoProducaoSelecionado & "'")
            dtInsumoFiltrado.ImportRow(row)
        Next

        Return dtInsumoFiltrado

    End Function

    Private Sub gravarobjConsumoAlt(ByVal linhaParaAlterar As DataRow)

        Dim dtConsumo As DataTable = CType(Session("objConsumoAlt" & HID.Value), DataTable)

        For Each row As DataRow In dtConsumo.Rows
            If row.Item("CodigoProdutoProducao") = linhaParaAlterar.Item("CodigoProdutoProducao") AndAlso
                row.Item("Produto") = linhaParaAlterar.Item("Produto") Then

                ' Atualiza os valores da linha existente com os da nova
                row = linhaParaAlterar

            End If
        Next

        Session("objConsumoAlt" & HID.Value) = dtConsumo
    End Sub

    Private Sub gravarobjInsumoAlt(ByVal linhaParaAlterar As DataRow)

        Dim dtInsumo As DataTable = CType(Session("objInsumoAlt" & HID.Value), DataTable)

        For Each row As DataRow In dtInsumo.Rows
            If row.Item("CodigoProdutoProducao") = linhaParaAlterar.Item("CodigoProdutoProducao") AndAlso
                row.Item("Produto") = linhaParaAlterar.Item("Produto") Then

                ' Atualiza os valores da linha existente com os da nova
                row = linhaParaAlterar

            End If
        Next

        Session("objInsumoAlt" & HID.Value) = dtInsumo
    End Sub

    Protected Sub imgRemoverProducao_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)
        Try
            SessaoRecuperaOrdem()

            If objOP.IUD = "U" Then
                MsgBox(Me.Page, "Item não pode ser removido, se desejar faça a exclusao do mesmo.", eTitulo.Info)
                Exit Sub
            End If

            Limpar()

        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try

    End Sub

    Protected Sub lnkAjustaProducao_Click(sender As Object, e As EventArgs) Handles lnkAjustaProducao.Click
        Try
            Dim j As Integer = 0
            While j < gridProducao.Rows.Count
                CType(gridProducao.Rows(j).FindControl("txtAjusteProducao"), TextBox).Enabled = True

                j += 1
            End While

            divConsumoLote.Visible = False
            lnkAjustaProducao.Parent.Visible = False
            lnkConfirmaProducao.Parent.Visible = True

        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkConfirmaProducao_Click(sender As Object, e As EventArgs) Handles lnkConfirmaProducao.Click
        Try

            If Funcoes.VerificaPermissao("OrdemDeProducao", "LIBERAR") Then

                SessaoRecuperaOrdem()

                For Each row In gridProducao.Rows

                    For Each prodProducao In objOP.ProdutosdaProducao
                        If prodProducao.CodigoProduto = row.Cells(1).Text Then

                            Dim tolerancia As Decimal = (prodProducao.Quantidade * 20) / 100
                            Dim quantidadeDeAjuste As Decimal = CDec(CType(gridProducao.Rows(row.RowIndex).FindControl("txtAjusteProducao"), TextBox).Text)

                            If quantidadeDeAjuste > 0 AndAlso quantidadeDeAjuste > (prodProducao.Quantidade + tolerancia) Then
                                MsgBox(Me.Page, "Ajuste não pode ser superior a 20%, da quantidade atual!", eTitulo.Info)
                                If Not Session("ssNomeUsuario") = "DOUGLAS" Then
                                    Exit Sub
                                End If
                            End If

                            If quantidadeDeAjuste > 0 AndAlso quantidadeDeAjuste < (prodProducao.Quantidade - tolerancia) Then
                                MsgBox(Me.Page, "Ajuste não pode ser inferior a 20%, da quantidade atual!", eTitulo.Info)
                                If Not Session("ssNomeUsuario") = "DOUGLAS" Then
                                    Exit Sub
                                End If
                            End If

                            prodProducao.QuantidadeDeAjuste = quantidadeDeAjuste

                            For Each drProduto As DataRow In CType(Session("objProducao" & HID.Value), DataTable).Rows

                                If prodProducao.CodigoProduto = drProduto("Produto") Then
                                    drProduto("AjusteProducao") = prodProducao.QuantidadeDeAjuste
                                    Exit For
                                End If

                            Next

                            Dim bRecalcularConsumo As Boolean = False
                            If bRecalcularConsumo Then

                                Dim dtConsumo As DataTable = CType(Session("objConsumo" & HID.Value), DataTable)
                                Dim CodigoUnidadeComercializacao As String = prodProducao.CodigoUnidadeComercializacao
                                Dim FatorConversao As Decimal = prodProducao.FatorConversao

                                For Each rowConsumo As DataRow In dtConsumo.Select("CodigoProdutoProducao = '" & CodigoProdutoProducaoSelecionado & "'")

                                    Dim qtde As Decimal

                                    If quantidadeDeAjuste = 0 Then
                                        qtde = Math.Round((rowConsumo("Percentual") * CDec(prodProducao.Quantidade)) / 100, 4, MidpointRounding.AwayFromZero)
                                    Else
                                        qtde = Math.Round((rowConsumo("Percentual") * CDec(prodProducao.QuantidadeDeAjuste)) / 100, 4, MidpointRounding.AwayFromZero)
                                    End If

                                    If CodigoUnidadeComercializacao = "KG" AndAlso FatorConversao > 0 Then
                                        qtde = (qtde * prodProducao.FatorConversao)
                                    End If

                                    rowConsumo("Quantidade") = qtde

                                    Session("objAjustarLote" & rowConsumo("Produto") & HID.Value) = False

                                Next

                                For Each pC In objOP.ItensDeConsumo
                                    For Each drItemC As DataRow In CType(Session("objConsumo" & HID.Value), DataTable).Rows
                                        If drItemC("CodigoProdutoProducao") = CodigoProdutoProducaoSelecionado And drItemC("Produto") = pC.CodigoProduto Then
                                            pC.Quantidade = drItemC("Quantidade")
                                        End If
                                    Next
                                Next

                            End If

                            gridConsumo.DataSource = dtgridConsumo()
                            gridConsumo.DataBind()

                            CType(gridProducao.Rows(row.RowIndex).FindControl("txtAjusteProducao"), TextBox).Enabled = False

                            Exit For

                        End If
                    Next
                Next

                SessaoSalvaOrdem()

                'lnkAjustaConsumo.Parent.Visible = False
                'lnkConfirmaConsumo.Parent.Visible = False
                'divAjusteConsumo.Visible = False

                lnkAjustaProducao.Parent.Visible = True
                lnkConfirmaProducao.Parent.Visible = False
                lnkAtualizar.Parent.Visible = True

                DisableSelecionarLote()

                Dim i As Integer = 0
                While i < gridConsumo.Rows.Count
                    If gridConsumo.Rows(i).Cells(3).Text = "INFORMADO" Then
                        gridConsumo.Rows(i).Cells(3).Text = "INFORMAR"
                    End If

                    i += 1
                End While

            End If

        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAjustaConsumo_Click(sender As Object, e As EventArgs) Handles lnkAjustaConsumo.Click
        Try
            SessaoRecuperaOrdem()

            For Each pC In objOP.ItensDeConsumo
                For Each drItemC As DataRow In CType(Session("objConsumo" & HID.Value), DataTable).Rows
                    If drItemC("CodigoProdutoProducao") = CodigoProdutoProducaoSelecionado And drItemC("Produto") = pC.CodigoProduto Then
                        drItemC("AjusteConsumo") = pC.QuantidadeDeAjuste
                    End If
                Next
            Next

            gridConsumo.DataSource = dtgridConsumo()
            gridConsumo.DataBind()

            If objOP.IUD = "U" Then
                Dim i As Integer = 0
                While i < gridConsumo.Rows.Count
                    If gridConsumo.Rows(i).Cells(3).Text = "INFORMADO" Or gridConsumo.Rows(i).Cells(3).Text = "INFORMAR" Then
                        CType(gridConsumo.Rows(i).FindControl("imgSelecionaLote"), ImageButton).ImageUrl = "~/images/select.jpg"
                        CType(gridConsumo.Rows(i).FindControl("imgSelecionaLote"), ImageButton).ToolTip = ""
                    Else
                        CType(gridConsumo.Rows(i).FindControl("imgSelecionaLote"), ImageButton).ImageUrl = "~/images/important.png"
                        CType(gridConsumo.Rows(i).FindControl("imgSelecionaLote"), ImageButton).ToolTip = "Não informa Lote"
                    End If

                    CType(gridConsumo.Rows(i).FindControl("imgSelecionaLote"), ImageButton).Enabled = False
                    CType(gridConsumo.Rows(i).FindControl("DdlSinal"), DropDownList).Enabled = True
                    CType(gridConsumo.Rows(i).FindControl("txtAjusteConsumo"), TextBox).Enabled = True

                    For Each drItemC As DataRow In CType(Session("objConsumo" & HID.Value), DataTable).Rows
                        If drItemC("CodigoProdutoProducao") = CodigoProdutoProducaoSelecionado And drItemC("Produto") = gridConsumo.Rows(i).Cells(1).Text Then
                            CType(gridConsumo.Rows(i).FindControl("DdlSinal"), DropDownList).SelectedValue = drItemC("Sinal")
                        End If
                    Next

                    i += 1
                End While
            End If

            lnkAjustaConsumo.Parent.Visible = False
            lnkConfirmaConsumo.Parent.Visible = True
            divConsumoLote.Visible = False

            Dim j As Integer = 0
            While j < gridLoteDeFornecedor.Rows.Count

                CType(gridLoteDeFornecedor.Rows(j).FindControl("txtConsumoLote"), TextBox).Enabled = True

                j += 1
            End While

            imgConfirmar.Visible = False

        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try

    End Sub

    Protected Sub lnkConfirmaConsumo_Click(sender As Object, e As EventArgs) Handles lnkConfirmaConsumo.Click
        Try
            SessaoRecuperaOrdem()

            For Each drItemC As DataRow In CType(Session("objConsumo" & HID.Value), DataTable).Rows
                If CDec(drItemC("AjusteConsumo")) > 0 Then
                    For Each drItemLote As DataRow In CType(Session("objLoteFornecedor" & HID.Value), DataTable).Rows
                        If drItemLote("CodigoProdutoProducao") = CodigoProdutoProducaoSelecionado And drItemLote("Produto") = drItemC("Produto") Then

                            'Usamos o OldConsumo, pq há momentos que o usuario clica varias vezes no ajuste, e dai o mesmo vai adicionando duplicado
                            If drItemC("Sinal") = "+" Then
                                drItemLote("Consumo") = drItemLote("OldConsumo") + drItemC("AjusteConsumo")
                            Else
                                drItemLote("Consumo") = drItemLote("OldConsumo") - drItemC("AjusteConsumo")
                            End If

                            Exit For
                        End If
                    Next
                End If
            Next

            For Each row In gridConsumo.Rows
                For Each pC In objOP.ItensDeConsumo
                    If pC.CodigoProdutoProducao = CodigoProdutoProducaoSelecionado And pC.CodigoProduto = row.Cells(1).Text Then

                        pC.QuantidadeDeAjuste = CDec(CType(gridConsumo.Rows(row.RowIndex).FindControl("txtAjusteConsumo"), TextBox).Text)
                        pC.Sinal = CType(gridConsumo.Rows(row.RowIndex).FindControl("DdlSinal"), DropDownList).SelectedValue

                        If CDec(CType(gridConsumo.Rows(row.RowIndex).FindControl("txtAjusteConsumo"), TextBox).Text) > 0 Then

                            If row.Cells(3).Text = "INFORMADO" Then

                                Dim qtdeConsumo As Decimal = 0
                                If pC.Sinal = "+" Then
                                    qtdeConsumo = pC.Quantidade + pC.QuantidadeDeAjuste
                                ElseIf pC.Sinal = "-" Then
                                    qtdeConsumo = pC.Quantidade - pC.QuantidadeDeAjuste
                                End If

                                Dim qtdeAjuste As Decimal = pC.QuantidadeDeAjuste

                                Session("objAjustarLote" & pC.CodigoProduto & HID.Value) = False
                                AjustarConsumoLotes(pC.CodigoProduto, pC.Produto.Nome, qtdeAjuste)

                                'Dim pLote = New OrdemParaProducaoXConsumo()

                                'Dim ds As New DataSet
                                'ds = pLote.buscarLoteDeFornecedor(objOP.CodigoEmpresa, objOP.EnderecoEmpresa, "'" & pC.CodigoProduto & "'")

                                'If ds.Tables(0).Rows.Count = 0 Then
                                '    MsgBox(Me.Page, "Não foi encontrado o lote do Produto " & pC.CodigoProduto & "-" & pC.Produto.Nome, eTitulo.Info)
                                '    Exit Sub
                                'End If

                                'For Each dr In ds.Tables(0).Rows

                                '    Dim temLote As Boolean = False

                                '    For Each drItemLote As DataRow In CType(Session("objLoteFornecedor" & HID.Value), DataTable).Rows
                                '        If drItemLote("CodigoProdutoProducao") = CodigoProdutoProducaoSelecionado And drItemLote("Produto") = dr("Produto") AndAlso drItemLote("Lote") = dr("Lote") Then
                                '            If qtdeAjuste > 0 Then
                                '                drItemLote("Quantidade") = dr("Quantidade")

                                '                drItemLote("Consumo") = pC.Quantidade

                                '                If pC.Sinal = "+" Then
                                '                    drItemLote("Consumo") += pC.QuantidadeDeAjuste
                                '                ElseIf pC.Sinal = "-" Then
                                '                    drItemLote("Consumo") -= pC.QuantidadeDeAjuste
                                '                End If
                                '                qtdeAjuste -= pC.QuantidadeDeAjuste
                                '            Else
                                '                drItemLote("Quantidade") = dr("Quantidade")
                                '                'drItemLote("Consumo") = dr("Consumo")
                                '            End If

                                '            temLote = True

                                '            Exit For
                                '        End If
                                '    Next

                                '    If Not temLote Then
                                '        Dim drItem As DataRow = CType(Session("objLoteFornecedor" & HID.Value), DataTable).NewRow()

                                '        drItem("CodigoProdutoProducao") = CodigoProdutoProducaoSelecionado
                                '        drItem("Produto") = dr("Produto")
                                '        drItem("Lote") = dr("Lote")
                                '        drItem("Fabricado") = CDate(dr("Validade")).ToString("dd/MM/yyyy")
                                '        drItem("Validade") = CDate(dr("Validade")).ToString("dd/MM/yyyy")
                                '        drItem("Quantidade") = dr("Quantidade")
                                '        drItem("Consumo") = dr("Consumo")

                                '        CType(Session("objLoteFornecedor" & HID.Value), DataTable).Rows.Add(drItem)
                                '    End If
                                'Next

                            End If
                        End If
                    End If
                Next
            Next

            Dim totalProduto As Decimal = 0

            For Each pC In objOP.ItensDeConsumo
                For Each drItemC As DataRow In CType(Session("objConsumo" & HID.Value), DataTable).Rows
                    If drItemC("CodigoProdutoProducao") = CodigoProdutoProducaoSelecionado And drItemC("Produto") = pC.CodigoProduto Then
                        drItemC("AjusteConsumo") = pC.QuantidadeDeAjuste
                        drItemC("Sinal") = pC.Sinal

                        If pC.Sinal = "+" Then
                            totalProduto += (pC.Quantidade + pC.QuantidadeDeAjuste)
                        ElseIf pC.Sinal = "-" Then
                            totalProduto += (pC.Quantidade - pC.QuantidadeDeAjuste)
                        End If
                    End If

                    If drItemC("NumeroDoLote") = "INFORMADO" Then
                        drItemC("NumeroDoLote") = "INFORMAR"
                    End If
                Next
            Next

            If objOP.ProdutosdaProducao.Count = 1 AndAlso Not Left(objOP.CodigoEmpresa, 8) = "05272759" Then

                For Each produto As OrdemParaProducaoXProduto In objOP.ProdutosdaProducao

                    'objOP.Quantidade = objOP.ItensDeConsumo.Sum(Function(s) s.Quantidade + s.QuantidadeDeAjuste)
                    'Só atualizar a Produção se o Peso do Item for ZERO - Furlan - 23/09/2024
                    If produto.FatorConversao = 1 Then produto.Quantidade = totalProduto

                    CType(Session("objProducao" & HID.Value), DataTable).Rows(0).Item("Quantidade") = produto.Quantidade

                Next

            End If

            gridProducao.DataSource = CType(Session("objProducao" & HID.Value), DataTable)
            gridProducao.DataBind()

            gridConsumo.DataSource = dtgridConsumo()
            gridConsumo.DataBind()

            Dim i As Integer = 0
            While i < gridConsumo.Rows.Count
                CType(gridConsumo.Rows(i).FindControl("imgSelecionaLote"), ImageButton).ToolTip = ""

                If gridConsumo.Rows(i).Cells(3).Text = "INFORMADO" Or gridConsumo.Rows(i).Cells(3).Text = "INFORMAR" Then
                    CType(gridConsumo.Rows(i).FindControl("imgSelecionaLote"), ImageButton).ImageUrl = "~/images/select.jpg"
                    CType(gridConsumo.Rows(i).FindControl("imgSelecionaLote"), ImageButton).ToolTip = ""
                    CType(gridConsumo.Rows(i).FindControl("imgSelecionaLote"), ImageButton).Enabled = True
                Else
                    CType(gridConsumo.Rows(i).FindControl("imgSelecionaLote"), ImageButton).ImageUrl = "~/images/important.png"
                    CType(gridConsumo.Rows(i).FindControl("imgSelecionaLote"), ImageButton).ToolTip = "Não informa Lote"
                    CType(gridConsumo.Rows(i).FindControl("imgSelecionaLote"), ImageButton).Enabled = False
                End If

                For Each drItemC As DataRow In CType(Session("objConsumo" & HID.Value), DataTable).Rows
                    If drItemC("CodigoProdutoProducao") = CodigoProdutoProducaoSelecionado And drItemC("Produto") = gridConsumo.Rows(i).Cells(1).Text Then
                        CType(gridConsumo.Rows(i).FindControl("DdlSinal"), DropDownList).SelectedValue = drItemC("Sinal")
                    End If
                Next

                i += 1
            End While

            'Dim dv As DataView = New DataView(CType(Session("objLoteFornecedor" & HID.Value), DataTable))

            'If divConsumoLote.Visible = False Then
            '    dv.RowFilter = "CodigoProdutoProducao = '" & CodigoProdutoProducaoSelecionado & "' AND Produto = '" & gridConsumo.Rows(0).Cells(1).Text & "'"
            'Else
            '    dv.RowFilter = "CodigoProdutoProducao = '" & CodigoProdutoProducaoSelecionado & "' AND Produto = '" & gridConsumo.Rows(Session("objLinha" & HID.Value)).Cells(1).Text & "'"
            'End If

            'gridLoteDeFornecedor.DataSource = dv.ToTable()
            'gridLoteDeFornecedor.DataBind()

            i = 0
            While i < gridLoteDeFornecedor.Rows.Count
                CType(gridLoteDeFornecedor.Rows(i).FindControl("txtConsumoLote"), TextBox).Enabled = False

                i += 1
            End While

            imgConfirmar.Visible = False

            SessaoSalvaOrdem()

            lnkConfirmaConsumo.Parent.Visible = False
            lnkAjustaConsumo.Parent.Visible = True

        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub imgSelecionaLote_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)
        Dim imgSelecionaLote As ImageButton = CType(sender, ImageButton)
        Dim row As GridViewRow = CType(imgSelecionaLote.NamingContainer, GridViewRow)

        Session("objLinha" & HID.Value) = row.RowIndex

        SessaoRecuperaOrdem()

        Dim produto As String = row.Cells(1).Text
        Dim nomeProduto As String = row.Cells(2).Text

        Dim qtdeConsumo As Decimal = CDec(row.Cells(4).Text)
        Dim qtdeAjuste As Decimal = 0

        If CType(gridConsumo.Rows(row.RowIndex).FindControl("DdlSinal"), DropDownList).SelectedValue = "+" Then
            qtdeAjuste = CDec(CType(gridConsumo.Rows(row.RowIndex).FindControl("txtAjusteConsumo"), TextBox).Text)
            qtdeConsumo += qtdeAjuste
        Else
            qtdeAjuste = CDec(CType(gridConsumo.Rows(row.RowIndex).FindControl("txtAjusteConsumo"), TextBox).Text)
            qtdeConsumo -= qtdeAjuste
        End If

        If objOP.IUD = "U" Then

            AjustarConsumoLotes(produto, nomeProduto, qtdeConsumo)

        Else

            If gridConsumo.Rows(row.RowIndex).Cells(3).Text = "INFORMAR" Then
                Dim pLote = New OrdemParaProducaoXConsumo()

                Dim ds As New DataSet
                ds = pLote.buscarLoteDeFornecedor(ddlEmpresa.SelectedValue.Split("-")(0), ddlEmpresa.SelectedValue.Split("-")(1), "'" & row.Cells(1).Text & "'")

                If ds.Tables(0).Rows.Count > 0 Then
                    If ds.Tables(0).Compute("SUM(Quantidade)", "") < qtdeConsumo Then
                        MsgBox(Me.Page, "Produto " & row.Cells(1).Text & "-" & row.Cells(2).Text & "não tem saldo suficiente para Consumo. Saldo " & ds.Tables(0).Compute("SUM(Quantidade)", "").ToString, eTitulo.Info)
                        Exit Sub
                    End If

                    For Each dr In ds.Tables(0).Rows
                        If objOP.IUD = "U" Then
                            For Each drItemLote As DataRow In CType(Session("objLoteFornecedor" & HID.Value), DataTable).Rows
                                If drItemLote("CodigoProdutoProducao") = CodigoProdutoProducaoSelecionado And drItemLote("Produto") = dr("Produto") AndAlso drItemLote("Lote") = dr("Lote") AndAlso CDate(drItemLote("Validade")) = dr("Validade") AndAlso qtdeAjuste > 0 Then
                                    drItemLote("Quantidade") = dr("Quantidade")
                                    drItemLote("Consumo") += qtdeAjuste

                                    qtdeAjuste = 0

                                    Exit For
                                End If
                            Next
                        Else
                            Dim temLote As Boolean = False

                            For Each drItemLote As DataRow In CType(Session("objLoteFornecedor" & HID.Value), DataTable).Rows
                                If drItemLote("CodigoProdutoProducao") = CodigoProdutoProducaoSelecionado And drItemLote("Produto") = dr("Produto") AndAlso drItemLote("Lote") = dr("Lote") AndAlso CDate(drItemLote("Validade")) = dr("Validade") Then
                                    temLote = True

                                    Exit For
                                End If
                            Next

                            If Not temLote Then
                                Dim drItem As DataRow = CType(Session("objLoteFornecedor" & HID.Value), DataTable).NewRow()

                                drItem("CodigoProdutoProducao") = CodigoProdutoProducaoSelecionado
                                drItem("Produto") = dr("Produto")
                                drItem("Lote") = dr("Lote")
                                drItem("Fabricado") = CDate(dr("Validade")).ToString("dd/MM/yyyy")
                                drItem("Validade") = CDate(dr("Validade")).ToString("dd/MM/yyyy")
                                drItem("Quantidade") = dr("Quantidade")

                                If qtdeConsumo > 0 Then
                                    If CDec(dr("Quantidade")) >= qtdeConsumo Then
                                        drItem("Consumo") = qtdeConsumo
                                    Else
                                        drItem("Consumo") = dr("Quantidade")
                                    End If

                                    qtdeConsumo -= drItem("Consumo")
                                End If

                                CType(Session("objLoteFornecedor" & HID.Value), DataTable).Rows.Add(drItem)

                            End If
                        End If
                    Next


                    Dim objProduto As New Produto(row.Cells(1).Text)

                    If objProduto.ControlarEstoque AndAlso objProduto.EstoqueMinimo > 0 Then
                        If (ds.Tables(0).Compute("SUM(Quantidade)", "") - CDec(row.Cells(4).Text)) < objProduto.EstoqueMinimo Then

                            Dim ItemEstoqueMinimo As DataRow = CType(Session("objEstoqueMinimo" & HID.Value), DataTable).NewRow()

                            ItemEstoqueMinimo("Produto") = objProduto.Codigo
                            ItemEstoqueMinimo("Nome") = objProduto.Nome
                            ItemEstoqueMinimo("EstoqueMinimo") = objProduto.EstoqueMinimo.ToString("N4")
                            ItemEstoqueMinimo("Faturando") = row.Cells(4).Text
                            Dim somaQuantidade As Decimal = ds.Tables(0).Compute("SUM(Quantidade)", "") - CDec(row.Cells(4).Text)
                            ItemEstoqueMinimo("Saldo") = somaQuantidade.ToString("N4")
                            CType(Session("objEstoqueMinimo" & HID.Value), DataTable).Rows.Add(ItemEstoqueMinimo)
                        End If
                    End If
                Else
                    MsgBox(Me.Page, "Não foi encontrato nenhum Lote para o Produto " & row.Cells(1).Text & "-" & row.Cells(2).Text, eTitulo.Info)
                    Exit Sub
                End If
            End If

            Dim dv As DataView = New DataView(CType(Session("objLoteFornecedor" & HID.Value), DataTable))
            dv.RowFilter = "CodigoProdutoProducao = '" & CodigoProdutoProducaoSelecionado & "' AND Produto = '" & gridConsumo.Rows(Session("objLinha" & HID.Value)).Cells(1).Text & "'"

            gridLoteDeFornecedor.DataSource = dv.ToTable()
            gridLoteDeFornecedor.DataBind()

        End If

        divConsumoLote.Visible = True
        divConsumo.Style.Value = "width: 69%;"
        divConsumoLote.Style.Value = "width: 29%;"

        imgConfirmar.Visible = True

        Dim i As Integer = 0
        While i < gridLoteDeFornecedor.Rows.Count

            If Not gridConsumo.Rows(row.RowIndex).Cells(3).Text = "INFORMAR" Then
                CType(gridLoteDeFornecedor.Rows(i).FindControl("txtConsumoLote"), TextBox).Enabled = False
                imgConfirmar.Visible = False
            Else
                CType(gridLoteDeFornecedor.Rows(i).FindControl("txtConsumoLote"), TextBox).Enabled = True
            End If

            i += 1
        End While

        'Enviar E-mail caso tenha chego no estoque mínimo - Furlan - 04/11/2020
        If CType(Session("objEstoqueMinimo" & HID.Value), DataTable).Rows.Count > 0 Then enviarEstoqueMinimo()

    End Sub

    Private Sub AjustarConsumoLotes(ByVal produto As String, ByVal nomeProduto As String, ByVal qtdConsumo As Decimal)

        If Session("objAjustarLote" & produto & HID.Value) IsNot Nothing AndAlso CType(Session("objAjustarLote" & produto & HID.Value), Boolean) = False Then

            Dim pLote = New OrdemParaProducaoXConsumo()

            Dim ds As New DataSet
            ds = pLote.buscarLoteDeFornecedor(ddlEmpresa.SelectedValue.Split("-")(0), ddlEmpresa.SelectedValue.Split("-")(1), "'" & produto & "'")

            Dim dtLotes As DataTable = CType(Session("objLoteFornecedor" & HID.Value), DataTable)

            Dim lotesComConsumo = dtLotes.AsEnumerable().
                                    Where(Function(r) r.Field(Of String)("CodigoProdutoProducao") = CodigoProdutoProducaoSelecionado.ToString() AndAlso
                                                      r.Field(Of String)("Produto") = produto AndAlso
                                                      Not IsDBNull(r("OldConsumo")) AndAlso
                                                      CDec(r("OldConsumo")) > 0).
                                    ToList()

            Dim saldoRestante As Decimal = qtdConsumo

            ' === 1º PASSO: usar lotes com consumo anterior (OldConsumo) ===
            For Each loteGravadoEmBanco As DataRow In lotesComConsumo
                
                produto = loteGravadoEmBanco("Produto").ToString()
                Dim lote = loteGravadoEmBanco("Lote").ToString()
                Dim validade = CDate(loteGravadoEmBanco("Validade"))

                ' Busca o lote correspondente no DataSet (saldo atual)
                Dim drBanco = ds.Tables(0).AsEnumerable().FirstOrDefault(Function(dr) dr("Produto").ToString() = produto AndAlso
                                                                             dr("Lote").ToString() = lote AndAlso
                                                                             CDate(dr("Validade")) = validade)

                Dim oldConsumo As Decimal = CDec(loteGravadoEmBanco("OldConsumo"))

                If drBanco IsNot Nothing Then

                    Dim saldoAtual As Decimal

                    If rbSim.Checked Then
                        saldoAtual = oldConsumo
                    Else
                        saldoAtual = CDec(drBanco("Quantidade")) + oldConsumo
                    End If

                    'A quantidade do lote sempre vai ser o OldConsumo que é o saldo em banco + o saldo atual
                    loteGravadoEmBanco("Quantidade") = CDec(drBanco("Quantidade")) + oldConsumo

                    If saldoRestante = 0 Then
                        loteGravadoEmBanco("Consumo") = 0
                    ElseIf saldoAtual - saldoRestante > 0 Then
                        loteGravadoEmBanco("Consumo") = saldoRestante
                        saldoRestante = 0
                    Else
                        saldoRestante -= saldoAtual
                        loteGravadoEmBanco("Consumo") = saldoAtual
                    End If

                Else

                    If rbSim.Checked Then

                        If saldoRestante < oldConsumo Then

                            loteGravadoEmBanco("Quantidade") = saldoRestante
                            loteGravadoEmBanco("Consumo") = saldoRestante

                        Else

                            loteGravadoEmBanco("Quantidade") = oldConsumo
                            loteGravadoEmBanco("Consumo") = oldConsumo

                        End If

                        saldoRestante -= oldConsumo

                        If saldoRestante <= 0 Then
                            saldoRestante = 0
                        End If

                    Else

                        'Se não tem mais saldo no lote, e não está dado baixa na OP, iremos consumir de outro lote, então zeramos os valores dos lotes iniciais
                        loteGravadoEmBanco("Quantidade") = 0
                        loteGravadoEmBanco("Consumo") = 0

                    End If

                End If

            Next

            If ds.Tables(0).Rows.Count = 0 And saldoRestante > 0 Then
                MsgBox(Me.Page, "Não foi encontrado o lote do Produto " & produto & "-" & nomeProduto, eTitulo.Info)
                Exit Sub
            End If

            ' === 2º PASSO: se ainda sobrou, usar novos lotes do banco (Pepsi) ===
            For Each dr In ds.Tables(0).Rows

                Dim produtoBanco = dr("Produto").ToString()
                Dim loteBanco = dr("Lote").ToString()
                Dim validadeBanco = CDate(dr("Validade"))
                Dim saldoBanco As Decimal = CDec(dr("Quantidade"))

                ' Verifica se esse lote já foi usado acima
                Dim dtExiste = dtLotes.AsEnumerable().Where(Function(r) r("CodigoProdutoProducao") = CodigoProdutoProducaoSelecionado AndAlso
                                r("Produto").ToString() = produtoBanco AndAlso
                                r("Lote").ToString() = loteBanco AndAlso
                                CDate(r("Validade")) = validadeBanco)

                If dtExiste.Count() = 0 Then

                    Dim drItem As DataRow = dtLotes.NewRow()

                    drItem("CodigoProdutoProducao") = CodigoProdutoProducaoSelecionado
                    drItem("Produto") = produtoBanco
                    drItem("Lote") = loteBanco
                    drItem("Fabricado") = validadeBanco.ToString("dd/MM/yyyy")
                    drItem("Validade") = validadeBanco.ToString("dd/MM/yyyy")
                    drItem("Quantidade") = saldoBanco

                    If saldoBanco >= saldoRestante Then

                        drItem("Consumo") = saldoRestante
                        saldoRestante = 0

                    Else

                        drItem("Consumo") = saldoBanco
                        saldoRestante -= saldoBanco

                    End If

                    dtLotes.Rows.Add(drItem)

                Else

                    ' Atualização
                    For Each drItem As DataRow In dtExiste

                        If Not IsDBNull(drItem("OldConsumo")) AndAlso drItem("OldConsumo") > 0 Then
                            Continue For
                        End If

                        Dim consumoAtual As Decimal = 0
                        Dim saldoDisponivel As Decimal = saldoBanco

                        If saldoBanco >= saldoRestante Then
                            drItem("Consumo") = saldoRestante
                            saldoRestante = 0
                        Else
                            saldoRestante -= saldoBanco
                        End If

                        If saldoRestante <= 0 Then
                            Exit For
                        End If

                    Next

                End If

            Next

            Session("objLoteFornecedor" & HID.Value) = dtLotes

        End If

        Dim dv As DataView = New DataView(CType(Session("objLoteFornecedor" & HID.Value), DataTable))
        dv.RowFilter = "CodigoProdutoProducao = '" & CodigoProdutoProducaoSelecionado & "' AND Produto = '" & produto & "'"

        gridLoteDeFornecedor.DataSource = dv.ToTable()
        gridLoteDeFornecedor.DataBind()

    End Sub

    Protected Sub imgConfirmar_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles imgConfirmar.Click
        SessaoRecuperaOrdem()

        Dim consumoLote As Decimal = 0
        Dim i As Integer = 0
        While i < gridLoteDeFornecedor.Rows.Count

            If Not String.IsNullOrWhiteSpace(CType(gridLoteDeFornecedor.Rows(i).FindControl("txtConsumoLote"), TextBox).Text) Then
                consumoLote += CDec(CType(gridLoteDeFornecedor.Rows(i).FindControl("txtConsumoLote"), TextBox).Text)
                If CDec(CType(gridLoteDeFornecedor.Rows(i).FindControl("txtConsumoLote"), TextBox).Text) > CDec(gridLoteDeFornecedor.Rows(i).Cells(2).Text()) Then
                    MsgBox(Me.Page, String.Format("Consumo de Lote: {0} informado não pode ser maior que o saldo!", gridLoteDeFornecedor.Rows(i).Cells(0).Text()), eTitulo.Info)
                    Exit Sub
                End If
            End If

            i += 1
        End While

        Dim totalProduto As Decimal = 0
        Dim totalAjusteProduto As Decimal = 0

        If objOP.IUD = "U" Then
            If CType(gridConsumo.Rows(Session("objLinha" & HID.Value)).FindControl("DdlSinal"), DropDownList).SelectedValue = "+" Then
                totalProduto = CDec(gridConsumo.Rows(Session("objLinha" & HID.Value)).Cells(4).Text())
                totalAjusteProduto = CDec(CType(gridConsumo.Rows(Session("objLinha" & HID.Value)).FindControl("txtAjusteConsumo"), TextBox).Text)
                totalProduto += totalAjusteProduto
            ElseIf CType(gridConsumo.Rows(Session("objLinha" & HID.Value)).FindControl("DdlSinal"), DropDownList).SelectedValue = "-" Then
                totalProduto = CDec(gridConsumo.Rows(Session("objLinha" & HID.Value)).Cells(4).Text())
                totalAjusteProduto = CDec(CType(gridConsumo.Rows(Session("objLinha" & HID.Value)).FindControl("txtAjusteConsumo"), TextBox).Text)
                totalProduto -= totalAjusteProduto
            End If
        Else
            totalProduto = CDec(gridConsumo.Rows(Session("objLinha" & HID.Value)).Cells(4).Text())
        End If

        For Each row In CType(Session("objLoteFornecedor" & HID.Value), DataTable).Rows
            If row("CodigoProdutoProducao") = CodigoProdutoProducaoSelecionado And row("Produto") = gridConsumo.Rows(Session("objLinha" & HID.Value)).Cells(1).Text() Then
                row("Consumo") = 0
            End If
        Next

        If totalProduto = consumoLote Then
            CType(gridConsumo.Rows(Session("objLinha" & HID.Value)).FindControl("imgSelecionaLote"), ImageButton).ToolTip = ""
            CType(gridConsumo.Rows(Session("objLinha" & HID.Value)).FindControl("imgSelecionaLote"), ImageButton).Style.Value = ""
            CType(gridConsumo.Rows(Session("objLinha" & HID.Value)).FindControl("imgSelecionaLote"), ImageButton).ImageUrl = "~/images/select.jpg"

            For Each drConsumo As DataRow In CType(Session("objConsumo" & HID.Value), DataTable).Rows

                If drConsumo("CodigoProdutoProducao") = CodigoProdutoProducaoSelecionado AndAlso drConsumo("Produto") = gridConsumo.Rows(Session("objLinha" & HID.Value)).Cells(1).Text() Then

                    drConsumo("NumeroDoLote") = "INFORMADO"
                    Exit For
                End If

            Next

            gridConsumo.Rows(Session("objLinha" & HID.Value)).Cells(3).Text() = "INFORMADO"

            i = 0
            While i < gridLoteDeFornecedor.Rows.Count
                If Not String.IsNullOrWhiteSpace(CType(gridLoteDeFornecedor.Rows(i).FindControl("txtConsumoLote"), TextBox).Text) AndAlso
                    CDec(CType(gridLoteDeFornecedor.Rows(i).FindControl("txtConsumoLote"), TextBox).Text) > 0 Then

                    For Each row In CType(Session("objLoteFornecedor" & HID.Value), DataTable).Rows

                        If row("CodigoProdutoProducao") = CodigoProdutoProducaoSelecionado And row("Produto") = gridConsumo.Rows(Session("objLinha" & HID.Value)).Cells(1).Text() _
                            AndAlso row("Lote") = gridLoteDeFornecedor.Rows(i).Cells(0).Text() AndAlso CDate(gridLoteDeFornecedor.Rows(i).Cells(1).Text()) = CDate(row("Validade")) Then

                            row("Consumo") = CDec(CType(gridLoteDeFornecedor.Rows(i).FindControl("txtConsumoLote"), TextBox).Text)
                        End If
                    Next
                End If

                CType(gridLoteDeFornecedor.Rows(i).FindControl("txtConsumoLote"), TextBox).Enabled = False

                i += 1
            End While

            Session("objAjustarLote" & gridConsumo.Rows(Session("objLinha" & HID.Value)).Cells(1).Text() & HID.Value) = True
            imgConfirmar.Visible = False

            gridLoteDeFornecedor.DataSource = Nothing
            gridLoteDeFornecedor.DataBind()
            divConsumo.Style.Value = "width:      99%;"
            divConsumoLote.Visible = False

        Else
            MsgBox(Me.Page, "Consumo de Lote(s) informado(s) difere do Total do Item.", eTitulo.Info)
        End If
    End Sub

    Protected Sub lnkAjustaEspecificacao_Click(sender As Object, e As EventArgs) Handles lnkAjustaEspecificacao.Click
        Try
            SessaoRecuperaOrdem()

            For Each pE In objOP.ItensDeEspecificacao
                For Each drEP As DataRow In CType(Session("objEspecificacaoDoProduto" & HID.Value), DataTable).Rows
                    If drEP("CodigoProdutoProducao") = CodigoProdutoProducaoSelecionado And drEP("Codigo") = pE.CodigoEspecificacao Then
                        drEP("Resultado") = pE.Resultado
                        Exit For
                    End If
                Next
            Next

            gridEspecificacaoDoProduto.DataSource = dtgridEspecificacaoDoProduto()
            gridEspecificacaoDoProduto.DataBind()

            For Each row In gridEspecificacaoDoProduto.Rows
                CType(gridEspecificacaoDoProduto.Rows(row.RowIndex).FindControl("txtResultado"), TextBox).Enabled = True
            Next

            lnkAjustaEspecificacao.Parent.Visible = False
            lnkConfirmaEspecificacao.Parent.Visible = True
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkConfirmaEspecificacao_Click(sender As Object, e As EventArgs) Handles lnkConfirmaEspecificacao.Click
        Try
            SessaoRecuperaOrdem()

            For Each row In gridEspecificacaoDoProduto.Rows
                For Each pE In objOP.ItensDeEspecificacao.Where(Function(x) x.CodigoProdutoProducao = CodigoProdutoProducaoSelecionado)
                    If pE.CodigoEspecificacao = row.Cells(1).Text Then
                        pE.Resultado = CDec(CType(gridEspecificacaoDoProduto.Rows(row.RowIndex).FindControl("txtResultado"), TextBox).Text)
                    End If
                Next
            Next

            For Each pE In objOP.ItensDeEspecificacao
                For Each drEP As DataRow In CType(Session("objEspecificacaoDoProduto" & HID.Value), DataTable).Rows
                    If drEP("CodigoProdutoProducao") = CodigoProdutoProducaoSelecionado And drEP("Codigo") = pE.CodigoEspecificacao Then
                        drEP("Resultado") = pE.Resultado
                        Exit For
                    End If
                Next
            Next

            gridEspecificacaoDoProduto.DataSource = dtgridEspecificacaoDoProduto()
            gridEspecificacaoDoProduto.DataBind()

            For Each row In gridEspecificacaoDoProduto.Rows
                CType(gridEspecificacaoDoProduto.Rows(row.RowIndex).FindControl("txtResultado"), TextBox).Enabled = False
            Next

            SessaoSalvaOrdem()

            lnkConfirmaEspecificacao.Parent.Visible = False
            lnkAjustaEspecificacao.Parent.Visible = True
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub ddlEPI_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ddlEPI.SelectedIndexChanged
        Try
            SessaoRecuperaOrdem()

            If ddlEPI.SelectedIndex > 0 Then
                Dim temItem As Boolean = False

                For Each dr As DataRow In CType(Session("objEPI" & HID.Value), DataTable).Rows
                    If dr("Codigo") = ddlEPI.SelectedItem.Value Then
                        temItem = True
                    End If
                Next

                If temItem Then
                    MsgBox(Me.Page, "EPI já foi selecionado.", eTitulo.Info)
                    Exit Sub
                End If

                If objOP.IUD = "U" Then
                    objOPXEpi = New OrdemParaProducaoXEPI(objOP)

                    objOPXEpi.IUD = "I"
                    objOPXEpi.CodigoEPI = ddlEPI.SelectedItem.Value

                    objOP.ItensDeEPI.Add(objOPXEpi)

                    SessaoSalvaOrdem()
                End If

                Dim drItem As DataRow = CType(Session("objEPI" & HID.Value), DataTable).NewRow()

                drItem("Codigo") = ddlEPI.SelectedItem.Value
                drItem("Descricao") = ddlEPI.SelectedItem.Text

                CType(Session("objEPI" & HID.Value), DataTable).Rows.Add(drItem)

                gridEPI.DataSource = CType(Session("objEPI" & HID.Value), DataTable)
                gridEPI.DataBind()

                ddlEPI.SelectedIndex = 0
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub imgRemoverEPI_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)
        Dim img As ImageButton = CType(sender, ImageButton)
        Dim row As GridViewRow = CType(img.NamingContainer, GridViewRow)

        CType(Session("objEPI" & HID.Value), DataTable).Rows(row.RowIndex).Delete()

        gridEPI.DataSource = CType(Session("objEPI" & HID.Value), DataTable)
        gridEPI.DataBind()
    End Sub

    Protected Sub imgAddEmbalagem_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)
        Try
            If ddlEmbalagem.SelectedIndex = 0 Then
                MsgBox(Me.Page, "Embalagem não foi selecionada.", eTitulo.Info)
                Exit Sub
            ElseIf ddlUnidadeDeComercializacao.SelectedIndex = 0 Then
                MsgBox(Me.Page, "Unidade de Comercialização não foi selecionada.", eTitulo.Info)
                Exit Sub
            ElseIf String.IsNullOrWhiteSpace(txtQuantidadeEmbalagem.Text) OrElse CDec(txtQuantidadeEmbalagem.Text) = 0 Then
                MsgBox(Me.Page, "Quantidade de Embalagem não foi informada.", eTitulo.Info)
                Exit Sub

            ElseIf String.IsNullOrWhiteSpace(txtCapacidade.Text) OrElse CDec(txtCapacidade.Text) = 0 Then
                MsgBox(Me.Page, "Capacidade da Embalagem não foi informada.", eTitulo.Info)
                Exit Sub
            End If

            Dim temItem As Boolean = False

            For Each dr As DataRow In CType(Session("objEmbalagem" & HID.Value), DataTable).Rows
                If dr("Codigo") = ddlEmbalagem.SelectedItem.Value AndAlso dr("CodigoUnidade") = ddlUnidadeDeComercializacao.SelectedItem.Value Then
                    temItem = True
                End If
            Next

            If temItem Then
                MsgBox(Me.Page, "Embalagem já foi selecionada.", eTitulo.Info)
                Exit Sub
            End If

            SessaoRecuperaOrdem()

            If objOP.IUD = "U" Then
                objOPXEmb = New OrdemParaProducaoXEmbalagens(objOP)

                objOPXEmb.IUD = "I"
                objOPXEmb.CodigoEmbalagem = ddlEmbalagem.SelectedItem.Value
                objOPXEmb.CodigoUnidade = ddlUnidadeDeComercializacao.SelectedItem.Value
                objOPXEmb.Quantidade = CInt(txtQuantidadeEmbalagem.Text)
                objOPXEmb.Capacidade = CDec(txtCapacidade.Text)

                objOP.ItensDeEmabalagem.Add(objOPXEmb)

                SessaoSalvaOrdem()
            End If

            Dim drItem As DataRow = CType(Session("objEmbalagem" & HID.Value), DataTable).NewRow()

            drItem("Codigo") = ddlEmbalagem.SelectedItem.Value
            drItem("Descricao") = ddlEmbalagem.SelectedItem.Text
            drItem("CodigoUnidade") = ddlUnidadeDeComercializacao.SelectedItem.Value
            drItem("DescricaoUnidade") = ddlUnidadeDeComercializacao.SelectedItem.Text
            drItem("Quantidade") = txtQuantidadeEmbalagem.Text
            drItem("Capacidade") = txtCapacidade.Text

            CType(Session("objEmbalagem" & HID.Value), DataTable).Rows.Add(drItem)

            gridEmbalagem.DataSource = CType(Session("objEmbalagem" & HID.Value), DataTable)
            gridEmbalagem.DataBind()

            ddlEmbalagem.SelectedIndex = 0
            ddlUnidadeDeComercializacao.SelectedIndex = 0
            txtQuantidadeEmbalagem.Text = String.Empty
            txtCapacidade.Text = String.Empty
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub imgRemoverEmbalagem_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)
        Dim img As ImageButton = CType(sender, ImageButton)
        Dim row As GridViewRow = CType(img.NamingContainer, GridViewRow)

        CType(Session("objEmbalagem" & HID.Value), DataTable).Rows(row.RowIndex).Delete()

        gridEmbalagem.DataSource = CType(Session("objEmbalagem" & HID.Value), DataTable)
        gridEmbalagem.DataBind()
    End Sub

    Protected Sub ddlProcedimento_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ddlProcedimento.SelectedIndexChanged
        Try
            SessaoRecuperaOrdem()

            If ddlProcedimento.SelectedIndex > 0 Then
                Dim temItem As Boolean = False

                For Each dr As DataRow In CType(Session("objProcedimento" & HID.Value), DataTable).Rows
                    If dr("Codigo") = ddlProcedimento.SelectedItem.Value Then
                        temItem = True
                    End If
                Next

                If temItem Then
                    MsgBox(Me.Page, "Procedimento já foi selecionado.", eTitulo.Info)
                    Exit Sub
                End If

                If objOP.IUD = "U" Then
                    objOPXProcedimento = New OrdemParaProducaoXProcedimento(objOP)

                    objOPXProcedimento.IUD = "I"
                    objOPXProcedimento.CodigoProcedimento = ddlProcedimento.SelectedItem.Value

                    objOP.ItensDeProcedimento.Add(objOPXProcedimento)

                    SessaoSalvaOrdem()
                End If

                Dim drItem As DataRow = CType(Session("objProcedimento" & HID.Value), DataTable).NewRow()

                drItem("Codigo") = ddlProcedimento.SelectedItem.Value
                drItem("Descricao") = ddlProcedimento.SelectedItem.Text

                CType(Session("objProcedimento" & HID.Value), DataTable).Rows.Add(drItem)

                gridProcedimento.DataSource = CType(Session("objProcedimento" & HID.Value), DataTable)
                gridProcedimento.DataBind()

                ddlProcedimento.SelectedIndex = 0
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub imgRemoverProcedimento_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)
        Dim img As ImageButton = CType(sender, ImageButton)
        Dim row As GridViewRow = CType(img.NamingContainer, GridViewRow)

        CType(Session("objProcedimento" & HID.Value), DataTable).Rows(row.RowIndex).Delete()

        gridProcedimento.DataSource = CType(Session("objProcedimento" & HID.Value), DataTable)
        gridProcedimento.DataBind()
    End Sub

    Function ValidarSelecao() As Boolean

        If ddlUnidadeNegocio.SelectedIndex = 0 Then
            MsgBox(Me.Page, "Unidade de Negócio não foi selecionada.", eTitulo.Info)
            Return False
        ElseIf ddlEmpresa.SelectedIndex = 0 Then
            MsgBox(Me.Page, "Empresa não foi selecionada.", eTitulo.Info)
            Return False
        End If

        For Each drItem As DataRow In CType(Session("objConsumo" & HID.Value), DataTable).Rows
            If drItem("NumeroDoLote") = "INFORMAR" Then
                MsgBox(Me.Page, "Lote do Produto " & drItem("Produto") & " - " & drItem("NomeProduto") & " não foi informado.", eTitulo.Info)
                Return False
            End If
        Next

        Dim i As Integer = 0
        While i < gridConsumo.Rows.Count

            If gridConsumo.Rows(i).Cells(3).Text = "INFORMAR" Then
                MsgBox(Me.Page, "Lote do Produto " & gridConsumo.Rows(i).Cells(1).Text & " - " & gridConsumo.Rows(i).Cells(2).Text & " não foi informado.", eTitulo.Info)
                Return False
            End If

            i += 1
        End While

        If rbSim.Checked AndAlso CDate(txtDataEstoque.Text) < CDate(txtDataProducao.Text) Then
            MsgBox(Me.Page, "Data para estoque não pode ser menor que a data da Ordem de Produção.", eTitulo.Info)

            Return False
        End If

        For Each drItem As DataRow In CType(Session("objConsumo" & HID.Value), DataTable).Rows

            Dim totalProduto As Decimal = 0

            If drItem("Sinal") = "+" Then
                totalProduto = drItem("Quantidade") + drItem("AjusteConsumo")
            ElseIf drItem("Sinal") = "-" Then
                totalProduto = drItem("Quantidade") - drItem("AjusteConsumo")
            End If

            Dim QuantidadeLote As Decimal = 0

            Dim objProduto = New Produto(drItem("Produto"))
            Dim bVerificarTotalProdutoXQuantidadeLote As Boolean = False

            If objProduto.ControlarNumeroDoLote Then
                For Each drLote As DataRow In CType(Session("objLoteFornecedor" & HID.Value), DataTable).Rows
                    If drItem("CodigoProdutoProducao") = drLote("CodigoProdutoProducao") And drItem("Produto") = drLote("Produto") AndAlso Not IsDBNull(drLote("Consumo")) AndAlso drLote("Consumo") > 0 Then
                        QuantidadeLote += drLote("Consumo")
                        bVerificarTotalProdutoXQuantidadeLote = True
                    End If
                Next

                If bVerificarTotalProdutoXQuantidadeLote AndAlso totalProduto <> QuantidadeLote Then
                    MsgBox(Me.Page, "Quantidade do Produto " & drItem("Produto") & " difere com a quantidade informada no lote, verifique.", eTitulo.Info)

                    Return False
                End If

            End If
        Next

        Return True

    End Function

    Private Sub verNumerador()
        If Not String.IsNullOrWhiteSpace(ddlEmpresa.SelectedValue) Then
            Dim num As New [Lib].Negocio.Numerador(ddlEmpresa.SelectedValue.Split("-")(0), ddlEmpresa.SelectedValue.Split("-")(1), 75)

            txtSequencia.Text = num.Sequencia + 1

            txtSequencia.Enabled = False

        End If
    End Sub

    Private Sub Limpar()

        bAlterarOrdemDeProducao = False
        bGerenciarOrdemDeProducao = False
        TabIncluirParaConsumo.Visible = False
        lnkConsultar.Parent.Visible = True
        lnkNovo.Parent.Visible = False
        lnkAtualizar.Parent.Visible = False
        lnkExcluir.Parent.Visible = False
        lnkImprimirOrdem.Parent.Visible = False
        lnkImprimirLaudo.Parent.Visible = False
        lnkImprimirLaudoManual.Parent.Visible = True

        divConsumo.Style.Value = "width: 99%;"
        divConsumoLote.Visible = False

        lnkAjustaProducao.Parent.Visible = False
        lnkConfirmaProducao.Parent.Visible = False
        divAjusteProducao.Visible = False

        lnkAjustaConsumo.Parent.Visible = False
        lnkConfirmaConsumo.Parent.Visible = False
        divAjusteConsumo.Visible = False

        lnkAjustaEspecificacao.Parent.Visible = False
        lnkConfirmaEspecificacao.Parent.Visible = False
        divEspecificacao.Visible = False

        rbNao.Checked = True
        rbSim.Checked = False
        rbNao.Enabled = False
        rbSim.Enabled = False
        dataEstoque.Visible = False

        Session.Remove("objProdutoxPRD" & HID.Value)
        Session.Remove("objConsultaXOrdem" & HID.Value)

        Session.Remove("objLinha" & HID.Value)
        Session.Remove("objLoteFornecedor" & HID.Value)
        Session.Remove("objEPI" & HID.Value)
        Session.Remove("objEmbalagem" & HID.Value)
        Session.Remove("objEspecificacaoDoProduto" & HID.Value)
        Session.Remove("objProducao" & HID.Value)
        Session.Remove("objConsumo" & HID.Value)
        Session.Remove("objInsumo" & HID.Value)
        Session.Remove("objProcedimento" & HID.Value)
        Session.Remove("objEstoqueMinimo" & HID.Value)

        HID.Value = Guid.NewGuid().ToString

        Dim dtEstoqueMinimo As New DataTable("ItemEstoqueMinimo")
        dtEstoqueMinimo.Columns.Add("Produto", Type.GetType("System.String"))
        dtEstoqueMinimo.Columns.Add("Nome", Type.GetType("System.String"))
        dtEstoqueMinimo.Columns.Add("EstoqueMinimo", Type.GetType("System.String"))
        dtEstoqueMinimo.Columns.Add("Faturando", Type.GetType("System.String"))
        dtEstoqueMinimo.Columns.Add("Saldo", Type.GetType("System.String"))
        Session("objEstoqueMinimo" & HID.Value) = dtEstoqueMinimo

        Dim dtLoteFornecedor As New DataTable("ItemLoteFornecedor")
        dtLoteFornecedor.Columns.Add("CodigoProdutoProducao", Type.GetType("System.String"))
        dtLoteFornecedor.Columns.Add("Produto", Type.GetType("System.String"))
        dtLoteFornecedor.Columns.Add("Lote", Type.GetType("System.String"))
        dtLoteFornecedor.Columns.Add("Fabricado", Type.GetType("System.String"))
        dtLoteFornecedor.Columns.Add("Validade", Type.GetType("System.String"))
        dtLoteFornecedor.Columns.Add("Quantidade", Type.GetType("System.Decimal"))
        dtLoteFornecedor.Columns.Add("Consumo", Type.GetType("System.Decimal"))
        dtLoteFornecedor.Columns.Add("OldConsumo", Type.GetType("System.Decimal"))
        dtLoteFornecedor.Columns.Add("DesabilitarEdicaoConsumo", Type.GetType("System.Decimal"))
        Session("objLoteFornecedor" & HID.Value) = dtLoteFornecedor

        Dim dtProcedimento As New DataTable("ItemProcedimento")
        dtProcedimento.Columns.Add("Codigo", Type.GetType("System.String"))
        dtProcedimento.Columns.Add("Descricao", Type.GetType("System.String"))
        Session("objProcedimento" & HID.Value) = dtProcedimento

        Dim dtEPI As New DataTable("ItemEPI")
        dtEPI.Columns.Add("Codigo", Type.GetType("System.String"))
        dtEPI.Columns.Add("Descricao", Type.GetType("System.String"))
        Session("objEPI" & HID.Value) = dtEPI

        Dim dtEmbalagem As New DataTable("ItemEmbalagem")
        dtEmbalagem.Columns.Add("Codigo", Type.GetType("System.String"))
        dtEmbalagem.Columns.Add("Descricao", Type.GetType("System.String"))
        dtEmbalagem.Columns.Add("CodigoUnidade", Type.GetType("System.String"))
        dtEmbalagem.Columns.Add("DescricaoUnidade", Type.GetType("System.String"))
        dtEmbalagem.Columns.Add("Quantidade", Type.GetType("System.Decimal"))
        dtEmbalagem.Columns.Add("Capacidade", Type.GetType("System.Decimal"))
        Session("objEmbalagem" & HID.Value) = dtEmbalagem

        Dim dtEspecificacaoDoProduto As New DataTable("ItemEspecificacaoDoProduto")
        dtEspecificacaoDoProduto.Columns.Add("CodigoProdutoProducao", Type.GetType("System.String"))
        dtEspecificacaoDoProduto.Columns.Add("Codigo", Type.GetType("System.String"))
        dtEspecificacaoDoProduto.Columns.Add("Descricao", Type.GetType("System.String"))
        dtEspecificacaoDoProduto.Columns.Add("ResultadoTxt", Type.GetType("System.String"))
        dtEspecificacaoDoProduto.Columns.Add("ReferenciaTxt", Type.GetType("System.String"))
        dtEspecificacaoDoProduto.Columns.Add("FaixaInicial", Type.GetType("System.Decimal"))
        dtEspecificacaoDoProduto.Columns.Add("FaixaFinal", Type.GetType("System.Decimal"))
        dtEspecificacaoDoProduto.Columns.Add("Resultado", Type.GetType("System.Decimal"))
        Session("objEspecificacaoDoProduto" & HID.Value) = dtEspecificacaoDoProduto

        Dim dtProducao As New DataTable("ItemProducao")
        dtProducao.Columns.Add("Produto", Type.GetType("System.String"))
        dtProducao.Columns.Add("NomeProduto", Type.GetType("System.String"))
        dtProducao.Columns.Add("CodigoUnidadeComercializacao", Type.GetType("System.String"))
        dtProducao.Columns.Add("FatorConversao", Type.GetType("System.Decimal"))
        dtProducao.Columns.Add("Quantidade", Type.GetType("System.Decimal"))
        dtProducao.Columns.Add("AjusteProducao", Type.GetType("System.Decimal"))
        dtProducao.Columns.Add("Lote", Type.GetType("System.String"))
        Session("objProducao" & HID.Value) = dtProducao

        Dim dtConsumo As New DataTable("ItemConsumo")
        dtConsumo.Columns.Add("CodigoProdutoProducao", Type.GetType("System.String"))
        dtConsumo.Columns.Add("Produto", Type.GetType("System.String"))
        dtConsumo.Columns.Add("NomeProduto", Type.GetType("System.String"))
        dtConsumo.Columns.Add("NumeroDoLote", Type.GetType("System.String"))
        dtConsumo.Columns.Add("Quantidade", Type.GetType("System.Decimal"))
        dtConsumo.Columns.Add("Percentual", Type.GetType("System.Decimal"))
        dtConsumo.Columns.Add("AjusteConsumo", Type.GetType("System.Decimal"))
        dtConsumo.Columns.Add("Sinal", Type.GetType("System.String"))
        Session("objConsumo" & HID.Value) = dtConsumo

        Dim dtInsumo As New DataTable("ItemInsumo")
        dtInsumo.Columns.Add("CodigoProdutoProducao", Type.GetType("System.String"))
        dtInsumo.Columns.Add("Produto", Type.GetType("System.String"))
        dtInsumo.Columns.Add("NomeProduto", Type.GetType("System.String"))
        dtInsumo.Columns.Add("Base", Type.GetType("System.Decimal"))
        dtInsumo.Columns.Add("Quantidade", Type.GetType("System.Decimal"))
        dtInsumo.Columns.Add("Estoque", Type.GetType("System.Decimal"))
        Session("objInsumo" & HID.Value) = dtInsumo

        Dim keysToRemove As New List(Of String)()

        ' Percorre todas as chaves da sessão
        For Each key As String In Session.Keys
            ' Verifica se o nome da chave contém "objLote"
            If key.Contains("objAjustarLote") Then
                keysToRemove.Add(key)
            End If
        Next

        ' Remove as chaves identificadas
        For Each key As String In keysToRemove
            Session.Remove(key)
        Next

        'ddlUnidadeNegocio.Enabled = True
        'ddlEmpresa.Enabled = True
        txtSequencia.Enabled = True
        txtDataProducao.Enabled = True
        txtDataValidade.Enabled = True
        txtQuantidadeProducao.Enabled = True
        ddlGrupoProdutoProducao.Enabled = True
        ddlEmbalagem.Enabled = True
        ddlUnidadeDeComercializacao.Enabled = True
        txtCapacidade.Enabled = True
        txtQuantidadeEmbalagem.Enabled = True
        lnkBuscaProdutoProducao.Enabled = True
        ddlUnidadeComercializacao.Items.Clear()
        ddlProdutosProducao.Enabled = True
        imgAdicionarDadosParaProducao.Enabled = True

        ddlEPI.Enabled = True
        ddlProcedimento.Enabled = True

        txtDataProducao.Text = Now.ToString("dd/MM/yyyy")

        Dim dtValida As Date = Now()
        dtValida = dtValida.AddYears(1)
        txtDataValidade.Text = dtValida.ToString("dd/MM/yyyy")

        txtSequencia.Text = String.Empty
        txtQuantidadeProducao.Text = String.Empty
        ddlGrupoProdutoProducao.SelectedIndex = 0
        ddlEPI.SelectedIndex = 0
        ddlEmbalagem.SelectedIndex = 0
        ddlUnidadeDeComercializacao.SelectedIndex = 0
        txtCapacidade.Text = String.Empty
        txtQuantidadeEmbalagem.Text = String.Empty
        txtObservacoes.Text = String.Empty
        CodigoProdutoProducaoSelecionado = String.Empty

        ddlProdutosProducao.Items.Clear()

        gridProducao.DataSource = Nothing
        gridProducao.DataBind()

        gridConsumo.DataSource = Nothing
        gridConsumo.DataBind()

        gridInsumo.DataSource = Nothing
        gridInsumo.DataBind()

        gridLoteDeFornecedor.DataSource = Nothing
        gridLoteDeFornecedor.DataBind()

        gridEspecificacaoDoProduto.DataSource = Nothing
        gridEspecificacaoDoProduto.DataBind()

        gridEPI.DataSource = Nothing
        gridEPI.DataBind()

        gridEmbalagem.DataSource = Nothing
        gridEmbalagem.DataBind()

        gridEstoque.DataSource = Nothing
        gridEstoque.DataBind()

        TabContainer1.ActiveTabIndex = 0
        TabInfConsumos.ActiveTabIndex = 0

        TabPanel7.Visible = False

        ucConsultaProduto.SetarHID(HID.Value)
        ucConsultaOrdemDeProducao.SetarHID(HID.Value)
        ucLaudoManual.SetarHID(HID.Value)

        txtDataInicial.Text = Now.ToString("dd/MM/yyyy")
        txtDataFinal.Text = Now.ToString("dd/MM/yyyy")

        ddlUsuarios.Items.Clear()
        ddlUsuarios.Items.Add("Inc.- " & Session("ssNomeUsuario"))

        objOP = New OrdemParaProducao
        objOP.IUD = "I"
        SessaoSalvaOrdem()

        LiberaEmpresa()

        LimparConsumo()

    End Sub

    Private Sub LimparConsumo()

        Session.Remove("objConsumoAlt" & HID.Value)
        Session.Remove("objInsumoAlt" & HID.Value)

        Dim dtConsumo As New DataTable("ItemConsumo")
        dtConsumo.Columns.Add("CodigoProdutoProducao", Type.GetType("System.String"))
        dtConsumo.Columns.Add("Produto", Type.GetType("System.String"))
        dtConsumo.Columns.Add("NomeProduto", Type.GetType("System.String"))
        dtConsumo.Columns.Add("Quantidade", Type.GetType("System.Decimal"))
        dtConsumo.Columns.Add("Lote", Type.GetType("System.String"))
        dtConsumo.Columns.Add("Validade", Type.GetType("System.String"))
        dtConsumo.Columns.Add("Sinal", Type.GetType("System.String"))
        Session("objConsumoAlt" & HID.Value) = dtConsumo

        Dim dtInsumo As New DataTable("ItemInsumo")
        dtInsumo.Columns.Add("CodigoProdutoProducao", Type.GetType("System.String"))
        dtInsumo.Columns.Add("Produto", Type.GetType("System.String"))
        dtInsumo.Columns.Add("NomeProduto", Type.GetType("System.String"))
        dtInsumo.Columns.Add("Quantidade", Type.GetType("System.Int64"))
        Session("objInsumoAlt" & HID.Value) = dtInsumo

        ddlGrupoProdutoProducao.Enabled = True
        lnkBuscaProdutoProducao.Enabled = True

        btnQuantidade.Visible = False
        txtQuantidadeAlt.Enabled = True
        ddlGrupoProdutoConsumo.Enabled = True
        lnkBuscaProdutoConsumo.Enabled = True
        ddlProdutosConsumo.Enabled = True

        ddlGrupoProdutoProducao.SelectedIndex = 0
        gridProducao.DataSource = Nothing
        gridProducao.DataBind()

        txtValidadeAlt.Text = Now.ToString("dd/MM/yyyy")
        txtLoteAlt.Text = String.Empty

        txtQuantidadeAlt.Text = String.Empty
        lblQuantidadeAlt.Text = String.Empty
        ddlGrupoProdutoConsumo.SelectedIndex = 0
        ddlProdutosConsumo.Items.Clear()
        gridConsumoAlt.DataSource = Nothing
        gridConsumoAlt.DataBind()

        btnQuantidadeIns.Visible = False
        txtQuantidadeIns.Text = String.Empty
        ddlGrupoProdutoInsumo.Enabled = True
        lnkBuscaProdutoInsumo.Enabled = True
        ddlProdutosInsumo.Enabled = True
        ddlGrupoProdutoInsumo.SelectedIndex = 0
        ddlProdutosInsumo.Items.Clear()
        gridInsumoAlt.DataSource = Nothing
        gridInsumoAlt.DataBind()

    End Sub

    Private Sub LiberaEmpresa()
        'If Not UsuarioServidor.LiberaEmpresa Then
        '    ddlUnidadeNegocio.Enabled = False
        '    ddlEmpresa.Enabled = False
        'End If
    End Sub

    Private Function RecuperarOrdem() As Boolean

        If String.IsNullOrWhiteSpace(objOP.ProdutosdaProducao.Count() = 0) Then
            MsgBox(Me.Page, "Ordem de Produção não foi encontrada!", eTitulo.Info)
            Return False
        Else
            TabPanel7.Visible = True

            ddlUnidadeNegocio.Enabled = False
            ddlEmpresa.Enabled = False
            txtDataProducao.Enabled = False
            txtSequencia.Enabled = False

            divAjusteConsumo.Visible = True

            objOP.IUD = "U"

            txtSequencia.Text = objOP.Codigo

            txtDataProducao.Text = objOP.Movimento.ToString("dd/MM/yyyy")
            txtDataProducao.Enabled = False

            txtDataValidade.Text = objOP.Validade.ToString("dd/MM/yyyy")

            dataEstoque.Visible = True
            rbNao.Enabled = True
            rbSim.Enabled = True

            If objOP.Estoque Then
                rbSim.Checked = True

                Dim ds As New DataSet
                ds = objOP.ItensEstoque()

                gridEstoque.DataSource = ds
                gridEstoque.DataBind()

                If ds.Tables(0).Rows.Count > 0 Then
                    txtDataEstoque.Text = CDate(ds.Tables(0).Rows(0).Item("Movimento")).ToString("dd/MM/yyyy")
                Else
                    rbNao.Checked = True
                    txtDataEstoque.Text = Now.ToString("dd/MM/yyyy")
                End If
            Else
                rbNao.Checked = True
                txtDataEstoque.Text = Now.ToString("dd/MM/yyyy")
            End If

            txtObservacoes.Text = objOP.Observacoes

            For Each produto As OrdemParaProducaoXProduto In objOP.ProdutosdaProducao

                Dim drItem As DataRow = CType(Session("objProducao" & HID.Value), DataTable).NewRow()
                drItem("Produto") = produto.CodigoProduto
                drItem("NomeProduto") = produto.Produto.Nome
                drItem("CodigoUnidadeComercializacao") = produto.CodigoUnidadeComercializacao
                drItem("FatorConversao") = produto.FatorConversao
                drItem("Quantidade") = produto.Quantidade
                drItem("AjusteProducao") = produto.QuantidadeDeAjuste
                drItem("Lote") = produto.Lote
                CType(Session("objProducao" & HID.Value), DataTable).Rows.Add(drItem)

            Next

            gridProducao.DataSource = CType(Session("objProducao" & HID.Value), DataTable)
            gridProducao.DataBind()

            If gridProducao.Rows.Count > 0 Then
                gridProducao.SelectedIndex = 0
                gridProducao_SelectedIndexChanged(Nothing, Nothing) ' Dispara o evento manualmente
            End If

            gridProducao.Columns(4).Visible = True
            gridProducao.Columns(5).Visible = True
            gridProducao.Columns(6).Visible = True

            divAjusteProducao.Visible = True

            For Each pC In objOP.ItensDeConsumo
                pC.IUD = objOP.IUD

                Dim drItemC As DataRow = CType(Session("objConsumo" & HID.Value), DataTable).NewRow()

                drItemC("CodigoProdutoProducao") = pC.CodigoProdutoProducao
                drItemC("Produto") = pC.CodigoProduto
                drItemC("NomeProduto") = pC.Produto.Nome

                If pC.ItensDeConsumoXLote.Count > 0 Then
                    drItemC("NumeroDoLote") = "INFORMADO"
                Else
                    drItemC("NumeroDoLote") = ""
                End If

                drItemC("Quantidade") = pC.Quantidade.ToString("N4")
                drItemC("Percentual") = pC.Percentual.ToString("N4")
                drItemC("AjusteConsumo") = pC.QuantidadeDeAjuste
                drItemC("Sinal") = pC.Sinal

                CType(Session("objConsumo" & HID.Value), DataTable).Rows.Add(drItemC)

                If pC.ItensDeConsumoXLote.Count > 0 Then
                    For Each pLote In pC.ItensDeConsumoXLote
                        pLote.IUD = pC.IUD

                        Dim drItemCL As DataRow = CType(Session("objLoteFornecedor" & HID.Value), DataTable).NewRow()

                        Dim pBuscarSaldoLote = New OrdemParaProducaoXConsumo()
                        Dim ds As New DataSet
                        ds = pBuscarSaldoLote.buscarLoteDeFornecedor(objOP.CodigoEmpresa, objOP.EnderecoEmpresa, "'" & pC.CodigoProduto & "'")

                        drItemCL("CodigoProdutoProducao") = pC.CodigoProdutoProducao
                        drItemCL("Produto") = pC.CodigoProduto
                        drItemCL("Lote") = pLote.Lote
                        drItemCL("Fabricado") = pLote.Validade.ToString("dd/MM/yyyy")
                        drItemCL("Validade") = pLote.Validade.ToString("dd/MM/yyyy")

                        'Se não houver saldo, o saldoserá o valor gravado no lote, em banco
                        drItemCL("Quantidade") = pLote.Quantidade
                        For Each dr In ds.Tables(0).Rows
                            If dr("Produto") = pC.CodigoProduto AndAlso pLote.Lote = dr("Lote") AndAlso dr("Quantidade") > pLote.Quantidade Then
                                drItemCL("Quantidade") = dr("Quantidade")
                            End If
                        Next
                        drItemCL("Consumo") = pLote.Quantidade
                        drItemCL("OldConsumo") = pLote.Quantidade
                        drItemCL("DesabilitarEdicaoConsumo") = 0

                        CType(Session("objLoteFornecedor" & HID.Value), DataTable).Rows.Add(drItemCL)
                    Next
                End If
            Next

            gridConsumo.DataSource = dtgridConsumo()
            gridConsumo.DataBind()

            'Gerenciamento exclusivo da ordem de produção
            If Funcoes.VerificaPermissao("GerenciarOrdemDeProducao", "ALTERAR") Then
                bGerenciarOrdemDeProducao = True
                TabIncluirParaConsumo.Visible = True
            Else
                bGerenciarOrdemDeProducao = False
                TabIncluirParaConsumo.Visible = False
            End If


            If Funcoes.VerificaPermissao("OrdemDeProducao", "ALTERAR") Then
                bAlterarOrdemDeProducao = True
            End If

            Session("objLinha" & HID.Value) = 0

            Dim i As Integer = 0
            While i < gridConsumo.Rows.Count

                If gridConsumo.Rows(i).Cells(3).Text = "INFORMADO" Then

                    CType(gridConsumo.Rows(i).FindControl("imgSelecionaLote"), ImageButton).ImageUrl = "~/images/select.jpg"
                    CType(gridConsumo.Rows(i).FindControl("imgSelecionaLote"), ImageButton).ToolTip = ""
                    CType(gridConsumo.Rows(i).FindControl("imgSelecionaLote"), ImageButton).Enabled = True

                ElseIf String.IsNullOrWhiteSpace(gridConsumo.Rows(i).Cells(3).Text.Replace("&nbsp;", "").Trim()) OrElse gridConsumo.Rows(i).Cells(3).Text.Length = 0 Then

                    CType(gridConsumo.Rows(i).FindControl("imgSelecionaLote"), ImageButton).ImageUrl = "~/images/important.png"
                    CType(gridConsumo.Rows(i).FindControl("imgSelecionaLote"), ImageButton).ToolTip = "Não informa Lote"
                    CType(gridConsumo.Rows(i).FindControl("imgSelecionaLote"), ImageButton).Enabled = False

                Else

                    Dim ds As New DataSet

                    objOPXC = New OrdemParaProducaoXConsumo(objOP)
                    ds = objOPXC.buscarLoteDeFornecedor(objOP.CodigoEmpresa, objOP.EnderecoEmpresa, "'" & gridConsumo.Rows(i).Cells(1).Text & "'")

                    If ds.Tables(0).Rows.Count > 0 Then
                        CType(gridConsumo.Rows(i).FindControl("imgSelecionaLote"), ImageButton).ImageUrl = "~/images/select.jpg"
                        CType(gridConsumo.Rows(i).FindControl("imgSelecionaLote"), ImageButton).ToolTip = ""
                        gridConsumo.Rows(i).Cells(3).Text = "INFORMAR"
                    Else
                        CType(gridConsumo.Rows(i).FindControl("imgSelecionaLote"), ImageButton).ImageUrl = "~/images/important.png"
                        CType(gridConsumo.Rows(i).FindControl("imgSelecionaLote"), ImageButton).ToolTip = "Não informa Lote"
                    End If

                End If

                CType(gridConsumo.Rows(i).FindControl("imgExcluir"), ImageButton).Visible = bGerenciarOrdemDeProducao
                'CType(gridConsumo.Rows(i).FindControl("imgSelecionaLote"), ImageButton).Enabled = bGerenciarOrdemDeProducao
                CType(gridConsumo.Rows(i).FindControl("DdlSinal"), DropDownList).Enabled = bGerenciarOrdemDeProducao
                CType(gridConsumo.Rows(i).FindControl("txtAjusteConsumo"), TextBox).Enabled = bGerenciarOrdemDeProducao

                For Each drItemC As DataRow In CType(Session("objConsumo" & HID.Value), DataTable).Rows
                    If drItemC("CodigoProdutoProducao") = CodigoProdutoProducaoSelecionado And drItemC("Produto") = gridConsumo.Rows(i).Cells(1).Text Then
                        CType(gridConsumo.Rows(i).FindControl("DdlSinal"), DropDownList).SelectedValue = drItemC("Sinal")
                    End If
                Next

                i += 1
            End While

            gridConsumo.Columns(6).Visible = True
            gridConsumo.Columns(7).Visible = True
            divConsumo.Style.Value = "width: 99%;"

            divConsumoLote.Visible = False

            Dim dv As DataView = New DataView(CType(Session("objLoteFornecedor" & HID.Value), DataTable))
            dv.RowFilter = "CodigoProdutoProducao = '" & CodigoProdutoProducaoSelecionado & "' AND Produto = '" & gridConsumo.Rows(Session("objLinha" & HID.Value)).Cells(1).Text & "'"

            gridLoteDeFornecedor.DataSource = dv.ToTable()
            gridLoteDeFornecedor.DataBind()

            imgConfirmar.Visible = False

            '1 - Carregar itens dos Insumos gravados
            Dim dataInsumo As Date = Now()
            Dim dataInsumoAnt As Date = dataInsumo.AddDays(-1)

            For Each pI In objOP.ItensDeInsumo
                pI.IUD = objOP.IUD

                Dim drItemI As DataRow = CType(Session("objInsumo" & HID.Value), DataTable).NewRow()

                drItemI("CodigoProdutoProducao") = pI.CodigoProdutoProducao
                drItemI("Produto") = pI.CodigoProduto
                drItemI("NomeProduto") = pI.Produto.Nome
                drItemI("Estoque") = CDec(0)
                'drItemI("Quantidade") = CInt(pI.Quantidade)
                drItemI("Quantidade") = pI.Quantidade

                Dim qInsumo = New OrdemParaProducaoXInsumo()

                drItemI("Estoque") = CDec(qInsumo.buscarEstoqueProduto(ddlEmpresa.SelectedValue.Split("-")(0), ddlEmpresa.SelectedValue.Split("-")(1), pI.CodigoProduto, dataInsumoAnt.ToString("yyyy-MM-dd"), dataInsumo.ToString("yyyy-MM-dd")))

                If Not objOP.Estoque Then drItemI("Estoque") += CInt(pI.Quantidade)

                CType(Session("objInsumo" & HID.Value), DataTable).Rows.Add(drItemI)
            Next

            '2 - Recuperar lista dos Insumos
            Dim objProducaoI = New ListProdutoXInsumos()

            For Each produto As OrdemParaProducaoXProduto In objOP.ProdutosdaProducao

                objProducaoI.CarregarProduto(produto.CodigoProduto)

                For Each row In objProducaoI

                    Dim temInsumo As Boolean = False

                    For Each pInsumo In objOP.ItensDeInsumo
                        If row.CodigoProdutoInsumo = pInsumo.CodigoProduto Then
                            temInsumo = True
                        End If
                    Next

                    If Not temInsumo Then
                        Dim drItemI As DataRow = CType(Session("objInsumo" & HID.Value), DataTable).NewRow()

                        drItemI("CodigoProdutoProducao") = produto.CodigoProduto
                        drItemI("Produto") = row.ProdutoInsumo.Codigo
                        drItemI("NomeProduto") = row.ProdutoInsumo.Nome
                        'drItemI("Base") = CInt(row.Base)
                        drItemI("Base") = row.Base
                        drItemI("Quantidade") = CDec(0)
                        drItemI("Estoque") = CDec(0)

                        Dim qInsumo = New OrdemParaProducaoXInsumo()

                        drItemI("Estoque") = CDec(qInsumo.buscarEstoqueProduto(ddlEmpresa.SelectedValue.Split("-")(0), ddlEmpresa.SelectedValue.Split("-")(1), row.ProdutoInsumo.Codigo, dataInsumoAnt.ToString("yyyy-MM-dd"), dataInsumo.ToString("yyyy-MM-dd")))

                        CType(Session("objInsumo" & HID.Value), DataTable).Rows.Add(drItemI)
                    End If

                Next

            Next

            For Each rowI In objProducaoI
                For Each drI As DataRow In CType(Session("objInsumo" & HID.Value), DataTable).Rows
                    If rowI.CodigoProduto = drI("CodigoProdutoProducao") And rowI.CodigoProdutoInsumo = drI("Produto") Then
                        drI("Base") = rowI.Base
                    End If
                Next
            Next

            gridInsumo.DataSource = dtgridInsumo()
            gridInsumo.DataBind()

            Dim j As Integer = 0
            While j < gridInsumo.Rows.Count

                If objOP.IUD = "U" And objOP.Estoque Then
                    'FICA DESABILITADO IZA PODER AJUSTAR A ORDEM DE PRODUÇÃO - FURLAN - 24/01/2023
                    CType(gridInsumo.Rows(j).FindControl("txtQuantidadeInsumo"), TextBox).Enabled = False
                Else
                    If CInt(gridInsumo.Rows(j).Cells(4).Text) > 0 Then
                        CType(gridInsumo.Rows(j).FindControl("txtQuantidadeInsumo"), TextBox).Enabled = True
                    Else
                        CType(gridInsumo.Rows(j).FindControl("txtQuantidadeInsumo"), TextBox).Enabled = False
                    End If
                End If

                j += 1
            End While


            If objOP.ItensDeEspecificacao.Count > 0 Then
                For Each pE In objOP.ItensDeEspecificacao
                    pE.IUD = objOP.IUD

                    Dim drEP As DataRow = CType(Session("objEspecificacaoDoProduto" & HID.Value), DataTable).NewRow()

                    drEP("CodigoProdutoProducao") = pE.CodigoProdutoProducao
                    drEP("Codigo") = pE.CodigoEspecificacao
                    drEP("Descricao") = pE.EspecificacaoDoProduto.Descricao
                    drEP("FaixaInicial") = pE.FaixaInicial
                    drEP("FaixaFinal") = pE.FaixaFinal
                    drEP("Resultado") = pE.Resultado

                    CType(Session("objEspecificacaoDoProduto" & HID.Value), DataTable).Rows.Add(drEP)

                Next

                gridEspecificacaoDoProduto.DataSource = dtgridEspecificacaoDoProduto()
                gridEspecificacaoDoProduto.DataBind()

                gridEspecificacaoDoProduto.Columns(4).Visible = True

                divEspecificacao.Visible = True
                lnkAjustaEspecificacao.Parent.Visible = True
            Else

                For Each produto As OrdemParaProducaoXProduto In objOP.ProdutosdaProducao

                    If produto.Produto.ProdutoXEspecificacao.Count > 0 Then

                        For Each eP In produto.Produto.ProdutoXEspecificacao.Where(Function(s) s.Ativo = True)
                            objOPXE = New OrdemParaProducaoXEspecificacao(objOP)

                            objOPXE.IUD = "I"
                            objOPXE.CodigoProdutoProducao = eP.CodigoProduto
                            objOPXE.CodigoEspecificacao = eP.CodigoEspecificacao
                            objOPXE.FaixaInicial = eP.FaixaInicial
                            objOPXE.FaixaFinal = eP.FaixaFinal

                            If objOPXE.Salvar Then

                                objOPXE.IUD = objOP.IUD
                                objOP.ItensDeEspecificacao.Add(objOPXE)

                                Dim drEP As DataRow = CType(Session("objEspecificacaoDoProduto" & HID.Value), DataTable).NewRow()

                                drEP("CodigoProdutoProducao") = eP.CodigoProduto
                                drEP("Codigo") = eP.CodigoEspecificacao
                                drEP("Descricao") = eP.EspecificacaoDoProduto.Descricao
                                drEP("FaixaInicial") = eP.FaixaInicial
                                drEP("FaixaFinal") = eP.FaixaFinal
                                drEP("Resultado") = 0

                                CType(Session("objEspecificacaoDoProduto" & HID.Value), DataTable).Rows.Add(drEP)
                            End If
                        Next

                        Dim dtEspecificacaoDoProduto As DataTable = dtgridEspecificacaoDoProduto()

                        If dtEspecificacaoDoProduto.Rows.Count > 0 Then
                            gridEspecificacaoDoProduto.DataSource = dtEspecificacaoDoProduto
                            gridEspecificacaoDoProduto.DataBind()

                            gridEspecificacaoDoProduto.Columns(4).Visible = True

                            divEspecificacao.Visible = True
                            lnkAjustaEspecificacao.Parent.Visible = True
                        End If
                    End If

                Next

            End If

            If objOP.ItensDeEPI.Count > 0 Then
                For Each drEPI In objOP.ItensDeEPI
                    drEPI.IUD = objOP.IUD

                    Dim drI As DataRow = CType(Session("objEPI" & HID.Value), DataTable).NewRow()

                    drI("Codigo") = drEPI.CodigoEPI
                    drI("Descricao") = drEPI.EPI.Descricao

                    CType(Session("objEPI" & HID.Value), DataTable).Rows.Add(drI)
                Next

                gridEPI.DataSource = CType(Session("objEPI" & HID.Value), DataTable)
                gridEPI.DataBind()
            End If

            If objOP.ItensDeEmabalagem.Count > 0 Then

                For Each drEmb In objOP.ItensDeEmabalagem
                    drEmb.IUD = objOP.IUD

                    Dim drE As DataRow = CType(Session("objEmbalagem" & HID.Value), DataTable).NewRow()

                    drE("Codigo") = drEmb.CodigoEmbalagem
                    drE("Descricao") = drEmb.Embalagem.Descricao
                    drE("CodigoUnidade") = drEmb.CodigoUnidade
                    drE("DescricaoUnidade") = drEmb.UnidadeDeMedida.Descricao
                    drE("Quantidade") = drEmb.Quantidade
                    drE("Capacidade") = drEmb.Capacidade

                    CType(Session("objEmbalagem" & HID.Value), DataTable).Rows.Add(drE)
                Next

                gridEmbalagem.DataSource = CType(Session("objEmbalagem" & HID.Value), DataTable)
                gridEmbalagem.DataBind()
            End If

            If objOP.ItensDeProcedimento.Count > 0 Then
                For Each drProcedimento In objOP.ItensDeProcedimento
                    drProcedimento.IUD = objOP.IUD

                    Dim drI As DataRow = CType(Session("objProcedimento" & HID.Value), DataTable).NewRow()

                    drI("Codigo") = drProcedimento.CodigoProcedimento
                    drI("Descricao") = drProcedimento.Procedimento.Descricao

                    CType(Session("objProcedimento" & HID.Value), DataTable).Rows.Add(drI)
                Next

                gridProcedimento.DataSource = CType(Session("objProcedimento" & HID.Value), DataTable)
                gridProcedimento.DataBind()
            End If

            If gridProducao.Rows.Count = 1 Then
                ddlGrupoProdutoProducao.Enabled = True
                ddlProdutosProducao.Enabled = True
                ddlUnidadeComercializacao.Enabled = True
                txtQuantidadeProducao.Enabled = True
                imgAdicionarDadosParaProducao.Enabled = True
                lnkBuscaProdutoProducao.Enabled = True
            Else
                ddlGrupoProdutoProducao.Enabled = False
                ddlProdutosProducao.Enabled = False
                ddlUnidadeComercializacao.Enabled = False
                txtQuantidadeProducao.Enabled = False
                imgAdicionarDadosParaProducao.Enabled = False
                lnkBuscaProdutoProducao.Enabled = False
            End If

            lnkConsultar.Parent.Visible = False
            lnkNovo.Parent.Visible = False

            ddlUsuarios.Items.Clear()

            If Not String.IsNullOrWhiteSpace(objOP.UsuarioCancelamento) Then
                ddlUsuarios.Items.Add("Can.- " & objOP.UsuarioCancelamento & " " & objOP.DataCancelamento.ToString("dd/MM/yyyy"))
            End If

            If Not String.IsNullOrWhiteSpace(objOP.UsuarioAlteracao) Then
                ddlUsuarios.Items.Add("Alt.- " & objOP.UsuarioAlteracao & " " & objOP.DataAlteracao.ToString("dd/MM/yyyy"))
            End If

            ddlUsuarios.Items.Add("Inc.- " & objOP.UsuarioInclusao & " " & objOP.DataInclusao.ToString("dd/MM/yyyy"))
            If objOP.CodigoSituacao = CInt(eSituacao.Excluido) Then
                lnkAtualizar.Parent.Visible = False
                lnkExcluir.Parent.Visible = False
                lnkImprimirOrdem.Parent.Visible = False
                lnkImprimirLaudo.Parent.Visible = False
                lnkImprimirLaudoManual.Parent.Visible = False
                MsgBox(Me.Page, "Ordem de Produção Excluída!", eTitulo.Info)
            Else
                lnkAtualizar.Parent.Visible = True
                lnkExcluir.Parent.Visible = True
                lnkImprimirOrdem.Parent.Visible = True
                lnkImprimirLaudo.Parent.Visible = True
                lnkImprimirLaudoManual.Parent.Visible = False
            End If

            lnkAjustaProducao.Parent.Visible = bAlterarOrdemDeProducao
            lnkAjustaConsumo.Parent.Visible = bAlterarOrdemDeProducao

            If objOP.Estoque Then

                If Not Funcoes.VerificaAcesso(objOP.CodigoEmpresa, objOP.EnderecoEmpresa, txtDataEstoque.Text, "PRODUCAO") Then

                    lnkAtualizar.Parent.Visible = False
                    lnkExcluir.Parent.Visible = False

                    rbNao.Enabled = False
                    rbSim.Enabled = False

                    txtDataEstoque.Enabled = False

                    lnkAjustaProducao.Parent.Visible = False
                    lnkAjustaConsumo.Parent.Visible = False
                    imgConfirmar.Visible = False

                    TabIncluirParaConsumo.Visible = False

                    MsgBox(Me.Page, "Movimento de Produção já Fechado, apenas Impressão ou Ajuste de Especificações.")

                End If

            End If

            SessaoSalvaOrdem()

            Return True
        End If
    End Function

    Private Sub ImprimirOrdem(Empresa As String, ByVal EndEmpresa As String, ByVal Ordem As Integer)
        'SessaoRecuperaOrdem()

        objOP = New OrdemParaProducao(Empresa, EndEmpresa, Ordem)

        If String.IsNullOrWhiteSpace(objOP.ProdutosdaProducao.Count = 0) Then
            MsgBox(Me.Page, "Ordem de Produção não foi selecionada!", eTitulo.Info)
            Exit Sub
        End If

        Dim html = New StringBuilder()

        html.Append("<html>")
        html.Append("<head>")
        html.Append("<meta http-equiv='Content-Type' content='text/html; charset=utf-8'>")
        html.Append("<title>Ordem de Produção</title>")
        html.Append("<style type='text/css'>")

        'styles
        html.Append("@page { size: A4; margin: 11mm 17mm 17mm 17mm; }")
        html.Append("@media print { html, body { width: 210mm; height: 297mm; padding-top: 15; } }")
        html.Append("html, body { width: 99.99%; height: 100%; margin: 0; padding-top: 15; }")
        html.Append("body { display: flex; flex-direction: column; }")
        html.Append("main { width: 100%; flex-grow: 1; }")
        html.Append(".table-report { font-family: 'Times New Roman', Times, serif; border-collapse: collapse; width: 100%; font-size: 12px; }")
        html.Append(".table-report td, .table-report th { border-right: 0.2px solid#696969; border-left: 0.2px solid #696969; border-bottom: 0.2px solid #696969; padding: 3px; }")
        html.Append(".table-report tr > th { background-color: #f2f2f2; }")
        html.Append(".table-report th { text-align: left; color: black; }")
        html.Append(".table-report tr { page-break-inside: avoid; }")
        html.Append("ol { padding: 5px; }")
        html.Append("ol li { padding: 3px; margin-left: 15px; }")
        html.Append("ul { padding: 5px; list-style-type: none; color: black; }")
        html.Append("ul li { margin: 5px; }")
        html.Append(".border-top { border-top: 0.2px solid #696969; }")
        html.Append(".td-align-top { font-weight: bold; text-align: start; vertical-align:top!important; width: 60%; } ")
        html.Append("p { padding: 3px; margin: 0px; } ")
        html.Append(".td-align-bottom { text-align: End; vertical-align:top!important; font-size: 8px; width: 20%;} ")
        html.Append(".font-size-14 { font-size: 14px; }")
        html.Append(".font-size-16 { font-size: 16px; }")
        html.Append(".width-50 {width: 50%; } ")
        html.Append(".content {position: relative; min-height: 100px; font-size: 8px; } ")
        html.Append(".content-bottom { position: absolute; bottom: 0; right: 0; } ")
        html.Append(".content-top { position: absolute; top: 0; right: 0; } ")
        'end styles
        html.Append("</style>")
        html.Append("</head>")
        html.Append("<body>")
        html.Append("<main>")

        Dim link As String = "http://www.ngssolucoes.com.br/Download/UltimaVersao/" & Session("ssImagemEmpresa")

        'Cabeçalho
        html.Append("<table class='table-report'>")
        html.Append("<tr Class='border-top'>")
        html.Append("<td> <img src='" & link & "' width='200' height='100'></td>")
        html.Append("<td Class='td-align-top'>")
        html.Append("<p Class='padding-p'>" & Session("ssNomeEmpresa") & "</p>")
        html.Append("<p Class='padding-p'>" & Session("ssCidadeEmpresa") & "/" & Session("ssEstadoEmpresa") & "</p>")
        html.Append("<p Class='padding-p'>ORDEM DE PRODUÇÃO: " & objOP.Codigo.ToString() & "</p>")

        If objOP.CodigoSituacao = eSituacao.Excluido Then
            html.Append("<p Class='padding-p' style='color:#FF0000'>***** ORDEM CANCELADA *****</p>")
        End If

        html.Append("</td>")
        html.Append("<td style='width: 20%;'>")
        html.Append("<div class='content'>")
        html.Append("<div class='content-top'>")

        If String.IsNullOrWhiteSpace(objOP.UsuarioAlteracao) Then
            html.Append("<p>Usuário: " & objOP.UsuarioInclusao & "</p>")
        Else
            html.Append("<p>Usuário: " & objOP.UsuarioAlteracao & "</p>")
        End If

        html.Append("</div>")
        html.Append("<div class='content-bottom'>")

        If String.IsNullOrWhiteSpace(objOP.UsuarioAlteracao) Then
            html.Append("<p>Data: " & objOP.DataInclusao.ToString("dd/MM/yyyy") & "</p>")
            html.Append("<p>Hora: " & objOP.DataInclusao.ToString("HH:mm") & "</p>")
        Else
            html.Append("<p>Data: " & objOP.DataAlteracao.ToString("dd/MM/yyyy") & "</p>")
            html.Append("<p>Hora: " & objOP.DataAlteracao.ToString("HH:mm") & "</p>")
        End If

        html.Append("</div>")
        html.Append("</div>")

        html.Append("</td>")
        html.Append("</tr>")
        html.Append("</table>")

        For Each produto As OrdemParaProducaoXProduto In objOP.ProdutosdaProducao

            'Produto
            html.Append("<table class='table-report'>")

            html.Append("<tr>")
            html.Append("<th colspan='7' class='font-size-14'>PRODUTO: " & produto.Produto.Nome & "</th>")
            html.Append("</tr>")

            html.Append("<tr>")
            html.Append("<td class='font-size-16'>Lote Nº</td>")
            html.Append("<td>Data de Produção</td>")
            html.Append("<td>Validade</td>")
            html.Append("<td>Hora de Início</td>")
            html.Append("<td>Hora de Término</td>")
            html.Append("<td>Data de Término</td>")
            html.Append("<td>Quantidade</td>")

            html.Append("</tr>")
            html.Append("<tr>")
            html.Append("<td class='font-size-16'>" & produto.Lote & "</td>")
            html.Append("<td>" & objOP.Movimento.ToString("dd/MM/yyyy") & "</td>")
            html.Append("<td>" & objOP.Validade.ToString("dd/MM/yyyy") & "</td>")

            html.Append("<td></td>")
            html.Append("<td></td>")
            html.Append("<td></td>")

            If produto.QuantidadeDeAjuste > 0 Then
                html.Append("<td>" & produto.QuantidadeDeAjuste.ToString() & "</td>")
            Else
                html.Append("<td>" & produto.Quantidade.ToString() & "</td>")
            End If

            html.Append("</tr>")

            html.Append("</table>")

            'Consumo
            html.Append("<table class='table-report'>")
            html.Append("<tr>")
            html.Append("<th colspan='7'> CONSUMO</th>")
            html.Append("</tr>")
            html.Append("<tr>")
            html.Append("<th></th>")
            html.Append("<th class='font-size-16'> Código</th>")
            html.Append("<th class='font-size-16'> Seqüência de adição de matéria-primas</th>")
            html.Append("<th class='font-size-16'> Lote Fornecedor</th>")
            'html.Append("<th style='text-align:right;'>% </th>")
            html.Append("<th style='text-align:right;'>Quantidade </th>")
            html.Append("<th style='text-align:right;'>Ajuste </th>")
            html.Append("</tr>")

            Dim i As Integer = 1
            Dim totalPecentual As Decimal = 0
            Dim totalKg As Decimal = 0

            For Each consumo In objOP.ItensDeConsumo.Where(Function(x) x.CodigoProdutoProducao = produto.CodigoProduto)
                i = (i + 1)
                html.Append("<tr>")
                html.Append("<td>" & i & "</td>")
                html.Append("<td class='font-size-16'>" & consumo.CodigoProduto & "</td>")
                html.Append("<td class='font-size-16'>" & consumo.Produto.Nome & "</td>")

                html.Append("<td class='font-size-16'>") 'Lote
                For Each nLote In consumo.ItensDeConsumoXLote
                    html.Append(nLote.Lote)
                    html.Append("<br />")
                Next
                html.Append("</td>") 'Lote

                'html.Append("<td style='text-align:right;'>" & consumo.Percentual & " </td>") '%
                html.Append("<td style='text-align:right; font-size: 16px;'>" & consumo.Quantidade & " </td>") 'Kg

                If consumo.Sinal = "+" Then
                    html.Append("<td style='text-align:right; font-size: 16px;'>" & consumo.QuantidadeDeAjuste & " </td>") 'Ajuste
                    totalKg += (consumo.Quantidade + consumo.QuantidadeDeAjuste)
                ElseIf consumo.Sinal = "-" Then
                    html.Append("<td style='text-align:right; font-size: 16px;'>" & consumo.QuantidadeDeAjuste & "(-) </td>") 'Ajuste
                    totalKg += (consumo.Quantidade - consumo.QuantidadeDeAjuste)
                End If

                html.Append("</tr>")

                totalPecentual += consumo.Percentual 'Total %
            Next

            html.Append("<tr>")
            html.Append("<td></td>")
            html.Append("<td></td>")
            html.Append("<td></td>")
            'html.Append("<td> TOTAL</td>")
            html.Append("<td></td>")
            'html.Append("<td style='text-align:right;'>" & totalPecentual & " </td>")
            html.Append("<td colspan='2' style='text-align:right; font-size: 16px;'>" & totalKg & " </td>")
            'html.Append("<td></td>")
            html.Append("</tr>")
            html.Append("</table>")

        Next

        'Especificações
        If objOP.ItensDeEspecificacao.Count > 0 Then
            html.Append("<table class='table-report'>")
            html.Append("<tr>")
            html.Append("<th>ESPECIFICAÇÕES</th>")
            html.Append("<th>ANÁLISE 1</th>")
            html.Append("<th>ANÁLISE 2</th>")
            html.Append("<th>ANÁLISE 3</th>")
            html.Append("<th>ANÁLISE 4</th>")
            html.Append("<th style='text-align:right;'>RESULTADOS ANALÍTICOS </th>")
            html.Append("</tr>")


            Dim prdEspecificacao = ""

            For Each ep In objOP.ItensDeEspecificacao

                If Not ep.CodigoProdutoProducao = prdEspecificacao Then
                    html.Append("<tr>")
                    html.Append("<td colspan='6' style='font-weight: bold;'> " & ep.ProdutoProducao.Nome & "</td>")
                    html.Append("</tr>")
                    prdEspecificacao = ep.CodigoProdutoProducao
                End If

                Dim descricao = ep.EspecificacaoDoProduto.Descricao & " " & ep.FaixaInicial & " - " & ep.FaixaFinal
                html.Append("<tr>")
                html.Append("<td> " & descricao & "</td>")
                html.Append("<td></td>")
                html.Append("<td></td>")
                html.Append("<td></td>")
                html.Append("<td></td>")
                html.Append("<td style='text-align:right;'>" & ep.Resultado & " </td>")
                html.Append("</tr>")
            Next

            html.Append("</table>")
        End If


        'Embalagem
        If objOP.ItensDeEmabalagem.Count > 0 Then
            html.Append("<table class='table-report'>")
            html.Append("<tr>")
            html.Append("<th colspan='4'> EMBALAGEM </th>")
            html.Append("</tr>")
            html.Append("<tr>")
            html.Append("<th> Descrição</th>")
            html.Append("<th> Quantidade</th>")
            html.Append("<th> Unidade</th>")
            html.Append("<th> Capacidade</th>")
            html.Append("</tr>")

            For Each emb In objOP.ItensDeEmabalagem
                html.Append("<tr>")
                html.Append("<td>" & emb.CodigoUnidade & " - " & emb.Embalagem.Descricao & "</td>")
                html.Append("<td>" & emb.Quantidade & "</td>")
                html.Append("<td>" & emb.CodigoUnidade & " - " & emb.Embalagem.Descricao & "</td>")
                html.Append("<td>" & emb.Capacidade & "</td>")
                html.Append("</tr>")
            Next

            html.Append("</table>")
        End If


        'EPI'
        If objOP.ItensDeEPI.Count > 0 Then
            html.Append("<table class='table-report'>")

            html.Append("<tr>")
            html.Append("<th colspan='3'> EPI</th>")
            html.Append("</tr>")

            For Each epi In objOP.ItensDeEPI
                html.Append("<tr>")
                html.Append("<td colspan='3'>" & epi.EPI.Descricao & "</td>")
                html.Append("</tr>")
            Next

            html.Append("</table>")
        End If


        'Procedimento
        If objOP.ItensDeProcedimento.Count > 0 Then
            html.Append("<table class='table-report'>")
            html.Append("<tr>")
            html.Append("<th colspan='3'> PROCEDIMENTO</th>")
            html.Append("</tr>")
            html.Append("<tr>")
            'html.Append("<td colspan='3'>")
            'html.Append("<ol>")

            For Each procedimento In objOP.ItensDeProcedimento
                html.Append("<tr>")
                html.Append("<td colspan='3'>" & procedimento.Procedimento.Descricao & "</td>")
                html.Append("</tr>")
            Next

            'html.Append("<li>Colocar em um container limpo 600 Kg de água gelada</li>")
            'html.Append("<li>adicionar lentamente 600 kg de ácido sulfurico 98% Deixar esfriar a té pelo menos 40 celsius	</li>")
            'html.Append("<li>Repetira sequencia dos itens 1 e 2 em novo container em  mais 5.0 container </li>")
            'html.Append("<li>Transferir o conteudo resfriado de todos os containers para o reator  de ácido </li>")
            'html.Append("<li>Homogeinizar por pelo menos trinta minutos</li>")
            'html.Append("<li>Retirar amostra para o controle de qualidade e aguardar instruções finais.</li>")

            'html.Append("</ol>")
            'html.Append("</td>")
            'html.Append("</tr>")
            html.Append("</table>")
        End If

        'Observações
        html.Append("<table class='table-report'>")
        html.Append("<tr>")
        html.Append("<th classcolspan='3'> OBERVAÇÕES</th>")
        html.Append("</tr>")
        html.Append("<tr>")
        html.Append("<td height='100px' colspan='3'>")
        html.Append(objOP.Observacoes)
        html.Append("</td>")
        html.Append("</tr>")
        html.Append("</table>")

        'Responsável Produção
        html.Append("<table class='table-report'>")

        html.Append("<tr>")
        html.Append("<th class='width-50' colspan='3'>RESPONSÁVEL PRODUÇÃO</th>")
        html.Append("<th colspan='3'> APROVAÇÃO CONTROLE DE QUALIDADE</th>")
        html.Append("</tr>")

        html.Append("<tr>")
        html.Append("<td colspan='3'>")
        html.Append("&nbsp;")
        html.Append("</td>")
        html.Append("<td colspan='3'>")
        html.Append("&nbsp;")
        html.Append("</td>")
        html.Append("</tr>")

        html.Append("</table>")

        html.Append("</main>")
        html.Append("</body>")
        html.Append("</html>")

        Dim strCaminho As String = Server.MapPath("~/Files/OrdemDeProducao.html")
        If Dir(strCaminho).Length > 0 Then Kill(strCaminho)

        Using strm As New StreamWriter(Server.MapPath("~/Files/OrdemDeProducao.html"), True)
            strm.WriteLine(html)
            strm.Close()
            strm.Dispose()
        End Using

        Dim strNomeArquivo As String = "Files/OrdemDeProducao.pdf"
        Dim pathPDF = Server.MapPath(strNomeArquivo)
        If Dir(pathPDF).Length > 0 Then Kill(pathPDF)

        Dim htmlContent As String = html.ToString()

        Dim generator = New NReco.PdfGenerator.HtmlToPdfConverter()
        generator.PageFooterHtml = "<div style=""width: 100%; text-align: center; padding-top: 10px; display: inline-block;""><img src=""http://localhost:4586/Images/logo_ngs_rpt.jpg"" width=""60"" height=""20""></div>"

        Dim pdf = generator.GeneratePdf(htmlContent)

        Using fs As FileStream = New FileStream(pathPDF, FileMode.Create)
            fs.Write(pdf, 0, pdf.Length)
            fs.Close()
        End Using

        'If IO.File.Exists(pathPDF) Then
        '    Response.Clear()
        '    Response.ContentType = "application/pdf"
        '    Response.AppendHeader("Content-Disposition", "attachment; filename=" & "OrdemDeProducao " & objOP.Codigo & ".pdf")
        '    Response.TransmitFile(Server.MapPath("~/Files/OrdemDeProducao.pdf"))
        '    Response.End()
        'End If

        ScriptManager.RegisterClientScriptBlock(Me.Page, Me.Page.GetType(), Guid.NewGuid().ToString, "window.open('" & strNomeArquivo & "');", True)

    End Sub

    Private Sub ImprimirLaudo()
        SessaoRecuperaOrdem()

        If String.IsNullOrWhiteSpace(objOP.ProdutosdaProducao.Count = 0) Then
            MsgBox(Me.Page, "Ordem de Produção não foi selecionada!", eTitulo.Info)
            Exit Sub
        ElseIf objOP.ItensDeEspecificacao.Count = 0 Then
            MsgBox(Me.Page, "Ordem de Produção não tem Especificações!", eTitulo.Info)
            Exit Sub
        End If

        Dim html = New StringBuilder()

        html.Append("<html>")
        html.Append("<head>")
        html.Append("<meta http-equiv='Content-Type' content='text/html; charset=utf-8'>")
        html.Append("<title>Laudo de Análises</title>")
        html.Append("<style type='text/css'>")

        'styles
        html.Append("@page { size: A4; margin: 11mm 17mm 17mm 17mm; }")
        html.Append("@media print { html, body { width: 210mm; height: 297mm; padding-top: 15; } }")
        html.Append("html, body { width: 99.99%; height: 100%; margin: 0; padding-top: 15; }")
        html.Append("body { display: flex; flex-direction: column; }")
        html.Append("main { width: 100%; flex-grow: 1; }")
        html.Append(".table-report { font-family: 'Times New Roman', Times, serif; border-collapse: collapse; width: 100%; font-size: 12px; }")
        html.Append(".table-report td, .table-report th { border-right: 0.2px solid#696969; border-left: 0.2px solid #696969; border-bottom: 0.2px solid #696969; padding: 3px; }")
        html.Append(".table-report tr > th { background-color: #f2f2f2; }")
        html.Append(".table-report th { text-align: left; color: black; }")
        html.Append(".table-report tr { page-break-inside: avoid; }")
        html.Append("ol { padding: 5px; }")
        html.Append("ol li { padding: 3px; margin-left: 15px; }")
        html.Append("ul { padding: 5px; list-style-type: none; color: black; }")
        html.Append("ul li { margin: 5px; }")
        html.Append(".border-top { border-top: 0.2px solid #696969; }")
        html.Append(".td-align-top { font-weight: bold; text-align: start; vertical-align:top!important; width: 60%; } ")
        html.Append("p { padding: 3px; margin: 0px; } ")
        html.Append(".td-align-bottom { text-align: End; vertical-align:top!important; font-size: 8px; width: 20%;} ")
        html.Append(".font-size-14 { font-size: 14px; }")
        html.Append(".width-50 {width: 50%; } ")
        html.Append(".content {position: relative; min-height: 100px; font-size: 8px; } ")
        html.Append(".content-bottom { position: absolute; bottom: 0; right: 0; } ")
        html.Append(".content-top { position: absolute; top: 0; right: 0; } ")
        'end styles
        html.Append("</style>")
        html.Append("</head>")
        html.Append("<body>")
        html.Append("<main>")

        Dim link As String = "http://www.ngssolucoes.com.br/Download/UltimaVersao/" & objOP.Empresa.Imagem

        'Cabeçalho
        html.Append("<table class='table-report'>")
        html.Append("<tr Class='border-top'>")
        html.Append("<td> <img src='" & link & "' width='200' height='100'></td>")
        html.Append("<td Class='td-align-top'>")
        html.Append("<p Class='padding-p'>" & objOP.Empresa.Nome & "</p>")
        html.Append("<p Class='padding-p'>" & objOP.Empresa.Cidade & "/" & objOP.Empresa.CodigoEstado & "</p>")
        html.Append("<p Class='padding-p'>" & objOP.Empresa.OutrosTelefones & "</p>")

        If objOP.CodigoSituacao = eSituacao.Excluido Then
            html.Append("<p Class='padding-p'>" & objOP.Empresa.Email & "</p>")
            html.Append("<p Class='padding-p' style='color:#FF0000'>***** ORDEM CANCELADA *****</p>")
        Else
            html.Append("<p Class='padding-p'>" & objOP.Empresa.Email & "</p>")
        End If

        html.Append("</td>")
        'html.Append("<td style='width: 20%;'>")
        'html.Append("<div class='content'>")
        'html.Append("<div class='content-top'>")
        'html.Append("<p>Usuário: " & IIf(String.IsNullOrWhiteSpace(objOP.UsuarioAlteracao), objOP.UsuarioInclusao, objOP.UsuarioAlteracao) & "</p>")
        'html.Append("</div>")
        'html.Append("<div class='content-bottom'>")
        'html.Append("<p>Data: " & Now().ToString("dd/MM/yyyy") & "</p>")
        'html.Append("<p>Hora: " & Now().ToString("HH:mm") & "</p>")
        'html.Append("</div>")
        'html.Append("</div>")
        'html.Append("</td>")
        html.Append("</tr>")
        html.Append("</table>")

        html.Append("<br>")
        html.Append("<br>")

        html.Append("<table class='table-report'>")
        html.Append("<tr style='border-top: 0.2px solid #696969;'>")
        html.Append("<th colspan='4' class='font-size-14' style='text-align:center;'>LAUDO DE ANÁLISES</th>")
        html.Append("</tr>")
        html.Append("</table>")

        html.Append("<br>")
        html.Append("<br>")

        'Produto
        html.Append("<table class='table-report'>")

        For Each produto As OrdemParaProducaoXProduto In objOP.ProdutosdaProducao

            html.Append("<tr style='border-top: 0.2px solid #696969;'>")
            html.Append("<th colspan='4' class='font-size-14'>PRODUTO: " & produto.Produto.Nome & "</th>")
            html.Append("</tr>")

            html.Append("<tr>")
            html.Append("<td>Lote Nº</td>")
            html.Append("<td>Data</td>")
            html.Append("<td>Validade</td>")
            html.Append("</tr>")

            html.Append("<tr>")
            html.Append("<td>" & produto.Lote & "</td>")
            html.Append("<td>" & objOP.Movimento.ToString("dd/MM/yyyy") & "</td>")
            html.Append("<td>" & objOP.Validade.ToString("dd/MM/yyyy") & "</td>")
            html.Append("</tr>")

        Next

        html.Append("</table>")

        html.Append("<br>")
        html.Append("<br>")

        'Especificações
        html.Append("<table class='table-report'>")
        html.Append("<tr style='border-top: 0.2px solid #696969;'>")
        html.Append("<th> ESPECIFICAÇÕES</th>")
        html.Append("<th style='text-align:center;'>Faixa Inicial</th>")
        html.Append("<th style='text-align:center;'>Faixa Final</th>")
        html.Append("<th style='text-align:right;'> RESULTADOS</th>")
        html.Append("</tr>")

        Dim prdEspecificacao = ""

        For Each ep In objOP.ItensDeEspecificacao

            If Not ep.CodigoProdutoProducao = prdEspecificacao Then
                html.Append("<tr>")
                html.Append("<td colspan='4' style='font-weight: bold;'> " & ep.ProdutoProducao.Nome & "</td>")
                html.Append("</tr>")
                prdEspecificacao = ep.CodigoProdutoProducao
            End If

            Dim descricao = ep.EspecificacaoDoProduto.Descricao
            Dim faixaInicial = ep.FaixaInicial
            Dim faixaFinal = ep.FaixaFinal
            Dim resultado = ep.Resultado
            html.Append("<tr>")
            html.Append("<td> " & descricao & "</td>")
            html.Append("<td> " & faixaInicial & "</td>")
            html.Append("<td> " & faixaFinal & "</td>")
            html.Append("<td style='text-align:right;'>" & resultado & " </td>")
            html.Append("</tr>")
        Next

        html.Append("</table>")

        html.Append("<br>")
        html.Append("<br>")
        html.Append("<br>")
        html.Append("<br>")
        html.Append("<br>")
        html.Append("<br>")

        html.Append("<table>")
        html.Append("<tr>")
        html.Append("<td style='text-align:center;'>_________________________________________</td>")
        html.Append("</tr>")
        html.Append("<tr>")
        html.Append("<td style='text-align:center;'>CONTROLE DE QUALIDADE</td>")
        html.Append("</tr>")
        html.Append("</table>")

        html.Append("</main>")
        html.Append("</body>")
        html.Append("</html>")

        Dim strCaminho As String = Server.MapPath("~/Files/LaudoDeAnalises.html")
        If Dir(strCaminho).Length > 0 Then Kill(strCaminho)

        Using strm As New StreamWriter(Server.MapPath("~/Files/LaudoDeAnalises.html"), True)
            strm.WriteLine(html)
            strm.Close()
            strm.Dispose()
        End Using

        Dim strNomeArquivo As String = "Files/LaudoDeAnalises.pdf"
        Dim pathPDF = Server.MapPath(strNomeArquivo)
        If Dir(pathPDF).Length > 0 Then Kill(pathPDF)

        Dim htmlContent As String = html.ToString()

        Dim generator = New NReco.PdfGenerator.HtmlToPdfConverter()

        generator.PageFooterHtml = "<div style=""width: 100%; text-align: center; padding-top: 10px; display: inline-block;""><img src=""http://localhost:4586/Images/logo_ngs_rpt.jpg"" width=""60"" height=""20""></div>"

        Dim pdf = generator.GeneratePdf(htmlContent)

        Using fs As FileStream = New FileStream(pathPDF, FileMode.Create)
            fs.Write(pdf, 0, pdf.Length)
            fs.Close()
        End Using

        ScriptManager.RegisterClientScriptBlock(Me.Page, Me.Page.GetType(), Guid.NewGuid().ToString, "window.open('" & strNomeArquivo & "');", True)

    End Sub

    Public Sub ImprimirLaudoManual()

        Dim html = New StringBuilder()

        html.Append("<html>")
        html.Append("<head>")
        html.Append("<meta http-equiv='Content-Type' content='text/html; charset=utf-8'>")
        html.Append("<title>Laudo de Análises</title>")
        html.Append("<style type='text/css'>")

        'styles
        html.Append("@page { size: A4; margin: 11mm 17mm 17mm 17mm; }")
        html.Append("@media print { html, body { width: 210mm; height: 297mm; padding-top: 15; } }")
        html.Append("html, body { width: 99.99%; height: 100%; margin: 0; padding-top: 15; }")
        html.Append("body { display: flex; flex-direction: column; }")
        html.Append("main { width: 100%; flex-grow: 1; }")
        html.Append(".table-report { font-family: 'Times New Roman', Times, serif; border-collapse: collapse; width: 100%; font-size: 12px; }")
        html.Append(".table-report td, .table-report th { border-right: 0.2px solid#696969; border-left: 0.2px solid #696969; border-bottom: 0.2px solid #696969; padding: 3px; }")
        html.Append(".table-report tr > th { background-color: #f2f2f2; }")
        html.Append(".table-report th { text-align: left; color: black; }")
        html.Append(".table-report tr { page-break-inside: avoid; }")
        html.Append("ol { padding: 5px; }")
        html.Append("ol li { padding: 3px; margin-left: 15px; }")
        html.Append("ul { padding: 5px; list-style-type: none; color: black; }")
        html.Append("ul li { margin: 5px; }")
        html.Append(".border-top { border-top: 0.2px solid #696969; }")
        html.Append(".td-align-top { font-weight: bold; text-align: start; vertical-align:top!important; width: 60%; } ")
        html.Append("p { padding: 3px; margin: 0px; } ")
        html.Append(".td-align-bottom { text-align: End; vertical-align:top!important; font-size: 8px; width: 20%;} ")
        html.Append(".font-size-14 { font-size: 14px; }")
        html.Append(".width-50 {width: 50%; } ")
        html.Append(".content {position: relative; min-height: 100px; font-size: 8px; } ")
        html.Append(".content-bottom { position: absolute; bottom: 0; right: 0; } ")
        html.Append(".content-top { position: absolute; top: 0; right: 0; } ")
        'end styles
        html.Append("</style>")
        html.Append("</head>")
        html.Append("<body>")
        html.Append("<main>")

        Dim objEmpresa As Cliente = New Cliente(ddlEmpresa.SelectedValue.ToString.Split("-")(0), ddlEmpresa.SelectedValue.ToString.Split("-")(1))

        Dim link As String = "http://www.ngssolucoes.com.br/Download/UltimaVersao/" & objEmpresa.Imagem

        'Cabeçalho
        html.Append("<table class='table-report'>")
        html.Append("<tr Class='border-top'>")
        html.Append("<td> <img src='" & link & "' width='200' height='100'></td>")
        html.Append("<td Class='td-align-top'>")
        html.Append("<p Class='padding-p'>" & objEmpresa.Nome & "</p>")
        html.Append("<p Class='padding-p'>" & objEmpresa.Cidade & "/" & objEmpresa.CodigoEstado & "</p>")
        html.Append("<p Class='padding-p'>" & objEmpresa.OutrosTelefones & "</p>")

        If objEmpresa.CodigoSituacao = eSituacao.Excluido Then
            html.Append("<p Class='padding-p'>" & objEmpresa.Email & "</p>")
            html.Append("<p Class='padding-p' style='color:#FF0000'>***** ORDEM CANCELADA *****</p>")
        Else
            html.Append("<p Class='padding-p'>" & objEmpresa.Email & "</p>")
        End If

        Dim objUsuario As New [Lib].Negocio.Usuario(HttpContext.Current.Session("ssNomeUsuario"))

        html.Append("</td>")
        'html.Append("<td style='width: 20%;'>")
        'html.Append("<div class='content'>")
        'html.Append("<div class='content-top'>")
        'html.Append("<p>Usuário: " & IIf(String.IsNullOrWhiteSpace(objUsuario.Usuario_Id), objUsuario.Usuario_Id, objUsuario.Usuario_Id) & "</p>")
        'html.Append("</div>")
        'html.Append("<div class='content-bottom'>")
        'html.Append("<p>Data: " & Now().ToString("dd/MM/yyyy") & "</p>")
        'html.Append("<p>Hora: " & Now().ToString("HH:mm") & "</p>")
        'html.Append("</div>")
        'html.Append("</div>")
        'html.Append("</td>")
        html.Append("</tr>")
        html.Append("</table>")

        html.Append("<br>")
        html.Append("<br>")

        'Produto
        Dim objProduto As DataRow = CType(Session("objLoteFornecedor" & HID.Value), DataTable).Rows(0)
        Dim prd As Produto = New Produto(objProduto("Produto"))

        html.Append("<table class='table-report'>")
        html.Append("<tr style='border-top: 0.2px solid #696969;'>")
        html.Append("<th colspan='4' class='font-size-14' style='text-align:center;'>LAUDO DE ANÁLISES</th>")
        html.Append("</tr>")
        html.Append("</table>")

        html.Append("<br>")
        html.Append("<br>")

        html.Append("<table class='table-report'>")
        html.Append("<tr style='border-top: 0.2px solid #696969;'>")
        html.Append("<th colspan='4' class='font-size-14'>PRODUTO: " & prd.Codigo & " - " & prd.Nome & "</th>")
        html.Append("</tr>")

        html.Append("<tr>")
        html.Append("<td>Lote Nº</td>")
        html.Append("<td>Data</td>")
        html.Append("<td>Validade</td>")
        html.Append("</tr>")

        html.Append("<tr>")
        html.Append("<td>" & objProduto("Lote") & "</td>")
        html.Append("<td>" & objProduto("Fabricado") & "</td>")
        html.Append("<td>" & objProduto("Validade") & "</td>")
        html.Append("</tr>")
        html.Append("</table>")

        html.Append("<br>")
        html.Append("<br>")

        'Especificações
        html.Append("<table class='table-report'>")
        html.Append("<tr style='border-top: 0.2px solid #696969;'>")
        html.Append("<th> ESPECIFICAÇÕES</th>")
        html.Append("<th style='text-align:center;'>REFERÊNCIAS</th>")
        html.Append("<th style='text-align:right;'> RESULTADOS</th>")
        html.Append("</tr>")

        For Each ep In CType(Session("objEspecificacao" & HID.Value), DataTable).Rows
            Dim descricao = ep("Descricao")
            Dim referencia = ep("Referencia")
            Dim resultado = ep("ResultadoTxt")
            html.Append("<tr>")
            html.Append("<td> " & descricao & "</td>")
            html.Append("<td style='text-align:right;'>" & referencia & " </td>")
            html.Append("<td style='text-align:right;'>" & resultado & " </td>")
            html.Append("</tr>")
        Next

        html.Append("</table>")

        html.Append("<br>")
        html.Append("<br>")
        html.Append("<br>")
        html.Append("<br>")
        html.Append("<br>")
        html.Append("<br>")

        html.Append("<table>")
        html.Append("<tr>")
        html.Append("<td style='text-align:center;'>_________________________________________</td>")
        html.Append("</tr>")
        html.Append("<tr>")
        html.Append("<td style='text-align:center;'>CONTROLE DE QUALIDADE</td>")
        html.Append("</tr>")
        html.Append("</table>")

        html.Append("</main>")
        html.Append("</body>")
        html.Append("</html>")

        Dim strCaminho As String = Server.MapPath("~/Files/LaudoDeAnalises.html")
        If Dir(strCaminho).Length > 0 Then Kill(strCaminho)

        Using strm As New StreamWriter(Server.MapPath("~/Files/LaudoDeAnalises.html"), True)
            strm.WriteLine(html)
            strm.Close()
            strm.Dispose()
        End Using

        Dim strNomeArquivo As String = "Files/LaudoDeAnalises.pdf"
        Dim pathPDF = Server.MapPath(strNomeArquivo)
        If Dir(pathPDF).Length > 0 Then Kill(pathPDF)

        Dim htmlContent As String = html.ToString()

        Dim generator = New NReco.PdfGenerator.HtmlToPdfConverter()

        generator.PageFooterHtml = "<div style=""width: 100%; text-align: center; padding-top: 10px; display: inline-block;""><img src=""http://localhost:4586/Images/logo_ngs_rpt.jpg"" width=""60"" height=""20""></div>"

        Dim pdf = generator.GeneratePdf(htmlContent)

        Using fs As FileStream = New FileStream(pathPDF, FileMode.Create)
            fs.Write(pdf, 0, pdf.Length)
            fs.Close()
        End Using

        ScriptManager.RegisterClientScriptBlock(Me.Page, Me.Page.GetType(), Guid.NewGuid().ToString, "window.open('" & strNomeArquivo & "');", True)

    End Sub

    Protected Sub lnkConsultar_Click(sender As Object, e As EventArgs) Handles lnkConsultar.Click
        Try
            If Funcoes.VerificaPermissao("OrdemDeProducao", "LEITURA") Then
                If ValidarSelecao() Then
                    If Not String.IsNullOrWhiteSpace(txtSequencia.Text) AndAlso CInt(txtSequencia.Text) > 0 Then
                        objOP = New OrdemParaProducao(ddlEmpresa.SelectedValue.Split("-")(0), ddlEmpresa.SelectedValue.Split("-")(1), CInt(txtSequencia.Text))
                        RecuperarOrdem()
                    Else
                        ucConsultaOrdemDeProducao.Limpar()
                        Dim txtSequencia As TextBox = CType(ucConsultaOrdemDeProducao.FindControlRecursive("txtSequencia"), TextBox)
                        Popup.ConsultaOrdemDeProducao(Me.Page, "objConsultaxOrdem" & HID.Value, txtSequencia.ClientID, True)
                    End If
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para consultar registro!", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkNovo_Click(sender As Object, e As EventArgs) Handles lnkNovo.Click

        If Funcoes.VerificaPermissao("OrdemDeProducao", "GRAVAR") Then

            If ValidarSelecao() Then
                objOP = New OrdemParaProducao

                objOP.IUD = "I"
                objOP.CodigoEmpresa = ddlEmpresa.SelectedValue.Split("-")(0)
                objOP.EnderecoEmpresa = ddlEmpresa.SelectedValue.Split("-")(1)

                'FURLAN - 18/10/2023
                '############################################
                '### ATUALIZAR SEQUÊNCIAS ANTES DE GRAVAR ###
                '############################################

                'Atualiza Sequência do Numerador
                verNumerador()

                objOP.Codigo = CInt(txtSequencia.Text)

                'Atualiza Sequência do Lote
                Dim produtoXLote = New ProdutoXSequenciaDeLote(CType(Session("objProducao" & HID.Value), DataTable).Rows(0).Item("Produto").ToString(), CType(txtDataProducao.Text, DateTime).Year)

                If String.IsNullOrWhiteSpace(produtoXLote.SequenciaDoProduto) Then
                    ddlProdutosProducao.SelectedIndex = 0
                    MsgBox(Me.Page, "Sequência do Lote do Produto não foi encontrada.", eTitulo.Info)
                    Exit Sub
                End If

                '############################################
                '###                ATÉ AQUI              ###
                '############################################

                objOP.Movimento = CDate(txtDataProducao.Text)
                objOP.Validade = CDate(txtDataValidade.Text)
                objOP.CodigoSituacao = CInt(eSituacao.Normal)

                For Each drProduto As DataRow In CType(Session("objProducao" & HID.Value), DataTable).Rows

                    objOPXP = New OrdemParaProducaoXProduto(objOP)

                    objOPXP.IUD = "I"
                    objOPXP.CodigoProduto = drProduto("Produto")
                    objOPXP.Produto.Nome = drProduto("NomeProduto")
                    objOPXP.CodigoUnidadeComercializacao = drProduto("CodigoUnidadeComercializacao")
                    objOPXP.FatorConversao = drProduto("FatorConversao")
                    objOPXP.Quantidade = drProduto("Quantidade")
                    objOPXP.QuantidadeDeAjuste = drProduto("AjusteProducao")

                    objOP.ProdutosdaProducao.Add(objOPXP)

                Next

                If Not String.IsNullOrWhiteSpace(txtObservacoes.Text) Then objOP.Observacoes = Funcoes.EliminarCaracteresEspeciais(txtObservacoes.Text.ToUpper())

                For Each drItem As DataRow In CType(Session("objConsumo" & HID.Value), DataTable).Rows
                    objOPXC = New OrdemParaProducaoXConsumo(objOP)

                    objOPXC.IUD = "I"
                    objOPXC.CodigoProdutoProducao = drItem("CodigoProdutoProducao")
                    objOPXC.CodigoProduto = drItem("Produto")
                    objOPXC.Quantidade = drItem("Quantidade")
                    objOPXC.Sinal = "+"

                    'FURLAN - 18/10/2023
                    '############################################
                    '### CONFERINDO O ESTOQUE ANTES DE GRAVAR ###
                    '############################################

                    Dim ds As New DataSet

                    If objOPXC.Produto.ControlarEstoque Then
                        ds = objOPXC.buscarLoteDeFornecedor(objOP.CodigoEmpresa, objOP.EnderecoEmpresa, "'" & objOPXC.CodigoProduto & "'")
                        If ds.Tables(0).Rows.Count = 0 Then
                            MsgBox(Me.Page, "Não foi encontrado o lote do Produto. Ordem de Produção não será gravada. " & objOPXC.CodigoProduto & "-" & objOPXC.Produto.Nome, eTitulo.Info)
                            Exit Sub
                        End If
                    End If

                    For Each drLote As DataRow In CType(Session("objLoteFornecedor" & HID.Value), DataTable).Rows
                        If drLote("CodigoProdutoProducao") = objOPXC.CodigoProdutoProducao And objOPXC.CodigoProduto = drLote("Produto") AndAlso Not IsDBNull(drLote("Consumo")) AndAlso drLote("Consumo") > 0 Then

                            Dim temEstoqueLote As Boolean = False

                            For Each dr In ds.Tables(0).Rows
                                If dr("Produto") = drLote("Produto") AndAlso FuncoesStrings.NormalizeLote(dr("Lote"), objOP.CodigoEmpresa) = FuncoesStrings.NormalizeLote(drLote("Lote"), objOP.CodigoEmpresa) AndAlso CDate(dr("Validade")).ToString("yyyy-MM-dd") = CDate(drLote("Validade")).ToString("yyyy-MM-dd") AndAlso dr("Quantidade") >= drLote("Consumo") Then
                                    temEstoqueLote = True
                                End If
                            Next

                            If Not temEstoqueLote Then
                                MsgBox(Me.Page, "Falta saldo no lote do Produto. Ordem de Produção não será gravada. Produto: " & objOPXC.CodigoProduto & "-" & objOPXC.Produto.Nome & " - Lote: " & drLote("Lote") & " - Validade: " & CDate(drLote("Validade")).ToString("dd/MM/yyyy"), eTitulo.Info)
                                Exit Sub
                            End If

                            objOPXCXLote = New OrdemParaProducaoXConsumoXLote(objOPXC)

                            objOPXCXLote.IUD = "I"
                            objOPXCXLote.Lote = drLote("Lote")
                            objOPXCXLote.Quantidade = drLote("Consumo")
                            objOPXCXLote.Validade = drLote("Validade")

                            If objOPXC.ItensDeConsumoXLote.Where(Function(x) x.Lote = drLote("Lote")).Count() = 0 Then
                                objOPXC.ItensDeConsumoXLote.Add(objOPXCXLote)
                            Else
                                MsgBox(Me.Page, "O Produto: " & objOPXC.CodigoProduto & " - " & objOPXC.Produto.Nome & " está com o Lote: " & drLote("Lote") & " informado duplicado, favor corrigir para gravar a ordem de produção!", eTitulo.Info)
                                Exit Sub
                            End If

                        End If
                    Next

                    objOP.ItensDeConsumo.Add(objOPXC)
                Next

                If gridInsumo.Rows.Count > 0 Then

                    Dim j As Integer = 0

                    While j < gridInsumo.Rows.Count
                        If CDec(CType(gridInsumo.Rows(j).FindControl("txtQuantidadeInsumo"), TextBox).Text) > 0 Then

                            If CDec(gridInsumo.Rows(j).Cells(4).Text) < CDec(CType(gridInsumo.Rows(j).FindControl("txtQuantidadeInsumo"), TextBox).Text) Then
                                MsgBox(Me.Page, "Quantidade de Insumo informanda no Produto " & Trim(gridInsumo.Rows(j).Cells(0).Text) & " não pode ser maior que a quantidade disponível em estoque.", eTitulo.Info)
                                Exit Sub
                            End If

                            objOPXI = New OrdemParaProducaoXInsumo(objOP)

                            objOPXI.IUD = "I"

                            objOPXI.CodigoProduto = Trim(gridInsumo.Rows(j).Cells(0).Text)

                            For Each drItem As DataRow In CType(Session("objInsumo" & HID.Value), DataTable).Rows

                                If objOPXI.CodigoProduto = drItem("Produto") Then
                                    objOPXI.CodigoProdutoProducao = drItem("CodigoProdutoProducao")
                                    Exit For
                                End If
                            Next

                            objOPXI.Quantidade = CDec(CType(gridInsumo.Rows(j).FindControl("txtQuantidadeInsumo"), TextBox).Text)

                            objOP.ItensDeInsumo.Add(objOPXI)

                        End If

                        j += 1
                    End While

                End If

                If CType(Session("objEspecificacaoDoProduto" & HID.Value), DataTable).Rows.Count > 0 Then
                    For Each drPXE As DataRow In CType(Session("objEspecificacaoDoProduto" & HID.Value), DataTable).Rows
                        objOPXE = New OrdemParaProducaoXEspecificacao(objOP)

                        objOPXE.IUD = "I"
                        objOPXE.CodigoProdutoProducao = drPXE("CodigoProdutoProducao")
                        objOPXE.CodigoEspecificacao = drPXE("Codigo")
                        objOPXE.FaixaInicial = drPXE("FaixaInicial")
                        objOPXE.FaixaFinal = drPXE("FaixaFinal")

                        objOP.ItensDeEspecificacao.Add(objOPXE)
                    Next
                End If

                If CType(Session("objEPI" & HID.Value), DataTable).Rows.Count > 0 Then
                    For Each drEPI As DataRow In CType(Session("objEPI" & HID.Value), DataTable).Rows
                        objOPXEpi = New OrdemParaProducaoXEPI(objOP)

                        objOPXEpi.IUD = "I"
                        objOPXEpi.CodigoEPI = drEPI("Codigo")

                        objOP.ItensDeEPI.Add(objOPXEpi)
                    Next
                End If

                If CType(Session("objEmbalagem" & HID.Value), DataTable).Rows.Count > 0 Then

                    For Each drEmb As DataRow In CType(Session("objEmbalagem" & HID.Value), DataTable).Rows
                        objOPXEmb = New OrdemParaProducaoXEmbalagens(objOP)

                        objOPXEmb.IUD = "I"
                        objOPXEmb.CodigoEmbalagem = drEmb("Codigo")
                        objOPXEmb.CodigoUnidade = drEmb("CodigoUnidade")
                        objOPXEmb.Quantidade = drEmb("Quantidade")
                        objOPXEmb.Capacidade = drEmb("Capacidade")

                        objOP.ItensDeEmabalagem.Add(objOPXEmb)
                    Next
                End If

                If CType(Session("objProcedimento" & HID.Value), DataTable).Rows.Count > 0 Then
                    For Each drProcedimento As DataRow In CType(Session("objProcedimento" & HID.Value), DataTable).Rows
                        objOPXProcedimento = New OrdemParaProducaoXProcedimento(objOP)

                        objOPXProcedimento.IUD = "I"
                        objOPXProcedimento.CodigoProcedimento = drProcedimento("Codigo")

                        objOP.ItensDeProcedimento.Add(objOPXProcedimento)
                    Next
                End If

                If objOP.Salvar Then
                    SessaoSalvaOrdem()

                    MsgBox(Me.Page, "Ordem de Produção " & objOP.Codigo & " incluída com Sucesso.", eTitulo.Sucess)

                    ImprimirOrdem(objOP.CodigoEmpresa, objOP.EnderecoEmpresa, objOP.Codigo)

                    Limpar()
                Else
                    MsgBox(Me.Page, HttpContext.Current.Session("ssMessage").ToString)
                End If
            End If
        Else
            MsgBox(Me.Page, "Usuário sem permissão para gravar registro.", eTitulo.Info)
        End If

    End Sub

    Protected Sub lnkAtualizar_Click(sender As Object, e As EventArgs) Handles lnkAtualizar.Click

        If Funcoes.VerificaPermissao("OrdemDeProducao", "ALTERAR") Then
            SessaoRecuperaOrdem()

            If rbSim.Checked Then
                If Not Funcoes.VerificaAcesso(objOP.CodigoEmpresa, objOP.EnderecoEmpresa, txtDataEstoque.Text, "PRODUCAO") Then
                    MsgBox(Me.Page, "Movimento de Produção já Fechado, apenas Impressão ou Ajuste de Especificações.")
                    Exit Sub
                End If
            End If

            If ValidarSelecao() Or Session("ssNomeUsuario") = "DOUGLAS" Then

                AtualizarOrdemProducao(False)

            End If
        Else
            MsgBox(Me.Page, "Usuário sem permissão para alterar registro.", eTitulo.Info)
        End If

    End Sub

    Protected Sub AtualizarOrdemProducao(ByVal bConsumo As Boolean)

        objOP.IUD = "U"
        objOP.Validade = CDate(txtDataValidade.Text)
        objOP.MovimentoEstoque = CDate(txtDataProducao.Text)

        If Not String.IsNullOrWhiteSpace(txtObservacoes.Text) Then objOP.Observacoes = Funcoes.EliminarCaracteresEspeciais(txtObservacoes.Text.ToUpper())

        If rbSim.Checked Then
            objOP.Estoque = True
            objOP.MovimentoEstoque = CDate(txtDataEstoque.Text)
        Else
            objOP.Estoque = False
        End If

        For Each drProduto As DataRow In CType(Session("objProducao" & HID.Value), DataTable).Rows

            objOPXP = objOP.ProdutosdaProducao.FirstOrDefault(Function(x) x.CodigoProduto = drProduto("Produto"))

            If objOPXP Is Nothing Then

                objOPXP = New OrdemParaProducaoXProduto(objOP)

                objOPXP.IUD = "I"
                objOPXP.CodigoProduto = drProduto("Produto")
                objOPXP.Produto.Nome = drProduto("NomeProduto")
                objOPXP.CodigoUnidadeComercializacao = drProduto("CodigoUnidadeComercializacao")
                objOPXP.FatorConversao = drProduto("FatorConversao")
                objOPXP.Quantidade = drProduto("Quantidade")
                objOPXP.QuantidadeDeAjuste = drProduto("AjusteProducao")

                objOP.ProdutosdaProducao.Add(objOPXP)

            Else

                If String.IsNullOrWhiteSpace(objOPXP.IUD) Then
                    objOPXP.IUD = "U"
                End If

                objOPXP.CodigoProduto = drProduto("Produto")
                objOPXP.Produto.Nome = drProduto("NomeProduto")
                objOPXP.CodigoUnidadeComercializacao = drProduto("CodigoUnidadeComercializacao")
                objOPXP.FatorConversao = drProduto("FatorConversao")
                objOPXP.Quantidade = drProduto("Quantidade")
                objOPXP.QuantidadeDeAjuste = drProduto("AjusteProducao")

            End If

        Next

        If Not String.IsNullOrWhiteSpace(txtObservacoes.Text) Then objOP.Observacoes = Funcoes.EliminarCaracteresEspeciais(txtObservacoes.Text.ToUpper())

        For Each drItem As DataRow In CType(Session("objConsumo" & HID.Value), DataTable).Rows

            objOPXC = objOP.ItensDeConsumo.FirstOrDefault(Function(x) x.CodigoProdutoProducao = drItem("CodigoProdutoProducao") And x.CodigoProduto = drItem("Produto"))

            If objOPXC Is Nothing Then

                objOPXC = New OrdemParaProducaoXConsumo(objOP)

                objOPXC.IUD = "I"
                objOPXC.CodigoProdutoProducao = drItem("CodigoProdutoProducao")
                objOPXC.CodigoProduto = drItem("Produto")
                objOPXC.Quantidade = drItem("Quantidade")
                objOPXC.Sinal = drItem("Sinal")

                objOP.ItensDeConsumo.Add(objOPXC)

            Else

                If String.IsNullOrWhiteSpace(objOPXC.IUD) Then
                    objOPXC.IUD = "U"
                End If

                objOPXC.CodigoProdutoProducao = drItem("CodigoProdutoProducao")
                objOPXC.CodigoProduto = drItem("Produto")
                objOPXC.Quantidade = drItem("Quantidade")
                objOPXC.Sinal = drItem("Sinal")

            End If

            Dim ds As New DataSet

            If objOPXC.Produto.ControlarEstoque Then
                ds = objOPXC.buscarLoteDeFornecedor(objOP.CodigoEmpresa, objOP.EnderecoEmpresa, "'" & objOPXC.CodigoProduto & "'")
                If ds.Tables(0).Rows.Count = 0 Then
                    MsgBox(Me.Page, "Não foi encontrado o lote do Produto. Ordem de Produção não será gravada. " & objOPXC.CodigoProduto & "-" & objOPXC.Produto.Nome, eTitulo.Info)
                    Exit Sub
                End If
            End If

            For Each drLote As DataRow In CType(Session("objLoteFornecedor" & HID.Value), DataTable).Rows

                If drLote("CodigoProdutoProducao") = objOPXC.CodigoProdutoProducao And objOPXC.CodigoProduto = drLote("Produto") AndAlso Not IsDBNull(drLote("Consumo")) AndAlso drLote("Consumo") > 0 Then

                    Dim temEstoqueLote As Boolean = False

                    For Each dr In ds.Tables(0).Rows

                        If rbSim.Checked Then
                            If dr("Produto") = drLote("Produto") AndAlso FuncoesStrings.NormalizeLote(dr("Lote"), objOP.CodigoEmpresa) = FuncoesStrings.NormalizeLote(drLote("Lote"), objOP.CodigoEmpresa) AndAlso CDate(dr("Validade")).ToString("yyyy-MM-dd") = CDate(drLote("Validade")).ToString("yyyy-MM-dd") AndAlso dr("Quantidade") + IIf(IsDBNull(drLote("OldConsumo")), 0, drLote("OldConsumo")) >= drLote("Consumo") Then
                                temEstoqueLote = True
                            End If

                            Dim qtdEstoque As Decimal = drLote("Quantidade")
                        Else
                            If dr("Produto") = drLote("Produto") AndAlso FuncoesStrings.NormalizeLote(dr("Lote"), objOP.CodigoEmpresa) = FuncoesStrings.NormalizeLote(drLote("Lote"), objOP.CodigoEmpresa) AndAlso CDate(dr("Validade")).ToString("yyyy-MM-dd") = CDate(drLote("Validade")).ToString("yyyy-MM-dd") AndAlso dr("Quantidade") >= drLote("Consumo") Then
                                temEstoqueLote = True
                            End If
                        End If

                    Next

                    'Quando não tem lote, o consumo não pode ser maior que o OldConsumo se o estoque estiver fechado
                    If (Not temEstoqueLote AndAlso Not IsDBNull(drLote("OldConsumo")) AndAlso drLote("Consumo") > drLote("OldConsumo") And rbSim.Checked) Or Not temEstoqueLote And rbSim.Checked = False Then
                        MsgBox(Me.Page, "Falta saldo no lote do Produto. Ordem de Produção não será gravada. Produto: " & objOPXC.CodigoProduto & "-" & objOPXC.Produto.Nome & " - Lote: " & drLote("Lote") & " - Validade: " & CDate(drLote("Validade")).ToString("dd/MM/yyyy"), eTitulo.Info)
                        Exit Sub
                    End If

                    objOPXCXLote = objOPXC.ItensDeConsumoXLote.FirstOrDefault(Function(x) x.Lote = drLote("Lote"))

                    If objOPXCXLote Is Nothing Then

                        objOPXCXLote = New OrdemParaProducaoXConsumoXLote(objOPXC)

                        objOPXCXLote.IUD = "I"
                        objOPXCXLote.Lote = drLote("Lote")
                        objOPXCXLote.Quantidade = drLote("Consumo")
                        objOPXCXLote.Validade = drLote("Validade")

                        objOPXC.ItensDeConsumoXLote.Add(objOPXCXLote)

                    Else

                        If String.IsNullOrWhiteSpace(objOPXCXLote.IUD) Then
                            objOPXCXLote.IUD = "U"
                        End If

                        objOPXCXLote.Lote = drLote("Lote")
                        objOPXCXLote.Quantidade = drLote("Consumo")
                        objOPXCXLote.Validade = drLote("Validade")

                    End If

                ElseIf drLote("Consumo") = 0 AndAlso Not IsDBNull(drLote("OldConsumo")) AndAlso drLote("OldConsumo") > 0 Then

                    'Precisamos remover o estoque do consume do lote, porque foi ajustado os lotes

                    objOPXCXLote = objOPXC.ItensDeConsumoXLote.FirstOrDefault(Function(x) x.Lote = drLote("Lote"))

                    If Not objOPXCXLote Is Nothing Then
                        objOPXCXLote.IUD = "D"
                    End If

                End If
            Next

        Next

        If gridInsumo.Rows.Count > 0 OrElse objOP.ItensDeInsumo.Count > 0 Then

            If objOP.ItensDeInsumo.Count > 0 Then

                Dim j As Integer = 0

                While j < gridInsumo.Rows.Count
                    Dim temInsumo As Boolean = False

                    For Each itemI In objOP.ItensDeInsumo
                        If itemI.CodigoProdutoProducao = CodigoProdutoProducaoSelecionado And itemI.CodigoProduto = Trim(gridInsumo.Rows(j).Cells(0).Text) Then

                            If CDec(gridInsumo.Rows(j).Cells(4).Text) < CDec(CType(gridInsumo.Rows(j).FindControl("txtQuantidadeInsumo"), TextBox).Text) AndAlso objOP.Estoque Then

                                If objOP.IUD = "U" And objOP.Estoque Then
                                    'NÃO FAZ NADA PARA IZA PODER AJUSTAR A ORDEM DE PRODUÇÃO - FURLAN - 24/01/2023
                                Else
                                    MsgBox(Me.Page, "Quantidade de Insumo informanda no Produto " & Trim(gridInsumo.Rows(j).Cells(0).Text) & " não pode ser maior que a quantidade disponível em estoque.", eTitulo.Info)

                                    Exit Sub
                                End If
                            End If

                            temInsumo = True
                            itemI.Quantidade = CDec(CType(gridInsumo.Rows(j).FindControl("txtQuantidadeInsumo"), TextBox).Text)
                        End If
                    Next

                    If Not temInsumo Then
                        objOPXI = New OrdemParaProducaoXInsumo(objOP)

                        objOPXI.IUD = "I"
                        objOPXI.CodigoProduto = Trim(gridInsumo.Rows(j).Cells(0).Text)
                        objOPXI.Quantidade = CDec(CType(gridInsumo.Rows(j).FindControl("txtQuantidadeInsumo"), TextBox).Text)
                        objOPXI.CodigoProdutoProducao = objOPXP.CodigoProduto

                        objOP.ItensDeInsumo.Add(objOPXI)
                    End If

                    j += 1
                End While

            Else

                Dim j As Integer = 0

                While j < gridInsumo.Rows.Count
                    If CDec(CType(gridInsumo.Rows(j).FindControl("txtQuantidadeInsumo"), TextBox).Text) > 0 Then


                        If CDec(gridInsumo.Rows(j).Cells(4).Text) < CDec(CType(gridInsumo.Rows(j).FindControl("txtQuantidadeInsumo"), TextBox).Text) Then
                            MsgBox(Me.Page, "Quantidade de Insumo informanda no Produto " & Trim(gridInsumo.Rows(j).Cells(0).Text) & " não pode ser maior que a quantidade disponível em estoque.", eTitulo.Info)

                            Exit Sub
                        End If

                        objOPXI = New OrdemParaProducaoXInsumo(objOP)

                        objOPXI.IUD = "I"
                        objOPXI.CodigoProduto = Trim(gridInsumo.Rows(j).Cells(0).Text)
                        objOPXI.Quantidade = CDec(CType(gridInsumo.Rows(j).FindControl("txtQuantidadeInsumo"), TextBox).Text)
                        objOPXI.CodigoProdutoProducao = objOPXP.CodigoProduto

                        objOP.ItensDeInsumo.Add(objOPXI)
                    End If

                    j += 1
                End While
            End If

            For Each itemInsumo In objOP.ItensDeInsumo
                If itemInsumo.Quantidade = 0 Then itemInsumo.IUD = "D"
            Next

        End If


        If CType(Session("objEPI" & HID.Value), DataTable).Rows.Count > 0 Then
            For Each itemEPI In objOP.ItensDeEPI
                If itemEPI.IUD = "U" Then
                    itemEPI.IUD = "D"

                    For Each drEPI As DataRow In CType(Session("objEPI" & HID.Value), DataTable).Rows
                        If drEPI("Codigo") = itemEPI.CodigoEPI Then
                            itemEPI.IUD = "U"
                        End If
                    Next
                End If
            Next
        End If

        If CType(Session("objEmbalagem" & HID.Value), DataTable).Rows.Count > 0 Then
            For Each itemEMB In objOP.ItensDeEmabalagem
                If itemEMB.IUD = "U" Then
                    itemEMB.IUD = "D"

                    For Each drEmb As DataRow In CType(Session("objEmbalagem" & HID.Value), DataTable).Rows
                        If drEmb("Codigo") = itemEMB.CodigoEmbalagem AndAlso drEmb("CodigoUnidade") = itemEMB.CodigoUnidade Then
                            itemEMB.IUD = "U"
                        End If
                    Next
                End If
            Next
        End If

        If CType(Session("objProcedimento" & HID.Value), DataTable).Rows.Count > 0 Then
            For Each itemProcedimento In objOP.ItensDeProcedimento
                If itemProcedimento.IUD = "U" Then
                    itemProcedimento.IUD = "D"

                    For Each drProcedimento As DataRow In CType(Session("objProcedimento" & HID.Value), DataTable).Rows
                        If drProcedimento("Codigo") = itemProcedimento.CodigoProcedimento Then
                            itemProcedimento.IUD = "U"
                        End If
                    Next
                End If
            Next
        End If

        If objOP.Salvar Then
            SessaoSalvaOrdem()

            MsgBox(Me.Page, "Ordem de Produção " & objOP.Codigo & " Alterada com Sucesso.", eTitulo.Sucess)

            If bConsumo = False Then
                ImprimirOrdem(objOP.CodigoEmpresa, objOP.EnderecoEmpresa, objOP.Codigo)
                Limpar()
            End If

        Else
            MsgBox(Me.Page, HttpContext.Current.Session("ssMessage").ToString)
        End If

    End Sub

    Protected Sub lnkExcluir_Click(sender As Object, e As EventArgs) Handles lnkExcluir.Click
        If Funcoes.VerificaPermissao("OrdemDeProducao", "EXCLUIR") Then
            SessaoRecuperaOrdem()

            If objOP.Estoque Then
                If Not Funcoes.VerificaAcesso(objOP.CodigoEmpresa, objOP.EnderecoEmpresa, txtDataEstoque.Text, "PRODUCAO") Then
                    MsgBox(Me.Page, "Movimento de Produção já Fechado, apenas Impressão ou Ajuste de Especificações.")
                    Exit Sub
                End If
            End If

            objOP.IUD = "D"
            objOP.CodigoSituacao = CInt(eSituacao.Excluido)
            If Not String.IsNullOrWhiteSpace(txtObservacoes.Text) Then objOP.Observacoes = Funcoes.EliminarCaracteresEspeciais(txtObservacoes.Text.ToUpper())

            If objOP.Salvar Then
                MsgBox(Me.Page, "Ordem de Produção " & objOP.Codigo & " Cancelada com Sucesso.", eTitulo.Sucess)
                Limpar()
            Else
                MsgBox(Me.Page, HttpContext.Current.Session("ssMessage").ToString)
            End If
        Else
            MsgBox(Me.Page, "Usuário sem permissão para excluir registro.", eTitulo.Info)
        End If
    End Sub

    Protected Sub lnkImprimirOrdem_Click(sender As Object, e As EventArgs) Handles lnkImprimirOrdem.Click
        If Funcoes.VerificaPermissao("OrdemDeProducao", "RELATORIO") Then
            SessaoRecuperaOrdem()

            If objOP.Codigo > 0 Then
                ImprimirOrdem(objOP.CodigoEmpresa, objOP.EnderecoEmpresa, objOP.Codigo)
            Else
                MsgBox(Me.Page, "Ordem de Produção " & objOP.Codigo & " não foi informada.", eTitulo.Sucess)
            End If
        Else
            MsgBox(Me.Page, "Usuário sem permissão para tirar Relatório", eTitulo.Info)
        End If
    End Sub

    Protected Sub lnkImprimirLaudo_Click(sender As Object, e As EventArgs) Handles lnkImprimirLaudo.Click
        If Funcoes.VerificaPermissao("OrdemDeProducao", "RELATORIO") Then
            ImprimirLaudo()
        Else
            MsgBox(Me.Page, "Usuário sem permissão para tirar Relatório", eTitulo.Info)
        End If
    End Sub

    Protected Sub lnkImprimirLaudoManual_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lnkImprimirLaudoManual.Click
        Try
            If Funcoes.VerificaPermissao("OrdemDeProducao", "RELATORIO") Then
                If ddlEmpresa.SelectedIndex = 0 Then
                    MsgBox(Me.Page, "Empresa não selecionada.")
                Else
                    ucLaudoManual.Limpar()
                    Popup.LaudoManual(Me.Page, "objLaudoManual" & HID.Value)
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para tirar Relatório", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkLimpar_Click(sender As Object, e As EventArgs) Handles lnkLimpar.Click
        Limpar()
    End Sub
    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "OrdemDeProducao")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub

    Protected Sub rbNao_CheckedChanged(sender As Object, e As EventArgs) Handles rbNao.CheckedChanged
        txtDataEstoque.Enabled = False
        dataEstoque.Visible = False
    End Sub

    Protected Sub rbSim_CheckedChanged(sender As Object, e As EventArgs) Handles rbSim.CheckedChanged
        dataEstoque.Visible = True
        txtDataEstoque.Enabled = True
        txtDataEstoque.Text = Now.ToString("dd/MM/yyyy")
    End Sub

    Protected Sub imgCRelatorio_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles imgCRelatorio.Click

        If Funcoes.VerificaPermissao("OrdemDeProducao", "RELATORIO") Then
            If ddlUnidadeNegocio.SelectedIndex = 0 Then
                MsgBox(Me.Page, "Unidade de Negócio não foi selecionada.", eTitulo.Info)
            ElseIf ddlEmpresa.SelectedIndex = 0 Then
                MsgBox(Me.Page, "Empresa não foi selecionada.", eTitulo.Info)
            Else
                Try
                    Dim objBanco As New AcessaBanco()

                    Dim strSQL As String = "select o.Ordem_Id AS Ordem, OPXP.Produto_Id AS Produto, p.Nome, OPXP.Lote, convert(varchar,o.Movimento,103) AS Movimento, convert(varchar,o.Validade,103) AS Validade, ISNULL(convert(varchar,PRO.Movimento_Id,103), '') AS Estoque, OPXP.Quantidade, OPXP.QuantidadeDeAjuste, " & vbCrLf &
                                            "		case" & vbCrLf &
                                            "			when o.Estoque = 0" & vbCrLf &
                                            "				then 'NÃO'" & vbCrLf &
                                            "				else 'SIM'" & vbCrLf &
                                            "			end AS Estoque," & vbCrLf &
                                            "		o.UsuarioInclusao, convert(varchar,o.UsuarioInclusaoData,103) AS UsuarioInclusaoData, " & vbCrLf &
                                            "		isnull(o.UsuarioAlteracao,'') AS UsuarioAlteracao, " & vbCrLf &
                                            "		case" & vbCrLf &
                                            "			when isnull(o.UsuarioAlteracaoData,'') = ''" & vbCrLf &
                                            "				then ''" & vbCrLf &
                                            "				else convert(varchar,o.UsuarioAlteracaoData,103)" & vbCrLf &
                                            "			end AS UsuarioAlteracaoData," & vbCrLf &
                                            "		isnull(o.UsuarioCancelamento,'') AS UsuarioCancelamento, " & vbCrLf &
                                            "		case" & vbCrLf &
                                            "			when isnull(o.UsuarioCancelamentoData,'') = ''" & vbCrLf &
                                            "				then ''" & vbCrLf &
                                            "				else convert(varchar,o.UsuarioCancelamentoData,103) " & vbCrLf &
                                            "			end AS UsuarioCancelamentoData" & vbCrLf &
                                            "FROM OrdemDeProducao o" & vbCrLf &
                                            "INNER JOIN OrdemDeProducaoXProduto OPXP " & vbCrLf &
                                            "   ON o.Empresa_Id		        = OPXP.Empresa_Id " & vbCrLf &
                                            "   AND o.EndEmpresa_Id	        = OPXP.EndEmpresa_Id " & vbCrLf &
                                            "   AND o.Ordem_Id			    = OPXP.Ordem_Id " & vbCrLf &
                                            "INNER JOIN Produtos p" & vbCrLf &
                                            "   ON p.Produto_id             = OPXP.Produto_Id" & vbCrLf &
                                            "LEFT JOIN Producao PRO" & vbCrLf &
                                            "		    ON o.Empresa_Id     = PRO.Empresa_Id " & vbCrLf &
                                            "   AND o.EndEmpresa_Id         = PRO.EndEmpresa_Id " & vbCrLf &
                                            "   AND o.Ordem_Id              = PRO.OrdemDeProducao " & vbCrLf &
                                            "   AND OPXP.Produto_Id         = PRO.Produto_Id " & vbCrLf &
                                            "   AND PRO.FisicoFiscal_Id     = 1 " & vbCrLf &
                                            "WHERE o.Empresa_Id             = '" & ddlEmpresa.SelectedValue.Split("-")(0) & "'" & vbCrLf &
                                            "   AND o.EndEmpresa_Id         = " & ddlEmpresa.SelectedValue.Split("-")(1) & vbCrLf

                    If rbCancelada.Checked Then
                        strSQL &= "AND o.Situacao = 3" & vbCrLf
                    Else
                        strSQL &= "AND o.Situacao = 1" & vbCrLf
                    End If

                    If rbAberta.Checked Then strSQL &= "AND o.Estoque = 0" & vbCrLf

                    If rbEncerrada.Checked Then strSQL &= "AND o.Estoque = 1" & vbCrLf

                    If rbDtMov.Checked Then strSQL &= "AND o.Movimento between '" & CDate(txtDataInicial.Text).ToString("yyyy-MM-dd") & "' and '" & CDate(txtDataFinal.Text).ToString("yyyy-MM-dd") & "'" & vbCrLf

                    If rbDtVen.Checked Then strSQL &= "AND o.Validade between '" & CDate(txtDataInicial.Text).ToString("yyyy-MM-dd") & "' and '" & CDate(txtDataFinal.Text).ToString("yyyy-MM-dd") & "'" & vbCrLf

                    If rbDtEstoque.Checked Then strSQL &= "AND PRO.Movimento_Id between '" & CDate(txtDataInicial.Text).ToString("yyyy-MM-dd") & "' and '" & CDate(txtDataFinal.Text).ToString("yyyy-MM-dd") & "'" & vbCrLf

                    If rbOrdem.Checked Then strSQL &= "ORDER BY o.Ordem_Id" & vbCrLf

                    If rbOPrd.Checked Then strSQL &= "ORDER BY OPXP.Produto_Id" & vbCrLf

                    If rbOPrdNome.Checked Then strSQL &= "ORDER BY p.Nome" & vbCrLf

                    Dim ds As DataSet = Banco.ConsultaDataSet(strSQL, "RelatorioDeOrdemDeProducao")

                    If ds.Tables(0).Rows.Count = 0 Then
                        MsgBox(Me.Page, "Sem registros para esse período.", eTitulo.Info)
                        Exit Sub
                    End If

                    Dim objEmpresa As New Cliente()

                    If ddlEmpresa.SelectedIndex > 0 Then
                        objEmpresa = New Cliente(ddlEmpresa.SelectedValue.ToString.Split("-")(0), ddlEmpresa.SelectedValue.ToString.Split("-")(1))
                    Else
                        objEmpresa = New Cliente(Session("ssEmpresa"), Session("ssEndEmpresa"))
                    End If

                    Dim fileName As String = Server.MapPath("~/PlanilhasExcel/" & Funcoes.GeraNomeArquivo & ".xlsx")

                    If File.Exists(fileName) Then
                        File.Delete(fileName)
                    End If

                    Using arquivo As New FileStream(fileName, FileMode.CreateNew)
                        Using package As New ExcelPackage(arquivo)
                            'criando planilha títulos
                            Dim worksheet As ExcelWorksheet = package.Workbook.Worksheets.Add("Resumo")

                            'criando linha com o cabeçalho da planilha
                            Dim rowIndex As Integer = 8
                            Dim columnIndex As Integer = 1

                            worksheet.Cells("A1:F1").Merge = True
                            worksheet.Cells("A2:F2").Merge = True
                            worksheet.Cells("A3:F3").Merge = True
                            worksheet.Cells("A4:F4").Merge = True
                            worksheet.Cells("A5:F5").Merge = True
                            worksheet.Cells("A6:F6").Merge = True
                            worksheet.Cells("A7:F7").Merge = True
                            worksheet.Cells("A1:F7").Style.Font.Bold = True
                            worksheet.Cells("A1:F7").Style.Font.Size = 14

                            worksheet.Cells("A1").Value = objEmpresa.Nome
                            worksheet.Cells("A2").Value = "           " & objEmpresa.Cidade & " - " & objEmpresa.CodigoEstado
                            worksheet.Cells("A3").Value = "           CNPJ - " & objEmpresa.CodigoFormatado

                            'criando linha que informa o título do relatório
                            worksheet.Cells("A5").Style.Font.Bold = True
                            worksheet.Cells("A5").Style.Font.Color.SetColor(Color.Red)
                            worksheet.Cells("A5").Value = "           RELATÓRIO DAS ORDENS DE PRODUÇÃO"

                            worksheet.Cells("A7").Value = "PERÍODO DE " & txtDataInicial.Text & " À " & txtDataFinal.Text

                            'criando linha com o cabeçalho da planilha
                            For Each col As DataColumn In ds.Tables(0).Columns
                                If col.ColumnName = "UsuarioInclusao" Then
                                    worksheet.Cells(rowIndex, columnIndex).Value = "USUÁRIO"
                                    worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                                    worksheet.Cells(rowIndex, columnIndex).Style.Font.Size = 12
                                ElseIf col.ColumnName = "UsuarioInclusaoData" Then
                                    worksheet.Cells(rowIndex, columnIndex).Value = "DATA INCLUSÃO"
                                    worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                                    worksheet.Cells(rowIndex, columnIndex).Style.Font.Size = 12
                                ElseIf col.ColumnName = "UsuarioAlteracao" Then
                                    worksheet.Cells(rowIndex, columnIndex).Value = "USUÁRIO"
                                    worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                                    worksheet.Cells(rowIndex, columnIndex).Style.Font.Size = 12
                                ElseIf col.ColumnName = "UsuarioAlteracaoData" Then
                                    worksheet.Cells(rowIndex, columnIndex).Value = "DATA ALTERAÇÃO"
                                    worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                                    worksheet.Cells(rowIndex, columnIndex).Style.Font.Size = 12
                                ElseIf col.ColumnName = "UsuarioCancelamento" Then
                                    worksheet.Cells(rowIndex, columnIndex).Value = "USUÁRIO"
                                    worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                                    worksheet.Cells(rowIndex, columnIndex).Style.Font.Size = 12
                                ElseIf col.ColumnName = "UsuarioCancelamentoData" Then
                                    worksheet.Cells(rowIndex, columnIndex).Value = "DATA CANCELAMENTO"
                                    worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                                    worksheet.Cells(rowIndex, columnIndex).Style.Font.Size = 12
                                ElseIf col.ColumnName = "QuantidadeDeAjuste" Then
                                    worksheet.Cells(rowIndex, columnIndex).Value = "AJUSTE"
                                    worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Right
                                    worksheet.Cells(rowIndex, columnIndex).Style.Font.Size = 12
                                Else
                                    worksheet.Cells(rowIndex, columnIndex).Value = col.ColumnName.ToUpper
                                    worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                                    worksheet.Cells(rowIndex, columnIndex).Style.Font.Size = 12
                                End If

                                columnIndex += 1
                            Next

                            Using range = worksheet.Cells(rowIndex, 1, rowIndex, ds.Tables(0).Columns.Count)
                                range.Style.Font.Bold = True
                                range.Style.Fill.PatternType = ExcelFillStyle.Solid
                                range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
                                range.Style.Font.Color.SetColor(Color.White)
                            End Using

                            rowIndex += 1

                            ' criando conteúdo da planilha com os dados do dataset
                            For Each row As DataRow In ds.Tables(0).Rows
                                columnIndex = 1
                                For Each col As DataColumn In ds.Tables(0).Columns
                                    worksheet.Cells(rowIndex, columnIndex).Value = row(col.ColumnName)
                                    columnIndex += 1
                                Next

                                'aplicando formatação nas células do conteúdo
                                If rowIndex Mod 2 = 0 Then
                                    Using range = worksheet.Cells(rowIndex, 1, rowIndex, ds.Tables(0).Columns.Count)
                                        range.Style.Fill.PatternType = ExcelFillStyle.Solid
                                        range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(153, 204, 255))
                                    End Using
                                End If
                                rowIndex += 1
                            Next

                            'setando autofit nas células da planilha
                            worksheet.Cells.AutoFitColumns(0)

                            'congelando quinta linha (cabeçalho)
                            worksheet.View.FreezePanes(8, 1)

                            'salvando planilha do excel
                            package.Save()
                        End Using
                    End Using

                    Funcoes.AbrirExcel(Me.Page, "PlanilhasExcel/" & Path.GetFileName(fileName))
                Catch ex As Exception
                    MsgBox(Me.Page, ex.Message, eTitulo.Erro)
                End Try
            End If
        Else
            MsgBox(Me.Page, "Usuário sem permissão para tirar Relatório", eTitulo.Info)
        End If
    End Sub

    Private Function enviarEstoqueMinimo() As Boolean

        Dim estoqueOK As Boolean = True

        Dim lstMail As New List(Of String)

        Dim Filial As String = String.Empty
        Dim bodyHTML = String.Empty

        bodyHTML &= "<h1 style='clear: left; color: red; display: block; line-height: 105%; margin: 2px; padding: 7px 5px; position: relative; font-size: 150%; font-weight: bold;'>Alerta para o Estoque mínimo</h1>" & vbCrLf

        bodyHTML &= "<br/>Sr(s), <br/><br/>" & vbCrLf

        If CType(Session("objEstoqueMinimo" & HID.Value), DataTable).Rows.Count = 1 Then
            bodyHTML &= "Emitindo Ordem para Produção, o Produto abaixo merece atenção: <br/>" & vbCrLf
        Else
            bodyHTML &= "Emitindo Ordem para Produção, os Produtos abaixo merecem atenção: <br/>" & vbCrLf
        End If

        'bodyHTML &= "<ul>"
        If String.IsNullOrEmpty(Filial) Then Filial = ddlEmpresa.SelectedValue.Split("-")(0)

        bodyHTML &= "<table border='1' bordercolor='#000000' cellpadding='0' cellspacing='0'>" & vbCrLf

        bodyHTML &= "<tr>" & vbCrLf
        bodyHTML &= "<td>&nbsp;PRODUTO&nbsp;</td>" & vbCrLf
        bodyHTML &= "<td style='text-align:right;'>&nbsp;ESTOQUE MÍNIMO&nbsp;</td>" & vbCrLf
        bodyHTML &= "<td style='text-align:right;'>&nbsp;FATURANDO&nbsp;</td>" & vbCrLf
        bodyHTML &= "<td style='text-align:right;'>&nbsp;SALDO&nbsp;</td>" & vbCrLf
        bodyHTML &= "</tr>" & vbCrLf

        For Each drItemE As DataRow In CType(Session("objEstoqueMinimo" & HID.Value), DataTable).Rows
            bodyHTML &= "<tr>" & vbCrLf
            bodyHTML &= "<td>&nbsp;" & drItemE("Produto") & "-" & drItemE("Nome") & "&nbsp;</td>" & vbCrLf
            bodyHTML &= "<td style='text-align:right;'>&nbsp;" & drItemE("EstoqueMinimo") & "&nbsp;</td>" & vbCrLf
            bodyHTML &= "<td style='text-align:right;'>&nbsp;" & drItemE("Faturando") & "&nbsp;</td>" & vbCrLf
            bodyHTML &= "<td style='text-align:right;'>&nbsp;" & drItemE("Saldo") & "&nbsp;</td>" & vbCrLf
            bodyHTML &= "</tr>" & vbCrLf
        Next

        bodyHTML &= "</table>" & vbCrLf

        bodyHTML &= "<br/><font size='2'>Este é um e-mail gerado automaticamente. Por favor, não responder.</font>" & vbCrLf

        bodyHTML &= "<br/><br/><font size='2'>© Copyright " & Now.ToString("yyyy") & " NGS Soluções Ltda - Todos os direitos reservados</font>" & vbCrLf

        Dim objUsuario As New [Lib].Negocio.Usuario(HttpContext.Current.Session("ssNomeUsuario"))
        If objUsuario IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(objUsuario.Email) AndAlso objUsuario.Email.Trim() <> "&nbsp;" Then
            lstMail.Add(objUsuario.Email)
        End If

        Dim Sql As String = "SELECT u.Email " & vbCrLf &
                              "  FROM ConfiguracaoXUsuario cxu " & vbCrLf &
                              " INNER JOIN Usuarios u " & vbCrLf &
                              "    ON (u.Usuario_Id = cxu.Usuario_Id) " & vbCrLf &
                              " WHERE cxu.Etapa_Id = " & eEtapa.EstoqueMinimo

        Dim dsMail As DataSet = Banco.ConsultaDataSet(Sql, "ConfiguracaoXUsuario")
        If dsMail IsNot Nothing AndAlso dsMail.Tables IsNot Nothing AndAlso dsMail.Tables.Count > 0 AndAlso dsMail.Tables(0).Rows.Count > 0 Then
            For Each row As DataRow In dsMail.Tables(0).Rows
                If Not String.IsNullOrWhiteSpace(row("Email")) Then
                    lstMail.Add(row("Email"))
                End If
            Next
        End If

        'bodyHTML &= "</ul>"

        Dim EmprMail = New [Lib].Negocio.Cliente(Filial, 0)
        Dim errorMsgMail As String = String.Empty
        Dim subject As String = String.Format("Empresa " & EmprMail.CodigoFormatado & " (" & EmprMail.Cidade & "/" & EmprMail.CodigoEstado & ") - Atenção para o Estoque mínimo - {0:dd/MM/yyyy HH:mm}", DateTime.Now)
        Dim smtp = Funcoes.GetSmtpSettings()
        Dim fromMail = Funcoes.GetFromMail()

        If lstMail IsNot Nothing AndAlso lstMail.Count > 0 Then
            Funcoes.SendMail(fromMail, "NGS SOLUÇÕES", lstMail, subject, bodyHTML, smtp, errorMsgMail)
        End If

        Return True

    End Function

    Protected Sub imgExcluir_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)

        Try
            If Funcoes.VerificaPermissao("GerenciarOrdemDeProducao", "EXCLUIR") Then

                Dim imgSelecionaExcluir As ImageButton = CType(sender, ImageButton)
                Dim row As GridViewRow = CType(imgSelecionaExcluir.NamingContainer, GridViewRow)

                Dim rowIndex As Integer = row.RowIndex
                Dim itemDeConsumo As New OrdemParaProducaoXConsumo

                SessaoRecuperaOrdem()

                itemDeConsumo = objOP.ItensDeConsumo.Where(Function(x) x.CodigoProduto = CType(Session("objConsumo" & HID.Value), DataTable).Rows(row.RowIndex).Item("Produto")).FirstOrDefault()
                itemDeConsumo.IUD = "DEL_CONS"

                If itemDeConsumo.ItensDeConsumoXLote.Any() Then
                    itemDeConsumo.ItensDeConsumoXLote.ForEach(Sub(item) item.IUD = "DEL_CONS")
                End If

                CType(Session("objConsumo" & HID.Value), DataTable).Rows(row.RowIndex).Delete()

                gridConsumo.DataSource = dtgridConsumo()
                gridConsumo.DataBind()

            Else
                MsgBox(Me.Page, "Usuário sem permissão para excluir registro.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try

    End Sub

    Protected Sub ddlGrupoProdutoConsumo_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddlGrupoProdutoConsumo.SelectedIndexChanged
        CarregarProdutoConsumo()
    End Sub

    Protected Sub gridProducao_SelectedIndexChanged(sender As Object, e As EventArgs) Handles gridProducao.SelectedIndexChanged
        Dim row As GridViewRow = gridProducao.SelectedRow
        If row IsNot Nothing Then
            CodigoProdutoProducaoSelecionado = row.Cells(1).Text ' Ajuste o índice da célula conforme necessário
        End If

        If Not sender Is Nothing Then
            gridConsumo.DataSource = dtgridConsumo()
            gridConsumo.DataBind()

            gridInsumo.DataSource = dtgridInsumo()
            gridInsumo.DataBind()

            gridEspecificacaoDoProduto.DataSource = dtgridEspecificacaoDoProduto()
            gridEspecificacaoDoProduto.DataBind()

        End If

        gridLoteDeFornecedor.DataSource = Nothing
        gridLoteDeFornecedor.DataBind()
        divConsumo.Style.Value = "width: 99%;"
        divConsumoLote.Visible = False

        DisableSelecionarLote()

    End Sub

    Private Sub DisableSelecionarLote()

        Dim i As Integer = 0
        While i < gridConsumo.Rows.Count

            If gridConsumo.Rows(i).Cells(3).Text = "INFORMADO" Or gridConsumo.Rows(i).Cells(3).Text = "INFORMAR" Then
                CType(gridConsumo.Rows(i).FindControl("imgSelecionaLote"), ImageButton).ImageUrl = "~/images/select.jpg"
                CType(gridConsumo.Rows(i).FindControl("imgSelecionaLote"), ImageButton).ToolTip = "Informar Lote do Fornecedor"
                CType(gridConsumo.Rows(i).FindControl("imgSelecionaLote"), ImageButton).Style.Value = "cursor: pointer; border: 0 none;"
            Else
                CType(gridConsumo.Rows(i).FindControl("imgSelecionaLote"), ImageButton).ImageUrl = "~/images/important.png"
                CType(gridConsumo.Rows(i).FindControl("imgSelecionaLote"), ImageButton).ToolTip = "Não informa Lote"
                CType(gridConsumo.Rows(i).FindControl("imgSelecionaLote"), ImageButton).Enabled = False
            End If

            For Each drItemC As DataRow In CType(Session("objConsumo" & HID.Value), DataTable).Rows
                If drItemC("CodigoProdutoProducao") = CodigoProdutoProducaoSelecionado And drItemC("Produto") = gridConsumo.Rows(i).Cells(1).Text Then
                    CType(gridConsumo.Rows(i).FindControl("DdlSinal"), DropDownList).SelectedValue = drItemC("Sinal")
                End If
            Next

            i += 1
        End While



    End Sub

    Private Sub CarregarProdutoConsumo()
        Try
            ddlProdutosConsumo.Enabled = True
            ddl.Carregar(ddlProdutosConsumo, CarregarDDL.Tabela.Produto, " ((ControlarEstoque = 'S') or Produto_Id in ( select pc.Produto_id from ProdutoDeConsumo pc inner join Produtos p on p.Produto_Id = pc.Produto_id where  p.Grupo = '" & ddlGrupoProdutoConsumo.SelectedValue & "')) and Grupo ='" & ddlGrupoProdutoConsumo.SelectedValue & "'")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkBuscaProdutoConsumo_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lnkBuscaProdutoConsumo.Click
        Try
            ucConsultaProduto.Limpar()
            Dim txtNome As TextBox = CType(ucConsultaProduto.FindControlRecursive("txtNome"), TextBox)
            Popup.ConsultaDeProduto(Me.Page, "objProdutoxCON" & HID.Value, txtNome.ClientID, True)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub ProdutosConsumo()
        Try

            SessaoRecuperaOrdem()

            If Session("objOrdemDeProducao" & HID.Value) Is Nothing Then
                MsgBox(Me.Page, "Selecione uma ordem de produção.", eTitulo.Info)
            ElseIf String.IsNullOrWhiteSpace(txtQuantidadeAlt.Text) OrElse CDec(txtQuantidadeAlt.Text) = 0 Then
                ddlProdutosConsumo.SelectedIndex = 0
                MsgBox(Me.Page, "Quantidade para Consumo não foi informado.", eTitulo.Info)
                'ElseIf CDec(txtQuantidadeAlt.Text) > 100 Then
                '    ddlProdutosConsumo.SelectedIndex = 0
                '    MsgBox(Me.Page, "Quantidade para Consumo não pode ser maior que 100%.", eTitulo.Info)
            ElseIf ddlProdutosConsumo.SelectedIndex = 0 Then
                ddlProdutosConsumo.SelectedIndex = 0
                MsgBox(Me.Page, "Produto para Consumo não foi informado.", eTitulo.Info)
            ElseIf String.IsNullOrWhiteSpace(txtQuantidadeAlt.Text) Then
                MsgBox(Me.Page, "Quantidade não foi informada", eTitulo.Info)
            ElseIf CDec(txtQuantidadeAlt.Text) = 0 Then
                MsgBox(Me.Page, "Quantidade não foi informada", eTitulo.Info)
            ElseIf ddlProdutosConsumo.SelectedValue Is Nothing OrElse ddlProdutosConsumo.SelectedValue = "" Then
                MsgBox(Me.Page, "O produto deve ser informado", eTitulo.Info)
            ElseIf txtLoteAlt.Text = "" Then
                MsgBox(Me.Page, "O lote deve ser informado", eTitulo.Info)
            Else

                Dim dtConsumo As DataTable = dtgridConsumoAlt()

                If dtConsumo.Rows.Count > 0 Then

                    If Not IsDBNull(dtConsumo.Compute("SUM(Quantidade)", "Produto = " & ddlProdutosConsumo.SelectedValue)) Then
                        ddlProdutosConsumo.SelectedIndex = 0
                        MsgBox(Me.Page, "Produto já existe na lista para Consumo.", eTitulo.Info)
                        Exit Sub
                        'ElseIf (CType(Session("objConsumoAlt" & HID.Value), DataTable).Compute("SUM(Quantidade)", "") + CDec(txtQuantidadeAlt.Text)) > 100 Then
                        '    ddlProdutosConsumo.SelectedIndex = 0
                        '    MsgBox(Me.Page, "A soma do Quantidade para Consumo dos Produtos não pode ser maior que 100%.", eTitulo.Info)
                        '    Exit Sub
                    End If
                End If

                objPrd = New Produto(ddlProdutosConsumo.SelectedValue)

                Dim drItem As DataRow = CType(Session("objConsumoAlt" & HID.Value), DataTable).NewRow()

                drItem("CodigoProdutoProducao") = CodigoProdutoProducaoSelecionado
                drItem("Produto") = objPrd.Codigo
                drItem("NomeProduto") = objPrd.Nome
                drItem("Quantidade") = txtQuantidadeAlt.Text
                drItem("Lote") = txtLoteAlt.Text
                drItem("Validade") = txtValidadeAlt.Text
                drItem("Sinal") = IIf(rbMais.Checked, "+", "-")

                CType(Session("objConsumoAlt" & HID.Value), DataTable).Rows.Add(drItem)

                Dim dtConsumoAlt As DataTable = dtgridConsumoAlt()

                gridConsumoAlt.DataSource = dtConsumoAlt
                gridConsumoAlt.DataBind()

                lblQuantidadeAlt.Text = "TOTAL = " & dtConsumoAlt.Compute("SUM(Quantidade)", "")

                ' Verifica se o DataTable não é Nothing
                If dtConsumoAlt IsNot Nothing Then

                End If
                LimparCamposInsumos()
                EstadosCamposConsumo(True)

            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub gridConsumoAlt_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles gridConsumoAlt.SelectedIndexChanged

        Dim dtConsumo As DataTable = dtgridConsumoAlt()
        txtQuantidadeAlt.Text = dtConsumo.Rows(gridConsumoAlt.SelectedIndex).Item("Quantidade")
        txtLoteAlt.Text = dtConsumo.Rows(gridConsumoAlt.SelectedIndex).Item("Lote")
        txtDataValidade.Text = dtConsumo.Rows(gridConsumoAlt.SelectedIndex).Item("Validade")

        objPrd = New Produto(dtConsumo.Rows(gridConsumoAlt.SelectedIndex).Item("Produto"))
        ddlGrupoProdutoConsumo.SelectedValue = objPrd.CodigoGrupo
        ddl.Carregar(ddlProdutosConsumo, CarregarDDL.Tabela.Produto, " Situacao = 1 and Grupo ='" & ddlGrupoProdutoConsumo.SelectedValue & "'")

        With ddlProdutosConsumo
            ddlProdutosConsumo.SelectedIndex = .Items.IndexOf(.Items.FindByValue(objPrd.Codigo))
        End With

        EstadosCamposConsumo(False)

    End Sub

    Protected Sub imgRemoverConsumo_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)

        Dim img As ImageButton = CType(sender, ImageButton)
        Dim row As GridViewRow = CType(img.NamingContainer, GridViewRow)

        Dim dtConsumo As DataTable = dtgridConsumoAlt()

        ' Verifica se o DataTable não é Nothing
        If dtConsumo IsNot Nothing Then

            dtConsumo.Rows(row.RowIndex).Delete()

            gridConsumoAlt.DataSource = dtConsumo
            gridConsumoAlt.DataBind()

            lblQuantidadeAlt.Text = "TOTAL = " & dtConsumo.Compute("SUM(Quantidade)", "")

            gravarobjConsumoAlt(dtConsumo.Rows(row.RowIndex))

            LimparCamposInsumos()
            EstadosCamposConsumo(True)

        Else

            gridConsumoAlt.DataSource = Nothing
            gridConsumoAlt.DataBind()

            lblQuantidadeAlt.Text = String.Empty

            LimparCamposInsumos()
            EstadosCamposConsumo(True)

        End If

    End Sub

    Private Sub EstadosCamposConsumo(ByVal bHabilita As Boolean)

        ddlGrupoProdutoConsumo.Enabled = bHabilita
        lnkBuscaProdutoConsumo.Enabled = bHabilita
        ddlProdutosConsumo.Enabled = bHabilita

        btnQuantidade.Visible = bHabilita = False

    End Sub

    Private Sub LimparCamposInsumos()

        txtQuantidadeAlt.Text = String.Empty
        ddlGrupoProdutoConsumo.SelectedIndex = 0
        ddlProdutosConsumo.Items.Clear()
        rbMais.Checked = True
        txtLoteAlt.Text = String.Empty
        txtValidadeAlt.Text = Now.ToString("dd/MM/yyyy")

    End Sub

    Protected Sub btnQuantidade_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Try

            Dim dtConsumo As DataTable = dtgridConsumoAlt()

            Dim total As Decimal = dtConsumo.Compute("SUM(Quantidade)", "") _
                                   - CDec(dtConsumo.Rows(gridConsumoAlt.SelectedIndex).Item("Quantidade")) _
                                   + CDec(txtQuantidadeAlt.Text)

            If total > 100 Then
                MsgBox(Me.Page, "Total do(s) Produto(s) para Consumo não pode ultrapassar 100%.", eTitulo.Info)
                Exit Sub
            End If

            dtConsumo.Rows(gridConsumoAlt.SelectedIndex).Item("Quantidade") = txtQuantidadeAlt.Text
            dtConsumo.Rows(gridConsumoAlt.SelectedIndex).Item("Lote") = txtLoteAlt.Text
            dtConsumo.Rows(gridConsumoAlt.SelectedIndex).Item("Validade") = txtValidadeAlt.Text
            dtConsumo.Rows(gridConsumoAlt.SelectedIndex).Item("Sinal") = IIf(rbMais.Checked, "+", "-")


            gridConsumoAlt.DataSource = dtConsumo
            gridConsumoAlt.DataBind()

            lblQuantidadeAlt.Text = "TOTAL = " & dtConsumo.Compute("SUM(Quantidade)", "")

            gravarobjInsumoAlt(dtConsumo.Rows(gridConsumoAlt.SelectedIndex))

            LimparCamposInsumos()
            EstadosCamposConsumo(True)

        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub ddlGrupoProdutoInsumo_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddlGrupoProdutoInsumo.SelectedIndexChanged
        CarregarProdutoInsumo()
    End Sub

    Private Sub CarregarProdutoInsumo()
        Try
            ddlProdutosInsumo.Enabled = True
            ddl.Carregar(ddlProdutosInsumo, CarregarDDL.Tabela.Produto, " ((ControlarEstoque = 'S') or Produto_Id in ( select pc.Produto_id from ProdutoDeConsumo pc inner join Produtos p on p.Produto_Id = pc.Produto_id where  p.Grupo = '" & ddlGrupoProdutoInsumo.SelectedValue & "')) and Grupo ='" & ddlGrupoProdutoInsumo.SelectedValue & "'")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkBuscaProdutoInsumo_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lnkBuscaProdutoInsumo.Click
        Try
            ucConsultaProduto.Limpar()
            Dim txtNome As TextBox = CType(ucConsultaProduto.FindControlRecursive("txtNome"), TextBox)
            Popup.ConsultaDeProduto(Me.Page, "objProdutoxINSU" & HID.Value, txtNome.ClientID, True)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub ProdutosInsumo()
        Try
            If ddlProdutosInsumo.SelectedIndex = 0 Then
                ddlProdutosInsumo.SelectedIndex = 0
                MsgBox(Me.Page, "Produto para Insumo não foi informado.", eTitulo.Info)
            Else

                Dim dtInsumo As DataTable = dtgridInsumoAlt()

                If dtInsumo.Rows.Count > 0 Then

                    If Not IsDBNull(dtInsumo.Compute("SUM(Quantidade)", "Produto = " & ddlProdutosInsumo.SelectedValue)) Then
                        ddlProdutosInsumo.SelectedIndex = 0
                        MsgBox(Me.Page, "Produto já existe na lista para Insumo.", eTitulo.Info)
                        Exit Sub
                    End If
                End If

                objPrd = New Produto(ddlProdutosInsumo.SelectedValue)

                Dim drItem As DataRow = CType(Session("objInsumoAlt" & HID.Value), DataTable).NewRow()

                drItem("CodigoProdutoProducao") = CodigoProdutoProducaoSelecionado
                drItem("Produto") = objPrd.Codigo
                drItem("NomeProduto") = objPrd.Nome
                drItem("Quantidade") = 0
                If txtQuantidadeIns.Text.Length > 0 Then drItem("Quantidade") = CInt(txtQuantidadeIns.Text)

                CType(Session("objInsumoAlt" & HID.Value), DataTable).Rows.Add(drItem)

                gridInsumoAlt.DataSource = dtgridInsumoAlt()
                gridInsumoAlt.DataBind()

                LimparCamposInsumo()
                EstadosCamposInsumo(True)

            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub gridInsumoAlt_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles gridInsumoAlt.SelectedIndexChanged

        Dim dtInsumo As DataTable = dtgridInsumoAlt()

        txtQuantidadeIns.Text = dtInsumo.Rows(gridInsumoAlt.SelectedIndex).Item("Quantidade")

        objPrd = New Produto(dtInsumo.Rows(gridInsumoAlt.SelectedIndex).Item("Produto"))
        ddlGrupoProdutoInsumo.SelectedValue = objPrd.CodigoGrupo
        ddl.Carregar(ddlProdutosInsumo, CarregarDDL.Tabela.Produto, " Situacao = 1 and Grupo ='" & ddlGrupoProdutoInsumo.SelectedValue & "'")

        With ddlProdutosInsumo
            ddlProdutosInsumo.SelectedIndex = .Items.IndexOf(.Items.FindByValue(objPrd.Codigo))
        End With

        EstadosCamposInsumo(False)

    End Sub

    Private Sub EstadosCamposInsumo(ByVal bHabilita As Boolean)

        ddlGrupoProdutoInsumo.Enabled = bHabilita
        lnkBuscaProdutoInsumo.Enabled = bHabilita
        ddlProdutosInsumo.Enabled = bHabilita

        btnQuantidadeIns.Visible = bHabilita = False

    End Sub

    Private Sub LimparCamposInsumo()

        txtQuantidadeIns.Text = String.Empty
        ddlGrupoProdutoInsumo.SelectedIndex = 0
        ddlProdutosInsumo.Items.Clear()

    End Sub

    Protected Sub imgRemoverInsumo_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)

        Dim img As ImageButton = CType(sender, ImageButton)
        Dim row As GridViewRow = CType(img.NamingContainer, GridViewRow)

        Dim dtInsumo As DataTable = dtgridInsumoAlt()

        dtInsumo.Rows(row.RowIndex).Delete()

        gridInsumoAlt.DataSource = dtInsumo
        gridInsumoAlt.DataBind()

        gravarobjInsumoAlt(dtInsumo.Rows(row.RowIndex))

        LimparCamposInsumos()
        EstadosCamposInsumo(True)

    End Sub

    Protected Sub btnQuantidadeIns_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Try

            Dim dtInsumo As DataTable = dtgridInsumoAlt()

            dtInsumo.Rows(gridInsumoAlt.SelectedIndex).Item("Quantidade") = txtQuantidadeIns.Text

            gridInsumoAlt.DataSource = dtInsumo
            gridInsumoAlt.DataBind()

            gravarobjInsumoAlt(dtInsumo.Rows(gridInsumoAlt.SelectedIndex))

            LimparCamposInsumo()
            EstadosCamposInsumo(True)

        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Function ValidarSelecaoConIns() As Boolean

        Dim dtConsumo As DataTable = dtgridInsumoAlt()

        For Each drItem As DataRow In dtConsumo.Rows
            If drItem("Lote").ToString().Length = 0 AndAlso drItem("Lote") = "0" Then
                MsgBox(Me.Page, "Lote do Produto " & drItem("Produto") & "-" & drItem("NomeProduto") & " não foi informado.", eTitulo.Info)

                Return False
            End If
        Next

        Return True

    End Function

    Protected Sub lnkGravarConsumo_Click(sender As Object, e As EventArgs) Handles lnkGravarConsumo.Click

        If Funcoes.VerificaPermissao("GerenciarOrdemDeProducao", "GRAVAR") Then

            If ValidarSelecaoConIns() Then

                SessaoRecuperaOrdem()

                Dim objOPConIns As New OrdemParaProducao
                objOPConIns = objOP.Clone()

                objOPConIns.IUD = "IN_CON_INS"
                objOPConIns.Validade = CDate(txtDataValidade.Text)
                objOPConIns.MovimentoEstoque = CDate(txtDataProducao.Text)

                Dim totalProduto As Decimal = 0

                objOPConIns.ItensDeConsumo.Clear()
                objOPConIns.ItensDeInsumo.Clear()

                '-------------------------------------------------------------
                For Each drItem As DataRow In dtgridConsumoAlt().Rows
                    objOPXC = New OrdemParaProducaoXConsumo(objOPConIns)

                    objOPXC.IUD = "I"
                    objOPXC.CodigoProdutoProducao = drItem("CodigoProdutoProducao")
                    objOPXC.CodigoProduto = drItem("Produto")
                    objOPXC.Quantidade = drItem("Quantidade")
                    objOPXC.Sinal = drItem("Sinal")

                    totalProduto += objOPXC.Quantidade

                    Dim ds As New DataSet

                    If objOPXC.Produto.ControlarEstoque Then
                        ds = objOPXC.buscarLoteDeFornecedor(objOPConIns.CodigoEmpresa, objOPConIns.EnderecoEmpresa, "'" & objOPXC.CodigoProduto & "'")
                        If ds.Tables(0).Rows.Count = 0 Then
                            MsgBox(Me.Page, "Não foi encontrado o lote do Produto. Ordem de Produção não será gravada. " & objOPXC.CodigoProduto & "-" & objOPXC.Produto.Nome, eTitulo.Info)
                            Exit Sub
                        End If

                        If objOPXC.CodigoProduto = drItem("Produto") Then

                            Dim temEstoqueLote As Boolean = False

                            For Each dr In ds.Tables(0).Rows
                                If dr("Produto") = drItem("Produto") AndAlso FuncoesStrings.NormalizeLote(dr("Lote"), objOPConIns.CodigoEmpresa) = FuncoesStrings.NormalizeLote(drItem("Lote"), objOP.CodigoEmpresa) AndAlso CDate(dr("Validade")).ToString("yyyy-MM-dd") = CDate(drItem("Validade")).ToString("yyyy-MM-dd") Then
                                    temEstoqueLote = True
                                End If
                            Next

                            If Not temEstoqueLote Then
                                MsgBox(Me.Page, "Falta saldo no lote do Produto. Ordem de Produção não será gravada. Produto: " & objOPXC.CodigoProduto & "-" & objOPXC.Produto.Nome & " - Lote: " & drItem("Lote") & " - Validade: " & CDate(drItem("Validade")).ToString("dd/MM/yyyy"), eTitulo.Info)
                                Exit Sub
                            End If

                            objOPXCXLote = New OrdemParaProducaoXConsumoXLote(objOPXC)

                            objOPXCXLote.IUD = "I"
                            objOPXCXLote.Lote = drItem("Lote")
                            objOPXCXLote.Quantidade = drItem("Quantidade")
                            objOPXCXLote.Validade = drItem("Validade")

                            objOPXC.ItensDeConsumoXLote.Add(objOPXCXLote)
                        End If

                    End If

                    objOPConIns.ItensDeConsumo.Add(objOPXC)

                Next

                If objOP.ProdutosdaProducao.Count = 1 AndAlso Not Left(objOP.CodigoEmpresa, 8) = "05272759" Then

                    For Each produto As OrdemParaProducaoXProduto In objOP.ProdutosdaProducao


                        'Só atualizar a Produção se o Peso do Item for ZERO - Furlan - 23/09/2024
                        If produto.FatorConversao = 1 Then
                            produto.IUD = "U"
                            produto.Quantidade += totalProduto
                        End If
                        Exit For

                    Next

                End If

                For Each drItem As DataRow In CType(Session("objInsumoAlt" & HID.Value), DataTable).Rows

                    objOPXI = New OrdemParaProducaoXInsumo(objOPConIns)

                    objOPXI.IUD = "I"
                    objOPXI.CodigoProdutoProducao = drItem("CodigoProdutoProducao")
                    objOPXI.CodigoProduto = drItem("Produto")
                    objOPXI.Quantidade = drItem("Quantidade")

                    objOPConIns.ItensDeInsumo.Add(objOPXI)

                Next

                If objOPConIns.Salvar Then

                    MsgBox(Me.Page, "Ordem de Produção " & objOPConIns.Codigo & " atualizada com Sucesso.", eTitulo.Sucess)

                    LimparConsumo()
                    Dim operacao As Integer = txtSequencia.Text
                    Limpar()
                    txtSequencia.Text = operacao
                    objOP = New OrdemParaProducao(ddlEmpresa.SelectedValue.Split("-")(0), ddlEmpresa.SelectedValue.Split("-")(1), CInt(txtSequencia.Text))
                    RecuperarOrdem()

                Else

                    MsgBox(Me.Page, HttpContext.Current.Session("ssMessage").ToString)

                End If

            End If

        Else
            MsgBox(Me.Page, "Usuário sem permissão para gravar registro.")
        End If

    End Sub

    Protected Sub lnkLimparConsumo_Click(sender As Object, e As EventArgs) Handles lnkLimparConsumo.Click
        LimparConsumo()
    End Sub

    Protected Sub imgSelecionaLoteAlt_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)

        If String.IsNullOrWhiteSpace(txtQuantidadeAlt.Text) Then
            MsgBox(Me.Page, "Quantidade não foi informada", eTitulo.Info)
            Exit Sub
        ElseIf CDec(txtQuantidadeAlt.Text) = 0 Then
            MsgBox(Me.Page, "Quantidade não foi informada", eTitulo.Info)
            Exit Sub
        ElseIf ddlProdutosConsumo.SelectedValue Is Nothing OrElse ddlProdutosConsumo.SelectedValue = "" Then
            MsgBox(Me.Page, "O produto deve ser informado", eTitulo.Info)
            Exit Sub
        End If

        Dim ucConsultaLote = CType(Me.Page.FindControlRecursive("ucConsultaLote"), ucConsultaLote)
        If ucConsultaLote IsNot Nothing Then
            ucConsultaLote.Limpar()
            ucConsultaLote.SetarHID(HID.Value)


            Dim nf As NotaFiscal = New NotaFiscal()
            Dim item As NotaFiscalXItem = New NotaFiscalXItem()
            item.NotaFiscal = nf
            item.NotaFiscal.CodigoEmpresa = ddlEmpresa.SelectedValue.Split("-")(0)
            item.NotaFiscal.EnderecoEmpresa = ddlEmpresa.SelectedValue.Split("-")(1)
            'item.NotaFiscal.CodigoRomaneio = 0
            item.CodigoProduto = CDec(ddlProdutosConsumo.SelectedValue)
            item.QuantidadeFiscal = txtQuantidadeAlt.Text

            ucConsultaLote.MostrarLotesParaOrdemProducao(item.NotaFiscal.CodigoEmpresa, item.NotaFiscal.EnderecoEmpresa, item)

        End If

        Popup.ConsultaDeLote(Me.Page, "objConsultaDeLote" & HID.Value)

    End Sub

    Protected Sub ddlProdutosInsumo_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddlProdutosInsumo.SelectedIndexChanged
        ProdutosInsumo()
    End Sub

End Class