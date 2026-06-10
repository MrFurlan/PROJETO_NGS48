Imports System.Data
Imports System.IO
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports GenCode128
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis
Imports OfficeOpenXml
Imports OfficeOpenXml.Style
Imports System.Drawing
Imports ThoughtWorks.QRCode.Codec

Partial Class Mercadorias
    Inherits BasePage

    Private objProduto As [Lib].Negocio.Produto
    Private Sql As String
    Private ds As DataSet

#Region "Events"

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.setMenu(eModulo.Gestao)
            If Not IsPostBack And IsConnect Then

                If Not UsuarioServidor.KeyCodeActive Then
                    MsgBox(Me.Page, "Sistema com chave de licença expirada. Entre em contato com o suporte.", "~/Gestao.aspx", eTitulo.Info)
                    Exit Sub
                End If

                If Funcoes.VerificaPermissao("Mercadorias", "ACESSAR") Then
                    CarregarGrupos()
                    CarregarUnidadeDeMedida()
                    CarregarSituacao()
                    CarregarEtapas()
                    CarregarEstadoFisico()
                    CarregarEmbalagem()
                    CarregarTipoDaEmbagem()
                    CarregarPesoQuantidade()
                    CarregarControleDeQualidade()
                    CarregarMarcas()
                    CarregarSeguimentos()
                    CarregarCarteiraDeCompras()
                    CarregarCarteiraDeVendas()
                    CarregarTipoDoItem()
                    CarregarCodigoDoGenero()
                    CarregarEspecificacoes()
                    CarregarPlanoConta()
                    CarregarCentroDeCusto()
                    ddl.Carregar(ddlCnae, CarregarDDL.Tabela.CNAE)
                    ddl.Carregar(ddlEPI, CarregarDDL.Tabela.EPI)
                    ddl.Carregar(ddlProcedimento, CarregarDDL.Tabela.ProcedimentoParaProducao)
                    LimparCampos(True)
                    LimparProdxCons()
                    TabContainer1.ActiveTabIndex = 0
                    lnkAtualizar.Parent.Visible = False
                    lnkExcluir.Parent.Visible = False
                Else
                    MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Gestao.aspx")
                    Exit Sub
                End If
            Else
                If ddlCodigoDoGenero IsNot Nothing Then
                    For Each li As ListItem In ddlCodigoDoGenero.Items
                        li.Attributes("title") = li.Text
                    Next
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub rdASim_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            Try
                rdASim.Checked = True
                rdANao.Checked = False
                chkControlarRomaneio.Checked = False
                chkControlarRomaneio.Enabled = False
            Catch ex As Exception
                MsgBox(Me.Page, ex.Message, eTitulo.Erro)
            End Try
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub rdANao_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            SessaoRecuperaProduto()

            If Not objProduto.ProdutoXEmbalagens Is Nothing AndAlso objProduto.ProdutoXEmbalagens.Count > 0 Then
                MsgBox(Me.Page, "Produto com embalagem não pode ser desagrupado.")
                rdASim.Checked = True
                rdANao.Checked = False
            Else
                rdASim.Checked = False
                rdANao.Checked = True
                chkControlarRomaneio.Enabled = True
                If chkControlarPesagem.Checked Then chkControlarRomaneio.Checked = True
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub imgSequencia_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles imgSequencia.Click
        Try
            If Funcoes.VerificaPermissao("Mercadorias", "LEITURA") Then
                If ddlGrupoProduto.SelectedIndex = 0 Then
                    MsgBox(Me.Page, "Grupo do produto deve ser selecionado!")
                ElseIf ddlGrupoProduto.Enabled = False Then
                    Exit Sub
                Else
                    SessaoRecuperaProduto()
                    objProduto.CodigoGrupo = ddlGrupoProduto.SelectedValue
                    objProduto.BuscarSequencia()
                    ddlGrupoProduto.Enabled = False
                    objProduto.IUD = "I"
                    txtCodigoProduto.Text = objProduto.Codigo
                    SessaoSalvaProduto()
                    txtCodigoProduto.Enabled = False
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para consultar sequência.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub chkFitossanitario_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            If chkFitossanitario.Checked Then
                txtCodigoIndea.Enabled = True
            Else
                txtCodigoIndea.Text = ""
                txtCodigoIndea.Enabled = False
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub ddlCodigoDoGenero_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            CarregarSubCodigoDoGenero()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkSelecionarPrd_Click(ByVal sender As Object, ByVal e As EventArgs)
        Try
            Dim lnk As LinkButton = CType(sender, LinkButton)
            Dim row As GridViewRow = CType(lnk.NamingContainer, GridViewRow)
            LimparCampos(False)

            objProduto = New [Lib].Negocio.Produto(gridConsulta.Rows(row.RowIndex).Cells(3).Text())

            objProduto.IUD = "U"
            lblConsolidaProduto.Parent.Visible = True
            lblConsolidaProduto.Text = "Produtos agrupados em " & objProduto.Codigo & " - " & objProduto.Descricao
            gridProdutoAgrupado.DataSource = objProduto.ProdutosAgrupados.ToArray
            gridProdutoAgrupado.DataBind()

            txtCodigoProduto.Enabled = False
            imgSequencia.Enabled = True

            ddlGrupoProduto.SelectedValue = objProduto.CodigoGrupo

            ddlGrupoProduto.Enabled = False
            txtCodigoProduto.Text = objProduto.Codigo

            txtCodigoProdutoTerceiro.Text = objProduto.CodigoProdutoTerceiro

            ddlUnidadeDeMedida.SelectedValue = objProduto.Unidade
            If objProduto.CodigoSituacao > 0 Then ddlSituacao.SelectedValue = objProduto.CodigoSituacao
            txtCodigoIndea.Text = objProduto.ProdutoIndea

            txtNome.Text = objProduto.Nome
            txtDescricao.Text = objProduto.Descricao
            txtInfaDProd.Text = objProduto.InfaDProd

            If objProduto.Etapa > 0 Then ddlEtapas.SelectedValue = objProduto.Etapa
            If objProduto.CodigoEmbalagem > 0 Then ddlEmbalagem.SelectedValue = objProduto.CodigoEmbalagem

            If objProduto.ControlarPesagem Then
                chkControlarPesagem.Checked = True
            Else
                chkControlarPesagem.Checked = False
            End If

            If objProduto.Agrupar = "S" Then
                rdASim.Checked = True
                chkControlarRomaneio.Checked = False
                chkControlarRomaneio.Enabled = False
            Else
                rdANao.Checked = True
                chkControlarRomaneio.Enabled = True
                If objProduto.ControlarRomaneio Then
                    chkControlarRomaneio.Checked = True
                Else
                    chkControlarRomaneio.Checked = False
                End If
            End If

            txtDescricaoMapa.Text = objProduto.DescricaoMapa
            txtNCM.Text = objProduto.NCM

            If Not String.IsNullOrWhiteSpace(objProduto.CodigoCnae) AndAlso (objProduto.CodigoCnae.Length = 7) Then
                ddlCnae.SelectedValue = objProduto.CodigoCnae
            Else
                ddlCnae.SelectedIndex = 0
            End If

            If objProduto.PesoQuantidade.Length > 0 Then ddlPesoQuantidade.SelectedValue = objProduto.PesoQuantidade

            txtEstoqueMinimo.Text = objProduto.EstoqueMinimo.ToString("N4")

            If objProduto.Qualidade > 0 Then ddlQualidade.SelectedValue = objProduto.Qualidade

            txtIPI.Text = objProduto.IPI.ToString("N2")

            chkControlarLote.Checked = objProduto.ControlarLote
            chkControlarEmbalagem.Checked = objProduto.ControlarEmbalagem
            chkFitossanitario.Checked = objProduto.Fitossanitario
            If chkFitossanitario.Checked = True Then txtCodigoIndea.Enabled = True
            chkControlarEstoque.Checked = objProduto.ControlarEstoque
            chkPrecoDePauta.Checked = objProduto.ControlarPrecoDePauta
            chkPecasMeios.Checked = objProduto.ControlarPecas
            chkDecimais.Checked = objProduto.ControlarDecimais
            chkAlmoxarifado.Checked = objProduto.Almoxarifado
            chkPrecoDoProduto.Checked = objProduto.PrecoDoProduto
            chkNumeroDoLote.Checked = objProduto.ControlarNumeroDoLote
            chkAutorizacaoDeRetirada.Checked = objProduto.AutorizacaoDeRetirada
            chkCustoIndireto.Checked = objProduto.CustoIndireto

            If objProduto.CodigoDaMarca > 0 Then ddlMarca.SelectedValue = objProduto.CodigoDaMarca
            If objProduto.CodigoCarteiraCompra.Length > 0 Then ddlCarteiraDeCompras.SelectedValue = objProduto.CodigoCarteiraCompra
            If objProduto.CodigoCarteiraVenda.Length > 0 Then ddlCarteiraDeVendas.SelectedValue = objProduto.CodigoCarteiraVenda
            ddlTipoDoItem.SelectedValue = objProduto.TipoItem

            txtCodigoDoServico.Text = objProduto.CodigoServico
            txtCodigoEX.Text = objProduto.CodigoEX

            ddlCodigoDoGenero.SelectedValue = objProduto.CodigoGenero
            CarregarSubCodigoDoGenero()
            If objProduto.SubCodigoGenero > 0 Then ddlSubCodigoDoGenero.SelectedValue = objProduto.SubCodigoGenero

            ddlEstadoFisico.SelectedValue = objProduto.CodigoEstadoFisico

            txtRegMinAgr.Text = objProduto.RegistroMinisterioAgricultura

            gridUnidadeComercializacao.DataSource = objProduto.UnidadesDeComercializacao
            gridUnidadeComercializacao.DataBind()

            If chkControlarEmbalagem.Checked Then
                ddlPXEmbalagem.Enabled = True
                ddlMaterialEmbalagem.Enabled = True
                txtQuantidadeDaEmbalagem.Enabled = True
                txtPesoBruto.Enabled = True
                txtPesoLiquido.Enabled = True
                lnkAdicionarEmbalagem.Parent.Parent.Visible = True

                If objProduto.ProdutoXEmbalagens.Count > 0 Then
                    gridEmbalagem.DataSource = objProduto.ProdutoXEmbalagens.ToArray()
                    TabEmbalagens.Visible = True
                Else
                    gridEmbalagem.DataSource = Nothing
                End If
            Else
                gridEmbalagem.DataSource = Nothing
            End If
            gridEmbalagem.DataBind()

            lnkAtualizar.Parent.Visible = True

            lnkNovo.Parent.Visible = False

            If objProduto.ExistePedido OrElse objProduto.ExisteNota OrElse objProduto.ExisteProducao OrElse objProduto.ExisteOperacoesXEncargos Then
                ddlUnidadeDeMedida.Enabled = False
                lnkExcluir.Parent.Visible = False
            Else
                ddlUnidadeDeMedida.Enabled = True
                lnkExcluir.Parent.Visible = True
            End If

            ddlUsuarios.Items.Clear()
            If objProduto.UsuarioAlteracao IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(objProduto.UsuarioAlteracao) AndAlso objProduto.UsuarioInclusao <> objProduto.UsuarioAlteracao Then
                ddlUsuarios.Items.Add("Alt.- " & objProduto.UsuarioAlteracao)
                ddlUsuarios.Items.Add("Inc.- " & objProduto.UsuarioInclusao)
            Else
                ddlUsuarios.Items.Add("Inc.- " & objProduto.UsuarioInclusao)
            End If

            SessaoSalvaProduto()

            txtGTIN8.Text = objProduto.Gtin8
            If txtGTIN8.Text.Length > 0 Then
                imgGTIN8.ImageUrl = gerarImgGtin(txtGTIN8.Text)
            End If

            txtGTIN12.Text = objProduto.Gtin12
            If txtGTIN12.Text.Length > 0 Then
                imgGTIN12.ImageUrl = gerarImgGtin(txtGTIN12.Text)
            End If

            txtGTIN13.Text = objProduto.Gtin13
            If txtGTIN13.Text.Length > 0 Then
                imgGTIN13.ImageUrl = gerarImgGtin(txtGTIN13.Text)
            End If

            txtGTIN14.Text = objProduto.Gtin14
            If txtGTIN14.Text.Length > 0 Then
                imgGTIN14.ImageUrl = gerarImgGtin(txtGTIN14.Text)
            End If

            chkDashboard.Checked = objProduto.Dashboard

            If objProduto.CodigoSeguimento > 0 Then ddlSeguimentos.SelectedValue = objProduto.CodigoSeguimento

            If objProduto.ProdutoXEspecificacao.Count > 0 Then AtualizarGridEPrd()

            If objProduto.ProdutoXEPI.Count > 0 Then AtualizarGridEPI()

            If objProduto.ProdutoXProcedimento.Count > 0 Then AtualizarGridProcedimento()

            TabContainer1.ActiveTabIndex = 0

            AtualizarGridProdxCons()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub chkControlarEmbalagem_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            If chkControlarEmbalagem.Checked Then
                ddlPXEmbalagem.Enabled = True
                ddlMaterialEmbalagem.Enabled = True
                txtQuantidadeDaEmbalagem.Enabled = True
                txtPesoBruto.Enabled = True
                txtPesoLiquido.Enabled = True
                lnkAdicionarEmbalagem.Parent.Parent.Visible = True
                TabEmbalagens.Visible = True
            Else
                If gridEmbalagem.Rows.Count > 0 Then
                    chkControlarEmbalagem.Checked = True
                    MsgBox(Me.Page, "Não pode ser desmarcado controlar embalagem com embalagens cadastradas.")
                Else
                    ddlPXEmbalagem.Enabled = False
                    ddlMaterialEmbalagem.Enabled = False
                    txtQuantidadeDaEmbalagem.Enabled = False
                    txtPesoBruto.Enabled = False
                    txtPesoLiquido.Enabled = False
                    lnkAdicionarEmbalagem.Parent.Parent.Visible = False
                    TabEmbalagens.Visible = False
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub imgExcluirEmbalagem_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)
        Try
            If Funcoes.VerificaPermissao("Mercadorias", "EXCLUIR") Then
                Dim img As ImageButton = CType(sender, ImageButton)
                Dim row As GridViewRow = CType(img.NamingContainer, GridViewRow)

                SessaoRecuperaProduto()

                If objProduto.ProdutoXEmbalagens(row.RowIndex).EmbalagemXNotaFiscal Then
                    MsgBox(Me.Page, "Embalagem com Nota Fiscal emitida não pode ser excluída.")
                Else
                    objProduto.ProdutoXEmbalagens(row.RowIndex).IUD = "D"

                    If objProduto.ProdutoXEmbalagens(row.RowIndex).Salvar() Then
                        objProduto.ProdutoXEmbalagens.Remove(objProduto.ProdutoXEmbalagens(row.RowIndex))
                        SessaoSalvaProduto()
                        MsgBox(Me.Page, "Embalagem excluída com Sucesso.", eTitulo.Sucess)
                        Limpar_PxE()
                    End If
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para excluir registros.")
                Exit Sub
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub chkControlarRomaneio_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            If chkControlarRomaneio.Checked = True AndAlso chkControlarPesagem.Checked = False Then
                chkControlarRomaneio.Checked = False
                MsgBox(Me.Page, "Romaneio só pode ser marcado se controlar pesagem estiver selecionado.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub chkControlarPesagem_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            If chkControlarPesagem.Checked = False AndAlso chkControlarRomaneio.Checked = True Then
                chkControlarRomaneio.Checked = False
                MsgBox(Me.Page, "Romaneio só pode ser marcado se controlar pesagem estiver selecionado, o mesmo será desmarcado.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub imgExcluir_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)
        Try
            If Funcoes.VerificaPermissao("Mercadorias", "EXCLUIR") Then
                Dim btnExcluir As ImageButton = CType(sender, ImageButton)
                Dim Gridrow As GridViewRow = CType(btnExcluir.NamingContainer, GridViewRow)

                SessaoRecuperaProduto()

                objProduto.ProdutosAgrupados(Gridrow.RowIndex).IUD = "D"
                If objProduto.ProdutosAgrupados(Gridrow.RowIndex).Salvar Then
                    objProduto.ProdutosAgrupados.RemoveAt(Gridrow.RowIndex)
                    SessaoSalvaProduto()
                    gridProdutoAgrupado.DataSource = objProduto.ProdutosAgrupados.ToArray
                    gridProdutoAgrupado.DataBind()
                    MsgBox(Me.Page, "Sucesso na exclusão.", eTitulo.Sucess)
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para excluir registros.")
                Exit Sub
            End If

        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub ddlGrupoProduto_SelectedIndexChanged(ByVal sender As Object, ByVal e As EventArgs) Handles ddlGrupoProduto.SelectedIndexChanged
        Try
            SessaoRecuperaProduto()
            If objProduto.IUD <> "U" Then
                imgSequencia_Click(imgSequencia, Nothing)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkLimpar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkLimpar.Click
        Try
            LimparCampos(True)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkPdf_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkPdf.Click
        Try
            EmitirRelatorio(True)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkExcel_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkExcel.Click
        Try
            EmitirRelatorio(False)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkExcelDados_Click(sender As Object, e As EventArgs) Handles lnkExcelDados.Click
        Try
            EmitirRelatorioDados() 'Excel Dados
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub EmitirRelatorio(ByVal Pdf As Boolean)
        Try
            If Funcoes.VerificaPermissao("Mercadorias", "RELATORIO") Then
                getDataset()

                Dim parameters = New Dictionary(Of String, Object)()
                parameters.Add("Titulo", "Relatório De Produtos e Mercadorias.")
                parameters.Add("ConsultaParametros", getParam())

                Funcoes.BindReport(Me.Page, ds, "Cr_Mercadorias", IIf(Pdf, eExportType.PDF, eExportType.ExcelCrystal), parameters)
            Else
                MsgBox(Me.Page, "Usuário sem permissão para emitir relatório.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Function getDataset() As DataSet
        Dim Where As String = "Where"

        Sql = "SELECT Produtos.Produto_Id, Produtos.Grupo, Produtos.Nome, Produtos.Descricao, Produtos.Unidade, " & vbCrLf &
                         "GruposDeEstoques.Descricao AS DescricaoGrp, Produtos.NCM, isnull(produtos.RegMinAgr, '') as RegistroMinAgr " & vbCrLf &
                         " FROM Produtos INNER JOIN " & vbCrLf &
                         "      GruposDeEstoques ON Produtos.Grupo = GruposDeEstoques.Grupo_Id " & vbCrLf

        'If txtCodigoProduto.Text.Length > 0 Then
        '    strSQL &= Where & " UPPER(Produtos.Produto_Id) LIKE '" & txtCodigoProduto.Text & "%' "
        '    Where = "And"
        'End If

        If txtNome.Text.Length > 0 Then
            Sql &= Where & " Produtos.Nome LIKE '" & RTrim(txtNome.Text) & "%' "
            Where = "And"
        End If
        If ddlGrupoProduto.SelectedIndex > 0 Then
            Sql &= Where & " Produtos.Grupo LIKE '" & RTrim(ddlGrupoProduto.SelectedValue) & "%' "
            Where = "And"
        End If
        If txtDescricao.Text.Length > 0 Then
            Sql &= Where & " Produtos.Descricao LIKE '" & RTrim(txtDescricao.Text) & "%' "
            Where = "And"
        End If
        If ddlSituacao.SelectedIndex > 0 Then
            Sql &= Where & " Produtos.Situacao = " & ddlSituacao.SelectedValue & " "
            Where = "And"
        End If
        If ddlEtapas.SelectedIndex > 0 Then
            Sql &= Where & " Produtos.Etapa = " & ddlEtapas.SelectedValue & " "
            Where = "And"
        End If
        If txtNCM.Text.Length > 0 Then
            Sql &= Where & " Produtos.NCM LIKE '%" & RTrim(txtNCM.Text) & "%' "
            Where = "And"
        End If

        If ddlCnae.SelectedIndex > 0 Then
            Sql &= Where & " Produtos.Cnae LIKE '%" & RTrim(ddlCnae.SelectedValue) & "%' "
            Where = "And"
        End If
        ds = Banco.ConsultaDataSet(Sql, "Mercadorias")

        Return ds
    End Function

    Private Sub EmitirRelatorioDados()
        Dim obj As Usuario = UsuarioServidor.Carregar(Session("ssNomeUsuario"))

        Dim dt As DataTable = New DataTable()

        ds = getDataset()
        dt = ds.Tables(0)

        If dt.Rows.Count = 0 Then
            MsgBox(Me.Page, "Nenhum registro encontrado para essa seleção.")
            Exit Sub
        End If

        Dim fileName As String = Server.MapPath("~/PlanilhasExcel/" & Funcoes.GeraNomeArquivo & ".xlsx")

        If File.Exists(fileName) Then
            File.Delete(fileName)
        End If

        Using arquivo As New FileStream(fileName, FileMode.CreateNew)
            Using package As New ExcelPackage(arquivo)

                'criando planilha e título
                Dim worksheet As ExcelWorksheet = package.Workbook.Worksheets.Add("Relatório de produtos.")

                'criando linha com o cabeçalho da planilha
                Dim rowIndex As Integer = 1
                Dim columnIndex As Integer = 1

                'criando linha que informa o nome da empresa e o cnpj
                worksheet.Cells(rowIndex, columnIndex).Style.Font.Bold = True
                worksheet.Cells(rowIndex, columnIndex).Style.Font.Color.SetColor(Color.Black)
                worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                worksheet.Cells(rowIndex, columnIndex).Value = String.Format("{0} - {1}", obj.Empresa.Nome, Funcoes.FormatarCpfCnpj(obj.Empresa.Codigo))
                worksheet.Cells(String.Format("A{0}:D{0}", rowIndex)).Merge = True
                worksheet.Cells(String.Format("A{0}:D{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                rowIndex += 1

                'criando linha que informa a cidade e o estado da empresa
                worksheet.Cells(rowIndex, columnIndex).Style.Font.Bold = True
                worksheet.Cells(rowIndex, columnIndex).Style.Font.Color.SetColor(Color.Black)
                worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                worksheet.Cells(rowIndex, columnIndex).Value = String.Format("{0}/{1}", obj.Empresa.Cidade, obj.Empresa.CodigoEstado)
                worksheet.Cells(String.Format("A{0}:D{0}", rowIndex)).Merge = True
                worksheet.Cells(String.Format("A{0}:D{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                rowIndex += 1

                'criando linha que informa o título do relatório
                worksheet.Cells(rowIndex, columnIndex).Style.Font.Bold = True
                worksheet.Cells(rowIndex, columnIndex).Style.Font.Color.SetColor(Color.Red)
                worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                worksheet.Cells(rowIndex, columnIndex).Value = String.Format("{0}", "PRODUTOS")
                worksheet.Cells(String.Format("A{0}:D{0}", rowIndex)).Merge = True
                worksheet.Cells(String.Format("A{0}:D{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                rowIndex += 1

                'criando linha que informa o período selecionado na página
                worksheet.Cells(rowIndex, columnIndex).Style.Font.Bold = True
                worksheet.Cells(rowIndex, columnIndex).Style.Font.Color.SetColor(Color.Black)
                worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                worksheet.Cells(rowIndex, columnIndex).Value = String.Format("{0}", "Data : " & Now().ToString("dd-MM-yyyy"))
                worksheet.Cells(String.Format("A{0}:D{0}", rowIndex)).Merge = True
                worksheet.Cells(String.Format("A{0}:D{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                rowIndex += 1

                'criando cabeçalho da planilha com os dados do dataset
                'criando linha com o cabeçalho da planilha
                For Each col As DataColumn In dt.Columns
                    worksheet.Cells(rowIndex, columnIndex).Value = col.ColumnName
                    columnIndex += 1
                Next

                'criando auto filtro na planilha
                worksheet.Cells("A5:F" & rowIndex).AutoFilter = True

                'formatando células numéricas
                worksheet.Cells(String.Format("H{0}", rowIndex)).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"
                worksheet.Cells(String.Format("I{0}", rowIndex)).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"

                'aplicando formatação nas células do cabeçalho
                Using range = worksheet.Cells(rowIndex, 1, rowIndex, dt.Columns.Count)
                    range.Style.Font.Bold = True
                    range.Style.Fill.PatternType = ExcelFillStyle.Solid
                    range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
                    range.Style.Font.Color.SetColor(Color.White)
                End Using
                rowIndex += 1

                ' criando conteúdo da planilha com os dados do dataset
                For Each row As DataRow In dt.Rows
                    columnIndex = 1

                    For Each col As DataColumn In dt.Columns
                        worksheet.Cells(rowIndex, columnIndex).Value = row(col.ColumnName)
                        columnIndex += 1
                    Next

                    'formatando células datas
                    worksheet.Cells(String.Format("A{0}", rowIndex)).Style.Numberformat.Format = "dd/MM/yyyy"

                    'aplicando formatação nas células do conteúdo
                    If rowIndex Mod 2 = 0 Then
                        Using range = worksheet.Cells(rowIndex, 1, rowIndex, dt.Columns.Count)
                            range.Style.Fill.PatternType = ExcelFillStyle.Solid
                            range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(153, 204, 255))
                        End Using
                    End If
                    rowIndex += 1
                Next

                'setando autofit nas células da planilha
                worksheet.Cells.AutoFitColumns(0)

                'congelando quinta linha (cabeçalho)
                worksheet.View.FreezePanes(6, 1)

                'salvando planilha do excel
                package.Save()
            End Using
        End Using
        Funcoes.AbrirExcel(Me.Page, "PlanilhasExcel/" & Path.GetFileName(fileName))
        'download do arquivo pelo browser
        'Download(fileName)
    End Sub

    Protected Sub lnkNovo_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkNovo.Click
        Try
            If Funcoes.VerificaPermissao("Mercadorias", "GRAVAR") Then
                GravarProduto("I")
            Else
                MsgBox(Me.Page, "Usuário sem permissão para gravar registro.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAtualizar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkAtualizar.Click
        Try
            If Funcoes.VerificaPermissao("Mercadorias", "ALTERAR") Then
                GravarProduto("U")
            Else
                MsgBox(Me.Page, "Usuário sem permissão para alterar registro.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkExcluir_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkExcluir.Click
        Try
            If Funcoes.VerificaPermissao("Mercadorias", "EXCLUIR") Then
                SessaoRecuperaProduto()

                If objProduto.IUD = "D" AndAlso objProduto.ExistePedido Then
                    MsgBox(Me.Page, "Produto com lançamento em Pedido(s) não pode ser excluído.")
                ElseIf objProduto.IUD = "D" AndAlso objProduto.ExisteNota Then
                    MsgBox(Me.Page, "Produto com lançamento em Nota(s) não pode ser excluído.")
                ElseIf objProduto.IUD = "D" AndAlso objProduto.ExisteProducao Then
                    MsgBox(Me.Page, "Produto com lançamento em Produção não pode ser excluído.")
                ElseIf objProduto.IUD = "D" AndAlso objProduto.ExisteOperacoesXEncargos Then
                    MsgBox(Me.Page, "Produto com lançamento em OperacoesXEncargos não pode ser excluído.")
                ElseIf objProduto.IUD = "D" AndAlso Not objProduto.ProdutoXEmbalagens Is Nothing AndAlso objProduto.ProdutoXEmbalagens.Count > 0 Then
                    MsgBox(Me.Page, "Produto com embalagem não pode ser excluído.")
                Else
                    objProduto.IUD = "D"
                    If objProduto.Salvar() Then
                        MsgBox(Me.Page, "Produto Excluído com Sucesso.", eTitulo.Sucess)
                        LimparCampos(True)
                    End If
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para excluir registro.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkConsultar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkConsultar.Click
        Try
            If Funcoes.VerificaPermissao("Mercadorias", "LEITURA") Then
                Dim pProduto As New [Lib].Negocio.Produto()
                Dim Grupo As String = ""
                Dim Agrupar As String = ""
                Dim ControlarLote As String = ""
                Dim ControlarEmbalagem As String = ""
                Dim Fitossanitario As String = ""

                If rdASim.Checked = True Then Agrupar = "S"
                If rdANao.Checked = True Then Agrupar = "N"
                If chkControlarLote.Checked Then ControlarLote = "S"
                If chkControlarEmbalagem.Checked Then ControlarEmbalagem = "S"
                If chkFitossanitario.Checked Then Fitossanitario = "S"

                If txtCodigoProduto.Text.Length > 0 Then pProduto.Codigo = txtCodigoProduto.Text
                If ddlGrupoProduto.SelectedIndex > 0 Then
                    Grupo = ddlGrupoProduto.SelectedValue
                    pProduto.Codigo = ""
                End If

                If ddlUnidadeDeMedida.SelectedIndex > 0 Then pProduto.Unidade = ddlUnidadeDeMedida.SelectedValue
                If ddlSituacao.SelectedIndex > 0 Then pProduto.CodigoSituacao = ddlSituacao.SelectedValue
                If txtCodigoIndea.Text.Length > 0 Then pProduto.ProdutoIndea = txtCodigoIndea.Text
                If txtNome.Text.Length > 0 Then pProduto.Nome = txtNome.Text
                If txtDescricao.Text.Length > 0 Then pProduto.Descricao = txtDescricao.Text
                If ddlEtapas.SelectedIndex > 0 Then pProduto.Etapa = ddlEtapas.SelectedValue
                If ddlEmbalagem.SelectedIndex > 0 Then pProduto.CodigoEmbalagem = ddlEmbalagem.SelectedValue
                If txtDescricaoMapa.Text.Length > 0 Then pProduto.DescricaoMapa = txtDescricaoMapa.Text
                If txtNCM.Text.Length > 0 Then pProduto.NCM = txtNCM.Text
                'If txtCnae.Text.Length > 0 Then pProduto.Cnae = txtCnae.Text
                If ddlCnae.SelectedIndex > 0 Then pProduto.CodigoCnae = ddlCnae.SelectedValue
                If ddlPesoQuantidade.SelectedIndex > 0 Then pProduto.PesoQuantidade = ddlPesoQuantidade.SelectedValue
                If chkControlarEstoque.Checked Then pProduto.ControlarEstoque = True
                If chkPrecoDePauta.Checked Then pProduto.ControlarPrecoDePauta = True
                If chkPecasMeios.Checked Then pProduto.ControlarPecas = True
                If chkDecimais.Checked Then pProduto.ControlarDecimais = True
                If chkAlmoxarifado.Checked Then pProduto.Almoxarifado = True
                If chkPrecoDoProduto.Checked Then pProduto.PrecoDoProduto = True
                If ddlMarca.SelectedIndex > 0 Then pProduto.CodigoDaMarca = ddlMarca.SelectedValue
                If ddlCodigoDoGenero.SelectedIndex > 0 Then pProduto.CodigoGenero = ddlCodigoDoGenero.SelectedValue
                If chkNumeroDoLote.Checked Then pProduto.ControlarNumeroDoLote = True
                If chkAutorizacaoDeRetirada.Checked Then pProduto.AutorizacaoDeRetirada = True
                If chkCustoIndireto.Checked Then pProduto.CustoIndireto = True
                If chkDashboard.Checked Then pProduto.Dashboard = True
                If ddlMarca.SelectedIndex > 0 Then pProduto.CodigoDaMarca = ddlMarca.SelectedValue

                Dim ListProdutos As New [Lib].Negocio.ListProduto(Grupo, Agrupar, ControlarLote, ControlarEmbalagem, Fitossanitario, "", False, pProduto)

                If ListProdutos.Count > 0 Then
                    gridConsulta.DataSource = ListProdutos.ToArray
                    gridConsulta.DataBind()
                    TabContainer1.ActiveTabIndex = 7
                Else
                    MsgBox(Me.Page, "Não foi encontrado nenhum Produto com os parâmetro(s) escolhido(s).")
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para consultar registro.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "Mercadorias")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkNovoEPrd_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkNovoEPrd.Click
        Try
            If Funcoes.VerificaPermissao("ProdutoXEspecificacao", "GRAVAR") Then

                If Not validarEPrd() Then Exit Sub

                SessaoRecuperaProduto()

                Dim objPrdXE = New ProdutoXEspecificacao(objProduto)
                objPrdXE.IUD = "I"
                objPrdXE.CodigoEspecificacao = ddlEspecificacao.SelectedValue
                objPrdXE.FaixaInicial = CDec(txtFaixaInicial.Text)
                objPrdXE.FaixaFinal = CDec(txtFaixaFinal.Text)
                objPrdXE.Ativo = True

                If objProduto.IUD = "I" Then

                    objProduto.ProdutoXEspecificacao.Add(objPrdXE)

                ElseIf objProduto.IUD = "U" Then

                    If Not objPrdXE.Salvar() Then
                        MsgBox(Me.Page, HttpContext.Current.Session("ssMessage").ToString())
                        Exit Sub
                    End If
                End If

                SessaoSalvaProduto()

                LimparEPrd()
            Else
                MsgBox(Me.Page, "Usuário sem permissão para gravar registro.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAtualizarEPrd_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkAtualizarEPrd.Click
        Try
            If Funcoes.VerificaPermissao("ProdutoXEspecificacao", "ALTERAR") Then

                If Not validarEPrd() Then Exit Sub

                SessaoRecuperaProduto()

                Dim objPrdXE = New ProdutoXEspecificacao(objProduto)
                objPrdXE.IUD = "U"
                objPrdXE.CodigoEspecificacao = ddlEspecificacao.SelectedValue
                objPrdXE.FaixaInicial = CDec(txtFaixaInicial.Text)
                objPrdXE.FaixaFinal = CDec(txtFaixaFinal.Text)
                objPrdXE.Ativo = True

                If Not objPrdXE.Salvar() Then
                    MsgBox(Me.Page, HttpContext.Current.Session("ssMessage").ToString())
                    Exit Sub
                End If

                SessaoSalvaProduto()

                LimparEPrd()
            Else
                MsgBox(Me.Page, "Usuário sem permissão para alterar registro.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkExcluirEPrd_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkExcluirEPrd.Click
        Try
            If Funcoes.VerificaPermissao("ProdutoXEspecificacao", "EXCLUIR") Then
                SessaoRecuperaProduto()

                Dim objPrdXE = New ProdutoXEspecificacao(objProduto)
                objPrdXE.IUD = "D"
                objPrdXE.CodigoEspecificacao = ddlEspecificacao.SelectedValue
                objPrdXE.FaixaInicial = CDec(txtFaixaInicial.Text)
                objPrdXE.FaixaFinal = CDec(txtFaixaFinal.Text)
                objPrdXE.Ativo = False

                If Not objPrdXE.Salvar() Then
                    MsgBox(Me.Page, HttpContext.Current.Session("ssMessage").ToString())
                    Exit Sub
                End If

                SessaoSalvaProduto()

                LimparEPrd()
            Else
                MsgBox(Me.Page, "Usuário sem permissão para excluir registro.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkLimparEPrd_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkLimparEPrd.Click
        LimparEPrd()
    End Sub

    Protected Sub lnkSelecionarEPrd_Click(ByVal sender As Object, ByVal e As EventArgs)
        Try
            Dim lnk As LinkButton = CType(sender, LinkButton)
            Dim row As GridViewRow = CType(lnk.NamingContainer, GridViewRow)

            SessaoRecuperaProduto()

            If Not objProduto.ProdutoXEspecificacao(row.RowIndex).Existe Then
                MsgBox(Me.Page, "Para alterar essa Especificação deve Ativar a mesma na Tabela de Especificações do Produto.")
                Exit Sub
            End If

            ddlEspecificacao.SelectedValue = objProduto.ProdutoXEspecificacao(row.RowIndex).CodigoEspecificacao
            txtFaixaInicial.Text = objProduto.ProdutoXEspecificacao(row.RowIndex).FaixaInicial
            txtFaixaFinal.Text = objProduto.ProdutoXEspecificacao(row.RowIndex).FaixaFinal

            ddlEspecificacao.Enabled = False

            If Not objProduto.ProdutoXEspecificacao(row.RowIndex).Ativo Then
                txtFaixaInicial.Enabled = False
                txtFaixaFinal.Enabled = False
            End If

            lnkAtualizarEPrd.Parent.Visible = True
            lnkExcluirEPrd.Parent.Visible = True
            lnkNovoEPrd.Parent.Visible = False

        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Function validarEPrd() As Boolean
        If ddlEspecificacao.SelectedIndex = 0 Then
            MsgBox(Me.Page, "Especificação do Produto não foi selecionada.")
            Return False
        ElseIf String.IsNullOrWhiteSpace(txtFaixaInicial.Text) OrElse CDec(txtFaixaInicial.Text) = 0 Then
            MsgBox(Me.Page, "Faixa inicial não foi informada.")
            Return False
        ElseIf String.IsNullOrWhiteSpace(txtFaixaFinal.Text) OrElse CDec(txtFaixaFinal.Text) = 0 Then
            MsgBox(Me.Page, "Faixa final não foi informada.")
            Return False
        Else
            Return True
        End If
    End Function

    Private Sub AtualizarGridEPrd()
        SessaoRecuperaProduto()

        Dim listProdutoXEspecificacao = New ListProdutoXEspecificacao(objProduto)

        objProduto.ProdutoXEspecificacao = listProdutoXEspecificacao

        gridEspecificacaoDoProduto.DataSource = objProduto.ProdutoXEspecificacao.ToArray()
        gridEspecificacaoDoProduto.DataBind()

        Dim i As Integer = 0
        While i < gridEspecificacaoDoProduto.Rows.Count
            If gridEspecificacaoDoProduto.Rows(i).Cells(4).Text.ToUpper() = "TRUE" Then
                gridEspecificacaoDoProduto.Rows(i).Cells(4).Text = "SIM"
            Else
                gridEspecificacaoDoProduto.Rows(i).Cells(4).Text = "NÃO"
            End If

            i += 1
        End While

        SessaoSalvaProduto()
    End Sub

    Private Sub LimparEPrd()
        lnkAtualizarEPrd.Parent.Visible = False
        lnkExcluirEPrd.Parent.Visible = False
        lnkNovoEPrd.Parent.Visible = True

        ddlEspecificacao.Enabled = True
        txtFaixaInicial.Enabled = True
        txtFaixaFinal.Enabled = True

        ddlEspecificacao.SelectedIndex = 0
        txtFaixaInicial.Text = String.Empty
        txtFaixaFinal.Text = String.Empty

        gridEspecificacaoDoProduto.DataSource = Nothing
        gridEspecificacaoDoProduto.DataBind()

        If Not String.IsNullOrWhiteSpace(txtCodigoProduto.Text) Then
            AtualizarGridEPrd()
        End If

    End Sub

    Private Sub carregarDDLobjProduto()
        SessaoRecuperaProduto()
    End Sub

    Private Function atulizarGtinProduto() As Boolean
        SessaoRecuperaProduto()

        'ddlCodeGtin.SelectedIndex = 1
        'objProduto.Gtin8 = ddlCodeGtin.SelectedValue
        'ddlCodeGtin.SelectedIndex = 2
        'objProduto.Gtin12 = ddlCodeGtin.SelectedValue
        'ddlCodeGtin.SelectedIndex = 3
        'objProduto.Gtin13 = ddlCodeGtin.SelectedValue
        'ddlCodeGtin.SelectedIndex = 4
        'objProduto.Gtin14 = ddlCodeGtin.SelectedValue

        objProduto.Gtin8 = txtGTIN8.Text
        objProduto.Gtin12 = txtGTIN12.Text
        objProduto.Gtin13 = txtGTIN13.Text
        objProduto.Gtin14 = txtGTIN14.Text

    End Function

    Private Function gerarImgGtin(cod As String) As String
        Dim sCodeOrdemDeCarregamento As String

        sCodeOrdemDeCarregamento = String.Format("{0}", cod)

        'BarCode
        Dim imgBarcode As Image = Code128Rendering.MakeBarcodeImage(sCodeOrdemDeCarregamento, 1, True)
        Dim imgConvertBarcode As New ImageConverter()

        Dim ByteBarCode As Byte() = DirectCast(imgConvertBarcode.ConvertTo(imgBarcode, GetType(Byte())), Byte())
        Dim imgBarras64 As String = Convert.ToBase64String(ByteBarCode)

        Return "data:image/png;base64," + imgBarras64
    End Function

    Protected Sub txtGTIN8_TextChanged(sender As Object, e As EventArgs) Handles txtGTIN8.TextChanged
        Try

            imgGTIN8.ImageUrl = ""
            If txtGTIN8.Text.Length > 0 Then
                imgGTIN8.ImageUrl = gerarImgGtin(txtGTIN8.Text)
            End If

        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub txtGTIN12_TextChanged(sender As Object, e As EventArgs) Handles txtGTIN12.TextChanged
        Try

            imgGTIN12.ImageUrl = ""
            If txtGTIN12.Text.Length > 0 Then
                imgGTIN12.ImageUrl = gerarImgGtin(txtGTIN12.Text)
            End If

        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub txtGTIN13_TextChanged(sender As Object, e As EventArgs) Handles txtGTIN13.TextChanged
        Try

            imgGTIN13.ImageUrl = ""
            If txtGTIN13.Text.Length > 0 Then
                imgGTIN13.ImageUrl = gerarImgGtin(txtGTIN13.Text)
            End If

        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub txtGTIN14_TextChanged(sender As Object, e As EventArgs) Handles txtGTIN14.TextChanged
        Try

            imgGTIN14.ImageUrl = ""
            If txtGTIN14.Text.Length > 0 Then
                imgGTIN14.ImageUrl = gerarImgGtin(txtGTIN14.Text)
            End If

        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnAdicionarUnidComercializacao_Click(sender As Object, e As System.Web.UI.ImageClickEventArgs) Handles btnAdicionarUnidComercializacao.Click
        Try
            If Funcoes.VerificaPermissao("Mercadorias", "GRAVAR") Then
                SessaoRecuperaProduto()

                If txtPesoDoProduto.Text.Length = 0 OrElse CType(txtPesoDoProduto.Text, Decimal) = 0 Then
                    MsgBox(Me.Page, "Informe o peso do produto!")
                    Exit Sub
                End If

                If ddlUnidadedeComercializacao.SelectedValue.Length = 0 Then
                    MsgBox(Me.Page, "Informe a unidade de comercialização!")
                    Exit Sub
                End If

                If ValidaUnidadeDeComercializacao(False) Then
                    Dim unid As ProdutosXUnidadeDeComercializacao = objProduto.UnidadesDeComercializacao.Where(Function(s) s.CodigoUnidade = ddlUnidadedeComercializacao.SelectedValue And s.FatorConversao = txtPesoDoProduto.Text).FirstOrDefault

                    If unid IsNot Nothing Then
                        objProduto.UnidadesDeComercializacao.Remove(objProduto.UnidadesDeComercializacao.Where(Function(s) s.CodigoUnidade = ddlUnidadedeComercializacao.SelectedValue And s.FatorConversao = txtPesoDoProduto.Text).FirstOrDefault)
                        unid.IUD = "U"
                    Else
                        unid = New ProdutosXUnidadeDeComercializacao(objProduto)
                        unid.IUD = "I"
                    End If

                    unid.CodigoProduto = objProduto.Codigo
                    unid.CodigoUnidade = ddlUnidadedeComercializacao.SelectedValue
                    unid.FatorConversao = txtPesoDoProduto.Text
                    unid.PesoDaEmbalagem = txtPesoDaEmbalagem.Text
                    objProduto.UnidadesDeComercializacao.Add(unid)

                    SessaoSalvaProduto()

                    gridUnidadeComercializacao.DataSource = objProduto.UnidadesDeComercializacao.Where(Function(s) s.IUD <> "D")
                    gridUnidadeComercializacao.DataBind()

                    ddlUnidadedeComercializacao.SelectedIndex = 0
                    txtPesoDoProduto.Text = 0
                    txtPesoDaEmbalagem.Text = 0

                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para gravar registros.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnExcluirUnidComercializacao_Click(sender As Object, e As System.Web.UI.ImageClickEventArgs)
        Try
            If Funcoes.VerificaPermissao("Mercadorias", "EXCLUIR") Then
                SessaoRecuperaProduto()
                Dim row As GridViewRow = CType(CType(sender, ImageButton).NamingContainer, GridViewRow)

                Dim codigoUnidade As String = gridUnidadeComercializacao.DataKeys(row.RowIndex).Values("CodigoUnidade").ToString()
                Dim fatorConversao As Decimal = gridUnidadeComercializacao.DataKeys(row.RowIndex).Values("FatorConversao")

                Dim unid As New ProdutosXUnidadeDeComercializacao(objProduto.Codigo, codigoUnidade, fatorConversao)

                If unid IsNot Nothing AndAlso unid.CodigoProduto IsNot Nothing AndAlso unid.CodigoProduto.Length > 0 Then
                    If unid.VerificarRelacionamento() Then
                        MsgBox(Me.Page, "O sistema não pode apagar a unidade de comercialização porque a mesma já está vinculada a um pedido!")
                        Exit Sub
                    End If
                End If

                If ddlUnidadeDeMedida.SelectedValue.Length > 0 Then
                    Dim unidadePadrao As String = ddlUnidadeDeMedida.SelectedItem.Text.Split("-")(0).Trim.Replace(".", "")

                    If unidadePadrao = unid.CodigoUnidade Then
                        If objProduto.UnidadesDeComercializacao.Where(Function(u) u.CodigoUnidade = unidadePadrao And u.IUD <> "D").Count <= 1 Then
                            MsgBox(Me.Page, "O sistema não pode apagar a unidade de comercialização igual à unidade padrão do produto! É necessário ter no mínimo uma unidade cadastrada.")
                            Exit Sub
                        End If
                    End If
                End If

                ' Encontra todos os itens que correspondem à condição
                Dim itensParaRemover = objProduto.UnidadesDeComercializacao.
                   Where(Function(s) s.CodigoUnidade = codigoUnidade And s.FatorConversao = fatorConversao).ToList()

                ' Remove cada item encontrado
                For Each item In itensParaRemover

                    If objProduto.IUD = "I" Then
                        objProduto.UnidadesDeComercializacao.Remove(item)
                    Else
                        item.IUD = "D"
                    End If

                Next

                SessaoSalvaProduto()

                gridUnidadeComercializacao.DataSource = objProduto.UnidadesDeComercializacao.Where(Function(s) s.IUD <> "D")
                gridUnidadeComercializacao.DataBind()

            Else
                MsgBox(Me.Page, "Usuário sem permissão para excluir registros.")
                Exit Sub
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnAtualizarUnidComercializacao_Click(sender As Object, e As System.Web.UI.ImageClickEventArgs)
        Try
            If Funcoes.VerificaPermissao("Mercadorias", "ALTERAR") Then
                SessaoRecuperaProduto()
                Dim row As GridViewRow = CType(CType(sender, ImageButton).NamingContainer, GridViewRow)

                Dim codigoUnidade As String = gridUnidadeComercializacao.DataKeys(row.RowIndex).Values("CodigoUnidade").ToString()
                Dim fatorConversao As Decimal = gridUnidadeComercializacao.DataKeys(row.RowIndex).Values("FatorConversao")
                Dim txtPesoDaEmbalagem As TextBox = CType(row.FindControl("PesoDaEmbalagem"), TextBox)
                Dim pesoDaEmbalagem As Decimal

                If Not Decimal.TryParse(txtPesoDaEmbalagem.Text, pesoDaEmbalagem) Then
                    pesoDaEmbalagem = 0
                End If

                ' Encontra todos os itens que correspondem à condição
                Dim itensParaAtualizar = objProduto.UnidadesDeComercializacao.
                       Where(Function(s) s.CodigoUnidade = codigoUnidade And s.FatorConversao = fatorConversao).ToList()

                ' Remove cada item encontrado
                For Each item In itensParaAtualizar

                    item.IUD = "U"
                    item.PesoDaEmbalagem = pesoDaEmbalagem

                Next

                SessaoSalvaProduto()

                gridUnidadeComercializacao.DataSource = objProduto.UnidadesDeComercializacao.Where(Function(s) s.IUD <> "D")
                gridUnidadeComercializacao.DataBind()

            Else
                MsgBox(Me.Page, "Usuário sem permissão para alterar registros.")
                Exit Sub
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub gridUnidadeComercializacao_RowDataBound(sender As Object, e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles gridUnidadeComercializacao.RowDataBound
        Try
            If e.Row.RowType = DataControlRowType.DataRow Then
                If e.Row.DataItem IsNot Nothing Then
                    Dim btn As Button = CType(e.Row.FindControl("BtnAtivo"), Button)
                    If btn.Text = "True" Then
                        btn.Text = "SIM"
                        btn.BackColor = System.Drawing.Color.Green
                        btn.BorderColor = System.Drawing.Color.Green
                        btn.ForeColor = System.Drawing.Color.White
                    Else
                        btn.Text = "NAO"
                        btn.BackColor = System.Drawing.Color.Salmon
                        btn.BorderColor = System.Drawing.Color.Salmon
                        btn.ForeColor = System.Drawing.Color.Black
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub BtnAtivo_Click(sender As Object, e As EventArgs)
        Try
            If Funcoes.VerificaPermissao("Mercadorias", "ALTERAR") Then
                SessaoRecuperaProduto()
                Dim btn As Button = CType(sender, Button)
                Dim row As GridViewRow = CType(btn.NamingContainer, GridViewRow)

                Dim unid As ProdutosXUnidadeDeComercializacao = objProduto.UnidadesDeComercializacao(row.RowIndex)
                If objProduto.IUD <> "I" Then
                    unid.IUD = "U"
                End If
                unid.Ativo = IIf(btn.Text = "SIM", 0, 1)

                If objProduto.UnidadesDeComercializacao.Where(Function(s) s.Ativo).FirstOrDefault() Is Nothing Then
                    MsgBox(Me.Page, "Ao menos 1(uma) unidade de comercialização deve permanecer ativa")
                    Exit Sub
                End If
                SessaoSalvaProduto()

                If btn.Text = "SIM" Then
                    btn.Text = "NAO"
                    btn.BackColor = System.Drawing.Color.Salmon
                    btn.BorderColor = System.Drawing.Color.Salmon
                    btn.ForeColor = System.Drawing.Color.Black
                Else
                    btn.Text = "SIM"
                    btn.BackColor = System.Drawing.Color.Green
                    btn.BorderColor = System.Drawing.Color.Green
                    btn.ForeColor = System.Drawing.Color.White
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para alterar registros.")
                Exit Sub
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAdicionarEmbalagem_Click(sender As Object, e As EventArgs) Handles lnkAdicionarEmbalagem.Click
        Try
            If Funcoes.VerificaPermissao("Mercadorias", "GRAVAR") Then
                SessaoRecuperaProduto()

                If ddlPXEmbalagem.SelectedIndex = 0 Then
                    MsgBox(Me.Page, "Embalagem não foi selecionada.")
                ElseIf ddlMaterialEmbalagem.SelectedIndex = 0 Then
                    MsgBox(Me.Page, "Material da Embalagem não foi selecionado.")
                ElseIf txtQuantidadeDaEmbalagem.Text.Length = 0 Then
                    MsgBox(Me.Page, "Quantidade da Embalagem não foi informada.")
                ElseIf txtPesoBruto.Text.Length = 0 Then
                    MsgBox(Me.Page, "Peso Bruto da Embalagem não foi informado.")
                ElseIf txtPesoLiquido.Text.Length = 0 Then
                    MsgBox(Me.Page, "Peso Liquido da Embalagem não foi informado.")
                Else
                    Dim objPxE As New [Lib].Negocio.ProdutoXEmbalagem(objProduto)

                    If objProduto.IUD = "I" Then
                        objPxE.IUD = "I"
                        objPxE.CodigoEmbalagem = ddlPXEmbalagem.SelectedValue
                        objPxE.CodigoTipoDeEmbalagem = ddlMaterialEmbalagem.SelectedValue
                        objPxE.Capacidade = txtQuantidadeDaEmbalagem.Text
                        objPxE.PesoBruto = txtPesoBruto.Text
                        objPxE.PesoLiquido = txtPesoLiquido.Text
                        objPxE.PesoVariavel = chkPesoVariavel.Checked
                        objProduto.ProdutoXEmbalagens.Add(objPxE)
                        SessaoSalvaProduto()
                        Limpar_PxE()
                    Else
                        objPxE.IUD = "I"
                        objPxE.CodigoEmbalagem = ddlPXEmbalagem.SelectedValue
                        objPxE.CodigoTipoDeEmbalagem = ddlMaterialEmbalagem.SelectedValue
                        objPxE.Capacidade = txtQuantidadeDaEmbalagem.Text
                        objPxE.PesoBruto = txtPesoBruto.Text
                        objPxE.PesoLiquido = txtPesoLiquido.Text
                        objPxE.PesoVariavel = chkPesoVariavel.Checked

                        If objPxE.Salvar() Then
                            objPxE.IUD = ""
                            objProduto.ProdutoXEmbalagens.Add(objPxE)
                            SessaoSalvaProduto()
                            Limpar_PxE()
                            MsgBox(Me.Page, "Embalagem adicionada com Sucesso.", eTitulo.Sucess)
                        End If
                    End If
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para gravar registros.")
                Exit Sub
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAdPrdAgrp_Click(sender As Object, e As EventArgs) Handles lnkAdPrdAgrp.Click
        Try
            If Funcoes.VerificaPermissao("Mercadorias", "GRAVAR") Then
                SessaoRecuperaProduto()
                If HidProdAgr.Value.Length > 0 Then
                    objProduto.ProdutosAgrupados.AdicionarProdutosAgrupados(HidProdAgr.Value)
                    gridProdutoAgrupado.DataSource = objProduto.ProdutosAgrupados.ToArray
                    gridProdutoAgrupado.DataBind()
                Else
                    MsgBox(Me.Page, "Selecione um Produto antes de continuar.")
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para gravar registros.")
                Exit Sub
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnSelPrdAgrp_Click(sender As Object, e As EventArgs) Handles btnSelPrdAgrp.Click
        Try
            Dim ucConsultaProduto As ucConsultaProduto = CType(Me.Page.FindControlRecursive("ucConsultaProduto"), ucConsultaProduto)
            If ucConsultaProduto IsNot Nothing Then
                ucConsultaProduto.Limpar()
                Dim txtNome As TextBox = CType(ucConsultaProduto.FindControlRecursive("txtNome"), TextBox)
                Popup.ConsultaDeProduto(Me.Page, "objProdutoCad" & HID.Value, txtNome.ClientID, True)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub ddlUnidadeDeMedida_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddlUnidadeDeMedida.SelectedIndexChanged
        Try
            If Funcoes.VerificaPermissao("Mercadorias", "GRAVAR") Then

                SessaoRecuperaProduto()

                If Not String.IsNullOrWhiteSpace(objProduto.Codigo) Then
                    If objProduto.UnidadesDeComercializacao IsNot Nothing AndAlso objProduto.UnidadesDeComercializacao.Count > 0 Then

                        objProduto.UnidadesDeComercializacao.RemoveAll(Function(s) s.IUD = "I")
                        objProduto.UnidadesDeComercializacao.ForEach(Function(s)
                                                                         s.IUD = "D"
                                                                         Return ""
                                                                     End Function)
                    End If

                    If ddlUnidadeDeMedida.SelectedValue.Length > 0 Then
                        ddlUnidadedeComercializacao.SelectedValue = ddlUnidadeDeMedida.SelectedValue
                        txtPesoDaEmbalagem.Text = "1,0000"
                        txtPesoDoProduto.Text = "1,0000"
                    Else
                        ddlUnidadedeComercializacao.SelectedValue = ""
                        txtPesoDaEmbalagem.Text = "0,0000"
                        txtPesoDoProduto.Text = "0,0000"
                    End If

                    '2024-12-17
                    'Comentado, o produto não será mais inserido automaticamente
                    'Pq empresas como a Baxi, informa uma quantidade diferente de 1


                    'Dim unid As ProdutosXUnidadeDeComercializacao = objProduto.UnidadesDeComercializacao.Where(Function(s) s.CodigoUnidade = ddlUnidadeDeMedida.SelectedValue).FirstOrDefault

                    'If unid IsNot Nothing Then
                    '    unid.IUD = "U"
                    'Else
                    '    unid = New ProdutosXUnidadeDeComercializacao(objProduto)
                    '    unid.IUD = "I"
                    '    unid.CodigoProduto = objProduto.Codigo
                    '    unid.CodigoUnidade = ddlUnidadeDeMedida.SelectedValue
                    'End If

                    'unid.FatorConversao = 1
                    'objProduto.Unidade = ddlUnidadeDeMedida.SelectedValue
                    'objProduto.UnidadesDeComercializacao.Add(unid)
                    'SessaoSalvaProduto()

                    'gridUnidadeComercializacao.DataSource = objProduto.UnidadesDeComercializacao.Where(Function(s) s.IUD <> "D")
                    'gridUnidadeComercializacao.DataBind()

                Else
                    MsgBox(Me.Page, "Informe o produto.")
                End If

            Else
                MsgBox(Me.Page, "Usuário sem permissão para gravar registros.")
                Exit Sub
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub AtualizarGridProcedimento()

        SessaoRecuperaProduto()

        Dim listPrdXProcedimento = New ListProdutoXProcedimento(txtCodigoProduto.Text)

        gridProcedimento.DataSource = listPrdXProcedimento
        gridProcedimento.DataBind()

        Dim i As Integer = 0
        While i < gridProcedimento.Rows.Count
            If gridProcedimento.Rows(i).Cells(3).Text.ToUpper() = "TRUE" Then
                gridProcedimento.Rows(i).Cells(3).Text = "SIM"
            Else
                gridProcedimento.Rows(i).Cells(3).Text = "NÃO"
            End If

            i += 1
        End While

        SessaoSalvaProduto()

    End Sub

    Private Sub LimparProcedimento()

        ddlProcedimento.Enabled = True
        ddlProcedimento.SelectedIndex = 0

        lnkExcluirProcedimento.Parent.Visible = False
        lnkAtivarProcedimento.Parent.Visible = False
        lnkNovoProcedimento.Parent.Visible = True

        gridProcedimento.DataSource = Nothing
        gridProcedimento.DataBind()

        If Not String.IsNullOrWhiteSpace(txtCodigoProduto.Text) Then
            AtualizarGridProcedimento()
        End If
    End Sub

    Protected Sub lnkNovoProcedimento_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkNovoProcedimento.Click
        Try
            If Funcoes.VerificaPermissao("ProdutoXProcedimento", "GRAVAR") Then

                If ddlProcedimento.SelectedIndex = 0 Then
                    MsgBox(Me.Page, "Procedimento do Produto não foi selecionado.")
                    Exit Sub
                End If

                SessaoRecuperaProduto()

                Dim objOPXProcedimento = New ProdutoXProcedimento()
                objOPXProcedimento.IUD = "I"
                objOPXProcedimento.CodigoProduto = txtCodigoProduto.Text
                objOPXProcedimento.CodigoProcedimento = ddlProcedimento.SelectedValue
                objOPXProcedimento.Ativo = True

                If objProduto.IUD = "I" Then

                    objProduto.ProdutoXProcedimento.Add(objOPXProcedimento)

                ElseIf objProduto.IUD = "U" Then

                    If Not objOPXProcedimento.Salvar() Then
                        MsgBox(Me.Page, HttpContext.Current.Session("ssMessage").ToString())
                        Exit Sub
                    End If
                End If

                SessaoSalvaProduto()

                LimparProcedimento()
            Else
                MsgBox(Me.Page, "Usuário sem permissão para gravar registro.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAtivarProcedimento_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkAtivarProcedimento.Click
        Try

            If Funcoes.VerificaPermissao("ProdutoXProcedimento", "ALTERAR") Then

                If ddlProcedimento.SelectedIndex = 0 Then
                    MsgBox(Me.Page, "Procedimento do Produto não foi selecionado.")
                    Exit Sub
                End If

                SessaoRecuperaProduto()

                Dim objOPXProcedimento = New ProdutoXProcedimento()
                objOPXProcedimento.IUD = "U"
                objOPXProcedimento.CodigoProduto = txtCodigoProduto.Text
                objOPXProcedimento.CodigoProcedimento = ddlProcedimento.SelectedValue

                If Not objOPXProcedimento.Salvar() Then
                    MsgBox(Me.Page, HttpContext.Current.Session("ssMessage").ToString())
                    Exit Sub
                End If

                SessaoSalvaProduto()

                LimparProcedimento()

            Else
                MsgBox(Me.Page, "Usuário sem permissão para alterar registro.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try

    End Sub

    Protected Sub lnkExcluirProcedimento_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkExcluirProcedimento.Click
        Try
            If Funcoes.VerificaPermissao("ProdutoXProcedimento", "EXCLUIR") Then

                If ddlProcedimento.SelectedIndex = 0 Then
                    MsgBox(Me.Page, "Procedimento do Produto não foi selecionado.")
                    Exit Sub
                End If

                SessaoRecuperaProduto()

                Dim objOPXProcedimento = New ProdutoXProcedimento()
                objOPXProcedimento.IUD = "D"
                objOPXProcedimento.CodigoProduto = txtCodigoProduto.Text
                objOPXProcedimento.CodigoProcedimento = ddlProcedimento.SelectedValue

                If Not objOPXProcedimento.Salvar() Then
                    MsgBox(Me.Page, HttpContext.Current.Session("ssMessage").ToString())
                    Exit Sub
                End If

                SessaoSalvaProduto()

                LimparProcedimento()

            Else
                MsgBox(Me.Page, "Usuário sem permissão para alterar registro.", eTitulo.Info)
            End If

        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkLimparProcedimento_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkLimparProcedimento.Click
        LimparProcedimento()
    End Sub

    Protected Sub lnkSelecionarProcedimento_Click(ByVal sender As Object, ByVal e As EventArgs)
        Try
            Dim lnk As LinkButton = CType(sender, LinkButton)
            Dim row As GridViewRow = CType(lnk.NamingContainer, GridViewRow)

            ddlProcedimento.SelectedValue = row.Cells(1).Text
            ddlProcedimento.Enabled = False

            lnkExcluirProcedimento.Parent.Visible = True
            lnkAtivarProcedimento.Parent.Visible = True
            lnkNovoProcedimento.Parent.Visible = False

        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub AtualizarGridEPI()

        SessaoRecuperaProduto()

        Dim listProdutoXEPI = New ListProdutoXEPI(txtCodigoProduto.Text)

        objProduto.ProdutoXEPI = listProdutoXEPI

        gridEPIProduto.DataSource = listProdutoXEPI
        gridEPIProduto.DataBind()

        Dim i As Integer = 0
        While i < gridEPIProduto.Rows.Count
            If gridEPIProduto.Rows(i).Cells(3).Text.ToUpper() = "TRUE" Then
                gridEPIProduto.Rows(i).Cells(3).Text = "SIM"
            Else
                gridEPIProduto.Rows(i).Cells(3).Text = "NÃO"
            End If

            i += 1
        End While

        SessaoSalvaProduto()

    End Sub

    Private Sub LimparEPI()

        ddlEPI.Enabled = True

        ddlEPI.SelectedIndex = 0

        lnkExcluirEPI.Parent.Visible = False
        lnkAtivarEPI.Parent.Visible = False
        lnkNovoEPI.Parent.Visible = True

        gridEPIProduto.DataSource = Nothing
        gridEPIProduto.DataBind()

        If Not String.IsNullOrWhiteSpace(txtCodigoProduto.Text) Then
            AtualizarGridEPI()
        End If
    End Sub

    Protected Sub lnkNovoEPI_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkNovoEPI.Click
        Try
            If Funcoes.VerificaPermissao("ProdutoXEPI", "GRAVAR") Then

                If ddlEPI.SelectedIndex = 0 Then
                    MsgBox(Me.Page, "EPI do Produto não foi selecionado.")
                    Exit Sub
                End If

                SessaoRecuperaProduto()

                Dim objPXEPI = New ProdutoXEPI()
                objPXEPI.IUD = "I"
                objPXEPI.CodigoProduto = txtCodigoProduto.Text
                objPXEPI.CodigoEPI = ddlEPI.SelectedValue
                objPXEPI.Ativo = True

                If objProduto.IUD = "I" Then

                    objProduto.ProdutoXEPI.Add(objPXEPI)

                ElseIf objProduto.IUD = "U" Then

                    If Not objPXEPI.Salvar() Then
                        MsgBox(Me.Page, HttpContext.Current.Session("ssMessage").ToString())
                        Exit Sub
                    End If

                End If

                SessaoSalvaProduto()

                LimparEPI()
            Else
                MsgBox(Me.Page, "Usuário sem permissão para gravar registro.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAtivarEPI_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkAtivarEPI.Click
        Try
            If Funcoes.VerificaPermissao("ProdutoXEPI", "ALTERAR") Then

                If ddlEPI.SelectedIndex = 0 Then
                    MsgBox(Me.Page, "EPI do Produto não foi selecionado.")
                    Exit Sub
                End If

                SessaoRecuperaProduto()

                Dim objPXE = New ProdutoXEPI()
                objPXE.IUD = "U"
                objPXE.CodigoProduto = txtCodigoProduto.Text
                objPXE.CodigoEPI = ddlEPI.SelectedValue

                If Not objPXE.Salvar() Then
                    MsgBox(Me.Page, HttpContext.Current.Session("ssMessage").ToString())
                    Exit Sub
                End If

                SessaoSalvaProduto()

                LimparEPI()
            Else
                MsgBox(Me.Page, "Usuário sem permissão para alterar registro.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkExcluirEPI_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkExcluirEPI.Click
        Try
            If Funcoes.VerificaPermissao("ProdutoXEPI", "EXCLUIR") Then

                SessaoRecuperaProduto()

                Dim objPXE = New ProdutoXEPI()
                objPXE.IUD = "D"
                objPXE.CodigoProduto = txtCodigoProduto.Text
                objPXE.CodigoEPI = ddlEPI.SelectedValue

                If Not objPXE.Salvar() Then
                    MsgBox(Me.Page, HttpContext.Current.Session("ssMessage").ToString())
                    Exit Sub
                End If

                SessaoSalvaProduto()

                LimparEPI()
            Else
                MsgBox(Me.Page, "Usuário sem permissão para desativar registro.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkLimparEPI_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkLimparEPI.Click
        LimparEPI()
    End Sub

    Protected Sub lnkSelecionarEPI_Click(ByVal sender As Object, ByVal e As EventArgs)
        Try
            Dim lnk As LinkButton = CType(sender, LinkButton)
            Dim row As GridViewRow = CType(lnk.NamingContainer, GridViewRow)

            ddlEPI.SelectedValue = row.Cells(1).Text
            ddlEPI.Enabled = False

            lnkExcluirEPI.Parent.Visible = True
            lnkAtivarEPI.Parent.Visible = True
            lnkNovoEPI.Parent.Visible = False

        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub AtualizarGridProdxCons()
        SessaoRecuperaProduto()

        Dim listProdxCons As ListProdutoXCentroDeCusto
        If String.IsNullOrWhiteSpace(txtCodigoProduto.Text) Then
            listProdxCons = Nothing
        Else
            listProdxCons = New ListProdutoXCentroDeCusto(txtCodigoProduto.Text)
        End If

        gridProdxCons.DataSource = listProdxCons
        gridProdxCons.DataBind()
    End Sub

    Private Sub LimparProdxCons()
        ddlPlanoContasProdxCons.Enabled = True
        ddlPlanoContasProdxCons.SelectedIndex = 0
        ddlCentroCustoProdxCons.Enabled = True
        ddlCentroCustoProdxCons.SelectedIndex = 0

        lnkNovoProdxCons.Parent.Visible = True
        lnkExcluirProdxCons.Parent.Visible = True

        gridEPIProduto.DataSource = Nothing
        gridEPIProduto.DataBind()

        AtualizarGridProdxCons()
    End Sub

    Protected Sub lnkNovoProdxCons_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkNovoProdxCons.Click
        Try
            If Funcoes.VerificaPermissao("ProdutoXCentrosDeCustos", "GRAVAR") Then

                If ddlPlanoContasProdxCons.SelectedIndex = 0 Then
                    MsgBox(Me.Page, "Plano de conta não foi selecionado.")
                    Exit Sub
                End If

                If ddlCentroCustoProdxCons.SelectedIndex = 0 Then
                    MsgBox(Me.Page, "Centro de Custo do Produto não foi selecionado.")
                    Exit Sub
                End If

                SessaoRecuperaProduto()

                Dim objProdxCons = New ProdutoXCentroDeCusto()
                objProdxCons.IUD = "I"
                objProdxCons.CodigoProduto = txtCodigoProduto.Text
                objProdxCons.CodigoConta = ddlPlanoContasProdxCons.SelectedValue
                objProdxCons.CodigoCusto = ddlCentroCustoProdxCons.SelectedValue

                If Not objProdxCons.Salvar() Then
                    MsgBox(Me.Page, HttpContext.Current.Session("ssMessage").ToString())
                    Exit Sub
                End If

                LimparProdxCons()
            Else
                MsgBox(Me.Page, "Usuário sem permissão para gravar registro.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkExcluirProdxCons_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkExcluirProdxCons.Click
        Try
            If Funcoes.VerificaPermissao("ProdutoXCentrosDeCustos", "EXCLUIR") Then

                SessaoRecuperaProduto()

                Dim objProdxCons = New ProdutoXCentroDeCusto()
                Dim objCentroDeCusto = New CentroDeCusto(ddlCentroCustoProdxCons.SelectedValue)
                objProdxCons.IUD = "D"
                objProdxCons.CodigoProduto = txtCodigoProduto.Text
                objProdxCons.CodigoConta = ddlPlanoContasProdxCons.SelectedValue
                objProdxCons.CodigoCusto = ddlCentroCustoProdxCons.SelectedValue

                If Not objProdxCons.Salvar() Then
                    MsgBox(Me.Page, HttpContext.Current.Session("ssMessage").ToString())
                    Exit Sub
                End If

                LimparProdxCons()
            Else
                MsgBox(Me.Page, "Usuário sem permissão para desativar registro.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkLimparProdxCons_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkLimparProdxCons.Click
        LimparProdxCons()
    End Sub

    Protected Sub lnkSelecionarProdxCons_Click(ByVal sender As Object, ByVal e As EventArgs)
        Try
            Dim lnk As LinkButton = CType(sender, LinkButton)
            Dim row As GridViewRow = CType(lnk.NamingContainer, GridViewRow)

            ddlCentroCustoProdxCons.SelectedValue = row.Cells(3).Text
            ddlCentroCustoProdxCons.Enabled = False

            ddlPlanoContasProdxCons.SelectedValue = row.Cells(4).Text
            ddlPlanoContasProdxCons.Enabled = False

            lnkExcluirProdxCons.Parent.Visible = True
            lnkNovoProdxCons.Parent.Visible = False
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub
#End Region

#Region "Methods"

    Private Sub SessaoSalvaProduto()
        Session("objProduto" & HID.Value) = objProduto
    End Sub

    Private Sub SessaoRecuperaProduto()
        objProduto = CType(Session("objProduto" & HID.Value), [Lib].Negocio.Produto)
    End Sub

    Public Overrides Sub carregar(obj As IBaseEntity)
        Try
            If Session("objProdutoCad" & HID.Value) IsNot Nothing Then
                Dim objProduto As [Lib].Negocio.Produto = obj
                HidProdAgr.Value = objProduto.Codigo
                txtPrdAgrp.Text = objProduto.Codigo.ToString & " - " & objProduto.Nome
                Session.Remove("objProdutoCad" & HID.Value)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub CarregarGrupos()
        ddl.Carregar(ddlGrupoProduto, CarregarDDL.Tabela.GrupoProdutoGeral, "", True)
    End Sub

    Private Sub CarregarUnidadeDeMedida()
        ddl.Carregar(ddlUnidadeDeMedida, CarregarDDL.Tabela.UnidadeDeMedida, "", True)
        ddl.Carregar(ddlUnidadedeComercializacao, CarregarDDL.Tabela.UnidadeDeMedida, "", True)
    End Sub

    Private Sub CarregarSituacao()
        ddl.Carregar(ddlSituacao, CarregarDDL.Tabela.Situacao, "Situacao_Id in (1,4,5)", True)
    End Sub

    Private Sub CarregarEstadoFisico()
        ddl.Carregar(ddlEstadoFisico, CarregarDDL.Tabela.EstadoFisicoIA, "", True)
    End Sub

    Private Sub CarregarEtapas()
        ddl.Carregar(ddlEtapas, CarregarDDL.Tabela.Etapas, "", True)
    End Sub

    Private Sub CarregarEmbalagem()
        ddl.Carregar(ddlEmbalagem, CarregarDDL.Tabela.Embalagem, "", True)
        ddl.Carregar(ddlPXEmbalagem, CarregarDDL.Tabela.Embalagem, "", True)
    End Sub

    Private Sub CarregarTipoDaEmbagem()
        ddl.Carregar(ddlMaterialEmbalagem, CarregarDDL.Tabela.TipoDeEmbalagem, "", True)
    End Sub

    Private Sub CarregarPesoQuantidade()
        ddlPesoQuantidade.Items.Add(New ListItem("Peso", "P"))
        ddlPesoQuantidade.Items.Add(New ListItem("Quantidade", "Q"))
        Funcoes.InserirLinhaEmBranco(ddlPesoQuantidade)
    End Sub

    Private Sub CarregarControleDeQualidade()
        ddl.Carregar(ddlQualidade, CarregarDDL.Tabela.ControleDeQualidade, "", True)
    End Sub

    Private Sub CarregarMarcas()
        ddl.Carregar(ddlMarca, CarregarDDL.Tabela.Marca, "", True)
    End Sub

    Private Sub CarregarSeguimentos()
        ddl.Carregar(ddlSeguimentos, CarregarDDL.Tabela.Seguimentos, "", True)
    End Sub

    Private Sub CarregarCarteiraDeCompras()
        ddl.Carregar(ddlCarteiraDeCompras, CarregarDDL.Tabela.CarteiraFinanceiraConta, "Classificacao = 'P'", True)
    End Sub

    Private Sub CarregarCarteiraDeVendas()
        ddl.Carregar(ddlCarteiraDeVendas, CarregarDDL.Tabela.CarteiraFinanceiraConta, "Classificacao = 'R'", True)
    End Sub

    Private Sub CarregarEspecificacoes()
        ddl.Carregar(ddlEspecificacao, CarregarDDL.Tabela.EspecificacaoDoProduto, "Ativo = 1", True)
    End Sub

    Private Sub CarregarPlanoConta()
        ddl.Carregar(ddlPlanoContasProdxCons, CarregarDDL.Tabela.PlanoDeContas, " (Cliente = 'S' and len(conta_id) = 7) or (len(conta_id) = 9 and left(conta_id, 1) in (1,2,3,4)) order by len(conta_id), conta_id", True)
    End Sub

    Private Sub CarregarCentroDeCusto()
        ddl.Carregar(ddlCentroCustoProdxCons, CarregarDDL.Tabela.CentroDeCustoDescricao, "LEN(CentroDeCusto_Id) = 5")
    End Sub

    Private Sub CarregarTipoDoItem()
        ddlTipoDoItem.Items.Add(New ListItem("00 - Mercadorias para Revenda", 0))
        ddlTipoDoItem.Items.Add(New ListItem("01 - Matéria Prima", 1))
        ddlTipoDoItem.Items.Add(New ListItem("02 - Embalagem", 2))
        ddlTipoDoItem.Items.Add(New ListItem("03 - Produto em Processo", 3))
        ddlTipoDoItem.Items.Add(New ListItem("04 - Produto Acabado", 4))
        ddlTipoDoItem.Items.Add(New ListItem("05 - SubProduto", 5))
        ddlTipoDoItem.Items.Add(New ListItem("06 - Produto Intermediário", 6))
        ddlTipoDoItem.Items.Add(New ListItem("07 - Material de Uso e Consumo", 7))
        ddlTipoDoItem.Items.Add(New ListItem("08 - Ativo Imobilizado", 8))
        ddlTipoDoItem.Items.Add(New ListItem("09 - Serviços", 9))
        ddlTipoDoItem.Items.Add(New ListItem("10 - Outros Insumos", 10))
        ddlTipoDoItem.Items.Add(New ListItem("99 - Outras", 99))
        Funcoes.InserirLinhaEmBranco(ddlTipoDoItem)
    End Sub

    Private Sub CarregarCodigoDoGenero()
        Dim strSQL As String = "SELECT Codigo_Id AS Codigo, Descricao " & vbCrLf &
                               "  FROM GeneroDoProduto " & vbCrLf
        strSQL &= "ORDER BY Codigo_Id"

        For Each Dr As DataRow In Banco.ConsultaDataSet(strSQL, "GeneroDoProduto").Tables(0).Rows
            ddlCodigoDoGenero.Items.Add(New ListItem(Funcoes.AlinharDireita(Dr("Codigo"), 2, "0") & " - " & Dr("Descricao"), Dr("Codigo")))
        Next
        If ddlCodigoDoGenero IsNot Nothing Then
            For Each li As ListItem In ddlCodigoDoGenero.Items
                li.Attributes("title") = li.Text
            Next
        End If

        Funcoes.InserirLinhaEmBranco(ddlCodigoDoGenero)
    End Sub

    Private Sub CarregarSubCodigoDoGenero()
        Dim strSQL As String = "SELECT SubCodigo_Id AS Codigo, Descricao " & vbCrLf &
                               "  FROM GeneroDoProdutoXSub " & vbCrLf &
                               " WHERE Codigo_Id = " & ddlCodigoDoGenero.SelectedValue & vbCrLf
        strSQL &= "ORDER BY Codigo_Id"

        ddlSubCodigoDoGenero.Items.Clear()
        For Each Dr As DataRow In Banco.ConsultaDataSet(strSQL, "GeneroDoProdutoXSub").Tables(0).Rows
            ddlSubCodigoDoGenero.Items.Add(New ListItem(Funcoes.AlinharDireita(Dr("Codigo"), 2, "0") & " - " & Dr("Descricao"), Dr("Codigo")))
        Next

        If ddlSubCodigoDoGenero IsNot Nothing Then
            For Each li As ListItem In ddlSubCodigoDoGenero.Items
                li.Attributes("title") = li.Text
            Next
        End If

        Funcoes.InserirLinhaEmBranco(ddlSubCodigoDoGenero)
    End Sub



    Private Sub LimparCampos(ByVal Tudo As Boolean)
        HID.Value = Guid.NewGuid().ToString
        ucConsultaProduto.SetarHID(HID.Value)

        lblConsolidaProduto.Parent.Visible = False

        ddlGrupoProduto.Enabled = True
        ddlGrupoProduto.SelectedIndex = 0
        txtCodigoProduto.Text = ""
        txtCodigoProduto.Enabled = True

        txtCodigoProdutoTerceiro.Text = ""

        imgSequencia.Enabled = True
        ddlUnidadeDeMedida.SelectedIndex = 0


        ddlUnidadeDeMedida.Enabled = True

        ddlSituacao.SelectedValue = 1
        txtCodigoIndea.Text = ""
        txtCodigoIndea.Enabled = False
        txtNome.Text = ""
        txtDescricao.Text = ""
        txtInfaDProd.Text = ""
        ddlEtapas.SelectedIndex = 0

        ddlEmbalagem.SelectedIndex = 0
        txtEstoqueMinimo.Text = "0,0000"

        'txtQuantidadeNaCaixa.Text = ""

        rdANao.Checked = False
        rdASim.Checked = False
        chkControlarPesagem.Checked = False
        chkControlarRomaneio.Checked = False
        chkControlarRomaneio.Enabled = False
        txtDescricaoMapa.Text = ""
        txtNCM.Text = ""
        'txtCnae.Text = String.Empty
        ddlCnae.SelectedValue = String.Empty
        ddlPesoQuantidade.SelectedIndex = 0
        ddlQualidade.SelectedIndex = 0
        txtIPI.Text = ""
        'rdINao.Checked = True
        'chkIntegral.Checked = False
        'chkPresumido.Checked = False
        chkControlarLote.Checked = False
        chkControlarEmbalagem.Checked = False
        chkCustoIndireto.Checked = False
        TabEmbalagens.Visible = False

        chkFitossanitario.Checked = False
        chkControlarEstoque.Checked = False
        chkPrecoDePauta.Checked = False
        chkPecasMeios.Checked = False
        chkDecimais.Checked = False
        chkAlmoxarifado.Checked = False
        chkPrecoDoProduto.Checked = False
        chkNumeroDoLote.Checked = False
        chkAutorizacaoDeRetirada.Checked = False
        ddlMarca.SelectedIndex = 0
        ddlCarteiraDeCompras.SelectedIndex = 0
        ddlCarteiraDeVendas.SelectedIndex = 0
        ddlTipoDoItem.SelectedIndex = 0
        txtCodigoDoServico.Text = ""
        txtCodigoEX.Text = ""
        ddlEstadoFisico.SelectedIndex = 0
        ddlCodigoDoGenero.SelectedIndex = 0
        ddlSubCodigoDoGenero.Items.Clear()
        ddlPXEmbalagem.SelectedIndex = 0
        ddlMaterialEmbalagem.SelectedIndex = 0
        txtQuantidadeDaEmbalagem.Text = ""
        txtPesoBruto.Text = ""
        txtPesoLiquido.Text = ""
        txtRegMinAgr.Text = ""
        ddlPXEmbalagem.Enabled = False
        ddlMaterialEmbalagem.Enabled = False
        txtQuantidadeDaEmbalagem.Enabled = False
        txtPesoBruto.Enabled = False
        txtPesoLiquido.Enabled = False
        lnkAdicionarEmbalagem.Parent.Parent.Visible = False
        gridEmbalagem.DataSource = Nothing
        gridEmbalagem.DataBind()
        gridProdutoAgrupado.DataSource = Nothing
        gridProdutoAgrupado.DataBind()
        lnkNovo.Parent.Visible = True
        lnkAtualizar.Parent.Visible = False
        lnkExcluir.Parent.Visible = False

        ddlUnidadedeComercializacao.SelectedIndex = 0
        txtPesoDoProduto.Text = 0
        txtPesoDaEmbalagem.Text = 0
        txtPrdAgrp.Text = String.Empty
        HidProdAgr.Value = String.Empty

        gridUnidadeComercializacao.DataSource = Nothing
        gridUnidadeComercializacao.DataBind()

        txtGTIN8.Text = ""
        imgGTIN8.ImageUrl = ""

        txtGTIN12.Text = ""
        imgGTIN12.ImageUrl = ""

        txtGTIN13.Text = ""
        imgGTIN13.ImageUrl = ""

        txtGTIN14.Text = ""
        imgGTIN14.ImageUrl = ""

        LimparEPrd()
        LimparEPI()
        LimparProcedimento()

        Session.Remove("objProduto" & HID.Value)

        If Tudo Then
            objProduto = New [Lib].Negocio.Produto()
            objProduto.IUD = "I"

            objProduto.UsuarioInclusao = Session("ssNomeUsuario")

            ddlUsuarios.Items.Clear()
            ddlUsuarios.Items.Add("Inc.- " & objProduto.UsuarioInclusao)

            SessaoSalvaProduto()
        End If
    End Sub

    Private Sub Limpar_PxE()
        ddlPXEmbalagem.SelectedIndex = 0
        ddlMaterialEmbalagem.SelectedIndex = 0
        txtQuantidadeDaEmbalagem.Text = ""
        txtPesoBruto.Text = ""
        txtPesoLiquido.Text = ""
        chkPesoVariavel.Checked = False
        SessaoRecuperaProduto()
        If objProduto.ProdutoXEmbalagens.Count > 0 Then
            gridEmbalagem.DataSource = objProduto.ProdutoXEmbalagens.ToArray()
        Else
            gridEmbalagem.DataSource = Nothing
        End If
        gridEmbalagem.DataBind()
    End Sub

    Private Function ValidarCampos() As Boolean


        Dim listaGTIN As List(Of String) = New List(Of String) From {txtGTIN8.Text, txtGTIN12.Text, txtGTIN13.Text, txtGTIN14.Text}
        listaGTIN = listaGTIN.Where(Function(gtin) Not String.IsNullOrEmpty(gtin)).ToList()

        If objProduto Is Nothing Then SessaoRecuperaProduto()

        If ddlSituacao.SelectedIndex = 0 Then
            MsgBox(Me.Page, "Situação não foi selecionada.")
            Return False
        ElseIf ddlGrupoProduto.SelectedIndex = 0 Then
            MsgBox(Me.Page, "Grupo do Produto não foi selecionado.")
            Return False
        ElseIf txtCodigoProduto.Text.Length = 0 Then
            MsgBox(Me.Page, "Código do Produto não foi informado.")
            Return False
        ElseIf txtNome.Text.Length = 0 Then
            MsgBox(Me.Page, "Nome do Produto não foi informado.")
            Return False
        ElseIf ddlEtapas.SelectedIndex = 0 Then
            MsgBox(Me.Page, "Etapa do Produto não foi selecionada.")
            Return False
        ElseIf rdANao.Checked = False And rdASim.Checked = False Then
            MsgBox(Me.Page, "Selecione Agrupar Sim ou Não.")
            Return False
        ElseIf chkFitossanitario.Checked AndAlso txtCodigoIndea.Text.Length = 0 Then
            MsgBox(Me.Page, "Caso marque Fitossanitário Código do Indea é obrigatório.")
            Return False
        ElseIf txtNCM.Text.Length = 0 Then
            MsgBox(Me.Page, "Código do NCM não foi informado.")
            txtNCM.Focus()
            Return False
        ElseIf txtNCM.Text.Length > 8 Then
            MsgBox(Me.Page, "Código do NCM deve ter no máximo 8 caracteres.")
            txtNCM.Focus()
            Return False
        ElseIf ddlCarteiraDeCompras.SelectedIndex = 0 Then
            MsgBox(Me.Page, "Carteira de Compra não foi selecionada.")
            Return False
        ElseIf ddlCarteiraDeVendas.SelectedIndex = 0 Then
            ddlCarteiraDeVendas.Focus()
            MsgBox(Me.Page, "Carteira de Venda não foi selecionada.")
            Return False
        ElseIf ddlTipoDoItem.SelectedIndex = 0 Then
            MsgBox(Me.Page, "Tipo do Item não foi selecionado.")
            Return False
        ElseIf ddlCodigoDoGenero.SelectedIndex = 0 Then
            MsgBox(Me.Page, "Gênero não foi selecionado.")
            Return False
        ElseIf ddlSubCodigoDoGenero.SelectedIndex = 0 Then
            MsgBox(Me.Page, "Sub do Gênero não foi selecionado.")
            Return False
        ElseIf ddlEmbalagem.SelectedIndex = 0 Then
            MsgBox(Me.Page, "Embalagem Padrão não foi selecionada.")
            Return False
        ElseIf chkControlarEmbalagem.Checked AndAlso gridEmbalagem.Rows.Count = 0 Then
            MsgBox(Me.Page, "Informe as embalagens a serem controladas.")
            Return False
        ElseIf ddlPesoQuantidade.SelectedIndex = 0 Then
            MsgBox(Me.Page, "Selecione Peso ou Quantidade para o Produto.")
            Return False
        ElseIf ddlUnidadeDeMedida.SelectedIndex = 0 Then
            MsgBox(Me.Page, "Unidade de Medida não foi selecionada.")
            Return False
        ElseIf ddlEstadoFisico.SelectedValue.Length = 0 Then
            MsgBox(Me.Page, "Estado Físico não foi selecionado.")
            Return False
        ElseIf ValidaUnidadeDeComercializacao(True) = False Then
            Return False
        ElseIf listaGTIN.Count > 2 Then
            MsgBox(Me.Page, "Só pode haver no maximo 2 GTIN informado!")
            Return False
        ElseIf txtGTIN8.Text.Length > 0 AndAlso txtGTIN8.Text.Length <> 8 Then
            MsgBox(Me.Page, "O código GTIN8 está invalido!")
            Return False
        ElseIf txtGTIN12.Text.Length > 0 AndAlso txtGTIN12.Text.Length <> 12 Then
            MsgBox(Me.Page, "O código GTIN12 está invalido!")
            Return False
        ElseIf txtGTIN13.Text.Length > 0 AndAlso txtGTIN13.Text.Length <> 13 Then
            Dim t = txtGTIN13.Text.Length
            MsgBox(Me.Page, "O código GTIN13 está invalido!")
            Return False
        ElseIf txtGTIN14.Text.Length > 0 AndAlso txtGTIN14.Text.Length <> 14 Then
            MsgBox(Me.Page, "O código GTIN14 está invalido!")
            Return False
        ElseIf txtGTIN8.Text.Length > 0 AndAlso ValidarGTIN(txtGTIN8.Text) = False Then
            MsgBox(Me.Page, "O código GTIN8 está invalido!")
            Return False
        ElseIf txtGTIN12.Text.Length > 0 AndAlso ValidarGTIN(txtGTIN12.Text) = False Then
            MsgBox(Me.Page, "O código GTIN12 está invalido!")
            Return False
        ElseIf txtGTIN13.Text.Length > 0 AndAlso ValidarGTIN(txtGTIN13.Text) = False Then
            MsgBox(Me.Page, "O código GTIN13 está invalido!")
            Return False
        ElseIf txtGTIN14.Text.Length > 0 AndAlso ValidarGTIN(txtGTIN14.Text) = False Then
            MsgBox(Me.Page, "O código GTIN14 está invalido!")
            Return False
        ElseIf chkDashboard.Checked AndAlso ddlSeguimentos.SelectedIndex = 0 Then
            MsgBox(Me.Page, "Seguimento não foi selecionado!")
            Return False
        Else
            Return True
        End If
    End Function

    Function ValidarGTIN(gtin As String) As Boolean
        ' Remove espaços extras
        gtin = gtin.Trim()

        ' Verifica se contém apenas números
        If Not IsNumeric(gtin) Then Return False

        ' Valida os tamanhos permitidos
        If gtin.Length <> 8 AndAlso gtin.Length <> 12 AndAlso gtin.Length <> 13 AndAlso gtin.Length <> 14 Then Return False

        ' Cálculo do dígito verificador
        Dim soma As Integer = 0
        Dim multiplicador As Integer
        For i As Integer = 0 To gtin.Length - 2
            multiplicador = If((gtin.Length - i) Mod 2 = 0, 3, 1)
            soma += CInt(gtin(i).ToString()) * multiplicador
        Next

        Dim digitoVerificador As Integer = (10 - (soma Mod 10)) Mod 10
        Return digitoVerificador = CInt(gtin.Last().ToString())
    End Function

    Private Function ValidaUnidadeDeComercializacao(ByVal bAtualizarProduto As Boolean) As Boolean

        Dim i As Integer = -1
        i = objProduto.UnidadesDeComercializacao.FindIndex(Function(s) s.CodigoUnidade = "SCS")

        Dim unidadePadrao As String = ddlUnidadeDeMedida.SelectedItem.Text.Split("-")(0).Trim.Replace(".", "")

        If Not IsNumeric(txtPesoDoProduto.Text) Then
            MsgBox(Me.Page, "Informe o peso do produto valido para Unidade de comercialização.")
            Return False
        ElseIf objProduto.UnidadesDeComercializacao.Where(Function(s) s.CodigoUnidade <> "SCS" And s.FatorConversao = 0).Count > 0 And chkControlarLote.Checked Then
            MsgBox(Me.Page, "Foi Informado uma unidade de comercializacao sem o peso do produto, o peso do produto só pode ser zerado se o produto for controlado por lote.")
            Return False
        ElseIf objProduto.UnidadesDeComercializacao.Where(Function(s) s.CodigoUnidade <> "SCS" And s.FatorConversao = 0).Count > 0 Then
            MsgBox(Me.Page, "O peso do produto não pode ser ZERO.")
            Return False
        ElseIf bAtualizarProduto AndAlso objProduto.UnidadesDeComercializacao.Where(Function(s) s.Ativo = True).Count = 0 Then
            MsgBox(Me.Page, "Ative ao menos 1(uma) unidade de comercialização.")
            Return False
        ElseIf bAtualizarProduto AndAlso gridUnidadeComercializacao.Rows.Count > 0 AndAlso objProduto.UnidadesDeComercializacao.Where(Function(u) u.CodigoUnidade = unidadePadrao And u.IUD <> "D").Count = 0 Then
            MsgBox(Me.Page, "Informe ao menos, uma unidade de comercialização com a unidade padrão")
            Return False
        ElseIf i >= 0 AndAlso objProduto.UnidadesDeComercializacao(i).FatorConversao = 0 AndAlso Not objProduto.ControlarLote Then
            MsgBox(Me.Page, "O Fator de Conversao na Unidade de Comercialização SCS só pode ser ZERO, se o produto for controlado por Lote.")
            Return False
        End If

        Return True

    End Function

    Private Sub GravarProduto(ByVal Tipo As String)
        Try
            If ValidarCampos() Then
                SessaoRecuperaProduto()
                objProduto.IUD = Tipo
                objProduto.CodigoGrupo = ddlGrupoProduto.SelectedValue
                objProduto.Codigo = txtCodigoProduto.Text
                objProduto.Unidade = ddlUnidadeDeMedida.SelectedValue
                objProduto.CodigoSituacao = ddlSituacao.SelectedValue
                objProduto.Nome = txtNome.Text
                objProduto.Descricao = txtDescricao.Text
                objProduto.CodigoProdutoTerceiro = txtCodigoProdutoTerceiro.Text

                'Levar informarção do InfaDProd do Produto para NFE caso tenha - Furlan - 21-12-2022
                If txtInfaDProd.Text.Length > 0 Then objProduto.InfaDProd = Funcoes.EliminarCaracteresEspeciaisNF(txtInfaDProd.Text)

                objProduto.Etapa = ddlEtapas.SelectedValue
                objProduto.CodigoEmbalagem = ddlEmbalagem.SelectedValue
                If rdASim.Checked = True Then objProduto.Agrupar = "S"
                If rdANao.Checked = True Then objProduto.Agrupar = "N"
                If chkControlarPesagem.Checked Then
                    objProduto.ControlarPesagem = True
                Else
                    objProduto.ControlarPesagem = False
                End If
                If chkControlarRomaneio.Checked Then
                    objProduto.ControlarRomaneio = True
                Else
                    objProduto.ControlarRomaneio = False
                End If
                objProduto.DescricaoMapa = txtDescricaoMapa.Text
                objProduto.NCM = txtNCM.Text
                objProduto.CodigoCnae = IIf(ddlCnae.SelectedIndex > 0, ddlCnae.SelectedValue, "")
                objProduto.PesoQuantidade = ddlPesoQuantidade.SelectedValue

                objProduto.Qualidade = IIf(ddlQualidade.SelectedIndex = 0, 0, ddlQualidade.SelectedValue)

                atulizarGtinProduto()

                'If txtIPI.Text.Length = 0 Then
                '    objProduto.IPI = 0
                'Else
                '    objProduto.IPI = CDbl(txtIPI.Text)
                'End If

                'If rdISim.Checked = True Then objProduto.IPITributado = "S"
                'If rdINao.Checked = True Then objProduto.IPITributado = "N"

                'If chkIntegral.Checked Then
                '    objProduto.PisCofinsIntegral = "S"
                'Else
                '    objProduto.PisCofinsIntegral = "N"
                'End If
                'If chkPresumido.Checked Then
                '    objProduto.PisCofinsPresumido = "S"
                'Else
                '    objProduto.PisCofinsPresumido = "N"
                'End If

                objProduto.ControlarLote = chkControlarLote.Checked
                objProduto.ControlarEmbalagem = chkControlarEmbalagem.Checked
                objProduto.Fitossanitario = chkFitossanitario.Checked
                objProduto.ControlarEstoque = chkControlarEstoque.Checked
                objProduto.ControlarPrecoDePauta = chkPrecoDePauta.Checked
                objProduto.ControlarPecas = chkPecasMeios.Checked
                objProduto.ControlarDecimais = chkDecimais.Checked
                objProduto.Almoxarifado = chkAlmoxarifado.Checked
                objProduto.PrecoDoProduto = chkPrecoDoProduto.Checked
                objProduto.ControlarNumeroDoLote = chkNumeroDoLote.Checked
                objProduto.AutorizacaoDeRetirada = chkAutorizacaoDeRetirada.Checked
                objProduto.CustoIndireto = chkCustoIndireto.Checked

                If objProduto.ControlarEstoque Then
                    objProduto.EstoqueMinimo = CDbl(txtEstoqueMinimo.Text)
                Else
                    objProduto.EstoqueMinimo = 0
                End If

                If chkFitossanitario.Checked Then objProduto.ProdutoIndea = txtCodigoIndea.Text
                If ddlMarca.SelectedIndex > 0 Then objProduto.CodigoDaMarca = ddlMarca.SelectedValue

                objProduto.CodigoCarteiraCompra = ddlCarteiraDeCompras.SelectedValue
                objProduto.CodigoCarteiraVenda = ddlCarteiraDeVendas.SelectedValue
                objProduto.TipoItem = ddlTipoDoItem.SelectedValue

                If txtCodigoDoServico.Text.Length > 0 Then objProduto.CodigoServico = txtCodigoDoServico.Text
                If txtCodigoEX.Text.Length > 0 Then objProduto.CodigoEX = txtCodigoEX.Text

                objProduto.CodigoGenero = ddlCodigoDoGenero.SelectedValue
                objProduto.SubCodigoGenero = ddlSubCodigoDoGenero.SelectedValue

                objProduto.CodigoEstadoFisico = ddlEstadoFisico.SelectedValue

                objProduto.RegistroMinisterioAgricultura = txtRegMinAgr.Text

                objProduto.Dashboard = chkDashboard.Checked
                If ddlSeguimentos.SelectedIndex > 0 Then
                    objProduto.CodigoSeguimento = ddlSeguimentos.SelectedValue
                Else
                    objProduto.CodigoSeguimento = 0
                End If

                If objProduto.Salvar Then
                    MsgBox(Me.Page, "Produto " & IIf(Tipo = "I", "Incluso", "Alterado") & " com Sucesso.", eTitulo.Sucess)
                    LimparCampos(True)
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Function getParam() As String
        Dim param As String = ""
        If Not String.IsNullOrWhiteSpace(txtCodigoProduto.Text) Then
            param &= "Produto: " & txtCodigoProduto.Text & " - "
        End If
        If Not String.IsNullOrWhiteSpace(txtNome.Text) Then
            param &= "Nome: " & txtNome.Text & " - "
        End If
        If Not String.IsNullOrWhiteSpace(ddlGrupoProduto.SelectedValue) Then
            param &= "Grupo: " & ddlGrupoProduto.SelectedValue & " - "
        End If
        If Not String.IsNullOrWhiteSpace(txtDescricao.Text) Then
            param &= "Descrição: " & txtDescricao.Text & vbCrLf
        End If
        If Not String.IsNullOrWhiteSpace(ddlSituacao.SelectedValue) Then
            param &= "Situação: " & ddlSituacao.SelectedValue & " - "
        End If
        If Not String.IsNullOrWhiteSpace(ddlEtapas.SelectedValue) Then
            param &= "Etapas: " & ddlEtapas.SelectedValue & " - "
        End If
        If Not String.IsNullOrWhiteSpace(txtNCM.Text) Then
            param &= "Código NCM: " & txtNCM.Text & "."
        End If

        If ddlCnae.SelectedIndex > 0 Then
            param &= "CNAE: " & ddlCnae.SelectedValue & "."
        End If

        Return param
    End Function

#End Region

    Protected Sub ConsultaPamCard(ByVal sender As Object, e As EventArgs) Handles lnkConsultarPamCard.Click
        Dim pamcard As PamcardNew = New PamcardNew()
        pamcard.ConsultarFavorecido("63358210000176", "48384601000171")
    End Sub

    Protected Sub txtCodigoProduto_TextChanged(sender As Object, e As EventArgs) Handles txtCodigoProduto.TextChanged
        SessaoRecuperaProduto()
        objProduto.Codigo = txtCodigoProduto.Text
        objProduto.UnidadesDeComercializacao = Nothing
        gridUnidadeComercializacao.DataSource = Nothing
        gridUnidadeComercializacao.DataBind()
        SessaoSalvaProduto()
    End Sub

    Protected Sub chkPesoVariavel_CheckedChanged(sender As Object, e As EventArgs) Handles chkPesoVariavel.CheckedChanged
        pnlPesoEmbalagem.Visible = Not chkPesoVariavel.Checked
        txtPesoBruto.Text = "1,0000"
        txtPesoLiquido.Text = "1,0000"
    End Sub
End Class