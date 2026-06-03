Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis
Public Class OrdemDeProducaoInterna
    Inherits BasePage

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.ProducaoEstoque)
        If Not IsPostBack And IsConnect Then
            'If Funcoes.VerificaPermissao("OrdemDeProducaoInterna", "ACESSAR") Then
            '    BuncarUnidadeDeNegocio()

            '    ddl.Carregar(cmbOperacao, CarregarDDL.Tabela.Operacao, "", True)
            '    cmbOperacao.SelectedValue = 40 'PRODUÇÃO
            '    ddl.Carregar(cmbSubOperacao, CarregarDDL.Tabela.OperacaoSubOperacao, "So.Operacao_Id = " & cmbOperacao.SelectedValue, True)

            '    Limpar(True)

            '    ddl.Carregar(ddlGrupoProdutoInterno, CarregarDDL.Tabela.GrupoProdutoXConsumo)

            'Else
            '    MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Producao.aspx")
            '    Exit Sub
            'End If

            MsgBox(Me.Page, "Programa desabilitado, dúvidas entre em contato com o Suporte.", "~/Producao.aspx")
            Exit Sub

        End If
    End Sub

    Public Overrides Sub Carregar(obj As IBaseEntity)
        If Session("objProdutoxPRDI" & HID.Value) IsNot Nothing Then
            Dim objProduto As [Lib].Negocio.Produto = CType(obj, [Lib].Negocio.Produto)

            If objProduto.ControlarEstoque Then
                ddlGrupoProdutoInterno.SelectedValue = objProduto.CodigoGrupo
                CarregarProdutoInterno()
                ddlProdutoInterno.SelectedValue = objProduto.Codigo
                Session.Remove("objProdutoxPRDI" & HID.Value)

                CarregarGridProdutoInterno(ddlProdutoInterno.SelectedValue, "")

                lnkNovoI.Parent.Visible = True
            Else
                MsgBox(Me.Page, "Produto selecionado não está marcado para CONTROLAR ESTOQUE.", eTitulo.Info)
            End If
        End If
    End Sub

    Private Sub BuncarUnidadeDeNegocio()
        Try
            ddl.Carregar(ddlUnidadeNegocio, CarregarDDL.Tabela.UnidadeDeNegocio)
            Funcoes.VerificaUnidadeDeNegocio(ddlUnidadeNegocio, ddlEmpresa)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub ddlUnidadeNegocio_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            If ddlUnidadeNegocio.SelectedIndex > 0 Then
                ddl.Carregar(ddlEmpresa, CarregarDDL.Tabela.Empresas, ddlUnidadeNegocio.SelectedValue.ToString)
            Else
                ddlEmpresa.Items.Clear()

                Limpar(True)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub ddlEmpresa_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            'CarregarGridProdutoInterno("", "")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub rdEntradas_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles rdEntradas.CheckedChanged
        cmbSubOperacao.SelectedValue = cmbOperacao.SelectedValue & "-" & 10 'ORDEM DE PRODUÇÃO INTERNA - ENTRADA
    End Sub

    Protected Sub rdSaidas_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles rdSaidas.CheckedChanged
        cmbSubOperacao.SelectedValue = cmbOperacao.SelectedValue & "-" & 60 'ORDEM DE PRODUÇÃO INTERNA - SAÍDA
    End Sub

    Private Sub CarregarGridProdutoInterno(Optional ByVal prd As String = "", Optional ByVal nLote As String = "")
        Try
            If ddlEmpresa.SelectedIndex > 0 Then
                Dim objBanco As New AcessaBanco()

                Dim strSQL As String = "SELECT opi.Produto_Id AS Produto, p.Nome AS NomeProduto, opi.Movimento, opi.Validade, opi.Lote_Id AS Lote, opi.Quantidade, opi.Observacoes, opi.EntradaSaida, opi.Operacao, opi.SubOperacao " & vbCrLf &
                                    "FROM OrdemDeProducaoInterna opi " & vbCrLf &
                                    "		inner join Produtos p " & vbCrLf &
                                    "				on p.Produto_Id = opi.Produto_Id " & vbCrLf &
                                    "Where opi.Empresa_Id = '" & ddlEmpresa.SelectedValue.Split("-")(0) & "'" & vbCrLf &
                                    "  and opi.EndEmpresa_Id = " & ddlEmpresa.SelectedValue.Split("-")(1) & vbCrLf

                If nLote.Length > 0 Then strSQL &= "  and opi.Lote_Id = '" & Trim(txtLoteInterno.Text) & "'" & vbCrLf

                If prd.Length > 0 Then strSQL &= "  and opi.Produto_Id = '" & ddlProdutoInterno.SelectedValue & "'" & vbCrLf

                Dim ds As DataSet = Banco.ConsultaDataSet(strSQL, "OrdemDeProducaoInterna")

                gridProducaoInterno.DataSource = ds
            Else
                gridProducaoInterno.DataSource = Nothing
            End If

            gridProducaoInterno.DataBind()

        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub imgSelecionaPrdInterno_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)
        Dim imgSelecionaPrdInterno As ImageButton = CType(sender, ImageButton)
        Dim row As GridViewRow = CType(imgSelecionaPrdInterno.NamingContainer, GridViewRow)

        txtLoteInterno.Text = Trim(gridProducaoInterno.Rows(row.RowIndex).Cells(5).Text)
        txtDataProducaoInterna.Text = gridProducaoInterno.Rows(row.RowIndex).Cells(3).Text
        txtDataValidadeInterna.Text = gridProducaoInterno.Rows(row.RowIndex).Cells(4).Text
        txtQuantidadeInterna.Text = gridProducaoInterno.Rows(row.RowIndex).Cells(6).Text

        If gridProducaoInterno.Rows(row.RowIndex).Cells(7).Text = "&NBSP;" Then
            txtObsProdutoInterno.Text = ""
        Else
            txtObsProdutoInterno.Text = gridProducaoInterno.Rows(row.RowIndex).Cells(7).Text
        End If

        If gridProducaoInterno.Rows(row.RowIndex).Cells(8).Text = "E" Then rdEntradas.Checked = True

        If gridProducaoInterno.Rows(row.RowIndex).Cells(8).Text = "S" Then rdSaidas.Checked = True

        cmbSubOperacao.SelectedValue = gridProducaoInterno.Rows(row.RowIndex).Cells(9).Text & "-" & gridProducaoInterno.Rows(row.RowIndex).Cells(10).Text

        Dim objProduto As [Lib].Negocio.Produto = New Produto(gridProducaoInterno.Rows(row.RowIndex).Cells(1).Text)

        ddlGrupoProdutoInterno.SelectedValue = objProduto.CodigoGrupo
        CarregarProdutoInterno()
        ddlProdutoInterno.SelectedValue = objProduto.Codigo

        txtLoteInterno.Enabled = False
        ddlGrupoProdutoInterno.Enabled = False
        ddlProdutoInterno.Enabled = False

        lnkAtualizarI.Parent.Visible = True
        lnkNovoI.Parent.Visible = False

    End Sub

    Protected Sub imgRemoverProdutoInterno_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)
        Try
            If Funcoes.VerificaPermissao("OrdemDeProducaoInterna", "EXCLUIR") Then
                Dim imgSelecionaPrdInterno As ImageButton = CType(sender, ImageButton)
                Dim row As GridViewRow = CType(imgSelecionaPrdInterno.NamingContainer, GridViewRow)

                Dim objBanco As New AcessaBanco()
                Dim sqls As New ArrayList

                Dim strSQL As String = "SELECT Lote_Id FROM OrdemDeProducaoXConsumoXLote " & vbCrLf &
                                          "Where Empresa_Id    = '" & ddlEmpresa.SelectedValue.Split("-")(0) & "'" & vbCrLf &
                                          "  and EndEmpresa_Id = " & ddlEmpresa.SelectedValue.Split("-")(1) & vbCrLf &
                                          "  and Produto_Id    = '" & ddlProdutoInterno.SelectedValue & "'" & vbCrLf &
                                          "  and Lote_Id       = '" & Trim(txtLoteInterno.Text) & "'"

                Dim ds As DataSet = Banco.ConsultaDataSet(strSQL, "OrdemDeProducaoXConsumoXLote")

                If ds.Tables(0).Rows.Count > 0 Then
                    MsgBox(Me.Page, "Lote interno não pode ser excluído porque já foi utilizado para consumo.", eTitulo.Info)
                    Exit Sub
                End If

                strSQL = "Delete from OrdemDeProducaoInterna " & vbCrLf &
                            "Where Empresa_Id    = '" & ddlEmpresa.SelectedValue.Split("-")(0) & "'" & vbCrLf &
                            "  and EndEmpresa_Id = " & ddlEmpresa.SelectedValue.Split("-")(1) & vbCrLf &
                            "  and Produto_Id    = '" & gridProducaoInterno.Rows(row.RowIndex).Cells(1).Text & "'" & vbCrLf &
                            "  and Lote_Id       = '" & gridProducaoInterno.Rows(row.RowIndex).Cells(5).Text & "'"

                sqls.Add(strSQL)

                If objBanco.GravaBanco(sqls) Then

                    Dim prd As String = gridProducaoInterno.Rows(row.RowIndex).Cells(1).Text

                    Limpar(False)

                    CarregarGridProdutoInterno(prd, "")

                Else
                    MsgBox(Me.Page, HttpContext.Current.Session("ssMessage").ToString, eTitulo.Erro)
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para excluir registro.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try

    End Sub

    Protected Sub lnkBuscaProdutoInterno_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lnkBuscaProdutoInterno.Click
        Try
            lnkNovoI.Parent.Visible = False

            ucConsultaProduto.Limpar()
            Dim txtNome As TextBox = CType(ucConsultaProduto.FindControlRecursive("txtNome"), TextBox)
            Popup.ConsultaDeProduto(Me.Page, "objProdutoxPRDI" & HID.Value, txtNome.ClientID, True)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub ddlGrupoProdutoInterno_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddlGrupoProdutoInterno.SelectedIndexChanged
        Try
            CarregarProdutoInterno()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub CarregarProdutoInterno()
        If ddlGrupoProdutoInterno.SelectedIndex > 0 Then
            ddl.Carregar(ddlProdutoInterno, CarregarDDL.Tabela.Produto, " Grupo = '" & ddlGrupoProdutoInterno.SelectedValue & "'")
        Else
            ddlProdutoInterno.Items.Clear()
        End If
    End Sub

    Protected Sub ddlProdutosInterno_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddlProdutoInterno.SelectedIndexChanged
        Try
            If ddlProdutoInterno.SelectedIndex = 0 Then
                lnkNovoI.Parent.Visible = False
            Else
                Dim objProduto As [Lib].Negocio.Produto = New Produto(ddlProdutoInterno.SelectedValue)

                If objProduto.ControlarEstoque Then

                    CarregarGridProdutoInterno(ddlProdutoInterno.SelectedValue, "")

                    lnkNovoI.Parent.Visible = True
                Else
                    lnkNovoI.Parent.Visible = False
                    MsgBox(Me.Page, "Produto selecionado não está marcado para CONTROLAR ESTOQUE.", eTitulo.Info)
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub Limpar(ByVal limpa As Boolean)

        HID.Value = Guid.NewGuid().ToString

        lnkNovoI.Parent.Visible = False
        lnkAtualizarI.Parent.Visible = False

        rdEntradas.Checked = False
        rdSaidas.Checked = False

        txtLoteInterno.Enabled = True
        txtDataProducaoInterna.Enabled = True
        txtDataValidadeInterna.Enabled = True
        txtQuantidadeInterna.Enabled = True
        ddlGrupoProdutoInterno.Enabled = True
        ddlProdutoInterno.Enabled = True
        lnkBuscaProdutoInterno.Enabled = True
        cmbSubOperacao.SelectedIndex = 0

        If limpa Then If ddlProdutoInterno.Items.Count > 0 Then ddlProdutoInterno.SelectedIndex = 0

        gridProducaoInterno.DataSource = Nothing
        gridProducaoInterno.DataBind()

        txtLoteInterno.Text = String.Empty
        txtQuantidadeInterna.Text = String.Empty
        txtObsProdutoInterno.Text = String.Empty

        txtDataProducaoInterna.Text = Now.ToString("dd/MM/yyyy")
        txtDataValidadeInterna.Text = Now.ToString("dd/MM/yyyy")

        ucConsultaProduto.SetarHID(HID.Value)

    End Sub

    Protected Sub lnkNovoI_Click(sender As Object, e As EventArgs) Handles lnkNovoI.Click
        If Funcoes.VerificaPermissao("OrdemDeProducaoInterna", "GRAVAR") Then

            If ddlUnidadeNegocio.SelectedIndex = 0 Then
                MsgBox(Me.Page, "Unidade de Negócio não foi selecionada.", eTitulo.Info)
            ElseIf ddlEmpresa.SelectedIndex = 0 Then
                MsgBox(Me.Page, "Empresa não foi selecionada.", eTitulo.Info)
            ElseIf String.IsNullOrWhiteSpace(txtQuantidadeInterna.Text) OrElse CDec(txtQuantidadeInterna.Text) = 0 Then
                MsgBox(Me.Page, "Quantidade não foi informada.", eTitulo.Info)
            ElseIf ddlProdutoInterno.SelectedIndex = 0 Then
                MsgBox(Me.Page, "Produto não foi informado.", eTitulo.Info)
            ElseIf ddlUnidadeNegocio.SelectedIndex = 0 Then
                MsgBox(Me.Page, "Unidade de Negócio não foi selecionada.", eTitulo.Info)
            ElseIf ddlEmpresa.SelectedIndex = 0 Then
                MsgBox(Me.Page, "Empresa não foi selecionada.", eTitulo.Info)
            ElseIf String.IsNullOrWhiteSpace(txtLoteInterno.Text) Then
                MsgBox(Me.Page, "Lote interno não foi informado.", eTitulo.Info)
            ElseIf rdEntradas.Checked = False AndAlso rdSaidas.Checked = False Then
                MsgBox(Me.Page, "Selecione Entrada ou Saída.", eTitulo.Info)
            Else
                Dim objBanco As New AcessaBanco()
                Dim sqls As New ArrayList

                Dim strSQL As String = "SELECT Lote_Id FROM OrdemDeProducaoInterna " & vbCrLf &
                                        "Where Empresa_Id    = '" & ddlEmpresa.SelectedValue.Split("-")(0) & "'" & vbCrLf &
                                        "  and EndEmpresa_Id = " & ddlEmpresa.SelectedValue.Split("-")(1) & vbCrLf &
                                        "  and Produto_Id    = '" & ddlProdutoInterno.SelectedValue & "'" & vbCrLf &
                                        "  and Lote_Id       = '" & Trim(txtLoteInterno.Text) & "'"

                Dim ds As DataSet = Banco.ConsultaDataSet(strSQL, "OrdemDeProducaoInterna")

                If ds.Tables(0).Rows.Count > 0 Then
                    MsgBox(Me.Page, "Lote interno já foi cadastrado.", eTitulo.Info)
                    Exit Sub
                End If

                Dim msg As String = String.Empty
                Dim EntSai As String = String.Empty

                If rdEntradas.Checked Then EntSai = "E"
                If rdSaidas.Checked Then EntSai = "S"

                If Not String.IsNullOrWhiteSpace(txtObsProdutoInterno.Text) Then msg = Funcoes.EliminarCaracteresEspeciais(txtObsProdutoInterno.Text)

                strSQL = "INSERT INTO OrdemDeProducaoInterna(Empresa_Id, EndEmpresa_Id, Produto_Id, Lote_Id, Movimento, Validade, Quantidade, Observacoes, UsuarioInclusao, UsuarioInclusaoData, EntradaSaida, Operacao, SubOperacao)" & vbCrLf &
                      "Values('" & ddlEmpresa.SelectedValue.Split("-")(0) & "'," & ddlEmpresa.SelectedValue.Split("-")(1) & ",'" & ddlProdutoInterno.SelectedValue & "','" & Trim(txtLoteInterno.Text) & "','" & CDate(txtDataProducaoInterna.Text).ToString("yyyy-MM-dd") & "','" & CDate(txtDataValidadeInterna.Text).ToString("yyyy-MM-dd") & "'" & vbCrLf &
                      "," & Str(CDec(txtQuantidadeInterna.Text)) & "," & IIf(msg.Length > 0, "'" & msg & "'", "''") & ",'" & HttpContext.Current.Session("ssNomeUsuario") & "', getdate(), '" & EntSai & "'," & cmbSubOperacao.SelectedValue.ToString.Split("-")(0) & ", " & cmbSubOperacao.SelectedValue.ToString.Split("-")(1) & ")"

                sqls.Add(strSQL)

                If objBanco.GravaBanco(sqls) Then

                    Dim prd As String = ddlProdutoInterno.SelectedValue

                    Limpar(False)

                    CarregarGridProdutoInterno(prd, "")

                Else
                    MsgBox(Me.Page, HttpContext.Current.Session("ssMessage").ToString, eTitulo.Erro)
                End If
            End If
        Else
            MsgBox(Me.Page, "Usuário sem permissão para gravar registro.", eTitulo.Info)
        End If
    End Sub

    Protected Sub lnkAtualizarI_Click(sender As Object, e As EventArgs) Handles lnkAtualizarI.Click
        If Funcoes.VerificaPermissao("OrdemDeProducaoInterna", "ALTERAR") Then
            Dim objBanco As New AcessaBanco()
            Dim sqls As New ArrayList

            Dim msg As String = String.Empty
            Dim EntSai As String = String.Empty

            If rdEntradas.Checked Then EntSai = "E"
            If rdSaidas.Checked Then EntSai = "S"

            If Not String.IsNullOrWhiteSpace(txtObsProdutoInterno.Text) Then msg = Funcoes.EliminarCaracteresEspeciais(txtObsProdutoInterno.Text)

            Dim strSQL As String = "Update OrdemDeProducaoInterna Set " & vbCrLf &
                                     "   Movimento            = '" & CDate(txtDataProducaoInterna.Text).ToString("yyyy-MM-dd") & "'" & vbCrLf &
                                     "  ,Validade             = '" & CDate(txtDataValidadeInterna.Text).ToString("yyyy-MM-dd") & "'" & vbCrLf &
                                     "  ,Quantidade           = " & Str(CDec(txtQuantidadeInterna.Text)) & vbCrLf &
                                     "  ,Observacoes          = " & IIf(msg.Length > 0, "'" & msg & "'", "''") & vbCrLf &
                                     "	,UsuarioAlteracao     = '" & HttpContext.Current.Session("ssNomeUsuario") & "'" & vbCrLf &
                                     "	,UsuarioAlteracaoData = getdate() " & vbCrLf &
                                     "	,EntradaSaida         = '" & EntSai & "'" & vbCrLf &
                                     "	,Operacao             = " & cmbSubOperacao.SelectedValue.ToString.Split("-")(0) & vbCrLf &
                                     "	,SubOperacao          = " & cmbSubOperacao.SelectedValue.ToString.Split("-")(1) & vbCrLf &
                                     "Where Empresa_Id    = '" & ddlEmpresa.SelectedValue.Split("-")(0) & "'" & vbCrLf &
                                     "  and EndEmpresa_Id = " & ddlEmpresa.SelectedValue.Split("-")(1) & vbCrLf &
                                     "  and Produto_Id    = '" & ddlProdutoInterno.SelectedValue & "'" & vbCrLf &
                                     "  and Lote_Id       = '" & Trim(txtLoteInterno.Text) & "'"

            sqls.Add(strSQL)

            If objBanco.GravaBanco(sqls) Then

                Dim prd As String = ddlProdutoInterno.SelectedValue

                Limpar(False)

                CarregarGridProdutoInterno(prd, "")

            Else
                MsgBox(Me.Page, HttpContext.Current.Session("ssMessage").ToString, eTitulo.Erro)
            End If
        Else
            MsgBox(Me.Page, "Usuário sem permissão para alterar registro.", eTitulo.Info)
        End If
    End Sub

    Protected Sub lnkConsultarI_Click(sender As Object, e As EventArgs) Handles lnkConsultarI.Click
        If Funcoes.VerificaPermissao("OrdemDeProducaoInterna", "LEITURA") Then

            lnkNovoI.Parent.Visible = False
            lnkAtualizarI.Parent.Visible = False

            rdEntradas.Checked = False
            rdSaidas.Checked = False

            Dim nLote As String = String.Empty
            Dim prd As String = String.Empty

            If txtLoteInterno.Text.Length > 0 Then nLote = Trim(txtLoteInterno.Text)

            If ddlProdutoInterno.SelectedValue.Length > 0 Then prd = ddlProdutoInterno.SelectedValue

            CarregarGridProdutoInterno(prd, nLote)

        Else
            MsgBox(Me.Page, "Usuário sem permissão para consultar registro.", eTitulo.Info)
        End If
    End Sub

    Protected Sub lnkLimparI_Click(sender As Object, e As EventArgs) Handles lnkLimparI.Click
        Limpar(True)
    End Sub

End Class