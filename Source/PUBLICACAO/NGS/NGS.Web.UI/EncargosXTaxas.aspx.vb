Imports System.Data
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class EncargosXTaxas
    Inherits BasePage

    Dim SqlArray As New ArrayList
    Dim Sql As String
    Dim ds As DataSet
    Dim Row As DataRow
    Dim Mensagem As String = ""

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Gestao)
        If Not IsPostBack And IsConnect Then
            If Funcoes.VerificaPermissao("EncargosXTaxas", "ACESSAR") Then
                CargaEncargosXTaxas()
                CarregarUF()
                CarregarEncargos()
                CarregarGrupoDeProdutos()
                Limpar()
            Else
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Gestao.aspx")
                Exit Sub
            End If
        End If
    End Sub

    Private Sub CarregarGrupoDeProdutos()
        ddl.Carregar(ddlGrupoProduto, CarregarDDL.Tabela.GrupoProduto, "", True)
    End Sub

    Private Sub CarregarProdutos()
        ddl.Carregar(ddlProdutos, CarregarDDL.Tabela.ProdutoPorGrupo, ddlGrupoProduto.SelectedValue, True)
    End Sub

    Private Sub CarregarEncargos()
        'If Funcoes.VerificaPermissao("EncargosXTaxas", "LEITURA") Then
        '    Dim strSQL As String
        '    strSQL = "SELECT Encargo_id as Descricao " & vbCrLf & _
        '             "FROM Encargos " & vbCrLf & _
        '             "ORDER BY Encargo_id" & vbCrLf

        '    Dim dsEncargo As DataSet = Banco.ConsultaDataSet(strSQL, "Encargos")

        '    CmbEncargo.DataSource = dsEncargo
        '    CmbEncargo.DataBind()


        '    CmbEncargo.Items.Insert(0, "")
        '    CmbEncargo.SelectedIndex = 0
        'Else
        '    MsgBox(Me.Page, "Usuário sem permissão para ler os registros!")
        'End If
        ddl.Carregar(CmbEncargo, CarregarDDL.Tabela.Encargos, "", True)
    End Sub

    Private Sub CarregarUF()
        'Dim strSQL As String

        'strSQL = "SELECT Estado_Id, Estado_Id + ' - ' + Descricao AS Descricao " & vbCrLf & _
        '         "FROM Estados " & vbCrLf & _
        '         "ORDER BY Estado_Id" & vbCrLf

        'Dim dsEstados As DataSet = Banco.ConsultaDataSet(strSQL, "Estados")

        'cmbEstado.DataSource = dsEstados
        'cmbEstado.DataBind()
        ddl.Carregar(cmbEstado, CarregarDDL.Tabela.Estados, "", True)
    End Sub

    Private Sub CargaEncargosXTaxas()
        ' incluir rotina de carga de grid. 
        Sql = "Select Estado_id , Encargo_id, Produto_Id, Data_id , Percentual, SimplesNacional" & vbCrLf & _
                "from EncargosXTaxas" & vbCrLf & _
                "order by Estado_id, Encargo_id, Produto_Id, Data_id desc"

        ds = Banco.ConsultaDataSet(Sql, "EncargosXTaxas")

        Gridencargosxtaxas.DataSource = Banco.ConsultaDataSet(Sql, "EncargosXTaxas")
        Gridencargosxtaxas.DataBind()
    End Sub

    Protected Sub lnkBuscaProduto_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lnkBuscaProduto.Click
        Try
            ucConsultaProduto.Limpar()
            Dim txtNome As TextBox = CType(ucConsultaProduto.FindControlRecursive("txtNome"), TextBox)
            Popup.ConsultaDeProduto(Me.Page, "objProdutoXTAXA" & HID.Value, txtNome.ClientID, True)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub ddlGrupoProduto_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddlGrupoProduto.SelectedIndexChanged
        Try
            CarregarProdutos()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Public Overrides Sub Carregar(obj As IBaseEntity)
        If Session("objProdutoXTAXA" & HID.Value) IsNot Nothing Then
            Dim objProduto As [Lib].Negocio.Produto = CType(obj, [Lib].Negocio.Produto)

            If objProduto.ControlarEstoque Then
                ddlGrupoProduto.SelectedValue = objProduto.CodigoGrupo
                CarregarProdutos()
                ddlProdutos.SelectedValue = objProduto.Codigo
                Session.Remove("objProdutoXTAXA" & HID.Value)
            Else
                MsgBox(Me.Page, "Produto selecionado não está marcado para CONTROLAR ESTOQUE.", eTitulo.Info)
            End If
        End If
    End Sub

    Protected Sub Gridencargosxtaxas_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            cmbEstado.SelectedValue = Gridencargosxtaxas.SelectedRow.Cells(1).Text()
            CmbEncargo.SelectedValue = Gridencargosxtaxas.SelectedRow.Cells(2).Text()

            If Not Gridencargosxtaxas.SelectedRow.Cells(3).Text() = "&nbsp;" AndAlso Gridencargosxtaxas.SelectedRow.Cells(3).Text().Length > 0 Then
                Dim prd As New Produto(Gridencargosxtaxas.SelectedRow.Cells(3).Text())
                ddlGrupoProduto.SelectedValue = prd.CodigoGrupo
                ddl.Carregar(ddlProdutos, CarregarDDL.Tabela.ProdutoPorGrupo, ddlGrupoProduto.SelectedValue, True)

                ddlProdutos.SelectedValue = Gridencargosxtaxas.SelectedRow.Cells(3).Text()
            End If

            txtData.Text = Gridencargosxtaxas.SelectedRow.Cells(4).Text()
            txtPercentual.Text = Gridencargosxtaxas.SelectedRow.Cells(5).Text()
            txtSimplesNacional.Text = Gridencargosxtaxas.SelectedRow.Cells(6).Text()

            cmbEstado.Enabled = False
            CmbEncargo.Enabled = False
            ddlGrupoProduto.Enabled = False
            ddlProdutos.Enabled = False
            lnkBuscaProduto.Enabled = False
            txtData.Enabled = False

            lnkNovo.Parent.Visible = False
            lnkAtualizar.Parent.Visible = True
            lnkExcluir.Parent.Visible = True
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub Limpar()
        txtData.Enabled = True
        cmbEstado.Enabled = True
        CmbEncargo.Enabled = True
        ddlGrupoProduto.Enabled = True
        ddlProdutos.Enabled = True
        lnkBuscaProduto.Enabled = True

        txtData.Text = String.Empty
        txtPercentual.Text = "0"
        txtSimplesNacional.Text = "0"

        cmbEstado.SelectedIndex = 0
        CmbEncargo.SelectedIndex = 0
        ddlGrupoProduto.SelectedIndex = 0

        ddlProdutos.Items.Clear()
        lnkNovo.Parent.Visible = True
        lnkAtualizar.Parent.Visible = False
        lnkExcluir.Parent.Visible = False

        Session.Remove("objProdutoXTAXA" & HID.Value)

        HID.Value = Guid.NewGuid().ToString
    End Sub

    Private Sub GravaEncargosXTaxas()
        If Funcoes.VerificaPermissao("EncargosXTaxas", "GRAVAR") Then

            'Dim i As Integer
            'ValidaCampos()
            'cmbEstado.SelectedIndex = -1
            'Dim x As ListItem = cmbEstado.Items.FindByValue(Gridencargosxtaxas.SelectedRow.Cells(1).Text)
            'x.Selected = True
            'cmbEstado.SelectedValue = Gridencargosxtaxas.SelectedRow.Cells(1).Text()
            'CmbEncargo.SelectedValue = Gridencargosxtaxas.SelectedRow.Cells(2).Text()
            'txtData.Text = Gridencargosxtaxas.SelectedRow.Cells(3).Text()
            'txtPercentual.Text = Gridencargosxtaxas.SelectedRow.Cells(4).Text()

            'ExcluirEncargosXTaxas()
            'If Mensagem = "" Then
            Sql = " Insert into EncargosXTaxas " & vbCrLf & _
                  " (Estado_id , " & vbCrLf & _
                  "Encargo_id , " & vbCrLf & _
                  "Produto_Id , " & vbCrLf & _
                  "Data_id , " & vbCrLf & _
                  "Percentual, SimplesNacional) " & vbCrLf & _
                  " Values (" & vbCrLf & _
                  "'" & cmbEstado.SelectedValue & "'," & vbCrLf & _
                  "'" & CmbEncargo.SelectedValue & "'," & vbCrLf & _
                  "'" & IIf(ddlProdutos.SelectedValue.Length > 0, ddlProdutos.SelectedValue, "") & "'," & vbCrLf & _
                  "'" & Convert.ToDateTime(txtData.Text).ToString("yyyy-MM-dd") & "'" & vbCrLf & _
                  ", " & Str(txtPercentual.Text) & vbCrLf & _
                  ", " & Str(txtSimplesNacional.Text) & ")" & vbCrLf

            SqlArray.Add(Sql)

            If Banco.GravaBanco(SqlArray) Then
                MsgBox(Me.Page, "Registro gravado com sucesso.", eTitulo.Sucess)
                Limpar()
                CargaEncargosXTaxas()
            Else
                MsgBox(Me.Page, "Erro: " & HttpContext.Current.Session("ssMessage"), eTitulo.Erro)
            End If
            'End If
            'Sql = "INSERT INTO EncargosXTaxas "
            'Sql &= " (Pagamento_Id"
            'Sql &= " ,Descricao"
            'Sql &= " ,Parcelas)"
            'Sql &= " VALUES( "
            'Sql &= CInt(txtCodigo.Text)
            'Sql &= ", '" & txtDescricao.Text & "'"
            'Sql &= ", " & CInt(txtParcelas.Text) & ")"
            'SqlArray.Add(Sql)
            'Sql &= " (Estado_id , "        ''While i < GridParcelas.Rows.Count
            'Sql = "INSERT INTO EncargosXTaxasXParcelas "
            ' Sql &= " (Pagamento_Id"
            'Sql &= " ,Sequencia_Id"
            'Sql &= " ,Dias)"
            'Sql &= " VALUES( "
            'Sql &= CInt(txtCodigo.Text)
            'Sql &= ", " & i + 1
            'Sql &= ", " & CType(GridParcelas.Rows(i).FindControl("txtDias"), TextBox).Text()
            'Sql &= ")"
            'SqlArray.Add(Sql)
            ' i += 1
            'End While
            'If Banco.GravaBanco(SqlArray) = False Then
            '    ScriptManager.RegisterClientScriptBlock(Me, Me.lnkNovo.GetType(), Guid.NewGuid().ToString(), "mensagemServidor = '" & Mensagem & "';", True)
            'Else
            '    Mensagem = "Registro Gravado/Atualizado com sucesso."
            '    ScriptManager.RegisterClientScriptBlock(Me, Me.lnkNovo.GetType(), Guid.NewGuid().ToString(), "mensagemServidor = '" & Mensagem & "';", True)
            '    Limpar()
            'End If
        Else
            MsgBox(Me.Page, "Usuário sem permissão para gravar o registro.")
        End If
    End Sub

    Private Sub AlterarEncargosXTaxas()
        Sql = " Update EncargosXTaxas " & vbCrLf & _
              "  set Percentual = " & Str(txtPercentual.Text) & vbCrLf & _
              "     ,SimplesNacional = " & Str(txtSimplesNacional.Text) & vbCrLf & _
              "  WHERE Estado_id = '" & cmbEstado.SelectedValue & "'" & vbCrLf & _
              "  and Encargo_id = '" & CmbEncargo.SelectedValue & "'" & vbCrLf & _
              "  and Produto_id = '" & IIf(ddlProdutos.SelectedValue.Length > 0, ddlProdutos.SelectedValue, "") & "'" & vbCrLf & _
              "  and Data_id = '" & Convert.ToDateTime(txtData.Text).ToString("yyyy-MM-dd") & "'"

        SqlArray.Add(Sql)

        If Banco.GravaBanco(SqlArray) Then
            MsgBox(Me.Page, "Registro alterado com sucesso.", eTitulo.Sucess)
            Limpar()
            CargaEncargosXTaxas()
        Else
            MsgBox(Me.Page, "Erro: " & HttpContext.Current.Session("ssMessage"), eTitulo.Erro)
        End If

    End Sub


    Private Sub ExcluirEncargosXTaxas()
        If Funcoes.VerificaPermissao("EncargosXTaxas", "EXCLUIR") Then

            Sql = " DELETE EncargosXTaxas " & vbCrLf & _
                  "  WHERE Estado_id = '" & cmbEstado.SelectedValue & "'" & vbCrLf & _
                  "  and Encargo_id = '" & CmbEncargo.SelectedValue & "'" & vbCrLf & _
                  "  and Produto_id = '" & IIf(ddlProdutos.SelectedValue.Length > 0, ddlProdutos.SelectedValue, "") & "'" & vbCrLf & _
                  "  and Data_id = '" & Convert.ToDateTime(txtData.Text).ToString("yyyy-MM-dd") & "'"

            SqlArray.Add(Sql)

            If Banco.GravaBanco(SqlArray) Then
                MsgBox(Me.Page, "Registro excluído com sucesso.", eTitulo.Sucess)
                Limpar()
                CargaEncargosXTaxas()
            Else
                MsgBox(Me.Page, "Erro: " & HttpContext.Current.Session("ssMessage"), eTitulo.Erro)
            End If
        Else
            MsgBox(Me.Page, "Usuário sem permissão para excluir o registro.")
        End If
    End Sub

    Private Sub CriaParcelas()
        'Dim i As Integer = 0
        'Dim dsParcelas As New DataSet
        'Dim dtParcelas As DataTable

        'If txtParcelas.Text <> "" Then
        '    dtParcelas = New DataTable("Parcelas")
        '    dtParcelas.Columns.Add("Sequencia", Type.GetType("System.String")).DefaultValue = ""

        '    While i < CInt(txtParcelas.Text)
        '        Row = dtParcelas.NewRow()
        '        Row("Sequencia") = i + 1
        '        dtParcelas.Rows.Add(Row)
        '        i += 1
        '    End While

        '    GridParcelas.DataSource = dtParcelas
        '    GridParcelas.DataBind()
        'End If
    End Sub

    Protected Sub lnkNovo_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkNovo.Click
        Try
            GravaEncargosXTaxas()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAtualizar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkAtualizar.Click
        Try
            AlterarEncargosXTaxas()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkExcluir_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkExcluir.Click
        Try
            ExcluirEncargosXTaxas()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkLimpar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkLimpar.Click
        Try
            Limpar()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkRelatorio_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkRelatorio.Click
        Try
            If Funcoes.VerificaPermissao("EncargosXTaxas", "RELATORIO") Then
                Sql = "SELECT Estado_id , Encargo_id , Data_id , Percentual, SimplesNacional " & vbCrLf & _
                      "FROM  EncargosXTaxas" & vbCrLf
                Dim Ds As DataSet = Banco.ConsultaDataSet(Sql, "EncargosXTaxas")

                Funcoes.BindReport(Me.Page, Ds, "Cr_EncargosXTaxas", eExportType.PDF)
            Else
                MsgBox(Me.Page, "Usuário sem permissão para visualizar o relatório.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAjuda_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "EncargosXTaxas")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub

End Class