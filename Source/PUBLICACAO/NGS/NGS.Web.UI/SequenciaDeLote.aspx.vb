Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis
Public Class SequenciaDeLote
    Inherits BasePage

    Private ds As DataSet
    Private objPrd As [Lib].Negocio.Produto

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Gestao)
        If Not IsPostBack And IsConnect Then
            If Funcoes.VerificaPermissao("SequenciaDeLote", "ACESSAR") Then
                ddl.Carregar(ddlGrupoProduto, CarregarDDL.Tabela.GrupoProdutoXConsumo)
                Limpar()
            Else
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Gestao.aspx")
                Exit Sub
            End If
        End If
    End Sub

    Public Overrides Sub Carregar(obj As IBaseEntity)
        If Session("objProdutoxPS" & HID.Value) IsNot Nothing Then
            Dim objProduto As [Lib].Negocio.Produto = CType(obj, [Lib].Negocio.Produto)

            If objProduto.ControlarEstoque Then
                ddlGrupoProduto.SelectedValue = objProduto.CodigoGrupo
                CarregarProduto()
                ddlProduto.SelectedValue = objProduto.Codigo
                Session.Remove("objProdutoxPS" & HID.Value)
            Else
                MsgBox(Me.Page, "Produto selecionado não está marcado para CONTROLAR ESTOQUE.", eTitulo.Info)
            End If
        End If
    End Sub

    Private Sub AtualizarGrid()
        objPrd = New Produto()

        Dim objLista = New ListProdutoXSequenciaDeLote(objPrd)

        If objLista.Count > 0 Then
            If rdNome.Checked Then gridSequenciaDeLote.DataSource = objLista.OrderBy(Function(s) s.Produto.Nome).ToList()

            If rdSequencia.Checked Then gridSequenciaDeLote.DataSource = objLista.OrderBy(Function(s) s.SequenciaDoProduto).ToList()
        Else
            gridSequenciaDeLote.DataSource = Nothing
        End If

        gridSequenciaDeLote.DataBind()
    End Sub

    Protected Sub ddlGrupoProduto_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddlGrupoProduto.SelectedIndexChanged
        CarregarProduto()
    End Sub

    Private Sub CarregarProduto()
        Try
            ddl.Carregar(ddlProduto, CarregarDDL.Tabela.ProdutoProducao, " P.Grupo = '" & ddlGrupoProduto.SelectedValue & "'")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkBuscaProdutoProducao_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lnkBuscaProduto.Click
        Try
            ucConsultaProduto.Limpar()
            Dim txtNome As TextBox = CType(ucConsultaProduto.FindControlRecursive("txtNome"), TextBox)
            Popup.ConsultaDeProduto(Me.Page, "objProdutoxPS" & HID.Value, txtNome.ClientID, True)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub gridSequenciaDeLote_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        lnkNovo.Parent.Visible = False
        lnkAtualizar.Parent.Visible = True
        lnkExcluir.Parent.Visible = True

        objPrd = New Produto(gridSequenciaDeLote.SelectedRow.Cells(1).Text())

        ddlGrupoProduto.SelectedValue = objPrd.CodigoGrupo
        CarregarProduto()

        With ddlProduto
            ddlProduto.SelectedIndex = .Items.IndexOf(.Items.FindByValue(objPrd.Codigo))
        End With

        txtSequenciaProduto.Text = gridSequenciaDeLote.SelectedRow.Cells(3).Text()
        txtSequenciaLote.Text = gridSequenciaDeLote.SelectedRow.Cells(4).Text()
        txtAno.Text = gridSequenciaDeLote.SelectedRow.Cells(5).Text()

        ddlGrupoProduto.Enabled = False
        lnkBuscaProduto.Enabled = False
        ddlProduto.Enabled = False
        txtSequenciaProduto.Enabled = False
        txtAno.Enabled = False

    End Sub

    Private Sub Limpar()
        Session.Remove("objProdutoxPS" & HID.Value)

        lnkNovo.Parent.Visible = True
        lnkAtualizar.Parent.Visible = False
        lnkExcluir.Parent.Visible = False

        ddlGrupoProduto.Enabled = True
        lnkBuscaProduto.Enabled = True
        ddlProduto.Enabled = True
        txtSequenciaProduto.Enabled = True
        txtSequenciaLote.Enabled = True
        txtAno.Enabled = True

        ddlGrupoProduto.SelectedIndex = 0
        ddlProduto.Items.Clear()

        txtSequenciaProduto.Text = String.Empty
        txtSequenciaLote.Text = String.Empty
        txtAno.Text = Now.Year.ToString()

        HID.Value = Guid.NewGuid().ToString

        AtualizarGrid()
    End Sub

    Function ValidarSelecao() As Boolean
        If ddlGrupoProduto.SelectedIndex = 0 Then
            MsgBox(Me.Page, "Grupo do Produto não foi selecionado.", eTitulo.Info)
            Return False
        ElseIf ddlProduto.SelectedIndex = 0 Then
            MsgBox(Me.Page, "Produto não foi selecionado.", eTitulo.Info)
            Return False
        ElseIf String.IsNullOrWhiteSpace(txtSequenciaProduto.Text) Then
            MsgBox(Me.Page, "Sequência do Produto não foi informada.", eTitulo.Info)
            Return False
        ElseIf String.IsNullOrWhiteSpace(txtSequenciaLote.Text) Then
            MsgBox(Me.Page, "Sequência do Lote não foi informada.", eTitulo.Info)
            Return False
        ElseIf String.IsNullOrWhiteSpace(txtAno.Text) Then
            MsgBox(Me.Page, "Ano não foi informado.", eTitulo.Info)
            txtAno.Text = Now.Year.ToString()
            Return False
        ElseIf ddlGrupoProduto.Enabled = True AndAlso CInt(txtAno.Text) < Now.Year Then
            MsgBox(Me.Page, "Ano não pode ser menor que o atual.", eTitulo.Info)
            Return False
        Else
            Return True
        End If
    End Function

    Protected Sub lnkNovo_Click(sender As Object, e As EventArgs) Handles lnkNovo.Click
        If Funcoes.VerificaPermissao("SequenciaDeLote", "GRAVAR") Then
            If ValidarSelecao() Then
                Dim objSequencia = New ProdutoXSequenciaDeLote()
                objSequencia.IUD = "I"
                objSequencia.Produto = New Produto(ddlProduto.SelectedValue)
                objSequencia.SequenciaDoProduto = Trim(txtSequenciaProduto.Text)
                objSequencia.SequenciaDoLote = Trim(txtSequenciaLote.Text)
                objSequencia.Ano = CInt(txtAno.Text)

                If objSequencia.Salvar Then
                    Limpar()
                Else
                    MsgBox(Me.Page, HttpContext.Current.Session("ssMessage").ToString())
                End If
            End If
        Else
            MsgBox(Me.Page, "Usuário sem permissão para gravar registro.")
        End If
    End Sub

    Protected Sub lnkAtualizar_Click(sender As Object, e As EventArgs) Handles lnkAtualizar.Click
        If Funcoes.VerificaPermissao("SequenciaDeLote", "ALTERAR") Then
            If ValidarSelecao() Then
                Dim objSequencia = New ProdutoXSequenciaDeLote()
                objSequencia.IUD = "U"
                objSequencia.Produto = New Produto(ddlProduto.SelectedValue)
                objSequencia.SequenciaDoProduto = Trim(txtSequenciaProduto.Text)
                objSequencia.SequenciaDoLote = Trim(txtSequenciaLote.Text)
                objSequencia.Ano = CInt(txtAno.Text)

                If objSequencia.Salvar Then
                    Limpar()
                Else
                    MsgBox(Me.Page, HttpContext.Current.Session("ssMessage").ToString())
                End If
            End If
        Else
            MsgBox(Me.Page, "Usuário sem permissão para alterar registro.")
        End If
    End Sub

    Protected Sub lnkExcluir_Click(sender As Object, e As EventArgs) Handles lnkExcluir.Click
        If Funcoes.VerificaPermissao("SequenciaDeLote", "EXCLUIR") Then
            If ValidarSelecao() Then
                Dim objSequencia = New ProdutoXSequenciaDeLote()
                objSequencia.IUD = "D"
                objSequencia.Produto = New Produto(ddlProduto.SelectedValue)
                objSequencia.SequenciaDoProduto = Trim(txtSequenciaProduto.Text)
                objSequencia.SequenciaDoLote = Trim(txtSequenciaLote.Text)
                objSequencia.Ano = CInt(txtAno.Text)

                If objSequencia.Salvar Then
                    Limpar()
                Else
                    MsgBox(Me.Page, HttpContext.Current.Session("ssMessage").ToString())
                End If
            End If
        Else
            MsgBox(Me.Page, "Usuário sem permissão para alterar registro.")
        End If
    End Sub

    Protected Sub lnkLimpar_Click(sender As Object, e As EventArgs) Handles lnkLimpar.Click
        Limpar()
    End Sub

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "SequenciaDeLote")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub

    Protected Sub rdNome_CheckedChanged(sender As Object, e As EventArgs) Handles rdNome.CheckedChanged
        Limpar()
    End Sub

    Protected Sub rdSequencia_CheckedChanged(sender As Object, e As EventArgs) Handles rdSequencia.CheckedChanged
        Limpar()
    End Sub
End Class