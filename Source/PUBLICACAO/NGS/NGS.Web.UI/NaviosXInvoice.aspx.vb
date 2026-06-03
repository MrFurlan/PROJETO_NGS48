Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class NaviosXInvoice
    Inherits BasePage

    Dim lstNavios As New ListNaviosXInvoice
    Dim ObjNavioxinvoice As New NavioXInvoice
    Dim ObjNavioxinvoiceXProduto As New NavioXInvoiceXProduto

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.setMenu(eModulo.Gestao)
            If Not IsPostBack And IsConnect Then

                If Not UsuarioServidor.KeyCodeActive Then
                    MsgBox(Me.Page, "Sistema com chave de licença expirada. Entre em contato com o suporte.", "~/Gestao.aspx", eTitulo.Info)
                    Exit Sub
                End If

                If Funcoes.VerificaPermissao("NaviosXInvoice", "ACESSAR") Then
                    ddl.Carregar(ddlGrupoProduto, CarregarDDL.Tabela.GrupoProduto)
                    ddl.Carregar(ddlPais, CarregarDDL.Tabela.Pais, "", True)

                    limpar()
                Else
                    MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Gestao.aspx")
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub ddlGrupoProduto_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddlGrupoProduto.SelectedIndexChanged
        Try
            CarregarProduto()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub CarregarProduto()
        ddl.Carregar(ddlProdutos, CarregarDDL.Tabela.Produto, " Grupo = '" & ddlGrupoProduto.SelectedValue & "'")
    End Sub

    Protected Sub lnkBuscaProduto_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lnkBuscaProduto.Click
        Try
            ucConsultaProduto.Limpar()
            Dim txtNome As TextBox = CType(ucConsultaProduto.FindControlRecursive("txtNome"), TextBox)
            Popup.ConsultaDeProduto(Me.Page, "objProdutoxINV" & HID.Value, txtNome.ClientID, True)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub limpar()
        Session.Remove("objConsultarNavios" & HID.Value)
        Session.Remove("objProdutoxINV" & HID.Value)

        HID.Value = Guid.NewGuid().ToString
        ucConsultarNavios.SetarHID(HID.Value)
        ucConsultaProduto.SetarHID(HID.Value)

        txtCodigo.Text = String.Empty
        txtObservacao.Text = String.Empty
        txtNavio.Text = String.Empty
        txtDescricao.Text = String.Empty
        txtDataDeChegada.Text = String.Empty
        txtCodigo.Enabled = False
        txtNavio.Enabled = False
        txtDescricao.Enabled = False
        chkAtivo.Checked = True
        chkAtivo.Enabled = False
        txtDataDeChegada.Enabled = True
        txtObservacao.Enabled = True

        ddlGrupoProduto.Enabled = True
        ddlProdutos.Enabled = True
        lnkBuscaProduto.Enabled = True

        ddlGrupoProduto.SelectedIndex = 0
        ddlProdutos.Items.Clear()

        ddlPais.Enabled = True
        ddlPais.SelectedIndex = 0

        lnkConsultarNavio.Enabled = True
        'Ver processo com o furlan.
        lnkConsultar.Parent.Visible = False

        lnkNovo.Parent.Visible = False
        lnkExcluir.Parent.Visible = False
        lnkAtualizar.Parent.Visible = False

        GridNaviosXInvoice.DataSource = Nothing
        GridNaviosXInvoice.DataBind()
        carregarInvoice()

        TabNavio.ActiveTabIndex = 0
    End Sub

    Protected Sub GridNaviosXInvoice_SelectedIndexChanged(ByVal sernder As Object, ByVal e As System.EventArgs) Handles GridNaviosXInvoice.SelectedIndexChanged
        Try
            ObjNavioxinvoice = New NavioXInvoice(GridNaviosXInvoice.SelectedRow.Cells(1).Text())

            txtCodigo.Text = ObjNavioxinvoice.Codigo
            txtObservacao.Text = ObjNavioxinvoice.Observacao
            txtNavio.Text = ObjNavioxinvoice.Navio_Id
            txtDataDeChegada.Text = ObjNavioxinvoice.DataDeChegada
            txtDescricao.Text = ObjNavioxinvoice.Descricao
            chkAtivo.Checked = ObjNavioxinvoice.Ativo

            ddlGrupoProduto.SelectedValue = ObjNavioxinvoice.NavioXInvoiceXProduto.Produto.CodigoGrupo
            CarregarProduto()
            ddlProdutos.SelectedValue = ObjNavioxinvoice.NavioXInvoiceXProduto.Produto.Codigo

            ddlPais.SelectedValue = ObjNavioxinvoice.CodigoPais
            'ddlPais.Enabled = False

            'ddlGrupoProduto.Enabled = False
            'ddlProdutos.Enabled = False
            'lnkBuscaProduto.Enabled = False

            lnkNovo.Parent.Visible = False
            lnkConsultar.Parent.Visible = False
            'lnkConsultarNavio.Enabled = False

            If chkAtivo.Checked Then
                txtObservacao.Enabled = True
                txtDataDeChegada.Enabled = True
                lnkAtualizar.Parent.Visible = True
                lnkExcluir.Parent.Visible = True
            Else
                txtObservacao.Enabled = False
                txtDataDeChegada.Enabled = False
                lnkAtualizar.Parent.Visible = False
                lnkExcluir.Parent.Visible = False
            End If

            TabNavio.ActiveTabIndex = 0

        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Public Overrides Sub Carregar(obj As IBaseEntity)

        If Session("objProdutoxINV" & HID.Value) IsNot Nothing Then
            Dim objProduto As [Lib].Negocio.Produto = CType(obj, [Lib].Negocio.Produto)

            ddlGrupoProduto.SelectedValue = objProduto.CodigoGrupo
            CarregarProduto()
            ddlProdutos.SelectedValue = objProduto.Codigo
            Session.Remove("objProdutoxINV" & HID.Value)
        ElseIf Session("objConsultarNavios" & HID.Value) IsNot Nothing Then
            Dim objNavio = CType(Session("objConsultarNavios" & HID.Value), [Lib].Negocio.Navio)
            txtNavio.Text = objNavio.Codigo
            txtDescricao.Text = objNavio.Descricao
            Session.Remove("objConsultarNavios" & HID.Value)
        End If
    End Sub

    Private Sub carregarInvoice()
        Dim lstNavios As ListNaviosXInvoice = New ListNaviosXInvoice(True)

        GridNaviosXInvoice.DataSource = lstNavios
        GridNaviosXInvoice.DataBind()
    End Sub

    Private Function validaCampos() As Boolean
        If ddlProdutos.SelectedValue.Length = 0 Then
            MsgBox(Me.Page, "Produto não foi selecionado.")
            Return False
        ElseIf ddlPais.SelectedValue.Length = 0 Then
            MsgBox(Me.Page, "País não foi selecionado.")
            Return False
        ElseIf String.IsNullOrWhiteSpace(txtDataDeChegada.Text) Then
            MsgBox(Me.Page, "Informe a data de chegada.")
            Return False
            'Retirado a pedido do Ricardo, verificar processo.
            'ElseIf CDate(txtDataDeChegada.Text) < CDate(Now().ToString("dd/MM/yyyy")) Then
            '    MsgBox(Me.Page, "Data de chegada não pode ser menor que hoje.")
            '    Return False
        ElseIf String.IsNullOrWhiteSpace(txtObservacao.Text) Then
            MsgBox(Me.Page, "Informe a observação.")
            Return False
        End If
        Return True
    End Function

    Protected Sub lnkConsultar_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lnkConsultar.Click
        Try
            If Funcoes.VerificaPermissao("NaviosXInvoice", "CONSULTAR") Then

                Dim sqlWhere As String
                If txtCodigo.Text <> String.Empty Then
                    sqlWhere = "Where NXI.Codigo_Id = " & txtCodigo.Text
                ElseIf txtNavio.Text <> String.Empty Then
                    sqlWhere = "Where NXI.Navio_Id = " & txtNavio.Text
                Else
                    MsgBox(Me.Page, "Código invoice ou do navio é obrigatório para consulta.", eTitulo.Info)
                    Exit Sub
                End If

                lstNavios = New ListNaviosXInvoice(True, sqlWhere)
                If lstNavios.Count = 0 Then
                    limpar()
                    MsgBox(Me.Page, "Invoice não encontrado!", eTitulo.Erro)
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para consultar registro.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkConsultarNavio_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lnkConsultarNavio.Click
        Try
            If Funcoes.VerificaPermissao("NaviosXInvoice", "LEITURA") Then

                ucConsultarNavios.Limpar()
                ucConsultarNavios.carregarNavios()
                Popup.ConsultarNavios(Me.Page, "objConsultarNavios" & HID.Value)
                lnkNovo.Parent.Visible = True
            Else
                MsgBox(Me.Page, "Usuário sem permissão para consultar navios.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkNovo_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lnkNovo.Click
        Try
            If Funcoes.VerificaPermissao("NaviosXInvoice", "GRAVAR") Then
                If validaCampos() Then

                    ObjNavioxinvoice = New NavioXInvoice()
                    ObjNavioxinvoice.DataDeChegada = txtDataDeChegada.Text
                    ObjNavioxinvoice.Observacao = txtObservacao.Text
                    ObjNavioxinvoice.Navio_Id = txtNavio.Text
                    ObjNavioxinvoice.Descricao = RTrim(txtDescricao.Text)
                    ObjNavioxinvoice.Ativo = chkAtivo.Checked
                    ObjNavioxinvoice.CodigoPais = ddlPais.SelectedValue
                    ObjNavioxinvoice.IUD = "I"

                    ObjNavioxinvoiceXProduto = New NavioXInvoiceXProduto()
                    ObjNavioxinvoiceXProduto.IUD = "I"
                    ObjNavioxinvoiceXProduto.CodigoProduto = ddlProdutos.SelectedValue

                    ObjNavioxinvoice.NavioXInvoiceXProduto = ObjNavioxinvoiceXProduto

                    If ObjNavioxinvoice.Salvar Then
                        MsgBox(Me.Page, "Registro inserido com sucesso.", eTitulo.Sucess)
                        limpar()
                    Else
                        MsgBox(Me.Page, HttpContext.Current.Session("ssMessage").ToString)
                    End If
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para gravar registro.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAtualizar_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lnkAtualizar.Click
        Try
            If Funcoes.VerificaPermissao("NaviosXInvoice", "ALTERAR") Then
                ObjNavioxinvoice = New NavioXInvoice()
                ObjNavioxinvoice.IUD = "U"
                ObjNavioxinvoice.Codigo = RTrim(txtCodigo.Text)
                ObjNavioxinvoice.Descricao = RTrim(txtDescricao.Text)
                ObjNavioxinvoice.Ativo = chkAtivo.Checked
                ObjNavioxinvoice.CodigoPais = ddlPais.SelectedValue
                ObjNavioxinvoice.Observacao = txtObservacao.Text
                ObjNavioxinvoice.DataDeChegada = txtDataDeChegada.Text

                If ObjNavioxinvoice.Salvar Then
                    MsgBox(Me.Page, "Registro alterado com sucesso.", eTitulo.Sucess)
                    limpar()
                Else
                    MsgBox(Me.Page, HttpContext.Current.Session("ssMessage").ToString)
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para alterar registro.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkExcluir_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkExcluir.Click
        Try
            If Funcoes.VerificaPermissao("NaviosXInvoice", "EXCLUIR") Then
                ObjNavioxinvoice = New NavioXInvoice(txtCodigo.Text)

                ObjNavioxinvoice.IUD = "D"
                ObjNavioxinvoice.Codigo = txtCodigo.Text

                If ObjNavioxinvoice.Salvar Then
                    MsgBox(Me.Page, "Registro excluído com sucesso.", eTitulo.Sucess)
                    limpar()
                Else
                    MsgBox(Me.Page, HttpContext.Current.Session("ssMessage").ToString)
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para excluir registro.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkLimpar_Click(ByVal sernder As Object, ByVal e As System.EventArgs) Handles lnkLimpar.Click
        limpar()
    End Sub

    Protected Sub lnkRelatorio_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkRelatorio.Click
        Try
            If Funcoes.VerificaPermissao("NaviosXInvoice", "RELATORIO") Then
                MsgBox(Me.Page, "Em desenvolvimento.")
            Else
                MsgBox(Me.Page, "Usuário sem permissão para emitir o relatório.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAjuda_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "Ñavios")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub
End Class