Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class PrecoDePauta
    Inherits BasePage

    Private objPrecoDePauta As [Lib].Negocio.PrecoDePauta

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.setMenu(eModulo.Gestao)
            If Not IsPostBack And IsConnect Then
                If Funcoes.VerificaPermissao("PrecoDePauta", "ACESSAR") Then
                    Carregar_Estados()
                    Carregar_Produtos()
                    Limpar()
                Else
                    MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Gestao.aspx")
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub Carregar_Produtos()
        Dim ListPrd As New [Lib].Negocio.ListProduto("", "", "", "", "", "ControlarPrecoDePauta = 'TRUE'")
        gridProduto.DataSource = ListPrd.ToArray
        gridProduto.DataBind()
    End Sub

    Private Sub Carregar_Estados()
        ddl.Carregar(ddlEstado, CarregarDDL.Tabela.Estados, "", True)
    End Sub

    Private Sub Limpar()
        lnkNovo.Parent.Visible = True
        lnkAtualizar.Parent.Visible = False
        lnkExcluir.Parent.Visible = False
        ddlEstado.Enabled = False
        txtData.ReadOnly = True
        txtPreco.ReadOnly = True
        txtData.Text = ""
        txtPreco.Text = ""
        ddlEstado.SelectedIndex = 0
        gridPrecoDePauta.DataBind()
    End Sub

    Protected Sub gridProduto_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles gridProduto.SelectedIndexChanged
        Try
            AtualizarGrid()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub AtualizarGrid()
        Dim ListPPauta As New [Lib].Negocio.ListPrecoDePauta(gridProduto.Rows(gridProduto.SelectedIndex).Cells(1).Text())
        gridPrecoDePauta.DataSource = ListPPauta
        gridPrecoDePauta.DataBind()
        lnkNovo.Parent.Visible = True
        lnkAtualizar.Parent.Visible = False
        lnkExcluir.Parent.Visible = False
        ddlEstado.Enabled = True
        txtData.ReadOnly = False
        txtPreco.ReadOnly = False
        txtData.Text = ""
        txtPreco.Text = ""
        ddlEstado.SelectedIndex = 0
    End Sub

    Protected Sub gridPrecoDePauta_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            txtData.Text = gridPrecoDePauta.Rows(gridPrecoDePauta.SelectedIndex).Cells(2).Text()
            txtPreco.Text = gridPrecoDePauta.Rows(gridPrecoDePauta.SelectedIndex).Cells(3).Text()
            ddlEstado.SelectedValue = gridPrecoDePauta.Rows(gridPrecoDePauta.SelectedIndex).Cells(1).Text()
            ddlEstado.Enabled = False
            txtData.ReadOnly = True
            txtPreco.ReadOnly = False
            lnkNovo.Parent.Visible = False
            lnkAtualizar.Parent.Visible = True
            lnkExcluir.Parent.Visible = True
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkNovo_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkNovo.Click
        Try
            If Funcoes.VerificaPermissao("PrecoDePauta", "GRAVAR") Then
                objPrecoDePauta = New [Lib].Negocio.PrecoDePauta()
                objPrecoDePauta.CodigoProduto = gridProduto.Rows(gridProduto.SelectedIndex).Cells(1).Text()
                objPrecoDePauta.CodigoEstado = ddlEstado.SelectedValue
                objPrecoDePauta.Data = txtData.Text
                objPrecoDePauta.Preco = txtPreco.Text
                objPrecoDePauta.IUD = "I"
                If objPrecoDePauta.Salvar Then
                    AtualizarGrid()
                Else
                    MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para gravar registro.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAtualizar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkAtualizar.Click
        Try
            If Funcoes.VerificaPermissao("PrecoDePauta", "ALTERAR") Then
                objPrecoDePauta = New [Lib].Negocio.PrecoDePauta()
                objPrecoDePauta.CodigoProduto = gridProduto.Rows(gridProduto.SelectedIndex).Cells(1).Text()
                objPrecoDePauta.CodigoEstado = ddlEstado.SelectedValue
                objPrecoDePauta.Data = txtData.Text
                objPrecoDePauta.Preco = txtPreco.Text
                objPrecoDePauta.IUD = "U"
                If objPrecoDePauta.Salvar Then
                    AtualizarGrid()
                Else
                    MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
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
            If Funcoes.VerificaPermissao("PrecoDePauta", "EXCLUIR") Then
                objPrecoDePauta = New [Lib].Negocio.PrecoDePauta()
                objPrecoDePauta.CodigoProduto = gridProduto.Rows(gridProduto.SelectedIndex).Cells(1).Text()
                objPrecoDePauta.CodigoEstado = ddlEstado.SelectedValue
                objPrecoDePauta.Data = txtData.Text
                objPrecoDePauta.Preco = txtPreco.Text
                objPrecoDePauta.IUD = "D"
                If objPrecoDePauta.Salvar Then
                    AtualizarGrid()
                Else
                    MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para excluir registro.", eTitulo.Info)
            End If
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

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs)
        Try
            Funcoes.Ajuda(Me.Page, "PrecoDePauta")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub
End Class