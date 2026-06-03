Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class UnidadeConversao
    Inherits BasePage

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Gestao)
        If Not IsPostBack And IsConnect Then
            If Funcoes.VerificaPermissao("UnidadeConversao", "ACESSAR") Then
                Limpar()
                CarregaUnidades()
                AtualizarGrid()
            Else
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Gestao.aspx")
                Exit Sub
            End If
        End If
    End Sub

    Private Sub CarregaUnidades()
        ddl.Carregar(ddlUnidadeOrigem, CarregarDDL.Tabela.UnidadeDeMedida)
        ddl.Carregar(ddlUnidadeDestino, CarregarDDL.Tabela.UnidadeDeMedida)
    End Sub

    Private Function ValidarCampos() As Boolean
        If ddlUnidadeOrigem.SelectedIndex = 0 Then
            MsgBox(Me.Page, "Unidade de Origem não informada")
            Return False
        ElseIf ddlUnidadeDestino.SelectedIndex = 0 Then
            MsgBox(Me.Page, "Unidade de Destino não informada")
            Return False
        ElseIf String.IsNullOrWhiteSpace(txtFator.Text) AndAlso Not CDec(txtFator.Text) Then
            MsgBox(Me.Page, "Fator de conversão")
            Return False
        Else
            Return True
        End If
    End Function

    Private Sub IniciarUnidade(ByVal Tipo As String)
        If ValidarCampos() Then
            Dim objUnidConversao As New [Lib].Negocio.UnidadeConversao
            objUnidConversao.CodigoUnidadeOrigem = ddlUnidadeOrigem.SelectedValue
            objUnidConversao.CodigoUnidadeDestino = ddlUnidadeDestino.SelectedValue
            objUnidConversao.Fator = txtFator.Text
            objUnidConversao.IUD = Tipo
            If objUnidConversao.Salvar Then
                Limpar()
                AtualizarGrid()
            Else
                MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
            End If
        End If
    End Sub

    Private Sub Limpar()
        ddlUnidadeOrigem.SelectedIndex = 0
        ddlUnidadeDestino.SelectedIndex = 0
        ddlUnidadeOrigem.Enabled = True
        ddlUnidadeDestino.Enabled = True
        txtFator.Text = 0
        lnkNovo.Parent.Visible = True
        lnkAtualizar.Parent.Visible = False
        lnkExcluir.Parent.Visible = False
    End Sub

    Protected Sub GridUnidadeConversao_SelectedIndexChanged(sender As Object, e As EventArgs) Handles GridUnidadeConversao.SelectedIndexChanged
        Try
            ddlUnidadeOrigem.SelectedValue = GridUnidadeConversao.SelectedRow.Cells(1).Text()
            ddlUnidadeDestino.SelectedValue = GridUnidadeConversao.SelectedRow.Cells(2).Text()
            txtFator.Text = CType(GridUnidadeConversao.SelectedRow.FindControl("txtFator"), TextBox).Text

            lnkNovo.Parent.Visible = False
            lnkAtualizar.Parent.Visible = True
            lnkExcluir.Parent.Visible = True
            ddlUnidadeOrigem.Enabled = False
            ddlUnidadeDestino.Enabled = False
            txtFator.Focus()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub AtualizarGrid()
        Dim Lista As New [Lib].Negocio.ListUnidadeConversao
        GridUnidadeConversao.DataSource = Lista.ToArray
        GridUnidadeConversao.DataBind()
    End Sub

    Protected Sub lnkNovo_Click(sender As Object, e As EventArgs) Handles lnkNovo.Click
        Try
            If Funcoes.VerificaPermissao("UnidadeConversao", "ALTERAR") Then
                IniciarUnidade("I")
            Else
                MsgBox(Me.Page, "Usuário sem permissão para alterar registro.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAtualizar_Click(sender As Object, e As EventArgs) Handles lnkAtualizar.Click
        Try
            If Funcoes.VerificaPermissao("UnidadeConversao", "ALTERAR") Then
                IniciarUnidade("U")
            Else
                MsgBox(Me.Page, "Usuário sem permissão para alterar registro.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkExcluir_Click(sender As Object, e As EventArgs) Handles lnkExcluir.Click
        Try
            If Funcoes.VerificaPermissao("UnidadeConversao", "EXCLUIR") Then
                IniciarUnidade("D")
            Else
                MsgBox(Me.Page, "Usuário sem permissão para excluir registro.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkLimpar_Click(sender As Object, e As EventArgs) Handles lnkLimpar.Click
        Try
            Limpar()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkRelatorio_Click(sender As Object, e As EventArgs) Handles lnkRelatorio.Click
        Try
            If Funcoes.VerificaPermissao("UnidadeConversao", "RELATORIO") Then
                Dim ds As New DataSet
                Dim Sql As String

                Sql = "Select UnidadeOrigem_Id as UnidadeOrigem, UnidadeDestino_Id as UnidadeDestino, Fator " & vbCrLf & _
                      "  From UnidadeConversao" & vbCrLf & _
                      "  Order by UnidadeOrigem_Id" & vbCrLf


                ds = Banco.ConsultaDataSet(Sql, "UnidadeConversao")

                Dim parameters = New Dictionary(Of String, Object)()

                Funcoes.BindReport(Me.Page, ds, "Cr_UnidadeConversao", eExportType.PDF, parameters)
            Else
                MsgBox(Me.Page, "Usuário sem permissão para emitir relatório.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "UnidadesDeMedida")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub ddlUnidadeOrigem_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddlUnidadeOrigem.SelectedIndexChanged
        If UnidadesIguais(CType(sender, DropDownList).SelectedValue, ddlUnidadeDestino.SelectedValue) Then
            MsgBox(Me.Page, "As Unidade de Origem e Destino não podem ser iguais.")
            Exit Sub
        End If
    End Sub

    Protected Sub ddlUnidadeDestino_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddlUnidadeDestino.SelectedIndexChanged
        If UnidadesIguais(ddlUnidadeDestino.SelectedValue, CType(sender, DropDownList).SelectedValue) Then
            MsgBox(Me.Page, "As Unidade de Origem e Destino não podem ser iguais.")
            Exit Sub
        End If
    End Sub

    Private Function UnidadesIguais(ByVal UnidadeOrigem As String, ByVal UnidadeDestino As String) As Boolean
        If Not String.IsNullOrWhiteSpace(UnidadeOrigem) AndAlso Not String.IsNullOrWhiteSpace(UnidadeDestino) Then
            If UnidadeOrigem.Equals(UnidadeDestino) Then
                Return True
            End If
        End If
        Return False
    End Function
End Class