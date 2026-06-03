Imports System.Data
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class ClassesDeOperacoes
    Inherits BasePage

    Dim SqlArray As New ArrayList
    Dim ds As DataSet
    Dim sql As String

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Gestao)
        If Not IsPostBack And IsConnect Then
            If Funcoes.VerificaPermissao("ClassesDeOperacoes", "ACESSAR") AndAlso Request.Url.Host.ToUpper.Contains("LOCALHOST") Then
                CarregarClassesDeOperacoes()
                CarregarClassificacoesDeFrete()
                Limpar()
            Else
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Gestao.aspx")
                Exit Sub
            End If
        End If
    End Sub

    Private Sub CarregarClassesDeOperacoes()
        If Funcoes.VerificaPermissao("ClassesDeOperacoes", "LEITURA") Then
            sql = "SELECT Classe_id as Codigo, Descricao, isnull(Operacao,0) as Operacao, isnull(SubOperacao,0) as SubOperacao, ClassificacaoDeFrete " & vbCrLf & _
                  "  FROM ClassesDeOperacoes " & vbCrLf & _
                  " ORDER BY Classe_id" & vbCrLf
            GridClassesDeOperacoes.DataSource = Banco.ConsultaDataSet(sql, "Consulta")
            GridClassesDeOperacoes.DataBind()
        Else
            MsgBox(Me.Page, "Usuário sem permissão para leitura registro.", eTitulo.Info)
        End If
    End Sub

    Private Sub CarregarClassificacoesDeFrete()
        ddlClassificacaoFrete.Items.Clear()
        ddlClassificacaoFrete.Items.Add(New ListItem("", ""))
        ddlClassificacaoFrete.Items.Add(New ListItem("COMPRAS", "COMPRAS"))
        ddlClassificacaoFrete.Items.Add(New ListItem("VENDAS", "VENDAS"))
        ddlClassificacaoFrete.Items.Add(New ListItem("DEPÓSITOS", "DEPOSITOS"))
        ddlClassificacaoFrete.Items.Add(New ListItem("TRANSFERÊNCIAS", "TRANSFERENCIAS"))
        ddlClassificacaoFrete.Items.Add(New ListItem("DOACOES/BONIFICACOES", "DOACAO"))
        ddlClassificacaoFrete.SelectedIndex = 0
    End Sub

    Private Sub Limpar()
        txtCodigo.Text = ""
        txtDescricao.Text = ""
        txtCodigo.Enabled = True
        chkOperacao.Checked = False
        chkSubOperacao.Checked = False
        ddlClassificacaoFrete.SelectedIndex = 0
        lnkNovo.Parent.Visible = True
        lnkAtualizar.Parent.Visible = False
        lnkExcluir.Parent.Visible = False
    End Sub

    Protected Sub GridClassesDeOperacoes_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            Limpar()
            txtCodigo.Text = GridClassesDeOperacoes.SelectedRow.Cells(1).Text()
            txtDescricao.Text = GridClassesDeOperacoes.SelectedRow.Cells(2).Text()
            chkOperacao.Checked = CType(GridClassesDeOperacoes.SelectedRow.Cells(3).FindControl("CheckBox1"), CheckBox).Checked
            chkSubOperacao.Checked = CType(GridClassesDeOperacoes.SelectedRow.Cells(4).FindControl("CheckBox2"), CheckBox).Checked
            If (Not String.IsNullOrWhiteSpace(GridClassesDeOperacoes.SelectedRow.Cells(5).Text())) AndAlso GridClassesDeOperacoes.SelectedRow.Cells(5).Text().Trim() <> "&nbsp;" Then
                ddlClassificacaoFrete.SelectedValue = GridClassesDeOperacoes.SelectedRow.Cells(5).Text().Trim()
            End If
            txtCodigo.Enabled = False
            txtDescricao.Focus()
            lnkNovo.Parent.Visible = False
            lnkAtualizar.Parent.Visible = True
            lnkExcluir.Parent.Visible = True
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkNovo_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkNovo.Click
        Try
            If Funcoes.VerificaPermissao("ClassesDeOperacoes", "GRAVAR") Then
                sql = "INSERT Into ClassesDeOperacoes(Classe_id, Descricao, Operacao, SubOperacao, ClassificacaoDeFrete) " & vbCrLf & _
                      " Values('" & txtCodigo.Text & "','" & UCase(txtDescricao.Text) & "'," & CByte(chkOperacao.Checked) & "," & CByte(chkSubOperacao.Checked) & ", " & IIf(String.IsNullOrWhiteSpace(ddlClassificacaoFrete.SelectedValue), " null ", "'" & ddlClassificacaoFrete.SelectedValue & "'") & ")"
                SqlArray.Add(sql)
                If Banco.GravaBanco(SqlArray) = False Then

                    MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                Else
                    Limpar()
                    CarregarClassesDeOperacoes()
                    CarregarClassificacoesDeFrete()
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
            If Funcoes.VerificaPermissao("ClassesDeOperacoes", "ALTERAR") Then
                sql = "UPDATE ClassesDeOperacoes Set" & vbCrLf & _
                      "   Descricao   ='" & txtDescricao.Text & "'" & vbCrLf & _
                      "  ,Operacao    = " & CByte(chkOperacao.Checked) & vbCrLf & _
                      "  ,SubOperacao = " & CByte(chkSubOperacao.Checked) & vbCrLf & _
                      "  ,ClassificacaoDeFrete = " & IIf(String.IsNullOrWhiteSpace(ddlClassificacaoFrete.SelectedValue), " null ", "'" & ddlClassificacaoFrete.SelectedValue & "'") & vbCrLf & _
                      " WHERE Classe_Id = '" & txtCodigo.Text & "' "
                SqlArray.Add(sql)
                If Banco.GravaBanco(SqlArray) = False Then

                    MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                Else
                    Limpar()
                    CarregarClassesDeOperacoes()
                    CarregarClassificacoesDeFrete()
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
            If Funcoes.VerificaPermissao("ClassesDeOperacoes", "EXCLUIR") Then
                sql = "DELETE FROM ClassesDeOperacoes" & vbCrLf & _
                      " WHERE Classe_Id = '" & txtCodigo.Text & "' " & vbCrLf
                SqlArray.Add(sql)
                If Banco.GravaBanco(SqlArray) = False Then

                    MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                Else
                    Limpar()
                    CarregarClassesDeOperacoes()
                    CarregarClassificacoesDeFrete()
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

    Protected Sub lnkRelatorio_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkRelatorio.Click
        Try
            If Funcoes.VerificaPermissao("ClassesDeOperacoes", "RELATORIO") Then
                sql = "Select CAST(Classe_Id AS varchar) as Codigo, Descricao From ClassesDeOperacoes Order by Classe_Id"
                ds = Banco.ConsultaDataSet(sql, "Tabelas")

                Dim parameters = New Dictionary(Of String, Object)()
                parameters.Add("Titulo", "Relatório De Classes De Operações.")
                parameters.Add("Codigo", "Código")
                parameters.Add("Descricao", "Descrição")
               
                Funcoes.BindReport(Me.Page, ds, "Cr_Tabelas", eExportType.PDF, parameters)
            Else
                MsgBox(Me.Page, "Usuário sem permissão para emitir relatório.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAjuda_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "ClassesDeOperacoes")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub

End Class