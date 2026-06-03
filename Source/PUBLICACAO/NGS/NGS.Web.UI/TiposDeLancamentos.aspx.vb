Imports System.Data
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class TiposDeLancamentos
    Inherits BasePage

    Dim Sql As String
    Dim SqlArray As New ArrayList
    Dim DS As DataSet

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Gestao)
        If Not IsPostBack And IsConnect Then
            If Funcoes.VerificaPermissao("TiposDeLancamentos", "ACESSAR") Then
                CarregarTiposDeLancamentos()
                Limpar()
            Else
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Gestao.aspx")
                Exit Sub
            End If
        End If
    End Sub

    Private Sub CarregarTiposDeLancamentos()
        If Funcoes.VerificaPermissao("TiposDeLancamentos", "LEITURA") Then
            Sql = "  SELECT TipoDeLancamento_id as Codigo, Descricao " & vbCrLf & _
                  " FROM TiposDeLancamentos " & vbCrLf & _
                  " ORDER BY TipoDeLancamento_id" & vbCrLf

            GridTiposDeLancamentos.DataSource = Banco.ConsultaDataSet(Sql, "Consulta")
            GridTiposDeLancamentos.DataBind()
        End If
    End Sub

    Private Sub Limpar()
        txtCodigo.Text = ""
        txtDescricao.Text = ""
        txtCodigo.Enabled = True
        lnkNovo.Parent.Visible = True
        lnkAtualizar.Parent.Visible = False
        lnkExcluir.Parent.Visible = False
    End Sub

    Protected Sub GridTiposDeLancamentos_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            Limpar()
            txtCodigo.Text = GridTiposDeLancamentos.SelectedRow.Cells(1).Text()
            txtDescricao.Text = GridTiposDeLancamentos.SelectedRow.Cells(2).Text()
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
            If Funcoes.VerificaPermissao("TiposDeLancamentos", "GRAVAR") Then
                Sql = "INSERT Into TiposDeLancamentos(TipoDeLancamento_id, Descricao) " & vbCrLf & _
                      " Values('" & txtCodigo.Text & "' " & vbCrLf & _
                      ",'" & UCase(txtDescricao.Text) & "')" & vbCrLf
                SqlArray.Add(Sql)

                If Banco.GravaBanco(SqlArray) = False Then
                    MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                Else
                    Limpar()
                    CarregarTiposDeLancamentos()
                End If
            Else
                MsgBox(Me, "Usuário sem permissão para gravar registro.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAtualizar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkAtualizar.Click
        Try
            If Funcoes.VerificaPermissao("TiposDeLancamentos", "ALTARER") Then
                Sql = "UPDATE TiposDeLancamentos" & vbCrLf & _
                      " Set Descricao = '" & txtDescricao.Text & "' " & vbCrLf & _
                      " WHERE TipoDeLancamento_Id = '" & txtCodigo.Text & "' " & vbCrLf
                SqlArray.Add(Sql)

                If Banco.GravaBanco(SqlArray) = False Then
                    MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                Else
                    Limpar()
                    CarregarTiposDeLancamentos()
                End If
            Else
                MsgBox(Me, "Usuário sem permissão para alterar registro.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkExcluir_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkExcluir.Click
        Try
            If Funcoes.VerificaPermissao("TiposDeLancamentos", "EXCLUIR") Then
                Sql = "DELETE FROM TiposDeLancamentos" & vbCrLf & _
                      " WHERE TipoDeLancamento_Id = '" & txtCodigo.Text & "' " & vbCrLf
                SqlArray.Add(Sql)

                If Banco.GravaBanco(SqlArray) = False Then
                    MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                Else
                    Limpar()
                    CarregarTiposDeLancamentos()
                End If
            Else
                MsgBox(Me, "Usuário sem permissão para excluir registro.")
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
            If Funcoes.VerificaPermissao("TiposDeLancamentos", "RELATORIO") Then
                Sql = "Select REPLICATE('0', 3 - LEN(CAST(TipoDeLancamento_Id AS varchar))) + CAST(TipoDeLancamento_Id AS varchar) as Codigo, Descricao From TiposDeLancamentos Order by TipoDeLancamento_Id"

                DS = Banco.ConsultaDataSet(Sql, "Tabelas")

                Dim parameters = New Dictionary(Of String, Object)()
                parameters.Add("Titulo", "Tabela de Tipos De Lançamentos")
                parameters.Add("Codigo", "Código")
                parameters.Add("Descricao", "Descrição")
               
                Funcoes.BindReport(Me.Page, DS, "Cr_Tabelas", eExportType.PDF, parameters)
            Else
                MsgBox(Me.Page, "Usuário sem permissão para emitir relatório.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAjuda_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "TiposDeLancamentos")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

End Class