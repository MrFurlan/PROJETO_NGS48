Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class EstadoFisicoIA
    Inherits BasePage

    Dim Sql As String
    Dim ds As DataSet
    Dim objEstadoFisicoIA As [Lib].Negocio.EstadoFisicoIA

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Revenda)
        If Not IsPostBack And IsConnect Then
            If Funcoes.VerificaPermissao("EstadoFisicoIA", "ACESSAR") Then
                AtualizarGrid()
                Limpar()
            Else
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Revenda.aspx")
                Exit Sub
            End If
        End If
    End Sub

    Protected Sub gridEstadoFisicoIA_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            txtCodigo.Enabled = False
            txtCodigo.Text = gridEstadoFisicoIA.SelectedRow.Cells(1).Text()
            txtDescricao.Text = gridEstadoFisicoIA.SelectedRow.Cells(2).Text()
            txtDescricao.Focus()
            lnkNovo.Parent.Visible = False
            lnkAtualizar.Parent.Visible = True
            lnkExcluir.Parent.Visible = True
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub AtualizarGrid()
        Dim Lista As New [Lib].Negocio.ListEstadoFisicoIA(True)
        gridEstadoFisicoIA.DataSource = Lista.ToArray()
        gridEstadoFisicoIA.DataBind()
    End Sub

    Private Sub Limpar()
        objEstadoFisicoIA = New [Lib].Negocio.EstadoFisicoIA
        txtCodigo.Text = ""
        txtDescricao.Text = ""
        txtCodigo.Enabled = True
        txtCodigo.Focus()
        lnkNovo.Parent.Visible = True
        lnkAtualizar.Parent.Visible = False
        lnkExcluir.Parent.Visible = False
    End Sub

    Private Function getParameters() As String
        Dim param As String = ""
        If Not String.IsNullOrWhiteSpace(txtCodigo.Text) Then
            param &= "Código: " & txtCodigo.Text & vbCrLf
        End If
        If Not String.IsNullOrWhiteSpace(txtDescricao.Text) Then
            param &= "Descrição: " & txtDescricao.Text & "."
        End If

        Return param
    End Function

    Protected Sub lnkNovo_Click(sender As Object, e As EventArgs) Handles lnkNovo.Click
        Try
            If Funcoes.VerificaPermissao("EstadoFisicoIA", "GRAVAR") Then
                objEstadoFisicoIA = New [Lib].Negocio.EstadoFisicoIA

                objEstadoFisicoIA.Codigo = txtCodigo.Text
                objEstadoFisicoIA.Descricao = txtDescricao.Text
                objEstadoFisicoIA.IUD = "I"
                If objEstadoFisicoIA.Salvar Then
                    MsgBox(Me.Page, "Registro incluso com sucesso.", eTitulo.Sucess)
                    AtualizarGrid()
                    Limpar()
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para gravar registro.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAtualizar_Click(sender As Object, e As EventArgs) Handles lnkAtualizar.Click
        Try
            If Funcoes.VerificaPermissao("EstadoFisicoIA", "ALTERAR") Then
                objEstadoFisicoIA = New [Lib].Negocio.EstadoFisicoIA

                objEstadoFisicoIA.Codigo = gridEstadoFisicoIA.SelectedRow.Cells(1).Text()
                objEstadoFisicoIA.Descricao = txtDescricao.Text
                objEstadoFisicoIA.IUD = "U"
                If objEstadoFisicoIA.Salvar Then
                    MsgBox(Me.Page, "Registro Alterado com sucesso.", eTitulo.Sucess)
                    AtualizarGrid()
                    Limpar()
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para alterar registro.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkExcluir_Click(sender As Object, e As EventArgs) Handles lnkExcluir.Click
        Try
            If Funcoes.VerificaPermissao("EstadoFisicoIA", "EXCLUIR") Then
                objEstadoFisicoIA = New [Lib].Negocio.EstadoFisicoIA

                objEstadoFisicoIA.Codigo = gridEstadoFisicoIA.SelectedRow.Cells(1).Text()
                objEstadoFisicoIA.IUD = "D"
                If objEstadoFisicoIA.Salvar Then
                    MsgBox(Me.Page, "Registro excluso com sucesso.", eTitulo.Sucess)
                    AtualizarGrid()
                    Limpar()
                End If
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
            If Funcoes.VerificaPermissao("EstadoFisicoIA", "RELATORIO") Then
                Sql = "Select EstadoFisicoIA_Id AS Codigo, Descricao From EstadoFisicoIA " & vbCrLf
                ds = Banco.ConsultaDataSet(Sql, "Tabelas")

                Dim parameters = New Dictionary(Of String, Object)()
                parameters.Add("Titulo", "Relatório De Estado Físico Ingrediente Ativo.")
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

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "EstadoFisicoIA")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub
End Class