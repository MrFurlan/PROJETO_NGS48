Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class Solo
    Inherits BasePage

    Dim objSolo As [Lib].Negocio.Solo

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Revenda)
        If Not IsPostBack And IsConnect Then
            If Funcoes.VerificaPermissao("Solo", "ACESSAR") Then
                Limpar()
                AtualizarGrid()
            Else
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Revenda.aspx")
                Exit Sub
            End If
        End If
    End Sub

    Protected Sub gridSolo_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            txtCodigo.Enabled = False
            txtCodigo.Text = gridSolo.SelectedRow.Cells(1).Text()
            txtDescricao.Text = gridSolo.SelectedRow.Cells(2).Text()
            txtDescricao.Focus()
            lnkNovo.Parent.Visible = False
            lnkAtualizar.Parent.Visible = True
            lnkExcluir.Parent.Visible = True
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub AtualizarGrid()
        If Funcoes.VerificaPermissao("Solo", "LEITURA") Then
            Dim Lista As New [Lib].Negocio.ListSolo(True)
            gridSolo.DataSource = Lista.ToArray()
            gridSolo.DataBind()
        Else
            MsgBox(Me.Page, "Usuário sem permissão para leitura.")
        End If
    End Sub

    Private Sub Limpar()
        objSolo = New [Lib].Negocio.Solo
        txtCodigo.Text = ""
        txtDescricao.Text = ""
        txtCodigo.Enabled = True
        txtCodigo.Focus()
        lnkNovo.Parent.Visible = True
        lnkAtualizar.Parent.Visible = False
        lnkExcluir.Parent.Visible = False
    End Sub

    Protected Sub lnkNovo_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkNovo.Click
        Try
            If Funcoes.VerificaPermissao("Solo", "GRAVAR") Then
                objSolo = New [Lib].Negocio.Solo

                objSolo.Codigo = txtCodigo.Text
                objSolo.Descricao = txtDescricao.Text
                objSolo.IUD = "I"
                If objSolo.Salvar Then
                    Limpar()
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
            If Funcoes.VerificaPermissao("Solo", "ALTERAR") Then
                objSolo = New [Lib].Negocio.Solo

                objSolo.Codigo = gridSolo.SelectedRow.Cells(1).Text()
                objSolo.Descricao = txtDescricao.Text
                objSolo.IUD = "U"
                If objSolo.Salvar Then
                    Limpar()
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
            If Funcoes.VerificaPermissao("Solo", "ALTERAR") Then
                objSolo = New [Lib].Negocio.Solo

                objSolo.Codigo = gridSolo.SelectedRow.Cells(1).Text()
                objSolo.IUD = "D"
                If objSolo.Salvar Then
                    Limpar()
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

    Protected Sub lnkRelatorio_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkRelatorio.Click
        Try
            If Funcoes.VerificaPermissao("Solo", "RELATORIO") Then
                Dim ds As DataSet = getDataSet()
                Dim parameters = New Dictionary(Of String, Object)()
                parameters.Add("Titulo", "Solo")
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

    Private Function getDataSet() As DataSet
        Dim sql As String = "Select Solo_Id as Codigo, Descricao From Solo where 1=1" & vbCrLf

        If Not String.IsNullOrWhiteSpace(txtCodigo.Text) Then
            sql &= "And Solo_Id = " & txtCodigo.Text
        End If
        If Not String.IsNullOrWhiteSpace(txtDescricao.Text) Then
            sql &= "And Descricao = '" & txtDescricao.Text & "'"
        End If

        sql &= "order by Solo_Id"

        Return Banco.ConsultaDataSet(sql, "Tabelas")
    End Function

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "Solo")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub
End Class