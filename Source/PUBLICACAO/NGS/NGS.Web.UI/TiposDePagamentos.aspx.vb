Imports System.Data
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class TiposDePagamentos
    Inherits BasePage

    Private Sql As String
    Private SqlArray As New ArrayList
    Private DS As DataSet

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Gestao)
        If Not IsPostBack And IsConnect Then
            If Session("ssNomeUsuario") = "FURLAN" Then
                CarregarTiposDePagamentos()
                Limpar()
            Else
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Gestao.aspx")
                Exit Sub
            End If
        End If
    End Sub

    Private Sub CarregarTiposDePagamentos()
        If Funcoes.VerificaPermissao("TiposDePagamentos", "LEITURA") Then
            Sql = "  SELECT TipoDePagamento_id as Codigo, Descricao, EnviaAoBanco, TPagSefaz " & vbCrLf & _
                  " FROM TiposDePagamentos " & vbCrLf & _
                  " ORDER BY TipoDePagamento_id" & vbCrLf

            GridTiposDePagamentos.DataSource = Banco.ConsultaDataSet(Sql, "Consulta")
            GridTiposDePagamentos.DataBind()
        End If
    End Sub

    Private Sub Limpar()
        txtCodigo.Text = ""
        txtDescricao.Text = ""
        txtCodigoSefaz.Text = "0"
        rbSim.Checked = False
        rbNao.Checked = True
        txtCodigo.Enabled = True
        lnkNovo.Parent.Visible = True
        lnkAtualizar.Parent.Visible = False
        lnkExcluir.Parent.Visible = False
    End Sub

    Protected Sub GridTiposDePagamentos_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            Limpar()
            txtCodigo.Text = GridTiposDePagamentos.SelectedRow.Cells(1).Text()
            txtDescricao.Text = GridTiposDePagamentos.SelectedRow.Cells(2).Text()

            If GridTiposDePagamentos.SelectedRow.Cells(3).Text() = "S" Then
                rbSim.Checked = True
            Else
                rbNao.Checked = True
            End If

            txtCodigoSefaz.Text = GridTiposDePagamentos.SelectedRow.Cells(4).Text()

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
            If Funcoes.VerificaPermissao("TiposDePagamentos", "GRAVAR") Then
                Sql = "INSERT Into TiposDePagamentos(TipoDePagamento_id, Descricao, EnviaAoBanco, TPagSefaz) " & vbCrLf & _
                      " Values('" & txtCodigo.Text & "' " & vbCrLf & _
                      ",'" & UCase(txtDescricao.Text) & "'" & vbCrLf

                If rbSim.Checked Then
                    Sql &= ",'S'" & vbCrLf
                Else
                    Sql &= ",'N'" & vbCrLf
                End If

                Sql &= "," & txtCodigoSefaz.Text & ")" & vbCrLf

                SqlArray.Add(Sql)

                If Banco.GravaBanco(SqlArray) = False Then
                    Limpar()
                    CarregarTiposDePagamentos()
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
            If Funcoes.VerificaPermissao("TiposDePagamentos", "ALTERAR") Then
                Sql = "UPDATE TiposDePagamentos" & vbCrLf & _
                      " Set Descricao = '" & txtDescricao.Text & "'" & vbCrLf & _
                      ", EnviaAoBanco = '" & IIf(rbSim.Checked, "S", "N") & "'" & vbCrLf & _
                      ", TPagSefaz = " & txtCodigoSefaz.Text & vbCrLf & _
                      " WHERE TipoDePagamento_Id = '" & txtCodigo.Text & "' " & vbCrLf
                SqlArray.Add(Sql)

                If Banco.GravaBanco(SqlArray) = False Then
                    Limpar()
                    CarregarTiposDePagamentos()
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
            If Funcoes.VerificaPermissao("TiposDePagamentos", "EXCLUIR") Then
                Sql = "DELETE FROM TiposDePagamentos" & vbCrLf & _
                      " WHERE TipoDePagamento_Id = '" & txtCodigo.Text & "' " & vbCrLf
                SqlArray.Add(Sql)

                If Banco.GravaBanco(SqlArray) = False Then
                    Limpar()
                    CarregarTiposDePagamentos()
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
            If Funcoes.VerificaPermissao("TiposDePagamentos", "RELATORIO") Then
                Sql = "Select REPLICATE('0', 3 - LEN(CAST(TipoDePagamento_Id AS varchar))) + CAST(TipoDePagamento_Id AS varchar) as Codigo, Descricao From TiposDePagamentos Order by TipoDePagamento_Id"

                DS = Banco.ConsultaDataSet(Sql, "Tabelas")

                Dim parameters = New Dictionary(Of String, Object)()
                parameters.Add("Titulo", "Tabela de Tipos De Pagamentos")
                parameters.Add("Codigo", "Código")
                parameters.Add("Descricao", "Descrição")
               
                Funcoes.BindReport(Me.Page, DS, "Cr_Tabelas", eExportType.PDF, parameters)
            Else
                MsgBox(Me.Page, "Usuário sem permisão para emitir relatório.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAjuda_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "TiposDePagamentos")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

End Class