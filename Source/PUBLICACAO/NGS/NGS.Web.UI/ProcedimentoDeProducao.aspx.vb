Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis
Public Class ProcedimentoDeProducao
    Inherits BasePage

    Dim Sql As String
    Dim ds As DataSet
    Dim objProcedimentoDeProducao As ProcedimentoParaProducao

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Gestao)
        If Not IsPostBack And IsConnect Then
            If Funcoes.VerificaPermissao("ProcedimentoDeProducao", "ACESSAR") Then
                AtualizarGrid()
                Limpar()
            Else
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Gestao.aspx")
                Exit Sub
            End If
        End If
    End Sub

    Private Sub AtualizarGrid()
        If Funcoes.VerificaPermissao("ProcedimentoDeProducao", "LEITURA") Then
            Dim Lista As New ListProcedimentoParaProducao(True)
            gridProcedimentoDeProducao.DataSource = Lista
            gridProcedimentoDeProducao.DataBind()

            Dim i As Integer = 0
            While i < gridProcedimentoDeProducao.Rows.Count
                If gridProcedimentoDeProducao.Rows(i).Cells(3).Text.ToUpper() = "TRUE" Then
                    gridProcedimentoDeProducao.Rows(i).Cells(3).Text = "SIM"
                Else
                    gridProcedimentoDeProducao.Rows(i).Cells(3).Text = "NÃO"
                End If

                i += 1
            End While
        End If
    End Sub

    Private Sub Limpar()
        objProcedimentoDeProducao = New ProcedimentoParaProducao()
        txtCodigo.Text = ""
        txtDescricao.Text = ""
        lnkNovo.Parent.Visible = True
        lnkAtualizar.Parent.Visible = False
        lnkExcluir.Parent.Visible = False
    End Sub

    Protected Sub gridProcedimentoDeProducao_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        txtCodigo.Text = gridProcedimentoDeProducao.SelectedRow.Cells(1).Text()
        txtDescricao.Text = gridProcedimentoDeProducao.SelectedRow.Cells(2).Text()

        If gridProcedimentoDeProducao.SelectedRow.Cells(3).Text.ToUpper() = "SIM" Then
            chkAtivo.Checked = True
        Else
            chkAtivo.Checked = False
        End If

        lnkNovo.Parent.Visible = False
        lnkAtualizar.Parent.Visible = True
        lnkExcluir.Parent.Visible = True

        txtDescricao.Focus()
    End Sub

    Protected Sub lnkNovo_Click(sender As Object, e As EventArgs) Handles lnkNovo.Click
        If Funcoes.VerificaPermissao("ProcedimentoDeProducao", "GRAVAR") Then
            objProcedimentoDeProducao = New ProcedimentoParaProducao()

            objProcedimentoDeProducao.Descricao = txtDescricao.Text.ToUpper()
            objProcedimentoDeProducao.Ativo = True
            objProcedimentoDeProducao.IUD = "I"
            If objProcedimentoDeProducao.Salvar Then
                Limpar()
                AtualizarGrid()
            Else
                MsgBox(Me.Page, HttpContext.Current.Session("ssMessage").ToString)
            End If
        Else
            MsgBox(Me.Page, "Usuário sem permissão para gravar registro.")
        End If
    End Sub

    Protected Sub lnkAtualizar_Click(sender As Object, e As EventArgs) Handles lnkAtualizar.Click
        If Funcoes.VerificaPermissao("ProcedimentoDeProducao", "ALTERAR") Then
            objProcedimentoDeProducao = New ProcedimentoParaProducao()

            objProcedimentoDeProducao.Codigo = txtCodigo.Text
            objProcedimentoDeProducao.Descricao = txtDescricao.Text.ToUpper()
            objProcedimentoDeProducao.Ativo = True
            objProcedimentoDeProducao.IUD = "U"
            If objProcedimentoDeProducao.Salvar Then
                Limpar()
                AtualizarGrid()
            Else
                MsgBox(Me.Page, HttpContext.Current.Session("ssMessage").ToString)
            End If
        Else
            MsgBox(Me.Page, "Usuário sem permissão para alterar registro.")
        End If
    End Sub

    Protected Sub lnkExcluir_Click(sender As Object, e As EventArgs) Handles lnkExcluir.Click
        If Funcoes.VerificaPermissao("ProcedimentoDeProducao", "ALTERAR") Then
            objProcedimentoDeProducao = New ProcedimentoParaProducao()

            objProcedimentoDeProducao.Codigo = txtCodigo.Text
            objProcedimentoDeProducao.Descricao = txtDescricao.Text
            objProcedimentoDeProducao.Ativo = False
            objProcedimentoDeProducao.IUD = "D"
            If objProcedimentoDeProducao.Salvar Then
                Limpar()
                AtualizarGrid()
            Else
                MsgBox(Me.Page, HttpContext.Current.Session("ssMessage").ToString)
            End If
        Else
            MsgBox(Me.Page, "Usuário sem permissão para alterar registro.")
        End If
    End Sub

    Protected Sub lnkLimpar_Click(sender As Object, e As EventArgs) Handles lnkLimpar.Click
        Limpar()
    End Sub

    Protected Sub lnkRelatorio_Click(sender As Object, e As EventArgs) Handles lnkRelatorio.Click
        Try
            If Funcoes.VerificaPermissao("ProcedimentoDeProducao", "RELATORIO") Then
                Sql = "Select Codigo_Id AS Codigo, Descricao from ProcedimentoDeProducao" & vbCrLf
                ds = Banco.ConsultaDataSet(Sql, "Tabelas")

                Dim parameters = New Dictionary(Of String, Object)()
                parameters.Add("Titulo", "Relatório de Procedimentos para Produção.")
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
            Funcoes.Ajuda(Me.Page, "ProcedimentoDeProducao")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub
End Class