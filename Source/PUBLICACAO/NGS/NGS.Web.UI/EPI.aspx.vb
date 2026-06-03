Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis
Public Class EPI
    Inherits BasePage

    Dim Sql As String
    Dim ds As DataSet
    Dim objEPI As [Lib].Negocio.EPI

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Gestao)
        If Not IsPostBack And IsConnect Then
            If Funcoes.VerificaPermissao("EPI", "ACESSAR") Then
                AtualizarGrid()
                Limpar()
            Else
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Gestao.aspx")
                Exit Sub
            End If
        End If
    End Sub

    Private Sub AtualizarGrid()
        If Funcoes.VerificaPermissao("EPI", "LEITURA") Then
            Dim Lista As New [Lib].Negocio.ListEPI(True)
            gridEPI.DataSource = Lista.ToArray()
            gridEPI.DataBind()

            Dim i As Integer = 0
            While i < gridEPI.Rows.Count
                If gridEPI.Rows(i).Cells(3).Text.ToUpper() = "TRUE" Then
                    gridEPI.Rows(i).Cells(3).Text = "SIM"
                Else
                    gridEPI.Rows(i).Cells(3).Text = "NÃO"
                End If

                i += 1
            End While
        End If
    End Sub

    Private Sub Limpar()
        objEPI = New [Lib].Negocio.EPI()
        txtCodigo.Text = ""
        txtDescricao.Text = ""
        lnkNovo.Parent.Visible = True
        lnkAtualizar.Parent.Visible = False
        lnkExcluir.Parent.Visible = False
    End Sub

    Protected Sub gridEPI_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        txtCodigo.Text = gridEPI.SelectedRow.Cells(1).Text()
        txtDescricao.Text = gridEPI.SelectedRow.Cells(2).Text()

        If gridEPI.SelectedRow.Cells(3).Text.ToUpper() = "SIM" Then
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
        If Funcoes.VerificaPermissao("EPI", "GRAVAR") Then
            objEPI = New [Lib].Negocio.EPI

            objEPI.Descricao = txtDescricao.Text
            objEPI.Ativo = True
            objEPI.IUD = "I"
            If objEPI.Salvar Then
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
        If Funcoes.VerificaPermissao("EPI", "ALTERAR") Then
            objEPI = New [Lib].Negocio.EPI

            objEPI.Codigo = txtCodigo.Text
            objEPI.Descricao = txtDescricao.Text
            objEPI.Ativo = True
            objEPI.IUD = "U"
            If objEPI.Salvar Then
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
        If Funcoes.VerificaPermissao("EPI", "ALTERAR") Then
            objEPI = New [Lib].Negocio.EPI

            objEPI.Codigo = txtCodigo.Text
            objEPI.Descricao = txtDescricao.Text
            objEPI.Ativo = False
            objEPI.IUD = "D"
            If objEPI.Salvar Then
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
            If Funcoes.VerificaPermissao("EPI", "RELATORIO") Then
                Sql = "Select Codigo_Id AS Codigo, Descricao from EPI" & vbCrLf
                ds = Banco.ConsultaDataSet(Sql, "Tabelas")

                Dim parameters = New Dictionary(Of String, Object)()
                parameters.Add("Titulo", "Relatório de Especificações do Produto.")
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
            Funcoes.Ajuda(Me.Page, "EPI")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub
End Class