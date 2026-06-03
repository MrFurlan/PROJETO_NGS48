Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis
Public Class EspecificacaoDoProduto
    Inherits BasePage

    Dim Sql As String
    Dim ds As DataSet
    Dim objEP As [Lib].Negocio.EspecificacaoDoProduto

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Gestao)
        If Not IsPostBack And IsConnect Then
            If Funcoes.VerificaPermissao("EspecificacaoDoProduto", "ACESSAR") Then
                AtualizarGrid()
                Limpar()
            Else
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Gestao.aspx")
                Exit Sub
            End If
        End If
    End Sub

    Private Sub AtualizarGrid()
        If Funcoes.VerificaPermissao("EspecificacaoDoProduto", "LEITURA") Then
            Dim Lista As New [Lib].Negocio.ListEspecificacaoDoProduto()
            gridEspecificacaoDoProduto.DataSource = Lista.OrderBy(Function(s) s.Descricao).ToList()
            gridEspecificacaoDoProduto.DataBind()

            Dim i As Integer = 0
            While i < gridEspecificacaoDoProduto.Rows.Count
                If gridEspecificacaoDoProduto.Rows(i).Cells(3).Text.ToUpper() = "TRUE" Then
                    gridEspecificacaoDoProduto.Rows(i).Cells(3).Text = "SIM"
                Else
                    gridEspecificacaoDoProduto.Rows(i).Cells(3).Text = "NÃO"
                End If

                i += 1
            End While
        End If
    End Sub

    Private Sub Limpar()
        objEP = New [Lib].Negocio.EspecificacaoDoProduto()
        txtCodigo.Text = ""
        txtDescricao.Text = ""
        lnkNovo.Parent.Visible = True
        lnkAtualizar.Parent.Visible = False
        lnkExcluir.Parent.Visible = False
    End Sub

    Protected Sub gridEspecificacaoDoProduto_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        txtCodigo.Text = gridEspecificacaoDoProduto.SelectedRow.Cells(1).Text()
        txtDescricao.Text = gridEspecificacaoDoProduto.SelectedRow.Cells(2).Text()

        If gridEspecificacaoDoProduto.SelectedRow.Cells(3).Text.ToUpper() = "SIM" Then
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
        If Funcoes.VerificaPermissao("EspecificacaoDoProduto", "GRAVAR") Then

            If Trim(txtDescricao.Text).Length = 0 Then
                MsgBox(Me.Page, "Descrição da Especificação é obrigatória.")
                Exit Sub
            End If

            objEP = New [Lib].Negocio.EspecificacaoDoProduto

            objEP.Descricao = LTrim(txtDescricao.Text)
            objEP.Descricao = RTrim(objEP.Descricao)
            objEP.Ativo = True
            objEP.IUD = "I"
            If objEP.Salvar Then
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
        If Funcoes.VerificaPermissao("EspecificacaoDoProduto", "ALTERAR") Then

            If Trim(txtDescricao.Text).Length = 0 Then
                MsgBox(Me.Page, "Descrição da Especificação é obrigatória.")
                Exit Sub
            End If

            objEP = New [Lib].Negocio.EspecificacaoDoProduto

            objEP.Codigo = txtCodigo.Text
            objEP.Descricao = LTrim(txtDescricao.Text)
            objEP.Descricao = RTrim(objEP.Descricao)
            objEP.Ativo = True
            objEP.IUD = "U"
            If objEP.Salvar Then
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
        If Funcoes.VerificaPermissao("EspecificacaoDoProduto", "EXCLUIR") Then
            objEP = New [Lib].Negocio.EspecificacaoDoProduto

            objEP.Codigo = txtCodigo.Text
            objEP.Descricao = txtDescricao.Text
            objEP.Ativo = False
            objEP.IUD = "D"
            If objEP.Salvar Then
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
            If Funcoes.VerificaPermissao("EspecificacaoDoProduto", "RELATORIO") Then
                Sql = "Select Codigo_Id AS Codigo, Descricao from EspecificacaoDoProduto" & vbCrLf
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
            Funcoes.Ajuda(Me.Page, "EspecificacaoDoProduto")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub
End Class