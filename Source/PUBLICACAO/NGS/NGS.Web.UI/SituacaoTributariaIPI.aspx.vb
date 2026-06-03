
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis
Imports NGS.Lib
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared

Public Class SituacaoTributariaIPI
    Inherits BasePage

    Dim Sql As String
    Dim ds As DataSet

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Gestao)
        If Not IsPostBack And IsConnect Then
            If Funcoes.VerificaPermissao("SituacaoTributariaIPI", "ACESSAR") Then
                CarregarSTIPI()
                Limpar()
            Else
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Gestao.aspx")
                Exit Sub
            End If
        End If
    End Sub

    Private Sub CarregarSTIPI()
        If Funcoes.VerificaPermissao("SituacaoTributariaIPI", "LEITURA") Then
            Dim objSituacaoTributariaIPI As New [Lib].Negocio.ListSituacaoTributariaIPI(True)
            GridSTIpi.DataSource = objSituacaoTributariaIPI.ToArray
            GridSTIpi.DataBind()
        End If
    End Sub

    Private Sub Limpar()
        Me.txtCodigo.Text = ""
        Me.txtDescricao.Text = ""
        Me.txtCodigo.Enabled = True
        lnkNovo.Parent.Visible = True
        lnkAtualizar.Parent.Visible = False
        lnkExcluir.Parent.Visible = False
    End Sub

    Function ValidarCampos() As Boolean
        If String.IsNullOrWhiteSpace(txtCodigo.Text) Then
            MsgBox(Me.Page, "Código da Situação Tributária do IPI não foi informado!")
            Return False
        ElseIf String.IsNullOrWhiteSpace(txtDescricao.Text) Then
            MsgBox(Me.Page, "Descrição da Situação Tributária do IPI não foi informada!")
            Return False
        Else
            Return True
        End If
    End Function

    Protected Sub GridSTIPI_SelectedIndexChanged(ByVal sender As Object, ByVal e As EventArgs) Handles GridSTIpi.SelectedIndexChanged
        Try
            txtCodigo.Text = GridSTIpi.SelectedRow.Cells(1).Text().Trim
            Dim ObjSituacaoTributariaIPI As New [Lib].Negocio.SituacaoTributariaIPI(txtCodigo.Text.Trim)
            If (ObjSituacaoTributariaIPI IsNot Nothing) Then
                txtDescricao.Text = ObjSituacaoTributariaIPI.Descricao
                txtCodigo.Enabled = False
                txtCodigo.Focus()
            End If
            Limpar()
            lnkNovo.Parent.Visible = False
            lnkAtualizar.Parent.Visible = True
            lnkExcluir.Parent.Visible = True
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkNovo_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkNovo.Click
        Try
            If Funcoes.VerificaPermissao("SituacaoTributariaIPI", "GRAVAR") Then
                If ValidarCampos() Then
                    Dim objSituacaoTributariaIPI As New [Negocio].SituacaoTributariaIPI()
                    objSituacaoTributariaIPI.Codigo = txtCodigo.Text.Trim
                    objSituacaoTributariaIPI.Descricao = txtDescricao.Text.Trim
                    objSituacaoTributariaIPI.IUD = "I"
                    If objSituacaoTributariaIPI.Salvar Then
                        Limpar()
                        CarregarSTIPI()
                    Else
                        MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                    End If
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissao para gravar registro.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAtualizar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkAtualizar.Click
        Try
            If Funcoes.VerificaPermissao("SituacaoTributariaIPI", "ALTERAR") Then
                If ValidarCampos() Then

                    Dim objSituacaoTributariaIPI As New [Negocio].SituacaoTributariaIPI()
                    objSituacaoTributariaIPI.Codigo = txtCodigo.Text.Trim
                    objSituacaoTributariaIPI.Descricao = txtDescricao.Text.Trim

                    objSituacaoTributariaIPI.IUD = "U"

                    If objSituacaoTributariaIPI.Salvar Then
                        Limpar()
                        CarregarSTIPI()
                    Else
                        MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                    End If
                End If
            Else
                MsgBox(Me.Page, "Usuario sem permissao para alterar registro.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkExcluir_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkExcluir.Click
        Try
            If Funcoes.VerificaPermissao("Estados", "EXCLUIR") Then
                If ValidarCampos() Then

                    Dim objSituacaoTributariaIPI As New [Negocio].SituacaoTributariaIPI(txtCodigo.Text.Trim)

                    objSituacaoTributariaIPI.Descricao = txtDescricao.Text.Trim

                    objSituacaoTributariaIPI.IUD = "D"
                    If objSituacaoTributariaIPI.Salvar Then
                        Limpar()
                        CarregarSTIPI()
                    Else
                        MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                    End If
                End If

            Else
                MsgBox(Me.Page, "Usuario sem permissao para excluir registro.")
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
            If Funcoes.VerificaPermissao("SituacaoTributariaIPI", "RELATORIO") Then
                Dim Titulo As String = "Tabela de Situações Tributárias do IPI"
                Dim Codigo As String = "Código"
                Dim Descricao As String = "Descrição"

                Titulo = "Tabela de Situações Tributárias do IPI"
                Sql = "Select REPLICATE('0', 3 - LEN(CAST(SituacaoTributariaIPI_Id AS varchar))) + CAST(SituacaoTributariaIPI_Id AS varchar) as Codigo, Descricao From SituacaoTributariaIPI Order by SituacaoTributariaIPI_Id"
                ds = Banco.ConsultaDataSet(Sql, "Tabelas")

                Dim parameters = New Dictionary(Of String, Object)()
                parameters.Add("Titulo", Titulo)
                parameters.Add("Codigo", Codigo)
                parameters.Add("Descricao", Descricao)

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
            Funcoes.Ajuda(Me.Page, "SituacaoTributariaIPI")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub
End Class