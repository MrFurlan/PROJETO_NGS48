Imports System.Data
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class CodigoBeneficioFiscal
    Inherits BasePage

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.setMenu(eModulo.Gestao)
            If Not IsPostBack And IsConnect Then
                If Funcoes.VerificaPermissao("CodigoBeneficioFiscal", "ACESSAR") Then
                    ddl.Carregar(ddlEstado, CarregarDDL.Tabela.Estados, "", True)

                    Limpar()

                    Dim objEmpresa = New Cliente(HttpContext.Current.Session("ssEmpresa"), HttpContext.Current.Session("ssEndEmpresa"))

                    ddlEstado.SelectedValue = objEmpresa.Estado.Codigo

                    carregarGrid()
                Else
                    MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Gestao.aspx")
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try

    End Sub

    Private Sub carregarGrid()

        If ddlEstado.SelectedValue.Length > 0 Then
            Dim lst As New ListAjustesDaApuracaoIcms(ddlEstado.SelectedValue)
            GridAjustesDaApuracaoIcms.DataSource = lst
        Else
            GridAjustesDaApuracaoIcms.DataSource = Nothing
        End If

        GridAjustesDaApuracaoIcms.DataBind()

    End Sub

    Private Sub Limpar()
        txtCodigo.Enabled = True
        txtDescricao.Enabled = True
        txtCodigo.Text = ""
        txtDescricao.Text = ""
        lnkNovo.Parent.Visible = True
        lnkAtualizar.Parent.Visible = False
        lnkExcluir.Parent.Visible = False
        ddlEstado.Enabled = True
    End Sub

    Private Function validaRegistro() As Boolean
        If String.IsNullOrWhiteSpace(txtCodigo.Text) Then
            MsgBox(Me.Page, "Código obrigatório.")
            txtCodigo.Focus()
            Return False
        ElseIf String.IsNullOrWhiteSpace(ddlEstado.SelectedValue) Then
            MsgBox(Me.Page, "Estado obrigatório")
            ddlEstado.Focus()
            Return False
        ElseIf String.IsNullOrWhiteSpace(txtDescricao.Text) Then
            MsgBox(Me.Page, "Descrição obrigatória.")
            txtDescricao.Focus()
            Return False
        ElseIf txtCodigo.Text.Length = 8 AndAlso ddlEstado.SelectedValue = "SC" Then
            Return True
        ElseIf Not txtCodigo.Text.Length = 6 Then
            MsgBox(Me.Page, "Informar código com 6 caracteres ou 8 se for SC, inclusive 0(zero) a esquerda se necessário. Ex.: 000001 OR 00000001.")
            txtCodigo.Focus()
            Return False
        Else
            Return True
        End If
    End Function

    Protected Sub gridAjustesDaApuracaoIcms_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles gridAjustesDaApuracaoIcms.SelectedIndexChanged
        Try

            Dim cod As New AjustesDaApuracaoIcms(gridAjustesDaApuracaoIcms.SelectedRow.Cells(1).Text)

            If ddlEstado.SelectedValue = "SC" Then
                txtCodigo.Text = Mid(cod.Codigo, 3, 8)
            Else
                txtCodigo.Text = Mid(cod.Codigo, 3, 6)
            End If

            txtDescricao.Text = cod.Descricao

            ddlEstado.Enabled = False
            txtCodigo.Enabled = False

            lnkNovo.Parent.Visible = False

            If cod.CodigoUtilizado Then
                txtDescricao.Enabled = False
                lnkAtualizar.Parent.Visible = False
                lnkExcluir.Parent.Visible = False
                MsgBox(Me.Page, "Código de Benefício Fiscal utilizado não pode ser alterado/excluído.", eTitulo.Sucess)
            Else
                txtDescricao.Enabled = True
                lnkAtualizar.Parent.Visible = True
                lnkExcluir.Parent.Visible = True
            End If

        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkNovo_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkNovo.Click
        Try
            If Funcoes.VerificaPermissao("CodigoBeneficioFiscal", "GRAVAR") Then
                If validaRegistro() Then

                    Dim cod As String = ddlEstado.SelectedValue & Trim(txtCodigo.Text)

                    Dim aIcms As New AjustesDaApuracaoIcms()
                    aIcms.IUD = "I"
                    aIcms.Codigo = cod
                    aIcms.Descricao = txtDescricao.Text.Trim

                    If aIcms.Salvar Then
                        MsgBox(Me.Page, "Código de Benefício Fiscal salvo com Sucesso.", eTitulo.Sucess)
                        Limpar()
                        carregarGrid()
                    Else
                        MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                    End If

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
            If Funcoes.VerificaPermissao("CodigoBeneficioFiscal", "ALTERAR") Then
                If validaRegistro() Then
                    Dim cod As String = ddlEstado.SelectedValue & Trim(txtCodigo.Text)

                    Dim aIcms As New AjustesDaApuracaoIcms()
                    aIcms.IUD = "U"
                    aIcms.Codigo = cod
                    aIcms.Descricao = txtDescricao.Text.Trim

                    If aIcms.Salvar Then
                        MsgBox(Me.Page, "Código de Benefício Fiscal atualizado com Sucesso.", eTitulo.Sucess)
                        Limpar()
                        carregarGrid()
                    Else
                        MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                    End If

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
            If Funcoes.VerificaPermissao("CodigoBeneficioFiscal", "EXCLUIR") Then
                If validaRegistro() Then
                    Dim cod As String = ddlEstado.SelectedValue & Trim(txtCodigo.Text)

                    Dim aIcms As New AjustesDaApuracaoIcms()
                    aIcms.IUD = "D"
                    aIcms.Codigo = cod
                    aIcms.Descricao = txtDescricao.Text.Trim

                    If aIcms.Salvar Then
                        MsgBox(Me.Page, "Código de Benefício Fiscal atualizado com Sucesso.", eTitulo.Sucess)
                        Limpar()
                        carregarGrid()
                    Else
                        MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                    End If

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
            If Funcoes.VerificaPermissao("CodigoBeneficioFiscal", "RELATORIO") Then

                If ddlEstado.SelectedValue.Length = 0 Then
                    MsgBox(Me.Page, "Estado obrigatório")
                    ddlEstado.Focus()
                    Exit Sub
                End If

                Dim ds As DataSet

                Dim Sql = "Select Codigo_Id as Codigo, Descricao" & vbCrLf & _
                            "From AjustesDaApuracaoIcms" & vbCrLf & _
                            "where left(Codigo_Id,2) = '" & ddlEstado.SelectedValue & "'" & vbCrLf & _
                            "Order by Codigo_Id"

                ds = Banco.ConsultaDataSet(Sql, "Tabelas")

                Dim Campo As String() = HttpContext.Current.Server.MapPath("").Split("\")
                Dim RootSite As String = ""
                For i As Integer = 0 To Campo.GetUpperBound(0) - 1
                    RootSite &= Campo(i) & "\"
                Next

                Dim parameters = New Dictionary(Of String, Object)()
                parameters.Add("Titulo", "Código de Benefício Fiscal")
                parameters.Add("Codigo", "Código")
                parameters.Add("Descricao", "Descrição")
                Funcoes.BindReport(Me.Page, ds, "Cr_Tabelas", eExportType.PDF, parameters)
            Else
                MsgBox(Me.Page, "Usuario sem permissão para emitir relatório.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAjuda_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "CodigoBeneficioFiscal")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Erro)
        End Try
    End Sub

    Protected Sub ddlEstado_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddlEstado.SelectedIndexChanged
        Try
            carregarGrid()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

End Class