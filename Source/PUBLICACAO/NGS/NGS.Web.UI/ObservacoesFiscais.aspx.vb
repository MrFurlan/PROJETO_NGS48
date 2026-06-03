Imports System.Data
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class ObservacoesFiscais
    Inherits BasePage

    Dim ds As DataSet
    Dim objObservacaoFiscal As ObservacaoFiscal

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.setMenu(eModulo.Fiscal)
            If Not IsPostBack And IsConnect Then
                If Funcoes.VerificaPermissao("ObservacoesFiscais", "ACESSAR") Then
                    CarregarObservacoesFiscais()
                    Limpar()
                Else
                    MsgBox(Me.Page, "Usuário sem permissão para acessar essa página", "~/Fiscal.aspx")
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub CarregarObservacoesFiscais()
        If Funcoes.VerificaPermissao("ObservacoesFiscais", "LEITURA") Then
            GridObservacoesFiscais.DataSource = New ListObservacaoFiscal()
            GridObservacoesFiscais.DataBind()
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

    Protected Sub GridObservacoesFiscais_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            Limpar()
            objObservacaoFiscal = New ObservacaoFiscal(GridObservacoesFiscais.SelectedRow.Cells(1).Text())
            txtCodigo.Text = objObservacaoFiscal.Codigo
            txtDescricao.Text = objObservacaoFiscal.Descricao
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
            If Funcoes.VerificaPermissao("ObservacoesFiscais", "GRAVAR") Then
                If ValidarCampos() Then
                    objObservacaoFiscal = New ObservacaoFiscal()
                    objObservacaoFiscal.IUD = "I"
                    objObservacaoFiscal.Codigo = txtCodigo.Text
                    objObservacaoFiscal.Descricao = txtDescricao.Text
                    If objObservacaoFiscal.Salvar Then
                        Limpar()
                        CarregarObservacoesFiscais()
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

    Function ValidarCampos() As Boolean
        If String.IsNullOrWhiteSpace(txtCodigo.Text) Then
            MsgBox(Me.Page, "Código não foi informado.")
            txtCodigo.Focus()
            Return False
        ElseIf String.IsNullOrWhiteSpace(txtDescricao.Text) Then
            MsgBox(Me.Page, "Descrição não foi informada.")
            txtDescricao.Focus()
            Return False
        Else
            Return True
        End If
    End Function

    Protected Sub lnkAtualizar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkAtualizar.Click
        Try
            If Funcoes.VerificaPermissao("ObservacoesFiscais", "ALTERAR") Then
                If ValidarCampos() Then
                    objObservacaoFiscal = New ObservacaoFiscal(txtCodigo.Text)
                    objObservacaoFiscal.IUD = "U"
                    objObservacaoFiscal.Codigo = txtCodigo.Text
                    objObservacaoFiscal.Descricao = txtDescricao.Text
                    If objObservacaoFiscal.Salvar Then
                        Limpar()
                        CarregarObservacoesFiscais()
                    Else
                        MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                    End If
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissao para alterar registro.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkExcluir_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkExcluir.Click
        Try
            If Funcoes.VerificaPermissao("ObservacoesFiscais", "EXCLUIR") Then
                objObservacaoFiscal = New ObservacaoFiscal(txtCodigo.Text)
                objObservacaoFiscal.IUD = "D"
                If objObservacaoFiscal.Salvar Then
                    Limpar()
                    CarregarObservacoesFiscais()
                Else
                    MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissao para excluir registro.")
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
            If Funcoes.VerificaPermissao("ObservacoesFiscais", "RELATORIO") Then
                Dim Sql = "Select REPLICATE('0', 3 - LEN(CAST(Codigo_Id AS varchar))) + CAST(Codigo_Id AS varchar) as Codigo, Descricao From ObservacoesFiscais Order by Codigo_Id"
                ds = Banco.ConsultaDataSet(Sql, "Tabelas")

                Dim parameters = New Dictionary(Of String, Object)()
                parameters.Add("Titulo", "Relatório De Observações Fiscais")
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
            Funcoes.Ajuda(Me.Page, "Numerador")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub

End Class