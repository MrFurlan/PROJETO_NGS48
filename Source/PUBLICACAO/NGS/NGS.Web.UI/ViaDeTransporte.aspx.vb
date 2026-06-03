Imports System.Data
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class ViaDeTransporte
    Inherits BasePage

    Dim Sql As String
    Dim SqlArray As New ArrayList
    Dim DS As DataSet
    Dim objViaDeTransporte As [Lib].Negocio.ViaDeTransporte

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Gestao)
        If Not IsPostBack And IsConnect Then
            If Funcoes.VerificaPermissao("ViaDeTransporte", "ACESSAR") Then
                CarregarViaDeTransportes()
                Limpar()
            Else
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Gestao.aspx")
                Exit Sub
            End If
        End If
    End Sub

    Private Sub CarregarViaDeTransportes()
        If Funcoes.VerificaPermissao("ViaDeTransporte", "LEITURA") Then
            Sql = "  SELECT Codigo_id as Codigo, Descricao " & vbCrLf & _
                  " FROM ViaDeTransportes " & vbCrLf & _
                  " ORDER BY Codigo_id" & vbCrLf
            GridViaDeTransportes.DataSource = Banco.ConsultaDataSet(Sql, "Consulta")
            GridViaDeTransportes.DataBind()
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

    Protected Sub GridViaDeTransportes_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            Limpar()
            txtCodigo.Text = GridViaDeTransportes.SelectedRow.Cells(1).Text()
            txtDescricao.Text = GridViaDeTransportes.SelectedRow.Cells(2).Text()
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
            If Funcoes.VerificaPermissao("ViaDeTransporte", "GRAVAR") Then
                If ValidarCampos() Then
                    objViaDeTransporte = New [Lib].Negocio.ViaDeTransporte()
                    objViaDeTransporte.IUD = "I"
                    objViaDeTransporte.Codigo = txtCodigo.Text
                    objViaDeTransporte.Descricao = txtDescricao.Text
                    If objViaDeTransporte.Salvar Then
                        Limpar()
                        CarregarViaDeTransportes()
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
        If Funcoes.VerificaPermissao("ViaDeTransporte", "ALTERAR") Then
            If ValidarCampos() Then
                objViaDeTransporte = New [Lib].Negocio.ViaDeTransporte()
                objViaDeTransporte.IUD = "U"
                objViaDeTransporte.Codigo = txtCodigo.Text
                objViaDeTransporte.Descricao = txtDescricao.Text
                If objViaDeTransporte.Salvar Then
                    Limpar()
                    CarregarViaDeTransportes()
                Else
                    MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                End If
            End If
        Else
            MsgBox(Me.Page, "Usuário sem permissão para alterar registro.", eTitulo.Info)
        End If
    End Sub

    Protected Sub lnkExcluir_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkExcluir.Click
        If Funcoes.VerificaPermissao("ViaDeTransporte", "EXCLUIR") Then
            objViaDeTransporte = New [Lib].Negocio.ViaDeTransporte(txtCodigo.Text)
            objViaDeTransporte.IUD = "D"
            objViaDeTransporte.Codigo = Trim(txtCodigo.Text)
            objViaDeTransporte.Descricao = RTrim(txtDescricao.Text)
            txtCodigo.Enabled = False
            If objViaDeTransporte.Salvar Then
                Limpar()
                CarregarViaDeTransportes()
            Else
                MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
            End If
        Else
            MsgBox(Me.Page, "Usuário sem permissão para excluir registro.", eTitulo.Info)
        End If
    End Sub

    Protected Sub lnkRelatorio_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkRelatorio.Click
        Try
            If Funcoes.VerificaPermissao("ViaDeTransporte", "RELATORIO") Then
                Sql = "Select REPLICATE('0', 3 - LEN(CAST(Codigo_Id AS varchar))) + CAST(Codigo_Id AS varchar) as Codigo, Descricao From ViaDeTransportes Order by Codigo_Id"

                DS = Banco.ConsultaDataSet(Sql, "Tabelas")

                Dim parameters = New Dictionary(Of String, Object)()
                parameters.Add("Titulo", "Tabela de Via De Transportes")
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

    Protected Sub lnkLimpar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkLimpar.Click
        Limpar()
    End Sub

    Protected Sub lnkAjuda_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "ViaDeTransporte")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Function ValidarCampos() As Boolean
        If String.IsNullOrWhiteSpace(txtCodigo.Text) Then
            MsgBox(Me.Page, "Código da Via de Transporte não foi informado!")
            Return False
        ElseIf String.IsNullOrWhiteSpace(txtDescricao.Text) Then
            MsgBox(Me.Page, "Descrição da Via de Transporte não foi informada!")
            Return False
        Else
            Return True
        End If
    End Function

End Class