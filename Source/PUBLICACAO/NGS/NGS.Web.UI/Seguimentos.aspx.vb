Imports System.Data
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class Seguimentos
    Inherits BasePage

    Private Sql As String
    Private SqlArray As New ArrayList
    Private DS As DataSet
    Private objSeguimento As [Lib].Negocio.Seguimento

    Public Property SessaoSeguimento As [Lib].Negocio.Seguimento
        Get
            Return Session("objSeguimento" & HID.Value)
        End Get
        Set(ByVal value As [Lib].Negocio.Seguimento)
            Session("objSeguimento" & HID.Value) = value
        End Set
    End Property

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Expedicao)
        If Not IsPostBack And IsConnect Then
            If Funcoes.VerificaPermissao("Seguimentos", "ACESSAR") Then
                CarregarSeguimentos()
                Limpar()
            Else
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página", "~/Expedicao.aspx")
                Exit Sub
            End If
        End If
    End Sub

    Private Sub CarregarSeguimentos()
        If Funcoes.VerificaPermissao("Seguimentos", "LEITURA") Then
            Dim objListaSeguimentos As New ListSeguimento(True)
            GridSeguimentos.DataSource = objListaSeguimentos
            GridSeguimentos.DataBind()
        End If
    End Sub

    Private Sub Limpar()
        txtCodigo.Text = ""
        txtDescricao.Text = ""
        lnkNovo.Parent.Visible = True
        lnkAtualizar.Parent.Visible = False
        lnkExcluir.Parent.Visible = False
    End Sub

    Protected Sub GridSeguimentos_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridSeguimentos.SelectedIndexChanged
        Try
            Limpar()
            txtCodigo.Text = GridSeguimentos.SelectedRow.Cells(1).Text()
            txtDescricao.Text = GridSeguimentos.SelectedRow.Cells(2).Text()
            objSeguimento = New [Lib].Negocio.Seguimento(GridSeguimentos.SelectedRow.Cells(1).Text())
            SessaoSeguimento = objSeguimento
            txtDescricao.Focus()
            lnkNovo.Parent.Visible = False
            lnkAtualizar.Parent.Visible = True
            lnkExcluir.Parent.Visible = True
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Function ValidaSeguimento() As Boolean
        If String.IsNullOrWhiteSpace(txtDescricao.Text) Then
            MsgBox(Me.Page, "A Descrição da Seguimento é Obrigatória.")
            txtDescricao.Focus()
            Return False
        End If
        Return True
    End Function

    Protected Sub lnkNovo_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkNovo.Click
        Try
            If Funcoes.VerificaPermissao("Seguimentos", "GRAVAR") Then
                If ValidaSeguimento() Then
                    objSeguimento = New [Lib].Negocio.Seguimento()
                    objSeguimento.Descricao = Trim(txtDescricao.Text.ToUpper)
                    objSeguimento.IUD = "I"
                    If Not objSeguimento.Salvar() Then
                        MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                    Else
                        Limpar()
                        CarregarSeguimentos()
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
            If Funcoes.VerificaPermissao("Seguimentos", "ALTERAR") Then
                If ValidaSeguimento() Then
                    objSeguimento = SessaoSeguimento
                    objSeguimento.Descricao = Trim(txtDescricao.Text.ToUpper)
                    objSeguimento.IUD = "U"

                    If Not objSeguimento.Salvar Then
                        MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                    Else
                        Limpar()
                        CarregarSeguimentos()
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
            If Funcoes.VerificaPermissao("Seguimentos", "EXCLUIR") Then
                objSeguimento = SessaoSeguimento

                'If objSeguimento.LiberaExclusaoDaSeguimento Then
                objSeguimento.IUD = "D"
                If Not objSeguimento.Salvar Then
                    MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                Else
                    Limpar()
                    CarregarSeguimentos()
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
            If Funcoes.VerificaPermissao("Seguimentos", "RELATORIO") Then
                Sql = "Select REPLICATE('0', 3 - LEN(CAST(Seguimento_Id AS varchar))) + CAST(Seguimento_Id AS varchar) as Codigo, Descricao From Seguimentos Order by Seguimento_Id"
                DS = Banco.ConsultaDataSet(Sql, "Tabelas")

                Dim parameters As New Dictionary(Of String, Object)
                parameters.Add("Titulo", "Relatório De Seguimentos De Produto.")
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
            Funcoes.Ajuda(Me.Page, "Seguimentos")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub

End Class