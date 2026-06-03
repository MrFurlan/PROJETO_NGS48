Imports System.IO
Imports System.Data
Imports System.Net.Configuration
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis
Imports System.Threading
Imports System.Xml


Public Class Setores
    Inherits BasePage

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            If Not IsPostBack And IsConnect Then
                'If Funcoes.VerificaPermissao("Setores", "ACESSAR") Then
                Limpar()
                'Else
                '    MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "")
                '    Exit Sub
                'End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub GridSetores_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try

        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub Limpar()
        lblEmpresa.Text = ""
        txtSetor.Text = ""

        Session.Remove("objEmpresaCON" & HID.Value)

        HID.Value = Guid.NewGuid().ToString
        ucConsultaEmpresas.SetarHID(HID.Value)
    End Sub

    Public Overrides Sub Carregar(ByVal obj As [Lib].Negocio.IBaseEntity)
        If Not Session("objEmpresaCON" & HID.Value) Is Nothing Then
            Dim itemEmpresa As ListItem = Funcoes.FormatarListItemCliente(CType(Session("objEmpresaCON" & HID.Value), [Lib].Negocio.Cliente))
            lblEmpresa.Text = itemEmpresa.Text
            Session.Remove("objEmpresaCON" & HID.Value)
        End If
    End Sub

    Protected Sub lnkNovo_Click(sender As Object, e As EventArgs) Handles lnkNovo.Click
        Try

        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAtualizar_Click(sender As Object, e As EventArgs) Handles lnkAtualizar.Click
        Try

        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkExcluir_Click(sender As Object, e As EventArgs) Handles lnkExcluir.Click
        Try

        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkLimpar_Click(sender As Object, e As EventArgs) Handles lnkLimpar.Click
        Try
            Limpar()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try

        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkEmpresa_Click(sender As Object, e As EventArgs) Handles lnkEmpresa.Click
        Try
            ucConsultaEmpresas.Limpar()
            Popup.ConsultaDeEmpresas(Me, "objEmpresaCON" & HID.Value)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

End Class