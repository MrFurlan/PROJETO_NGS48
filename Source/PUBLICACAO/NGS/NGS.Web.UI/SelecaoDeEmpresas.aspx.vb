Imports System.Data
Imports System.IO
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class SelecaoDeEmpresas
    Inherits BasePage

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Gestao)
        If Not IsPostBack And IsConnect Then
            If Not Funcoes.VerificaPermissao("SelEmpresas", "ACESSAR") Then
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página!", "~/Gestao.aspx")
                Exit Sub
            End If
            Limpar()
            CarregarEmpresa()
        End If
    End Sub

    Protected Sub btnEmpresa_Click(sender As Object, e As EventArgs) Handles btnEmpresa.Click
        ucConsultaEmpresas.Limpar()
        Popup.ConsultaDeEmpresas(Me.Page, "objSelecaoDeEmpresas" & HID.Value)
    End Sub

    Public Overrides Sub Carregar(ByVal obj As [Lib].Negocio.IBaseEntity)
        If Not Session("objSelecaoDeEmpresas" & HID.Value) Is Nothing Then
            Dim objEmpresa As [Lib].Negocio.Cliente = CType(Session("objSelecaoDeEmpresas" & HID.Value), [Lib].Negocio.Cliente)
            HttpContext.Current.Session("Empresa") = objEmpresa
            HttpContext.Current.Session("ssEmpresa") = objEmpresa.Codigo
            HttpContext.Current.Session("ssEndEmpresa") = objEmpresa.CodigoEndereco
            UsuarioServidor.VerificarEmpresa()

            Dim Sqls As New ArrayList
            If UsuarioServidor.Usuario IsNot Nothing Then
                Dim objUsuario As New [Lib].Negocio.Usuario(UsuarioServidor.Usuario.Usuario_Id)
                If objUsuario IsNot Nothing Then
                    objUsuario.IUD = "U"
                    objUsuario.Empresa_Id = objEmpresa.Codigo
                    objUsuario.EnderecoEmpresa = objEmpresa.CodigoEndereco
                    objUsuario.AcessoUnidade = objEmpresa.UnidadeDeNegocio.Codigo
                    objUsuario.AcessoEmpresa = objEmpresa.Codigo
                    objUsuario.AcessoEnderecoEmpresa = objEmpresa.CodigoEndereco
                    objUsuario.SalvarSql(Sqls)
                End If
            End If

            If Not Banco.GravaBanco(Sqls) Then
                MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
            End If
            MsgBox(Me.Page, "Empresa alterada com sucesso!", "~/Index.aspx")
        End If
    End Sub

    Private Sub CarregarEmpresa()
        Dim objEmpresa As New [Lib].Negocio.Cliente(HttpContext.Current.Session("ssEmpresa"), HttpContext.Current.Session("ssEndEmpresa"))
        txtEmpresa.Text = Funcoes.FormatarListItemCliente(objEmpresa).Text
    End Sub

    Private Sub Limpar()
        Session.Remove("objSelecaoDeEmpresas" & HID.Value)
        HID.Value = Guid.NewGuid.ToString()
        txtEmpresa.Text = String.Empty
    End Sub

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "SelecaoDeEmpresas")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub
End Class