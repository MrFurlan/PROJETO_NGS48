Imports System.Web.Script.Services
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class Gestao
    Inherits BasePage

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Gestao)
        If Not IsPostBack And IsConnect Then
            If Not Funcoes.VerificaPermissao("Gestao", "ACESSAR") Then
                MsgBox(Me.Page, "Usuário sem permissão para acessar esse módulo", IIf(Not String.IsNullOrEmpty(Session("ssMenuDeAcesso")), Session("ssMenuDeAcesso") & ".aspx", "~/Index.aspx"), eTitulo.Info)
                Exit Sub
            End If
        End If
    End Sub

    <Services.WebMethod()>
    <ScriptMethod(UseHttpGet:=True, ResponseFormat:=ResponseFormat.Json)>
    Public Shared Function GetEmpresa()

        If HttpContext.Current.Session("ssEmpresa") Is Nothing Then Return ""

        Return HttpContext.Current.Session("ssEmpresa").ToString()
    End Function

End Class