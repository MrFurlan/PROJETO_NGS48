Imports System.Data
Imports System.Collections.Generic
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis
Imports System.Web.UI.Page

Public Class ucConsultaEstados
    Inherits BaseUserControl

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not IsPostBack And IsConnect Then
            Dim lstEstados As New [Lib].Negocio.Estados(True)
            GridEstados.DataSource = lstEstados
            GridEstados.DataBind()
        End If
    End Sub

    Public Overrides Sub SetarHID(ByVal guid As String)
        HID.Value = guid.ToString
    End Sub

    Public Overrides Sub Limpar()
        Session.Remove(HttpContext.Current.Session("ssTipoRetorno"))
        Session.Remove("_MainUserControl")
    End Sub

    Protected Overrides Sub Selecionar()
        Try
            Dim objEstado As New [Lib].Negocio.Estado(GridEstados.SelectedRow.Cells(1).Text())
            Session(HttpContext.Current.Session("ssTipoRetorno")) = objEstado
            If Session("ssTipoRetorno") IsNot Nothing Then
                If MainUserControl IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(MainUserControl.ClientID) Then
                    Dim uc As UserControl = CType(WebHelpers.FindControlRecursive(Me.Page, MainUserControl.ClientID.Split("_")(1)), UserControl)
                    CType(uc, IBaseUserControl).Carregar(objEstado)
                Else
                    CType(Me.Page, IBasePage).Carregar(objEstado)
                End If
                Popup.CloseDialog(Me.Page, "divConsultaEstados")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub GridEstados_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Selecionar()
    End Sub

End Class