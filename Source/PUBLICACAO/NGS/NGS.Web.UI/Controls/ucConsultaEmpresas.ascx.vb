Imports System.Data
Imports System.Collections.Generic
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class ucConsultaEmpresas
    Inherits BaseUserControl

    Dim sql As String = String.Empty

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not IsPostBack And IsConnect Then
            CarregarUnidadesNegocio()
        End If
    End Sub

    Public Overrides Sub SetarHID(ByVal guid As String)
        HID.Value = guid.ToString
    End Sub

    Public Overrides Sub Limpar()
        CarregarUnidadesNegocio()
        GridClientes.DataSource = New List(Of Object)
        GridClientes.DataBind()
        GridClientes.SelectedIndex = -1
        Session.Remove("_MainUserControl")
    End Sub

    Protected Overrides Sub Selecionar(ByVal args As String)
        Try
            Dim strEmpresa As String() = args.Split(";")
            Dim objEmpresa As New [Lib].Negocio.Cliente(strEmpresa(0), Convert.ToInt32(strEmpresa(1)))
            objEmpresa.UnidadeDeNegocio = New [Lib].Negocio.Cliente(LbxUnidades.SelectedValue, 0)
            Session(Session("ssTipoRetorno")) = objEmpresa
            If Session("ssTipoRetorno") IsNot Nothing Then
                If MainUserControl IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(MainUserControl.ClientID) Then
                    Dim uc As UserControl = CType(WebHelpers.FindControlRecursive(Me.Page, MainUserControl.ClientID.Split("_")(1)), UserControl)
                    CType(uc, IBaseUserControl).Carregar(objEmpresa)
                Else
                    CType(Me.Page, IBasePage).Carregar(objEmpresa)
                End If
                Popup.CloseDialog(Me.Page, "divConsultaEmpresas")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub CarregarUnidadesNegocio()
        LbxUnidades.Items.Clear()
        Dim Tipos As New List(Of [Lib].Negocio.eTipoCliente)
        Tipos.Add([Lib].Negocio.eTipoCliente.UnidadesNegocio)
        Dim objUnidadesNegocio As New [Lib].Negocio.ListCliente("", Tipos)
        If objUnidadesNegocio.Count > 0 Then
            For Each objUnidade As [Lib].Negocio.Cliente In objUnidadesNegocio
                LbxUnidades.Items.Add(New ListItem(objUnidade.Nome, objUnidade.Codigo.Trim()))
            Next
        Else : MsgBox(Me.Page, objUnidadesNegocio.Erro.Message, eTitulo.Info)
        End If
        objUnidadesNegocio = Nothing
    End Sub

    Protected Sub LbxUnidades_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles LbxUnidades.SelectedIndexChanged
        Dim objEmpresas As New [Lib].Negocio.ListCliente(LbxUnidades.SelectedValue)
        If objEmpresas.Count > 0 Then
            GridClientes.DataSource = objEmpresas.ToArray()
            GridClientes.DataBind()
            GridClientes.SelectedIndex = -1
        End If
    End Sub

    Protected Sub GridClientes_RowCommand(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewCommandEventArgs)
        Selecionar(e.CommandArgument.ToString())
    End Sub

    Protected Sub btnFechar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnFechar.Click
        Popup.CloseDialog(Me.Page, "divConsultaEmpresas")
    End Sub

End Class