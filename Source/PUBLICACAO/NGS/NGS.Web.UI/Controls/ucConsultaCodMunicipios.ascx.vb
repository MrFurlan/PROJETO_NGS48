Imports System.Data
Imports System.Collections.Generic
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class ucConsultaCodMunicipios
    Inherits BaseUserControl

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not IsPostBack And IsConnect Then
            If Not Session("ssUF" & HID.Value) Is Nothing Then CargaMunicipios()
        End If
    End Sub

    Public Overrides Sub SetarHID(ByVal guid As String)
        HID.Value = guid.ToString
    End Sub

    Private Sub CargaMunicipios()
        Dim ds As New DataSet
        Dim Municipio() As String = Session("ssUF" & HID.Value).ToString.Split(";")

        Dim strCampos As String() = Nothing
        If Not Request.QueryString("cmp") Is Nothing Then
            strCampos = Request.QueryString("cmp").Split("|")
        End If

        If Not Request.QueryString("mun") Is Nothing Then
            Municipio(1) = Request.QueryString("mun")
        End If

        If Request.QueryString("tipo") = "PDXI" Or Request.QueryString("tipo") = "CLI" Or Request.QueryString("tipo") = "CLI2" Or Request.QueryString("tipo") = "CLIxIMO" Or Request.QueryString("tipo") = "CLIxA" Or Request.QueryString("tipo") = "CLIxCTA" Then
            Session("ssUF" & HID.Value) = strCampos(0)
        End If

        Dim sql As String = "SELECT Codigo_id AS Codigo, Municipio_id AS Descricao FROM Municipios "
        sql &= "WHERE Estado_Id = '" & Municipio(0) & "' "
        If Municipio.Count > 1 AndAlso Municipio(1).Length > 0 Then
            sql &= "AND Municipio_id LIKE '" & Municipio(1) & "%' "
        End If
        sql &= "ORDER BY Municipio_id"

        ds = Banco.ConsultaDataSet(sql, "Municipios")
        GridCodMunicipio.DataSource = ds
        GridCodMunicipio.DataBind()
    End Sub

    Public Overrides Sub Limpar()
        TxtMunicipio.Text = ""
        GridCodMunicipio.DataSource = New List(Of Object)
        GridCodMunicipio.DataBind()
        Session.Remove(Session("ssTipoRetorno"))
        TxtMunicipio.Focus()
        If Not Session("ssUF" & HID.Value) Is Nothing Then CargaMunicipios()
    End Sub

    Protected Overrides Sub Selecionar()
        Try
            Dim objMunicipio = New [Lib].Negocio.Municipio(Session("ssUF" & HID.Value).ToString.Split(";")(0), GridCodMunicipio.SelectedRow.Cells(2).Text())
            Session(HttpContext.Current.Session("ssTipoRetorno")) = objMunicipio
            If Session("ssTipoRetorno") IsNot Nothing Then
                If MainUserControl IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(MainUserControl.ClientID) Then
                    Dim uc As UserControl = CType(WebHelpers.FindControlRecursive(Me.Page, MainUserControl.ClientID.Split("_")(1)), UserControl)
                    CType(uc, IBaseUserControl).Carregar(objMunicipio)
                Else
                    CType(Me.Page, IBasePage).Carregar(objMunicipio)
                End If
                Popup.CloseDialog(Me.Page, "divConsultaCodMunicipios")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub GridCodMunicipio_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridCodMunicipio.SelectedIndexChanged
        Selecionar()
    End Sub

    Protected Sub BtnOK_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim ds As New DataSet
        Dim sql As String = "SELECT Codigo_id AS Codigo, Municipio_id AS Descricao FROM Municipios "
        sql &= "WHERE Estado_Id = '" & Session("ssUF" & HID.Value) & "' AND "
        sql &= "Municipio_id LIKE '" & TxtMunicipio.Text & "%' "
        sql &= "ORDER BY Municipio_id"
        ds = Banco.ConsultaDataSet(sql, "Municipios")
        GridCodMunicipio.DataSource = ds
        GridCodMunicipio.DataBind()
    End Sub

    Protected Sub lnkFechar_Click(sender As Object, e As EventArgs) Handles lnkFechar.Click
        Popup.CloseDialog(Me.Page, "divConsultaCodMunicipios")
    End Sub
End Class