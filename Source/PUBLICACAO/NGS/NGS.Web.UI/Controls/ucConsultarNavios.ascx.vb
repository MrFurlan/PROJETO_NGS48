Imports System.Data
Imports System.Collections.Generic
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class ucConsultarNavios
    Inherits BaseUserControl

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Public Overrides Sub Limpar()
        gridNavios.DataSource = Nothing
        gridNavios.DataBind()
    End Sub

    Public Overrides Sub SetarHID(ByVal guid As String)
        HID.Value = guid.ToString
    End Sub

    Public Sub carregarNavios()
        Dim sql As String = "SELECT Codigo_Id, Descricao, Ativo " & vbCrLf &
                            "FROM Navios" & vbCrLf &
                            "WHERE Ativo = 1"

        Dim banco As New AcessaBanco
        Dim ds As New DataSet

        ds = banco.ConsultaDataSet(sql, "Navios")

        gridNavios.DataSource = ds
        gridNavios.DataBind()
    End Sub

    Private Sub carregarGrid(codigo As String)
        Dim sql As String = "SELECT Codigo_Id, Descricao, Ativo " & vbCrLf &
                            "FROM Navios" & vbCrLf &
                            "WHERE Ativo = 1"

        If Not String.IsNullOrWhiteSpace(codigo) Then
            sql += "AND Codigo_Id = '" + codigo + "'"
        End If

        Dim banco As New AcessaBanco
        Dim ds As New DataSet

        ds = banco.ConsultaDataSet(sql, "Navios")

        gridNavios.DataSource = ds
        gridNavios.DataBind()
    End Sub

    Protected Sub gridNavios_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles gridNavios.SelectedIndexChanged
        Try
            Dim objNavio = New Navio(gridNavios.SelectedRow.Cells(1).Text())
            Session("objConsultarNavios" & HID.Value) = objNavio
            If Session("objConsultarNavios" & HID.Value) IsNot Nothing Then
                If MainUserControl IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(MainUserControl.ClientID) Then
                    Dim uc As UserControl = CType(WebHelpers.FindControlRecursive(Me.Page, MainUserControl.ClientID.Split("_")(1)), UserControl)
                    CType(uc, IBaseUserControl).Carregar(objNavio)
                Else
                    CType(Me.Page, IBasePage).Carregar(objNavio)
                End If
                Popup.CloseDialog(Me.Page, "divConsultarNavios")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkFechar_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lnkFechar.Click
        Popup.CloseDialog(Me.Page, "divConsultarNavios")
    End Sub

End Class