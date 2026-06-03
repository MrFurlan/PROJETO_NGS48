Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class ucConsultarNaviosXInvoice
    Inherits BaseUserControl
    Private Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            If Not IsPostBack And IsConnect Then
                Limpar()
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Public Overrides Sub Limpar()
        ddlNavios.Items.Clear()
    End Sub

    Public Sub carregarNavios(ByVal codProduto As String)
        Session("objProduto" & HID.Value) = ""
        ddl.Carregar(ddlNavios, CarregarDDL.Tabela.NaviosXInvoice, codProduto)
        Session("objProduto" & HID.Value) = codProduto

        gridNaviosXInvoice.DataSource = Nothing
        gridNaviosXInvoice.DataBind()
    End Sub

    Private Sub carregarNaviosXInvoice(codigo As String)
        Dim sql As String = "SELECT nxi.Codigo_Id, nxi.Navio_Id, n.Descricao, nxi.DataDeChegada, nxi.Observacao " & vbCrLf &
                            "FROM NaviosXInvoice nxi" & vbCrLf &
                            "INNER JOIN Navios n" & vbCrLf &
                            "ON n.Codigo_Id = nxi.Navio_Id" & vbCrLf &
                            "INNER JOIN NavioXInvoiceXProduto nxixp " & vbCrLf &
                            "ON nxixp.Codigo_Id  = nxi.Codigo_Id " & vbCrLf &
                            "AND nxixp.Produto_Id = '" & Session("objProduto" & HID.Value) & "'" & vbCrLf &
                            "WHERE nxi.Ativo = 1" & vbCrLf

        If Not String.IsNullOrWhiteSpace(codigo) Then
            sql += "AND nxi.Navio_Id = " + codigo
        End If

        Dim banco As New AcessaBanco
        Dim ds As New DataSet
        ds = banco.ConsultaDataSet(sql, "NaviosXInvoice")

        If ds IsNot Nothing AndAlso ds.Tables IsNot Nothing AndAlso ds.Tables.Count > 0 And ds.Tables(0).Rows.Count > 0 Then
            gridNaviosXInvoice.DataSource = ds
            gridNaviosXInvoice.DataBind()
        Else
            MsgBox(Me.Page, "Não foram encontrados invoices, para o navio selecionado.", eTitulo.Info)
            Exit Sub
        End If

    End Sub

    Protected Sub ddlNavios_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ddlNavios.SelectedIndexChanged
        Try
            Dim codigoNavio = ddlNavios.SelectedValue.ToString().Split(" ")
            carregarNaviosXInvoice(codigoNavio(0))
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Public Overrides Sub SetarHID(ByVal guid As String)
        HID.Value = guid.ToString
    End Sub

    Protected Sub gridNaviosXInvoice_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles gridNaviosXInvoice.SelectedIndexChanged
        Try
            Dim objNaviosXInvoice = New NavioXInvoice(gridNaviosXInvoice.SelectedRow.Cells(1).Text())
            Session("objNavioXInvoice" & HID.Value) = objNaviosXInvoice
            If Session("objNavioXInvoice" & HID.Value) IsNot Nothing Then
                If MainUserControl IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(MainUserControl.ClientID) Then
                    Dim uc As UserControl = CType(WebHelpers.FindControlRecursive(Me.Page, MainUserControl.ClientID.Split("_")(1)), UserControl)
                    CType(uc, IBaseUserControl).Carregar(objNaviosXInvoice)
                Else
                    CType(Me.Page, IBasePage).Carregar(objNaviosXInvoice)
                End If
                Popup.CloseDialog(Me.Page, "divConsultarNaviosXInvoice")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkLimpar_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lnkLimpar.Click
        Try
            Limpar()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkFechar_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lnkFechar.Click
        Try
            Limpar()
            Popup.CloseDialog(Me.Page, "divConsultarNaviosXInvoice")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub
End Class