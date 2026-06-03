Imports System.Data
Imports System.Collections.Generic
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class ucConsultaProduto
    Inherits BaseUserControl

    Private Where As String = String.Empty

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If IsPostBack Then
            Return
        End If
    End Sub

    'Protected Sub btnProduto_Click(ByVal sender As Object, ByVal e As System.EventArgs)
    '    Session("Where" & HID.Value) = ""
    '    BuscarProduto()
    'End Sub

    Protected Sub lnkProduto_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lnkProduto.Click
        Try
            Session("Where" & HID.Value) = ""
            BuscarProduto()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Public Sub BuscarProduto(Optional ByVal xmlNota As Boolean = False)
        Where = Session("Where" & HID.Value)
        Where = IIf(Where Is Nothing OrElse Where.Length = 0, "", Where)
        Dim ds As New DataSet
        Dim Sql As String
        Sql = "SELECT top(50) Produto_Id, " & vbCrLf & _
              " case len(DescricaoMapa) when 0 " & vbCrLf & _
              "     then Nome " & vbCrLf & _
              " else Nome + ' - ' + DescricaoMapa " & vbCrLf & _
              " end AS Nome, " & vbCrLf & _
              " CodigoProdutoTerceiro AS ProdutoDeTerceiro " & vbCrLf & _
              "  FROM Produtos " & vbCrLf & _
              " WHERE Situacao = 1" & vbCrLf

        If txtNome.Text.Length > 0 Then
            Sql &= "And Nome LIKE '%" & txtNome.Text & "%' " & vbCrLf
        End If

        If txtTerceiro.Text.Length > 0 Then
            Sql &= "And CodigoProdutoTerceiro LIKE '%" & txtTerceiro.Text & "%' " & vbCrLf
        End If

        If Where.Trim.Length > 0 Then
            Sql &= "And " & Where
        End If
        Sql &= " ORDER BY Nome"
        ds = Banco.ConsultaDataSet(Sql, "Produtos")
        gridProduto.DataSource = ds
        gridProduto.DataBind()

        If xmlNota Then MsgBox(Me.Page, "Para facilitar foi carregado o(s) Produto(s) com o mesmo NCM informado no XML recebido da Nota Fiscal")
    End Sub

    Public Overrides Sub SetarHID(ByVal guid As String)
        HID.Value = guid.ToString
    End Sub

    Public Overrides Sub Limpar()
        gridProduto.DataSource = New List(Of Object)
        gridProduto.DataBind()
        txtNome.Text = String.Empty
        txtTerceiro.Text = String.Empty
        Session.Remove("_MainUserControl")
        txtNome.Focus()
    End Sub

    Protected Overrides Sub Selecionar()
        Try
            Dim objProduto As New [Lib].Negocio.Produto(gridProduto.SelectedRow.Cells(1).Text())
            Session(Session("ssTipoRetorno")) = objProduto
            If Session("ssTipoRetorno") IsNot Nothing Then
                If MainUserControl IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(MainUserControl.ClientID) Then
                    Dim uc As UserControl = CType(WebHelpers.FindControlRecursive(Me.Page, MainUserControl.ClientID.Split("_")(1)), UserControl)
                    CType(uc, IBaseUserControl).Carregar(objProduto)
                Else
                    CType(Me.Page, IBasePage).Carregar(objProduto)
                End If
                Popup.CloseDialog(Me.Page, "divConsultaProduto")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub Fechar()
        Try
            Dim objProduto As New [Lib].Negocio.Produto()
            CType(Me.Page, IBasePage).Carregar(objProduto)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Protected Sub gridProduto_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Selecionar()
    End Sub

    Protected Sub txtNome_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        BuscarProduto()
    End Sub

    Protected Sub btnFechar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnFechar.Click
        Fechar()
        Popup.CloseDialog(Me.Page, "divConsultaProduto")
    End Sub

End Class