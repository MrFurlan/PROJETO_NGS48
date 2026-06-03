Imports System.Data
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class ProdutoXFito
    Inherits BasePage

    Private Sql As String

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.setMenu(eModulo.Revenda)
            If Not IsPostBack And IsConnect Then
                If Funcoes.VerificaPermissao("ProdutoXFito", "ACESSAR") Then
                    If Funcoes.VerificaPermissao("ProdutoXFito", "LEITURA") Then
                        Limpar()
                        carregarProduto()
                    End If
                Else
                    MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Revenda.aspx")
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub carregarProduto()
        Dim Lista As New [Lib].Negocio.ListProduto("", "S", "", "", "S")
        ddlProduto.Items.Clear()

        For Each dr As [Lib].Negocio.Produto In Lista
            ddlProduto.Items.Add(New ListItem(Funcoes.AlinharEsquerda(dr.Nome, 50, ".") & "-" & dr.Codigo, dr.Codigo))
        Next

        Funcoes.InserirLinhaEmBranco(ddlProduto)
    End Sub

    Private Sub carregarFito(Optional ByVal pCodigoIndea As String = "")
        Dim Lista As New [Lib].Negocio.ListaFito("", 0, "", "", pCodigoIndea)

        ddlFito.Items.Clear()

        For Each dr As [Lib].Negocio.Fito In Lista
            ddlFito.Items.Add(New ListItem(Funcoes.AlinharEsquerda(dr.NomeComercial, 50, ".") & "-" & dr.CodigoFito, dr.CodigoFito))
        Next

        Funcoes.InserirLinhaEmBranco(ddlFito)
    End Sub

    Private Sub carregarProdutoXFito(ByVal Produto As String)
        Dim ds As New DataSet

        Sql = "SELECT     ProdutoXFito.Produto_Id, Produtos.Nome, ProdutoXFito.Fito_Id, Fito.NomeComercial, Fito.CodigoIndeaMT " & vbCrLf & _
              "FROM         ProdutoXFito INNER JOIN " & vbCrLf & _
              "                      Produtos ON ProdutoXFito.Produto_Id = Produtos.Produto_Id INNER JOIN " & vbCrLf & _
              "                      Fito ON ProdutoXFito.Fito_Id = Fito.Fito_Id " & vbCrLf & _
              "WHERE      ProdutoXFito.Produto_Id = '" & Produto & "'" & vbCrLf
        ds = Banco.ConsultaDataSet(Sql, "ProdutoXFito")

        gridProdutoXFito.DataSource = ds
        gridProdutoXFito.DataBind()
    End Sub

    Private Sub Limpar()
        'lnkNovo.Parent.Visible = True
        'lnkExcluir.Parent.Visible = False
        'lnkLimpar.Parent.Visible = True
        'ddlProduto.Parent.Visible = True
        'ddlFito.Parent.Visible = True

        'lnkNovo.Visible = True
        'lnkExcluir.Visible = False
        'lnkLimpar.Visible = True

        ddlProduto.Enabled = True
        ddlFito.Enabled = True

        ddlProduto.SelectedIndex = 0
        ddlFito.SelectedIndex = 0
        carregarProdutoXFito(ddlProduto.SelectedValue)
    End Sub

    Protected Sub gridProdutoXFito_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            Dim objProduto As New [Lib].Negocio.Produto(gridProdutoXFito.SelectedRow.Cells(1).Text())
            Dim objFito As New [Lib].Negocio.Fito("", gridProdutoXFito.SelectedRow.Cells(3).Text())

            'lnkNovo.Parent.Visible = False
            'lnkExcluir.Parent.Visible = True
            'ddlProduto.Parent.Visible = False
            'ddlFito.Parent.Visible = False

            If objProduto.ProdutoIndea <> objFito.CodigoIndeaMT Then
                MsgBox(Me.Page, "Produto selecionado está com o código do Indea diferente do fito relacionado, verifique.")
            End If
            ddlProduto.SelectedValue = objProduto.Codigo
            ddlFito.SelectedValue = objFito.CodigoFito
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub ddlProduto_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            Dim objProduto As New [Lib].Negocio.Produto(ddlProduto.SelectedValue)
            If String.IsNullOrWhiteSpace(objProduto.ProdutoIndea.ToString) Then
                MsgBox(Me.Page, "Produto selecionado não possui código do indea informado, verifique.")
            Else
                carregarProdutoXFito(objProduto.Codigo)
                carregarFito(objProduto.ProdutoIndea)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkNovo_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkNovo.Click
        Try
            If Funcoes.VerificaPermissao("ProdutoXFito", "GRAVAR") Then
                If ddlProduto.SelectedIndex = 0 Then
                    MsgBox(Me.Page, "Selecione um produto.")
                ElseIf ddlFito.SelectedIndex = 0 Then
                    MsgBox(Me.Page, "Selecione um fito.")
                Else
                    Dim objProduto As New [Lib].Negocio.Produto(ddlProduto.SelectedValue)
                    Dim objFito As New [Lib].Negocio.Fito("", ddlFito.SelectedValue)

                    If objProduto.ProdutoIndea = objFito.CodigoIndeaMT Then

                        Dim arrSQL As New ArrayList()

                        Sql = "INSERT INTO ProdutoXFito (Produto_Id, Fito_Id, ProdutoCS) " & vbCrLf & _
                              "VALUES " & vbCrLf & _
                              "('" & objProduto.Codigo & "', " & objFito.CodigoFito & ", '')"
                        arrSQL.Add(Sql)

                        If Banco.GravaBanco(arrSQL) Then
                            Limpar()
                            carregarProdutoXFito(objProduto.Codigo)
                        Else
                            MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                        End If
                    ElseIf objProduto.ProdutoIndea <> objFito.CodigoIndeaMT Then
                        MsgBox(Me.Page, "Código indea do produto selecionado é diferento do informado no fito.")
                    ElseIf objProduto.ProdutoIndea.ToString.Length = 0 Then
                        MsgBox(Me.Page, "Produto selecionado não possui código do Indea informado, verifique.")
                    ElseIf objFito.CodigoIndeaMT.Length = 0 Then
                        MsgBox(Me.Page, "Fito selecionado não possui código do indea informado, verifique.")
                    End If
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para incluir registro.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkExcluir_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkExcluir.Click
        Try
            If Funcoes.VerificaPermissao("ProdutoXFito", "EXCLUIR") Then
                Dim arrSQL As New ArrayList()
                Dim objProduto As New [Lib].Negocio.Produto(gridProdutoXFito.SelectedRow.Cells(1).Text())
                Dim objFito As New [Lib].Negocio.Fito("", gridProdutoXFito.SelectedRow.Cells(3).Text())

                Sql = "DELETE FROM ProdutoXFito WHERE Produto_Id = '" & objProduto.Codigo & "' AND Fito_Id = " & objFito.CodigoFito
                arrSQL.Add(Sql)

                If Banco.GravaBanco(arrSQL) Then
                    'LimparCampos()
                    ddlFito.SelectedIndex = 0
                    carregarProdutoXFito(ddlProduto.SelectedValue)

                Else
                    MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
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

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "ProdutoXFito")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub
End Class