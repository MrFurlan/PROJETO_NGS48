Imports System.Data
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class ProdutosXAnalises
    Inherits BasePage

    Dim Sql As String

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.setMenu(eModulo.Gestao)
            If Not IsPostBack And IsConnect Then
                If Funcoes.VerificaPermissao("ProdutosXAnalises", "ACESSAR") Then
                    Produtos()
                    Analises()
                    BindGridView()
                    Limpar()
                Else
                    MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Gestao")
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub grdProdutoXAnalise_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdProdutoXAnalise.SelectedIndexChanged
        Try
            Dim produto As String = grdProdutoXAnalise.SelectedRow.Cells(1).Text().Trim()
            Dim analise As String = grdProdutoXAnalise.SelectedRow.Cells(2).Text().Trim()
            If Not String.IsNullOrWhiteSpace(produto) AndAlso Not String.IsNullOrWhiteSpace(analise) Then
                ddlProduto.SelectedValue = produto
                ddlAnalise.SelectedValue = analise
            End If
            lnkNovo.Parent.Visible = False
            lnkAtualizar.Parent.Visible = True
            lnkExcluir.Parent.Visible = True
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub Limpar()
        ddlProduto.SelectedIndex = 0
        ddlAnalise.SelectedIndex = 0
        BindGridView()
        lnkNovo.Parent.Visible = True
        lnkAtualizar.Parent.Visible = False
        lnkExcluir.Parent.Visible = False
    End Sub

    Private Function Validar() As Boolean
        Dim aux As Boolean = True

        If Not ddlProduto.SelectedIndex > 0 Then
            aux = False
            MsgBox(Me.Page, "É necessário selecionar um produto.")
        End If
        If Not ddlAnalise.SelectedIndex > 0 Then
            aux = False
            MsgBox(Me.Page, "É necessário selecionar uma análise.")
        End If

        Return aux
    End Function

    Private Sub BindGridView()
        Sql = "SELECT     ProdutosXAnalises.Produto_Id AS Produto, " & vbCrLf & _
              "          ProdutosXAnalises.Analise_Id AS Analise, " & vbCrLf & _
              "          Analises.Descricao AS AnaliseDescricao, " & vbCrLf & _
              "          Produtos.Descricao AS ProdutoDescricao, " & vbCrLf & _
              "          cast(ProdutosXAnalises.Produto_Id as varchar(100)) + ' - ' + cast(Produtos.Descricao as varchar(100)) as ProdutoCompleto, " & vbCrLf & _
              "          cast(ProdutosXAnalises.Analise_Id as varchar(100)) + ' - ' + cast(Analises.Descricao as varchar(100)) as AnaliseCompleto " & vbCrLf & _
              "FROM      ProdutosXAnalises INNER JOIN " & vbCrLf & _
              "          Analises ON Analises.Analise_Id = ProdutosXAnalises.Analise_Id INNER JOIN " & vbCrLf & _
              "          Produtos ON ProdutosXAnalises.Produto_Id = Produtos.Produto_Id order by ProdutosXAnalises.Produto_Id, ProdutosXAnalises.Analise_Id" & vbCrLf
        grdProdutoXAnalise.DataSource = Banco.ConsultaDataSet(Sql, "ProdutosXAnalises")
        grdProdutoXAnalise.DataBind()
    End Sub

    Private Sub Analises()
        Sql = "select Analise_Id, Descricao, cast(Analise_Id as varchar(50)) + ' - ' + cast(Descricao as varchar(100)) as NomeCompleto from Analises order by Analise_Id"
        ddlAnalise.Items.Clear()
        ddlAnalise.Items.Add(New ListItem("", ""))
        For Each row As DataRow In Banco.ConsultaDataSet(Sql, "Analises").Tables(0).Rows
            ddlAnalise.Items.Add(New ListItem(row("NomeCompleto"), row("NomeCompleto")))
        Next
    End Sub

    Private Sub Produtos()
        Sql = "SELECT Produto_Id As Codigo, Nome as Descricao, cast(Produto_Id as varchar(50)) + ' - ' + cast(Nome as varchar(100)) as NomeCompleto FROM Produtos order by Produto_Id"
        ddlProduto.Items.Clear()
        ddlProduto.Items.Add(New ListItem("", ""))
        For Each row As DataRow In Banco.ConsultaDataSet(Sql, "Produtos").Tables(0).Rows
            ddlProduto.Items.Add(New ListItem(row("NomeCompleto"), row("NomeCompleto")))
        Next
    End Sub

    Private Sub Incluir(ByVal Produto As String, ByVal Analise As String)
        If Funcoes.VerificaPermissao("ProdutosXAnalises", "GRAVAR") Then
            Dim SqlArray As New ArrayList

            Sql = "INSERT INTO ProdutosXAnalises(Produto_Id, Analise_Id) " & vbCrLf & _
                  "Values('" & RTrim(Produto) & "', " & RTrim(Analise) & ")" & vbCrLf
            SqlArray.Add(Sql)

            If Banco.GravaBanco(SqlArray) = False Then
                MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
            End If
        Else
            MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
        End If
    End Sub

    Private Sub Excluir(ByVal Produto As String, ByVal Analise As String)
        If Funcoes.VerificaPermissao("ProdutosXAnalises", "EXCLUIR") Then
            Dim SqlArray As New ArrayList
            Sql = "DELETE FROM ProdutosXAnalises WHERE Produto_Id = '" & RTrim(Produto) & "' AND "
            Sql &= " Analise_Id = " & RTrim(Analise)

            SqlArray.Add(Sql)

            If Banco.GravaBanco(SqlArray) = False Then
                MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
            End If
        Else
            MsgBox(Me.Page, Funcoes.EliminarCaracteresEspeciais(Session("ssMessage").ToString()))
        End If
    End Sub

    Private Sub Relatorio()
        Dim ds As New DataSet

        Sql = "SELECT    ProdutosXAnalises.Produto_Id AS Produto, " & vbCrLf & _
              "          ProdutosXAnalises.Analise_Id AS Analise, " & vbCrLf & _
              "          Analises.Descricao AS AnaliseDescricao, " & vbCrLf & _
              "          Produtos.Descricao AS ProdutoDescricao " & vbCrLf & _
              "FROM      ProdutosXAnalises INNER JOIN " & vbCrLf & _
              "          Analises ON Analises.Analise_Id = ProdutosXAnalises.Analise_Id INNER JOIN " & vbCrLf & _
              "          Produtos ON ProdutosXAnalises.Produto_Id = Produtos.Produto_Id " & vbCrLf & _
              "          Order by ProdutosXAnalises.Produto_Id, ProdutosXAnalises.Analise_Id " & vbCrLf
        ds = Banco.ConsultaDataSet(Sql, "ProdutosXAnalises")

        Dim parameters = New Dictionary(Of String, Object)()

        Funcoes.BindReport(Me.Page, ds, "Cr_ProdutosXAnalises", eExportType.PDF)
    End Sub

    Protected Sub lnkNovo_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkNovo.Click
        Try
            If Not Validar() Then
                Return
            End If
            Dim produto As String() = ddlProduto.SelectedValue.Split(New Char() {"-"c}, StringSplitOptions.RemoveEmptyEntries)
            Dim analise As String() = ddlAnalise.SelectedValue.Split(New Char() {"-"c}, StringSplitOptions.RemoveEmptyEntries)
            If produto IsNot Nothing AndAlso produto.Length > 0 AndAlso analise IsNot Nothing AndAlso analise.Length > 0 Then
                Incluir(produto(0).Trim(), analise(0).Trim())
                BindGridView()
            End If
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

    Protected Sub lnkExcluir_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkExcluir.Click
        Try
            If Not Validar() Then
                Return
            End If
            Dim produto As String() = ddlProduto.SelectedValue.Split(New Char() {"-"c}, StringSplitOptions.RemoveEmptyEntries)
            Dim analise As String() = ddlAnalise.SelectedValue.Split(New Char() {"-"c}, StringSplitOptions.RemoveEmptyEntries)
            If produto IsNot Nothing AndAlso produto.Length > 0 AndAlso analise IsNot Nothing AndAlso analise.Length > 0 Then
                Excluir(produto(0).Trim(), analise(0).Trim())
                BindGridView()
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

    Protected Sub lnkRelatorio_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkRelatorio.Click
        Try
            If Funcoes.VerificaPermissao("ProdutosXAnalises", "RELATORIO") Then
                Relatorio()
            Else
                MsgBox(Me.Page, "Usuário sem permissão para emitir relatório.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAjuda_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "ProdutosXAnalises")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

End Class