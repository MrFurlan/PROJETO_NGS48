Imports System.Data
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class ClassificacaoDeProdutos
    Inherits BasePage

#Region "Eventos"

#Region "Page_Load"

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.setMenu(eModulo.Gestao)
            If Not IsPostBack And IsConnect Then
                If Funcoes.VerificaPermissao("ClassificacaoDeProdutos", "ACESSAR") Then
                    BindCombos()
                    GerenciarBotoes()
                    CargaGrupo()
                    Limpar()
                Else
                    MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Gestao.aspx")
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

#End Region

#Region "DropDownList"

    Protected Sub ddlProdutoDestino_SelectedIndexChanged(ByVal sender As Object, ByVal e As EventArgs) Handles ddlProdutoDestino.SelectedIndexChanged
        Try
            If (ConsultarReplicar(ddlTabelaDestino.SelectedValue, ddlProdutoDestino.SelectedValue)) Then
                chkTabelaProdutoExiste.Visible = True
            Else
                chkTabelaProdutoExiste.Visible = False
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub ddlGrupoOrigem_SelectedIndexChanged(ByVal sender As Object, ByVal e As EventArgs) Handles ddlGrupoOrigem.SelectedIndexChanged
        Try
            CarregarProduto(ddlProdutoOrigem, ddlGrupoOrigem)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub ddlGrupoDestino_SelectedIndexChanged(ByVal sender As Object, ByVal e As EventArgs) Handles ddlGrupoDestino.SelectedIndexChanged
        Try
            chkTabelaProdutoExiste.Visible = False
            CarregarProduto(ddlProdutoDestino, ddlGrupoDestino)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub ddlTabela_SelectedIndexChanged(ByVal sender As Object, ByVal e As EventArgs) Handles ddlTabela.SelectedIndexChanged
        Try
            grdAbaCadastro.DataBind()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    'Replicação
    Protected Sub ddlTabelaOrigem_SelectedIndexChanged(ByVal sender As Object, ByVal e As EventArgs) Handles ddlTabelaOrigem.SelectedIndexChanged
        Try
            CarregarProduto(ddlProdutoOrigem, ddlGrupoOrigem)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub ddlTabelaDestino_SelectedIndexChanged(ByVal sender As Object, ByVal e As EventArgs) Handles ddlTabelaDestino.SelectedIndexChanged
        Try
            CarregarProduto(ddlProdutoDestino, ddlGrupoDestino)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

#End Region

#End Region

#Region "GridView"

    Protected Sub grdAbaCadastro_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdAbaCadastro.SelectedIndexChanged
        Try
            ddlTabela.SelectedValue = grdAbaCadastro.SelectedRow.Cells(1).Text
            ddlAnalise.SelectedValue = grdAbaCadastro.SelectedRow.Cells(4).Text
            lblSequencia.Text = grdAbaCadastro.SelectedRow.Cells(8).Text
            txtFaixaInicial.Text = Server.HtmlDecode(grdAbaCadastro.SelectedRow.Cells(9).Text).Split("á").ToArray(0).Trim
            txtFaixaFinal.Text = Server.HtmlDecode(grdAbaCadastro.SelectedRow.Cells(9).Text).Split("á").ToArray(1).Trim
            txtIndice.Text = grdAbaCadastro.SelectedRow.Cells(10).Text
            txtTolerancia.Text = grdAbaCadastro.SelectedRow.Cells(11).Text
            If (grdAbaCadastro.SelectedRow.Cells(12).Text.Equals("S")) Then
                radSim.Checked = True
            Else
                radNao.Checked = True
            End If
            txtObservacao.Text = Server.HtmlDecode(grdAbaCadastro.SelectedRow.Cells(13).Text).Trim
            ddlTabela.Enabled = False
            ddlAnalise.Enabled = False
            GerenciarBotoes()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

#End Region

#Region "Métodos"

    Private Sub BindCombos()
        ddl.Carregar(ddlTabela, CarregarDDL.Tabela.TabelaDeClassificacoes, "", True)
        ddl.Carregar(ddlAnalise, CarregarDDL.Tabela.Analises, "", True)
        'Origem
        ddl.Carregar(ddlTabelaOrigem, CarregarDDL.Tabela.TabelaDeClassificacoes, "", True)
        ddl.Carregar(ddlGrupoOrigem, CarregarDDL.Tabela.GrupoProduto, "", True)
        'Destino
        ddl.Carregar(ddlTabelaDestino, CarregarDDL.Tabela.TabelaDeClassificacoes, "", True)
        ddl.Carregar(ddlGrupoDestino, CarregarDDL.Tabela.GrupoProduto, "", True)
    End Sub

    Private Sub CarregarProduto(ByVal DropDownList As DropDownList, ByVal ddlGrupoSelecionado As DropDownList)
        ddl.Carregar(DropDownList, CarregarDDL.Tabela.Produto, " GRUPO = '" & ddlGrupoSelecionado.SelectedValue & "'", True)
    End Sub

    Private Sub BindGridView()
        Dim ds As New DataSet

        Dim sql As String

        sql = "SELECT     Convert(varchar(4),Classificacoes.Tabela_Id) as Tabela_Id, " & vbCrLf & _
              "          Classificacoes.Produto_Id, " & vbCrLf & _
              "          Classificacoes.Analise_Id, " & vbCrLf & _
              "          Classificacoes.Sequencia_Id, " & vbCrLf & _
              "          Produtos.Grupo, " & vbCrLf & _
              "          Convert(varchar(6),Classificacoes.FaixaInicial) as FaixaInicial, " & vbCrLf & _
              "          Convert(varchar(6),Classificacoes.FaixaFinal) as FaixaFinal, " & vbCrLf & _
              "          Convert(varchar(8),Classificacoes.Indice) as Indice, " & vbCrLf & _
              "          Convert(varchar(6),Classificacoes.Tolerancia) as Tolerancia, " & vbCrLf & _
              "          Classificacoes.FaixaValida, " & vbCrLf & _
              "          Classificacoes.Observacoes, " & vbCrLf & _
              "          isnull(TabelaDeClassificacoes.Descricao, 0) as DescricaoClassificacao, " & vbCrLf & _
              "          Produtos.Descricao AS DescricaoProduto, " & vbCrLf & _
              "          isnull(Analises.Descricao, '') AS DescricaoAnalise " & vbCrLf & _
              "FROM        Classificacoes LEFT OUTER JOIN " & vbCrLf & _
              "Analises ON Classificacoes.Analise_Id = Analises.Analise_Id LEFT OUTER JOIN " & vbCrLf & _
              "Produtos ON Classificacoes.Produto_Id COLLATE Latin1_General_CI_AS = Produtos.Produto_Id LEFT OUTER JOIN " & vbCrLf & _
              "TabelaDeClassificacoes ON Classificacoes.Tabela_Id = TabelaDeClassificacoes.Codigo_Id " & vbCrLf

        If Not String.IsNullOrWhiteSpace(ddlProduto.SelectedValue) Then
            sql &= " Where Classificacoes.Tabela_Id = '" & ddlTabela.SelectedValue & "'" & vbCrLf
        End If

        If Not String.IsNullOrWhiteSpace(ddlProduto.SelectedValue) Then
            sql &= " AND Classificacoes.Produto_Id = '" & ddlProduto.SelectedValue & "'" & vbCrLf
        End If

        ds = Banco.ConsultaDataSet(sql, "Consulta")

        Dim dr As DataRow

        For Each dr In ds.Tables(0).Rows
            dr("Tabela_Id") = dr("Tabela_Id").ToString
            dr("Grupo") = dr("Grupo").ToString
            dr("DescricaoClassificacao") = Funcoes.AlinharDireita(dr("Tabela_Id"), 4, "0") & "-" & Funcoes.AlinharEsquerda(dr("DescricaoClassificacao"), 10, " ")
            dr("Analise_Id") = dr("Analise_Id").ToString
            dr("DescricaoAnalise") = Format(dr("Analise_Id"), "0000") & "-" & Funcoes.AlinharEsquerda(dr("DescricaoAnalise"), 20, " ")
            dr("Produto_Id") = dr("Produto_Id").ToString
            dr("DescricaoProduto") = Funcoes.AlinharEsquerda(dr("Produto_Id"), 9, " ") & "-" & Funcoes.AlinharEsquerda(dr("DescricaoProduto"), 20, " ")
            dr("FaixaInicial") = Funcoes.AlinharEsquerda(Replace(dr("FaixaInicial"), ".", ","), 5, " ") & " á " & Funcoes.AlinharEsquerda(Replace(dr("FaixaFinal"), ".", ","), 5, " ")
            dr("FaixaFinal") = Funcoes.AlinharEsquerda(Replace(dr("FaixaFinal"), ".", ","), 5, " ")
            dr("Sequencia_Id") = dr("Sequencia_Id").ToString
            dr("Indice") = Funcoes.AlinharEsquerda(Replace(dr("Indice"), ".", ","), 7, " ")
            dr("Tolerancia") = Replace(dr("Tolerancia"), ".", ",")
            dr("FaixaValida") = dr("FaixaValida").ToString
            dr("Observacoes") = Funcoes.EliminarCaracteresEspeciais(dr("Observacoes"))
            ds.AcceptChanges()
        Next
        grdAbaCadastro.DataSource = ds
        grdAbaCadastro.DataBind()

    End Sub

    Private Sub ExecutarIUD(ByVal IUD As String)
        Dim obj As New ClassificacaoDeProduto
        obj.IUD = IUD
        obj.Classificacao.Codigo = ddlTabela.SelectedValue
        obj.Analise.Codigo = ddlAnalise.SelectedValue
        obj.Produto.Codigo = ddlProduto.SelectedValue
        If IUD.Equals("I") Then
            obj.Sequencia = PegarSequencia(obj.Classificacao.Codigo.ToString, obj.Produto.Codigo.ToString, obj.Analise.Codigo.ToString)
        Else
            obj.Sequencia = lblSequencia.Text
        End If

        obj.FaixaInicial = txtFaixaInicial.Text
        obj.FaixaFinal = txtFaixaFinal.Text
        obj.Indide = txtIndice.Text
        obj.Tolerancia = txtTolerancia.Text
        If (radSim.Checked) Then
            obj.FaixaValida = "S"
        Else
            obj.FaixaValida = "N"
        End If
        obj.Observacoes = Funcoes.EliminarCaracteresEspeciais(txtObservacao.Text)
        If Not obj.Salvar() Then
            MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
            Exit Sub
        End If

        If (IUD.Equals("D")) Then
            RefazerSequencia(obj)
        End If

        lblSequencia.Text = String.Empty
        txtFaixaInicial.Text = String.Empty
        txtFaixaFinal.Text = String.Empty
        txtIndice.Text = String.Empty
        txtTolerancia.Text = String.Empty
        radSim.Checked = False
        radNao.Checked = False
        txtObservacao.Text = String.Empty
        ddlAnalise.Enabled = True
        If ddlAnalise.Items.Count > 0 Then ddlAnalise.SelectedIndex = 0

        GerenciarBotoes()

        BindGridView()
    End Sub

    Private Sub Limpar()
        ddlAnalise.ClearSelection()
        ddlTabela.Enabled = True
        ddlAnalise.Enabled = True
        lblSequencia.Text = String.Empty
        txtFaixaInicial.Text = String.Empty
        txtFaixaFinal.Text = String.Empty
        txtIndice.Text = String.Empty
        txtTolerancia.Text = String.Empty
        radSim.Checked = False
        radNao.Checked = False
        txtObservacao.Text = String.Empty
        ddlGrupo.SelectedIndex = 0
        ddlProduto.SelectedIndex = 0
        grdAbaCadastro.DataSource = Nothing
        grdAbaCadastro.DataBind()
        GerenciarBotoes()
        CargaGrupo()
    End Sub

    Private Sub LimparReplicacao()
        'Origem
        ddlTabelaOrigem.SelectedIndex = 0
        ddlGrupoOrigem.SelectedIndex = 0
        ddlProdutoOrigem.Items.Clear()
        'Destino
        ddlTabelaDestino.SelectedIndex = 0
        ddlGrupoDestino.SelectedIndex = 0
        ddlProdutoDestino.Items.Clear()
    End Sub

    Private Function getParam() As String
        Dim param As String = ""
        If Not String.IsNullOrWhiteSpace(ddlTabela.SelectedValue) Then
            param &= "Tabela: " & ddlTabela.SelectedValue & " - "
        End If
        If Not String.IsNullOrWhiteSpace(ddlAnalise.SelectedValue) Then
            param &= "Análise: " & ddlAnalise.SelectedValue & " - "
        End If
        If Not String.IsNullOrWhiteSpace(lblSequencia.Text) Then
            param &= "Sequência: " & lblSequencia.Text & vbCrLf
        End If
        If Not String.IsNullOrWhiteSpace(txtFaixaInicial.Text) Then
            param &= "Faixa: " & txtFaixaInicial.Text & " - "
        End If
        If Not String.IsNullOrWhiteSpace(txtFaixaFinal.Text) Then
            param &= "Á: " & txtFaixaFinal.Text & " - "
        End If
        If Not String.IsNullOrWhiteSpace(txtIndice.Text) Then
            param &= "Índice: " & txtIndice.Text & " - "
        End If
        If Not String.IsNullOrWhiteSpace(txtTolerancia.Text) Then
            param &= "Tolerância: " & txtTolerancia.Text & vbCrLf
        End If
        If Not String.IsNullOrWhiteSpace(txtObservacao.Text) Then
            param &= "Observação: " & txtObservacao.Text & " - "
        End If
        param &= IIf(radSim.Checked, "Faixa Válida: Sim", "Faixa Válida: Não")

        Return param
    End Function

    Protected Sub ddlGrupo_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        ddl.Carregar(ddlProduto, CarregarDDL.Tabela.Produto, " Grupo = '" & ddlGrupo.SelectedValue & "'")
    End Sub

    Private Sub CargaGrupo()
        ddl.Carregar(ddlGrupo, CarregarDDL.Tabela.GrupoProduto, "N")
    End Sub

    Protected Sub ddlProduto_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddlProduto.SelectedIndexChanged
        If ddlProduto.SelectedIndex > 0 Then
            BindGridView()
        End If
    End Sub

    Public Sub Relatorio()
        Dim Sql As String
        Dim ds As New DataSet

        Sql = "SELECT Classificacoes.Tabela_Id, " & vbCrLf & _
              "       Classificacoes.Produto_Id, " & vbCrLf & _
              "       Classificacoes.Analise_Id, " & vbCrLf & _
              "       Classificacoes.Sequencia_Id, " & vbCrLf & _
              "       Classificacoes.FaixaInicial, " & vbCrLf & _
              "       Classificacoes.FaixaFinal, " & vbCrLf & _
              "       Classificacoes.Indice, " & vbCrLf & _
              "       Classificacoes.Tolerancia, " & vbCrLf & _
              "       Classificacoes.FaixaValida, " & vbCrLf & _
              "       Classificacoes.Observacoes, " & vbCrLf & _
              "       TabelaDeClassificacoes.Descricao as DescricaoClassificacao, " & vbCrLf & _
              "       Produtos.Descricao AS DescricaoProduto, " & vbCrLf & _
              "       Analises.Descricao AS DescricaoAnalise " & vbCrLf & _
              "  FROM Classificacoes " & vbCrLf & _
              "  LEFT OUTER JOIN Analises " & vbCrLf & _
              "    ON Classificacoes.Analise_Id = Analises.Analise_Id " & vbCrLf & _
              "  LEFT OUTER JOIN Produtos " & vbCrLf & _
              "    ON Classificacoes.Produto_Id COLLATE Latin1_General_CI_AS = Produtos.Produto_Id " & vbCrLf & _
              "  LEFT OUTER JOIN TabelaDeClassificacoes " & vbCrLf & _
              "    ON Classificacoes.Tabela_Id = TabelaDeClassificacoes.Codigo_Id " & vbCrLf

        If Not String.IsNullOrWhiteSpace(ddlTabela.SelectedValue) Then
            Sql &= " Where Classificacoes.Tabela_Id = '" & ddlTabela.SelectedValue & "'" & vbCrLf
        End If

        If Not String.IsNullOrWhiteSpace(ddlProduto.SelectedValue) Then
            Sql &= " AND Classificacoes.Produto_Id = '" & ddlProduto.SelectedValue & "'" & vbCrLf
        End If

        ds = Banco.ConsultaDataSet(Sql, "ClassificacaoDeProdutos")

        Funcoes.BindReport(Me.Page, ds, "Cr_ClassificacaoDeProdutos", eExportType.PDF)
    End Sub

    Public Function PegarSequencia(ByVal Tabela As String, ByVal Produto As String, ByVal Analise As String)
        Dim ds As New DataSet
        Dim sql As String
        Dim Texto As String

        sql = "select count(*)+1 as sequencia from Classificacoes where Tabela_Id=" & Tabela & vbCrLf & _
              " AND Produto_Id = '" & RTrim(Produto) & "'" & vbCrLf & _
              " AND  Analise_Id=" & Analise & vbCrLf

        Texto = "0"

        For Each Dr As DataRow In Banco.ConsultaDataSet(sql, "Tabelas").Tables(0).Rows
            Texto = Format(Dr("sequencia"), "0000")
        Next
        lblSequencia.Text = Texto
        Return Texto
    End Function

    Public Function ValidaCampos(ByVal IUD As String)
        Dim aux As String = String.Empty
        'Fazer a validação de todos os campos obrigatorios
        If (ddlTabela.SelectedValue.Equals("")) Then
            aux += "O campo Tabela é obrigatório. \n"
        End If

        If (ddlAnalise.SelectedValue.Equals("")) Then
            aux += "O campo Análise é obrigatório. \n"
        End If

        If (ddlProduto.SelectedValue.Equals("")) Then
            aux += "O campo Produto é obrigatório. \n"
        End If

        If (IUD IsNot "D") Then
            If (String.IsNullOrWhiteSpace(txtFaixaInicial.Text) Or String.IsNullOrWhiteSpace(txtFaixaFinal.Text)) Then
                aux += "O campo Faixa é obrigatório. \n"
            End If
            If (String.IsNullOrWhiteSpace(txtIndice.Text)) Then
                aux += "O campo Índice é obrigatório. \n"
            End If
            If (String.IsNullOrWhiteSpace(txtTolerancia.Text)) Then
                aux += "O campo Tolerância é obrigatório. \n"
            End If
            If (Not radSim.Checked And Not radNao.Checked) Then
                aux += "O campo Faixa Válida é obrigatório. \n"
            End If
        End If
        Return aux
    End Function

    Public Function ValidaCamposReplicacao()
        Dim aux As String = String.Empty
        'Origem
        If (ddlTabelaOrigem.SelectedValue.Equals("")) Then
            aux += "O campo Tabela/Origem é obrigatório. \n"
        End If
        If (ddlGrupoOrigem.SelectedValue.Equals("")) Then
            aux += "O campo Grupo/Origem é obrigatório. \n"
        End If
        If (ddlProdutoOrigem.SelectedValue.Equals("")) Then
            aux += "O campo Produto/Origem é obrigatório. \n"
        End If
        'Destino
        If (ddlTabelaDestino.SelectedValue.Equals("")) Then
            aux += "O campo Tabela/Destino é obrigatório. \n"
        End If
        If (ddlGrupoDestino.SelectedValue.Equals("")) Then
            aux += "O campo Grupo/Destino é obrigatório. \n"
        End If
        If (ddlProdutoDestino.SelectedValue.Equals("")) Then
            aux += "O campo Produto/Destino é obrigatório. \n"
        End If
        If (chkTabelaProdutoExiste.Visible And Not chkTabelaProdutoExiste.Checked) Then
            aux += "Marque a opção se deseja regravá-la. \n"
        End If
        Return aux
    End Function

    Public Function ConsultarReplicar(ByVal Tabela As String, ByVal Produto As String) As Boolean
        Dim Texto As String
        Dim Sql As String = "Select count(*) as Existe from Classificacoes WHERE Tabela_Id = "
        Sql &= Tabela & " AND Produto_Id = '" & RTrim(Produto) & "'"

        Texto = ""

        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Classificacoes").Tables(0).Rows
            If Dr("Existe") = "0" Then
                Return False
            Else
                Return True
            End If
        Next
        Return True
    End Function

    Public Function Replicar(ByVal Tabori As String, ByVal Prdori As String, ByVal Tabdes As String, ByVal Prddes As String)
        Dim SqlArray As New ArrayList
        Dim strSQL As String
        Dim Ds As New DataSet

        strSQL = "DELETE FROM  Classificacoes WHERE( " & vbCrLf & _
                 "Tabela_Id = " & Tabdes & vbCrLf & _
                 " AND Produto_Id = '" & RTrim(Prddes) & "')" & vbCrLf

        SqlArray.Add(strSQL)

        strSQL = "Select Tabela_Id, Produto_Id, Analise_Id, Sequencia_Id, FaixaInicial, FaixaFinal, Indice, Tolerancia, FaixaValida, Observacoes from Classificacoes WHERE (Tabela_Id="
        strSQL &= Tabori & " And Produto_Id='" & RTrim(Prdori) & "')"

        Ds = Banco.ConsultaDataSet(strSQL, "Classificacoes")

        If Ds.Tables(0).Rows.Count > 0 Then
            For Each Dr As DataRow In Ds.Tables(0).Rows
                strSQL = "INSERT INTO Classificacoes " & vbCrLf & _
                         "(Tabela_Id, Produto_Id, Analise_Id, Sequencia_Id, " & vbCrLf & _
                         "FaixaInicial, FaixaFinal, " & vbCrLf & _
                         "Indice, Tolerancia, FaixaValida, Observacoes) VALUES (" & vbCrLf & _
                         Tabdes & ", '" & Prddes & "', " & Dr("Analise_Id") & ", " & Dr("Sequencia_Id") & ", " & vbCrLf & _
                         Replace(Dr("FaixaInicial"), ",", ".") & ", " & Replace(Dr("FaixaFinal"), ",", ".") & ", " & vbCrLf & _
                         Replace(Dr("Indice"), ",", ".") & ", " & Replace(Dr("Tolerancia"), ",", ".") & ", '" & Dr("FaixaValida") & "', '" & Dr("Observacoes") & "')" & vbCrLf
                SqlArray.Add(strSQL)
            Next
        End If

        If Banco.GravaBanco(SqlArray) Then
            Return "Replicação concluída."
        Else
            Return ("Erro! " & HttpContext.Current.Session("ssMessage"))
        End If

    End Function

    Public Sub GerenciarBotoes()
        If (String.IsNullOrEmpty(lblSequencia.Text)) Then
            lnkAtualizar.Parent.Visible = False
            lnkExcluir.Parent.Visible = False
            lnkNovo.Parent.Visible = True
        Else
            lnkAtualizar.Parent.Visible = True
            lnkExcluir.Parent.Visible = True
            lnkNovo.Parent.Visible = False
        End If
    End Sub

    Public Sub RefazerSequencia(ByVal obj As ClassificacaoDeProduto)
        Dim Sql As String
        Sql = " UPDATE Classificacoes " & vbCrLf & _
                " SET Sequencia_Id = Sequencia_Id - 1 " & vbCrLf & _
                " WHERE Sequencia_Id > " & obj.Sequencia & vbCrLf & _
                " AND Tabela_Id = " & obj.Classificacao.Codigo & vbCrLf & _
                " AND Produto_Id = " & obj.Produto.Codigo & vbCrLf & _
                " AND Analise_Id = " & obj.Analise.Codigo & vbCrLf
        Banco.GravaBanco(Sql)
    End Sub

#End Region

#Region "Botoes"

    Protected Sub lnkNovo_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkNovo.Click
        Try
            If Funcoes.VerificaPermissao("ClassificacaoDeProdutos", "GRAVAR") Then
                Dim Erro = ValidaCampos("I")
                If (Erro.Equals(String.Empty)) Then
                    ExecutarIUD("I")
                    MsgBox(Me.Page, "Inclusão realizada com Sucesso.", eTitulo.Sucess)
                Else
                    MsgBox(Me.Page, Erro)
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
            Exit Sub
        End Try
    End Sub

    Protected Sub lnkConsultar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkConsultar.Click
        Try
            If String.IsNullOrWhiteSpace(ddlTabela.SelectedValue) Then
                MsgBox(Me.Page, "Informe a Tabela.")
            ElseIf String.IsNullOrWhiteSpace(ddlProduto.SelectedValue) Then
                MsgBox(Me.Page, "Informe o Produto.")
            Else
                BindGridView()
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAtualizar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkAtualizar.Click
        Try
            If Funcoes.VerificaPermissao("ClassificacaoDeProdutos", "ALTERAR") Then
                Dim Erro = ValidaCampos("U")
                If (Erro.Equals(String.Empty)) Then
                    ExecutarIUD("U")
                    MsgBox(Me.Page, "Atualização realizada com Sucesso.", eTitulo.Sucess)
                Else
                    MsgBox(Me.Page, Erro)
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
            Exit Sub
        End Try
    End Sub

    Protected Sub lnkExcluir_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkExcluir.Click
        Try
            If Funcoes.VerificaPermissao("ClassificacaoDeProdutos", "EXCLUIR") Then
                Dim Erro = ValidaCampos("D")
                If (Erro.Equals(String.Empty)) Then
                    ExecutarIUD("D")
                    MsgBox(Me.Page, "Exclusão realizada com Sucesso.", eTitulo.Sucess)
                Else
                    MsgBox(Me.Page, Erro)
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
            Exit Sub
        End Try
    End Sub

    Protected Sub lnkLimpar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkLimpar.Click
        Try
            Limpar()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
            Exit Sub
        End Try
    End Sub

    Protected Sub lnkRelatorio_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkRelatorio.Click
        Try
            If Funcoes.VerificaPermissao("ClassificacaoDeProdutos", "RELATORIO") Then
                If String.IsNullOrWhiteSpace(ddlTabela.SelectedValue) Then
                    MsgBox(Me.Page, "Informe a Tabela.")
                Else
                    Relatorio()
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para emitir relatório.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAjuda_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "ClassificacaoDeProdutos")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub

    Protected Sub btnReplicar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnReplicar.Click
        Try
            Dim Erro = ValidaCamposReplicacao()
            If (Erro.Equals(String.Empty)) Then
                If (Not ConsultarReplicar(ddlTabelaOrigem.SelectedValue, ddlProdutoOrigem.SelectedValue)) Then
                    Erro = "Não há registros na tabela/produto de origem para serem replicados. \n"
                End If
                If (Erro.Equals(String.Empty)) Then
                    Replicar(ddlTabelaOrigem.SelectedValue, ddlProdutoOrigem.SelectedValue, ddlTabelaDestino.SelectedValue, ddlProdutoDestino.SelectedValue)
                    LimparReplicacao()
                    MsgBox(Me.Page, "Replicação realizada com Sucesso.", eTitulo.Sucess)
                Else
                    MsgBox(Me.Page, Erro)
                End If
            Else
                MsgBox(Me.Page, Erro)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
            Exit Sub
        End Try
    End Sub

#End Region

End Class