Imports NGS.Lib.Negocio

Public Class RelacaoDeProducao
    Inherits BasePage

#Region "Events"

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.setMenu(eModulo.Gerencial)
            If Not IsPostBack And IsConnect Then
                If Funcoes.VerificaPermissao("RelacaoDeProducao", "ACESSAR") Then
                    txtDataInicial.Text = String.Format("01/01/{0}", Year(Now))
                    txtDataFinal.Text = Format(Today, "dd/MM/yyyy")
                    CargaUnidade()
                    LiberaEmpresa()
                Else
                    MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Gerencial.aspx")
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub DdlUnidade_SelectedIndexChanged(sender As Object, e As EventArgs) Handles DdlUnidade.SelectedIndexChanged
        Try
            CargaEmpresas()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkProcessar_Click(sender As Object, e As EventArgs) Handles lnkProcessar.Click
        Try
            If ValidaCampos() Then
                Dim ds As DataSet = getDataSet()

                Funcoes.BindExcelOffice(Me.Page, ds, "Analisador")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkLimpar_Click(sender As Object, e As EventArgs) Handles lnkLimpar.Click
        Try
            txtDataInicial.Text = String.Format("01/01/{0}", Year(Now))
            txtDataFinal.Text = Format(Today, "dd/MM/yyyy")
            LiberaEmpresa()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "RelacaoDeProducao")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

#End Region

#Region "Methods"

    Private Function ValidaCampos() As Boolean
        Try
            If String.IsNullOrWhiteSpace(DdlEmpresa.SelectedValue) Then
                MsgBox(Me.Page, "Informe a empresa.")
                Return False
            ElseIf String.IsNullOrWhiteSpace(txtDataInicial.Text) OrElse String.IsNullOrWhiteSpace(txtDataFinal.Text) Then
                MsgBox(Me.Page, "Informe o período.")
                Return False
            End If
            Return True
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Function

    Private Sub CargaUnidade()
        Try
            ddl.Carregar(DdlUnidade, CarregarDDL.Tabela.UnidadeDeNegocio, "", True)
            Funcoes.VerificaUnidadeDeNegocio(DdlUnidade, DdlEmpresa)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub CargaEmpresas()
        Try
            ddl.Carregar(DdlEmpresa, CarregarDDL.Tabela.Empresas, DdlUnidade.SelectedValue, True)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Function getDataSet() As DataSet
        Dim sql As String = String.Empty

        sql = " SELECT     Empresa.Reduzido AS RedEmpresa, Producao.Empresa_Id AS CNPJEmpresa, Producao.EndEmpresa_Id AS EndEmpresa," & vbCrLf &
              "                  Empresa.Nome AS NomeEmpresa, Empresa.Cidade AS CidadeEmpresa, Empresa.Estado AS UFEmpresa, Deposito.Reduzido AS RedDeposito," & vbCrLf &
              "                  Producao.Deposito_Id AS Deposito, Producao.EndDeposito_Id AS EndDeposito, Deposito.Nome AS NomeDeposito," & vbCrLf &
              "                  Deposito.Cidade AS CidadeDeposito, Deposito.Estado AS UFDeposito, Producao.Produto_Id AS Produto, Produtos.Nome AS NomeProduto," & vbCrLf &
              "                  Producao.Operacao_Id AS Operacao, Producao.SubOperacao_Id AS SubOperacao, Operacao.Descricao AS NomeOperacao," & vbCrLf &
              "                  Producao.ProdutoDerivado_Id as ProdutoDerivado, ProdutoDerivado.Nome AS NomeProdutoDerivado, Producao.Movimento_Id As Movimento, " & vbCrLf &
              "                  Producao.Entradas, Producao.Saidas,  isnull(Producao.Observacao, '') as Observacao, " & vbCrLf &
              "                  SUM(CASE WHEN ProducaoXAnalises.Analise_ID = 101 THEN ProducaoXAnalises.Indice ELSE 0 END) AS Oleo," & vbCrLf &
              "                  SUM(CASE WHEN ProducaoXAnalises.Analise_ID = 102 THEN ProducaoXAnalises.Indice ELSE 0 END) AS Umidade," & vbCrLf &
              "                  SUM(CASE WHEN ProducaoXAnalises.Analise_ID = 103 THEN ProducaoXAnalises.Indice ELSE 0 END) AS MateriaGraxa," & vbCrLf &
              "                  SUM(CASE WHEN ProducaoXAnalises.Analise_ID = 104 THEN ProducaoXAnalises.Indice ELSE 0 END) AS ProteinaTotal," & vbCrLf &
              "                  SUM(CASE WHEN ProducaoXAnalises.Analise_ID = 105 THEN ProducaoXAnalises.Indice ELSE 0 END) AS MateriaMineral," & vbCrLf &
              "                  SUM(CASE WHEN ProducaoXAnalises.Analise_ID = 106 THEN ProducaoXAnalises.Indice ELSE 0 END) AS Fibras," & vbCrLf &
              "                  SUM(CASE WHEN ProducaoXAnalises.Analise_ID = 107 THEN ProducaoXAnalises.Indice ELSE 0 END) AS Impurezas," & vbCrLf &
              "                  SUM(CASE WHEN ProducaoXAnalises.Analise_ID = 108 THEN ProducaoXAnalises.Indice ELSE 0 END) AS Ardidos," & vbCrLf &
              "                  SUM(CASE WHEN ProducaoXAnalises.Analise_ID = 109 THEN ProducaoXAnalises.Indice ELSE 0 END) AS Avariados," & vbCrLf &
              "                  SUM(CASE WHEN ProducaoXAnalises.Analise_ID = 110 THEN ProducaoXAnalises.Indice ELSE 0 END) AS AtividadeUreatica," & vbCrLf &
              "                  SUM(CASE WHEN ProducaoXAnalises.Analise_ID = 111 THEN ProducaoXAnalises.Indice ELSE 0 END) AS Acidez," & vbCrLf &
              "                  SUM(CASE WHEN ProducaoXAnalises.Analise_ID = 113 THEN ProducaoXAnalises.Indice ELSE 0 END) AS Saboes," & vbCrLf &
              "                  SUM(CASE WHEN ProducaoXAnalises.Analise_ID = 119 THEN ProducaoXAnalises.Indice ELSE 0 END) AS OleoBaseSeca," & vbCrLf &
              "                  SUM(CASE WHEN ProducaoXAnalises.Analise_ID = 125 THEN ProducaoXAnalises.Indice ELSE 0 END) AS Lex," & vbCrLf &
              "                  SUM(CASE WHEN ProducaoXAnalises.Analise_ID = 128 THEN ProducaoXAnalises.Indice ELSE 0 END) AS Farelo," & vbCrLf &
              "                  SUM(CASE WHEN ProducaoXAnalises.Analise_ID = 132 THEN ProducaoXAnalises.Indice ELSE 0 END) AS Profat " & vbCrLf &
              "   FROM           SubOperacoes AS Operacao INNER JOIN" & vbCrLf &
              "                  Produtos INNER JOIN" & vbCrLf &
              "                  Clientes AS Empresa INNER JOIN" & vbCrLf &
              "                  Producao ON Empresa.Cliente_Id = Producao.Empresa_Id AND Empresa.Endereco_Id = Producao.EndEmpresa_Id INNER JOIN" & vbCrLf &
              "                  Clientes AS Deposito ON Producao.Deposito_Id = Deposito.Cliente_Id AND Producao.EndDeposito_Id = Deposito.Endereco_Id ON" & vbCrLf &
              "                  Produtos.Produto_Id = Producao.Produto_Id ON Operacao.Operacao_Id = Producao.Operacao_Id AND" & vbCrLf &
              "                  Operacao.SubOperacoes_Id = Producao.SubOperacao_Id LEFT OUTER JOIN" & vbCrLf &
              "                  ProducaoXAnalises ON Producao.Empresa_Id = ProducaoXAnalises.Empresa_Id AND" & vbCrLf &
              "                  Producao.EndEmpresa_Id = ProducaoXAnalises.EndEmpresa_Id AND Producao.Deposito_Id = ProducaoXAnalises.Deposito_Id AND" & vbCrLf &
              "                  Producao.EndDeposito_Id = ProducaoXAnalises.EndDeposito_Id AND Producao.Produto_Id = ProducaoXAnalises.Produto_Id AND" & vbCrLf &
              "                  Producao.Operacao_Id = ProducaoXAnalises.Operacao_Id AND Producao.SubOperacao_Id = ProducaoXAnalises.SubOperacao_Id AND" & vbCrLf &
              "                  Producao.Movimento_Id = ProducaoXAnalises.Movimento_Id AND Producao.FisicoFiscal_Id = ProducaoXAnalises.FisicoFiscal_Id AND" & vbCrLf &
              "                  Producao.ProdutoDerivado_Id = ProducaoXAnalises.ProdutoDerivado_Id LEFT OUTER JOIN" & vbCrLf &
              "                  Produtos AS ProdutoDerivado ON Producao.ProdutoDerivado_Id = ProdutoDerivado.Produto_Id" & vbCrLf &
              " WHERE            (Producao.Empresa_Id = '" & DdlEmpresa.SelectedValue.Split("-")(0) & "')" & vbCrLf &
              "                  And (Producao.Movimento_Id  between '" & txtDataInicial.Text.ToSqlDate() & "' AND '" & txtDataFinal.Text.ToSqlDate() & "') " & vbCrLf &
              "   GROUP BY Empresa.Reduzido, Producao.Empresa_Id, Producao.EndEmpresa_Id, Empresa.Nome, Empresa.Cidade, Empresa.Estado, Deposito.Reduzido," & vbCrLf &
              "                  Producao.Deposito_Id, Producao.EndDeposito_Id, Deposito.Nome, Deposito.Cidade, Deposito.Estado, Producao.Produto_Id, Produtos.Nome," & vbCrLf &
              "                  Producao.Operacao_Id, Producao.SubOperacao_Id, Operacao.Descricao, Producao.ProdutoDerivado_Id, ProdutoDerivado.Nome," & vbCrLf &
              "                  Producao.Movimento_Id , Producao.Entradas, Producao.Saidas, Producao.Observacao" & vbCrLf &
              "   Order By       Producao.Movimento_Id " & vbCrLf

        Return Banco.ConsultaDataSet(sql, "Analisador")
    End Function

    Private Sub LiberaEmpresa()
        If Not UsuarioServidor.LiberaEmpresa Then
            DdlUnidade.Enabled = False
            DdlEmpresa.Enabled = False
        End If
    End Sub

#End Region

End Class