Imports System.Web
Imports NGS.Lib.Negocio

Public Class ListOrdemParaProducaoXProduto
    Inherits List(Of OrdemParaProducaoXProduto)
    Implements IBaseEntity

#Region "Contrutor"
    Public Sub New()
    End Sub

    Public Sub New(ByVal objOrdemDeProducao As OrdemParaProducao)

        Me.OrdemParaProducao = objOrdemDeProducao

        Dim db As New AcessaBanco()

        Try
            Dim strSQL As String

            strSQL = "SELECT Empresa_Id, EndEmpresa_Id, Ordem_Id, Produto_Id, Quantidade, QuantidadeDeAjuste, Lote, ISNULL(UnidadeComercializacao, '') AS UnidadeComercializacao, ISNULL(FatorConversao, 0) AS FatorConversao " & vbCrLf &
                     "FROM OrdemDeProducaoXProduto " & vbCrLf &
                     "WHERE Empresa_id    = '" & Me.OrdemParaProducao.CodigoEmpresa & "'" & vbCrLf &
                     "  AND EndEmpresa_Id = " & Me.OrdemParaProducao.EnderecoEmpresa & vbCrLf

            If Me.OrdemParaProducao.Codigo > 0 Then strSQL &= "  AND Ordem_Id      = " & Me.OrdemParaProducao.Codigo & vbCrLf

            Dim ds As DataSet = db.ConsultaDataSet(strSQL, "OrdemDeProducaoXProduto")

            For Each row As DataRow In ds.Tables(0).Rows
                Dim oPP As New OrdemParaProducaoXProduto(Me.OrdemParaProducao)
                oPP.CodigoProduto = row("Produto_Id").ToString()
                oPP.Produto = New Produto(row("Produto_Id").ToString())
                oPP.Lote = row("Lote").ToString()
                oPP.Quantidade = row("Quantidade")
                oPP.QuantidadeDeAjuste = row("QuantidadeDeAjuste")
                oPP.CodigoUnidadeComercializacao = row("UnidadeComercializacao")
                oPP.FatorConversao = row("FatorConversao")

                Me.Add(oPP)
            Next

        Catch ex As Exception
            Throw ex
        Finally
            db = Nothing
        End Try
    End Sub

    Public Sub SalvarSql(ByRef Sqls As ArrayList)
        For Each item As OrdemParaProducaoXProduto In Me
            item.SalvarSql(Sqls)
        Next
    End Sub
#End Region

#Region "Fields"
    Private _OrdemParaProducao As OrdemParaProducao
#End Region

#Region "Property"
    Public Property OrdemParaProducao() As OrdemParaProducao
        Get
            Return _OrdemParaProducao
        End Get
        Set(ByVal value As OrdemParaProducao)
            _OrdemParaProducao = value
        End Set
    End Property

#End Region

End Class

<Serializable()>
Public Class OrdemParaProducaoXProduto
    Implements IBaseEntity

#Region "Fields"
    Private _IUD As String = ""

    Private _OrdemParaProducao As OrdemParaProducao

    Private _CodigoProduto As String
    Private _Produto As Produto
    Private _CodigoUnidadeComercializacao As String
    Private _FatorConversao As Decimal
    Private _Lote As String
    Private _Quantidade As Decimal
    Private _QuantidadeDeAjuste As Decimal

#End Region

#Region "Property"
    Public Property IUD() As String
        Get
            Return _IUD
        End Get
        Set(ByVal value As String)
            _IUD = value
        End Set
    End Property

    Public Property OrdemParaProducao() As OrdemParaProducao
        Get
            Return _OrdemParaProducao
        End Get
        Set(ByVal value As OrdemParaProducao)
            _OrdemParaProducao = value
        End Set
    End Property

    Public Property CodigoProduto() As String
        Get
            Return _CodigoProduto
        End Get
        Set(ByVal value As String)
            _CodigoProduto = value
        End Set
    End Property

    Public Property Produto() As Produto
        Get
            If _Produto Is Nothing And Not Me.CodigoProduto Is Nothing Then _Produto = New Produto(Me.CodigoProduto)
            Return _Produto
        End Get
        Set(ByVal value As Produto)
            _Produto = value
        End Set
    End Property

    Public Property CodigoUnidadeComercializacao As String
        Get
            Return _CodigoUnidadeComercializacao
        End Get
        Set(value As String)
            _CodigoUnidadeComercializacao = value
        End Set
    End Property

    Public Property FatorConversao As Decimal
        Get
            Return _FatorConversao
        End Get
        Set(value As Decimal)
            _FatorConversao = value
        End Set
    End Property

    Public Property Lote As String
        Get
            Return _Lote
        End Get
        Set(value As String)
            _Lote = value
        End Set
    End Property

    Public Property Quantidade() As Decimal
        Get
            Return _Quantidade
        End Get
        Set(ByVal value As Decimal)
            _Quantidade = value
        End Set
    End Property

    Public Property QuantidadeDeAjuste() As Decimal
        Get
            Return _QuantidadeDeAjuste
        End Get
        Set(ByVal value As Decimal)
            _QuantidadeDeAjuste = value
        End Set
    End Property

#End Region

#Region "Construtor"
    Public Sub New()
    End Sub

    Public Sub New(ByVal OrdemParaProducao As OrdemParaProducao)
        Me.OrdemParaProducao = OrdemParaProducao
    End Sub
#End Region

#Region "Methods"

    Public Function Salvar() As Boolean
        Dim objBanco As New AcessaBanco()
        Dim sqls As New ArrayList

        SalvarSql(sqls)
        Return objBanco.GravaBanco(sqls)
    End Function

    Public Sub SalvarSql(ByRef Sqls As ArrayList)
        Dim strSql As String = ""

        Select Case Me.IUD
            Case "I"

                Dim produtoXLote = New ProdutoXSequenciaDeLote(Me.CodigoProduto, OrdemParaProducao.Movimento.Year)

                If String.IsNullOrWhiteSpace(produtoXLote.SequenciaDoProduto) Then
                    HttpContext.Current.Session("ssMessage") = "Sequência do Lote do Produto não foi encontrada."
                    Exit Sub
                End If

                Dim seq As Integer = produtoXLote.SequenciaDoLote + 1

                If Left(OrdemParaProducao.CodigoEmpresa, 8) = "40938762" OrElse Left(OrdemParaProducao.CodigoEmpresa, 8) = "49673784" Then
                    Me.Lote = seq
                Else
                    If IsNumeric(produtoXLote.SequenciaDoProduto) Then
                        Me.Lote = CInt(produtoXLote.SequenciaDoProduto).ToString("000") & "." & seq.ToString("0000") & "." & Now.ToString("yyyy")
                    Else
                        Me.Lote = produtoXLote.SequenciaDoProduto & "." & seq.ToString("0000") & "." & Now.ToString("yyyy")
                    End If
                End If

                strSql = "INSERT INTO OrdemDeProducaoXProduto(Empresa_Id, EndEmpresa_Id, Ordem_Id, Produto_Id, Lote, Quantidade, QuantidadeDeAjuste, UnidadeComercializacao, FatorConversao)" & vbCrLf &
                         "VALUES('" & OrdemParaProducao.CodigoEmpresa & "'," & OrdemParaProducao.EnderecoEmpresa & "," & OrdemParaProducao.Codigo & ",'" & Me.CodigoProduto & "','" & Me.Lote & "', " & Str(Me.Quantidade) & "," & Str(Me.QuantidadeDeAjuste) & ", '" & Me.CodigoUnidadeComercializacao & "', " & Str(Me.FatorConversao) & " )"

                Sqls.Add(strSql)

                strSql = " Update ProdutoXSequenciaDeLote set Sequencia = " & seq & vbCrLf &
                          "  WHERE Produto_Id        ='" & Me.CodigoProduto & "'" & vbCrLf &
                          "    AND Ano_Id            = " & produtoXLote.Ano
                Sqls.Add(strSql)

                salvarRelacionados(Sqls)

            Case "U"
                strSql = "UPDATE OrdemDeProducaoXProduto SET " & vbCrLf &
                         "   Quantidade           = " & Str(Me.Quantidade) & vbCrLf &
                         "  ,QuantidadeDeAjuste   = " & Str(Me.QuantidadeDeAjuste) & vbCrLf &
                         " WHERE Empresa_Id    = '" & OrdemParaProducao.CodigoEmpresa & "'" & vbCrLf &
                         "   AND EndEmpresa_Id = " & OrdemParaProducao.EnderecoEmpresa & vbCrLf &
                         "   AND Ordem_Id      = " & OrdemParaProducao.Codigo & vbCrLf &
                         "   AND Produto_Id    = '" & Me.CodigoProduto & "'"

                Sqls.Add(strSql)

                salvarRelacionados(Sqls)

            Case "D"
                salvarRelacionados(Sqls)

                strSql = "DELETE OrdemDeProducaoXProduto " & vbCrLf &
                         " WHERE Empresa_Id    = '" & OrdemParaProducao.CodigoEmpresa & "'" & vbCrLf &
                         "   AND EndEmpresa_Id = " & OrdemParaProducao.EnderecoEmpresa & vbCrLf &
                         "   AND Ordem_Id      = " & OrdemParaProducao.Codigo & vbCrLf &
                         "   AND Produto_Id    = '" & Me.CodigoProduto & "'"

                Sqls.Add(strSql)

        End Select

    End Sub

    Private Sub salvarRelacionados(ByRef sqls As ArrayList)

        GeraEstoqueProducao(sqls)

    End Sub

    Private Sub GeraEstoqueProducao(ByRef sqls As ArrayList)
        Dim strSql As String = ""

        If Me.IUD = "D" Then
            Exit Sub
        End If

        If OrdemParaProducao.Estoque Then
            If OrdemParaProducao.SubOperacaoProducao.EstoqueFisico Then 'Estoque Físico
                strSql = "INSERT INTO Producao " &
                 "(Empresa_Id, EndEmpresa_Id, Deposito_Id, EndDeposito_Id, Produto_Id, Operacao_Id, SubOperacao_Id, " & vbCrLf &
                 "Movimento_Id, FisicoFiscal_Id, ProdutoDerivado_Id, Lote_Id, Classificacao_Id, Embalagem_Id, TipoDeEmbalagem_Id, " & vbCrLf &
                 " CapacidadeEmbalagem_Id, Etapa, Safra, Entradas, Saidas, Observacao, UsuarioInclusao, UsuarioInclusaoData, OrdemDeProducao) " & vbCrLf &
                 "VALUES ('" & OrdemParaProducao.CodigoEmpresa & "', " & OrdemParaProducao.EnderecoEmpresa & ", '" & OrdemParaProducao.CodigoEmpresa & "', " & OrdemParaProducao.EnderecoEmpresa & ", '" & Me.CodigoProduto & "'," & OrdemParaProducao.SubOperacaoProducao.CodigoOperacao & "," & OrdemParaProducao.SubOperacaoProducao.Codigo & vbCrLf &
                 ",'" & OrdemParaProducao.MovimentoEstoque.ToString("yyyy-MM-dd") & "', 1, '', '" & Me.Lote & "', '" & OrdemParaProducao.Codigo & "', 0, ''" & vbCrLf &
                 ", 0, 1, 'NENHUM', " & IIf(Me.QuantidadeDeAjuste > 0, Str(Me.QuantidadeDeAjuste), Str(Me.Quantidade)) & ", 0, '', '" & UsuarioServidor.NomeUsuario & "', GETDATE(), " & OrdemParaProducao.Codigo & ")"

                sqls.Add(strSql)
            End If

            If OrdemParaProducao.SubOperacaoProducao.EstoqueFiscal Then 'Estoque Fiscal
                strSql = "INSERT INTO Producao " &
                 "(Empresa_Id, EndEmpresa_Id, Deposito_Id, EndDeposito_Id, Produto_Id, Operacao_Id, SubOperacao_Id, " & vbCrLf &
                 "Movimento_Id, FisicoFiscal_Id, ProdutoDerivado_Id, Lote_Id, Classificacao_Id, Embalagem_Id, TipoDeEmbalagem_Id, " & vbCrLf &
                 " CapacidadeEmbalagem_Id, Etapa, Safra, Entradas, Saidas, Observacao, UsuarioInclusao, UsuarioInclusaoData, OrdemDeProducao) " & vbCrLf &
                 "VALUES ('" & OrdemParaProducao.CodigoEmpresa & "', " & OrdemParaProducao.EnderecoEmpresa & ", '" & OrdemParaProducao.CodigoEmpresa & "', " & OrdemParaProducao.EnderecoEmpresa & ", '" & Me.CodigoProduto & "'," & OrdemParaProducao.SubOperacaoProducao.CodigoOperacao & "," & OrdemParaProducao.SubOperacaoProducao.Codigo & vbCrLf &
                 ",'" & OrdemParaProducao.MovimentoEstoque.ToString("yyyy-MM-dd") & "', 2, '', '" & Me.Lote & "', '" & OrdemParaProducao.Codigo & "', 0, ''" & vbCrLf &
                 ", 0, 1, 'NENHUM', " & IIf(Me.QuantidadeDeAjuste > 0, Str(Me.QuantidadeDeAjuste), Str(Me.Quantidade)) & ", 0, '', '" & UsuarioServidor.NomeUsuario & "', GETDATE(), " & OrdemParaProducao.Codigo & ")"

                sqls.Add(strSql)
            End If
        End If

    End Sub

#End Region

End Class
