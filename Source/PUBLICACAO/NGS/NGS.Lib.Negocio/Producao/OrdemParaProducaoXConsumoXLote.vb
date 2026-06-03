Public Class ListOrdemParaProducaoXConsumoXLote
    Inherits List(Of OrdemParaProducaoXConsumoXLote)
    Implements IBaseEntity

#Region "Contrutor"
    Public Sub New()
    End Sub

    Public Sub New(ByVal objOrdemDeProducaoXConsumo As OrdemParaProducaoXConsumo)

        Me.OrdemParaProducaoXConsumo = objOrdemDeProducaoXConsumo

        Dim db As New AcessaBanco()

        Try
            Dim strSQL As String

            strSQL = "Select Empresa_Id, EndEmpresa_Id, Ordem_Id, ProdutoProducao_Id, Produto_Id, Lote_Id, Quantidade, Validade " & vbCrLf &
                     "From OrdemDeProducaoXConsumoXLote " & vbCrLf &
                     "Where Empresa_id    = '" & Me.OrdemParaProducaoXConsumo.OrdemParaProducao.CodigoEmpresa & "'" & vbCrLf &
                     "  and EndEmpresa_Id = " & Me.OrdemParaProducaoXConsumo.OrdemParaProducao.EnderecoEmpresa & vbCrLf

            If Me.OrdemParaProducaoXConsumo.OrdemParaProducao.Codigo > 0 Then strSQL &= "  and Ordem_Id      = " & Me.OrdemParaProducaoXConsumo.OrdemParaProducao.Codigo & vbCrLf

            If Me.OrdemParaProducaoXConsumo.CodigoProduto.Length > 0 Then strSQL &= "  and Produto_Id     = '" & Me.OrdemParaProducaoXConsumo.CodigoProduto & "'" & vbCrLf

            If Me.OrdemParaProducaoXConsumo.CodigoProdutoProducao.Length > 0 Then strSQL &= "  and ProdutoProducao_Id     = '" & Me.OrdemParaProducaoXConsumo.CodigoProdutoProducao & "'" & vbCrLf

            Dim ds As DataSet = db.ConsultaDataSet(strSQL, "OrdemParaProducaoXConsumoXLote")

            For Each row As DataRow In ds.Tables(0).Rows
                Dim oPC As New OrdemParaProducaoXConsumoXLote(Me.OrdemParaProducaoXConsumo)
                oPC.Lote = row("Lote_Id")
                oPC.Quantidade = row("Quantidade")
                oPC.Validade = row("Validade")

                Me.Add(oPC)
            Next

        Catch ex As Exception
            Throw ex
        Finally
            db = Nothing
        End Try
    End Sub

    Public Sub SalvarSql(ByRef Sqls As ArrayList)
        For Each item As OrdemParaProducaoXConsumoXLote In Me
            item.SalvarSql(Sqls)
        Next
    End Sub
#End Region

#Region "Fields"
    Private _OrdemParaProducaoXConsumo As OrdemParaProducaoXConsumo
#End Region

#Region "Property"
    Public Property OrdemParaProducaoXConsumo() As OrdemParaProducaoXConsumo
        Get
            Return _OrdemParaProducaoXConsumo
        End Get
        Set(ByVal value As OrdemParaProducaoXConsumo)
            _OrdemParaProducaoXConsumo = value
        End Set
    End Property

#End Region

End Class

<Serializable()> _
Public Class OrdemParaProducaoXConsumoXLote
    Implements IBaseEntity

#Region "Fields"
    Private _IUD As String = ""

    Private _OrdemParaProducaoXConsumo As OrdemParaProducaoXConsumo

    Private _Lote As String
    Private _Quantidade As Decimal
    Private _Validade As DateTime = Date.Today

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

    Public Property OrdemParaProducaoXConsumo() As OrdemParaProducaoXConsumo
        Get
            Return _OrdemParaProducaoXConsumo
        End Get
        Set(ByVal value As OrdemParaProducaoXConsumo)
            _OrdemParaProducaoXConsumo = value
        End Set
    End Property

    Public Property Lote() As String
        Get
            Return _Lote
        End Get
        Set(ByVal value As String)
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

    Public Property Validade() As DateTime
        Get
            Return _Validade
        End Get
        Set(ByVal value As DateTime)
            _Validade = value
        End Set
    End Property

#End Region

#Region "Construtor"
    Public Sub New()
    End Sub

    Public Sub New(ByVal OrdemParaProducaoXConsumo As OrdemParaProducaoXConsumo)
        Me.OrdemParaProducaoXConsumo = OrdemParaProducaoXConsumo
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

                strSql = "INSERT INTO OrdemDeProducaoXConsumoXLote(Empresa_Id, EndEmpresa_Id, Ordem_Id, ProdutoProducao_Id, Produto_Id, Lote_Id, Quantidade, Validade)" & vbCrLf &
                         "Values('" & OrdemParaProducaoXConsumo.OrdemParaProducao.CodigoEmpresa & "'," & OrdemParaProducaoXConsumo.OrdemParaProducao.EnderecoEmpresa & "," & OrdemParaProducaoXConsumo.OrdemParaProducao.Codigo & ",'" & OrdemParaProducaoXConsumo.CodigoProdutoProducao & "','" & OrdemParaProducaoXConsumo.CodigoProduto & "','" & Me.Lote & "'," & Str(Me.Quantidade) & ",'" & Me.Validade.ToString("yyyy-MM-dd") & "')"

            Case "U"

                If Me.Quantidade > 0 Then
                    strSql = "Update OrdemDeProducaoXConsumoXLote Set " & vbCrLf &
                             "   Quantidade = " & Str(Me.Quantidade) & vbCrLf &
                             "  ,Validade   = '" & Me.Validade.ToString("yyyy-MM-dd") & "'" & vbCrLf &
                             " Where Empresa_Id    = '" & OrdemParaProducaoXConsumo.OrdemParaProducao.CodigoEmpresa & "'" & vbCrLf &
                             "   and EndEmpresa_Id = " & OrdemParaProducaoXConsumo.OrdemParaProducao.EnderecoEmpresa & vbCrLf &
                             "   and Ordem_Id      = " & OrdemParaProducaoXConsumo.OrdemParaProducao.Codigo & vbCrLf &
                             "   and ProdutoProducao_Id    = '" & OrdemParaProducaoXConsumo.CodigoProdutoProducao & "'" & vbCrLf &
                             "   and Produto_Id    = '" & OrdemParaProducaoXConsumo.CodigoProduto & "'" & vbCrLf &
                             "   and Lote_Id       = '" & Me.Lote & "'"
                Else
                    strSql = "Delete OrdemDeProducaoXConsumoXLote " & vbCrLf &
                             " Where Empresa_Id    = '" & OrdemParaProducaoXConsumo.OrdemParaProducao.CodigoEmpresa & "'" & vbCrLf &
                             "   and EndEmpresa_Id = " & OrdemParaProducaoXConsumo.OrdemParaProducao.EnderecoEmpresa & vbCrLf &
                             "   and Ordem_Id      = " & OrdemParaProducaoXConsumo.OrdemParaProducao.Codigo & vbCrLf &
                             "   and ProdutoProducao_Id    = '" & OrdemParaProducaoXConsumo.CodigoProdutoProducao & "'" & vbCrLf &
                             "   and Produto_Id    = '" & OrdemParaProducaoXConsumo.CodigoProduto & "'" & vbCrLf &
                             "   and Lote_Id       = '" & Me.Lote & "'"
                End If

            Case "D"
                strSql = "Delete OrdemDeProducaoXConsumoXLote " & vbCrLf &
                         " Where Empresa_Id    = '" & OrdemParaProducaoXConsumo.OrdemParaProducao.CodigoEmpresa & "'" & vbCrLf &
                         "   and EndEmpresa_Id = " & OrdemParaProducaoXConsumo.OrdemParaProducao.EnderecoEmpresa & vbCrLf &
                         "   and Ordem_Id      = " & OrdemParaProducaoXConsumo.OrdemParaProducao.Codigo & vbCrLf &
                         "   and ProdutoProducao_Id    = '" & OrdemParaProducaoXConsumo.CodigoProdutoProducao & "'" & vbCrLf &
                         "   and Produto_Id    = '" & OrdemParaProducaoXConsumo.CodigoProduto & "'" & vbCrLf &
                         "   and Lote_Id       = '" & Me.Lote & "'"

            Case "DEL_CONS"

                strSql = "Delete OrdemDeProducaoXConsumoXLote " & vbCrLf &
                         " Where Empresa_Id    = '" & OrdemParaProducaoXConsumo.OrdemParaProducao.CodigoEmpresa & "'" & vbCrLf &
                         "   and EndEmpresa_Id = " & OrdemParaProducaoXConsumo.OrdemParaProducao.EnderecoEmpresa & vbCrLf &
                         "   and Ordem_Id      = " & OrdemParaProducaoXConsumo.OrdemParaProducao.Codigo & vbCrLf &
                         "   and ProdutoProducao_Id    = '" & OrdemParaProducaoXConsumo.CodigoProdutoProducao & "'" & vbCrLf &
                         "   and Produto_Id    = '" & OrdemParaProducaoXConsumo.CodigoProduto & "'" & vbCrLf &
                         "   and Lote_Id       = '" & Me.Lote & "'"

        End Select

        Sqls.Add(strSql)

        If Me.IUD <> "D" And Me.IUD.Length > 0 Then

            If OrdemParaProducaoXConsumo.OrdemParaProducao.Estoque AndAlso OrdemParaProducaoXConsumo.Produto.ControlarEstoque Then GeraEstoqueConsumo(Sqls)

        End If

    End Sub

    Private Sub GeraEstoqueConsumo(ByRef sqls As ArrayList)

        Dim strSql As String = ""

        If Me.IUD = "DEL_CONS" Then

            strSql = "Delete Producao " & vbCrLf &
                 " Where Empresa_Id         = '" & OrdemParaProducaoXConsumo.OrdemParaProducao.CodigoEmpresa & "'" & vbCrLf &
                 "   and EndEmpresa_Id      = " & OrdemParaProducaoXConsumo.OrdemParaProducao.EnderecoEmpresa & vbCrLf &
                 "   and OrdemDeProducao    = " & OrdemParaProducaoXConsumo.OrdemParaProducao.Codigo & vbCrLf &
                 "   and Movimento_Id       = '" & OrdemParaProducaoXConsumo.OrdemParaProducao.MovimentoEstoque.ToString("yyyy-MM-dd") & "'" & vbCrLf &
                 "   and Produto_Id         = '" & OrdemParaProducaoXConsumo.CodigoProduto & "'" & vbCrLf &
                 "   and Lote_Id            = '" & Me.Lote & "'"

            sqls.Add(strSql)

        Else

            If OrdemParaProducaoXConsumo.SubOperacaoConsumo.EstoqueFisico AndAlso Me.Quantidade > 0 Then 'Estoque Físico
                strSql = "INSERT INTO Producao " &
                         "(Empresa_Id, EndEmpresa_Id, Deposito_Id, EndDeposito_Id, Produto_Id, Operacao_Id, SubOperacao_Id, " & vbCrLf &
                         "Movimento_Id, FisicoFiscal_Id, ProdutoDerivado_Id, Lote_Id, Classificacao_Id, Embalagem_Id, TipoDeEmbalagem_Id, " & vbCrLf &
                         " CapacidadeEmbalagem_Id, Etapa, Safra, Entradas, Saidas, Observacao, UsuarioInclusao, UsuarioInclusaoData, OrdemDeProducao) " & vbCrLf &
                         "VALUES ('" & OrdemParaProducaoXConsumo.OrdemParaProducao.CodigoEmpresa & "', " & OrdemParaProducaoXConsumo.OrdemParaProducao.EnderecoEmpresa & ", '" & OrdemParaProducaoXConsumo.OrdemParaProducao.CodigoEmpresa & "', " & OrdemParaProducaoXConsumo.OrdemParaProducao.EnderecoEmpresa & ", " & vbCrLf &
                         "'" & OrdemParaProducaoXConsumo.CodigoProduto & "', " & OrdemParaProducaoXConsumo.SubOperacaoConsumo.CodigoOperacao & ", " & OrdemParaProducaoXConsumo.SubOperacaoConsumo.Codigo & ", '" & OrdemParaProducaoXConsumo.OrdemParaProducao.MovimentoEstoque.ToString("yyyy-MM-dd") & "', " & vbCrLf &
                         " 1, '" & OrdemParaProducaoXConsumo.CodigoProdutoProducao & "', '" & Me.Lote & "', '" & OrdemParaProducaoXConsumo.OrdemParaProducao.Codigo & "', 0, '', 0, 1, 'NENHUM', 0, " & Str(Me.Quantidade) & ", '', " & vbCrLf &
                         "'" & UsuarioServidor.NomeUsuario & "', GETDATE(), " & OrdemParaProducaoXConsumo.OrdemParaProducao.Codigo & ")"

                sqls.Add(strSql)
            End If

            If OrdemParaProducaoXConsumo.SubOperacaoConsumo.EstoqueFiscal AndAlso Me.Quantidade > 0 Then 'Estoque Fiscal
                strSql = "INSERT INTO Producao " &
                         "(Empresa_Id, EndEmpresa_Id, Deposito_Id, EndDeposito_Id, Produto_Id, Operacao_Id, SubOperacao_Id, " & vbCrLf &
                         "Movimento_Id, FisicoFiscal_Id, ProdutoDerivado_Id, Lote_Id, Classificacao_Id, Embalagem_Id, TipoDeEmbalagem_Id, " & vbCrLf &
                         " CapacidadeEmbalagem_Id, Etapa, Safra, Entradas, Saidas, Observacao, UsuarioInclusao, UsuarioInclusaoData, OrdemDeProducao) " & vbCrLf &
                         "VALUES ('" & OrdemParaProducaoXConsumo.OrdemParaProducao.CodigoEmpresa & "', " & OrdemParaProducaoXConsumo.OrdemParaProducao.EnderecoEmpresa & ", '" & OrdemParaProducaoXConsumo.OrdemParaProducao.CodigoEmpresa & "', " & OrdemParaProducaoXConsumo.OrdemParaProducao.EnderecoEmpresa & ", " & vbCrLf &
                         "'" & OrdemParaProducaoXConsumo.CodigoProduto & "', " & OrdemParaProducaoXConsumo.SubOperacaoConsumo.CodigoOperacao & ", " & OrdemParaProducaoXConsumo.SubOperacaoConsumo.Codigo & ", '" & OrdemParaProducaoXConsumo.OrdemParaProducao.MovimentoEstoque.ToString("yyyy-MM-dd") & "', " & vbCrLf &
                         " 2, '" & OrdemParaProducaoXConsumo.CodigoProdutoProducao & "', '" & Me.Lote & "', '" & OrdemParaProducaoXConsumo.OrdemParaProducao.Codigo & "', 0, '', 0, 1, 'NENHUM', 0, " & Str(Me.Quantidade) & ", '', " & vbCrLf &
                         "'" & UsuarioServidor.NomeUsuario & "', GETDATE(), " & OrdemParaProducaoXConsumo.OrdemParaProducao.Codigo & ")"

                sqls.Add(strSql)
            End If

        End If

    End Sub
#End Region

End Class
