Imports System.Web
Imports NGS.Lib.Negocio

<Serializable()>
Public Class ListOrdemParaProducao
    Inherits List(Of OrdemParaProducao)

#Region "Construtor"
    Public Sub New()

    End Sub

    Public Sub New(ByVal Empresa As String, ByVal EndEmpresa As String, Optional ByVal CodigoProduto As String = "", Optional ByVal Where As String = "")
        Dim objBanco As New AcessaBanco()

        Try
            Dim strSQL As String = "SELECT Empresa_Id, EndEmpresa_Id, Ordem_Id, Movimento, Validade, Observacoes, Situacao, Estoque, " & vbCrLf &
                                   "       isnull(UsuarioInclusao,'') as UsuarioInclusao, isnull(UsuarioInclusaoData,GetDate()) as UsuarioInclusaoData," & vbCrLf &
                                   "       isnull(UsuarioAlteracao,'') as UsuarioAlteracao, isnull(UsuarioAlteracaoData,GetDate()) as UsuarioAlteracaoData," & vbCrLf &
                                   "       isnull(UsuarioCancelamento,'') as UsuarioCancelamento, isnull(UsuarioCancelamentoData,GetDate()) as UsuarioCancelamentoData " & vbCrLf &
                       "  FROM OrdemDeProducao WHERE 1 = 1 " & vbCrLf

            If Empresa.Length > 0 Then
                strSQL &= "  AND Empresa_Id    = '" & Empresa & "' " & vbCrLf &
                          "  AND EndEmpresa_Id = " & EndEmpresa & " " & vbCrLf
            End If

            If CodigoProduto.Length > 0 Then
                strSQL &= "  AND Produto       = '" & CodigoProduto & "' " & vbCrLf
            End If

            If Where.Length > 0 Then
                strSQL &= "  AND " & Where
            End If

            Dim dsOP As DataSet = objBanco.ConsultaDataSet(strSQL, "OrdemDeProducao")
            For Each row As DataRow In dsOP.Tables(0).Rows
                Dim op As New OrdemParaProducao

                op.CodigoEmpresa = row("Empresa_Id")
                op.EnderecoEmpresa = row("EndEmpresa_Id")
                op.Codigo = row("Ordem_Id")
                op.Movimento = CDate(row("Movimento"))
                op.Validade = CDate(row("Validade"))
                op.Observacoes = row("Observacoes")
                op.CodigoSituacao = row("Situacao")
                op.Estoque = row("Estoque")
                op.UsuarioInclusao = row("UsuarioInclusao")
                op.DataInclusao = row("UsuarioInclusaoData")
                op.UsuarioAlteracao = row("UsuarioAlteracao")
                op.DataAlteracao = row("UsuarioAlteracaoData")
                op.UsuarioCancelamento = row("UsuarioCancelamento")
                op.DataCancelamento = row("UsuarioCancelamentoData")

                Me.Add(op)
            Next
        Catch ex As Exception
        Finally
            objBanco = Nothing
        End Try
    End Sub
#End Region

End Class

<Serializable()>
Public Class OrdemParaProducao
    Implements IBaseEntity

#Region "Construtor"
    Public Sub New()

    End Sub

    Public Sub New(Empresa As String, ByVal EndEmpresa As String, ByVal Ordem As Integer)
        Dim objBanco As New AcessaBanco()

        Try
            Dim strSQL As String = "SELECT Empresa_Id, EndEmpresa_Id, Ordem_Id, Movimento, Validade, Observacoes, Situacao, Estoque, " & vbCrLf &
                                   "       isnull(UsuarioInclusao,'') as UsuarioInclusao, isnull(UsuarioInclusaoData,GetDate()) as UsuarioInclusaoData," & vbCrLf &
                                   "       isnull(UsuarioAlteracao,'') as UsuarioAlteracao, isnull(UsuarioAlteracaoData,GetDate()) as UsuarioAlteracaoData," & vbCrLf &
                                   "       isnull(UsuarioCancelamento,'') as UsuarioCancelamento, isnull(UsuarioCancelamentoData,GetDate()) as UsuarioCancelamentoData " & vbCrLf &
                                   "FROM OrdemDeProducao " & vbCrLf &
                                   "WHERE Empresa_Id    = '" & Empresa & "' " & vbCrLf &
                                   "  AND EndEmpresa_Id = " & EndEmpresa & " " & vbCrLf &
                                   "  AND Ordem_Id      = " & Ordem

            Dim dsOP As DataSet = objBanco.ConsultaDataSet(strSQL, "OrdemDeProducao")
            For Each row As DataRow In dsOP.Tables(0).Rows

                _CodigoEmpresa = row("Empresa_Id")
                _EnderecoEmpresa = row("EndEmpresa_Id")
                _Codigo = row("Ordem_Id")
                _Movimento = CDate(row("Movimento"))
                _Validade = CDate(row("Validade"))
                _Observacoes = row("Observacoes")
                _CodigoSituacao = row("Situacao")
                _Estoque = row("Estoque")
                _UsuarioInclusao = row("UsuarioInclusao")
                _DataInclusao = row("UsuarioInclusaoData")
                _UsuarioAlteracao = row("UsuarioAlteracao")
                _DataAlteracao = row("UsuarioAlteracaoData")
                _UsuarioCancelamento = row("UsuarioCancelamento")
                _DataCancelamento = row("UsuarioCancelamentoData")

            Next
        Catch ex As Exception
        Finally
            objBanco = Nothing
        End Try
    End Sub

    Public Sub New(Empresa As String, ByVal EndEmpresa As String, ByVal Lote As String)
        Dim objBanco As New AcessaBanco()

        Try
            Dim strSQL As String = "SELECT Empresa_Id, EndEmpresa_Id, Ordem_Id, Movimento, Validade, Observacoes, Situacao, Estoque, " & vbCrLf &
                                   "       isnull(UsuarioInclusao,'') as UsuarioInclusao, isnull(UsuarioInclusaoData,GetDate()) as UsuarioInclusaoData," & vbCrLf &
                                   "       isnull(UsuarioAlteracao,'') as UsuarioAlteracao, isnull(UsuarioAlteracaoData,GetDate()) as UsuarioAlteracaoData," & vbCrLf &
                                   "       isnull(UsuarioCancelamento,'') as UsuarioCancelamento, isnull(UsuarioCancelamentoData,GetDate()) as UsuarioCancelamentoData " & vbCrLf &
                                   "FROM OrdemDeProducao " & vbCrLf &
                                   "WHERE Empresa_Id    = '" & Empresa & "' " & vbCrLf &
                                   "  AND EndEmpresa_Id = " & EndEmpresa & vbCrLf &
                                   "  AND Lote          = '" & Lote & "'"

            Dim dsOP As DataSet = objBanco.ConsultaDataSet(strSQL, "OrdemDeProducao")
            For Each row As DataRow In dsOP.Tables(0).Rows

                _CodigoEmpresa = row("Empresa_Id")
                _EnderecoEmpresa = row("EndEmpresa_Id")
                _Codigo = row("Ordem_Id")
                _Movimento = CDate(row("Movimento"))
                _Validade = CDate(row("Validade"))
                _Observacoes = row("Observacoes")
                _CodigoSituacao = row("Situacao")
                _Estoque = row("Estoque")
                _UsuarioInclusao = row("UsuarioInclusao")
                _DataInclusao = row("UsuarioInclusaoData")
                _UsuarioAlteracao = row("UsuarioAlteracao")
                _DataAlteracao = row("UsuarioAlteracaoData")
                _UsuarioCancelamento = row("UsuarioCancelamento")
                _DataCancelamento = row("UsuarioCancelamentoData")

            Next
        Catch ex As Exception
        Finally
            objBanco = Nothing
        End Try
    End Sub

#End Region

#Region "Fields"
    Private _IUD As String = ""

    Private _CodigoEmpresa As String = ""
    Private _EnderecoEmpresa As Integer
    Private _Empresa As Negocio.Cliente

    Private _Codigo As Integer
    Private _Movimento As DateTime
    Private _Validade As DateTime

    Private _Observacoes As String = ""

    Private _CodigoSituacao As Integer
    Private _Situacao As Situacao

    Private _Estoque As Boolean = False
    Private _MovimentoEstoque As DateTime

    Private _UsuarioInclusao As String = ""
    Private _DataInclusao As DateTime
    Private _UsuarioAlteracao As String = ""
    Private _DataAlteracao As DateTime
    Private _UsuarioCancelamento As String = ""
    Private _DataCancelamento As DateTime

    Private _SubOperacaoProducao As SubOperacao

    Private _ItensDeConsumo As ListOrdemParaProducaoXConsumo
    Private _ItensDeInsumo As ListOrdemParaProducaoXInsumo
    Private _ItensDeEspecificacao As ListOrdemParaProducaoXEspecificacao
    Private _ItensDeEPI As ListOrdemParaProducaoXEPI
    Private _ItensDeEmabalagem As ListOrdemParaProducaoXEmbalagens
    Private _ItensDeProcedimento As ListOrdemParaProducaoXProcedimento
    Private _ProdutosdaProducao As ListOrdemParaProducaoXProduto
    Private _ItensEstoque As DataSet


#End Region

#Region "Property"
    Public Property IUD As String
        Get
            Return _IUD
        End Get
        Set(value As String)
            _IUD = value
        End Set
    End Property

    Public Property CodigoEmpresa() As String
        Get
            Return _CodigoEmpresa
        End Get
        Set(ByVal value As String)
            _CodigoEmpresa = value
        End Set
    End Property

    Public Property EnderecoEmpresa() As Integer
        Get
            Return _EnderecoEmpresa
        End Get
        Set(ByVal value As Integer)
            _EnderecoEmpresa = value
        End Set
    End Property

    Public Property Empresa() As Negocio.Cliente
        Get
            If _Empresa Is Nothing And Me.CodigoEmpresa.Trim.Length > 0 Then _Empresa = New Negocio.Cliente(Me.CodigoEmpresa, Me.EnderecoEmpresa)
            Return _Empresa
        End Get
        Set(ByVal value As Negocio.Cliente)
            _Empresa = value
        End Set
    End Property

    Public Property Codigo() As Integer
        Get
            Return _Codigo
        End Get
        Set(ByVal value As Integer)
            _Codigo = value
        End Set
    End Property

    Public Property Movimento() As DateTime
        Get
            Return _Movimento
        End Get
        Set(ByVal value As DateTime)
            _Movimento = value
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

    Public Property Observacoes() As String
        Get
            Return _Observacoes
        End Get
        Set(ByVal value As String)
            _Observacoes = value
        End Set
    End Property

    Public Property CodigoSituacao() As Integer
        Get
            Return _CodigoSituacao
        End Get
        Set(ByVal value As Integer)
            _CodigoSituacao = value
            _Situacao = Nothing
        End Set
    End Property

    Public ReadOnly Property Situacao() As Situacao
        Get
            If _Situacao Is Nothing And _CodigoSituacao > 0 Then _Situacao = New Situacao(_CodigoSituacao)
            Return _Situacao
        End Get
    End Property

    Public Property Estoque() As Boolean
        Get
            Return _Estoque
        End Get
        Set(ByVal value As Boolean)
            _Estoque = value
        End Set
    End Property

    Public Property MovimentoEstoque() As DateTime
        Get
            Return _MovimentoEstoque
        End Get
        Set(ByVal value As DateTime)
            _MovimentoEstoque = value
        End Set
    End Property

    Public Property UsuarioInclusao() As String
        Get
            Return _UsuarioInclusao
        End Get
        Set(ByVal value As String)
            _UsuarioInclusao = value
        End Set
    End Property

    Public Property DataInclusao() As DateTime
        Get
            Return _DataInclusao
        End Get
        Set(ByVal value As DateTime)
            _DataInclusao = value
        End Set
    End Property

    Public Property UsuarioAlteracao() As String
        Get
            Return _UsuarioAlteracao
        End Get
        Set(ByVal value As String)
            _UsuarioAlteracao = value
        End Set
    End Property

    Public Property DataAlteracao() As DateTime
        Get
            Return _DataAlteracao
        End Get
        Set(ByVal value As DateTime)
            _DataAlteracao = value
        End Set
    End Property

    Public Property UsuarioCancelamento() As String
        Get
            Return _UsuarioCancelamento
        End Get
        Set(ByVal value As String)
            _UsuarioCancelamento = value
        End Set
    End Property

    Public Property DataCancelamento() As DateTime
        Get
            Return _DataCancelamento
        End Get
        Set(ByVal value As DateTime)
            _DataCancelamento = value
        End Set
    End Property

    Public Property SubOperacaoProducao() As SubOperacao
        Get
            If _SubOperacaoProducao Is Nothing OrElse _SubOperacaoProducao.Codigo = 0 Then _SubOperacaoProducao = New SubOperacao(40, 2)

            Return _SubOperacaoProducao
        End Get
        Set(value As SubOperacao)
            _SubOperacaoProducao = value
        End Set
    End Property

    Public Property ItensDeConsumo() As ListOrdemParaProducaoXConsumo
        Get
            If _ItensDeConsumo Is Nothing Then _ItensDeConsumo = New ListOrdemParaProducaoXConsumo(Me)

            Return _ItensDeConsumo
        End Get
        Set(ByVal value As ListOrdemParaProducaoXConsumo)
            _ItensDeConsumo = value
        End Set
    End Property

    Public Property ItensDeInsumo() As ListOrdemParaProducaoXInsumo
        Get
            If _ItensDeInsumo Is Nothing Then _ItensDeInsumo = New ListOrdemParaProducaoXInsumo(Me)

            Return _ItensDeInsumo
        End Get
        Set(ByVal value As ListOrdemParaProducaoXInsumo)
            _ItensDeInsumo = value
        End Set
    End Property

    Public Property ItensDeEspecificacao() As ListOrdemParaProducaoXEspecificacao
        Get
            If _ItensDeEspecificacao Is Nothing Then _ItensDeEspecificacao = New ListOrdemParaProducaoXEspecificacao(Me)

            Return _ItensDeEspecificacao
        End Get
        Set(ByVal value As ListOrdemParaProducaoXEspecificacao)
            _ItensDeEspecificacao = value
        End Set
    End Property

    Public Property ItensDeEPI() As ListOrdemParaProducaoXEPI
        Get
            If _ItensDeEPI Is Nothing Then _ItensDeEPI = New ListOrdemParaProducaoXEPI(Me)

            Return _ItensDeEPI
        End Get
        Set(ByVal value As ListOrdemParaProducaoXEPI)
            _ItensDeEPI = value
        End Set
    End Property

    Public Property ItensDeEmabalagem() As ListOrdemParaProducaoXEmbalagens
        Get
            If _ItensDeEmabalagem Is Nothing Then _ItensDeEmabalagem = New ListOrdemParaProducaoXEmbalagens(Me)

            Return _ItensDeEmabalagem
        End Get
        Set(ByVal value As ListOrdemParaProducaoXEmbalagens)
            _ItensDeEmabalagem = value
        End Set
    End Property

    Public Property ItensDeProcedimento() As ListOrdemParaProducaoXProcedimento
        Get
            If _ItensDeProcedimento Is Nothing Then _ItensDeProcedimento = New ListOrdemParaProducaoXProcedimento(Me)

            Return _ItensDeProcedimento
        End Get
        Set(ByVal value As ListOrdemParaProducaoXProcedimento)
            _ItensDeProcedimento = value
        End Set
    End Property

    Public Property ProdutosdaProducao As ListOrdemParaProducaoXProduto
        Get
            If _ProdutosdaProducao Is Nothing Then _ProdutosdaProducao = New ListOrdemParaProducaoXProduto(Me)

            Return _ProdutosdaProducao
        End Get
        Set(value As ListOrdemParaProducaoXProduto)
            _ProdutosdaProducao = value
        End Set
    End Property

    Public Property ItensEstoque As DataSet
        Get
            Return ListarEstoque()
        End Get
        Set(value As DataSet)
            _ItensEstoque = value
        End Set
    End Property

#End Region

#Region "Methods"

    Function ListarEstoque() As DataSet
        Dim objBanco As New AcessaBanco()
        Dim estoque As New DataSet

        Try
            Dim strSQL As String = "SELECT p.Produto_Id + ' - ' + prd.Descricao AS Produto, p.Operacao_Id AS Operacao, p.SubOperacao_Id AS SubOperacao, so.Descricao AS Tipo, p.Movimento_Id AS Movimento, " & vbCrLf &
                                    "		case " & vbCrLf &
                                    "			when p.FisicoFiscal_Id = 1 " & vbCrLf &
                                    "				then 'FÍSICO' " & vbCrLf &
                                    "				else 'FISCAL' " & vbCrLf &
                                    "		end as Estoque, " & vbCrLf &
                                    "		case " & vbCrLf &
                                    "			when len(isnull(prdDerivado.Descricao,'')) > 0 " & vbCrLf &
                                    "				then p.ProdutoDerivado_Id + ' - ' + prdDerivado.Descricao " & vbCrLf &
                                    "				else '' " & vbCrLf &
                                    "		end as ProdutoDerivado, p.Entradas, p.Saidas " & vbCrLf &
                                    "FROM   Producao p " & vbCrLf &
                                    "	INNER JOIN SubOperacoes so " & vbCrLf &
                                    "			on so.Operacao_Id      = p.Operacao_Id " & vbCrLf &
                                    "			and so.SubOperacoes_Id = p.SubOperacao_Id " & vbCrLf &
                                    "	INNER JOIN Produtos AS prd" & vbCrLf &
                                    "			on prd.Produto_Id = p.Produto_Id" & vbCrLf &
                                    "	LEFT JOIN Produtos AS prdDerivado" & vbCrLf &
                                    "			on prdDerivado.Produto_Id = p.ProdutoDerivado_Id" & vbCrLf &
                                    "WHERE p.Empresa_Id      = '" & Me.CodigoEmpresa & "'" & vbCrLf &
                                    "  and p.EndEmpresa_Id   = " & Me.EnderecoEmpresa & vbCrLf &
                                    "  and p.OrdemDeProducao = " & Me.Codigo & vbCrLf &
                                    "ORDER BY  p.Operacao_Id, p.SubOperacao_Id"

            estoque = objBanco.ConsultaDataSet(strSQL, "Producao")

        Catch ex As Exception
        Finally
            objBanco = Nothing
        End Try

        Return estoque
    End Function

    Public Function Salvar() As Boolean
        Dim objBanco As New AcessaBanco()
        Dim sqls As New ArrayList

        salvarSql(sqls)
        Return objBanco.GravaBanco(sqls)
    End Function

    Public Sub salvarSql(ByRef sqls As ArrayList)
        Dim strSql As String = ""

        Select Case IUD
            Case "I"
                Dim n = New Numerador()
                n = New Numerador(Me.CodigoEmpresa, Me.EnderecoEmpresa, 75)

                n.Sequencia += 1

                sqls.Add(n.IncrementarNumeradorSql())

                Me.Codigo = n.Sequencia

                strSql = "INSERT INTO OrdemDeProducao(Empresa_Id, EndEmpresa_Id, Ordem_Id, Movimento, Validade, Observacoes, Situacao, Estoque, " & vbCrLf &
                         "                              UsuarioInclusao, UsuarioInclusaoData)" & vbCrLf &
                         "Values('" & Me.CodigoEmpresa & "'," & Me.EnderecoEmpresa & "," & Me.Codigo & ",'" & Me.Movimento.ToString("yyyy-MM-dd") & "','" & Me.Validade.ToString("yyyy-MM-dd") & "'" & vbCrLf &
                         ", '" & Me.Observacoes & "'," & Me.CodigoSituacao & "," & CByte(Me.Estoque) & ",'" & HttpContext.Current.Session("ssNomeUsuario") & "', getdate())"

                sqls.Add(strSql)

                salvarRelacionados(sqls)

            Case "U"

                salvarRelacionados(sqls)

                strSql = "Update OrdemDeProducao SET " & vbCrLf &
                         "  Movimento            = '" & Me.Movimento.ToString("yyyy-MM-dd") & "'" & vbCrLf &
                         "  ,Validade             = '" & Me.Validade.ToString("yyyy-MM-dd") & "'" & vbCrLf &
                         "  ,Observacoes          = '" & Me.Observacoes & "'" & vbCrLf &
                         "  ,Situacao             = " & Me.CodigoSituacao & vbCrLf &
                         "  ,Estoque              = " & CByte(Me.Estoque) & vbCrLf &
                         "	,UsuarioAlteracao     = '" & HttpContext.Current.Session("ssNomeUsuario") & "'" & vbCrLf &
                         "	,UsuarioAlteracaoData = getdate() " & vbCrLf &
                         " Where Empresa_Id    = '" & Me.CodigoEmpresa & "'" & vbCrLf &
                         "   and EndEmpresa_Id = " & Me.EnderecoEmpresa & vbCrLf &
                         "   and Ordem_Id      = " & Me.Codigo

                sqls.Add(strSql)

            Case "D"

                If Me.Estoque Then
                    strSql = "Delete Producao " & vbCrLf &
                               " Where Empresa_Id      = '" & Me.CodigoEmpresa & "'" & vbCrLf &
                               "   and EndEmpresa_Id   = " & Me.EnderecoEmpresa & vbCrLf &
                               "   and OrdemDeProducao = " & Me.Codigo

                    sqls.Add(strSql)
                End If

                strSql = "Update OrdemDeProducao Set " & vbCrLf &
                         "   Situacao                = " & Me.CodigoSituacao & vbCrLf &
                         "  ,Observacoes             = '" & Me.Observacoes & "'" & vbCrLf &
                         "	,UsuarioCancelamento     = '" & HttpContext.Current.Session("ssNomeUsuario") & "'" & vbCrLf &
                         "	,UsuarioCancelamentoData = getdate() " & vbCrLf &
                         " Where Empresa_Id    = '" & Me.CodigoEmpresa & "'" & vbCrLf &
                         "   and EndEmpresa_Id = " & Me.EnderecoEmpresa & vbCrLf &
                         "   and Ordem_Id      = " & Me.Codigo

                sqls.Add(strSql)

            Case "IN_CON_INS"

                If Me.ItensDeConsumo.Count > 0 Then Me.ItensDeConsumo.SalvarSql(sqls)

                If Me.ItensDeInsumo.Count > 0 Then Me.ItensDeInsumo.SalvarSql(sqls)

                strSql = "Update OrdemDeProducao Set " & vbCrLf &
                         "  Movimento            = '" & Me.Movimento.ToString("yyyy-MM-dd") & "'" & vbCrLf &
                         "  ,Validade             = '" & Me.Validade.ToString("yyyy-MM-dd") & "'" & vbCrLf &
                         "  ,Observacoes          = '" & Me.Observacoes & "'" & vbCrLf &
                         "  ,Situacao             = " & Me.CodigoSituacao & vbCrLf &
                         "  ,Estoque              = " & CByte(Me.Estoque) & vbCrLf &
                         "	,UsuarioAlteracao     ='" & HttpContext.Current.Session("ssNomeUsuario") & "'" & vbCrLf &
                         "	,UsuarioAlteracaoData = getdate() " & vbCrLf &
                         " Where Empresa_Id    = '" & Me.CodigoEmpresa & "'" & vbCrLf &
                         "   and EndEmpresa_Id = " & Me.EnderecoEmpresa & vbCrLf &
                         "   and Ordem_Id      = " & Me.Codigo

                sqls.Add(strSql)

        End Select

    End Sub

    Private Sub salvarRelacionados(ByRef sqls As ArrayList)

        removeProducao(sqls)

        If Me.ProdutosdaProducao.Count > 0 Then Me.ProdutosdaProducao.SalvarSql(sqls)

        If Me.ItensDeConsumo.Count > 0 Then Me.ItensDeConsumo.SalvarSql(sqls)

        If Me.ItensDeInsumo.Count > 0 Then Me.ItensDeInsumo.SalvarSql(sqls)

        If Me.ItensDeEspecificacao.Count > 0 Then Me.ItensDeEspecificacao.SalvarSql(sqls)

        If Me.ItensDeEPI.Count > 0 Then Me.ItensDeEPI.SalvarSql(sqls)

        If Me.ItensDeEmabalagem.Count > 0 Then Me.ItensDeEmabalagem.SalvarSql(sqls)

        If Me.ItensDeProcedimento.Count > 0 Then Me.ItensDeProcedimento.SalvarSql(sqls)

    End Sub

    Private Sub removeProducao(ByRef sqls As ArrayList)
        Dim strSql As String = ""

        strSql = "Delete Producao " & vbCrLf &
                  " Where Empresa_Id      = '" & Me.CodigoEmpresa & "'" & vbCrLf &
                  "   and EndEmpresa_Id   = " & Me.EnderecoEmpresa & vbCrLf &
                  "   and OrdemDeProducao = " & Me.Codigo

        sqls.Add(strSql)
    End Sub

    Public Function Clone() As OrdemParaProducao
        Dim ordemParaProducao As OrdemParaProducao = CType(Me.MemberwiseClone(), OrdemParaProducao)
        Return ordemParaProducao
    End Function

#End Region

End Class
