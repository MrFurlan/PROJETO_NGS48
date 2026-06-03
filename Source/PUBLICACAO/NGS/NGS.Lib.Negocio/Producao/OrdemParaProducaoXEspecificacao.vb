Public Class ListOrdemParaProducaoXEspecificacao
    Inherits List(Of OrdemParaProducaoXEspecificacao)
    Implements IBaseEntity

#Region "Contrutor"
    Public Sub New()
    End Sub

    Public Sub New(ByVal objOrdemDeProducao As OrdemParaProducao)

        Me.OrdemParaProducao = objOrdemDeProducao

        Dim db As New AcessaBanco()

        Try
            Dim strSQL As String

            strSQL = "Select Empresa_Id, EndEmpresa_Id, Ordem_Id, ProdutoProducao_Id, CodigoEspecificacao_Id, FaixaInicial, FaixaFinal, Resultado " & vbCrLf &
                     "From OrdemDeProducaoXEspecificacao " & vbCrLf &
                     "Where Empresa_id    = '" & Me.OrdemParaProducao.CodigoEmpresa & "'" & vbCrLf &
                     "  and EndEmpresa_Id = " & Me.OrdemParaProducao.EnderecoEmpresa & vbCrLf

            If Me.OrdemParaProducao.Codigo > 0 Then strSQL &= "  and Ordem_Id      = " & Me.OrdemParaProducao.Codigo & vbCrLf

            Dim ds As DataSet = db.ConsultaDataSet(strSQL, "OrdemParaProducaoXEspecificacao")

            For Each row As DataRow In ds.Tables(0).Rows
                Dim oPE As New OrdemParaProducaoXEspecificacao(Me.OrdemParaProducao)

                oPE.CodigoProdutoProducao = row("ProdutoProducao_Id")
                oPE.CodigoEspecificacao = row("CodigoEspecificacao_Id")
                oPE.EspecificacaoDoProduto = New EspecificacaoDoProduto(row("CodigoEspecificacao_Id"))

                oPE.FaixaInicial = row("FaixaInicial")
                oPE.FaixaFinal = row("FaixaFinal")
                oPE.Resultado = row("Resultado")

                Me.Add(oPE)
            Next

        Catch ex As Exception
            Throw ex
        Finally
            db = Nothing
        End Try
    End Sub

    Public Sub SalvarSql(ByRef Sqls As ArrayList)
        For Each item As OrdemParaProducaoXEspecificacao In Me
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

<Serializable()> _
Public Class OrdemParaProducaoXEspecificacao
    Implements IBaseEntity

#Region "Fields"
    Private _IUD As String = ""

    Private _OrdemParaProducao As OrdemParaProducao

    Private _CodigoEspecificacao As Integer
    Private _EspecificacaoDoProduto As EspecificacaoDoProduto

    Private _CodigoProdutoProducao As String
    Private _ProdutoProducao As Produto

    Private _FaixaInicial As Decimal
    Private _FaixaFinal As Decimal

    Private _Resultado As Decimal

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

    Public Property CodigoEspecificacao As Integer
        Get
            Return _CodigoEspecificacao
        End Get
        Set(value As Integer)
            _CodigoEspecificacao = value
        End Set
    End Property

    Public Property EspecificacaoDoProduto As EspecificacaoDoProduto
        Get
            If _EspecificacaoDoProduto Is Nothing And _CodigoEspecificacao > 0 Then _EspecificacaoDoProduto = New EspecificacaoDoProduto(_CodigoEspecificacao)
            Return _EspecificacaoDoProduto
        End Get
        Set(value As EspecificacaoDoProduto)
            _EspecificacaoDoProduto = value
        End Set
    End Property

    Public Property FaixaInicial As Decimal
        Get
            Return _FaixaInicial
        End Get
        Set(value As Decimal)
            _FaixaInicial = value
        End Set
    End Property

    Public Property FaixaFinal As Decimal
        Get
            Return _FaixaFinal
        End Get
        Set(value As Decimal)
            _FaixaFinal = value
        End Set
    End Property

    Public Property Resultado() As Decimal
        Get
            Return _Resultado
        End Get
        Set(ByVal value As Decimal)
            _Resultado = value
        End Set
    End Property

    Public Property CodigoProdutoProducao As String
        Get
            Return _CodigoProdutoProducao
        End Get
        Set(value As String)
            _CodigoProdutoProducao = value
        End Set
    End Property

    Public Property ProdutoProducao() As Produto
        Get
            If _ProdutoProducao Is Nothing And Not Me.CodigoProdutoProducao Is Nothing Then _ProdutoProducao = New Produto(Me.CodigoProdutoProducao)
            Return _ProdutoProducao
        End Get
        Set(ByVal value As Produto)
            _ProdutoProducao = value
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

                strSql = "INSERT INTO OrdemDeProducaoXEspecificacao(Empresa_Id, EndEmpresa_Id, Ordem_Id, ProdutoProducao_Id, CodigoEspecificacao_Id, FaixaInicial, FaixaFinal, Resultado)" & vbCrLf &
                         "Values('" & OrdemParaProducao.CodigoEmpresa & "'," & OrdemParaProducao.EnderecoEmpresa & "," & OrdemParaProducao.Codigo & ",'" & Me.CodigoProdutoProducao & "'," & Me.CodigoEspecificacao & "," & Str(Me.FaixaInicial) & "," & Str(Me.FaixaFinal) & "," & Str(Me.Resultado) & ")"

            Case "U"
                strSql = "UPDATE OrdemDeProducaoXEspecificacao Set " & vbCrLf &
                         "   FaixaInicial                           = " & Str(Me.FaixaInicial) & vbCrLf &
                         "  ,FaixaFinal                             = " & Str(Me.FaixaFinal) & vbCrLf &
                         "  ,Resultado                              = " & Str(Me.Resultado) & vbCrLf &
                         " WHERE Empresa_Id                         = '" & OrdemParaProducao.CodigoEmpresa & "'" & vbCrLf &
                         "      AND EndEmpresa_Id                      = " & OrdemParaProducao.EnderecoEmpresa & vbCrLf &
                         "      AND Ordem_Id                           = " & OrdemParaProducao.Codigo & vbCrLf &
                         "      AND ProdutoProducao_Id                 = '" & Me.CodigoProdutoProducao & "'" & vbCrLf &
                         "      AND CodigoEspecificacao_Id               = " & Me.CodigoEspecificacao

            Case "D"
                strSql = "DELETE OrdemDeProducaoXEspecificacao " & vbCrLf &
                         " WHERE Empresa_Id                         = '" & OrdemParaProducao.CodigoEmpresa & "'" & vbCrLf &
                         "   AND EndEmpresa_Id                      = " & OrdemParaProducao.EnderecoEmpresa & vbCrLf &
                         "   AND Ordem_Id                           = " & OrdemParaProducao.Codigo & vbCrLf &
                         "   AND ProdutoProducao_Id                 = '" & Me.CodigoProdutoProducao & "'" & vbCrLf &
                         "   AND CodigoEspecificacao_Id             = " & Me.CodigoEspecificacao
        End Select

        Sqls.Add(strSql)
    End Sub
#End Region

End Class
