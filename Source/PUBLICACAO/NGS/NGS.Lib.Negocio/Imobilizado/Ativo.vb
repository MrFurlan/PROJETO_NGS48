Public Class Ativo

#Region "Fields"

    Private _CodigoEmpresa As String
    Private _CodigoGrupo As String
    Private _Codigo As Integer
    Private _Sequencia As Integer
    Private _Situacao As Integer
    Private _UnidadeDeNegocio As String
    Private _CodigoEmpresaResp As String
    Private _EnderecoEmpresaResp As String
    Private _EmpresaResp As Cliente
    Private _CodigoCentroDeCusto As Integer
    Private _Conta As String
    Private _Descricao As String
    Private _Historico As String
    Private _Identificacao As String
    Private _DataAquisicao As DateTime
    Private _DataInicioDeUso As DateTime
    Private _ValorOriginal As Decimal
    Private _Depreciar As String
    Private _DataAtualizacao As DateTime
    Private _QuemLancou As String
    Private _QuandoLancou As DateTime
    Private _QuemAlterou As String
    Private _QuandoAlterou As DateTime
    Private _Seguro As String
    Private _DataBaixa As DateTime

#End Region

#Region "Properties"

    Public Property CodigoEmpresa() As String
        Get
            Return _CodigoEmpresa
        End Get
        Set(ByVal value As String)
            _CodigoEmpresa = value
        End Set
    End Property

    Public Property CodigoGrupo() As String
        Get
            Return _CodigoGrupo
        End Get
        Set(ByVal value As String)
            _CodigoGrupo = value
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

    Public Property Sequencia() As Integer
        Get
            Return _Sequencia
        End Get
        Set(ByVal value As Integer)
            _Sequencia = value
        End Set
    End Property

    Public Property Situacao() As Integer
        Get
            Return _Situacao
        End Get
        Set(ByVal value As Integer)
            _Situacao = value
        End Set
    End Property

    Public Property UnidadeDeNegocio() As String
        Get
            Return _UnidadeDeNegocio
        End Get
        Set(ByVal value As String)
            _UnidadeDeNegocio = value
        End Set
    End Property

    Public Property CodigoEmpresaResp() As String
        Get
            Return _CodigoEmpresaResp
        End Get
        Set(ByVal value As String)
            _CodigoEmpresaResp = value
        End Set
    End Property

    Public Property EnderecoEmpresaResp() As String
        Get
            Return _EnderecoEmpresaResp
        End Get
        Set(ByVal value As String)
            _EnderecoEmpresaResp = value
        End Set
    End Property

    Public Property EmpresaResp() As Cliente
        Get
            If _EmpresaResp Is Nothing AndAlso Not String.IsNullOrWhiteSpace(_CodigoEmpresaResp) Then _EmpresaResp = New Cliente(_CodigoEmpresaResp, _EnderecoEmpresaResp)
            Return _EmpresaResp
        End Get
        Set(ByVal value As Cliente)
            _EmpresaResp = value
        End Set
    End Property

    Public Property CodigoCentroDeCusto() As Integer
        Get
            Return _CodigoCentroDeCusto
        End Get
        Set(ByVal value As Integer)
            _CodigoCentroDeCusto = value
        End Set
    End Property

    Public Property Conta() As String
        Get
            Return _Conta
        End Get
        Set(ByVal value As String)
            _Conta = value
        End Set
    End Property

    Public Property Descricao() As String
        Get
            Return _Descricao
        End Get
        Set(ByVal value As String)
            _Descricao = value
        End Set
    End Property

    Public Property Historico() As String
        Get
            Return _Historico
        End Get
        Set(ByVal value As String)
            _Historico = value
        End Set
    End Property

    Public Property Identificacao() As String
        Get
            Return _Identificacao
        End Get
        Set(ByVal value As String)
            _Identificacao = value
        End Set
    End Property

    Public Property DataAquisicao() As DateTime
        Get
            Return _DataAquisicao
        End Get
        Set(ByVal value As DateTime)
            _DataAquisicao = value
        End Set
    End Property

    Public Property DataInicioDeUso() As DateTime
        Get
            Return _DataInicioDeUso
        End Get
        Set(ByVal value As DateTime)
            _DataInicioDeUso = value
        End Set
    End Property

    Public Property ValorOriginal() As Decimal
        Get
            Return _ValorOriginal
        End Get
        Set(ByVal value As Decimal)
            _ValorOriginal = value
        End Set
    End Property

    Public Property Depreciar() As String
        Get
            Return _Depreciar
        End Get
        Set(ByVal value As String)
            _Depreciar = value
        End Set
    End Property

    Public Property DataAtualizacao() As DateTime
        Get
            Return _DataAtualizacao
        End Get
        Set(ByVal value As DateTime)
            _DataAtualizacao = value
        End Set
    End Property

    Public Property QuemLancou() As String
        Get
            Return _QuemLancou
        End Get
        Set(ByVal value As String)
            _QuemLancou = value
        End Set
    End Property

    Public Property QuandoLancou() As DateTime
        Get
            Return _QuandoLancou
        End Get
        Set(ByVal value As DateTime)
            _QuandoLancou = value
        End Set
    End Property

    Public Property QuemAlterou() As String
        Get
            Return _QuemAlterou
        End Get
        Set(ByVal value As String)
            _QuemAlterou = value
        End Set
    End Property

    Public Property QuandoAlterou() As DateTime
        Get
            Return _QuandoAlterou
        End Get
        Set(ByVal value As DateTime)
            _QuandoAlterou = value
        End Set
    End Property

    Public Property Seguro() As String
        Get
            Return _Seguro
        End Get
        Set(ByVal value As String)
            _Seguro = value
        End Set
    End Property

    Public Property DataBaixa() As DateTime
        Get
            Return _DataBaixa
        End Get
        Set(ByVal value As DateTime)
            _DataBaixa = value
        End Set
    End Property

#End Region

End Class
