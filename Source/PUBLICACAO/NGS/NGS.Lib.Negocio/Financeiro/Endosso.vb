Imports System.Web

<Serializable()> _
Public Class ListEndosso
    Inherits List(Of Endosso)

#Region "Construtor"
    Public Sub New()

    End Sub

    Public Sub New(ByVal Empresa As String, ByVal EndEmpresa As String, Optional ByVal Where As String = "")
        Dim objBanco As New AcessaBanco()

        Try
            Dim strSQL As String = "SELECT Empresa_Id, EndEmpresa_Id, Codigo_Id, Movimento, ClienteEndosso_Id, EndClienteEndosso_Id, NumeroEndosso, Vencimento, Valor, Observacoes, Situacao," & vbCrLf &
                                   "       isnull(UsuarioInclusao,'') as UsuarioInclusao, isnull(UsuarioInclusaoData,GetDate()) as UsuarioInclusaoData," & vbCrLf &
                                   "       isnull(UsuarioAlteracao,'') as UsuarioAlteracao, isnull(UsuarioAlteracaoData,GetDate()) as UsuarioAlteracaoData," & vbCrLf &
                                   "       isnull(UsuarioCancelamento,'') as UsuarioCancelamento, isnull(UsuarioCancelamentoData,GetDate()) as UsuarioCancelamentoData, ObservacoesInterna " & vbCrLf &
                       "  FROM Endosso WHERE 1 = 1 " & vbCrLf

            If Empresa.Length > 0 Then
                strSQL &= "  AND Empresa_Id    = '" & Empresa & "' " & vbCrLf & _
                          "  AND EndEmpresa_Id = " & EndEmpresa & " " & vbCrLf
            End If

            If Where.Length > 0 Then
                strSQL &= "  AND " & Where
            End If

            Dim dsE As DataSet = objBanco.ConsultaDataSet(strSQL, "Endosso")
            For Each row As DataRow In dsE.Tables(0).Rows
                Dim e As New Endosso

                e.CodigoEmpresa = row("Empresa_Id")
                e.EnderecoEmpresa = row("EndEmpresa_Id")
                e.Codigo = row("Codigo_Id")
                'e.CodigoCliente = row("Cliente_Id")
                'e.EnderecoCliente = row("EndCliente_Id")
                e.Movimento = CDate(row("Movimento"))
                e.CodigoClienteEndosso = row("ClienteEndosso_Id")
                e.EnderecoClienteEndosso = row("EndClienteEndosso_Id")
                e.NumeroEndosso = row("NumeroEndosso")
                e.Vencimento = CDate(row("Vencimento"))
                e.Valor = row("Valor")
                e.Observacoes = row("Observacoes")
                e.CodigoSituacao = row("Situacao")
                e.UsuarioInclusao = row("UsuarioInclusao")
                e.DataInclusao = row("UsuarioInclusaoData")
                e.UsuarioAlteracao = row("UsuarioAlteracao")
                e.DataAlteracao = row("UsuarioAlteracaoData")
                e.UsuarioCancelamento = row("UsuarioCancelamento")
                e.DataCancelamento = row("UsuarioCancelamentoData")
                e.ObservacoesInterna = row("ObservacoesInterna")

                Me.Add(e)
            Next
        Catch ex As Exception
        Finally
            objBanco = Nothing
        End Try
    End Sub
#End Region

End Class

<Serializable()> _
Public Class Endosso
    Implements IBaseEntity

#Region "Construtor"
    Public Sub New()

    End Sub

    Public Sub New(Empresa As String, ByVal EndEmpresa As String, ByVal CodigoEndosso As Integer)
        Dim objBanco As New AcessaBanco()

        Try
            Dim strSQL As String = "SELECT Empresa_Id, EndEmpresa_Id, Codigo_Id, Movimento, ClienteEndosso_Id, EndClienteEndosso_Id, NumeroEndosso, Vencimento, Valor, Observacoes, Situacao," & vbCrLf &
                                   "       isnull(UsuarioInclusao,'') as UsuarioInclusao, isnull(UsuarioInclusaoData,GetDate()) as UsuarioInclusaoData," & vbCrLf &
                                   "       isnull(UsuarioAlteracao,'') as UsuarioAlteracao, isnull(UsuarioAlteracaoData,GetDate()) as UsuarioAlteracaoData," & vbCrLf &
                                   "       isnull(UsuarioCancelamento,'') as UsuarioCancelamento, isnull(UsuarioCancelamentoData,GetDate()) as UsuarioCancelamentoData, ObservacoesInterna " & vbCrLf &
                                   "FROM Endosso " & vbCrLf &
                                   "WHERE Empresa_Id    = '" & Empresa & "' " & vbCrLf &
                                   "  AND EndEmpresa_Id = " & EndEmpresa & vbCrLf &
                                   "  AND Codigo_Id     = " & CodigoEndosso

            Dim dsE As DataSet = objBanco.ConsultaDataSet(strSQL, "Endosso")
            For Each row As DataRow In dsE.Tables(0).Rows
                _CodigoEmpresa = row("Empresa_Id")
                _EnderecoEmpresa = row("EndEmpresa_Id")
                _Codigo = row("Codigo_Id")
                '_CodigoCliente = row("Cliente_Id")
                '_EnderecoCliente = row("EndCliente_Id")
                _Movimento = CDate(row("Movimento"))
                _CodigoClienteEndosso = row("ClienteEndosso_Id")
                _EnderecoClienteEndosso = row("EndClienteEndosso_Id")
                _NumeroEndosso = row("NumeroEndosso")
                _Vencimento = CDate(row("Vencimento"))
                _Valor = row("Valor")
                _Observacoes = row("Observacoes")
                _CodigoSituacao = row("Situacao")
                _UsuarioInclusao = row("UsuarioInclusao")
                _DataInclusao = row("UsuarioInclusaoData")
                _UsuarioAlteracao = row("UsuarioAlteracao")
                _DataAlteracao = row("UsuarioAlteracaoData")
                _UsuarioCancelamento = row("UsuarioCancelamento")
                _DataCancelamento = row("UsuarioCancelamentoData")
                _ObservacoesInterna = row("ObservacoesInterna")
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

    'Private _CodigoCliente As String = ""
    'Private _EnderecoCliente As Integer
    'Private _Cliente As Negocio.Cliente
    'Private _ClienteDescricao As String = ""

    Private _Movimento As DateTime

    Private _CodigoClienteEndosso As String = ""
    Private _EnderecoClienteEndosso As Integer
    Private _ClienteEndosso As Negocio.Cliente
    Private _ClienteEndossoDescricao As String = ""

    Private _NumeroEndosso As String = ""

    Private _Vencimento As DateTime

    Private _Valor As Decimal

    Private _Observacoes As String = ""

    Private _CodigoSituacao As Integer
    Private _Situacao As Situacao

    Private _UsuarioInclusao As String = ""
    Private _DataInclusao As DateTime
    Private _UsuarioAlteracao As String = ""
    Private _DataAlteracao As DateTime
    Private _UsuarioCancelamento As String = ""
    Private _DataCancelamento As DateTime

    Private _ObservacoesInterna As String = ""

    Private _TitulosXEndosso As ListEndossoXTitulo

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

    'Public Property CodigoCliente() As String
    '    Get
    '        Return _CodigoCliente
    '    End Get
    '    Set(ByVal value As String)
    '        _CodigoCliente = value
    '        _Cliente = Nothing
    '    End Set
    'End Property

    'Public Property EnderecoCliente() As Integer
    '    Get
    '        Return _EnderecoCliente
    '    End Get
    '    Set(ByVal value As Integer)
    '        _EnderecoCliente = value
    '        _Cliente = Nothing
    '    End Set
    'End Property

    'Public Property Cliente() As Negocio.Cliente
    '    Get
    '        If _Cliente Is Nothing And Me.CodigoCliente.Trim.Length > 0 Then _Cliente = New Negocio.Cliente(Me.CodigoCliente, Me.EnderecoCliente)
    '        Return _Cliente
    '    End Get
    '    Set(ByVal value As Negocio.Cliente)
    '        _Cliente = value
    '    End Set
    'End Property

    'Public ReadOnly Property ClienteDescricao() As String
    '    Get
    '        If Me.CodigoCliente.Trim.Length > 0 AndAlso _Cliente IsNot Nothing Then _ClienteDescricao = _Cliente.CodigoFormatado & "-" & _Cliente.Nome
    '        Return _ClienteDescricao
    '    End Get
    'End Property


    Public Property Movimento() As DateTime
        Get
            Return _Movimento
        End Get
        Set(ByVal value As DateTime)
            _Movimento = value
        End Set
    End Property

    Public Property CodigoClienteEndosso() As String
        Get
            Return _CodigoClienteEndosso
        End Get
        Set(ByVal value As String)
            _CodigoClienteEndosso = value
            _ClienteEndosso = Nothing
        End Set
    End Property

    Public Property EnderecoClienteEndosso() As Integer
        Get
            Return _EnderecoClienteEndosso
        End Get
        Set(ByVal value As Integer)
            _EnderecoClienteEndosso = value
            _ClienteEndosso = Nothing
        End Set
    End Property

    Public Property ClienteEndosso() As Negocio.Cliente
        Get
            If _ClienteEndosso Is Nothing And Me.CodigoClienteEndosso.Trim.Length > 0 Then _ClienteEndosso = New Negocio.Cliente(Me.CodigoClienteEndosso, Me.EnderecoClienteEndosso)
            Return _ClienteEndosso
        End Get
        Set(ByVal value As Negocio.Cliente)
            _ClienteEndosso = value
        End Set
    End Property

    Public ReadOnly Property ClienteEndossoDescricao() As String
        Get
            If Me.CodigoClienteEndosso.Trim.Length > 0 AndAlso _ClienteEndosso IsNot Nothing Then _ClienteEndossoDescricao = _ClienteEndosso.CodigoFormatado & "-" & _ClienteEndosso.Nome
            Return _ClienteEndossoDescricao
        End Get
    End Property

    Public Property NumeroEndosso As String
        Get
            Return _NumeroEndosso
        End Get
        Set(value As String)
            _NumeroEndosso = value
        End Set
    End Property

    Public Property Vencimento() As DateTime
        Get
            Return _Vencimento
        End Get
        Set(ByVal value As DateTime)
            _Vencimento = value
        End Set
    End Property

    Public Property Valor() As Decimal
        Get
            Return _Valor
        End Get
        Set(ByVal value As Decimal)
            _Valor = value
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

    Public Property ObservacoesInterna() As String
        Get
            Return _ObservacoesInterna
        End Get
        Set(ByVal value As String)
            _ObservacoesInterna = value
        End Set
    End Property

    Public Property TitulosXEndosso() As ListEndossoXTitulo
        Get
            If _TitulosXEndosso Is Nothing Then _TitulosXEndosso = New ListEndossoXTitulo(Me)

            Return _TitulosXEndosso
        End Get
        Set(ByVal value As ListEndossoXTitulo)
            _TitulosXEndosso = value
        End Set
    End Property

#End Region

#Region "Methods"

    Private Sub CarregarEndossoXTitulo(ByVal oPro As Endosso)
        Dim objBanco As New AcessaBanco()

        'Try
        '    'Dim strSQL As String = "SELECT Empresa_Id, EndEmpresa_Id, Ordem_Id, Lote, Movimento, Validade, Produto, Quantidade, QuantidadeDeAjuste, Observacoes, Situacao, Estoque, " & vbCrLf &
        '    '                       "       isnull(UsuarioInclusao,'') as UsuarioInclusao, isnull(UsuarioInclusaoData,GetDate()) as UsuarioInclusaoData," & vbCrLf &
        '    '                       "       isnull(UsuarioAlteracao,'') as UsuarioAlteracao, isnull(UsuarioAlteracaoData,GetDate()) as UsuarioAlteracaoData," & vbCrLf &
        '    '                       "       isnull(UsuarioCancelamento,'') as UsuarioCancelamento, isnull(UsuarioCancelamentoData,GetDate()) as UsuarioCancelamentoData " & vbCrLf &
        '    '                       "FROM OrdemDeProducao " & vbCrLf &
        '    '                       "WHERE Empresa_Id    = '" & oPro.CodigoEmpresa & "' " & vbCrLf &
        '    '                       "  AND EndEmpresa_Id = " & oPro.EnderecoEmpresa & " " & vbCrLf &
        '    '                       "  AND Produto       = " & oPro.CodigoProduto

        '    'Dim dsOP As DataSet = objBanco.ConsultaDataSet(strSQL, "OrdemDeProducao")
        '    'For Each row As DataRow In dsOP.Tables(0).Rows
        '    '    Dim op As New Endosso

        '    '    op.CodigoEmpresa = row("Empresa_Id")
        '    '    op.EnderecoEmpresa = row("EndEmpresa_Id")
        '    '    op.Codigo = row("Ordem_Id")
        '    '    op.Lote = row("Lote")
        '    '    op.Movimento = CDate(row("Movimento"))
        '    '    op.Validade = CDate(row("Validade"))
        '    '    op.CodigoProduto = row("Produto")
        '    '    op.Quantidade = row("Quantidade")
        '    '    op.QuantidadeAjuste = row("QuantidadeDeAjuste")
        '    '    op.Observacoes = row("Observacoes")
        '    '    op.CodigoSituacao = row("Situacao")
        '    '    op.Estoque = row("Estoque")
        '    '    op.UsuarioInclusao = row("UsuarioInclusao")
        '    '    op.DataInclusao = row("UsuarioInclusaoData")
        '    '    op.UsuarioAlteracao = row("UsuarioAlteracao")
        '    '    op.DataAlteracao = row("UsuarioAlteracaoData")
        '    '    op.UsuarioCancelamento = row("UsuarioCancelamento")
        '    '    op.DataCancelamento = row("UsuarioCancelamentoData")
        '    Next
        'Catch ex As Exception
        'Finally
        '    objBanco = Nothing
        'End Try

    End Sub

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
                n = New Numerador(Me.CodigoEmpresa, Me.EnderecoEmpresa, 200)

                n.Sequencia += 1

                sqls.Add(n.IncrementarNumeradorSql())

                Me.Codigo = n.Sequencia

                strSql = "INSERT INTO Endosso(Empresa_Id, EndEmpresa_Id, Codigo_Id, Movimento, ClienteEndosso_Id, EndClienteEndosso_Id, NumeroEndosso, Vencimento, Valor, Observacoes, Situacao, ObservacoesInterna, " & vbCrLf & _
                         "                    UsuarioInclusao, UsuarioInclusaoData)" & vbCrLf & _
                         "Values('" & Me.CodigoEmpresa & "'," & Me.EnderecoEmpresa & "," & Me.Codigo & ",'" & Me.Movimento.ToString("yyyy-MM-dd") & "'" & vbCrLf & _
                          ",'" & Me.CodigoClienteEndosso & "'," & Me.EnderecoClienteEndosso & ",'" & Me.NumeroEndosso & "','" & Me.Vencimento.ToString("yyyy-MM-dd") & "'" & vbCrLf & _
                         "," & Str(Me.Valor) & ",'" & Me.Observacoes & "'," & Me.CodigoSituacao & ",'" & Me.ObservacoesInterna & "','" & HttpContext.Current.Session("ssNomeUsuario") & "', getdate())"

                sqls.Add(strSql)

                salvarRelacionados(sqls)

            Case "U"

                strSql = "Update Endosso Set " & vbCrLf & _
                         "  ,Movimento            = '" & Me.Movimento.ToString("yyyy-MM-dd") & "'" & vbCrLf & _
                         "  ,ClienteEndosso_Id    = '" & Me.CodigoClienteEndosso & "'" & vbCrLf & _
                         "  ,EndClienteEndosso_Id = " & Me.EnderecoClienteEndosso & vbCrLf & _
                         "  ,NumeroEndosso        = '" & Me.NumeroEndosso & "'" & vbCrLf & _
                         "  ,Vencimento           = '" & Me.Vencimento.ToString("yyyy-MM-dd") & "'" & vbCrLf & _
                         "  ,Valor                = " & Str(Me.Valor) & vbCrLf & _
                         "  ,Observacoes          = '" & Me.Observacoes & "'" & vbCrLf & _
                         "  ,Situacao             = " & Me.CodigoSituacao & vbCrLf & _
                         "	,UsuarioAlteracao     ='" & HttpContext.Current.Session("ssNomeUsuario") & "'" & vbCrLf & _
                         "	,UsuarioAlteracaoData = getdate() " & vbCrLf & _
                         "	,ObservacoesInterna   ='" & Me.ObservacoesInterna & "'" & vbCrLf & _
                         " Where Empresa_Id    = '" & Me.CodigoEmpresa & "'" & vbCrLf & _
                         "   and EndEmpresa_Id = " & Me.EnderecoEmpresa & vbCrLf & _
                         "   and Codigo_Id     = " & Me.Codigo

                sqls.Add(strSql)

                salvarAlteracao(sqls)

                salvarRelacionados(sqls)

            Case "D"

                strSql = "Update Endosso Set " & vbCrLf & _
                         "   Situacao                = " & Me.CodigoSituacao & vbCrLf & _
                         "  ,Observacoes             = '" & Me.Observacoes & "'" & vbCrLf & _
                         "	,UsuarioCancelamento     = '" & HttpContext.Current.Session("ssNomeUsuario") & "'" & vbCrLf & _
                         "	,UsuarioCancelamentoData = getdate() " & vbCrLf & _
                         "	,ObservacoesInterna      ='" & Me.ObservacoesInterna & "'" & vbCrLf & _
                         " Where Empresa_Id    = '" & Me.CodigoEmpresa & "'" & vbCrLf & _
                         "   and EndEmpresa_Id = " & Me.EnderecoEmpresa & vbCrLf & _
                         "   and Codigo_Id     = " & Me.Codigo

                sqls.Add(strSql)

                salvarRelacionados(sqls)

            Case "F"

                strSql = "Update Endosso Set " & vbCrLf & _
                         "   Situacao                = " & Me.CodigoSituacao & vbCrLf & _
                         "  ,Observacoes             = '" & Me.Observacoes & "'" & vbCrLf & _
                         "	,UsuarioCancelamento     = '" & HttpContext.Current.Session("ssNomeUsuario") & "'" & vbCrLf & _
                         "	,UsuarioCancelamentoData = getdate() " & vbCrLf & _
                         "	,ObservacoesInterna      ='" & Me.ObservacoesInterna & "'" & vbCrLf & _
                         " Where Empresa_Id    = '" & Me.CodigoEmpresa & "'" & vbCrLf & _
                         "   and EndEmpresa_Id = " & Me.EnderecoEmpresa & vbCrLf & _
                         "   and Codigo_Id     = " & Me.Codigo

                sqls.Add(strSql)

                salvarRelacionados(sqls)

        End Select

    End Sub


    Private Sub salvarAlteracao(ByRef sqls As ArrayList)
        Dim sql As String = ""

        Dim nrEndosso As New Endosso()
        nrEndosso.CodigoEmpresa = Me.CodigoEmpresa
        nrEndosso.EnderecoEmpresa = Me.EnderecoEmpresa
        nrEndosso.Codigo = Me.Codigo

        Dim listaDosTitulos = New ListEndossoXTitulo(nrEndosso)

        For Each endXT In listaDosTitulos
            endXT.Titulo.IUD = "U"
            endXT.Titulo.CodigoSituacao = eSituacao.Normal

            Dim obs As String = endXT.Titulo.ObservacoesControleInterno
            If Not String.IsNullOrWhiteSpace(obs) AndAlso obs.Length > 0 Then
                obs = obs & ". Endosso alterou a Situação do Título para " & endXT.Titulo.CodigoSituacao.ToString() & "-" & endXT.Titulo.Situacao.Descricao & " em " & Format(Now, "dd/MM/yyyy HH:mm:ss") & " feita por " & HttpContext.Current.Session("ssNomeUsuario")
            Else
                obs = "Endosso alterou a Situação do Título para " & endXT.Titulo.CodigoSituacao.ToString() & "-" & endXT.Titulo.Situacao.Descricao & " em " & Format(Now, "dd/MM/yyyy HH:mm:ss") & " feita por " & HttpContext.Current.Session("ssNomeUsuario")
            End If

            endXT.Titulo.ObservacoesControleInterno = obs

            endXT.Titulo.SalvarSql(sqls)
        Next

        sql = "Delete EndossoXTitulo " & vbCrLf & _
                       " Where Empresa_Id    = '" & Me.CodigoEmpresa & "'" & vbCrLf & _
                       "   and EndEmpresa_Id = " & Me.EnderecoEmpresa & vbCrLf & _
                       "   and Codigo_Id     = " & Me.Codigo

        sqls.Add(sql)
    End Sub


    Private Sub salvarRelacionados(ByRef sqls As ArrayList)

        If Me.TitulosXEndosso.Count > 0 Then Me.TitulosXEndosso.SalvarSql(sqls)

    End Sub

#End Region

End Class
