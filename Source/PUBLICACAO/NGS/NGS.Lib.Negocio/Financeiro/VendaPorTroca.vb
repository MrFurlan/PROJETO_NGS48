Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Linq
Imports System.Collections.Generic
Imports NGS.Lib.Uteis

<Serializable()> _
Public Class ListVendaPorTroca
    Inherits List(Of VendaPorTroca)

End Class

<Serializable()> _
Public Class VendaPorTroca
    Implements IBaseEntity

#Region "Contrutor"
    Public Sub New()

    End Sub

    Public Sub New(ByVal pRegistro As Integer, Optional ByVal NF As Negocio.NotaFiscal = Nothing)
        _NotaFiscal = NF
        Dim sql As String
        sql = "SELECT Registro_Id, Sequencia_Id, Indexador, Moeda, Situacao, Lote, Movimento, Vencimento, UnidadeDebito, EmpresaDebito, EndEmpresaDebito, ClienteDebito, " & vbCrLf & _
              "       EndClienteDebito, PedidoDebito, CarteiraDebito, UnidadeCredito, EmpresaCredito, EndEmpresaCredito, ClienteCredito, EndClienteCredito, PedidoCredito, " & vbCrLf & _
              "       CarteiraCredito, ValorOficial, ValorMoeda, TaxaDeJuros, Historico, UsuarioInclusao, UsuarioInclusaoData, isnull(UsuarioAlteracao,'') as UsuarioAlteracao, UsuarioAlteracaoData, " & vbCrLf & _
              "       isnull(UsuarioCancelamento,'') as UsuarioCancelamento, UsuarioCancelamentoData, Observacoes, isnull(Nota,0) as Nota, isnull(Serie,'') as Serie, isnull(EntradaSaida,'') as EntradaSaida" & vbCrLf & _
              "  FROM VendasPorTrocas" & vbCrLf

        If pRegistro > 0 Then
            sql &= " WHERE Registro_Id = " & pRegistro
        ElseIf Not NF Is Nothing Then
            If NF.SubOperacao.VendaTroca Then
                sql &= " WHERE EmpresaDebito    ='" & NF.CodigoEmpresa & "'" & vbCrLf & _
                                       "   And EndEmpresaDebito = " & NF.EnderecoEmpresa & vbCrLf & _
                                       "   And ClienteDebito    ='" & NF.CodigoCliente & "'" & vbCrLf & _
                                       "   And EndClienteDebito = " & NF.EnderecoCliente & vbCrLf
            End If

            If NF.SubOperacao.CompraTroca Then
                sql &= " WHERE EmpresaCredito    ='" & NF.CodigoEmpresa & "'" & vbCrLf & _
                       "   And EndEmpresaCredito = " & NF.EnderecoEmpresa & vbCrLf & _
                       "   And ClienteCredito    ='" & NF.CodigoCliente & "'" & vbCrLf & _
                       "   And EndClienteCredito = " & NF.EnderecoCliente & vbCrLf
            End If

            sql &= " And Nota             = " & NF.Codigo & vbCrLf & _
                   " And Serie            ='" & NF.Serie & "'" & vbCrLf & _
                   " And EntradaSaida     ='" & NF.EntradaSaida & "'"
        End If

        Dim ds As DataSet
        Dim Banco As New AcessaBanco
        ds = Banco.ConsultaDataSet(sql, "VPT")

        If ds Is Nothing OrElse ds.Tables(0).Rows.Count > 0 Then
            Dim row As DataRow = ds.Tables(0).Rows(0)

            _Registro = row("Registro_Id")
            _Sequencia = row("Sequencia_Id")
            _CodigoIndexador = row("Indexador")
            _CodigoMoeda = row("Moeda")
            _CodigoSituacao = row("Situacao")
            _Lote = row("Lote")
            _Movimento = row("Movimento")
            _Vencimento = row("Vencimento")

            _CodigoUnidadeDebito = row("UnidadeDebito")
            _CodigoEmpresaDebito = row("EmpresaDebito")
            _EndEmpresaDebito = row("EndEmpresaDebito")
            _CodigoClienteDebito = row("ClienteDebito")
            _EndClienteDebito = row("EndClienteDebito")
            _CodigoPedidoDebito = row("PedidoDebito")
            _CodigoCarteiraDebito = row("CarteiraDebito")
            _NumeroNota = row("Nota")
            _Serie = row("Serie")
            _EntradaSaida = row("EntradaSaida")

            _CodigoUnidadeCredito = row("UnidadeCredito")
            _CodigoEmpresaCredito = row("EmpresaCredito")
            _EndEmpresaCredito = row("EndEmpresaCredito")
            _CodigoClienteCredito = row("ClienteCredito")
            _EndClienteCredito = row("EndClienteCredito")
            _CodigoPedidoCredito = row("PedidoCredito")
            _CodigoCarteiraCredito = row("CarteiraCredito")

            _ValorOficial = row("ValorOficial")
            _ValorMoeda = row("ValorMoeda")
            _TaxaDeJuro = row("TaxaDeJuros")
            _Historico = row("Historico")
            _Observacao = row("Observacoes")

            _UsuarioInclusao = row("UsuarioInclusao")
            _UsuarioInclusaoData = row("UsuarioInclusaoData")
            _UsuarioAlteracao = row("UsuarioAlteracao")
            _UsuarioAlteracaoData = row("UsuarioAlteracaoData")
            _UsuarioCancelamento = row("UsuarioCancelamento")
            _UsuarioCancelamentoData = row("UsuarioCancelamentoData")
        ElseIf Not NF Is Nothing Then
            '_CodigoIndexador = NF.Pedido.CodigoIndexador
            '_CodigoMoeda = NF.Pedido.CodigoMoeda
            '_CodigoSituacao = NF.CodigoSituacao
            '_Lote = 75
            '_Movimento = NF.Movimento
            '_Vencimento = NF.Pedido.VencimentoPedido
            '_CodigoUnidadeDebito = NF.Pedido.CodigoUnidadeNegocio
            '_CodigoEmpresaDebito = NF.CodigoEmpresa
            '_EndEmpresaDebito = NF.EnderecoEmpresa
            '_CodigoClienteDebito = row("ClienteDebito")
            '_EndClienteDebito = row("EndClienteDebito")
            '_CodigoPedidoDebito = row("PedidoDebito")
            '_CodigoCarteiraDebito = row("CarteiraDebito")
            '_NumeroNota = row("Nota")
            '_Serie = row("Serie")
            '_EntradaSaida = row("EntradaSaida")
            '_CodigoUnidadeCredito = row("UnidadeCredito")
            '_CodigoEmpresaCredito = row("EmpresaCredito")
            '_EndEmpresaCredito = row("EndEmpresaCredito")
            '_CodigoClienteCredito = row("ClienteCredito")
            '_EndClienteCredito = row("EndClienteCredito")
            '_CodigoPedidoCredito = row("PedidoCredito")
            '_CodigoCarteiraCredito = row("CarteiraCredito")
            '_ValorOficial = row("ValorOficial")
            '_ValorMoeda = row("ValorMoeda")
            '_TaxaDeJuro = row("TaxaDeJuros")
            '_Historico = row("Historico")
            '_Observacao = row("Observacoes")
            '_UsuarioInclusao = row("UsuarioInclusao")
            '_UsuarioInclusaoData = row("UsuarioInclusaoData")
            '_UsuarioAlteracao = row("UsuarioAlteracao")
            '_UsuarioAlteracaoData = row("UsuarioAlteracaoData")
            '_UsuarioCancelamento = row("UsuarioCancelamento")
            '_UsuarioCancelamentoData = row("UsuarioCancelamentoData")
        End If
    End Sub
#End Region

#Region "Fields"
    Private _NF As NotaFiscal
    Private _IUD As String
    Private _Registro As Integer
    Private _Sequencia As Integer
    Private _CodigoIndexador As Integer
    Private _indexador As Negocio.Indexador
    Private _CodigoMoeda As Integer
    Private _Moeda As Negocio.Moeda
    Private _CodigoSituacao As Integer
    Private _Situacao As Negocio.Situacao
    Private _Lote As Integer
    Private _Movimento As DateTime
    Private _Vencimento As DateTime

    Private _CodigoUnidadeDebito As String
    Private _UnidadeDebito As Negocio.Cliente
    Private _CodigoEmpresaDebito As String
    Private _EndEmpresaDebito As Integer
    Private _EmpresaDebito As Negocio.Cliente
    Private _CodigoClienteDebito As String
    Private _EndClienteDebito As Integer
    Private _ClienteDebito As Negocio.Cliente
    Private _CodigoPedidoDebito As Integer
    Private _PedidoDebito As Negocio.Pedido
    Private _CodigoCarteiraDebito As String
    Private _CarteiraDebito As Negocio.CarteiraFinanceira

    Private _NumeroNota As Integer
    Private _Serie As String
    Private _EntradaSaida As String
    Private _NotaFiscal As Negocio.NotaFiscal

    Private _CodigoUnidadeCredito As String
    Private _UnidadeCredito As Negocio.Cliente
    Private _CodigoEmpresaCredito As String
    Private _EndEmpresaCredito As Integer
    Private _EmpresaCredito As Negocio.Cliente
    Private _CodigoClienteCredito As String
    Private _EndClienteCredito As Integer
    Private _ClienteCredito As Negocio.Cliente
    Private _CodigoPedidoCredito As Integer
    Private _PedidoCredito As Negocio.Pedido
    Private _CodigoCarteiraCredito As String
    Private _CarteiraCredito As Negocio.CarteiraFinanceira

    Private _ValorOficial As Decimal
    Private _ValorMoeda As Decimal
    Private _TaxaDeJuro As Decimal
    Private _Historico As String
    Private _Observacao As String

    Private _UsuarioInclusao As String
    Private _UsuarioInclusaoData As DateTime
    Private _UsuarioAlteracao As String
    Private _UsuarioAlteracaoData As DateTime
    Private _UsuarioCancelamento As String
    Private _UsuarioCancelamentoData As String

    Private _Adiantamento As Negocio.Adiantamento

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

    Public Property Registro() As Integer
        Get
            Return _Registro
        End Get
        Set(ByVal value As Integer)
            _Registro = value
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

    Public Property CodigoIndexador() As Integer
        Get
            Return _CodigoIndexador
        End Get
        Set(ByVal value As Integer)
            _CodigoIndexador = value
            _indexador = Nothing
        End Set
    End Property

    Public Property indexador() As Negocio.Indexador
        Get
            If _indexador Is Nothing And _CodigoIndexador > 0 Then _indexador = New Negocio.Indexador(_CodigoIndexador)
            Return _indexador
        End Get
        Set(ByVal value As Negocio.Indexador)
            _indexador = value
        End Set
    End Property

    Public Property CodigoMoeda() As Integer
        Get
            Return _CodigoMoeda
        End Get
        Set(ByVal value As Integer)
            _CodigoMoeda = value
            _Moeda = Nothing
        End Set
    End Property

    Public Property Moeda() As Negocio.Moeda
        Get
            If _Moeda Is Nothing And _CodigoMoeda > 0 Then _Moeda = New Negocio.Moeda(_CodigoMoeda)
            Return _Moeda
        End Get
        Set(ByVal value As Negocio.Moeda)
            _Moeda = value
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

    Public Property Situacao() As Negocio.Situacao
        Get
            If _Situacao Is Nothing And _CodigoSituacao > 0 Then _Situacao = New Negocio.Situacao(_CodigoSituacao)
            Return _Situacao
        End Get
        Set(ByVal value As Negocio.Situacao)
            _Situacao = value
        End Set
    End Property

    Public Property Lote() As Integer
        Get
            Return _Lote
        End Get
        Set(ByVal value As Integer)
            _Lote = value
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

    Public Property Vencimento() As DateTime
        Get
            Return _Vencimento
        End Get
        Set(ByVal value As DateTime)
            _Vencimento = value
        End Set
    End Property

    '********** Debito  *********************************

    Public Property CodigoUnidadeDebito() As String
        Get
            Return _CodigoUnidadeDebito
        End Get
        Set(ByVal value As String)
            _CodigoUnidadeDebito = value
            _UnidadeDebito = Nothing
        End Set
    End Property

    Public Property UnidadeDebito() As Negocio.Cliente
        Get
            If _UnidadeDebito Is Nothing And _CodigoUnidadeDebito.Length > 0 Then _UnidadeDebito = New Negocio.Cliente(_CodigoUnidadeDebito, 0)
            Return _UnidadeDebito
        End Get
        Set(ByVal value As Negocio.Cliente)
            _UnidadeDebito = value
        End Set
    End Property

    Public Property CodigoEmpresaDebito() As String
        Get
            Return _CodigoEmpresaDebito
        End Get
        Set(ByVal value As String)
            _CodigoEmpresaDebito = value
            _EmpresaDebito = Nothing
            _NotaFiscal = Nothing
        End Set
    End Property

    Public Property EndEmpresaDebito() As Integer
        Get
            Return _EndEmpresaDebito
        End Get
        Set(ByVal value As Integer)
            _EndEmpresaDebito = value
            _EmpresaDebito = Nothing
            _NotaFiscal = Nothing
        End Set
    End Property

    Public Property EmpresaDebito() As Negocio.Cliente
        Get
            If _EmpresaDebito Is Nothing And _CodigoEmpresaDebito.Length > 0 Then _EmpresaDebito = New Negocio.Cliente(_CodigoEmpresaDebito, _EndEmpresaDebito)
            Return _EmpresaDebito
        End Get
        Set(ByVal value As Negocio.Cliente)
            _EmpresaDebito = value
        End Set
    End Property

    Public Property CodigoClienteDebito() As String
        Get
            Return _CodigoClienteDebito
        End Get
        Set(ByVal value As String)
            _CodigoClienteDebito = value
            _ClienteDebito = Nothing
            _NotaFiscal = Nothing
        End Set
    End Property

    Public Property EndClienteDebito() As Integer
        Get
            Return _EndClienteDebito
        End Get
        Set(ByVal value As Integer)
            _EndClienteDebito = value
            _ClienteDebito = Nothing
            _NotaFiscal = Nothing
        End Set
    End Property

    Public Property ClienteDebito() As Negocio.Cliente
        Get
            If _ClienteDebito Is Nothing And _CodigoClienteDebito.Length > 0 Then _ClienteDebito = New Negocio.Cliente(_CodigoClienteDebito, _EndClienteDebito)
            Return _ClienteDebito
        End Get
        Set(ByVal value As Negocio.Cliente)
            _ClienteDebito = value
        End Set
    End Property

    Public Property CodigoPedidoDebito() As Integer
        Get
            Return _CodigoPedidoDebito
        End Get
        Set(ByVal value As Integer)
            _CodigoPedidoDebito = value
            _PedidoDebito = Nothing
        End Set
    End Property

    Public Property PedidoDebito() As Negocio.Pedido
        Get
            If _PedidoDebito Is Nothing And _CodigoPedidoDebito > 0 Then _PedidoDebito = New Negocio.Pedido(_CodigoEmpresaDebito, _EndEmpresaDebito, _CodigoPedidoDebito)
            Return _PedidoDebito
        End Get
        Set(ByVal value As Negocio.Pedido)
            _PedidoDebito = value
        End Set
    End Property

    Public Property CodigoCarteiraDebito() As String
        Get
            Return _CodigoCarteiraDebito
        End Get
        Set(ByVal value As String)
            _CodigoCarteiraDebito = value
            _CarteiraDebito = Nothing
        End Set
    End Property

    Public Property CarteiraDebito() As Negocio.CarteiraFinanceira
        Get
            Return _CarteiraDebito
        End Get
        Set(ByVal value As Negocio.CarteiraFinanceira)
            _CarteiraDebito = value
        End Set
    End Property

    Public Property NumeroNota() As Integer
        Get
            Return _NumeroNota
        End Get
        Set(ByVal value As Integer)
            _NumeroNota = value
            _NotaFiscal = Nothing
        End Set
    End Property

    Public Property Serie() As String
        Get
            Return _Serie
        End Get
        Set(ByVal value As String)
            _Serie = value
            _NotaFiscal = Nothing
        End Set
    End Property

    Public Property EntradaSaida() As String
        Get
            Return _EntradaSaida
        End Get
        Set(ByVal value As String)
            _EntradaSaida = value
            _NotaFiscal = Nothing
        End Set
    End Property

    Public ReadOnly Property NotaFiscal() As Negocio.NotaFiscal
        Get
            If _NotaFiscal Is Nothing And _CodigoEmpresaDebito.Length > 0 And _CodigoClienteDebito.Length > 0 And _NumeroNota > 0 And _Serie.Length > 0 And _EntradaSaida.Length > 0 Then
                Dim nf As New Negocio.NotaFiscal
                nf.CodigoEmpresa = _CodigoEmpresaDebito
                nf.EnderecoEmpresa = _EndEmpresaDebito
                nf.CodigoCliente = _CodigoClienteDebito
                nf.EnderecoCliente = _EndClienteDebito
                nf.Codigo = _NumeroNota
                nf.Serie = _Serie
                nf.EntradaSaida = IIf(_EntradaSaida = "E", Negocio.eEntradaSaida.Entrada, Negocio.eEntradaSaida.Saida)
                _NotaFiscal = New Negocio.NotaFiscal(nf)
            End If
            Return _NotaFiscal
        End Get
    End Property


    '********** Credito  *********************************

    Public Property CodigoUnidadeCredito() As String
        Get
            Return _CodigoUnidadeCredito
        End Get
        Set(ByVal value As String)
            _CodigoUnidadeCredito = value
            _UnidadeCredito = Nothing
        End Set
    End Property

    Public Property UnidadeCredito() As Negocio.Cliente
        Get
            If _UnidadeCredito Is Nothing And _CodigoUnidadeCredito.Length > 0 Then _UnidadeCredito = New Negocio.Cliente(_CodigoUnidadeCredito, 0)
            Return _UnidadeCredito
        End Get
        Set(ByVal value As Negocio.Cliente)
            _UnidadeCredito = value
        End Set
    End Property

    Public Property CodigoEmpresaCredito() As String
        Get
            Return _CodigoEmpresaCredito
        End Get
        Set(ByVal value As String)
            _CodigoEmpresaCredito = value
            _EmpresaCredito = Nothing
        End Set
    End Property

    Public Property EndEmpresaCredito() As Integer
        Get
            Return _EndEmpresaCredito
        End Get
        Set(ByVal value As Integer)
            _EndEmpresaCredito = value
            _EmpresaCredito = Nothing
        End Set
    End Property

    Public Property EmpresaCredito() As Negocio.Cliente
        Get
            If _EmpresaCredito Is Nothing And _CodigoEmpresaCredito.Length > 0 Then _EmpresaCredito = New Negocio.Cliente(_CodigoEmpresaCredito, _EndEmpresaCredito)
            Return _EmpresaCredito
        End Get
        Set(ByVal value As Negocio.Cliente)
            _EmpresaCredito = value
        End Set
    End Property

    Public Property CodigoClienteCredito() As String
        Get
            Return _CodigoClienteCredito
        End Get
        Set(ByVal value As String)
            _CodigoClienteCredito = value
            _ClienteCredito = Nothing
        End Set
    End Property

    Public Property EndClienteCredito() As Integer
        Get
            Return _EndClienteCredito
        End Get
        Set(ByVal value As Integer)
            _EndClienteCredito = value
            _ClienteCredito = Nothing
        End Set
    End Property

    Public Property ClienteCredito() As Negocio.Cliente
        Get
            If _ClienteCredito Is Nothing And _CodigoClienteCredito.Length > 0 Then _ClienteCredito = New Negocio.Cliente(_CodigoClienteCredito, _EndClienteCredito)
            Return _ClienteCredito
        End Get
        Set(ByVal value As Negocio.Cliente)
            _ClienteCredito = value
        End Set
    End Property

    Public Property CodigoPedidoCredito() As Integer
        Get
            Return _CodigoPedidoCredito
        End Get
        Set(ByVal value As Integer)
            _CodigoPedidoCredito = value
            _PedidoCredito = Nothing
        End Set
    End Property

    Public Property PedidoCredito() As Negocio.Pedido
        Get
            If _PedidoCredito Is Nothing And _CodigoPedidoCredito > 0 Then _PedidoCredito = New Negocio.Pedido(_CodigoEmpresaCredito, _EndEmpresaCredito, _CodigoPedidoCredito)
            Return _PedidoCredito
        End Get
        Set(ByVal value As Negocio.Pedido)
            _PedidoCredito = value
        End Set
    End Property

    Public Property CodigoCarteiraCredito() As String
        Get
            Return _CodigoCarteiraCredito
        End Get
        Set(ByVal value As String)
            _CodigoCarteiraCredito = value
            _CarteiraCredito = Nothing
        End Set
    End Property

    Public Property CarteiraCredito() As Negocio.CarteiraFinanceira
        Get
            Return _CarteiraCredito
        End Get
        Set(ByVal value As Negocio.CarteiraFinanceira)
            _CarteiraCredito = value
        End Set
    End Property

    Public Property ValorOficial() As Decimal
        Get
            Return _ValorOficial
        End Get
        Set(ByVal value As Decimal)
            _ValorOficial = value
        End Set
    End Property

    Public Property ValorMoeda() As Decimal
        Get
            Return _ValorMoeda
        End Get
        Set(ByVal value As Decimal)
            _ValorMoeda = value
        End Set
    End Property

    Public Property TaxaDeJuro() As Decimal
        Get
            Return _TaxaDeJuro
        End Get
        Set(ByVal value As Decimal)
            _TaxaDeJuro = value
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

    Public Property Observacao() As String
        Get
            Return _Observacao
        End Get
        Set(ByVal value As String)
            _Observacao = value
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

    Public Property UsuarioInclusaoData() As DateTime
        Get
            Return _UsuarioInclusaoData
        End Get
        Set(ByVal value As DateTime)
            _UsuarioInclusaoData = value
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

    Public Property UsuarioAlteracaoData() As DateTime
        Get
            Return _UsuarioAlteracaoData
        End Get
        Set(ByVal value As DateTime)
            _UsuarioAlteracaoData = value
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

    Public Property UsuarioCancelamentoData() As String
        Get
            Return _UsuarioCancelamentoData
        End Get
        Set(ByVal value As String)
            _UsuarioCancelamentoData = value
        End Set
    End Property

    Public Property Adiantamento() As Negocio.Adiantamento
        Get
            If _Adiantamento Is Nothing And _Registro > 0 Then
                Dim Ad As New Negocio.Adiantamento
                Ad.CodigoEmpresa = _CodigoEmpresaDebito
                Ad.EndEmpresa = _EndEmpresaDebito
                Ad.CodigoCliente = _CodigoClienteDebito
                Ad.EndCliente = _CodigoClienteDebito
                Ad.CodigoTitulo = _Registro
                _Adiantamento = New Negocio.Adiantamento(Ad)
            End If

            Return _Adiantamento
        End Get
        Set(ByVal value As Negocio.Adiantamento)
            _Adiantamento = value
        End Set
    End Property

#End Region

#Region "Methods"
    Public Function Salvar() As Boolean
        If IUD = Nothing Then Return True
        Dim Banco As New AcessaBanco
        Dim Sqls As New ArrayList

        Sqls.Clear()
        SalvarSql(Sqls)
        Return Banco.GravaBanco(Sqls)
    End Function

    Public Sub SalvarSql(ByRef Sqls As ArrayList)
        Dim Sql As String = ""
        Select Case Me.IUD
            Case "I"
                Dim num As New Negocio.Numerador(1)
                _Registro = num.Sequencia
                Sqls.Add(num.IncrementarNumeradorSql)

                Sql = "Insert Into VendasPorTrocas(Registro_Id, Sequencia_Id, Indexador, Moeda, Situacao, Lote, Movimento, Vencimento, UnidadeDebito, EmpresaDebito, EndEmpresaDebito, ClienteDebito, " & vbCrLf & _
                      "       EndClienteDebito, PedidoDebito, CarteiraDebito, Nota, Serie, EntradaSaida, UnidadeCredito, EmpresaCredito, EndEmpresaCredito, ClienteCredito, EndClienteCredito, PedidoCredito, " & vbCrLf & _
                      "       CarteiraCredito, ValorOficial, ValorMoeda, TaxaDeJuros, Historico, Observacoes, UsuarioInclusao, UsuarioInclusaoData)" & vbCrLf & _
                      " values (" & _Registro & "," & _Sequencia & "," & _CodigoIndexador & "," & _CodigoMoeda & "," & _CodigoSituacao & "," & _Lote & ",'" & _Movimento.ToString("yyyy-MM-dd") & "','" & _Vencimento.ToString("yyyy-MM-dd") & "','" & _CodigoUnidadeDebito & "','" & _CodigoEmpresaDebito & "'," & _EndEmpresaDebito & ",'" & _CodigoClienteDebito & "'," & vbCrLf & _
                      "         " & _EndClienteDebito & "," & _CodigoPedidoDebito & ",'" & _CodigoCarteiraDebito & "'," & _NumeroNota & ",'" & _Serie & "','" & _EntradaSaida & "','" & CodigoUnidadeCredito & "','" & _CodigoEmpresaCredito & "'," & _EndEmpresaCredito & ",'" & _CodigoClienteCredito & "'," & _EndClienteCredito & "," & _CodigoPedidoCredito & "," & vbCrLf & _
                      "        '" & _CodigoCarteiraCredito & "'," & Str(_ValorOficial) & "," & Str(_ValorMoeda) & "," & Str(_TaxaDeJuro) & ",'" & _Historico & "','" & _Observacao & "','" & _UsuarioInclusao & "','" & _UsuarioInclusaoData.ToString("yyyy-MM-dd") & "'"

                Sqls.Add(Sql)
                SalvaTabelasRelacionadas(Sqls)
            Case "U"
                Sql = " UPDATE VendasPorTrocas SET" & vbCrLf & _
                      "    Indexador            = " & _CodigoIndexador & vbCrLf & _
                      "   ,Moeda                = " & _CodigoMoeda & vbCrLf & _
                      "   ,Situacao             = " & _CodigoSituacao & vbCrLf & _
                      "   ,Lote                 = " & _Lote & vbCrLf & _
                      "   ,Movimento            ='" & _Movimento.ToString("yyyy-MM-dd") & "'" & vbCrLf & _
                      "   ,Vencimento           ='" & _Vencimento.ToString("yyyy-MM-dd") & "'" & vbCrLf & _
                      "   ,UnidadeDebito        ='" & _CodigoUnidadeDebito & "'" & vbCrLf & _
                      "   ,EmpresaDebito        ='" & _CodigoEmpresaDebito & "'" & vbCrLf & _
                      "   ,EndEmpresaDebito     = " & _EndEmpresaDebito & vbCrLf & _
                      "   ,ClienteDebito        ='" & _CodigoCarteiraDebito & "'" & vbCrLf & _
                      "   ,EndClienteDebito     = " & _EndClienteDebito & vbCrLf & _
                      "   ,PedidoDebito         = " & _CodigoPedidoDebito & vbCrLf & _
                      "   ,CarteiraDebito       ='" & _CodigoCarteiraDebito & "'" & vbCrLf & _
                      "   ,Nota                 = " & _NumeroNota & vbCrLf & _
                      "   ,Serie                ='" & _Serie & "'" & vbCrLf & _
                      "   ,EntradaSaida         ='" & _EntradaSaida & "'" & vbCrLf & _
                      "   ,UnidadeCredito       ='" & _CodigoUnidadeCredito & "'" & vbCrLf & _
                      "   ,EmpresaCredito       ='" & _CodigoEmpresaCredito & "'" & vbCrLf & _
                      "   ,EndEmpresaCredito    = " & _EndEmpresaCredito & vbCrLf & _
                      "   ,ClienteCredito       ='" & _CodigoClienteCredito & "'" & vbCrLf & _
                      "   ,EndClienteCredito    = " & _EndClienteCredito & vbCrLf & _
                      "   ,PedidoCredito        = " & _CodigoPedidoCredito & vbCrLf & _
                      "   ,CarteiraCredito      ='" & _CodigoCarteiraCredito & "'" & vbCrLf & _
                      "   ,ValorOficial         = " & Str(_ValorOficial) & vbCrLf & _
                      "   ,ValorMoeda           = " & Str(_ValorMoeda) & vbCrLf & _
                      "   ,TaxaDeJuros          = " & Str(_TaxaDeJuro) & vbCrLf & _
                      "   ,Historico            ='" & _Historico & "'" & vbCrLf & _
                      "   ,Observacoes          ='" & _Observacao & "'" & vbCrLf & _
                      "   ,UsuarioAlteracao     ='" & _UsuarioAlteracao & "'" & vbCrLf & _
                      "   ,UsuarioAlteracaoData ='" & _UsuarioAlteracaoData.ToString("yyyy-MM-dd") & vbCrLf & _
                      "  WHERE Registro_Id  = " & _Registro & vbCrLf & _
                      "    And Sequencia_Id = " & _Sequencia
                Sqls.Add(Sql)
                SalvaTabelasRelacionadas(Sqls)
            Case "C"
                SalvaTabelasRelacionadas(Sqls)
                Sql = " UPDATE VendasPorTrocas SET" & vbCrLf & _
                      "    Situacao                = " & _CodigoSituacao & vbCrLf & _
                      "   ,UsuarioCancelamento     ='" & _UsuarioCancelamento & "'" & vbCrLf & _
                      "   ,UsuarioCancelamentoData ='" & String.Format("{0:yyyy-MM-dd}", _UsuarioCancelamentoData) & vbCrLf & _
                      "  WHERE Registro_Id  = " & _Registro & vbCrLf & _
                      "    And Sequencia_Id = " & _Sequencia
                Sqls.Add(Sql)
            Case "D"
                SalvaTabelasRelacionadas(Sqls)
                Sql = " DELETE VendasPorTrocas" & vbCrLf & _
                      "  WHERE Registro_Id  = " & _Registro & vbCrLf & _
                      "    And Sequencia_Id = " & _Sequencia
                Sqls.Add(Sql)
        End Select
    End Sub

    Public Sub SalvaTabelasRelacionadas(ByRef Sqls As ArrayList)
        '*******************************************************************************************************************
        '******************************************  ADIANTAMENTO / RAZAO  *************************************************
        '*******************************************************************************************************************
        Adiantamento.IUD = IUD
        Select Case IUD
            Case "I"
                If Not _NotaFiscal Is Nothing And _NotaFiscal.SubOperacao.Adiantamento Then
                    Adiantamento.CodigoEmpresa = _CodigoEmpresaDebito
                    Adiantamento.EndEmpresa = _EndEmpresaDebito
                    Adiantamento.CodigoCliente = _CodigoClienteDebito
                    Adiantamento.EndCliente = _EndClienteDebito
                    Adiantamento.RegistroPedido = _CodigoPedidoDebito
                    Adiantamento.CodigoTitulo = _Registro
                    Adiantamento.Recibo = 0
                    Adiantamento.CodigoSafra = PedidoDebito.CodigoSafra
                    Adiantamento.Movimento = _Movimento
                    Adiantamento.Vencimento = _Vencimento
                    Adiantamento.ValorOficial = _ValorOficial
                    Adiantamento.ValorMoeda = _ValorMoeda
                    Adiantamento.CodigoIndexador = _CodigoIndexador
                    Adiantamento.CodigoMoeda = _CodigoMoeda
                    Adiantamento.Taxa = _TaxaDeJuro
                    Adiantamento.Periodicidade = 1
                    Adiantamento.TipoDeTaxa = 0
                    Adiantamento.UltimaAtualizacao = _Movimento
                    Adiantamento.Duplicata = ""
                    Adiantamento.EmissaoDuplicata = _Movimento
                    Adiantamento.Numero_Nf = _NumeroNota
                    Adiantamento.Serie_Nf = _Serie
                    Adiantamento.EntradaSaida_Nf = _EntradaSaida
                    Adiantamento.Informacoes = "Adiantamento..."
                    Adiantamento.Observacoes = "Adiantamento..."
                    Adiantamento.DataDeCalculo = _Vencimento
                    Adiantamento.SalvarSql(Sqls)

                    'Dim SaldoOrigem As New Negocio.ListSaldoPedidoXNota(Adiantamento.CodigoEmpresa, Adiantamento.EndEmpresa, _NF.Pedido.PedidosDeTroca(0).CodigoPedidoOrigem)
                    'Dim SaldoDestino As New Negocio.ListSaldoPedidoXNota(Adiantamento.CodigoEmpresa, Adiantamento.EndEmpresa, _NF.Pedido.PedidosDeTroca(0).CodigoPedidoDestino)
                    Dim SaldoOrigem As New Negocio.ListSaldoPedidoXNota()
                    Dim SaldoDestino As New Negocio.ListSaldoPedidoXNota()

                    If SaldoDestino(0).EntregueFiscal > SaldoOrigem(0).EntregueFiscal Then
                        Dim VlrBaixaAdiantamento As Decimal
                        If SaldoDestino(0).EntregueFiscal - SaldoOrigem(0).EntregueFiscal > _NF.TotalNota Then
                            VlrBaixaAdiantamento = _NF.TotalNota
                        Else
                            VlrBaixaAdiantamento = (SaldoDestino(0).EntregueFiscal - SaldoOrigem(0).EntregueFiscal)
                        End If
                        Dim AxB As New Negocio.AdiantamentoBaixa

                    End If

                End If
                LanctosRazao(Sqls)
            Case "U"
                If Not _NotaFiscal Is Nothing And _NotaFiscal.SubOperacao.Adiantamento Then Adiantamento.SalvarSql(Sqls)
                DeletaRazao(Sqls)
                LanctosRazao(Sqls)
            Case "C" Or "D"
                DeletaRazao(Sqls)
                If Not _NotaFiscal Is Nothing And _NotaFiscal.SubOperacao.Adiantamento Then Adiantamento.SalvarSql(Sqls)
        End Select
    End Sub

    Public Sub DeletaRazao(ByRef sqls As ArrayList)
        Dim sql As String
        sql = "Delete Razao " & vbCrLf & _
              " Where Empresa_Id    ='" & _CodigoEmpresaDebito & "'" & vbCrLf & _
              "   and EndEmpresa_Id = " & _EndEmpresaDebito & vbCrLf & _
              "   and Conta_Id      ='" & CarteiraDebito.CodigoContaCliente & "'" & vbCrLf & _
              "   and Cliente_Id    ='" & _CodigoClienteDebito & "'" & vbCrLf & _
              "   and EndCliente_Id = " & _EndClienteDebito & vbCrLf & _
              "   and Movimento_Id  ='" & _Movimento.ToString("yyyy-MM-dd") & "'" & vbCrLf & _
              "   and Lote_Id       = " & _Lote & vbCrLf & _
              "   and Sequencia_Id  = " & _Sequencia & vbCrLf & _
              "   and Titulo        = " & _Registro

        sqls.Add(sql)
    End Sub

    Public Sub LanctosRazao(ByRef Sqls As ArrayList)
        '********* Debito ************************
        Dim RazDeb As New Negocio.Razao
        RazDeb.IUD = IUD
        RazDeb.CodigoEmpresa = _CodigoEmpresaDebito
        RazDeb.EndEmpresa = _EndEmpresaDebito
        RazDeb.CodigoConta = CarteiraDebito.CodigoContaCliente
        RazDeb.CodigoCliente = _CodigoClienteDebito
        RazDeb.EnderecoCliente = _EndClienteDebito
        RazDeb.Movimento = _Movimento
        RazDeb.Lote = _Lote
        RazDeb.Sequencia = _Sequencia
        RazDeb.CodigoTitulo = _Registro
        RazDeb.CodigoUnidadeDeNegocio = _CodigoUnidadeDebito
        RazDeb.CodigoIndexador = _CodigoIndexador
        RazDeb.DataMoeda = _Movimento
        RazDeb.DebitoOficial = _ValorOficial
        RazDeb.CreditoOficial = 0
        RazDeb.DebitoMoeda = _ValorMoeda
        RazDeb.CreditoMoeda = 0
        RazDeb.Historico = _Historico
        RazDeb.PrevistoRealizado = "P"
        RazDeb.SalvarSql(Sqls)

        '********* Credito ***********************
        Dim RazCred As New Negocio.Razao
        RazCred.IUD = IUD
        RazCred.CodigoEmpresa = _CodigoEmpresaCredito
        RazCred.EndEmpresa = _EndEmpresaCredito
        RazCred.CodigoConta = CarteiraCredito.CodigoContaCliente
        RazCred.CodigoCliente = _CodigoClienteCredito
        RazCred.EnderecoCliente = _EndClienteCredito
        RazCred.Movimento = _Movimento
        RazCred.Lote = _Lote
        RazCred.Sequencia = _Sequencia
        RazCred.CodigoTitulo = _Registro
        RazCred.CodigoUnidadeDeNegocio = _CodigoUnidadeCredito
        RazCred.CodigoIndexador = _CodigoIndexador
        RazCred.DataMoeda = _Movimento
        RazCred.DebitoOficial = 0
        RazCred.CreditoOficial = _ValorOficial
        RazCred.DebitoMoeda = 0
        RazCred.CreditoMoeda = _ValorMoeda
        RazCred.Historico = _Historico
        RazCred.PrevistoRealizado = "P"
        RazCred.SalvarSql(Sqls)

        '******* Transferencias Financeiras ******
        Dim Trans As New Negocio.TransferenciaFinanceira
        If Not _NotaFiscal.SubOperacao.Devolucao Then
            Trans.CodigoEmpresaDebito = _CodigoEmpresaDebito
            Trans.EndEmpresaDebito = _EndEmpresaDebito
            Trans.CodigoEmpresaCredito = _CodigoEmpresaCredito
            Trans.EndEmpresaCredito = _EndEmpresaCredito
        Else
            Trans.CodigoEmpresaDebito = _CodigoEmpresaCredito
            Trans.EndEmpresaDebito = _EndEmpresaCredito
            Trans.CodigoEmpresaCredito = CodigoEmpresaDebito
            Trans.EndEmpresaCredito = _EndEmpresaDebito
        End If

        Dim ListTrans As New Negocio.ListTransferenciaFinanceira(EmpresaDebito, EmpresaCredito)

        For Each T As Negocio.TransferenciaFinanceira In ListTrans
            Dim RazTrans As New Negocio.Razao
            RazTrans.IUD = IUD
            RazTrans.CodigoEmpresa = T.CodigoEmpresaContabil
            RazTrans.EndEmpresa = T.EndEmpresaContabil
            RazTrans.CodigoConta = T.ContaContabil
            RazTrans.CodigoCliente = T.CodigoClienteContabil
            RazTrans.EnderecoCliente = T.EndClienteContabil
            RazTrans.Movimento = _Movimento
            RazTrans.Lote = _Lote
            RazTrans.Sequencia = _Sequencia
            RazTrans.CodigoTitulo = _Registro
            RazTrans.CodigoUnidadeDeNegocio = _CodigoUnidadeCredito
            RazTrans.CodigoIndexador = _CodigoIndexador
            RazTrans.DataMoeda = _Movimento

            If T.DebitoCredito = "D" Then
                RazDeb.DebitoOficial = _ValorOficial
                RazDeb.CreditoOficial = 0
                RazDeb.DebitoMoeda = _ValorMoeda
                RazDeb.CreditoMoeda = 0
            Else
                RazTrans.DebitoOficial = 0
                RazTrans.CreditoOficial = _ValorOficial
                RazTrans.DebitoMoeda = 0
                RazTrans.CreditoMoeda = _ValorMoeda
            End If

            RazTrans.Historico = _Historico
            RazTrans.PrevistoRealizado = "P"
            RazTrans.SalvarSql(Sqls)
        Next
    End Sub
#End Region

End Class