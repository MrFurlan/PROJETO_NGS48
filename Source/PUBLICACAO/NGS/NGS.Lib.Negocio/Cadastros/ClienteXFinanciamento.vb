Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Collections.Generic
Imports NGS.Lib.Uteis

<Serializable()> _
Public Class ListClienteXFinanciamento
    Inherits List(Of ClienteXFinanciamento)

#Region "Construtor"
    Public Sub New()

    End Sub

    Public Sub New(ByVal pcliente As Cliente)
        _Cliente = pcliente
        Dim sql As String
        sql = "SELECT CxF.Cliente_Id, CxF.EndCliente_Id, CxF.Financiador_Id, CxF.EndFinanciador_Id, Financiador.Nome AS NomeFinanciador, " & vbCrLf & _
              "       CxF.Financiamento_Id, CxF.TipoFinanciamento, CxF.Safra, CxF.DataFinanciamento, CxF.DataVencimento, CxF.NumeroDeParcelas, " & vbCrLf & _
              "       CxF.Produto, Produtos.Nome AS NomeProduto, CxF.Quantidade, CxF.Moeda, CxF.ValorOficial, CxF.ValorMoeda, CxF.Observacao " & vbCrLf & _
              "  FROM ClienteXFinanciamento AS CxF " & vbCrLf & _
              "  INNER JOIN Clientes AS Financiador " & vbCrLf & _
              "          ON CxF.Financiador_Id = Financiador.Cliente_Id " & vbCrLf & _
              "         AND CxF.EndFinanciador_Id = Financiador.Endereco_Id " & vbCrLf & _
              "  INNER JOIN Produtos " & vbCrLf & _
              "          ON CxF.Produto = Produtos.Produto_Id " & vbCrLf & _
              " Where CxF.Cliente_Id    ='" & pcliente.Codigo & "'" & vbCrLf & _
              "   and CxF.EndCliente_Id = " & pcliente.CodigoEndereco

        Dim Banco As New AcessaBanco
        Dim ds As New DataSet
        ds = Banco.ConsultaDataSet(sql, "ClienteXFinanciamento")

        For Each row As DataRow In ds.Tables(0).Rows
            Dim CxF As New ClienteXFinanciamento(pcliente)
            CxF.CodigoFinanciador = row("Financiador_Id")
            CxF.EndFinanciador = row("EndFinanciador_Id")
            CxF.NomeFinanciador = row("NomeFinanciador")
            CxF.CodigoFinanciamento = row("Financiamento_Id")
            CxF.TipoFinanciamento = row("TipoFinanciamento")
            CxF.CodigoSafra = row("Safra")
            CxF.DataFinanciamento = row("DataFinanciamento")
            CxF.DataVencimento = row("DataVencimento")
            CxF.NumeroDeParcelas = row("NumeroDeParcelas")
            CxF.CodigoProduto = row("Produto")
            CxF.Quantidade = row("Quantidade")
            CxF.CodigoMoeda = row("Moeda")
            CxF.ValorOficial = row("ValorOficial")
            CxF.ValorMoeda = row("ValorMoeda")
            CxF.Observacao = row("Observacao")
            Me.Add(CxF)
        Next
    End Sub
#End Region

#Region "Fields"
    Private _Cliente As Cliente
#End Region

#Region "Property"
    Public Property Cliente() As Cliente
        Get
            Return _Cliente
        End Get
        Set(ByVal value As Cliente)
            _Cliente = value
        End Set
    End Property
#End Region

#Region "Methods"
    Public Sub SalvarSql(ByVal Sqls As ArrayList)
        For Each CxF As ClienteXFinanciamento In Me
            If _Cliente.IUD = "D" Or _Cliente.IUD = "I" Then CxF.IUD = _Cliente.IUD
            If CxF.IUD <> "" Then
                CxF.SalvarSql(Sqls)
            End If
        Next
    End Sub
#End Region

End Class

<Serializable()> _
Public Class ClienteXFinanciamento
    Implements IBaseEntity

#Region "Construtor"
    Public Sub New()

    End Sub

    Public Sub New(ByVal pCliente As Cliente)
        _Cliente = pCliente
    End Sub
#End Region

#Region "Fields"
    Private _IUD As String = ""
    Private _Cliente As Cliente
    Private _CodigoFinanciador As String
    Private _EndFinanciador As Integer
    Private _NomeFinanciador As String
    Private _Financiador As Cliente
    Private _CodigoFinanciamento As Integer
    Private _TipoFinanciamento As String
    Private _CodigoSafra As String
    Private _Safra As Safra
    Private _DataFinanciamento As DateTime
    Private _DataVencimento As DateTime
    Private _NumeroDeParcelas As Integer
    Private _CodigoProduto As String
    Private _NomeProduto As String
    Private _Produto As Produto
    Private _Quantidade As Decimal
    Private _CodigoMoeda As Integer
    Private _Moeda As Moeda
    Private _ValorOficial As Decimal
    Private _ValorMoeda As Decimal
    Private _Observacao As String
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

    Public Property Cliente() As Cliente
        Get
            Return _Cliente
        End Get
        Set(ByVal value As Cliente)
            _Cliente = value
        End Set
    End Property

    Public Property CodigoFinanciador() As String
        Get
            Return _CodigoFinanciador
        End Get
        Set(ByVal value As String)
            _CodigoFinanciador = value
        End Set
    End Property

    Public Property EndFinanciador() As Integer
        Get
            Return _EndFinanciador
        End Get
        Set(ByVal value As Integer)
            _EndFinanciador = value
        End Set
    End Property

    Public Property NomeFinanciador() As String
        Get
            Return _NomeFinanciador
        End Get
        Set(ByVal value As String)
            _NomeFinanciador = value
        End Set
    End Property

    Public Property Financiador() As Cliente
        Get
            If _Financiador Is Nothing And _CodigoFinanciador.Trim.Length > 0 Then _Financiador = New Cliente(_CodigoFinanciador, _EndFinanciador)
            Return _Financiador
        End Get
        Set(ByVal value As Cliente)
            _Financiador = value
        End Set
    End Property

    Public Property CodigoFinanciamento() As Integer
        Get
            Return _CodigoFinanciamento
        End Get
        Set(ByVal value As Integer)
            _CodigoFinanciamento = value
        End Set
    End Property

    Public Property TipoFinanciamento() As String
        Get
            Return _TipoFinanciamento
        End Get
        Set(ByVal value As String)
            _TipoFinanciamento = value
        End Set
    End Property

    Public Property CodigoSafra() As String
        Get
            Return _CodigoSafra
        End Get
        Set(ByVal value As String)
            _CodigoSafra = value
        End Set
    End Property

    Public Property Safra() As Safra
        Get
            If _Safra Is Nothing And _CodigoSafra.Trim.Length > 0 Then _Safra = New Safra(_CodigoSafra)
            Return _Safra
        End Get
        Set(ByVal value As Safra)
            _Safra = value
        End Set
    End Property

    Public Property DataFinanciamento() As DateTime
        Get
            Return _DataFinanciamento
        End Get
        Set(ByVal value As DateTime)
            _DataFinanciamento = value
        End Set
    End Property

    Public Property DataVencimento() As DateTime
        Get
            Return _DataVencimento
        End Get
        Set(ByVal value As DateTime)
            _DataVencimento = value
        End Set
    End Property

    Public Property NumeroDeParcelas() As Integer
        Get
            Return _NumeroDeParcelas
        End Get
        Set(ByVal value As Integer)
            _NumeroDeParcelas = value
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

    Public Property NomeProduto() As String
        Get
            Return _NomeProduto
        End Get
        Set(ByVal value As String)
            _NomeProduto = value
        End Set
    End Property

    Public Property Produto() As Produto
        Get
            If _Produto Is Nothing And _CodigoProduto.Trim.Length > 0 Then _Produto = New Produto(_CodigoProduto)
            Return _Produto
        End Get
        Set(ByVal value As Produto)
            _Produto = value
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

    Public Property CodigoMoeda() As Integer
        Get
            Return _CodigoMoeda
        End Get
        Set(ByVal value As Integer)
            _CodigoMoeda = value
        End Set
    End Property

    Public Property Moeda() As Moeda
        Get
            If _Moeda Is Nothing And _CodigoMoeda > 0 Then _Moeda = New Moeda(_CodigoMoeda)
            Return _Moeda
        End Get
        Set(ByVal value As Moeda)
            _Moeda = value
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

    Public Property Observacao() As String
        Get
            Return _Observacao
        End Get
        Set(ByVal value As String)
            _Observacao = value
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

    Public Sub SalvarSql(ByVal Sqls As ArrayList)
        Dim sql As String = ""

        Select Case _IUD
            Case "I"
                sql = "Insert into ClienteXFinanciamento(Cliente_Id, EndCliente_Id, Financiador_Id, EndFinanciador_Id, Financiamento_Id, TipoFinanciamento, Safra, DataFinanciamento, DataVencimento, NumeroDeParcelas, Produto, Quantidade, Moeda, ValorOficial, ValorMoeda, Observacao)" & vbCrLf & _
                      " values('" & _Cliente.Codigo & "'," & _Cliente.CodigoEndereco & ",'" & _CodigoFinanciador & "'," & _EndFinanciador & "," & _CodigoFinanciamento & ",'" & _TipoFinanciamento & "','" & _CodigoSafra & "','" & _DataFinanciamento.ToString("yyyy-MM-dd") & "','" & _DataVencimento.ToString("yyyy-MM-dd") & "'," & _NumeroDeParcelas & ",'" & _CodigoProduto & "'," & Str(_Quantidade) & "," & _CodigoMoeda & "," & Str(_ValorOficial) & "," & Str(_ValorMoeda) & ",'" & _Observacao & "')"
                Sqls.Add(sql)
            Case "U"
                sql = "Update ClienteXFinanciamento set " & vbCrLf & _
                      "    TipoFinanciamento ='" & _TipoFinanciamento & "'" & vbCrLf & _
                      "   ,Safra             ='" & _CodigoSafra & "'" & vbCrLf & _
                      "   ,DataFinanciamento ='" & _DataFinanciamento.ToString("yyyy-MM-dd") & "'" & vbCrLf & _
                      "   ,DataVencimento    ='" & _DataVencimento.ToString("yyyy-MM-dd") & "'" & vbCrLf & _
                      "   ,NumeroDeParcelas  = " & _NumeroDeParcelas & vbCrLf & _
                      "   ,Produto           ='" & _CodigoProduto & "'" & vbCrLf & _
                      "   ,Quantidade        = " & Str(_Quantidade) & vbCrLf & _
                      "   ,Moeda             = " & _CodigoMoeda & vbCrLf & _
                      "   ,ValorOficial      = " & Str(_ValorOficial) & vbCrLf & _
                      "   ,ValorMoeda        = " & Str(_ValorMoeda) & vbCrLf & _
                      "   ,Observacao        ='" & _Observacao & "'" & vbCrLf & _
                      " Where Cliente_Id        ='" & _Cliente.Codigo & "'" & vbCrLf & _
                      "   and EndCliente_Id     = " & _Cliente.CodigoEndereco & vbCrLf & _
                      "   and Financiador_Id    ='" & _CodigoFinanciador & "'" & vbCrLf & _
                      "   and EndFinanciador_Id = " & _EndFinanciador & vbCrLf & _
                      "   and Financiamento_Id  = " & _CodigoFinanciamento
                Sqls.Add(sql)
            Case "D"
                sql = "Delete ClienteXFinanciamento " & vbCrLf & _
                      " Where Cliente_Id        ='" & _Cliente.Codigo & "'" & vbCrLf & _
                      "   and EndCliente_Id     = " & _Cliente.CodigoEndereco & vbCrLf & _
                      "   and Financiador_Id    ='" & _CodigoFinanciador & "'" & vbCrLf & _
                      "   and EndFinanciador_Id = " & _EndFinanciador & vbCrLf & _
                      "   and Financiamento_Id  = " & _CodigoFinanciamento
                Sqls.Add(sql)
        End Select
    End Sub
#End Region

End Class