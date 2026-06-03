Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Collections.Generic
Imports NGS.Lib.Uteis

<Serializable()> _
Public Class ListClienteXVeiculo
    Inherits List(Of ClienteXVeiculo)

#Region "Construtor"
    Public Sub New(ByVal pCliente As Cliente)
        _Cliente = pCliente

        Dim Banco As New AcessaBanco()

        Dim sql As String
        sql = " Select Cliente_id, EndCliente_id, Placa_Id, TipoDeVeiculo, TpV.Descricao AS DescricaoTipoVeiculo, Ano, MarcaModelo, Fabricante, ValorOficial, ValorMoeda, DataAvaliacao, isnull(Onerado,0) as Onerado " & vbCrLf & _
              "   FROM ClienteXVeiculo " & vbCrLf & _
              "  INNER JOIN TiposDeVeiculos TpV " & vbCrLf & _
              "     ON TipoDeVeiculo = TpV.Codigo_Id " & vbCrLf & _
              "  Where Cliente_Id    ='" & pCliente.Codigo & "'" & vbCrLf & _
              "    and EndCliente_Id = " & pCliente.CodigoEndereco

        Dim ds As DataSet = Banco.ConsultaDataSet(sql, "ClienteXVeiculo")

        For Each row As DataRow In ds.Tables(0).Rows
            Dim CV As New ClienteXVeiculo(pCliente)

            CV.CodigoPlaca = row("Placa_Id")
            CV.CodigoTipoDeVeiculo = row("TipoDeVeiculo")
            CV.DescricaoTipoDeVeiculo = row("DescricaoTipoVeiculo")
            CV.Ano = row("Ano")
            CV.MarcaModelo = row("MarcaModelo")
            CV.Fabricante = row("Fabricante")
            CV.ValorOficial = row("ValorOficial")
            CV.ValorMoeda = row("ValorMoeda")
            CV.DataAvaliacao = row("DataAvaliacao")
            CV.Onerado = row("Onerado")
            Me.Add(CV)
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
        For Each CxV As ClienteXVeiculo In Me
            If _Cliente.IUD = "D" Or _Cliente.IUD = "I" Then CxV.IUD = _Cliente.IUD
            If CxV.IUD <> "" Then
                CxV.SalvarSql(Sqls)
            End If
        Next
    End Sub
#End Region

End Class

<Serializable()> _
Public Class ClienteXVeiculo
    Implements IBaseEntity

#Region "Construtor"
    Public Sub New()

    End Sub

    Public Sub New(ByVal Cliente As Cliente)
        _Cliente = Cliente
    End Sub
#End Region

#Region "Fields"
    Private _IUD As String = ""
    Private _Cliente As Cliente
    Private _CodigoPlaca As String
    Private _CodigoTipoDeVeiculo As Integer
    Private _DescricaoTipoDeVeiculo As String
    Private _TipoDeVeiculo As TipoDeVeiculo
    Private _Ano As Integer
    Private _MarcaModelo As String
    Private _Fabricante As String
    Private _ValorOficial As Decimal
    Private _ValorMoeda As Decimal
    Private _DataAvaliacao As DateTime
    Private _Onerado As Boolean

    Public Erro As Exception
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

    Public Property CodigoPlaca() As String
        Get
            Return _CodigoPlaca
        End Get
        Set(ByVal value As String)
            _CodigoPlaca = value
        End Set
    End Property

    Public Property CodigoTipoDeVeiculo() As Integer
        Get
            Return _CodigoTipoDeVeiculo
        End Get
        Set(ByVal value As Integer)
            _CodigoTipoDeVeiculo = value
        End Set
    End Property

    Public Property DescricaoTipoDeVeiculo() As String
        Get
            Return _DescricaoTipoDeVeiculo
        End Get
        Set(ByVal value As String)
            _DescricaoTipoDeVeiculo = value
        End Set
    End Property

    Public Property TipoDeVeiculo() As TipoDeVeiculo
        Get
            If _TipoDeVeiculo Is Nothing And _CodigoTipoDeVeiculo > 0 Then _TipoDeVeiculo = New TipoDeVeiculo(_CodigoTipoDeVeiculo)
            Return _TipoDeVeiculo
        End Get
        Set(ByVal value As TipoDeVeiculo)
            _TipoDeVeiculo = value
        End Set
    End Property

    Public Property Ano() As Integer
        Get
            Return _Ano
        End Get
        Set(ByVal value As Integer)
            _Ano = value
        End Set
    End Property

    Public Property MarcaModelo() As String
        Get
            Return _MarcaModelo
        End Get
        Set(ByVal value As String)
            _MarcaModelo = value
        End Set
    End Property

    Public Property Fabricante() As String
        Get
            Return _Fabricante
        End Get
        Set(ByVal value As String)
            _Fabricante = value
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

    Public Property DataAvaliacao() As DateTime
        Get
            Return _DataAvaliacao
        End Get
        Set(ByVal value As DateTime)
            _DataAvaliacao = value
        End Set
    End Property

    Public Property Onerado() As Boolean
        Get
            Return _Onerado
        End Get
        Set(ByVal value As Boolean)
            _Onerado = value
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
                sql = "INSERT INTO ClienteXVeiculo(Cliente_id, EndCliente_id, Placa_Id, TipoDeVeiculo, Ano, MarcaModelo, Fabricante, ValorOficial, ValorMoeda, DataAvaliacao, Onerado)" & vbCrLf & _
                      "VALUES ('" & _Cliente.Codigo & "'," & _Cliente.CodigoEndereco & ",'" & _CodigoPlaca & "'," & _CodigoTipoDeVeiculo & "," & _Ano & ",'" & _MarcaModelo & "','" & _Fabricante & "'," & Str(_ValorOficial) & "," & Str(_ValorMoeda) & ",'" & _DataAvaliacao.ToString("yyyy-MM-dd") & "'," & CByte(_Onerado) & ")"
                Sqls.Add(sql)
            Case "U"
                sql = "UPDATE ClienteXVeiculo SET" & _
                      "    TipoDeVeiculo = " & _CodigoTipoDeVeiculo & vbCrLf & _
                      "   ,Ano           = " & _Ano & vbCrLf & _
                      "   ,MarcaModelo   ='" & _MarcaModelo & "'" & vbCrLf & _
                      "   ,Fabricante    ='" & _Fabricante & "'" & vbCrLf & _
                      "   ,ValorOficial  = " & Str(_ValorOficial) & vbCrLf & _
                      "   ,ValorMoeda    = " & Str(_ValorMoeda) & vbCrLf & _
                      "   ,DataAvaliacao ='" & _DataAvaliacao.ToString("yyyy-MM-dd") & "'" & vbCrLf & _
                      "   ,Onerado       = " & CByte(_Onerado) & vbCrLf & _
                      " WHERE Cliente_Id    = '" & _Cliente.Codigo & "'" & vbCrLf & _
                      "   AND EndCliente_id = " & _Cliente.CodigoEndereco & vbCrLf & _
                      "   AND Placa_Id      = '" & _CodigoPlaca & "'"
                Sqls.Add(sql)
            Case "D"
                sql = "DELETE ClienteXVeiculo " & _
                      " WHERE Cliente_Id       = '" & _Cliente.Codigo & "'" & vbCrLf & _
                      "   AND EndCliente_id    = " & _Cliente.CodigoEndereco & vbCrLf & _
                      "   AND Placa_Id         = '" & _CodigoPlaca & "'"
                Sqls.Add(sql)
        End Select
    End Sub
#End Region

End Class