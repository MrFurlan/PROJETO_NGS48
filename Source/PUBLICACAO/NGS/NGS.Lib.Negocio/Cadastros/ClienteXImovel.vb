Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Collections.Generic
Imports NGS.Lib.Uteis

<Serializable()> _
Public Class ListClienteXImovel
    Inherits List(Of ClienteXImovel)

#Region "Construtor"
    Public Sub New()

    End Sub

    Public Sub New(ByVal pcliente As Cliente)
        _Cliente = pcliente
        Dim sql As String
        sql = "Select Cliente_Id, EndCliente_Id, Imovel_Id, Descricao, Onerado, Cidade, Estado, AreaTotal, AreaConstruida, Unidade, NumeroRegistro, Cartorio, ValorOficial, ValorMoeda, DataAvaliacao " & vbCrLf & _
              "  from ClienteXImovel " & vbCrLf & _
              " Where Cliente_Id  ='" & pcliente.Codigo & "'" & vbCrLf & _
              "   and EndCliente_Id = " & pcliente.CodigoEndereco

        Dim Banco As New AcessaBanco
        Dim ds As New DataSet
        ds = Banco.ConsultaDataSet(sql, "ClienteXImovel")

        For Each row As DataRow In ds.Tables(0).Rows
            Dim CxI As New ClienteXImovel(pcliente)
            CxI.CodigoImovel = row("Imovel_Id")
            CxI.Descricao = row("Descricao")
            CxI.Onerado = row("Onerado")
            CxI.CodigoCidade = row("Cidade")
            CxI.CodigoEstado = row("Estado")
            CxI.AreaTotal = row("AreaTotal")
            CxI.AreaConstruida = row("AreaConstruida")
            CxI.CodigoUnidadeDeMedida = row("Unidade")
            CxI.NumeroRegistro = row("NumeroRegistro")
            CxI.Cartorio = row("Cartorio")
            CxI.DataAvaliacao = row("DataAvaliacao")
            CxI.ValorOficial = row("ValorOficial")
            CxI.ValorMoeda = row("ValorMoeda")
            Me.Add(CxI)
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
        For Each CxC As ClienteXImovel In Me
            If _Cliente.IUD = "D" Or _Cliente.IUD = "I" Then CxC.IUD = _Cliente.IUD
            If CxC.IUD <> "" Then
                CxC.SalvarSql(Sqls)
            End If
        Next
    End Sub
#End Region

End Class

<Serializable()> _
Public Class ClienteXImovel
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
    Private _CodigoImovel As Integer
    Private _Descricao As String
    Private _Onerado As Boolean
    Private _CodigoCidade As String
    Private _CodigoEstado As String
    Private _Estado As Estado
    Private _AreaTotal As Decimal
    Private _AreaConstruida As Decimal
    Private _CodigoUnidadeDeMedida As String
    Private _UnidadeDeMedida As UnidadeDeMedida
    Private _NumeroRegistro As String
    Private _Cartorio As String
    Private _DataAvaliacao As DateTime
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

    Public Property CodigoImovel() As Integer
        Get
            Return _CodigoImovel
        End Get
        Set(ByVal value As Integer)
            _CodigoImovel = value
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

    Public Property Onerado() As Boolean
        Get
            Return _Onerado
        End Get
        Set(ByVal value As Boolean)
            _Onerado = value
        End Set
    End Property

    Public Property CodigoCidade() As String
        Get
            Return _CodigoCidade
        End Get
        Set(ByVal value As String)
            _CodigoCidade = value
        End Set
    End Property

    Public Property CodigoEstado() As String
        Get
            Return _CodigoEstado
        End Get
        Set(ByVal value As String)
            _CodigoEstado = value
            _Estado = Nothing
        End Set
    End Property

    Public ReadOnly Property Estado() As Estado
        Get
            If _Estado Is Nothing And _CodigoEstado.Trim.Length > 0 Then _Estado = New Estado(_CodigoEstado)
            Return _Estado
        End Get
    End Property

    Public Property AreaTotal() As Decimal
        Get
            Return _AreaTotal
        End Get
        Set(ByVal value As Decimal)
            _AreaTotal = value
        End Set
    End Property

    Public Property AreaConstruida() As Decimal
        Get
            Return _AreaConstruida
        End Get
        Set(ByVal value As Decimal)
            _AreaConstruida = value
        End Set
    End Property

    Public Property CodigoUnidadeDeMedida() As String
        Get
            Return _CodigoUnidadeDeMedida
        End Get
        Set(ByVal value As String)
            _CodigoUnidadeDeMedida = value
            _UnidadeDeMedida = Nothing
        End Set
    End Property

    Public Property UnidadeDeMedida() As UnidadeDeMedida
        Get
            If _UnidadeDeMedida Is Nothing And _CodigoUnidadeDeMedida.Trim.Length > 0 Then _UnidadeDeMedida = New UnidadeDeMedida(_CodigoUnidadeDeMedida)
            Return _UnidadeDeMedida
        End Get
        Set(ByVal value As UnidadeDeMedida)
            _UnidadeDeMedida = value
        End Set
    End Property

    Public Property NumeroRegistro() As String
        Get
            Return _NumeroRegistro
        End Get
        Set(ByVal value As String)
            _NumeroRegistro = value
        End Set
    End Property

    Public Property Cartorio() As String
        Get
            Return _Cartorio
        End Get
        Set(ByVal value As String)
            _Cartorio = value
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
                sql = "Insert into ClienteXImovel(Cliente_Id, EndCliente_Id,  Descricao, Imovel_id, Onerado, Cidade, Estado, AreaTotal, AreaConstruida, Unidade, NumeroRegistro, Cartorio, ValorOficial, ValorMoeda, DataAvaliacao, Observacao)" & vbCrLf & _
                      " values('" & _Cliente.Codigo & "'," & _Cliente.CodigoEndereco & ",'" & _Descricao & "'," & _CodigoImovel & "," & CByte(_Onerado) & ",'" & _CodigoCidade & "','" & _CodigoEstado & "'," & Str(_AreaTotal) & "," & Str(AreaConstruida) & ",'" & _CodigoUnidadeDeMedida & "','" & _NumeroRegistro & "','" & _Cartorio & "'," & Str(_ValorOficial) & "," & Str(_ValorMoeda) & ",'" & _DataAvaliacao.ToString("yyyy-MM-dd") & "', '" & _Observacao & "')"
                Sqls.Add(sql)
            Case "U"
                sql = "Update ClienteXImovel set " & vbCrLf & _
                      "    Descricao      ='" & _Descricao & "'" & vbCrLf & _
                      "   ,Cidade         ='" & _CodigoCidade & "'" & vbCrLf & _
                      "   ,Estado         ='" & _CodigoEstado & "'" & vbCrLf & _
                      "   ,AreaTotal      = " & Str(_AreaTotal) & vbCrLf & _
                      "   ,AreaConstruida = " & Str(_AreaConstruida) & vbCrLf & _
                      "   ,Unidade        ='" & _CodigoUnidadeDeMedida & "'" & vbCrLf & _
                      "   ,NumeroRegistro ='" & _NumeroRegistro & "'" & vbCrLf & _
                      "   ,Cartorio       ='" & _Cartorio & "'" & vbCrLf & _
                      "   ,ValorOficial   = " & Str(_ValorOficial) & vbCrLf & _
                      "   ,ValorMoeda     = " & Str(_ValorMoeda) & vbCrLf & _
                      "   ,DataAvaliacao  ='" & _DataAvaliacao.ToString("yyyy-MM-dd") & "'" & vbCrLf & _
                      "   ,Observacao     ='" & _Observacao & "'" & vbCrLf & _
                      " Where Cliente_Id    ='" & _Cliente.Codigo & "'" & vbCrLf & _
                      "   and EndCliente_Id = " & _Cliente.CodigoEndereco & vbCrLf & _
                      "   and Imovel_Id     = " & _CodigoImovel
                Sqls.Add(sql)
            Case "D"
                sql = "Delete ClienteXImovel " & vbCrLf & _
                      " Where Cliente_Id    ='" & _Cliente.Codigo & "'" & vbCrLf & _
                      "   and EndCliente_Id = " & _Cliente.CodigoEndereco & vbCrLf & _
                      "   and Imovel_Id     = " & _CodigoImovel
                Sqls.Add(sql)
        End Select
    End Sub
#End Region

End Class