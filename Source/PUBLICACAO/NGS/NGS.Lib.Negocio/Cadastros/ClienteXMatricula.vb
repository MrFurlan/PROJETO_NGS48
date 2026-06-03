Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Collections.Generic
Imports NGS.Lib.Uteis

<Serializable()> _
Public Class ListClienteXMatricula
    Inherits List(Of ClienteXMatricula)

#Region "Construtor"
    Public Sub New()

    End Sub

    Public Sub New(ByVal pcliente As Cliente)
        _Cliente = pcliente
        Dim sql As String
        sql = "SELECT Cliente_Id, EndCliente_Id, Matricula_Id, Area, DataAvaliacao, ValorOficial, isnull(Registro,'') as Registro, isnull(Livro,'') as Livro, isnull(Folha,'') as Folha, isnull(Municipio,'') as Municipio, isnull(Estado,'') as Estado" & vbCrLf & _
              "  FROM ClienteXMatricula " & vbCrLf & _
              " Where Cliente_Id    ='" & pcliente.Codigo & "'" & vbCrLf & _
              "   and EndCliente_Id = " & pcliente.CodigoEndereco

        Dim Banco As New AcessaBanco
        Dim ds As New DataSet
        ds = Banco.ConsultaDataSet(sql, "ClienteXMatricula")

        For Each row As DataRow In ds.Tables(0).Rows
            Dim CxM As New ClienteXMatricula(pcliente)
            CxM.CodigoMatricula = row("Matricula_Id")
            CxM.Area = row("Area")
            CxM.DataAvaliacao = row("DataAvaliacao")
            CxM.ValorOficial = row("ValorOficial")
            CxM.Registro = row("Registro")
            CxM.Livro = row("Livro")
            CxM.Folha = row("Folha")
            CxM.CodigoMunicipio = row("Municipio")
            CxM.CodigoEstado = row("Estado")
            Me.Add(CxM)
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
        For Each CxM As ClienteXMatricula In Me
            If _Cliente.IUD = "D" Or _Cliente.IUD = "I" Then CxM.IUD = _Cliente.IUD
            If CxM.IUD <> "" Then
                CxM.SalvarSql(Sqls)
            End If
        Next
    End Sub
#End Region

End Class

<Serializable()> _
Public Class ClienteXMatricula
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
    Private _CodigoMatricula As String
    Private _Area As Decimal
    Private _DataAvaliacao As DateTime
    Private _ValorOficial As Decimal
    Private _Registro As String
    Private _Livro As String
    Private _Folha As String
    Private _CodigoMunicipio As String
    Private _CodigoEstado As String
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

    Public Property CodigoMatricula() As String
        Get
            Return _CodigoMatricula
        End Get
        Set(ByVal value As String)
            _CodigoMatricula = value
        End Set
    End Property

    Public Property Area() As Decimal
        Get
            Return _Area
        End Get
        Set(ByVal value As Decimal)
            _Area = value
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

    Public Property Registro() As String
        Get
            Return _Registro
        End Get
        Set(ByVal value As String)
            _Registro = value
        End Set
    End Property

    Public Property Livro() As String
        Get
            Return _Livro
        End Get
        Set(ByVal value As String)
            _Livro = value
        End Set
    End Property

    Public Property Folha() As String
        Get
            Return _Folha
        End Get
        Set(ByVal value As String)
            _Folha = value
        End Set
    End Property

    Public Property CodigoMunicipio() As String
        Get
            Return _CodigoMunicipio
        End Get
        Set(ByVal value As String)
            _CodigoMunicipio = value
        End Set
    End Property

    Public Property CodigoEstado() As String
        Get
            Return _CodigoEstado
        End Get
        Set(ByVal value As String)
            _CodigoEstado = value
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
                sql = "Insert into ClienteXMatricula( Cliente_Id, EndCliente_Id, Matricula_Id, Area, DataAvaliacao, ValorOficial, Registro, Livro, Folha, Municipio, Estado)" & vbCrLf & _
                      " values('" & _Cliente.Codigo & "'," & _Cliente.CodigoEndereco & ",'" & _CodigoMatricula & "'," & Str(_Area) & ", '" & _DataAvaliacao.ToString("yyyy-MM-dd") & "', " & Str(_ValorOficial) & ",'" & _Registro & "','" & _Livro & "','" & _Folha & "','" & _CodigoMunicipio & "','" & _CodigoEstado & "')"
                Sqls.Add(sql)
            Case "U"
                sql = "Update ClienteXMatricula set " & vbCrLf & _
                      "    Area           = " & Str(_Area) & vbCrLf & _
                      "    ,DataAvaliacao = '" & _DataAvaliacao.ToString("yyyy-MM-dd") & "'" & vbCrLf & _
                      "    ,ValorOficial  = " & Str(_ValorOficial) & vbCrLf & _
                      "    ,Registro      ='" & _Registro & "'" & vbCrLf & _
                      "    ,Livro         ='" & _Livro & "'" & vbCrLf & _
                      "    ,Folha         ='" & _Folha & "'" & vbCrLf & _
                      "    ,Municipio     ='" & _CodigoMunicipio & "'" & vbCrLf & _
                      "    ,Estado        ='" & _CodigoEstado & "'" & vbCrLf & _
                      " Where Cliente_Id     = '" & _Cliente.Codigo & "'" & vbCrLf & _
                      "   and EndCliente_Id  = " & _Cliente.CodigoEndereco & vbCrLf & _
                      "   and Matricula_Id   = '" & _CodigoMatricula & "'" & vbCrLf
                Sqls.Add(sql)
            Case "D"
                sql = "Delete ClienteXMatricula " & vbCrLf & _
                      " Where Cliente_Id     = '" & _Cliente.Codigo & "'" & vbCrLf & _
                      "   and EndCliente_Id  = " & _Cliente.CodigoEndereco & vbCrLf & _
                      "   and Matricula_Id   = '" & _CodigoMatricula & "'" & vbCrLf
                Sqls.Add(sql)
        End Select
    End Sub
#End Region

End Class