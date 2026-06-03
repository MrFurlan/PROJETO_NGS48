Imports Microsoft.VisualBasic
Imports System.Data
Imports NGS.Lib.Uteis

<Serializable()> _
Public Class ListArt
    Inherits List(Of ART)

    Public Sub New(Optional ByVal CarregarDados As Boolean = False, Optional ByVal Empresa As String = "", Optional ByVal EndEmpresa As String = "", Optional ByVal Cliente As String = "", Optional ByVal EndCliente As String = "", Optional ByVal Status As Boolean = False)
        If CarregarDados Then
            Dim strWhereAnd As String = "WHERE"
            Dim sql As String = "Select  ART_Id, Empresa, EndEmpresa, Agronomo, EndAgronomo, ARTInicial, ARTFinal, ARTValidade, Status, UsuarioInclusao, UsuarioInclusaoData From ART "

            If Empresa.Length > 0 Then
                sql &= strWhereAnd & " Empresa = '" & Empresa & "' AND EndEmpresa = " & EndEmpresa & ""
                strWhereAnd = "AND"
            End If

            If Cliente.Length > 0 Then
                sql &= strWhereAnd & " Agronomo = '" & Cliente & "' AND EndAgronomo = " & EndCliente & ""
            End If

            If Status = True Then
                sql &= strWhereAnd & " Status = '1'"
            End If

            Dim Banco As New AcessaBanco
            Dim ds As DataSet

            ds = Banco.ConsultaDataSet(sql, "ART")

            For Each row As DataRow In ds.Tables(0).Rows
                Dim Ar As New ART
                Ar.CodigoArt = row("ART_Id")
                Ar.CodigoEmpresa = row("Empresa")
                Ar.EndEmpresa = row("EndEmpresa")
                Ar.CodigoAgronomo = row("Agronomo")
                Ar.EndAgronomo = row("EndAgronomo")
                Ar.ARTInicial = row("ARTInicial")
                Ar.ARTFinal = row("ARTFinal")
                Ar.ARTValidade = row("ARTValidade")
                Ar.Status = row("Status")
                Ar.UsuarioInclusao = row("UsuarioInclusao")
                Ar.UsuarioInclusaoData = row("UsuarioInclusaoData")
                Me.Add(Ar)
            Next
        End If
    End Sub

End Class

<Serializable()> _
Public Class ART

#Region "Contrutor"
    Public Sub New()

    End Sub

    Public Sub New(ByVal pCodigoArt As Integer)
        Dim Banco As New AcessaBanco
        Dim ds As New DataSet

        ds = Banco.ConsultaDataSet("Select  ART_Id, Empresa, EndEmpresa, Agronomo, EndAgronomo, ARTInicial, ARTFinal, ARTValidade, Status, UsuarioInclusao, UsuarioInclusaoData From ART Where ART_id = " & pCodigoArt, "ART")

        If ds.Tables(0).Rows.Count > 0 Then
            Dim row As DataRow = ds.Tables(0).Rows(0)
            _CodigoArt = row("ART_Id")
            _CodigoEmpresa = row("Empresa")
            _EndEmpresa = row("EndEmpresa")
            _CodigoAgronomo = row("Agronomo")
            _EndAgronomo = row("EndAgronomo")
            _ARTInicial = row("ARTInicial")
            _ARTFinal = row("ARTFinal")
            _ARTValidade = row("ARTValidade")
            _Status = row("Status")
            _UsuarioInclusao = row("UsuarioInclusao")
            _UsuarioInclusaoData = row("UsuarioInclusaoData")
        End If
    End Sub
#End Region

#Region "Fieds"
    Private _IUD As String
    Private _CodigoArt As String
    Private _Empresa As Cliente
    Private _CodigoEmpresa As String
    Private _EndEmpresa As Integer
    Private _Agronomo As Cliente
    Private _CodigoAgronomo As String
    Private _EndAgronomo As Integer
    Private _ARTInicial As Integer
    Private _ARTFinal As Integer
    Private _ARTValidade As DateTime
    Private _Status As Boolean
    Private _UsuarioInclusao As String
    Private _UsuarioInclusaoData As DateTime
    Private _ArtAtual As Integer
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

    Public Property CodigoArt() As String
        Get
            Return _CodigoArt
        End Get
        Set(ByVal value As String)
            _CodigoArt = value
        End Set
    End Property

    Public ReadOnly Property Empresa() As Cliente
        Get
            If _Empresa Is Nothing And _CodigoEmpresa.Length > 0 Then _Empresa = New Cliente(_CodigoEmpresa, _EndEmpresa)
            Return _Empresa
        End Get
    End Property

    Public Property CodigoEmpresa() As String
        Get
            Return _CodigoEmpresa
        End Get
        Set(ByVal value As String)
            _CodigoEmpresa = value
            _Empresa = Nothing
        End Set
    End Property

    Public Property EndEmpresa() As Integer
        Get
            Return _EndEmpresa
        End Get
        Set(ByVal value As Integer)
            _EndEmpresa = value
            _Empresa = Nothing
        End Set
    End Property

    Public ReadOnly Property Agronomo() As Cliente
        Get
            If _Agronomo Is Nothing And _CodigoAgronomo.Length > 0 Then _Agronomo = New Cliente(_CodigoAgronomo, _EndAgronomo)
            Return _Agronomo
        End Get
    End Property

    Public Property CodigoAgronomo() As String
        Get
            Return _CodigoAgronomo
        End Get
        Set(ByVal value As String)
            _CodigoAgronomo = value
            _Agronomo = Nothing
        End Set
    End Property

    Public Property EndAgronomo() As Integer
        Get
            Return _EndAgronomo
        End Get
        Set(ByVal value As Integer)
            _EndAgronomo = value
        End Set
    End Property

    Public ReadOnly Property NomeAgronomo() As String
        Get
            If _Agronomo Is Nothing And _CodigoAgronomo.Length > 0 Then _Agronomo = New Cliente(_CodigoAgronomo, _EndAgronomo)
            Return _Agronomo.Nome
        End Get
    End Property

    Public Property ARTInicial() As Integer
        Get
            Return _ARTInicial
        End Get
        Set(ByVal value As Integer)
            _ARTInicial = value
        End Set
    End Property

    Public Property ARTFinal() As Integer
        Get
            Return _ARTFinal
        End Get
        Set(ByVal value As Integer)
            _ARTFinal = value
        End Set
    End Property

    Public Property ARTValidade() As DateTime
        Get
            Return _ARTValidade
        End Get
        Set(ByVal value As DateTime)
            _ARTValidade = value
        End Set
    End Property

    Public Property Status() As Boolean
        Get
            Return _Status
        End Get
        Set(ByVal value As Boolean)
            _Status = value
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

    Public ReadOnly Property ARTAtual() As Integer
        Get
            If _ArtAtual = 0 Then
                Dim Sql As String
                Sql = "select isnull(max(ArtReceita),0) as NumART" & _
                      "  from receita" & _
                      " Where ART = " & _CodigoArt
                Dim ds As DataSet
                Dim Banco As New AcessaBanco
                ds = Banco.ConsultaDataSet(Sql, "Art")
                _ArtAtual = ds.Tables(0).Rows(0)("NumART")
                'If _ArtAtual = 0 Then
                '    _ArtAtual = _ARTInicial
                'End If
            End If
            Return _ArtAtual
        End Get
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
                Sql = " INSERT INTO ART( ART_Id, Empresa, EndEmpresa, Agronomo, EndAgronomo, ARTInicial, ARTFinal, ARTValidade, Status, UsuarioInclusao, UsuarioInclusaoData) " & vbCrLf & _
                      " VALUES (" & _CodigoArt & ",'" & _CodigoEmpresa & "'," & _EndEmpresa & ",'" & _CodigoAgronomo & "'," & _EndAgronomo & "," & _ARTInicial & "," & _ARTFinal & ",'" & _ARTValidade.ToSqlDate() & "'," & IIf(_Status, "1", "0") & ",'" & _UsuarioInclusao & "','" & _UsuarioInclusaoData.ToSqlDate() & "')"
                Sqls.Add(Sql)
            Case "U"
                Sql = " UPDATE ART SET" & vbCrLf & _
                      "   Empresa      ='" & _CodigoEmpresa & "'" & vbCrLf & _
                      "  ,EndEmpresa   = " & _EndEmpresa & vbCrLf & _
                      "  ,Agronomo     ='" & _CodigoAgronomo & "'" & vbCrLf & _
                      "  ,EndAgronomo  = " & _EndAgronomo & vbCrLf & _
                      "  ,ARTInicial   = " & _ARTInicial & vbCrLf & _
                      "  ,ARTFinal     = " & _ARTFinal & vbCrLf & _
                      "  ,ARTValidade  ='" & _ARTValidade.ToSqlDate() & "'" & vbCrLf & _
                      "  ,Status       = " & IIf(_Status, "1", "0") & vbCrLf & _
                      "  ,UsuarioInclusao     = '" & _UsuarioInclusao & "'" & vbCrLf & _
                      "  ,UsuarioInclusaoData ='" & _UsuarioInclusaoData.ToSqlDate() & "'" & vbCrLf & _
                      "  WHERE ART_Id =" & _CodigoArt
                Sqls.Add(Sql)
            Case "D"
                Sql = " DELETE ART" & vbCrLf & _
                      "  WHERE ART_Id =" & _CodigoArt
                Sqls.Add(Sql)
        End Select
    End Sub
#End Region

End Class