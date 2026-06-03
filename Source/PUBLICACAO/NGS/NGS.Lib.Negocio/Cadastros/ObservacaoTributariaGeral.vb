Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Collections.Generic
Imports NGS.Lib.Uteis

<Serializable()> _
Public Class ListObservacaoTributariaGeral
    Inherits List(Of ObservacaoTributariaGeral)

    Public Sub New()

    End Sub

    Public Sub New(ByVal estado As String, ByVal encargo As String)
        Dim sql As String
        sql = "SELECT codigo_id, Descricao, Estado, IsNull(Encargo, 'PRODUTO') as Encargo" & vbCrLf & _
              "  FROM ObservacoesTributarias " & vbCrLf & _
              " Where 1 = 1 " & vbCrLf

        If Not String.IsNullOrWhiteSpace(estado) Then
            sql &= "   And Estado = '" & estado & "'" & vbCrLf
        End If

        If Not String.IsNullOrWhiteSpace(encargo) Then
            sql &= "   And Encargo = '" & encargo & "'" & vbCrLf
        End If

        sql &= " Order By codigo_id " & vbCrLf

        Dim Banco As New AcessaBanco
        Dim ds As DataSet

        ds = Banco.ConsultaDataSet(sql, "ObservacoesTributarias")

        For Each row As DataRow In ds.Tables(0).Rows
            Dim Obs As New ObservacaoTributariaGeral
            Obs.Codigo = row("codigo_id")
            Obs.Descricao = row("Descricao")
            Obs.CodigoEstado = row("Estado")
            Obs.CodigoEncargo = row("Encargo")
            Me.Add(Obs)
        Next
    End Sub

    Public Sub New(ByVal pWhere As String)
        Dim sql As String
        sql = "SELECT codigo_id, Descricao, Estado, Encargo " & vbCrLf & _
              "  FROM ObservacoesTributarias " & vbCrLf & _
              " Where " & pWhere & "" & vbCrLf & _
              " Order By codigo_id " & vbCrLf

        Dim Banco As New AcessaBanco
        Dim ds As DataSet

        ds = Banco.ConsultaDataSet(sql, "ObservacoesTributarias")

        For Each row As DataRow In ds.Tables(0).Rows
            Dim Obs As New ObservacaoTributariaGeral
            Obs.Codigo = row("codigo_id")
            Obs.Descricao = row("Descricao")
            Obs.CodigoEstado = row("Estado")
            Obs.CodigoEncargo = row("Encargo")
            Me.Add(Obs)
        Next

    End Sub
End Class

<Serializable()> _
Public Class ObservacaoTributariaGeral
    Implements IBaseEntity

#Region "Construtor"
    Public Sub New()

    End Sub

    Public Sub New(ByVal pCodigo As Integer)
        Dim Banco As New AcessaBanco
        Dim ds As DataSet
        Dim sql As String
        sql = "Select codigo_id, Descricao, Estado, isnull(Encargo,'') as Encargo " & vbCrLf & _
              "  from ObservacoesTributarias " & vbCrLf & _
              " where codigo_id = " & pCodigo

        ds = Banco.ConsultaDataSet(sql, "ObservacoesTributarias")
        If ds.Tables(0).Rows.Count > 0 Then
            Me.Codigo = ds.Tables(0).Rows(0)("codigo_id")
            Me.Descricao = ds.Tables(0).Rows(0)("Descricao")
            Me.CodigoEstado = ds.Tables(0).Rows(0)("Estado")
            Me.CodigoEncargo = ds.Tables(0).Rows(0)("Encargo")
        End If
    End Sub
#End Region

#Region "Fields"
    Private _IUD As String = ""
    Private _Codigo As Integer = 0
    Private _Descricao As String = ""
    Private _CodigoEstado As String = ""
    Private _Estado As Estado
    Private _CodigoEncargo As String = ""
    Private _Encargo As Encargo
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

    Public Property Codigo() As Integer
        Get
            Return _Codigo
        End Get
        Set(ByVal value As Integer)
            _Codigo = value
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

    Public ReadOnly Property DescricaoObj() As String
        Get
            If _Descricao.Length > 200 Then
                Return Format(_Codigo, "000") & " - " & _Descricao.Substring(0, 199) & "..."
            Else
                Return Format(_Codigo, "000") & " - " & _Descricao
            End If
        End Get
    End Property

    Public Property CodigoEstado() As String
        Get
            Return _CodigoEstado
        End Get
        Set(ByVal value As String)
            _CodigoEstado = value
        End Set
    End Property

    Public Property Estado() As Estado
        Get
            If _Estado Is Nothing And CodigoEstado.Length > 0 Then _Estado = New Estado(CodigoEstado)
            Return _Estado

        End Get
        Set(ByVal value As Estado)
            _Estado = value
        End Set
    End Property

    Public Property CodigoEncargo() As String
        Get
            Return _CodigoEncargo
        End Get
        Set(ByVal value As String)
            _CodigoEncargo = value
        End Set
    End Property

    Public Property Encargo() As Encargo
        Get
            If _Encargo Is Nothing And CodigoEncargo.Length > 0 Then _Encargo = New Encargo(CodigoEncargo)
            Return _Encargo
        End Get
        Set(ByVal value As Encargo)
            _Encargo = value
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
                Sql = " INSERT INTO ObservacoesTributarias(codigo_id, Descricao, Estado, Encargo) " & vbCrLf & _
                      " VALUES (" & Me.Codigo & ",'" & Me.Descricao & "','" & Me.CodigoEstado & "','" & Me.CodigoEncargo & "')"
                Sqls.Add(Sql)
            Case "U"
                Sql = " UPDATE ObservacoesTributarias SET" & vbCrLf & _
                      "     Descricao ='" & _Descricao & "'" & vbCrLf & _
                      "    ,Estado    ='" & _CodigoEstado & "'" & vbCrLf & _
                      "    ,Encargo   ='" & _CodigoEncargo & "'" & vbCrLf & _
                      "  WHERE codigo_id = " & _Codigo & vbCrLf
                Sqls.Add(Sql)
            Case "D"
                Sql = " DELETE ObservacoesTributarias" & vbCrLf & _
                      "  WHERE codigo_id = " & _Codigo & vbCrLf
                Sqls.Add(Sql)
        End Select
    End Sub

    Public Sub buscarSequencia()
        Dim Banco As New AcessaBanco
        Dim ds As DataSet
        Dim sql As String
        sql = "Select max(codigo_id + 1) AS Codigo " & vbCrLf & _
              "  from ObservacoesTributarias "

        ds = Banco.ConsultaDataSet(sql, "ObservacoesTributarias")
        If ds.Tables(0).Rows.Count > 0 Then
            Me.Codigo = ds.Tables(0).Rows(0)("Codigo")
        End If
    End Sub
#End Region

End Class
