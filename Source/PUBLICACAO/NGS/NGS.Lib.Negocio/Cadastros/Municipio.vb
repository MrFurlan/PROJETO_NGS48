Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Collections.Generic
Imports NGS.Lib.Uteis

<Serializable()> _
Public Class ListMunicipio
    Inherits List(Of Municipio)

    Public Sub New()

    End Sub

    Public Sub New(ByVal CarregarDados As Boolean, ByVal Estado As String, Optional ByVal Municipio As String = "")
        If CarregarDados Then
            Dim Banco As New AcessaBanco
            Dim ds As DataSet
            Dim sql As String = "Select Estado_Id, Municipio_Id, Codigo_Id, EstadoIbge from Municipios "
            sql &= " Where Estado_Id = '" & Estado & "' "

            If Municipio.Length > 0 Then
                sql &= "AND Municipio_Id = '" & Municipio & "' "
            End If

            ds = Banco.ConsultaDataSet(sql, "Municipios")

            For Each row As DataRow In ds.Tables(0).Rows
                Dim Mun As New Municipio
                Mun.CodigoEstado = row("Estado_Id")
                Mun.CodigoMunicipio = row("Municipio_Id")
                Mun.CodigoIbge = row("Codigo_Id")
                Mun.EstadoIbge = row("EstadoIbge")
                Me.Add(Mun)
            Next
        End If
    End Sub

End Class

<Serializable()> _
Public Class Municipio
    Implements IBaseEntity

#Region "Construtor"
    Public Sub New()

    End Sub

    Public Sub New(ByVal pEstado As String, ByVal pMunicipio As String)
        Dim Banco As New AcessaBanco
        Dim ds As DataSet
        Dim sql As String = "SELECT Estado_Id, Municipio_Id, Codigo_Id, EstadoIbge, ISNULL(CodigoIbgeCompleto, '') CodigoIbgeCompleto " & vbCrLf & _
                            "  FROM Municipios " & vbCrLf & _
                            " WHERE Estado_Id = '" & pEstado & "'" & vbCrLf & _
                            "   AND Municipio_Id = '" & pMunicipio & "'" & vbCrLf

        ds = Banco.ConsultaDataSet(sql, "Municipios")

        If ds.Tables(0).Rows.Count > 0 Then
            Dim row As DataRow = ds.Tables(0).Rows(0)
            _CodigoEstado = row("Estado_Id")
            _CodigoMunicipio = row("Municipio_Id")
            _CodigoIbge = row("Codigo_Id")
            _EstadoIbge = row("EstadoIbge")
            _CodigoIbgeCompleto = row("CodigoIbgeCompleto")
        End If
    End Sub
#End Region

#Region "Fields"
    Private _IUD As String = ""
    Private _CodigoEstado As String = ""
    Private _Estado As Estado
    Private _CodigoMunicipio As String = ""
    Private _CodigoIbge As Integer = 0
    Private _EstadoIbge As Integer = 0
    Private _CodigoIbgeCompleto As Integer = 0
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
            If _Estado Is Nothing And _CodigoEstado.Length > 0 Then _Estado = New Estado(_CodigoEstado)
            Return _Estado
        End Get
        Set(ByVal value As Estado)
            _Estado = value
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

    Public Property CodigoIbge() As Integer
        Get
            Return _CodigoIbge
        End Get
        Set(ByVal value As Integer)
            _CodigoIbge = value
        End Set
    End Property

    Public Property EstadoIbge() As Integer
        Get
            Return _EstadoIbge
        End Get
        Set(ByVal value As Integer)
            _EstadoIbge = value
        End Set
    End Property


    Public Property CodigoIbgeCompleto() As Integer
        Get
            Return _CodigoIbgeCompleto
        End Get
        Set(ByVal value As Integer)
            _CodigoIbgeCompleto = value
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
                sql = "Insert into Municipios(Estado_Id, Municipio_Id, Codigo_Id, EstadoIbge)" & vbCrLf & _
                      " values('" & _CodigoEstado & "','" & _CodigoMunicipio & "'," & _CodigoIbge & "," & _EstadoIbge & ")"
                Sqls.Add(sql)
            Case "U"
                sql = "Update Municipios set " & vbCrLf & _
                      "    EstadoIbge = " & _EstadoIbge & vbCrLf & _
                      " Where Estado_Id    ='" & _CodigoEstado & "'" & vbCrLf & _
                      "   and Municipio_Id ='" & _CodigoMunicipio & "'" & vbCrLf & _
                      "   and Codigo_Id    = " & _CodigoIbge
                Sqls.Add(sql)
            Case "D"
                sql = "Delete Municipios " & vbCrLf & _
                      " Where Estado_Id    ='" & _CodigoEstado & "'" & vbCrLf & _
                      "   and Municipio_Id ='" & _CodigoMunicipio & "'" & vbCrLf & _
                      "   and Codigo_Id    = " & _CodigoIbge
                Sqls.Add(sql)
        End Select
    End Sub
#End Region

End Class
