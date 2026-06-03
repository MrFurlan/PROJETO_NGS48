Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Collections.Generic
Imports NGS.Lib.Uteis

<Serializable()> _
Public Class ListTipoDeCliente
    Inherits List(Of TipoDeCliente)

#Region "Construtor"
    Public Sub New(Optional ByVal Carregar As Boolean = False)
        Dim sql As String
        Dim ds As DataSet
        Dim Banco As New AcessaBanco

        sql = "SELECT Tipo_Id, Descricao" & _
              "  FROM TiposDeClientes" & _
              " Order by Descricao"

        ds = Banco.ConsultaDataSet(sql, "TipoDeCliente")

        For Each row As DataRow In ds.Tables(0).Rows
            Dim TC As New TipoDeCliente
            TC.CodigoTipo = row("Tipo_Id")
            TC.Descricao = row("Descricao")
            Me.Add(TC)
        Next
    End Sub
#End Region

#Region "Methods"
    Public Function Salvar() As Boolean
        Dim Banco As New AcessaBanco
        Dim Sqls As New ArrayList

        Sqls.Clear()
        SalvarSql(Sqls)

        Return Banco.GravaBanco(Sqls)
    End Function

    Public Sub SalvarSql(ByVal Sqls As ArrayList)
        For Each Cli As TipoDeCliente In Me
            If Cli.IUD <> "" Then
                Cli.SalvarSql(Sqls)
            End If
        Next
    End Sub
#End Region

End Class

<Serializable()> _
Public Class TipoDeCliente
    Implements IBaseEntity

#Region "Contrutor"
    Public Sub New()

    End Sub

    Public Sub New(ByVal pCodigoTipo As Integer)
        Dim Banco As New AcessaBanco
        Dim ds As DataSet
        Dim sql As String

        sql = "SELECT Tipo_Id, Descricao" & vbCrLf & _
              "  FROM TiposDeClientes " & vbCrLf & _
              " Where Tipo_Id = " & pCodigoTipo

        ds = Banco.ConsultaDataSet(sql, "TipoDeCliente")
        For Each row As DataRow In ds.Tables(0).Rows
            _CodigoTipo = row("Tipo_Id")
            _Descricao = row("Descricao")
        Next
    End Sub
#End Region

#Region "Fields"
    Private _IUD As String = ""
    Private _CodigoTipo As Integer
    Private _Descricao As String
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

    Public Property CodigoTipo() As Integer
        Get
            Return _CodigoTipo
        End Get
        Set(ByVal value As Integer)
            _CodigoTipo = value
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
                sql = "INSERT INTO TiposDeClientes (Tipo_Id, Descricao) " & vbCrLf & _
                      "VALUES (" & _CodigoTipo & ",'" & _Descricao & "')"
                Sqls.Add(sql)
            Case "U"
                sql = "UPDATE TiposDeClientes " & _
                      "   SET Descricao      ='" & _Descricao & "'" & vbCrLf & _
                      " WHERE Tipo_Id        = " & _CodigoTipo
                Sqls.Add(sql)
            Case "D"
                sql = "DELETE TiposDeClientes " & _
                      " WHERE Tipo_Id        = " & _CodigoTipo
                Sqls.Add(sql)
        End Select
    End Sub
#End Region

End Class