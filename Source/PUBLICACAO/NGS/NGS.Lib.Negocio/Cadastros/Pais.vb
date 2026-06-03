Imports Microsoft.VisualBasic
Imports System.Data
Imports NGS.Lib.Uteis

<Serializable()> _
Public Class ListPais
    Inherits List(Of Pais)

    Public Sub New()

    End Sub

    Public Sub New(ByVal CarregarDados As Boolean)
        If CarregarDados Then
            Dim sql As String = "Select Pais_Id, Descricao from Pais "

            Dim Banco As New AcessaBanco
            Dim ds As DataSet

            ds = Banco.ConsultaDataSet(sql, "Pais")

            For Each row As DataRow In ds.Tables(0).Rows
                Dim P As New Pais
                P.Codigo = row("Pais_Id")
                P.Descricao = row("Descricao")
                Me.Add(P)
            Next
        End If
    End Sub
End Class

<Serializable()> _
Public Class Pais
    Implements IBaseEntity

#Region "Construtor"
    Public Sub New()

    End Sub

    Public Sub New(ByVal pCodigo As String)
        Dim Banco As New AcessaBanco
        Dim ds As DataSet
        Dim sql As String = "Select Pais_Id, Descricao from Pais where Pais_Id = " & pCodigo

        ds = Banco.ConsultaDataSet(sql, "Pais")
        If ds.Tables(0).Rows.Count > 0 Then
            _Codigo = ds.Tables(0).Rows(0)("Pais_Id")
            _Descricao = ds.Tables(0).Rows(0)("Descricao")
        End If
    End Sub
#End Region

#Region "Fields"
    Private _IUD As String = ""
    Private _Codigo As Integer = 0
    Private _Descricao As String = ""
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
                Sql = " INSERT INTO Pais(Pais_Id, Descricao) " & vbCrLf & _
                      " VALUES (" & _Codigo & ",'" & _Descricao & "')"
                Sqls.Add(Sql)
            Case "U"
                Sql = " UPDATE Pais SET" & vbCrLf & _
                      "    Descricao   = '" & _Descricao & "'" & vbCrLf & _
                      "  WHERE Pais_Id = " & _Codigo & vbCrLf
                Sqls.Add(Sql)
            Case "D"
                Sql = " DELETE Pais" & vbCrLf & _
                      "  WHERE Pais_Id = " & _Codigo & vbCrLf
                Sqls.Add(Sql)
        End Select
    End Sub
#End Region

End Class