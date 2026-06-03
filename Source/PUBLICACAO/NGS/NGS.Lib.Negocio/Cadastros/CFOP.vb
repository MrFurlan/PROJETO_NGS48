Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Collections.Generic
Imports NGS.Lib.Uteis

<Serializable()> _
Public Class ListCFOP
    Inherits List(Of CFOP)

    Public Sub New()

    End Sub

    Public Sub New(ByVal CarregarDados As Boolean, Optional ByVal GrupoCFOP As String = "")
        If CarregarDados Then
            Dim sql As String = "Select CFOP_Id, GrupoCfop_Id, Descricao From CFOP " & vbCrLf
            If GrupoCFOP.Length > 0 Then
                sql &= " Where GrupoCfop_Id = " & GrupoCFOP & vbCrLf
            End If
            sql &= " Order By CFOP_Id"

            Dim Banco As New AcessaBanco
            Dim ds As DataSet

            ds = Banco.ConsultaDataSet(sql, "CFOP")

            For Each row As DataRow In ds.Tables(0).Rows
                Dim CF As New CFOP
                CF.Codigo = row("CFOP_Id")
                CF.CodigoGrupo = row("GrupoCfop_Id")
                CF.Descricao = row("Descricao")
                Me.Add(CF)
            Next
        End If
    End Sub

End Class

<Serializable()> _
Public Class CFOP
    Implements IBaseEntity

#Region "Construtor"
    Public Sub New()

    End Sub

    Public Sub New(ByVal pCodigo As String)
        Dim Banco As New AcessaBanco
        Dim ds As DataSet
        Dim sql As String = "Select CFOP_Id, GrupoCfop_Id, Descricao from CFOP where CFOP_Id = " & pCodigo

        ds = Banco.ConsultaDataSet(sql, "CFOP")
        If ds.Tables(0).Rows.Count > 0 Then
            _Codigo = ds.Tables(0).Rows(0)("CFOP_Id")
            _CodigoGrupo = ds.Tables(0).Rows(0)("GrupoCfop_Id")
            _Descricao = ds.Tables(0).Rows(0)("Descricao")
        End If
    End Sub
#End Region

#Region "Fields"
    Private _IUD As String = ""
    Private _Codigo As Integer = 0
    Private _CodigoGrupo As Integer = 0
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

    Public Property CodigoGrupo() As Integer
        Get
            Return _CodigoGrupo
        End Get
        Set(ByVal value As Integer)
            _CodigoGrupo = value
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
                Sql = " INSERT INTO CFOP(CFOP_Id, GrupoCfop_Id, Descricao) " & vbCrLf & _
                      " VALUES (" & _Codigo & "," & _CodigoGrupo & ",'" & _Descricao & "')"
                Sqls.Add(Sql)
            Case "U"
                Sql = " UPDATE CFOP SET" & vbCrLf & _
                      "    GrupoCfop_Id  = " & _CodigoGrupo & "," & vbCrLf & _
                      "    Descricao     = '" & _Descricao & "'" & vbCrLf & _
                      "  WHERE CFOP_Id = " & _Codigo & vbCrLf
                Sqls.Add(Sql)
            Case "D"
                Sql = " DELETE CFOP" & vbCrLf & _
                      "  WHERE CFOP_Id = " & _Codigo & vbCrLf
                Sqls.Add(Sql)
        End Select
    End Sub
#End Region

End Class