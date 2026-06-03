Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Collections.Generic
Imports NGS.Lib.Uteis

'***********************************************************************************************
'************************************ Classe LISTA *********************************************
'***********************************************************************************************
<Serializable()> _
Public Class ListCNAE
    Inherits List(Of CNAE)

    Public Sub New()

    End Sub

    Public Sub New(ByVal CarregarDados As Boolean)
        If Not CarregarDados Then Exit Sub

        Dim sql As String
        sql = "Select CNAE_Id, Descricao, Dt_Ini, Dt_Fin" & vbCrLf & _
              "  From CNAE" & vbCrLf

        Dim Banco As New AcessaBanco
        Dim ds As DataSet

        ds = Banco.ConsultaDataSet(sql, "CNAE")

        For Each row As DataRow In ds.Tables(0).Rows
            Dim CN As New CNAE
            CN.Codigo = row("CNAE_Id")
            CN.Descricao = row("Descricao")
            CN.DataIni = row("Dt_Ini")
            CN.DataFin = row("Dt_Fin")
            Me.Add(CN)
        Next
    End Sub

End Class


'***********************************************************************************************
'************************************ Classe BASE **********************************************
'***********************************************************************************************
<Serializable()> _
Public Class CNAE
    Implements IBaseEntity

#Region "Construtor"
    Public Sub New()

    End Sub

    Public Sub New(ByVal pCodigo As String)
        Dim Banco As New AcessaBanco
        Dim ds As DataSet
        Dim sql As String
        sql = "Select CNAE_Id, Descricao, Dt_Ini, Dt_Fin" & vbCrLf & _
              "  From CNAE" & vbCrLf & _
              " where CNAE_Id = " & pCodigo

        ds = Banco.ConsultaDataSet(sql, "CNAE")
        If ds.Tables(0).Rows.Count = 0 Then Exit Sub

        Me.Codigo = ds.Tables(0).Rows(0)("CNAE_Id")
        Me.Descricao = ds.Tables(0).Rows(0)("Descricao")
        Me.DataIni = ds.Tables(0).Rows(0)("Dt_Ini")
        Me.DataFin = ds.Tables(0).Rows(0)("Dt_Fin")

    End Sub
#End Region

#Region "Fields"
    Private _IUD As String = ""
    Private _Codigo As String = ""
    Private _Descricao As String = ""
    Private _DataIni As Date
    Private _DataFin As Date
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

    Public Property Codigo() As String
        Get
            Return _Codigo
        End Get
        Set(ByVal value As String)
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

    Public Property DataIni As Date
        Get
            Return _DataIni
        End Get
        Set(value As Date)
            _DataIni = value
        End Set
    End Property

    Public Property DataFin As Date
        Get
            Return _DataFin
        End Get
        Set(value As Date)
            _DataFin = value
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
                Sql = " INSERT INTO CNAE(CNAE_Id, Descricao, Dt_Ini, Dt_Fin) " & vbCrLf & _
                      " VALUES ('" & Me.Codigo & "','" & Me.Descricao & "','" & Me.DataIni.ToSqlNULL & "','" & Me.DataFin.ToSqlNULL & "')"
                Sqls.Add(Sql)

            Case "U"
                Sql = " UPDATE CNAE SET" & vbCrLf & _
                      "    Descricao ='" & Me.Descricao & "'" & vbCrLf & _
                      "   ,Dt_Ini    ='" & Me.DataIni.ToSqlNULL & "'" & vbCrLf & _
                      "   ,Dt_Fin    ='" & Me.DataFin.ToSqlNULL & "'" & vbCrLf & _
                      "  WHERE CNAE_Id ='" & Me.Codigo & "'" & vbCrLf
                Sqls.Add(Sql)

            Case "D"
                Sql = " DELETE CNAE" & vbCrLf & _
                      "  WHERE CNAE_Id ='" & Me.Codigo & "'" & vbCrLf
                Sqls.Add(Sql)
        End Select
    End Sub
#End Region

End Class