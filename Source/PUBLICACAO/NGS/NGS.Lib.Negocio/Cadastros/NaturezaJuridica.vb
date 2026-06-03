Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Collections.Generic
Imports NGS.Lib.Uteis

'***********************************************************************************************
'************************************ Classe LISTA *********************************************
'***********************************************************************************************
<Serializable()> _
Public Class ListNaturezaJuridica
    Inherits List(Of NaturezaJuridica)

    Public Sub New()

    End Sub

    Public Sub New(ByVal CarregarDados As Boolean, Optional ByVal GrupoNatJur As String = "")
        If Not CarregarDados Then Exit Sub

        Dim sql As String
        sql = "Select NatJur_Id, Nome, Data_Ini, Data_Fin, Grupo" & vbCrLf & _
              "  From NaturezaJuridica" & vbCrLf

        If GrupoNatJur.Length > 0 Then
            sql &= " Where Grupo = " & GrupoNatJur & vbCrLf
        End If
        sql &= " Order By NatJur_Id"

        Dim Banco As New AcessaBanco
        Dim ds As DataSet

        ds = Banco.ConsultaDataSet(sql, "NatJur")

        For Each row As DataRow In ds.Tables(0).Rows
            Dim NatJur As New NaturezaJuridica

            NatJur.Codigo = row("NatJur_Id")
            NatJur.CodigoGrupo = row("Grupo")
            NatJur.Nome = row("Nome")
            NatJur.DataIni = row("Data_Ini")
            NatJur.DataFin = row("Data_Fin")

            Me.Add(NatJur)
        Next
    End Sub

End Class


'***********************************************************************************************
'************************************ Classe BASE **********************************************
'***********************************************************************************************
<Serializable()> _
Public Class NaturezaJuridica
    Implements IBaseEntity

#Region "Construtor"
    Public Sub New()

    End Sub

    Public Sub New(ByVal pCodigo As String)
        Dim Banco As New AcessaBanco
        Dim ds As DataSet
        Dim sql As String
        sql = "Select NatJur_Id, Nome, Data_Ini, Data_Fin, Grupo" & vbCrLf & _
              "  From NaturezaJuridica" & vbCrLf & _
              " Where NatJur_Id = " & pCodigo

        ds = Banco.ConsultaDataSet(sql, "NatJur")
        If ds.Tables(0).Rows.Count = 0 Then Exit Sub

        Me.Codigo = ds.Tables(0).Rows(0)("NatJur_Id")
        Me.CodigoGrupo = ds.Tables(0).Rows(0)("Grupo")
        Me.Nome = ds.Tables(0).Rows(0)("Nome")
        Me.DataIni = ds.Tables(0).Rows(0)("Data_Ini")
        Me.DataFin = ds.Tables(0).Rows(0)("Data_Fin")
    End Sub
#End Region

#Region "Fields"
    Private _IUD As String = ""
    Private _Codigo As Integer
    Private _Nome As String = ""
    Private _DataIni As Date
    Private _DataFin As Date
    Private _CodigoGrupo As Integer
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

    Public Property Nome() As String
        Get
            Return _Nome
        End Get
        Set(ByVal value As String)
            _Nome = value
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
                Sql = " INSERT INTO NaturezaJuridica(NatJur_Id, Nome, Data_Ini, Data_Fin, Grupo) " & vbCrLf & _
                      " VALUES (" & Me.Codigo & ",'" & Me.Nome & "'," & Me.DataIni.ToSqlNULL & "," & Me.DataFin.ToSqlNULL & "," & Me.CodigoGrupo & ")"
                Sqls.Add(Sql)
            Case "U"
                Sql = " UPDATE NaturezaJuridica SET" & vbCrLf & _
                      "     nome      ='" & Me.Nome & "'" & vbCrLf & _
                      "    ,Data_Ini  = " & Me.DataIni.ToSqlNULL & vbCrLf & _
                      "    ,Data_Fin  = " & Me.DataFin.ToSqlNULL & vbCrLf & _
                      "    ,Grupo     = " & Me.CodigoGrupo & vbCrLf & _
                      "  WHERE NatJur_Id = " & Me.Codigo & vbCrLf
                Sqls.Add(Sql)
            Case "D"
                Sql = " DELETE NaturezaJuridica" & vbCrLf & _
                      "  WHERE NatJur_Id ='" & Me.Codigo & "'" & vbCrLf
                Sqls.Add(Sql)
        End Select
    End Sub
#End Region

End Class