Imports Microsoft.VisualBasic
Imports System.Collections.Generic
Imports System.Data
Imports NGS.Lib.Uteis

<Serializable()> _
Public Class ListCPRxMatricula
    Inherits List(Of CPRxMatricula)

#Region "Fields"
    Public Erro As Exception
    Private _CPRxFazenda As CPRxFazenda
#End Region

#Region "Contrutor"
    Public Sub New(ByVal pCPRxFaz As CPRxFazenda)
        Dim Banco As New AcessaBanco()
        _CPRxFazenda = pCPRxFaz
        Try
            Dim Sql As String = ""
            Dim Where As String = ""

            Sql = "SELECT Cartorio_Id, EndCartorio_Id, CPR_Id, Fazenda_id, EndFazenda_Id, Matricula_Id, area" & vbCrLf & _
                  "  FROM CPRxMatricula" & vbCrLf & _
                  " Where Cartorio_Id    ='" & pCPRxFaz.Cpr.CodigoCartorio & "'" & vbCrLf & _
                  "   and EndCartorio_Id = " & pCPRxFaz.Cpr.EndCartorio & vbCrLf & _
                  "   and CPR_Id         ='" & pCPRxFaz.Cpr.CodigoCPR & "'" & vbCrLf & _
                  "   and Fazenda_Id     ='" & pCPRxFaz.CodigoFazenda & "'" & vbCrLf & _
                  "   and EndFazenda_Id  = " & pCPRxFaz.EndFazenda & vbCrLf

            Dim ds As DataSet = Banco.ConsultaDataSet(Sql, "CPRxFaz")

            For Each row As DataRow In ds.Tables(0).Rows
                Dim CxM As New CPRxMatricula(pCPRxFaz)
                CxM.CodigoMatricula = row("Matricula_Id")
                CxM.Area = row("area")
                Me.Add(CxM)
            Next

        Catch ex As Exception
            Me.Erro = ex
        Finally
            Banco = Nothing
        End Try
    End Sub
#End Region

#Region "Property"
    Public ReadOnly Property CPRxFazenda() As CPRxFazenda
        Get
            Return _CPRxFazenda
        End Get
    End Property
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
        For Each CPxM As CPRxMatricula In Me
            If _CPRxFazenda.IUD = "D" Or _CPRxFazenda.IUD = "I" Then CPxM.IUD = _CPRxFazenda.IUD
            If CPxM.IUD <> "" Then
                CPxM.SalvarSql(Sqls)
            End If
        Next
    End Sub
#End Region

End Class

<Serializable()> _
Public Class CPRxMatricula
    Implements IBaseEntity

#Region "Construtor"
    Public Sub New(ByVal pCPRxFazenda As CPRxFazenda)
        _CprXFazenda = pCPRxFazenda
    End Sub

    Public Sub New()

    End Sub
#End Region

#Region "Fields"
    Private _IUD As String
    Private _CprXFazenda As CPRxFazenda
    Private _CodigoMatricula As String
    Private _Area As Decimal
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

    Public Property CprXFazenda() As CPRxFazenda
        Get
            Return _CprXFazenda
        End Get
        Set(ByVal value As CPRxFazenda)
            _CprXFazenda = value
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
        Dim sql As String
        Select Case Me.IUD
            Case "I"
                sql = " Insert Into CPRxMatricula(Cartorio_Id, EndCartorio_Id, CPR_Id, Fazenda_id, EndFazenda_Id, Matricula_id, Area)" & vbCrLf & _
                      " Values ('" & _CprXFazenda.Cpr.CodigoCartorio & "'," & _CprXFazenda.Cpr.EndCartorio & ",'" & _CprXFazenda.Cpr.CodigoCPR & "','" & _CprXFazenda.CodigoFazenda & "'," & _CprXFazenda.EndFazenda & ",'" & _CodigoMatricula & "'," & Str(_Area) & ")"
                Sqls.Add(sql)
            Case "D"
                sql = " Delete CPRxMatricula" & vbCrLf & _
                      "	 Where Cartorio_Id    ='" & _CprXFazenda.Cpr.CodigoCartorio & "'" & vbCrLf & _
                      "    and EndCartorio_Id = " & _CprXFazenda.Cpr.EndCartorio & vbCrLf & _
                      "    and CPR_Id         ='" & _CprXFazenda.Cpr.CodigoCPR & "'" & vbCrLf & _
                      "    and Fazenda_id     ='" & _CprXFazenda.CodigoFazenda & "'" & vbCrLf & _
                      "    and EndFazenda_Id  = " & _CprXFazenda.EndFazenda & vbCrLf & _
                      "    and Matricula_Id   ='" & _CodigoMatricula & "'" & vbCrLf
                Sqls.Add(sql)
        End Select

    End Sub

#End Region

End Class