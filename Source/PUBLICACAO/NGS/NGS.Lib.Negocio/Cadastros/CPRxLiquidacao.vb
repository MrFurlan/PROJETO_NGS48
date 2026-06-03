Imports Microsoft.VisualBasic
Imports System.Collections.Generic
Imports System.Data
Imports NGS.Lib.Uteis

<Serializable()> _
Public Class ListCPRxLiquidacao
    Inherits List(Of CPRxLiquidacao)

#Region "Fields"
    Public Erro As Exception
    Private _CPR As CPR
#End Region

#Region "Contrutor"
    Public Sub New(ByVal pCPR As CPR)
        Dim Banco As New AcessaBanco()
        _CPR = pCPR
        Try
            Dim Sql As String = ""
            Dim Where As String = ""

            Sql = "SELECT Cartorio_Id, EndCartorio_Id, CPR_Id, Liquidacao_Id, Data, Quantidade" & vbCrLf & _
                  "  FROM CPRxLiquidacao" & vbCrLf & _
                  " Where Cartorio_Id    ='" & Me.CPR.CodigoCartorio & "'" & vbCrLf & _
                  "   and EndCartorio_Id = " & Me.CPR.EndCartorio & vbCrLf & _
                  "   and CPR_Id         ='" & Me.CPR.CodigoCPR & "'" & vbCrLf & _
                  " order by Liquidacao_id"

            Dim ds As DataSet = Banco.ConsultaDataSet(Sql, "CPRxLiquidacao")

            For Each row As DataRow In ds.Tables(0).Rows
                Dim CxL As New CPRxLiquidacao(pCPR)
                CxL.Liquidacao = row("Liquidacao_Id")
                CxL.Data = row("Data")
                CxL.Quantidade = row("Quantidade")
                Me.Add(CxL)
            Next

        Catch ex As Exception
            Me.Erro = ex
        Finally
            Banco = Nothing
        End Try
    End Sub
#End Region

#Region "Property"
    Public ReadOnly Property CPR() As CPR
        Get
            Return _CPR
        End Get
    End Property

    Public ReadOnly Property QuantidadeTotalLiquidada
        Get
            Return (From c In Me Select c.Quantidade).Sum
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
        For Each CPxG As CPRxLiquidacao In Me
            If _CPR.IUD = "D" Or _CPR.IUD = "I" Then CPxG.IUD = _CPR.IUD
            If CPxG.IUD <> "" Then
                CPxG.SalvarSql(Sqls)
            End If
        Next
    End Sub
#End Region

End Class

<Serializable()> _
Public Class CPRxLiquidacao
    Implements IBaseEntity

#Region "Construtor"
    Public Sub New(ByVal pCPR As CPR)
        _Cpr = pCPR
    End Sub

    Public Sub New()

    End Sub
#End Region

#Region "Fields"
    Private _IUD As String
    Private _Cpr As CPR
    Private _Liquidacao As Integer
    Private _Data As Date
    Private _Quantidade As Decimal
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

    Public Property Cpr() As CPR
        Get
            Return _Cpr
        End Get
        Set(ByVal value As CPR)
            _Cpr = value
        End Set
    End Property

    Public Property Liquidacao() As Integer
        Get
            Return _Liquidacao
        End Get
        Set(ByVal value As Integer)
            _Liquidacao = value
        End Set
    End Property

    Public Property Data() As Date
        Get
            Return _Data
        End Get
        Set(ByVal value As Date)
            _Data = value
        End Set
    End Property

    Public Property Quantidade() As Decimal
        Get
            Return _Quantidade
        End Get
        Set(ByVal value As Decimal)
            _Quantidade = value
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
                sql = " Insert Into CPRxLiquidacao(Cartorio_Id, EndCartorio_Id, CPR_Id, Liquidacao_Id, Data, Quantidade)" & vbCrLf & _
                      " Values ('" & Me.Cpr.CodigoCartorio & "'," & Me.Cpr.EndCartorio & ",'" & Me.Cpr.CodigoCPR & "'," & Me.Liquidacao & ",'" & Me.Data.ToString("yyyy-MM-dd") & "'," & Str(Me.Quantidade) & ")"

                Sqls.Add(sql)
            Case "D"
                sql = " Delete CPRxLiquidacao" & vbCrLf & _
                      "	 Where Cartorio_Id    ='" & Me.Cpr.CodigoCartorio & "'" & vbCrLf & _
                      "    and EndCartorio_Id = " & Me.Cpr.EndCartorio & vbCrLf & _
                      "    and CPR_Id         ='" & Me.Cpr.CodigoCPR & "'" & vbCrLf & _
                      "    and Liquidacao_Id        = " & Me.Liquidacao & vbCrLf
                Sqls.Add(sql)
        End Select

    End Sub

#End Region

End Class
