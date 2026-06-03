Imports Microsoft.VisualBasic
Imports System.Collections.Generic
Imports System.Data
Imports NGS.Lib.Uteis

<Serializable()> _
Public Class ListCPRxAvalista
    Inherits List(Of CPRxAvalista)

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

            Sql = "SELECT Cartorio_Id, EndCartorio_Id, CPR_Id, Avalista_Id, EndAvalista_Id, TipoRelacao" & vbCrLf & _
                  "  FROM CPRxAvalista" & vbCrLf & _
                  " Where Cartorio_Id    ='" & _CPR.CodigoCartorio & "'" & vbCrLf & _
                  "   and EndCartorio_Id = " & _CPR.EndCartorio & vbCrLf & _
                  "   and CPR_Id         ='" & _CPR.CodigoCPR & "'"

            Dim ds As DataSet = Banco.ConsultaDataSet(Sql, "CPRxAva")

            For Each row As DataRow In ds.Tables(0).Rows
                Dim CxA As New CPRxAvalista(pCPR)
                CxA.CodigoAvalista = row("Avalista_id")
                CxA.EndAvalista = row("EndAvalista_Id")
                CxA.TipoDeRelacao = CType(CInt(row("TipoRelacao")), eTipoDeRelacaoCPR)
                Me.Add(CxA)
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
        For Each CPxA As CPRxAvalista In Me
            If _CPR.IUD = "D" Or _CPR.IUD = "I" Then CPxA.IUD = _CPR.IUD
            If CPxA.IUD <> "" Then
                CPxA.SalvarSql(Sqls)
            End If
        Next
    End Sub

    Public Function JaExiste(ByVal Ava As CPRxAvalista) As Boolean
        For Each row As CPRxAvalista In Me
            If row.CodigoAvalista = Ava.CodigoAvalista And row.EndAvalista = Ava.EndAvalista Then Return True
        Next
        Return False
    End Function
#End Region

End Class

<Serializable()> _
Public Class CPRxAvalista
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
    Private _CodigoAvalista As String = ""
    Private _EndAvalista As Integer
    Private _Avalista As Cliente
    Private _DescAvalista As String = ""
    Private _TipoRelacao As System.Nullable(Of eTipoDeRelacaoCPR)
    'Private _DescTipoRelacao As String
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

    Public Property CodigoAvalista() As String
        Get
            Return _CodigoAvalista
        End Get
        Set(ByVal value As String)
            _CodigoAvalista = value
        End Set
    End Property

    Public Property EndAvalista() As Integer
        Get
            Return _EndAvalista
        End Get
        Set(ByVal value As Integer)
            _EndAvalista = value
        End Set
    End Property

    Public Property Avalista() As Cliente
        Get
            If _Avalista Is Nothing And _CodigoAvalista.Length > 0 Then _Avalista = New Cliente(_CodigoAvalista, _EndAvalista)
            Return _Avalista
        End Get
        Set(ByVal value As Cliente)
            _Avalista = value
        End Set
    End Property

    Public ReadOnly Property DescAvalista()
        Get
            If _DescAvalista.Length = 0 Then _DescAvalista = Avalista.Nome
            Return _DescAvalista
        End Get
    End Property

    Public Property TipoDeRelacao() As eTipoDeRelacaoCPR
        Get
            Return _TipoRelacao
        End Get
        Set(ByVal value As eTipoDeRelacaoCPR)
            _TipoRelacao = value
        End Set
    End Property

    Public ReadOnly Property DescTipoRelacao()
        Get
            Return WebHelpers.GetEnumDescription(TipoDeRelacao)
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
        Dim sql As String
        Select Case Me.IUD
            Case "I"
                sql = " Insert Into CPRxAvalista(Cartorio_Id, EndCartorio_Id, CPR_Id, Avalista_id, EndAvalista_Id, TipoRelacao)" & vbCrLf & _
                      " Values ('" & _Cpr.CodigoCartorio & "'," & _Cpr.EndCartorio & ",'" & _Cpr.CodigoCPR & "','" & Me.CodigoAvalista & "'," & Me.EndAvalista & "," & Me.TipoDeRelacao & ")"

                Sqls.Add(sql)
            Case "D"
                sql = " Delete CPRxAvalista" & vbCrLf & _
                      "	 Where Cartorio_Id    ='" & _Cpr.CodigoCartorio & "'" & vbCrLf & _
                      "    and EndCartorio_Id = " & _Cpr.EndCartorio & vbCrLf & _
                      "    and CPR_Id         ='" & _Cpr.CodigoCPR & "'" & vbCrLf & _
                      "    and Avalista_id    ='" & _CodigoAvalista & "'" & vbCrLf & _
                      "    and EndAvalista_Id = " & _EndAvalista
                Sqls.Add(sql)
        End Select

    End Sub

#End Region

End Class