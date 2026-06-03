Imports Microsoft.VisualBasic
Imports System.Collections.Generic
Imports System.Data
Imports NGS.Lib.Uteis

<Serializable()> _
Public Class ListCPRxGrau
    Inherits List(Of CPRxGrau)

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

            Sql = "SELECT Cartorio_Id, EndCartorio_Id, CPR_Id, Grau_Id, Cliente, EndCliente" & vbCrLf & _
                  "  FROM CPRxGrau" & vbCrLf & _
                  " Where Cartorio_Id    ='" & _CPR.CodigoCartorio & "'" & vbCrLf & _
                  "   and EndCartorio_Id = " & _CPR.EndCartorio & vbCrLf & _
                  "   and CPR_Id         ='" & _CPR.CodigoCPR & "'" & vbCrLf & _
                  " order by Grau_id"

            Dim ds As DataSet = Banco.ConsultaDataSet(Sql, "CPRxGrau")

            For Each row As DataRow In ds.Tables(0).Rows
                Dim CxG As New CPRxGrau(pCPR)
                CxG.Grau = row("Grau_Id")
                CxG.CodigoCliente = row("Cliente")
                CxG.EndCliente = row("EndCliente")
                Me.Add(CxG)
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
        For Each CPxG As CPRxGrau In Me
            If _CPR.IUD = "D" Or _CPR.IUD = "I" Then CPxG.IUD = _CPR.IUD
            If CPxG.IUD <> "" Then
                CPxG.SalvarSql(Sqls)
            End If
        Next
    End Sub

    Public Function JaExiste(ByVal Gra As CPRxGrau) As Boolean
        For Each row As CPRxGrau In Me
            If row.CodigoCliente = Gra.CodigoCliente And row.EndCliente = Gra.EndCliente Then Return True
        Next
        Return False
    End Function
#End Region

End Class

<Serializable()> _
Public Class CPRxGrau

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
    Private _Grau As Integer
    Private _CodigoCliente As String
    Private _EndCliente As Integer
    Private _Cliente As Cliente
    Private _DescCliente As String = ""
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

    Public Property CodigoCliente() As String
        Get
            Return _CodigoCliente
        End Get
        Set(ByVal value As String)
            _CodigoCliente = value
            _Cliente = Nothing
            _DescCliente = ""
        End Set
    End Property

    Public Property EndCliente() As Integer
        Get
            Return _EndCliente
        End Get
        Set(ByVal value As Integer)
            _EndCliente = value
            _Cliente = Nothing
            _DescCliente = ""
        End Set
    End Property

    Public Property Cliente() As Cliente
        Get
            If _Cliente Is Nothing And _CodigoCliente.Length > 0 Then _Cliente = New Cliente(_CodigoCliente, _EndCliente)
            Return _Cliente
        End Get
        Set(ByVal value As Cliente)
            _Cliente = value
        End Set
    End Property

    Public Property DescCliente() As String
        Get
            If _DescCliente.Length = 0 Then _DescCliente = Cliente.Nome
            Return _DescCliente
        End Get
        Set(ByVal value As String)
            _DescCliente = value
        End Set
    End Property

    Public Property Grau() As Integer
        Get
            Return _Grau
        End Get
        Set(ByVal value As Integer)
            _Grau = value
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
                sql = " Insert Into CPRxGrau(Cartorio_Id, EndCartorio_Id, CPR_Id, Grau_Id, Cliente, EndCliente)" & vbCrLf & _
                      " Values ('" & _Cpr.CodigoCartorio & "'," & _Cpr.EndCartorio & ",'" & _Cpr.CodigoCPR & "'," & _Grau & ",'" & _CodigoCliente & "'," & _EndCliente & ")"

                Sqls.Add(sql)
            Case "D"
                sql = " Delete CPRxGrau" & vbCrLf & _
                      "	 Where Cartorio_Id    ='" & _Cpr.CodigoCartorio & "'" & vbCrLf & _
                      "    and EndCartorio_Id = " & _Cpr.EndCartorio & vbCrLf & _
                      "    and CPR_Id         ='" & _Cpr.CodigoCPR & "'" & vbCrLf & _
                      "    and Grau_Id        = " & _Grau & vbCrLf
                Sqls.Add(sql)
        End Select

    End Sub

#End Region

End Class
