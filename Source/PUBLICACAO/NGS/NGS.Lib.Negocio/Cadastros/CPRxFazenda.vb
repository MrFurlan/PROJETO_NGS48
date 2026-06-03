Imports Microsoft.VisualBasic
Imports System.Collections.Generic
Imports System.Data
Imports NGS.Lib.Uteis

Public Class ListCPRxFazenda
    Inherits List(Of CPRxFazenda)

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

            Sql = "SELECT Cartorio_Id, EndCartorio_Id, CPR_Id, Fazenda_id, EndFazenda_Id" & vbCrLf & _
                  "  FROM CPRxFazenda" & vbCrLf & _
                  " Where Cartorio_Id    ='" & _CPR.CodigoCartorio & "'" & vbCrLf & _
                  "   and EndCartorio_Id = " & _CPR.EndCartorio & vbCrLf & _
                  "   and CPR_Id         ='" & _CPR.CodigoCPR & "'"

            Dim ds As DataSet = Banco.ConsultaDataSet(Sql, "CPRxFaz")

            For Each row As DataRow In ds.Tables(0).Rows
                Dim CxF As New CPRxFazenda(pCPR)
                CxF.CodigoFazenda = row("Fazenda_id")
                CxF.EndFazenda = row("EndFazenda_Id")
                Me.Add(CxF)
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

    Public ReadOnly Property AreaTotal
        Get
            Return (From c In Me Select c.Area).Sum
        End Get
    End Property

#Region "Methods"
    Public Function Salvar() As Boolean
        Dim Banco As New AcessaBanco
        Dim Sqls As New ArrayList

        Sqls.Clear()
        SalvarSql(Sqls)

        Return Banco.GravaBanco(Sqls)
    End Function

    Public Sub SalvarSql(ByVal Sqls As ArrayList)
        For Each CPxFaz As CPRxFazenda In Me
            If _CPR.IUD = "D" Or _CPR.IUD = "I" Then CPxFaz.IUD = _CPR.IUD
            If CPxFaz.IUD <> "" Then
                CPxFaz.SalvarSql(Sqls)
            End If
        Next
    End Sub
#End Region

End Class


<Serializable()> _
Public Class CPRxFazenda
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
    Private _CodigoFazenda As String = ""
    Private _EndFazenda As Integer
    Private _Fazenda As Cliente
    Private _DescFazenda As String = ""
    Private _Clientes As ListCPRxCliente
    Private _Matriculas As ListCPRxMatricula
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

    Public Property CodigoFazenda() As String
        Get
            Return _CodigoFazenda
        End Get
        Set(ByVal value As String)
            _CodigoFazenda = value
            _Fazenda = Nothing
            _DescFazenda = ""
        End Set
    End Property

    Public Property EndFazenda() As Integer
        Get
            Return _EndFazenda
        End Get
        Set(ByVal value As Integer)
            _EndFazenda = value
            _Fazenda = Nothing
            _DescFazenda = ""
        End Set
    End Property

    Public Property Fazenda() As Cliente
        Get
            If _Fazenda Is Nothing And _CodigoFazenda.Length > 0 Then _Fazenda = New Cliente(_CodigoFazenda, _EndFazenda)
            Return _Fazenda
        End Get
        Set(ByVal value As Cliente)
            _Fazenda = value
        End Set
    End Property

    Public ReadOnly Property DescFazenda() As String
        Get
            If _DescFazenda.Length = 0 Then _DescFazenda = Fazenda.Nome + " - " + Fazenda.Complemento
            Return _DescFazenda
        End Get
    End Property

    Public Property Clientes() As ListCPRxCliente
        Get
            If _Clientes Is Nothing Then _Clientes = New ListCPRxCliente(Me)
            Return _Clientes
        End Get
        Set(ByVal value As ListCPRxCliente)
            _Clientes = value
        End Set
    End Property

    Public Property Matriculas() As ListCPRxMatricula
        Get
            If _Matriculas Is Nothing Then _Matriculas = New ListCPRxMatricula(Me)
            Return _Matriculas
        End Get
        Set(ByVal value As ListCPRxMatricula)
            _Matriculas = value
        End Set
    End Property

    Public ReadOnly Property Area() As Decimal
        Get
            Dim x As Decimal = 0
            For Each CxM As CPRxMatricula In Me.Matriculas
                x += CxM.Area
            Next
            Return x
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
                sql = " Insert Into CPRxFazenda(Cartorio_Id, EndCartorio_Id, CPR_Id, Fazenda_id, EndFazenda_Id)" & vbCrLf & _
                      " Values ('" & _Cpr.CodigoCartorio & "'," & _Cpr.EndCartorio & ",'" & _Cpr.CodigoCPR & "','" & _CodigoFazenda & "'," & _EndFazenda & ")"

                Sqls.Add(sql)
                '***********************************************************************
                '********* Procedimento para Salvar as tabelas relacionadas   **********
                '***********************************************************************
                SalvarTabelasRelacionadasSql(Sqls)
            Case "D"
                SalvarTabelasRelacionadasSql(Sqls)
                sql = " Delete CPRxFazenda" & vbCrLf & _
                      "	 Where Cartorio_Id    ='" & _Cpr.CodigoCartorio & "'" & vbCrLf & _
                      "    and EndCartorio_Id = " & _Cpr.EndCartorio & vbCrLf & _
                      "    and CPR_Id         ='" & _Cpr.CodigoCPR & "'" & vbCrLf & _
                      "    and Fazenda_id     ='" & _CodigoFazenda & "'" & vbCrLf & _
                      "    and EndFazenda_Id  = " & _EndFazenda
                Sqls.Add(sql)
        End Select

    End Sub

    Private Sub SalvarTabelasRelacionadasSql(ByRef Sqls As ArrayList)
        If Not Clientes Is Nothing Then Clientes.SalvarSql(Sqls)
        If Not Matriculas Is Nothing Then Matriculas.SalvarSql(Sqls)
    End Sub

#End Region

End Class
