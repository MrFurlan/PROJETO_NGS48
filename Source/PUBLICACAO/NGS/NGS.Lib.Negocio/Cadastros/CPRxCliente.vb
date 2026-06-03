Imports Microsoft.VisualBasic
Imports System.Collections.Generic
Imports System.Data
Imports NGS.Lib.Uteis

<Serializable()> _
Public Class ListCPRxCliente
    Inherits List(Of CPRxCliente)

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

            Sql = "SELECT Cartorio_Id, EndCartorio_Id, CPR_Id, Fazenda_id, EndFazenda_Id, Cliente_Id, EndCliente_Id" & vbCrLf & _
                  "  FROM CPRxCliente" & vbCrLf & _
                  " Where Cartorio_Id    ='" & pCPRxFaz.Cpr.CodigoCartorio & "'" & vbCrLf & _
                  "   and EndCartorio_Id = " & pCPRxFaz.Cpr.EndCartorio & vbCrLf & _
                  "   and CPR_Id         ='" & pCPRxFaz.Cpr.CodigoCPR & "'" & vbCrLf & _
                  "   and Fazenda_Id     ='" & pCPRxFaz.CodigoFazenda & "'" & vbCrLf & _
                  "   and EndFazenda_Id  = " & pCPRxFaz.EndFazenda & vbCrLf

            Dim ds As DataSet = Banco.ConsultaDataSet(Sql, "CPRxFaz")

            For Each row As DataRow In ds.Tables(0).Rows
                Dim CxC As New CPRxCliente(pCPRxFaz)
                CxC.CodigoCliente = row("Cliente_Id")
                CxC.EndCliente = row("EndCliente_Id")
                Me.Add(CxC)
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
        For Each CPxCli As CPRxCliente In Me
            If _CPRxFazenda.IUD = "D" Or _CPRxFazenda.IUD = "I" Then CPxCli.IUD = _CPRxFazenda.IUD
            If CPxCli.IUD <> "" Then
                CPxCli.SalvarSql(Sqls)
            End If
        Next
    End Sub
#End Region

End Class

<Serializable()> _
Public Class CPRxCliente
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
    Private _CodigoCliente As String = ""
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

    Public Property CprXFazenda() As CPRxFazenda
        Get
            Return _CprXFazenda
        End Get
        Set(ByVal value As CPRxFazenda)
            _CprXFazenda = value
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

    Public ReadOnly Property DescCliente() As String
        Get
            If _DescCliente.Length = 0 Then _DescCliente = Cliente.Nome
            Return _DescCliente
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
                sql = " Insert Into CPRxCliente(Cartorio_Id, EndCartorio_Id, CPR_Id, Fazenda_id, EndFazenda_Id, Cliente_id, EndCliente_Id)" & vbCrLf & _
                      " Values ('" & _CprXFazenda.Cpr.CodigoCartorio & "'," & _CprXFazenda.Cpr.EndCartorio & ",'" & _CprXFazenda.Cpr.CodigoCPR & "','" & _CprXFazenda.CodigoFazenda & "'," & _CprXFazenda.EndFazenda & ",'" & _CodigoCliente & "'," & _EndCliente & ")"

                Sqls.Add(sql)
            Case "D"
                sql = " Delete CPRxCliente" & vbCrLf & _
                      "	 Where Cartorio_Id    ='" & _CprXFazenda.Cpr.CodigoCartorio & "'" & vbCrLf & _
                      "    and EndCartorio_Id = " & _CprXFazenda.Cpr.EndCartorio & vbCrLf & _
                      "    and CPR_Id         ='" & _CprXFazenda.Cpr.CodigoCPR & "'" & vbCrLf & _
                      "    and Fazenda_id     ='" & _CprXFazenda.CodigoFazenda & "'" & vbCrLf & _
                      "    and EndFazenda_Id  = " & _CprXFazenda.EndFazenda & vbCrLf & _
                      "    and Cliente_id     ='" & _CodigoCliente & "'" & vbCrLf & _
                      "    and EndCliente_Id  = " & _EndCliente
                Sqls.Add(sql)
        End Select

    End Sub

#End Region

End Class
