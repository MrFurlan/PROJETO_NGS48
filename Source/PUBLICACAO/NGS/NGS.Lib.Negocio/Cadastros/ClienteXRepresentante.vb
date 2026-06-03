Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Collections.Generic
Imports NGS.Lib.Uteis

<Serializable()> _
Public Class ListClienteXRepresentante
    Inherits List(Of ClienteXRepresentante)

#Region "Construtor"
    Public Sub New()

    End Sub

    Public Sub New(ByVal pcliente As Cliente)
        _Cliente = pcliente
        Dim sql As String
        sql = "SELECT CXR.Cliente_Id, CXR.EndCliente_Id, CXR.Representante_Id, CXR.EndRepresentante_Id, Representante.Nome AS NomeRepresentante, CXR.Principal" & vbCrLf & _
              "  FROM ClienteXRepresentante CXR " & vbCrLf & _
              "  INNER JOIN Clientes AS Representante " & vbCrLf & _
              "          ON CXR.Representante_Id    = Representante.Cliente_Id " & vbCrLf & _
              "         AND CXR.EndRepresentante_Id = Representante.Endereco_Id " & vbCrLf & _
              " Where CXR.Cliente_Id    ='" & pcliente.Codigo & "'" & vbCrLf & _
              "   and CXR.EndCliente_Id = " & pcliente.CodigoEndereco

        Dim Banco As New AcessaBanco
        Dim ds As New DataSet
        ds = Banco.ConsultaDataSet(sql, "ClienteXRepresentante")

        For Each row As DataRow In ds.Tables(0).Rows
            Dim CxR As New ClienteXRepresentante(pcliente)
            CxR.CodigoRepresentante = row("Representante_Id")
            CxR.EndRepresentante = row("EndRepresentante_Id")
            CxR.NomeRepresentante = row("NomeRepresentante")
            CxR.Principal = row("Principal")
            Me.Add(CxR)
        Next
    End Sub
#End Region

#Region "Fields"
    Private _Cliente As Cliente
#End Region

#Region "Property"
    Public Property Cliente() As Cliente
        Get
            Return _Cliente
        End Get
        Set(ByVal value As Cliente)
            _Cliente = value
        End Set
    End Property
#End Region

#Region "Methods"
    Public Sub SalvarSql(ByVal Sqls As ArrayList)
        For Each CxA As ClienteXRepresentante In Me
            If _Cliente.IUD = "D" Or _Cliente.IUD = "I" Then CxA.IUD = _Cliente.IUD
            If CxA.IUD <> "" Then
                CxA.SalvarSql(Sqls)
            End If
        Next
    End Sub
#End Region

End Class

<Serializable()> _
Public Class ClienteXRepresentante
    Implements IBaseEntity

#Region "Construtor"
    Public Sub New()

    End Sub

    Public Sub New(ByVal pCliente As Cliente)
        _Cliente = pCliente
    End Sub
#End Region

#Region "Fields"
    Private _IUD As String = ""
    Private _Cliente As Cliente
    Private _CodigoRepresentante As String
    Private _EndRepresentante As Integer
    Private _NomeRepresentante As String
    Private _Representante As Cliente
    Private _Principal As Boolean
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

    Public Property Cliente() As Cliente
        Get
            Return _Cliente
        End Get
        Set(ByVal value As Cliente)
            _Cliente = value
        End Set
    End Property

    Public Property CodigoRepresentante() As String
        Get
            Return _CodigoRepresentante
        End Get
        Set(ByVal value As String)
            _CodigoRepresentante = value
        End Set
    End Property

    Public Property EndRepresentante() As Integer
        Get
            Return _EndRepresentante
        End Get
        Set(ByVal value As Integer)
            _EndRepresentante = value
        End Set
    End Property

    Public Property NomeRepresentante() As String
        Get
            Return _NomeRepresentante
        End Get
        Set(ByVal value As String)
            _NomeRepresentante = value
        End Set
    End Property

    Public Property Representante() As Cliente
        Get
            If _Representante Is Nothing And _CodigoRepresentante.Trim.Length > 0 Then _Representante = New Cliente(_CodigoRepresentante, _EndRepresentante)
            Return _Representante
        End Get
        Set(ByVal value As Cliente)
            _Representante = value
        End Set
    End Property

    Public Property Principal() As Boolean
        Get
            Return _Principal
        End Get
        Set(ByVal value As Boolean)
            _Principal = value
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

    Public Sub SalvarSql(ByVal Sqls As ArrayList)
        Dim sql As String = ""

        Select Case _IUD
            Case "I"
                sql = "Insert into ClienteXRepresentante( Cliente_Id, EndCliente_Id, Representante_Id, EndRepresentante_Id, Principal)" & vbCrLf & _
                      " values('" & _Cliente.Codigo & "'," & _Cliente.CodigoEndereco & ",'" & _CodigoRepresentante & "'," & _EndRepresentante & "," & CByte(_Principal) & ")"
                Sqls.Add(sql)
            Case "U"
                sql = "Update ClienteXRepresentante set " & vbCrLf & _
                      "    Principal = " & CByte(_Principal) & vbCrLf & _
                      " Where Cliente_Id          ='" & _Cliente.Codigo & "'" & vbCrLf & _
                      "   and EndCliente_Id       = " & _Cliente.CodigoEndereco & vbCrLf & _
                      "   and Representante_Id    ='" & _CodigoRepresentante & "'" & vbCrLf & _
                      "   and EndRepresentante_Id = " & _EndRepresentante
                Sqls.Add(sql)
            Case "D"
                sql = "Delete ClienteXRepresentante " & vbCrLf & _
                      " Where Cliente_Id          ='" & _Cliente.Codigo & "'" & vbCrLf & _
                      "   and EndCliente_Id       = " & _Cliente.CodigoEndereco & vbCrLf & _
                      "   and Representante_Id    ='" & _CodigoRepresentante & "'" & vbCrLf & _
                      "   and EndRepresentante_Id = " & _EndRepresentante
                Sqls.Add(sql)
        End Select
    End Sub
#End Region

End Class