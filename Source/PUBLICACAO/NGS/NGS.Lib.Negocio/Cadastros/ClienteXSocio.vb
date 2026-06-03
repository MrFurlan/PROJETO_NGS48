Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Collections.Generic
Imports NGS.Lib.Uteis

<Serializable()> _
Public Class ListClienteXSocio
    Inherits List(Of ClienteXSocio)

#Region "Construtor"
    Public Sub New()

    End Sub

    Public Sub New(ByVal pcliente As Cliente)
        _Cliente = pcliente
        Dim sql As String
        sql = "SELECT CxS.Cliente_Id, CxS.EndCliente_Id, CxS.Socio_Id, CxS.EndSocio_Id, Socio.Nome AS NomeSocio, CxS.Participacao" & vbCrLf & _
              "  FROM ClienteXSocio AS CxS" & vbCrLf & _
              "  INNER JOIN Clientes AS Socio " & vbCrLf & _
              "          ON CxS.Socio_Id    = Socio.Cliente_Id " & vbCrLf & _
              "         AND CxS.EndSocio_Id = Socio.Endereco_Id " & vbCrLf & _
              " Where CxS.Cliente_Id    = '" & pcliente.Codigo & "'" & vbCrLf & _
              "   and CxS.EndCliente_Id = " & pcliente.CodigoEndereco

        Dim Banco As New AcessaBanco
        Dim ds As New DataSet
        ds = Banco.ConsultaDataSet(sql, "ClienteXSocio")

        For Each row As DataRow In ds.Tables(0).Rows
            Dim CxS As New ClienteXSocio(pcliente)
            CxS.CodigoSocio = row("Socio_Id")
            CxS.EndSocio = row("EndSocio_Id")
            CxS.NomeSocio = row("NomeSocio")
            CxS.Participacao = row("Participacao")
            Me.Add(CxS)
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
        For Each CxS As ClienteXSocio In Me
            If _Cliente.IUD = "D" Or _Cliente.IUD = "I" Then CxS.IUD = _Cliente.IUD
            If CxS.IUD <> "" Then
                CxS.SalvarSql(Sqls)
            End If
        Next
    End Sub
#End Region

End Class

<Serializable()> _
Public Class ClienteXSocio
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
    Private _CodigoSocio As String
    Private _EndSocio As Integer
    Private _NomeSocio As String
    Private _Socio As Cliente
    Private _Participacao As Decimal
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

    Public Property CodigoSocio() As String
        Get
            Return _CodigoSocio
        End Get
        Set(ByVal value As String)
            _CodigoSocio = value
        End Set
    End Property

    Public Property EndSocio() As Integer
        Get
            Return _EndSocio
        End Get
        Set(ByVal value As Integer)
            _EndSocio = value
        End Set
    End Property

    Public Property NomeSocio() As String
        Get
            Return _NomeSocio
        End Get
        Set(ByVal value As String)
            _NomeSocio = value
        End Set
    End Property

    Public Property Socio() As Cliente
        Get
            If _Socio Is Nothing And _CodigoSocio.Trim.Length > 0 Then _Socio = New Cliente(_CodigoSocio, _EndSocio)
            Return _Socio
        End Get
        Set(ByVal value As Cliente)
            _Socio = value
        End Set
    End Property

    Public Property Participacao() As Decimal
        Get
            Return _Participacao
        End Get
        Set(ByVal value As Decimal)
            _Participacao = value
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
                sql = "Insert into ClienteXSocio(Cliente_Id, EndCliente_Id, Socio_Id, EndSocio_Id, Participacao)" & vbCrLf & _
                      " values('" & _Cliente.Codigo & "'," & _Cliente.CodigoEndereco & ",'" & _CodigoSocio & "'," & _EndSocio & "," & Str(_Participacao) & ")"
                Sqls.Add(sql)
            Case "U"
                sql = "Update ClienteXSocio set " & vbCrLf & _
                      "    Participacao = " & Str(_Participacao) & vbCrLf & _
                      " Where Cliente_Id    ='" & _Cliente.Codigo & "'" & vbCrLf & _
                      "   and EndCliente_Id = " & _Cliente.CodigoEndereco & vbCrLf & _
                      "   and Socio_Id      ='" & _CodigoSocio & "'" & vbCrLf & _
                      "   and EndSocio_Id   = " & _EndSocio
                Sqls.Add(sql)
            Case "D"
                sql = "Delete ClienteXSocio " & vbCrLf & _
                      " Where Cliente_Id    ='" & _Cliente.Codigo & "'" & vbCrLf & _
                      "   and EndCliente_Id = " & _Cliente.CodigoEndereco & vbCrLf & _
                      "   and Socio_Id      ='" & _CodigoSocio & "'" & vbCrLf & _
                      "   and EndSocio_Id   = " & _EndSocio
                Sqls.Add(sql)
        End Select
    End Sub
#End Region

End Class