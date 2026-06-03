Imports Microsoft.VisualBasic
Imports System.Data
Imports NGS.Lib.Uteis

'***********************************************************************************************************
'*********************************  Lista de Depositos do Pedido  ******************************************
'***********************************************************************************************************
Public Class ListPedidoxDeposito
    Inherits List(Of PedidoXDeposito)

#Region "Construtor"
    Public Sub New(ByVal pPedido As Negocio.Pedido)
        _Pedido = pPedido
        If _Pedido.Codigo = 0 Then Exit Sub

        Dim Banco As New AcessaBanco
        Try
            Dim SQL As String = ""
            SQL = "SELECT Deposito_Id, EndDeposito_Id, Quantidade, isnull(Tipo,'DE') AS Tipo, isnull(Principal,1) AS Principal " & _
                  "  FROM PedidosXDepositos " & _
                  " WHERE Empresa_Id    ='" & Me.Pedido.CodigoEmpresa & "' " & _
                  "   AND EndEmpresa_Id = " & Me.Pedido.EnderecoEmpresa.ToString() & " " & _
                  "   AND Pedido_Id     = " & Me.Pedido.Codigo.ToString()

            Dim ds As DataSet = Banco.ConsultaDataSet(SQL, "Comissoes")

            For Each row As DataRow In ds.Tables(0).Rows
                Dim dep As New Negocio.PedidoXDeposito(Pedido)

                dep.Codigo = row("Deposito_Id")
                dep.CodigoEndereco = row("EndDeposito_Id")
                dep.Quantidade = row("Quantidade")
                dep.Tipo = row("Tipo")
                dep.Principal = row("Principal")
                Me.Add(dep)
            Next
        Catch ex As Exception
        Finally
            Banco = Nothing
        End Try
    End Sub
#End Region

#Region "Field"
    Private _Pedido As Negocio.Pedido
#End Region

#Region "Property"
    Public Property Pedido As Negocio.Pedido
        Get
            Return _Pedido
        End Get
        Set(value As Negocio.Pedido)
            _Pedido = value
        End Set
    End Property
#End Region

#Region "Methods"
    Public Sub SalvarSql(ByVal Sqls As ArrayList)
        For Each dep As PedidoXDeposito In Me
            If Pedido.IUD = "I" Or Pedido.IUD = "D" Then dep.IUD = Pedido.IUD
            dep.SalvarSql(Sqls)
        Next
    End Sub
#End Region

End Class

'***********************************************************************************************************
'**************************************   Depositos do Pedido   ********************************************
'***********************************************************************************************************
Public Class PedidoXDeposito

#Region "Contrutor"
    Public Sub New(ByVal pPedido As Negocio.Pedido)
        _Pedido = pPedido
    End Sub
#End Region

#Region "Fields"
    Private _IUD As String = ""
    Private _Pedido As Negocio.Pedido
    Private _Codigo As String = ""
    Private _Endereco As Integer
    Private _Deposito As Negocio.Cliente
    Private _Quantidade As Decimal
    Private _Tipo As String = ""
    Private _Principal As Boolean = False
#End Region

#Region "Property"
    Public Property IUD As String
        Get
            Return _IUD
        End Get
        Set(value As String)
            _IUD = value
        End Set
    End Property

    Public ReadOnly Property Pedido As Negocio.Pedido
        Get
            Return _Pedido
        End Get
    End Property

    Public Property Codigo() As String
        Get
            Return _Codigo
        End Get
        Set(ByVal value As String)
            _Codigo = value
        End Set
    End Property

    Public Property CodigoEndereco() As Integer
        Get
            Return _Endereco
        End Get
        Set(ByVal value As Integer)
            _Endereco = value
        End Set
    End Property

    Public Property Deposito() As Negocio.Cliente
        Get
            If _Deposito Is Nothing And Not Codigo Is Nothing Then _Deposito = New Negocio.Cliente(Me.Codigo, Me.CodigoEndereco)
            Return _Deposito
        End Get
        Set(ByVal value As Negocio.Cliente)
            _Deposito = value
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

    Public Property Tipo() As String
        Get
            Return _Tipo
        End Get
        Set(ByVal value As String)
            _Tipo = value
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

    Public Sub SalvarSql(ByRef Sqls As ArrayList)
        Dim Sql As String = ""
        Select Case Me.IUD
            Case "I"
                Sql = "INSERT INTO PedidosXDepositos " & vbCrLf & _
                      "(Empresa_Id, EndEmpresa_Id, Pedido_Id, Deposito_Id, EndDeposito_Id, Quantidade, Tipo, Principal)" & vbCrLf & _
                      " VALUES('" & Me.Pedido.CodigoEmpresa & "'," & Me.Pedido.EnderecoEmpresa.ToString() & "," & vbCrLf & _
                      Me.Pedido.Codigo.ToString() & "," & vbCrLf & _
                      "'" & Me.Codigo & "'," & Me.CodigoEndereco.ToString() & "," & vbCrLf & _
                      Str(Me.Quantidade) & ",'" & Me.Tipo & "'," & IIf(Me.Principal, 1, 0) & ")" & vbCrLf
                Sqls.Add(Sql)
            Case "U"
                Sql = "Update PedidosXDepositos Set" & vbCrLf & _
                      "   Principal =" & IIf(Me.Principal, 1, 0) & vbCrLf & _
                      " WHERE Empresa_Id     ='" & Me.Pedido.CodigoEmpresa & "'" & vbCrLf & _
                      "   AND EndEmpresa_Id  = " & Me.Pedido.EnderecoEmpresa & vbCrLf & _
                      "   AND Pedido_Id      = " & Me.Pedido.Codigo & vbCrLf & _
                      "   AND Deposito_Id    ='" & Me.Codigo & "'" & vbCrLf & _
                      "   AND EndDeposito_Id = " & Me.CodigoEndereco & vbCrLf
                Sqls.Add(Sql)
            Case "D"
                Sql = "DELETE PedidosXDepositos " & _
                      " WHERE Empresa_Id     ='" & Me.Pedido.CodigoEmpresa & "'" & vbCrLf & _
                      "   AND EndEmpresa_Id  = " & Me.Pedido.EnderecoEmpresa & vbCrLf & _
                      "   AND Pedido_Id      = " & Me.Pedido.Codigo & vbCrLf & _
                      "   AND Deposito_Id    ='" & Me.Codigo & "'" & vbCrLf & _
                      "   AND EndDeposito_Id = " & Me.CodigoEndereco & vbCrLf
                Sqls.Add(Sql)
        End Select
    End Sub
#End Region

End Class


