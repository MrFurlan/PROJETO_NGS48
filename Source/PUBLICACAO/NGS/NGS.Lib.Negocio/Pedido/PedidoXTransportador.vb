Imports Microsoft.VisualBasic
Imports System.Data
Imports NGS.Lib.Uteis

'*****************************************************************************************************************
'************************************ List Transportadores do Pedido   *******************************************
'*****************************************************************************************************************
Public Class ListPedidoXTransportador
    Inherits List(Of PedidoXTransportador)

#Region "Construtor"
    Public Sub New(ByVal pPedido As Pedido)
        _Pedido = pPedido
        If Me._Pedido.Codigo = 0 Then Exit Sub

        Dim Banco As New AcessaBanco()

        Try
            Dim sql As String
            sql = "SELECT Transportador_Id, EndTransportador_Id, Quantidade, QuotaDiaria, Redespacho, DataFrete_Id, isnull(UnitarioFrete,0) AS UnitarioFrete " & vbCrLf & _
                  "  FROM PedidosXTransportadores " & vbCrLf & _
                  " WHERE Empresa_Id    ='" & Me._Pedido.CodigoEmpresa & "'" & vbCrLf & _
                  "   AND EndEmpresa_Id = " & Me._Pedido.EnderecoEmpresa & vbCrLf & _
                  "   AND Pedido_Id     = " & Me._Pedido.Codigo & vbCrLf

            Dim ds As DataSet = Banco.ConsultaDataSet(sql, "Comissoes")

            For Each row As DataRow In ds.Tables(0).Rows
                Dim tra As New PedidoXTransportador(Me._Pedido)

                tra.Codigo = row("Transportador_Id").ToString()
                tra.CodigoEndereco = row("EndTransportador_Id")
                tra.Quantidade = row("Quantidade")
                tra.QuotaDiaria = row("QuotaDiaria")
                tra.Redespacho = row("Redespacho").ToString() = "S"
                tra.DataFrete = row("DataFrete_Id")
                tra.UnitarioFrete = row("UnitarioFrete")

                Me.Add(tra)
            Next
        Catch ex As Exception
        Finally
            Banco = Nothing
        End Try

    End Sub
#End Region

#Region "Fields"
    Private _Pedido As Pedido
#End Region

#Region "Property"
    Public ReadOnly Property Pedido As Pedido
        Get
            Return _Pedido
        End Get
    End Property
#End Region
    
#Region "Methods"
    Public Sub SalvarSql(ByVal Sqls As ArrayList)
        For Each PxT As PedidoXTransportador In Me
            If _Pedido.IUD.Contains("D") Or _Pedido.IUD = "I" Then PxT.IUD = _Pedido.IUD
            PxT.SalvarSql(Sqls)
        Next
    End Sub
#End Region

End Class

'*****************************************************************************************************************
'**************************************    Transportadores do Pedido    ******************************************
'*****************************************************************************************************************
Public Class PedidoXTransportador
#Region "Construtor"
    Public Sub New(ByVal pPedido As Pedido)
        _Pedido = pPedido
    End Sub
#End Region

#Region "Fields"
    Private _IUD As String = ""
    Private _Pedido As Pedido
    Private _Codigo As String
    Private _CodigoEndereco As Integer
    Private _Transportador As Cliente
    Private _Quantidade As Decimal
    Private _QuotaDiaria As Decimal
    Private _Redespacho As Boolean = False
    Private _DataFrete As DateTime
    Private _UnitarioFrete As Decimal
    Private _UsuarioInclusao As String
    Private _UsuarioInclusaoData As DateTime
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

    Public ReadOnly Property Pedido() As Pedido
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
            Return _CodigoEndereco
        End Get
        Set(ByVal value As Integer)
            _CodigoEndereco = value
        End Set
    End Property

    Public Property Transportador() As Cliente
        Get
            If _Transportador Is Nothing And Not Codigo Is Nothing Then _Transportador = New Cliente(Me.Codigo, Me.CodigoEndereco)
            Return _Transportador
        End Get
        Set(ByVal value As Cliente)
            _Transportador = value
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

    Public Property QuotaDiaria() As Decimal
        Get
            Return _QuotaDiaria
        End Get
        Set(ByVal value As Decimal)
            _QuotaDiaria = value
        End Set
    End Property

    Public Property Redespacho() As Boolean
        Get
            Return _Redespacho
        End Get
        Set(ByVal value As Boolean)
            _Redespacho = value
        End Set
    End Property

    Public Property DataFrete() As DateTime
        Get
            Return _DataFrete
        End Get
        Set(ByVal value As DateTime)
            _DataFrete = value
        End Set
    End Property

    Public Property UnitarioFrete() As Decimal
        Get
            Return _UnitarioFrete
        End Get
        Set(ByVal value As Decimal)
            _UnitarioFrete = value
        End Set
    End Property

    Public Property UsuarioInclusao() As String
        Get
            Return _UsuarioInclusao
        End Get
        Set(ByVal value As String)
            _UsuarioInclusao = value
        End Set
    End Property

    Public Property DataInclusao() As DateTime
        Get
            Return _UsuarioInclusaoData
        End Get
        Set(ByVal value As DateTime)
            _UsuarioInclusaoData = value
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
                sql = "INSERT INTO PedidosXTransportadores " & vbCrLf & _
                      "        (Empresa_Id, EndEmpresa_Id, Pedido_Id, Transportador_Id, EndTransportador_Id, DataFrete_Id, UnitarioFrete, " & vbCrLf & _
                      "         Quantidade, QuotaDiaria, Redespacho, UsuarioInclusao, UsuarioInclusaoData) " & vbCrLf & _
                      " VALUES ('" & Me.Pedido.CodigoEmpresa & "', " & Me.Pedido.EnderecoEmpresa & "," & vbCrLf & _
                      Me.Pedido.Codigo & "," & vbCrLf & _
                      "'" & Me.Codigo & "', " & Me.CodigoEndereco.ToString() & "," & vbCrLf & _
                      "'" & Me.DataFrete.ToString("yyyy-MM-dd HH:mm:ss") & "', " & Str(Me.UnitarioFrete) & "," & vbCrLf & _
                      Str(Me.Quantidade) & "," & vbCrLf & _
                      Str(Me.QuotaDiaria) & ", '" & IIf(Me.Redespacho, "S", "N") & "'," & vbCrLf & _
                      "'" & Me.UsuarioInclusao & "', CONVERT(varchar, CURRENT_TIMESTAMP, 112))"
                Sqls.Add(sql)
            Case "U"
                sql = "Update PedidosXTransportadores Set" & _
                      "    UnitarioFrete = " & Str(Me.UnitarioFrete) & vbCrLf & _
                      "   ,Quantidade    = " & Str(Me.Quantidade) & vbCrLf & _
                      "   ,QuotaDiaria   = " & Str(Me.QuotaDiaria) & vbCrLf & _
                      "   ,Redespacho    ='" & IIf(Me.Redespacho, "S", "N") & "'" & vbCrLf & _
                      " WHERE Empresa_Id          ='" & Me.Pedido.CodigoEmpresa & "' " & _
                      "   AND EndEmpresa_Id       = " & Me.Pedido.EnderecoEmpresa.ToString() & " " & _
                      "   AND Pedido_Id           = " & Me.Pedido.Codigo.ToString() & " " & _
                      "   AND Transportador_Id    ='" & Me.Codigo & "' " & _
                      "   AND EndTransportador_Id = " & Me.CodigoEndereco.ToString() & _
                      "   AND DataFrete_Id        ='" & Me.DataFrete.ToString("yyyy-MM-dd HH:mm:ss") & "'"
                Sqls.Add(sql)
            Case "D"
                sql = "DELETE PedidosXTransportadores " & _
                      "WHERE Empresa_Id          ='" & Me.Pedido.CodigoEmpresa & "' " & _
                      "  AND EndEmpresa_Id       = " & Me.Pedido.EnderecoEmpresa.ToString() & " " & _
                      "  AND Pedido_Id           = " & Me.Pedido.Codigo.ToString() & " " & _
                      "  AND Transportador_Id    ='" & Me.Codigo & "' " & _
                      "  AND EndTransportador_Id = " & Me.CodigoEndereco.ToString() & _
                      "  AND DataFrete_Id        ='" & Me.DataFrete.ToString("yyyy-MM-dd HH:mm:ss") & "'"
                Sqls.Add(sql)
        End Select
    End Sub
#End Region
End Class

