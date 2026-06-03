Imports Microsoft.VisualBasic
Imports System.Data
Imports NGS.Lib.Uteis
Imports NGS.Lib.Negocio
Imports System.IO

'***********************************************************************************************************
'*********************************  Lista de Contratos do Pedido  ******************************************
'***********************************************************************************************************
<Serializable()> _
Public Class ListPedidoXContrato
    Inherits List(Of PedidoXContrato)

#Region "Construtor"
    Public Sub New(ByVal pPedido As Negocio.Pedido)
        _Pedido = pPedido
        If _Pedido.Codigo = 0 Then Exit Sub

        Dim Banco As New AcessaBanco
        Try
            Dim SQL As String = ""
            SQL = "SELECT Contrato_Id, Descricao, NomeDoArquivo, Arquivo " & _
                  "  FROM PedidosXContratos " & _
                  " WHERE Empresa_Id    ='" & Me.Pedido.CodigoEmpresa & "' " & _
                  "   AND EndEmpresa_Id = " & Me.Pedido.EnderecoEmpresa.ToString() & " " & _
                  "   AND Pedido_Id     = " & Me.Pedido.Codigo.ToString()

            Dim ds As DataSet = Banco.ConsultaDataSet(SQL, "Contratos")

            For Each row As DataRow In ds.Tables(0).Rows
                Dim pXc As New Negocio.PedidoXContrato(Pedido)

                pXc.Codigo = row("Contrato_Id")
                pXc.Descricao = row("Descricao")
                pXc.NomeDoArquivo = row("NomeDoArquivo")
                pXc.Arquivo = DirectCast(row("Arquivo"), Byte())

                Me.Add(pXc)
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
        For Each pXc As PedidoXContrato In Me
            pXc.SalvarSql(Sqls)
        Next
    End Sub
#End Region

End Class

'***********************************************************************************************************
'**************************************   Contratos do Pedido   ********************************************
'***********************************************************************************************************
<Serializable()> _
Public Class PedidoXContrato

#Region "Contrutor"
    Public Sub New(ByVal pPedido As Negocio.Pedido)
        _Pedido = pPedido
    End Sub
#End Region

#Region "Fields"
    Private _IUD As String = ""

    Private _Pedido As Negocio.Pedido

    Private _Codigo As Integer
    Private _Descricao As String
    Private _NomeDoArquivo As String
    Private _Arquivo As Byte()


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

    Public Property Codigo() As Integer
        Get
            Return _Codigo
        End Get
        Set(ByVal value As Integer)
            _Codigo = value
        End Set
    End Property

    Public Property Arquivo() As Byte()
        Get
            Return _Arquivo
        End Get
        Set(ByVal value As Byte())
            _Arquivo = value
        End Set
    End Property

    Public Property Descricao() As String
        Get
            Return _Descricao
        End Get
        Set(ByVal value As String)
            _Descricao = value
        End Set
    End Property

    Public Property NomeDoArquivo() As String
        Get
            Return _NomeDoArquivo
        End Get
        Set(ByVal value As String)
            _NomeDoArquivo = value
        End Set
    End Property

#End Region

#Region "Methods"

    Private Function BytesToString(ByVal Input As Byte()) As String
        Dim Result As New System.Text.StringBuilder(Input.Length * 2)
        Dim Part As String
        For Each b As Byte In Input
            Part = Conversion.Hex(b)
            If Part.Length = 1 Then Part = "0" & Part
            Result.Append(Part)
        Next
        Return "0x" & Result.ToString()
    End Function

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
                Sql = "INSERT INTO PedidosXContratos " & vbCrLf & _
                      "(Empresa_Id, EndEmpresa_Id, Pedido_Id, Contrato_Id, Descricao, NomeDoArquivo, Arquivo)" & vbCrLf & _
                      " VALUES('" & Me.Pedido.CodigoEmpresa & "'," & Me.Pedido.EnderecoEmpresa.ToString() & "," & vbCrLf & _
                      Me.Pedido.Codigo.ToString() & "," & Me.Codigo & ",'" & Me.Descricao.ToString() & "','" & Me.NomeDoArquivo.ToString() & "'," & BytesToString(Arquivo) & ")" & vbCrLf
                Sqls.Add(Sql)
            Case "U"
                Sql = "Update PedidosXContratos Set" & vbCrLf & _
                      "   Descricao     ='" & Me.Descricao.ToString() & "'" & vbCrLf & _
                      "  ,NomeDoArquivo ='" & Me.NomeDoArquivo.ToString() & "'" & vbCrLf & _
                      "  ,Arquivo       =" & BytesToString(Arquivo) & vbCrLf & _
                      " WHERE Empresa_Id     ='" & Me.Pedido.CodigoEmpresa & "'" & vbCrLf & _
                      "   AND EndEmpresa_Id  = " & Me.Pedido.EnderecoEmpresa & vbCrLf & _
                      "   AND Pedido_Id      = " & Me.Pedido.Codigo & vbCrLf & _
                      "   AND Contrato_Id    = " & Me.Codigo
                Sqls.Add(Sql)
            Case "D"
                Sql = "DELETE PedidosXContratos " & _
                      " WHERE Empresa_Id     ='" & Me.Pedido.CodigoEmpresa & "'" & vbCrLf & _
                      "   AND EndEmpresa_Id  = " & Me.Pedido.EnderecoEmpresa & vbCrLf & _
                      "   AND Pedido_Id      = " & Me.Pedido.Codigo & vbCrLf & _
                      "   AND Contrato_Id    = " & Me.Codigo
                Sqls.Add(Sql)
        End Select
    End Sub
#End Region

End Class


