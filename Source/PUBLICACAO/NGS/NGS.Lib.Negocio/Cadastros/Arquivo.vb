Imports Microsoft.VisualBasic
Imports System.Data
Imports NGS.Lib.Uteis
Imports System.Data.SqlClient


<Serializable()>
Public Class ListArquivo
    Inherits List(Of Arquivo)
    Implements IBaseEntity

    Public Sub New()

    End Sub

    Public Sub New(ByVal NotaFiscal As NotaFiscal)
        Dim Banco As New AcessaBanco
        Dim ds As New DataSet
        Dim sql As String = " SELECT Empresa_Id, EndEmpresa_Id, Cliente_Id, EndCliente_Id, Nota_Id, Serie_Id, Arquivo_Id, Pedido_Id, Descricao, Arquivo " & vbCrLf &
                            " FROM VW_Documentos " & vbCrLf &
                            " WHERE Empresa_Id  = '" & NotaFiscal.CodigoEmpresa & "'" & vbCrLf &
                            " AND EndEmpresa_Id = " & NotaFiscal.EnderecoEmpresa & vbCrLf &
                            " AND Cliente_Id    = '" & NotaFiscal.CodigoCliente & "'" & vbCrLf &
                            " AND EndCliente_Id = " & NotaFiscal.EnderecoCliente & vbCrLf &
                            " AND Nota_Id       = " & NotaFiscal.Codigo & vbCrLf &
                            " AND Serie_Id      = '" & NotaFiscal.Serie & "'"

        ds = Banco.ConsultaDataSet(sql, "Arquivo")

        For Each row As DataRow In ds.Tables("Arquivo").Rows
            Dim Arquivo As New Arquivo
            Arquivo.IUD = "U"

            Arquivo.CodigoEmpresa = row("Empresa_Id")
            Arquivo.EnderecoEmpresa = row("EndEmpresa_Id")
            Arquivo.CodigoCliente = row("Cliente_ID")
            Arquivo.EnderecoCliente = row("EndCliente_ID")
            Arquivo.CodigoNota = row("Nota_ID")
            Arquivo.Serie = row("Serie_ID")
            Arquivo.Codigo = row("Arquivo_Id").ToString
            Arquivo.CodigoPedido = row("Pedido_Id")
            Arquivo.Descricao = row("Descricao")
            Arquivo.Arquivo = row("Arquivo")

            Me.Add(Arquivo)
        Next
    End Sub

    Public Sub New(ByVal Pedido As Pedido)

    End Sub

End Class

<Serializable()>
Public Class Arquivo
    Implements IBaseEntity

    Public Sub New()

    End Sub

#Region "Fields"

    Private _IUD As String
    Private _Codigo As String
    Private _Descricao As String
    Private _Arquivo As Byte()
    Private _CodigoEmpresa As String
    Private _EnderecoEmpresa As Integer
    Private _CodigoCliente As String
    Private _EnderecoCliente As Integer
    Private _CodigoNota As Integer
    Private _Serie As String
    Private _CodigoPedido As Integer

    'Arquivo em String
    Private _ArquivoTexto As String

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

    Public Property Codigo() As String
        Get
            Return _Codigo
        End Get
        Set(ByVal value As String)
            _Codigo = value
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

    Public Property Arquivo() As Byte()
        Get
            Return _Arquivo
        End Get
        Set(ByVal value As Byte())
            _Arquivo = value
        End Set
    End Property

    Public Property ArquivoTexto() As String
        Get
            Return _ArquivoTexto
        End Get
        Set(ByVal value As String)
            _ArquivoTexto = value
        End Set
    End Property

    Public Property CodigoEmpresa() As String
        Get
            Return _CodigoEmpresa
        End Get
        Set(ByVal value As String)
            _CodigoEmpresa = value
        End Set
    End Property

    Public Property EnderecoEmpresa() As Integer
        Get
            Return _EnderecoEmpresa
        End Get
        Set(ByVal value As Integer)
            _EnderecoEmpresa = value
        End Set
    End Property

    Public Property CodigoCliente() As String
        Get
            Return _CodigoCliente
        End Get
        Set(ByVal value As String)
            _CodigoCliente = value
        End Set
    End Property

    Public Property EnderecoCliente() As Integer
        Get
            Return _EnderecoCliente
        End Get
        Set(ByVal value As Integer)
            _EnderecoCliente = value
        End Set
    End Property

    Public Property CodigoNota() As Integer
        Get
            Return _CodigoNota
        End Get
        Set(ByVal value As Integer)
            _CodigoNota = value
        End Set
    End Property

    Public Property Serie() As String
        Get
            Return _Serie
        End Get
        Set(ByVal value As String)
            _Serie = value
        End Set
    End Property

    Public Property CodigoPedido() As Integer
        Get
            Return _CodigoPedido
        End Get
        Set(ByVal value As Integer)
            _CodigoPedido = value
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
        Dim Sql As String = String.Empty
        Select Case Me.IUD
            Case "I"
                Sql = " INSERT INTO VW_Documentos (Empresa_Id, EndEmpresa_Id, Cliente_Id, EndCliente_Id, Nota_Id, Serie_Id, Pedido_Id, Descricao, Arquivo) " & vbCrLf &
                      "              VALUES ('" & Me.CodigoEmpresa & "', " & Me.EnderecoEmpresa & ",'" & Me.CodigoCliente & "'," & Me.EnderecoCliente & "," & Me.CodigoNota & ",'" & Me.Serie & "'," & Me.CodigoPedido & ",'" & Me.Descricao & "'," & BytesToString(Me.Arquivo) & ")" & vbCrLf
            Case "U"
                Sql = " UPDATE VW_Documentos " & vbCrLf &
                      "    SET Empresa_Id    = '" & Me.CodigoEmpresa & "'" & vbCrLf &
                      "        ,EndEmpresa_Id =  " & Me.EnderecoEmpresa & vbCrLf &
                      "        ,Cliente_Id    = '" & Me.CodigoCliente & "'" & vbCrLf &
                      "        ,EndCliente_Id =  " & Me.EnderecoCliente & vbCrLf &
                      "        ,Nota_Id       =  " & Me.CodigoNota & vbCrLf &
                      "        ,Serie_Id      = '" & Me.Serie & "'" & vbCrLf &
                      "        ,Pedido_Id     =  " & Me.CodigoPedido & vbCrLf &
                      " WHERE Arquivo_Id      = '" & Me.Codigo & "'" & vbCrLf
            Case "D"
                Sql = " DELETE VW_Documentos " & vbCrLf &
                      "  WHERE Arquivo_Id = '" & Me.Codigo & "'" & vbCrLf
        End Select
        If Not String.IsNullOrWhiteSpace(Sql) Then Sqls.Add(Sql)
    End Sub

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
#End Region
End Class
