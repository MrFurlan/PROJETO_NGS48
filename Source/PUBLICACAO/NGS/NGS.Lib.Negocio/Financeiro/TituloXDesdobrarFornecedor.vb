Imports Microsoft.VisualBasic
Imports System.Data
Imports NGS.Lib.Uteis

<Serializable()> _
Public Class TituloXDesdobrarFornecedor
    Implements IBaseEntity

#Region "Construtor"
    Public Sub New(ByVal pCodigoTitulo As Integer)
        _CodigoTitulo = pCodigoTitulo
        Carregar()
    End Sub

    Public Sub New(ByVal pTitulo As Titulo)
        _CodigoTitulo = pTitulo.Codigo
        _Titulo = pTitulo
        Carregar()
    End Sub
#End Region

#Region "Fields"
    Private _IUD As String = ""
    Private _CodigoTitulo As Integer
    Private _Titulo As Titulo
    Private _CodigoCliente As String = ""
    Private _EndCliente As Integer
    Private _Cliente As Cliente
    Private _CodigoPedido As Integer
    Private _Pedido As Pedido
    Private _CodigoCarteiraFinanceira As String = ""
    Private _CarteiraFinanceira As CarteiraFinanceira
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

    Public Property CodigoTitulo() As Integer
        Get
            Return _CodigoTitulo
        End Get
        Set(ByVal value As Integer)
            _CodigoTitulo = value
            _Titulo = Nothing
            _Pedido = Nothing
        End Set
    End Property

    Public Property Titulo() As Titulo
        Get
            If _Titulo Is Nothing And _CodigoTitulo > 0 Then _Titulo = New Titulo(_CodigoTitulo)
            Return _Titulo
        End Get
        Set(ByVal value As Titulo)
            _Titulo = value
        End Set
    End Property

    Public Property CodigoCliente() As String
        Get
            Return _CodigoCliente
        End Get
        Set(ByVal value As String)
            _CodigoCliente = value
            _Cliente = Nothing
        End Set
    End Property

    Public Property EndCliente() As Integer
        Get
            Return _EndCliente
        End Get
        Set(ByVal value As Integer)
            _EndCliente = value
            _Cliente = Nothing
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

    Public Property CodigoPedido() As Integer
        Get
            Return _CodigoPedido
        End Get
        Set(ByVal value As Integer)
            _CodigoPedido = value
            _Pedido = Nothing
        End Set
    End Property

    Public Property Pedido() As Pedido
        Get
            If _Pedido Is Nothing And _CodigoPedido > 0 And _CodigoTitulo > 0 Then _Pedido = New Pedido(Titulo.CodigoEmpresa, Titulo.EnderecoEmpresa, _CodigoPedido)
            Return _Pedido
        End Get
        Set(ByVal value As Pedido)
            _Pedido = value
        End Set
    End Property

    Public Property CodigoCarteiraFinanceira() As String
        Get
            Return _CodigoCarteiraFinanceira
        End Get
        Set(ByVal value As String)
            _CodigoCarteiraFinanceira = value
            _CarteiraFinanceira = Nothing
        End Set
    End Property

    Public Property CarteiraFinanceira() As CarteiraFinanceira
        Get
            If _CarteiraFinanceira Is Nothing And _CodigoCarteiraFinanceira.Length > 0 Then _CarteiraFinanceira = New CarteiraFinanceira(_CodigoCarteiraFinanceira)
            Return _CarteiraFinanceira
        End Get
        Set(ByVal value As CarteiraFinanceira)
            _CarteiraFinanceira = value
        End Set
    End Property
#End Region

#Region "Methods"
    Public Sub Carregar()
        Dim sql As String = "SELECT Registro_Id, Cliente, EndCliente, Pedido, Carteira" & vbCrLf & _
                            "  FROM TitulosXDesdobrarFornecedor " & vbCrLf & _
                            " Where Registro_Id = " & _CodigoTitulo
        Dim Banco As New AcessaBanco
        Dim ds As DataSet
        ds = Banco.ConsultaDataSet(sql, "TxDF")
        If ds.Tables(0).Rows.Count = 0 Then Exit Sub
        Dim row As DataRow = ds.Tables(0).Rows(0)

        _CodigoTitulo = row("Registro")
        _CodigoCliente = row("Cliente")
        _EndCliente = row("EndCliente")
        _CodigoPedido = row("Pedido")
        _CodigoCarteiraFinanceira = row("Carteira")
    End Sub

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
                Sql = " INSERT INTO TitulosXDesdobrarFornecedor(Registro_Id, Cliente, EndCliente, Pedido, Carteira) " & vbCrLf & _
                      " VALUES (" & _CodigoTitulo & ",'" & _CodigoCliente & "'," & _EndCliente & ",'" & _CodigoPedido & "','" & _CodigoCarteiraFinanceira & "')"
                Sqls.Add(Sql)
            Case "U"
                Sql = " UPDATE TitulosXDesdobrarFornecedor SET" & vbCrLf & _
                      "    Cliente    ='" & _CodigoCliente & "'" & vbCrLf & _
                      "   ,EndCliente = " & _EndCliente & vbCrLf & _
                      "   ,Pedido     ='" & _CodigoPedido & "'" & vbCrLf & _
                      "   ,Carteira   ='" & _CodigoCarteiraFinanceira & "'" & vbCrLf & _
                      "  WHERE Registro_Id = " & _CodigoTitulo & vbCrLf
                Sqls.Add(Sql)
            Case "D"
                Sql = " DELETE TitulosXDesdobrarFornecedor" & vbCrLf & _
                      "  WHERE Registro_Id = " & _CodigoTitulo & vbCrLf
                Sqls.Add(Sql)
        End Select
    End Sub
#End Region

End Class