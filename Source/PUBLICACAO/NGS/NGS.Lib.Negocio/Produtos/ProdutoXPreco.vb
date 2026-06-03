Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Collections.Generic
Imports NGS.Lib.Uteis
Imports System.Web

<Serializable()>
Public Class ListProdutoXPreco
    Inherits List(Of ProdutoXPreco)

#Region "Construtor"

    Public Sub New()

    End Sub

    Public Sub New(Optional ByVal tabelaDePreco As String = "", Optional ByVal codigoCliente As String = "", Optional ByVal enderecoCliente As String = "", Optional ByVal Produto As String = "", Optional ByVal Movimento As String = "", Optional ByVal pOrderby As String = "")
        Dim objBanco As New AcessaBanco()
        Dim strSQL As String = "SELECT Tabela_Id, Cliente_Id, EndCliente_Id, Produto_Id, Data_Id, Moeda, Valor, isnull(MargemMenor,0) as MargemMenor," & vbCrLf &
                               "isnull(MargemMaior,0) as MargemMaior, isnull(FixoOperacional,0) as FixoOperacional " & vbCrLf &
                               "  FROM ProdutosXPrecos " & vbCrLf &
                               " WHERE 1=1"

        If tabelaDePreco.Length > 0 Then
            strSQL &= " AND Tabela_Id    = " & tabelaDePreco
        End If

        If codigoCliente.Length > 0 Then
            strSQL &= "AND Cliente_Id    = '" & codigoCliente.ToString() & "'" & vbCrLf &
                      "AND EndCliente_id = " & enderecoCliente & vbCrLf
        End If

        If Produto.Length > 0 Then
            strSQL &= "AND Produto_Id = '" & Produto.ToString() & "'"

        End If

        If Movimento.Length > 0 Then
            strSQL &= "AND Data_Id >='" & Movimento.ToString() & "'"
        End If

        If pOrderby.Length > 0 Then
            strSQL &= " Order by " & pOrderby
        End If

        Dim ds As DataSet = objBanco.ConsultaDataSet(strSQL, "ProdutosXPrecos")

        For Each dr As DataRow In ds.Tables(0).Rows
            Dim PxP As New ProdutoXPreco
            PxP.CodigoTabelaDePreco = dr("Tabela_Id")
            PxP.CodigoCliente = dr("Cliente_Id")
            PxP.EnderecoCliente = dr("EndCliente_Id")
            PxP.CodigoProduto = dr("Produto_Id")
            PxP.Movimento = dr("Data_Id")
            PxP.CodigoMoeda = dr("Moeda")
            PxP.Valor = dr("Valor")
            PxP.MargemMenor = dr("MargemMenor")
            PxP.MargemMaior = dr("MargemMaior")
            PxP.FixoOperacional = dr("FixoOperacional")
            Me.Add(PxP)
        Next
    End Sub
#End Region

#Region "Fields"

#End Region

#Region "Property"

#End Region

#Region "Methods"

#End Region

End Class

<Serializable()>
Public Class ProdutoXPreco

#Region "Construtor"
    Public Sub New()

    End Sub

    Public Sub New(ByVal tabelaDePreco As Integer, ByVal codigoCliente As String, ByVal enderecoCliente As Integer, ByVal Produto As String, ByVal Movimento As String)
        Dim objBanco As New AcessaBanco()

        Dim Sql As String = "SELECT Tabela_Id, Cliente_Id, EndCliente_Id, Produto_Id, Data_Id, Moeda, Valor, isnull(MargemMenor,0) as MargemMenor, isnull(MargemMaior,0) as MargemMaior, isnull(FixoOperacional,0) as FixoOperacional " & vbCrLf &
                            "  FROM ProdutosXPrecos " &
                            " WHERE Tabela_Id     = " & tabelaDePreco & vbCrLf &
                            "   AND Cliente_Id    = '" & codigoCliente & "'" & vbCrLf &
                            "   AND EndCliente_id = " & enderecoCliente & vbCrLf &
                            "   AND Produto_Id    = '" & Produto & "'" & vbCrLf &
                            "   AND Data_id       = '" & Movimento & "'"

        Dim ds As DataSet = objBanco.ConsultaDataSet(Sql, "ProdutosXPrecos")

        For Each dr As DataRow In ds.Tables(0).Rows
            _CodigoTabelaDePreco = dr("Tabela_Id")
            _CodigoCliente = dr("Cliente_id")
            _EnderecoCliente = dr("EndCliente_Id")
            _CodigoProduto = dr("Produto_Id")
            _Movimento = dr("Data_Id")
            _CodigoMoeda = dr("Moeda")
            _Valor = dr("Valor")
            _MargemMenor = dr("MargemMenor")
            _MargemMaior = dr("MargemMaior")
            _FixoOperacional = dr("FixoOperacional")
        Next
    End Sub
#End Region

#Region "Fields"
    Private _IUD As String = ""
    Private _CodigoProduto As String = ""
    Private _Produto As Produto
    Private _Movimento As DateTime = Date.Today
    Private _CodigoMoeda As Integer = 0
    Private _Moeda As Moeda
    Private _Valor As Decimal
    Private _NomeProduto As String
    Private _MargemMenor As Decimal
    Private _MargemMaior As Decimal
    Private _FixoOperacional As Decimal
    Private _CodigoCliente As String
    Private _EnderecoCliente As Integer
    Private _Cliente As Cliente
    Private _CodigoTabelaDePreco As Integer
    Private _TabelaDePreco As TabelaDePreco

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

    Public Property CodigoProduto() As String
        Get
            Return _CodigoProduto
        End Get
        Set(ByVal value As String)
            _CodigoProduto = value
        End Set
    End Property

    Public Property Produto() As Produto
        Get
            If _Produto Is Nothing And Not Me.CodigoProduto Is Nothing Then _Produto = New Produto(_CodigoProduto)
            Return _Produto
        End Get
        Set(ByVal value As Produto)
            _Produto = value
        End Set
    End Property

    Public Property Movimento() As DateTime
        Get
            Return _Movimento
        End Get
        Set(ByVal value As DateTime)
            _Movimento = value
        End Set
    End Property

    Public Property CodigoMoeda() As Integer
        Get
            Return _CodigoMoeda
        End Get
        Set(ByVal value As Integer)
            _CodigoMoeda = value
        End Set
    End Property

    Public Property Moeda() As Moeda
        Get
            If _Moeda Is Nothing And Me.CodigoMoeda > 0 Then _Moeda = New Moeda(_CodigoMoeda)
            Return _Moeda
        End Get
        Set(ByVal value As Moeda)
            _Moeda = value
        End Set
    End Property

    Public Property Valor() As Decimal
        Get
            Return _Valor
        End Get
        Set(ByVal value As Decimal)
            _Valor = value
        End Set
    End Property

    Public ReadOnly Property NomeProduto() As String
        Get
            If CodigoProduto Is Nothing Then Return ""
            Return Produto.Nome
        End Get
    End Property

    Public ReadOnly Property NomeMoeda() As String
        Get
            If CodigoMoeda = 0 Then Return ""
            Return Moeda.Descricao
        End Get
    End Property

    Public Property MargemMenor() As Decimal
        Get
            Return _MargemMenor
        End Get
        Set(ByVal value As Decimal)
            _MargemMenor = value
        End Set
    End Property

    Public Property MargemMaior() As Decimal
        Get
            Return _MargemMaior
        End Get
        Set(ByVal value As Decimal)
            _MargemMaior = value
        End Set
    End Property

    Public Property FixoOperacional() As Decimal
        Get
            Return _FixoOperacional
        End Get
        Set(ByVal value As Decimal)
            _FixoOperacional = value
        End Set
    End Property

    Public ReadOnly Property ValorMargemMenor() As Decimal
        Get
            Return Math.Round((_Valor * (1 + _FixoOperacional / 100)) * (1 + _MargemMenor / 100), 6)
        End Get
    End Property

    Public ReadOnly Property ValorMargemMaior() As Decimal
        Get
            Return Math.Round((_Valor * (1 + _FixoOperacional / 100)) * (1 + _MargemMaior / 100), 6)
        End Get
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

    Public Property Cliente() As Cliente
        Get
            If _Cliente Is Nothing And Not Me.CodigoCliente Is Nothing Then _Cliente = New Cliente(_CodigoCliente, _EnderecoCliente)
            Return _Cliente
        End Get
        Set(ByVal value As Cliente)
            _Cliente = value
        End Set
    End Property

    Public Property CodigoTabelaDePreco() As Integer
        Get
            Return _CodigoTabelaDePreco
        End Get
        Set(ByVal value As Integer)
            _CodigoTabelaDePreco = value
        End Set
    End Property

    Public Property TabelaDePreco() As TabelaDePreco
        Get
            If _TabelaDePreco Is Nothing And Me.CodigoTabelaDePreco > 0 Then _TabelaDePreco = New TabelaDePreco(_CodigoMoeda)
            Return _TabelaDePreco
        End Get
        Set(ByVal value As TabelaDePreco)
            _TabelaDePreco = value
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
                Sql = "Insert Into ProdutosXPrecos(Tabela_Id, Cliente_Id, EndCliente_Id, Produto_Id, Data_Id, Moeda, Valor, MargemMenor, MargemMaior, FixoOperacional, UsuarioInclusao, UsuarioInclusaoData)" & vbCrLf &
                      " values(" & _CodigoTabelaDePreco & ",'" & _CodigoCliente & "'," & _EnderecoCliente & ",'" & _CodigoProduto & "','" & _Movimento.ToString("yyyy-MM-dd") & "'," & _CodigoMoeda & "," & Str(_Valor) & "," & Str(_MargemMenor) & "," & Str(_MargemMaior) & "," & Str(_FixoOperacional) & ",'" & HttpContext.Current.Session("ssNomeUsuario") & "', " & "'" & Now().ToString("yyyy-MM-dd") & "')"
                Sqls.Add(Sql)
            Case "U"
                Sql = "Update ProdutosXPrecos Set  " & vbCrLf &
                      "  Moeda           = " & _CodigoMoeda & vbCrLf &
                      ", Valor           = " & Str(_Valor) & vbCrLf &
                      ", MargemMenor     = " & Str(_MargemMenor) & vbCrLf &
                      ", MargemMaior     = " & Str(_MargemMaior) & vbCrLf &
                      ", FixoOperacional = " & Str(_FixoOperacional) & vbCrLf &
                      " Where Tabela_Id     = " & _CodigoTabelaDePreco & vbCrLf &
                      "   and Cliente_Id    = '" & _CodigoCliente & "'" & vbCrLf &
                      "   and EndCliente_Id =" & _EnderecoCliente & vbCrLf &
                      "   and Produto_Id    = '" & _CodigoProduto & "'" & vbCrLf &
                      "   and Data_Id       ='" & _Movimento.ToString("yyyy-MM-dd") & "'" & vbCrLf &
                      "   and UsuarioAlteracao       ='" & HttpContext.Current.Session("ssNomeUsuario") & "'" & vbCrLf &
                      "   and UsuarioAlteracaoData       ='" & Now().ToString("yyyy-MM-dd") & "'"
                Sqls.Add(Sql)
            Case "D"
                Sql = "Delete ProdutosXPrecos " & vbCrLf &
                      " Where Tabela_Id     = " & _CodigoProduto & vbCrLf &
                      "   and Cliente_Id    = '" & _CodigoCliente & "'" & vbCrLf &
                      "   and EndCliente_Id = " & _EnderecoCliente & vbCrLf &
                      "   and Produto_Id    = '" & _CodigoProduto & "'" & vbCrLf &
                      "   and Data_Id       = '" & _Movimento.ToString("yyyy-MM-dd") & "'"
                Sqls.Add(Sql)
        End Select
    End Sub
#End Region

End Class