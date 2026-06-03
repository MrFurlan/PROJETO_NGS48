Imports System.Data
Imports NGS.Lib.Uteis

<Serializable()> _
Public Class ListMetasXSafraXProdutos
    Inherits List(Of MetasXSafraXProdutos)
End Class

<Serializable()> _
Public Class MetasXSafraXProdutos
    Implements IBaseEntity

#Region "Fields"
    Private _IUD As String = ""
    Private _CodigoMeta As Integer
    Private _CodigoSafra As String = ""
    Private _CodigoCultura As Integer
    Private _CodigoProduto As String
    Private _Portifolio As Boolean
    Private _PrecoVenda As Decimal
    Private _Dosagem As Decimal
    Private _NumAplicacao As Integer
    Private _Valor As Decimal
#End Region

#Region "Properties"
    Property IUD As String
        Get
            Return _IUD
        End Get
        Set(ByVal value As String)
            _IUD = value
        End Set
    End Property

    Public Property CodigoMeta() As Integer
        Get
            Return _CodigoMeta
        End Get
        Set(ByVal value As Integer)
            _CodigoMeta = value
        End Set
    End Property

    Public Property CodigoSafra() As String
        Get
            Return _CodigoSafra
        End Get
        Set(ByVal value As String)
            _CodigoSafra = value
        End Set
    End Property

    Public Property CodigoCultura() As Integer
        Get
            Return _CodigoCultura
        End Get
        Set(ByVal value As Integer)
            _CodigoCultura = value
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

    Public Property Portifolio() As Boolean
        Get
            Return _Portifolio
        End Get
        Set(ByVal value As Boolean)
            _Portifolio = value
        End Set
    End Property

    Public Property PrecoVenda() As Decimal
        Get
            Return _PrecoVenda
        End Get
        Set(ByVal value As Decimal)
            _PrecoVenda = value
        End Set
    End Property

    Public Property Dosagem() As Decimal
        Get
            Return _Dosagem
        End Get
        Set(ByVal value As Decimal)
            _Dosagem = value
        End Set
    End Property

    Public Property NumAplicacao() As Integer
        Get
            Return _NumAplicacao
        End Get
        Set(ByVal value As Integer)
            _NumAplicacao = value
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
                Sql = " INSERT INTO MetasXSafrasXProdutos (Meta_Id, Safra_Id, Cultura_Id, Produto_Id, Potifolio, Preco_Venda," & vbCrLf & _
                      "                                         Dosagem, Num_Aplicacao, Valor) " & vbCrLf & _
                      " VALUES (" & Me.CodigoMeta & ", '" & Me.CodigoSafra & "', " & Me.CodigoCultura & ", '" & Me.CodigoProduto & "'," & Me.Portifolio & "," & Me.PrecoVenda & "," & vbCrLf & _
                                    Me.Dosagem & "," & Me.NumAplicacao & "," & Me.Valor & " )" & vbCrLf
                Sqls.Add(Sql)
            Case "U"
                Sql = " UPDATE MetasXSafrasXProdutos SET" & vbCrLf & _
                      "    Portifolio             = " & Me.Portifolio & vbCrLf & _
                      "    Preco_Venda             = " & Me.PrecoVenda & vbCrLf & _
                      "    Dosagem                = " & Me.Dosagem & vbCrLf & _
                      "    Num_Aplicacao           = " & Me.NumAplicacao & vbCrLf & _
                      "    Valor                  = " & Me.Valor & vbCrLf & _
                      "  WHERE      Meta_Id       = " & Me.CodigoMeta & vbCrLf & _
                      "     And     Safra_Id    = " & Str(Me.CodigoSafra) & vbCrLf & _
                      "     And     Cultura_Id    = " & Str(Me.CodigoCultura) & vbCrLf & _
                      "     And     Produto_Id    = " & Str(Me.CodigoProduto) & vbCrLf
                Sqls.Add(Sql)
            Case "D"
                Sql = " DELETE MetasXSafrasXProdutos" & vbCrLf & _
                       "  WHERE      Meta_Id       = " & Me.CodigoMeta & vbCrLf & _
                      "     And     Safra_Id    = " & Str(Me.CodigoSafra) & vbCrLf & _
                      "     And     Cultura_Id    = " & Str(Me.CodigoCultura) & vbCrLf & _
                      "     And     Produto_Id    = " & Str(Me.CodigoProduto) & vbCrLf
                Sqls.Add(Sql)
        End Select
    End Sub
#End Region

End Class
