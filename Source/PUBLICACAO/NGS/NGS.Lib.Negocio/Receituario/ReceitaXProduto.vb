Imports Microsoft.VisualBasic
Imports System.Data
Imports NGS.Lib.Uteis

<Serializable()> _
Public Class ListReceitaXProduto
    Inherits List(Of ReceitaXProduto)

#Region "Contrutor"
    Public Sub New()

    End Sub

    Public Sub New(ByRef pReceita As Receita)
        _Parent = pReceita
        Dim sql As String
        sql = "SELECT Receita_Id, Produto_Id, CulturaPragaFito, FormaDeAplicacao, Quantidade, Dosagem, DosagemRecomendada, UnidadeDeMedida, Vazao, AreaTratada, AreaTotal, ModoAplicacao, " & vbCrLf & _
              "       EpocaAplicacao, IntervaloSeguranca, InstrucaoDeUso, UsuarioInclusao, UsuarioInclusaoData, NumeroDeAplicacao " & vbCrLf & _
              "  FROM ReceitaXProduto " & vbCrLf & _
              " Where Receita_Id = " & pReceita.CodigoReceita & vbCrLf

        Dim ds As DataSet
        Dim Banco As New AcessaBanco
        ds = Banco.ConsultaDataSet(sql, "Produtos")

        For Each row As DataRow In ds.Tables(0).Rows
            Dim RxP As New ReceitaXProduto
            RxP.CodigoReceita = row("Receita_Id")
            RxP.CodigoProduto = row("Produto_Id")
            RxP.CodigoCulturaPragaFito = row("CulturaPragaFito")
            RxP.CodigoFormaDeAplicacao = row("FormaDeAplicacao")
            RxP.Quantidade = row("Quantidade")
            RxP.CodigoDosagem = row("Dosagem")
            RxP.DosagemRecomendada = row("DosagemRecomendada")
            RxP.UnidadeDeMedida = row("UnidadeDeMedida")
            RxP.Vazao = row("Vazao")
            RxP.AreaTratada = row("AreaTratada")
            RxP.AreaTotal = row("AreaTotal")
            RxP.ModoAplicacao = row("ModoAplicacao")
            RxP.EpocaAplicacao = row("EpocaAplicacao")
            RxP.IntervaloDeSeguranca = row("IntervaloSeguranca")
            RxP.InstrucaoDeUso = row("InstrucaoDeUso")
            RxP.UsuarioInclusao = row("UsuarioInclusao")
            RxP.UsuarioInclusaoData = row("UsuarioInclusaoData")
            RxP.NumeroDeAplicacao = row("NumeroDeAplicacao")
            Me.Add(RxP)
        Next
    End Sub
#End Region

#Region "Fields"
    Private _Parent As Receita
#End Region

#Region "Property"
    Public Property Parent() As Receita
        Get
            Return _Parent
        End Get
        Set(ByVal value As Receita)
            _Parent = value
        End Set
    End Property
#End Region

#Region "Methods"
    Public Sub SalvarSql(ByVal Sqls As ArrayList)
        For Each produto As ReceitaXProduto In Me
            If Parent.IUD = "D" Or Parent.IUD = "I" Then produto.IUD = Parent.IUD
            If produto.IUD <> "" Then
                produto.SalvarSql(Sqls)
            End If
        Next
    End Sub
#End Region

End Class

<Serializable()> _
Public Class ReceitaXProduto

#Region "Construtor"
    Public Sub New()

    End Sub

#End Region

#Region "Fields"
    Private _IUD As String
    Private _CodigoReceita As Integer
    Private _Produto As Produto
    Private _CodigoProduto As String
    Private _CodigoCulturaPragaFito As Integer
    Private _CodigoDosagem As Integer
    Private _FormaDeAplicacao As FormaDeAplicacao
    Private _CodigoFormaDeAplicacao As Integer
    Private _Quantidade As Decimal
    Private _DosagemRecomendada As Decimal
    Private _UnidadeDeMedida As String
    Private _Vazao As String
    Private _AreaTratada As Decimal
    Private _AreaTotal As Decimal
    Private _ModoAplicacao As String
    Private _EpocaAplicacao As String
    Private _IntervaloDeSeguranca As String
    Private _InstrucaoDeUso As String
    Private _UsuarioInclusao As String
    Private _UsuarioInclusaoData As DateTime
    Private _NumeroDeAplicacao As Integer
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

    Public Property CodigoReceita() As Integer
        Get
            Return _CodigoReceita
        End Get
        Set(ByVal value As Integer)
            _CodigoReceita = value
        End Set
    End Property

    Public ReadOnly Property Produto() As Produto
        Get
            If _Produto Is Nothing And _CodigoProduto > 0 Then _Produto = New Produto(_CodigoProduto)
            Return _Produto
        End Get
    End Property

    Public Property CodigoProduto() As String
        Get
            Return _CodigoProduto
        End Get
        Set(ByVal value As String)
            _CodigoProduto = value
            _Produto = Nothing
        End Set
    End Property

    Public Property CodigoCulturaPragaFito() As Integer
        Get
            Return _CodigoCulturaPragaFito
        End Get
        Set(ByVal value As Integer)
            _CodigoCulturaPragaFito = value
        End Set
    End Property

    Public Property CodigoDosagem() As Integer
        Get
            Return _CodigoDosagem
        End Get
        Set(ByVal value As Integer)
            _CodigoDosagem = value
        End Set
    End Property

    Public ReadOnly Property FormaDeAplicacao() As FormaDeAplicacao
        Get
            If _FormaDeAplicacao Is Nothing And _CodigoFormaDeAplicacao > 0 Then _FormaDeAplicacao = New FormaDeAplicacao(_CodigoFormaDeAplicacao)
            Return _FormaDeAplicacao
        End Get
    End Property

    Public Property CodigoFormaDeAplicacao() As Integer
        Get
            Return _CodigoFormaDeAplicacao
        End Get
        Set(ByVal value As Integer)
            _CodigoFormaDeAplicacao = value
            _FormaDeAplicacao = Nothing
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

    Public Property DosagemRecomendada() As Decimal
        Get
            Return _DosagemRecomendada
        End Get
        Set(ByVal value As Decimal)
            _DosagemRecomendada = value
        End Set
    End Property

    Public Property UnidadeDeMedida() As String
        Get
            Return _UnidadeDeMedida
        End Get
        Set(ByVal value As String)
            _UnidadeDeMedida = value
        End Set
    End Property

    Public Property Vazao() As String
        Get
            Return _Vazao
        End Get
        Set(ByVal value As String)
            _Vazao = value
        End Set
    End Property

    Public Property AreaTratada() As Decimal
        Get
            Return _AreaTratada
        End Get
        Set(ByVal value As Decimal)
            _AreaTratada = value
        End Set
    End Property

    Public Property AreaTotal() As Decimal
        Get
            Return _AreaTotal
        End Get
        Set(ByVal value As Decimal)
            _AreaTotal = value
        End Set
    End Property

    Public Property ModoAplicacao() As String
        Get
            Return _ModoAplicacao
        End Get
        Set(ByVal value As String)
            _ModoAplicacao = value
        End Set
    End Property

    Public Property EpocaAplicacao() As String
        Get
            Return _EpocaAplicacao
        End Get
        Set(ByVal value As String)
            _EpocaAplicacao = value
        End Set
    End Property

    Public Property IntervaloDeSeguranca() As String
        Get
            Return _IntervaloDeSeguranca
        End Get
        Set(ByVal value As String)
            _IntervaloDeSeguranca = value
        End Set
    End Property

    Public Property InstrucaoDeUso() As String
        Get
            Return _InstrucaoDeUso
        End Get
        Set(ByVal value As String)
            _InstrucaoDeUso = value
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

    Public Property UsuarioInclusaoData() As DateTime
        Get
            Return _UsuarioInclusaoData
        End Get
        Set(ByVal value As DateTime)
            _UsuarioInclusaoData = value
        End Set
    End Property

    Public Property NumeroDeAplicacao() As Integer
        Get
            Return _NumeroDeAplicacao
        End Get
        Set(ByVal value As Integer)
            _NumeroDeAplicacao = value
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
                Sql = " INSERT INTO ReceitaXProduto (Receita_Id, Produto_Id, CulturaPragaFito, FormaDeAplicacao, Quantidade, Dosagem, UnidadeDeMedida, Vazao, AreaTratada, AreaTotal, ModoAplicacao, EpocaAplicacao, IntervaloSeguranca, InstrucaoDeUso, UsuarioInclusao, UsuarioInclusaoData, DosagemRecomendada, NumeroDeAplicacao) " & vbCrLf & _
                      " VALUES (" & _CodigoReceita & ",'" & _CodigoProduto & "'," & _CodigoCulturaPragaFito & "," & _CodigoFormaDeAplicacao & "," & Str(_Quantidade) & "," & _CodigoDosagem & ",'" & Left(_UnidadeDeMedida.ToString.Replace("'", ""), 20) & "','" & Left(_Vazao, 50) & "'," & Str(_AreaTratada) & "," & Str(_AreaTotal) & ",'" & _ModoAplicacao & "','" & _EpocaAplicacao & "','" & Left(_IntervaloDeSeguranca, 50) & "','" & _InstrucaoDeUso & "','" & _UsuarioInclusao & "','" & _UsuarioInclusaoData.ToString("yyyy-MM-dd") & "'," & Str(_DosagemRecomendada) & "," & _NumeroDeAplicacao & ")"
                Sqls.Add(Sql)
            Case "U"
                Sql = " UPDATE ReceitaXProduto SET" & vbCrLf & _
                      "    CulturaPragaFito   = " & _CodigoCulturaPragaFito & vbCrLf & _
                      "   ,FormaDeAplicacao   = " & _CodigoFormaDeAplicacao & vbCrLf & _
                      "   ,Quantidade         = " & Str(_Quantidade) & vbCrLf & _
                      "   ,Dosagem            = " & _CodigoDosagem & vbCrLf & _
                      "   ,UnidadeDeMedida    ='" & Left(_UnidadeDeMedida, 20) & "'" & vbCrLf & _
                      "   ,Vazao              ='" & Left(_Vazao, 50) & "'" & vbCrLf & _
                      "   ,AreaTratada        = " & Str(_AreaTratada) & vbCrLf & _
                      "   ,AreaTotal          = " & Str(_AreaTotal) & vbCrLf & _
                      "   ,ModoAplicacao      ='" & _ModoAplicacao & "'" & vbCrLf & _
                      "   ,EpocaAplicacao     ='" & _EpocaAplicacao & "'" & vbCrLf & _
                      "   ,IntervaloSeguranca ='" & Left(_IntervaloDeSeguranca, 50) & "'" & vbCrLf & _
                      "   ,InstrucaoDeUso     ='" & _InstrucaoDeUso & "'" & vbCrLf & _
                      "   ,UsuarioInclusao    ='" & _UsuarioInclusao & "'" & vbCrLf & _
                      "   ,UsuarioInclusaoData='" & _UsuarioInclusaoData.ToString("yyyy-MM-dd") & "'" & vbCrLf & _
                      "   ,DosagemRecomendada = " & Str(_DosagemRecomendada) & vbCrLf & _
                      "   ,NumeroDeAplicacao = " & _NumeroDeAplicacao & vbCrLf & _
                      "  WHERE Receita_Id =" & _CodigoReceita & "" & vbCrLf & _
                      "    AND Produto_Id ='" & _CodigoProduto & "'"
                Sqls.Add(Sql)
            Case "D"
                Sql = " DELETE ReceitaXProduto" & vbCrLf & _
                      "  WHERE Receita_Id =" & _CodigoReceita & "" & vbCrLf & _
                      "    AND Produto_Id ='" & _CodigoProduto & "'"
                Sqls.Add(Sql)
        End Select
    End Sub
#End Region

End Class