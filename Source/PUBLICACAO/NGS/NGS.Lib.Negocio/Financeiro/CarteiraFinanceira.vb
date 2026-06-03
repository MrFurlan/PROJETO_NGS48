Imports Microsoft.VisualBasic
Imports System.Collections.Generic
Imports System.Data
Imports NGS.Lib.Uteis
Imports System.Web


<Serializable()> _
Public Class ListCarteiraFinanceira
    Inherits List(Of CarteiraFinanceira)

#Region "Contrutor"
    Public Sub New(Optional ByVal Carregar As Boolean = False, Optional ByVal pWhere As String = "")
        If Not Carregar Then Exit Sub

        Dim Banco As New AcessaBanco
        Dim sql As String
        sql = " SELECT Produto_Id, Descricao, Situacao, ContaClientes, ContaDescontos, ContaDeducoes, ContaJuros, ContaAcrescimos, Classificacao, Isnull(UsuarioInclusao,'') as UsuarioInclusao, " & vbCrLf & _
              "        isnull(UsuarioInclusaoData, CONVERT(varchar, CURRENT_TIMESTAMP, 112)) as UsuarioInclusaoData, isnull(UsuarioAlteracao,'') AS UsuarioAlteracao, isnull(UsuarioAlteracaoData," & vbCrLf & _
              "        CONVERT(varchar, CURRENT_TIMESTAMP, 112)) AS UsuarioAlteracaoData, isnull(Adiantamento,'N') as Adiantamento," & vbCrLf & _
              "        isnull(Pedido,0) as Pedido, isnull(BaixaAdiantamento,0) as BaixaAdiantamento, isnull(CarteiraBaixaAdiantamento,'') as CarteiraBaixaAdiantamento " & vbCrLf & _
              "   FROM ComprasXProdutos" & vbCrLf

        If pWhere.Length > 0 Then
            sql &= " WHERE " & pWhere & vbCrLf
        End If

        Dim ds As DataSet = Banco.ConsultaDataSet(sql, "Carteira")

        For Each row As DataRow In ds.Tables(0).Rows
            Dim CF As New CarteiraFinanceira
            CF.CodigoCarteira = row("Produto_Id")
            CF.Descricao = row("Descricao")
            CF.CodigoSituacao = row("Situacao")
            CF.isAdiantamento = row("Adiantamento") = "S"
            CF.CodigoContaCliente = row("ContaClientes")
            CF.CodigoContaDesconto = row("ContaDescontos")
            CF.CodigoContaDeducao = row("ContaDeducoes")
            CF.CodigoContaJuro = row("ContaJuros")
            CF.CodigoContaAcrescimo = row("ContaAcrescimos")
            CF.Classificacao = row("Classificacao")
            CF.TemPedido = row("Pedido")
            CF.BaixaAdiantamento = row("BaixaAdiantamento")
            CF.CodigoCarteiraBaixaAdiantamento = row("CarteiraBaixaAdiantamento")
            CF.UsuarioInclusao = row("UsuarioInclusao")
            CF.UsuarioInclusaoData = row("UsuarioInclusaoData")
            CF.UsuarioAlteracao = row("UsuarioAlteracao")
            CF.UsuarioAlteracaoData = row("UsuarioAlteracaoData")
            Me.Add(CF)
        Next
    End Sub
#End Region

End Class

<Serializable()> _
Public Class CarteiraFinanceira
    Implements IBaseEntity

#Region "Contrutor"
    Public Sub New()

    End Sub

    Public Sub New(ByVal pProduto As String, Optional ByVal pWhere As String = "")
        Dim Banco As New AcessaBanco
        Dim sql As String
        sql = "SELECT Produto_Id, Descricao, Situacao, ContaClientes, ContaDescontos, ContaDeducoes, ContaJuros, ContaAcrescimos, Classificacao, isnull(UsuarioInclusao,'') AS UsuarioInclusao, " & vbCrLf & _
              "       isnull(UsuarioInclusaoData,CONVERT(varchar, CURRENT_TIMESTAMP, 112)) AS UsuarioInclusaoData, isnull(UsuarioAlteracao,'') AS UsuarioAlteracao, isnull(UsuarioAlteracaoData,CONVERT(varchar, CURRENT_TIMESTAMP, 112)) AS UsuarioAlteracaoData, isnull(Adiantamento,'N') as Adiantamento," & vbCrLf & _
              "       isnull(Pedido,0) as Pedido, isnull(BaixaAdiantamento,0) as BaixaAdiantamento, isnull(CarteiraBaixaAdiantamento,'') as CarteiraBaixaAdiantamento " & vbCrLf & _
              "  FROM ComprasXProdutos" & vbCrLf & _
              " Where Produto_id = '" & pProduto & "'"

        If pWhere.Length > 0 Then
            sql &= " AND " & pWhere & vbCrLf
        End If

        Dim ds As DataSet = Banco.ConsultaDataSet(sql, "Carteira")

        With ds.Tables(0).Rows
            If .Count > 0 Then
                CodigoCarteira = .Item(0)("Produto_Id")
                Descricao = .Item(0)("Descricao")
                CodigoSituacao = .Item(0)("Situacao")
                isAdiantamento = .Item(0)("Adiantamento") = "S"
                CodigoContaCliente = .Item(0)("ContaClientes")
                CodigoContaDesconto = .Item(0)("ContaDescontos")
                CodigoContaDeducao = .Item(0)("ContaDeducoes")
                CodigoContaJuro = .Item(0)("ContaJuros")
                CodigoContaAcrescimo = .Item(0)("ContaAcrescimos")
                Classificacao = .Item(0)("Classificacao")
                TemPedido = .Item(0)("Pedido")
                BaixaAdiantamento = .Item(0)("BaixaAdiantamento")
                CodigoCarteiraBaixaAdiantamento = .Item(0)("CarteiraBaixaAdiantamento")
                UsuarioInclusao = .Item(0)("UsuarioInclusao")
                UsuarioInclusaoData = .Item(0)("UsuarioInclusaoData")
                UsuarioAlteracao = .Item(0)("UsuarioAlteracao")
                UsuarioAlteracaoData = .Item(0)("UsuarioAlteracaoData")
            End If
        End With
    End Sub
#End Region

#Region "Fields"
    Private _IUD As String = ""
    Private _CodigoCarteira As String
    Private _Descricao As String

    Private _Situacao As Situacao
    Private _CodigoSituacao As Integer

    Private _isAdiantamento As Boolean

    Private _CodigoContaCliente As String = ""
    Private _CodigoContaDesconto As String = ""
    Private _CodigoContaDeducao As String = ""
    Private _CodigoContaJuro As String = ""
    Private _CodigoContaAcrescimo As String = ""

    Private _ContaCliente As PlanoDeConta
    Private _ContaDesconto As PlanoDeConta
    Private _ContaDeducao As PlanoDeConta
    Private _ContaJuro As PlanoDeConta
    Private _ContaAcrescimo As PlanoDeConta

    Private _TemPedido As Boolean = False
    Private _BaixaAdiantamento As Boolean = False
    Private _CodigoCarteiraBaixaAdiantamento As String
    Private _CarteiraBaixaAdiantamento As CarteiraFinanceira

    Private _Classificacao As String
    Private _Tributos As ListCarteiraFinanceiraXTributoEncargo
    Private _UsuarioInclusao As String
    Private _UsuarioInclusaoData As DateTime
    Private _UsuarioAlteracao As String
    Private _UsuarioAlteracaoData As DateTime

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

    Public Property CodigoCarteira() As String
        Get
            Return _CodigoCarteira
        End Get
        Set(ByVal value As String)
            _CodigoCarteira = value
            _Tributos = Nothing
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

    Public Property CodigoSituacao() As Integer
        Get
            Return _CodigoSituacao
        End Get
        Set(ByVal value As Integer)
            _CodigoSituacao = value
            _Situacao = Nothing
        End Set
    End Property

    Public Property Situacao() As Situacao
        Get
            If _Situacao Is Nothing And _CodigoSituacao > 0 Then _Situacao = New Situacao(_CodigoSituacao)
            Return _Situacao
        End Get
        Set(ByVal value As Situacao)
            _Situacao = value
        End Set
    End Property

    Public Property isAdiantamento() As Boolean
        Get
            Return _isAdiantamento
        End Get
        Set(ByVal value As Boolean)
            _isAdiantamento = value
        End Set
    End Property

    Public Property CodigoContaCliente() As String
        Get
            Return _CodigoContaCliente
        End Get
        Set(ByVal value As String)
            _CodigoContaCliente = value
            _ContaCliente = Nothing
        End Set
    End Property

    Public Property ContaCliente() As PlanoDeConta
        Get
            If _ContaCliente Is Nothing And _CodigoContaCliente.Length > 0 Then _ContaCliente = New PlanoDeConta("", 0, _CodigoContaCliente)
            Return _ContaCliente
        End Get
        Set(ByVal value As PlanoDeConta)
            _ContaCliente = value
        End Set
    End Property

    Public Property CodigoContaDesconto() As String
        Get
            Return _CodigoContaDesconto
        End Get
        Set(ByVal value As String)
            _CodigoContaDesconto = value
            _ContaDesconto = Nothing
        End Set
    End Property

    Public Property ContaDesconto() As PlanoDeConta
        Get
            If _ContaDesconto Is Nothing And _CodigoContaDesconto.Length > 0 Then
                _ContaDesconto = New PlanoDeConta("", 0, _CodigoContaDesconto)
                If _ContaDesconto.Titulo.Length = 0 Then _ContaDesconto = Nothing
            End If

            Return _ContaDesconto
        End Get
        Set(ByVal value As PlanoDeConta)
            _ContaDesconto = value
        End Set
    End Property

    Public Property CodigoContaDeducao() As String
        Get
            Return _CodigoContaDeducao
        End Get
        Set(ByVal value As String)
            _CodigoContaDeducao = value
            _ContaDeducao = Nothing
        End Set
    End Property

    Public Property ContaDeducao() As PlanoDeConta
        Get
            If _ContaDeducao Is Nothing And _CodigoContaDeducao.Length > 0 Then
                _ContaDeducao = New PlanoDeConta("", 0, _CodigoContaDeducao)
                If _ContaDeducao.Titulo.Length = 0 Then _ContaDeducao = Nothing
            End If

            Return _ContaDeducao
        End Get
        Set(ByVal value As PlanoDeConta)
            _ContaDeducao = value
        End Set
    End Property

    Public Property CodigoContaJuro() As String
        Get
            Return _CodigoContaJuro
        End Get
        Set(ByVal value As String)
            _CodigoContaJuro = value
            _ContaJuro = Nothing
        End Set
    End Property

    Public Property ContaJuro() As PlanoDeConta
        Get
            If _ContaJuro Is Nothing And _CodigoContaJuro.Length > 0 Then
                _ContaJuro = New PlanoDeConta("", 0, _CodigoContaJuro)
                If _ContaJuro.Titulo.Length = 0 Then _ContaJuro = Nothing
            End If

            Return _ContaJuro
        End Get
        Set(ByVal value As PlanoDeConta)
            _ContaJuro = value
        End Set
    End Property

    Public Property CodigoContaAcrescimo() As String
        Get
            Return _CodigoContaAcrescimo
        End Get
        Set(ByVal value As String)
            _CodigoContaAcrescimo = value
            _ContaAcrescimo = Nothing
        End Set
    End Property

    Public Property ContaAcrescimo() As PlanoDeConta
        Get
            If _ContaAcrescimo Is Nothing And _CodigoContaAcrescimo.Length > 0 Then
                _ContaAcrescimo = New PlanoDeConta("", 0, _CodigoContaAcrescimo)
                If _ContaAcrescimo.Titulo.Length = 0 Then _ContaAcrescimo = Nothing
            End If

            Return _ContaAcrescimo
        End Get
        Set(ByVal value As PlanoDeConta)
            _ContaAcrescimo = value
        End Set
    End Property

    Public Property Classificacao() As String
        Get
            Return _Classificacao
        End Get
        Set(ByVal value As String)
            _Classificacao = value
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

    Public Property UsuarioAlteracao() As String
        Get
            Return _UsuarioAlteracao
        End Get
        Set(ByVal value As String)
            _UsuarioAlteracao = value
        End Set
    End Property

    Public Property UsuarioAlteracaoData() As DateTime
        Get
            Return _UsuarioAlteracaoData
        End Get
        Set(ByVal value As DateTime)
            _UsuarioAlteracaoData = value
        End Set
    End Property

    Public Property Tributos() As ListCarteiraFinanceiraXTributoEncargo
        Get
            If _Tributos Is Nothing And _CodigoCarteira.Length > 0 Then _Tributos = New ListCarteiraFinanceiraXTributoEncargo(_CodigoCarteira)
            Return _Tributos
        End Get
        Set(ByVal value As ListCarteiraFinanceiraXTributoEncargo)
            _Tributos = value
        End Set
    End Property

    Public Property TemPedido() As Boolean
        Get
            Return _TemPedido
        End Get
        Set(ByVal value As Boolean)
            _TemPedido = value
        End Set
    End Property

    Public Property BaixaAdiantamento() As Boolean
        Get
            Return _BaixaAdiantamento
        End Get
        Set(ByVal value As Boolean)
            _BaixaAdiantamento = value
        End Set
    End Property

    Public Property CodigoCarteiraBaixaAdiantamento() As String
        Get
            Return _CodigoCarteiraBaixaAdiantamento
        End Get
        Set(ByVal value As String)
            _CodigoCarteiraBaixaAdiantamento = value
            _CarteiraBaixaAdiantamento = Nothing
        End Set
    End Property

    Public Property CarteiraBaixaAdiantamento() As CarteiraFinanceira
        Get
            If _CarteiraBaixaAdiantamento Is Nothing And _CodigoCarteiraBaixaAdiantamento.Length > 0 Then _CarteiraBaixaAdiantamento = New CarteiraFinanceira(_CodigoCarteiraBaixaAdiantamento)
            Return _CarteiraBaixaAdiantamento
        End Get
        Set(ByVal value As CarteiraFinanceira)
            _CarteiraBaixaAdiantamento = value
        End Set
    End Property

#End Region

#Region "Methods"

    Public Function VerificaCarteiraTitulos() As Boolean
        Dim Banco As New AcessaBanco
        Dim strSQL As String

        strSQL = "Select Top 1 Registro_Id from ContasAPagar " & vbCrLf & _
                 "Where Carteira = '" & _CodigoCarteira & "' OR CarteiraAdto = '" & _CodigoCarteira & "'" & vbCrLf

        If _CodigoCarteiraBaixaAdiantamento.Length > 0 Then
            strSQL &= " OR Carteira = '" & _CodigoCarteiraBaixaAdiantamento & "' OR CarteiraAdto = '" & _CodigoCarteiraBaixaAdiantamento & "'" & vbCrLf
        End If

        Dim dsP As New DataSet
        dsP = Banco.ConsultaDataSet(strSQL, "CarteiraPagar")

        If Not dsP Is Nothing AndAlso dsP.Tables(0).Rows.Count > 0 Then Return True

        strSQL = "Select Top 1 Registro_Id from ContasAReceber " & vbCrLf & _
                 "Where Carteira = '" & _CodigoCarteira & "' OR CarteiraAdto = '" & _CodigoCarteira & "'" & vbCrLf

        If _CodigoCarteiraBaixaAdiantamento.Length > 0 Then
            strSQL &= " OR Carteira = '" & _CodigoCarteiraBaixaAdiantamento & "' OR CarteiraAdto = '" & _CodigoCarteiraBaixaAdiantamento & "'" & vbCrLf
        End If

        Dim dsR As New DataSet
        dsR = Banco.ConsultaDataSet(strSQL, "CarteiraReceber")

        If Not dsR Is Nothing AndAlso dsR.Tables(0).Rows.Count > 0 Then Return True

        strSQL = "Select Top 1 Registro_Id from MovimentacoesFinanceiras " & vbCrLf & _
                 "Where Carteira = '" & _CodigoCarteira & "' OR CarteiraAdto = '" & _CodigoCarteira & "'" & vbCrLf

        If _CodigoCarteiraBaixaAdiantamento.Length > 0 Then
            strSQL &= " OR Carteira = '" & _CodigoCarteiraBaixaAdiantamento & "' OR CarteiraAdto = '" & _CodigoCarteiraBaixaAdiantamento & "'" & vbCrLf
        End If

        Dim dsM As New DataSet
        dsM = Banco.ConsultaDataSet(strSQL, "CarteiraMFinanceiras")

        If Not dsM Is Nothing AndAlso dsM.Tables(0).Rows.Count > 0 Then Return True

        Return False
    End Function

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
                sql = "INSERT INTO ComprasXProdutos (Produto_Id, Descricao, Situacao, ContaCLientes, ContaDescontos, " & vbCrLf & _
                      "            ContaDeducoes, ContaJuros, ContaAcrescimos, Classificacao," & vbCrLf & _
                      "            UsuarioInclusao, UsuarioInclusaoData, " & vbCrLf & _
                      "            Adiantamento, Pedido, BaixaAdiantamento, CarteiraBaixaAdiantamento) " & vbCrLf & _
                      "Values ('" & Me.CodigoCarteira & "','" & Me.Descricao & "','" & Me.CodigoSituacao & "','" & Me.CodigoContaCliente & "','" & Me.CodigoContaDesconto & "'," & vbCrLf & _
                      "'" & Me.CodigoContaDeducao & "','" & Me.CodigoContaJuro & "','" & Me.CodigoContaAcrescimo & "','" & Me.Classificacao & "'," & vbCrLf & _
                      "'" & HttpContext.Current.Session("ssNomeUsuario") & "', '" & Now().ToString("yyyy-MM-dd hh:mm:ss") & "'," & vbCrLf & _
                      "'" & IIf(Me.isAdiantamento, "S", "N") & "'," & CByte(Me.TemPedido) & "," & CByte(Me.BaixaAdiantamento) & ",'" & Me.CodigoCarteiraBaixaAdiantamento & "')"
                Sqls.Add(sql)
            Case "U"
                sql = "Update ComprasXProdutos Set Descricao            = '" & Me.Descricao & "'," & vbCrLf & _
                                                  "Situacao             = '" & Me.CodigoSituacao & "'," & vbCrLf & _
                                                  "ContaCLientes        = '" & Me.CodigoContaCliente & "'," & vbCrLf & _
                                                  "ContaDescontos       = '" & Me.CodigoContaDesconto & "'," & vbCrLf & _
                                                  "ContaDeducoes        = '" & Me.CodigoContaDeducao & "'," & vbCrLf & _
                                                  "ContaJuros           = '" & Me.CodigoContaJuro & "'," & vbCrLf & _
                                                  "ContaAcrescimos      = '" & Me.CodigoContaAcrescimo & "'," & vbCrLf & _
                                                  "Classificacao        = '" & Me.Classificacao & "'," & vbCrLf & _
                                                  "UsuarioAlteracao     = '" & HttpContext.Current.Session("ssNomeUsuario") & "'," & vbCrLf & _
                                                  "UsuarioAlteracaoData = '" & Now().ToString("yyyy-MM-dd hh:mm:ss") & "'," & vbCrLf & _
                                                  "Adiantamento         = '" & IIf(Me.isAdiantamento, "S", "N") & "'" & vbCrLf & _
                                                  "Pedido               = " & CByte(Me.TemPedido) & vbCrLf & _
                                                  "BaixaAdiantamento    = " & CByte(Me.BaixaAdiantamento) & vbCrLf & _
                                                  "CarteiraBaixaAdiantamento = '" & Me.CodigoCarteiraBaixaAdiantamento & "'" & vbCrLf & _
                             "Where Produto_Id = '" & Me.CodigoCarteira & "'"
                Sqls.Add(sql)
            Case "D"
                sql = "Delete From ComprasXProdutos Where Produto_Id = '" & Me.CodigoCarteira & "'"
                Sqls.Add(sql)
        End Select
    End Sub
#End Region

End Class