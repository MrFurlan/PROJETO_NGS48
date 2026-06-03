Imports System.Data
Imports System.Collections.Generic
Imports Microsoft.VisualBasic
Imports NGS.Lib.Uteis

<Serializable()> _
Public Class ListPlanoDeConta
    Inherits List(Of PlanoDeConta)
    Implements IBaseEntity

#Region "Contrutor"
    Public Sub New()

    End Sub

    Public Sub New(ByVal pWhere As String)
        Selecionar("", 0, pWhere)
    End Sub

    Public Sub New(ByVal pCodigoEmpresa As String, ByVal pEndEmpresa As Integer)
        Selecionar(pCodigoEmpresa, pEndEmpresa, "")
    End Sub
#End Region

#Region "Methods"

    Private Sub Selecionar(ByVal pCodigoEmpresa As String, ByVal pEndEmpresa As Integer, ByVal Where As String)
        Dim objBanco As New AcessaBanco()
        Dim strSQL As String

        strSQL = "SELECT Empresa_Id, EndEmpresa_Id, Conta_Id, Titulo, isnull(Cliente,'N') AS Cliente, isnull(Produto,'N') AS Produto, " & vbCrLf & _
                 "       isnull(CentroDeCusto,'N') AS CentroDeCusto, ISNULL(TipoDeCliente, 0) AS TipoDeCliente, Responsabilidade, ContaOrcamentaria," & vbCrLf & _
                 "       isnull(ContaAnterior,'') as ContaAnterior, isnull(Dacon,0) as Dacon, isnull(TipoDeCusto,'NENHUM') as TipoDeCusto,   " & vbCrLf & _
                 "       isnull(Pagar,'') as Pagar, isnull(Receber,'') as Receber, isnull(TipoDeConta,'') as TipoDeConta," & vbCrLf & _
                 "       isnull(Adiantamento,0) as Adiantamento, isnull(Pedido,0) as Pedido, " & vbCrLf & _
                 "       isnull(Encargo,0) as Encargo, isnull(TemEncargo,0) as TemEncargo, " & vbCrLf & _
                 "       (Select count(*) from EncargosPlanoDeContas EPC where EPC.Conta_Id = PlanoDeContas.Conta_Id ) as NrEncargos, isnull(AdiantamentoSoContabil,0) as AdiantamentoSoContabil, isnull(Ativo,1) AS Ativo, " & vbCrLf & _
                 "       isnull(UsuarioInclusao, '') as UsuarioInclusao, isnull(UsuarioInclusaoData, '') as UsuarioInclusaoData, isnull(UsuarioAlteracao, '') as UsuarioAlteracao, isnull(UsuarioAlteracaoData, '') as UsuarioAlteracaoData, " & vbCrLf & _
                 "       isnull(ProdutoParaCusto,0) as ProdutoParaCusto" & vbCrLf & _
                 "  FROM PlanoDeContas " & vbCrLf

        If pCodigoEmpresa.Length > 0 Then
            strSQL = " Where Empresa_id    ='" & pCodigoEmpresa & "'" & _
                     "   and EndEmpresa_id = " & pEndEmpresa
            If Where.Length > 0 Then strSQL &= " " & Where
        Else
            If Where.Length > 0 Then strSQL &= "Where " & Where
        End If

        strSQL &= " ORDER BY Conta_Id "

        Dim dsPlanoDeContas As DataSet = objBanco.ConsultaDataSet(strSQL, "PlanoDeContas")

        If dsPlanoDeContas.Tables(0).Rows.Count = 0 Then
            strSQL = "SELECT Empresa_Id, EndEmpresa_Id, Conta_Id, Titulo, isnull(Cliente,'N') AS Cliente, isnull(Produto,'N') AS Produto, " & vbCrLf & _
                     "       isnull(CentroDeCusto,'N') AS CentroDeCusto, ISNULL(TipoDeCliente, 0) AS TipoDeCliente, Responsabilidade, ContaOrcamentaria," & vbCrLf & _
                     "       isnull(ContaAnterior,'') as ContaAnterior, isnull(Dacon,0) as Dacon, isnull(TipoDeCusto,'NENHUMA') TipoDeCusto, " & vbCrLf & _
                     "       isnull(Pagar,'') as Pagar, isnull(Receber,'') as Receber, isnull(TipoDeConta,'') as TipoDeConta," & vbCrLf & _
                     "       isnull(Adiantamento,0) as Adiantamento, isnull(Pedido,0) as Pedido, " & vbCrLf & _
                     "       isnull(Encargo,0) as Encargo, isnull(TemEncargo,0) as TemEncargo," & vbCrLf & _
                     "       (Select count(*) from EncargosPlanoDeContas EPC where EPC.Conta_Id = PlanoDeContas.Conta_Id ) as NrEncargos, isnull(AdiantamentoSoContabil,0) as AdiantamentoSoContabil, isnull(Ativo,1) AS Ativo, " & vbCrLf & _
                     "       isnull(UsuarioInclusao, '') as UsuarioInclusao, isnull(UsuarioInclusaoData, '') as UsuarioInclusaoData, isnull(UsuarioAlteracao, '') as UsuarioAlteracao, isnull(UsuarioAlteracaoData, '') as UsuarioAlteracaoData, " & vbCrLf & _
                     "       isnull(ProdutoParaCusto,0) as ProdutoParaCusto" & vbCrLf & _
                     "  FROM PlanoDeContas " & vbCrLf & _
                     " Where Empresa_id    = '99999999999999'" & vbCrLf & _
                     "   and EndEmpresa_id = 0" & vbCrLf

            If Where.Length > 0 Then strSQL &= " and " & Where

            strSQL &= " ORDER BY Conta_Id "
            dsPlanoDeContas = objBanco.ConsultaDataSet(strSQL, "PlanoDeContas")
        End If

        For Each row As DataRow In dsPlanoDeContas.Tables(0).Rows
            Dim Conta As New PlanoDeConta
            Conta.CodigoEmpresa = row("Empresa_Id")
            Conta.EnderecoEmpresa = row("EndEmpresa_Id")
            Conta.Conta = row("Conta_Id")
            Conta.Titulo = row("Titulo")
            Conta.TemCliente = row("Cliente") = "S"
            Conta.TemProduto = row("Produto") = "S"
            Conta.TemCentroDeCusto = row("CentroDeCusto") = "S"
            Conta.TipoDeCusto = row("TipoDeCusto")
            Conta.CodigoTipoDeCliente = row("TipoDeCliente")
            Conta.Responsabilidade = row("Responsabilidade")
            Conta.ContaOrcamentaria = row("ContaOrcamentaria")
            Conta.ContaAnterior = row("ContaAnterior")
            Conta.Dacon = row("Dacon")
            Conta.Pagar = row("Pagar")
            Conta.Receber = row("Receber")
            Conta.TipodeConta = row("TipoDeConta")
            Conta.Adiantamento = row("Adiantamento")
            Conta.AdiantamentoSoContabil = row("AdiantamentoSoContabil")
            Conta.TemPedido = row("Pedido")
            Conta.Encargo = row("Encargo")
            Conta.TemEncargo = row("TemEncargo")
            Conta.NrEncargos = IIf(row("NrEncargos") = 0, "", row("NrEncargos"))
            Conta.Ativo = row("Ativo")
            Conta.UsuarioInclusao = row("UsuarioInclusao")
            Conta.UsuarioInclusaoData = row("UsuarioInclusaoData")
            Conta.UsuarioAlteracao = row("UsuarioAlteracao")
            Conta.UsuarioAlteracaoData = row("UsuarioAlteracaoData")
            Conta.ProdutoParaCusto = row("ProdutoParaCusto")
            Add(Conta)
        Next
    End Sub
#End Region

End Class

<Serializable()> _
Public Class PlanoDeConta
    Implements IBaseEntity

#Region "Construtor"

    Public Sub New()

    End Sub

    Public Sub New(ByVal pCodigoEmpresa As String, ByVal pEndEmpresa As Integer, ByVal pConta As String)
        Dim objBanco As New AcessaBanco()
        Dim strSQL As String

        strSQL = "SELECT Empresa_Id, EndEmpresa_Id, Conta_Id, Titulo, Cliente, Produto, CentroDeCusto, " & vbCrLf & _
                 "       ISNULL(TipoDeCliente, 0) AS TipoDeCliente, Responsabilidade, ContaOrcamentaria, " & vbCrLf & _
                 "       isnull(ContaAnterior,'') as ContaAnterior, isnull(Dacon,0) as Dacon, isnull(TipoDeCusto,'NENHUMA') TipoDeCusto," & vbCrLf & _
                 "       isnull(Pagar,'') as Pagar, isnull(Receber,'') as Receber, isnull(TipoDeConta,'') as TipoDeConta," & vbCrLf & _
                 "       isnull(Adiantamento,0) as Adiantamento, isnull(AdiantamentoSoContabil,0) as AdiantamentoSoContabil, isnull(Pedido,0) as Pedido," & vbCrLf & _
                 "       isnull(Encargo,0) as Encargo, isnull(TemEncargo,0) as TemEncargo, " & vbCrLf & _
                 "       (Select count(*) from EncargosPlanoDeContas EPC where EPC.Conta_Id = PlanoDeContas.Conta_Id ) as NrEncargos, isnull(Ativo,1) AS Ativo, " & vbCrLf & _
                 "       isnull(UsuarioInclusao, '') as UsuarioInclusao, isnull(UsuarioInclusaoData, '') as UsuarioInclusaoData, isnull(UsuarioAlteracao, '')as UsuarioAlteracao, isnull(UsuarioAlteracaoData, '') as UsuarioAlteracaoData, " & vbCrLf & _
                 "       isnull(ProdutoParaCusto,0) as ProdutoParaCusto" & vbCrLf & _
                 "  FROM PlanoDeContas " & vbCrLf
        If pCodigoEmpresa.Length > 0 Then
            strSQL &= " Where Empresa_id    ='" & pCodigoEmpresa & "'" & vbCrLf & _
                      "   and EndEmpresa_id = " & pEndEmpresa & vbCrLf & _
                      "   and Conta_id      ='" & pConta & "'" & vbCrLf
        Else
            strSQL &= " Where Empresa_id    = '99999999999999'" & vbCrLf & _
                      "   and EndEmpresa_id = 0" & vbCrLf & _
                      "   and Conta_id      ='" & pConta & "'" & vbCrLf
        End If

        strSQL &= " ORDER BY Conta_Id "

        Dim dsPlanoDeContas As DataSet = objBanco.ConsultaDataSet(strSQL, "PlanoDeContas")

        If dsPlanoDeContas.Tables(0).Rows.Count = 0 Then Exit Sub

        Dim row As DataRow = dsPlanoDeContas.Tables(0).Rows(0)

        Me.CodigoEmpresa = row("Empresa_Id")
        Me.EnderecoEmpresa = row("EndEmpresa_Id")
        Me.Conta = row("Conta_Id")
        Me.Titulo = row("Titulo")
        Me.TemCliente = row("Cliente") = "S"
        Me.TemProduto = row("Produto") = "S"
        Me.TemCentroDeCusto = row("CentroDeCusto") = "S"
        Me.TipoDeCusto = row("TipoDeCusto")
        Me.CodigoTipoDeCliente = row("TipoDeCliente")
        Me.Responsabilidade = row("Responsabilidade")
        Me.ContaOrcamentaria = row("ContaOrcamentaria")
        Me.ContaAnterior = row("ContaAnterior")
        Me.Dacon = row("Dacon")
        Me.Pagar = row("Pagar")
        Me.Receber = row("Receber")
        Me.TipodeConta = row("TipoDeConta")
        Me.Adiantamento = row("Adiantamento")
        Me.AdiantamentoSoContabil = row("AdiantamentoSoContabil")
        Me.TemPedido = row("Pedido")
        Me.Encargo = row("Encargo")
        Me.TemEncargo = row("TemEncargo")
        Me.NrEncargos = IIf(row("NrEncargos") = 0, "", row("NrEncargos"))
        Me.Ativo = row("Ativo")
        Me.UsuarioInclusao = row("UsuarioInclusao")
        Me.UsuarioInclusaoData = row("UsuarioInclusaoData")
        Me.UsuarioAlteracao = row("UsuarioAlteracao")
        Me.UsuarioAlteracaoData = row("UsuarioAlteracaoData")
        Me.ProdutoParaCusto = row("ProdutoParaCusto")
    End Sub

#End Region

#Region "Fields"

    Private _IUD As String
    Private _Parent As PlanoDeConta

    Private _CodigoEmpresa As String = ""
    Private _EnderecoEmpresa As Integer
    Private _Empresa As Cliente

    Private _Conta As String = ""
    Private _Titulo As String = ""

    Private _Cliente As Boolean = False
    Private _Produto As Boolean = False
    Private _CentroDeCusto As Boolean = False
    Private _TipoDeCusto As String

    Private _CodigoTipoDeCliente As Integer
    Private _TipoDeCliente As TipoDeCliente
    Private _Responsabilidade As String = ""
    Private _ContaOrcamentaria As String = ""
    Private _ContaAnterior As String = ""
    Private _Dacon As Integer

    Private _Pagar As String = ""
    Private _Receber As String = ""
    Private _TipodeConta As String = ""
    Private _Adiantamento As Boolean = False
    Private _AdiantamentoSoContabil As Boolean = False
    Private _TemPedido As Boolean = False
    Private _Encargo As Boolean = False
    Private _TemEncargo As Boolean = False
    Private _ProdutoParaCusto As Boolean = False
    Private _NrEncargos As String = ""
    Private _EncargosPlanoDeContas As ListEncargosPlanoDeContas
    Private _Ativo As Boolean
    Private _UsuarioInclusao As String = ""
    Private _UsuarioInclusaoData As DateTime
    Private _UsuarioAlteracao As String = ""
    Private _UsuarioAlteracaoData As DateTime
#End Region

#Region "Propriedades"

    Public Property IUD() As String
        Get
            Return _IUD
        End Get
        Set(ByVal value As String)
            _IUD = value
        End Set
    End Property

    Public Property Parent() As PlanoDeConta
        Get
            Return _Parent
        End Get
        Set(ByVal value As PlanoDeConta)
            _Parent = value
        End Set
    End Property

    Public Property CodigoEmpresa() As String
        Get
            Return _CodigoEmpresa
        End Get
        Set(ByVal value As String)
            _CodigoEmpresa = value
            _Empresa = Nothing
        End Set
    End Property

    Public Property EnderecoEmpresa() As Integer
        Get
            Return _EnderecoEmpresa
        End Get
        Set(ByVal value As Integer)
            _EnderecoEmpresa = value
            _Empresa = Nothing
        End Set
    End Property

    Public Property Empresa() As Cliente
        Get
            If _Empresa Is Nothing And _CodigoEmpresa.Length > 0 Then _Empresa = New Cliente(_CodigoEmpresa, _EnderecoEmpresa)
            Return _Empresa
        End Get
        Set(ByVal value As Cliente)
            _Empresa = value
        End Set
    End Property

    Public Property Conta() As String
        Get
            Return _Conta
        End Get
        Set(ByVal value As String)
            _Conta = value
        End Set
    End Property

    Public Property Titulo() As String
        Get
            Return _Titulo
        End Get
        Set(ByVal value As String)
            _Titulo = value
        End Set
    End Property

    Public Property TemCliente() As Boolean
        Get
            Return _Cliente
        End Get
        Set(ByVal value As Boolean)
            _Cliente = value
        End Set
    End Property

    Public Property TemProduto() As Boolean
        Get
            Return _Produto
        End Get
        Set(ByVal value As Boolean)
            _Produto = value
        End Set
    End Property

    Public Property TemCentroDeCusto() As Boolean
        Get
            Return _CentroDeCusto
        End Get
        Set(ByVal value As Boolean)
            _CentroDeCusto = value
        End Set
    End Property

    Public Property TipoDeCusto() As String
        Get
            Return _TipoDeCusto
        End Get
        Set(ByVal value As String)
            _TipoDeCusto = value
        End Set
    End Property

    Public Property CodigoTipoDeCliente() As Integer
        Get
            Return _CodigoTipoDeCliente
        End Get
        Set(ByVal value As Integer)
            _CodigoTipoDeCliente = value
            _TipoDeCliente = Nothing
        End Set
    End Property

    Public ReadOnly Property TipoDeCliente()
        Get
            If _TipoDeCliente Is Nothing And _CodigoTipoDeCliente > 0 Then _TipoDeCliente = New TipoDeCliente(_CodigoTipoDeCliente)
            Return _TipoDeCliente
        End Get
    End Property

    Public Property Responsabilidade() As String
        Get
            Return _Responsabilidade
        End Get
        Set(ByVal value As String)
            _Responsabilidade = value
        End Set
    End Property

    Public Property ContaOrcamentaria() As String
        Get
            Return _ContaOrcamentaria
        End Get
        Set(ByVal value As String)
            _ContaOrcamentaria = value
        End Set
    End Property

    Public Property ContaAnterior() As String
        Get
            Return _ContaAnterior
        End Get
        Set(ByVal value As String)
            _ContaAnterior = value
        End Set
    End Property

    Public Property Dacon() As Integer
        Get
            Return _Dacon
        End Get
        Set(ByVal value As Integer)
            _Dacon = value
        End Set
    End Property

    Public Property Pagar() As String
        Get
            Return _Pagar
        End Get
        Set(ByVal value As String)
            _Pagar = value
        End Set
    End Property

    Public Property Receber() As String
        Get
            Return _Receber
        End Get
        Set(ByVal value As String)
            _Receber = value
        End Set
    End Property

    Public Property TipodeConta() As String
        Get
            Return _TipodeConta
        End Get
        Set(ByVal value As String)
            _TipodeConta = value
        End Set
    End Property

    Public Property Adiantamento() As Boolean
        Get
            Return _Adiantamento
        End Get
        Set(ByVal value As Boolean)
            _Adiantamento = value
        End Set
    End Property

    Public Property AdiantamentoSoContabil() As Boolean
        Get
            Return _AdiantamentoSoContabil
        End Get
        Set(ByVal value As Boolean)
            _AdiantamentoSoContabil = value
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

    Public Property Encargo() As Boolean
        Get
            Return _Encargo
        End Get
        Set(ByVal value As Boolean)
            _Encargo = value
        End Set
    End Property

    Public Property TemEncargo() As Boolean
        Get
            Return _TemEncargo
        End Get
        Set(ByVal value As Boolean)
            _TemEncargo = value
        End Set
    End Property

    Public Property NrEncargos() As String
        Get
            Return _NrEncargos
        End Get
        Set(ByVal value As String)
            _NrEncargos = value
        End Set
    End Property

    Public Property EncargosPlanoDeContas() As ListEncargosPlanoDeContas
        Get
            If _EncargosPlanoDeContas Is Nothing Then _EncargosPlanoDeContas = New ListEncargosPlanoDeContas(Me)
            Return _EncargosPlanoDeContas
        End Get
        Set(ByVal value As ListEncargosPlanoDeContas)
            _EncargosPlanoDeContas = value
        End Set
    End Property

    Public Property Ativo() As Boolean
        Get
            Return _Ativo
        End Get
        Set(ByVal value As Boolean)
            _Ativo = value
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

    Public Property ProdutoParaCusto() As Boolean
        Get
            Return _ProdutoParaCusto
        End Get
        Set(ByVal value As Boolean)
            _ProdutoParaCusto = value
        End Set
    End Property

#End Region

#Region "Methods"

    Public Function VerificaRazao() As Boolean
        Dim strSQL As String = "Select Conta_id From Razao " & vbCrLf & _
                               "Where Conta_Id = '" & Me.Conta & "'"

        Dim ObjBanco As New AcessaBanco
        Dim dsConta As New DataSet
        dsConta = ObjBanco.ConsultaDataSet(strSQL, "Razao")

        If Not dsConta Is Nothing AndAlso dsConta.Tables(0).Rows.Count > 0 Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Function VerificaOperacao() As Boolean
        Dim strSQL As String = "Select GrupoDeContas From SubOperacoes " & vbCrLf & _
                               "Where GrupoDeContas = '" & Me.Conta & "'"

        Dim ObjBanco As New AcessaBanco
        Dim dsOperacao As New DataSet
        dsOperacao = ObjBanco.ConsultaDataSet(strSQL, "SubOperacoes")

        If Not dsOperacao Is Nothing AndAlso dsOperacao.Tables(0).Rows.Count > 0 Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Function VerificaOperacoesXEncargos() As Boolean
        Dim strSQL As String = "Select DebitaConta, CreditaConta" & vbCrLf & _
                               "  From OperacaoXEstadoXEncargo " & vbCrLf & _
                               " Where DebitaConta = '" & Me.Conta & "' OR CreditaConta = '" & Me.Conta & "'"

        Dim ObjBanco As New AcessaBanco
        Dim dsOpXEnc As New DataSet
        dsOpXEnc = ObjBanco.ConsultaDataSet(strSQL, "OperacoesXEncargos")

        If Not dsOpXEnc Is Nothing AndAlso dsOpXEnc.Tables(0).Rows.Count > 0 Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Function Salvar() As Boolean
        Dim Banco As New AcessaBanco
        Dim sqls As New ArrayList
        SalvarSql(sqls)
        Return Banco.GravaBanco(sqls)
    End Function

    Public Function SalvarSql(ByRef sqls As ArrayList) As ArrayList
        Dim strSQL As String
        Dim ObjBanco As New AcessaBanco

        Select Case Me.IUD
            Case "I"
                strSQL = "INSERT INTO PlanoDeContas(Empresa_Id, EndEmpresa_Id, Conta_Id, Titulo, Cliente, Produto," & vbCrLf & _
                         "                          CentroDeCusto, TipodeCliente, Responsabilidade, ContaOrcamentaria," & vbCrLf & _
                         "                          ContaAnterior, Dacon, Pagar, Receber, TipoDeConta, TipoDeCusto, Adiantamento, AdiantamentoSoContabil," & vbCrLf & _
                         "                          Pedido, Encargo, TemEncargo, Ativo, UsuarioInclusao, UsuarioInclusaoData, ProdutoParaCusto)" & vbCrLf & _
                         " VALUES ('" & Me.CodigoEmpresa & "'," & Me.EnderecoEmpresa & ",'" & Me.Conta & "','" & Me.Titulo & "','" & IIf(Me.TemCliente, "S", "N") & "','" & IIf(Me.TemProduto, "S", "N") & "'," & vbCrLf & _
                         "'" & IIf(Me.TemCentroDeCusto, "S", "N") & "'," & Me.CodigoTipoDeCliente & ",'" & Me.Responsabilidade & "','" & Me.ContaOrcamentaria & "'," & vbCrLf & _
                         "'" & Me.ContaAnterior & "'," & Me.Dacon & ",'" & Me.Pagar & "','" & Me.Receber & "'," & IIf(Me.TipodeConta.Length = 0, "NULL", "'" & Me.TipodeConta & "'") & ",'" & Me.TipoDeCusto & "'," & IIf(Me.Adiantamento, 1, 0) & "," & IIf(Me.AdiantamentoSoContabil, 1, 0) & "," & vbCrLf & _
                         IIf(Me.TemPedido, 1, 0) & "," & IIf(Me.Encargo, 1, 0) & "," & IIf(Me.TemEncargo, 1, 0) & "," & IIf(Me.Ativo, 1, 0) & ", '" & Me.UsuarioInclusao & "' , '" & Me.UsuarioInclusaoData.ToString("yyyy-MM-dd H:mm:ss") & "', " & IIf(Me.ProdutoParaCusto, 1, 0) & ")"
                sqls.Add(strSQL)
                SalvarTabelasRelacionada(sqls)
            Case "U"
                strSQL = " UPDATE PlanoDeContas Set " & vbCrLf & _
                         "   Titulo                  ='" & Me.Titulo & "'" & vbCrLf & _
                         "  ,Cliente                 ='" & IIf(Me.TemCliente, "S", "N") & "'" & vbCrLf & _
                         "  ,Produto                 ='" & IIf(Me.TemProduto, "S", "N") & "'" & vbCrLf & _
                         "  ,CentroDeCusto           ='" & IIf(Me.TemCentroDeCusto, "S", "N") & "'" & vbCrLf & _
                         "  ,TipodeCliente           = " & Me.CodigoTipoDeCliente & vbCrLf & _
                         "  ,TipoDeCusto             ='" & Me.TipoDeCusto & "'" & vbCrLf & _
                         "  ,Responsabilidade        ='" & Me.Responsabilidade & "'" & vbCrLf & _
                         "  ,ContaOrcamentaria       ='" & Me.ContaOrcamentaria & "'" & vbCrLf & _
                         "  ,ContaAnterior           ='" & Me.ContaAnterior & "'" & vbCrLf & _
                         "  ,Dacon                   = " & Me.Dacon & vbCrLf & _
                         "  ,Pagar                   ='" & Me.Pagar & "'" & vbCrLf & _
                         "  ,Receber                 ='" & Me.Receber & "'" & vbCrLf & _
                         "  ,TipoDeConta             = " & IIf(Me.TipodeConta.Length = 0, "NULL", "'" & Me.TipodeConta & "'") & "" & vbCrLf & _
                         "  ,Adiantamento            = " & IIf(Me.Adiantamento, 1, 0) & vbCrLf & _
                         "  ,Pedido                  = " & IIf(Me.TemPedido, 1, 0) & vbCrLf & _
                         "  ,Encargo                 = " & IIf(Me.Encargo, 1, 0) & vbCrLf & _
                         "  ,TemEncargo              = " & IIf(Me.TemEncargo, 1, 0) & vbCrLf & _
                         "  ,Ativo                   = " & IIf(Me.Ativo, 1, 0) & vbCrLf & _
                         "  ,UsuarioAlteracao        ='" & Me.UsuarioAlteracao & "'" & vbCrLf & _
                         "  ,UsuarioAlteracaoData    ='" & Me.UsuarioAlteracaoData.ToString("yyyy-MM-dd H:mm:ss") & "'" & vbCrLf & _
                         "  ,ProdutoParaCusto        = " & IIf(Me.ProdutoParaCusto, 1, 0) & vbCrLf & _
                         " WHERE Empresa_Id          ='" & Me.CodigoEmpresa & "'" & vbCrLf & _
                         "   AND EndEmpresa_Id       = " & Me.EnderecoEmpresa & vbCrLf & _
                         "   AND Conta_Id            ='" & Me.Conta & "'"
                sqls.Add(strSQL)
                SalvarTabelasRelacionada(sqls)
            Case "D"
                SalvarTabelasRelacionada(sqls)
                strSQL = "Delete PlanoDeContas " & vbCrLf & _
                         " WHERE Empresa_Id          ='" & Me.CodigoEmpresa & "'" & vbCrLf & _
                         "   AND EndEmpresa_Id       = " & Me.EnderecoEmpresa & vbCrLf & _
                         "   AND Conta_Id            ='" & Me.Conta & "'"
                sqls.Add(strSQL)
            Case Else
                SalvarTabelasRelacionada(sqls)
        End Select

        Return sqls
    End Function

    Public Function SalvarTabelasRelacionada(ByRef sqls As ArrayList) As ArrayList
        Me.EncargosPlanoDeContas.SalvarSql(sqls)
        Return sqls
    End Function

#End Region

End Class