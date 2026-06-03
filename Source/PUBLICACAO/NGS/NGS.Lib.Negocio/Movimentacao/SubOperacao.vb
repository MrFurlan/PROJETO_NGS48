Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Collections.Generic
Imports NGS.Lib.Uteis

<Serializable()> _
Public Class ListSubOperacao
    Inherits List(Of SubOperacao)

    Public Erro As Exception

#Region "Construtor"
    Public Sub New()

    End Sub

    Public Sub New(ByVal Operacao As Integer)
        Selecionar(Operacao)
    End Sub

    Public Sub New(ByVal Where As String)
        Selecionar(0, Where)
    End Sub
#End Region

    Public Function Selecionar(ByVal Operacao As Integer, Optional ByVal Where As String = "") As Boolean
        Dim objDados As New AcessaBanco()

        Try
            Dim strSQL As String = "SELECT SO.Operacao_Id, SO.SubOperacoes_Id, SO.Descricao, SO.EntradaSaida, SO.Devolucao, SO.Classe, SO.PrecoFixo, SO.Laudo, " & vbCrLf &
                                   "       SO.EstoqueInicial, SO.EstoqueFisico, SO.EstoqueFiscal, SO.QuantidadeFisico, SO.QuantidadeFiscal, " & vbCrLf &
                                   "       SO.QuantidadePedido, SO.UnitarioPedido, SO.GrupoDeContas, SO.Financeiro, SO.Contabil,  " & vbCrLf &
                                   "       SO.ApuracaoDeCustos, SO.ApuracaodeCustosContraPartida, SO.Deposito, " & vbCrLf &
                                   "       isnull(SO.Situacao,1) as Situacao, isnull(SO.ControlarPecas,0) as ControlarPecas, " & vbCrLf &
                                   "       isnull(SO.Memorando,0) as Memorando, isnull(SO.consignacao,0) as consignacao, isnull(AmostraGratis, 0) as AmostraGratis," & vbCrLf &
                                   "       isnull(SO.OperacaoDestino,0) as OperacaoDestino, isnull(SO.SubOperacaoDestino,0) as SubOperacaoDestino, " & vbCrLf &
                                   "       isnull(SO.Pedido,0) as Pedido,  isnull(SO.CobraServico,0) as CobraServico," & vbCrLf &
                                   "       isnull(SO.Liminar,0) as Liminar, isnull(SO.ProdutoDeTerceiro,0) as ProdutoDeTerceiro, " & vbCrLf &
                                   "       Isnull(SO.FinalidadeDaNota,0) FinalidadeDaNota, ContaDeAdiantamento, isnull(Representante,0) AS Representante, " & vbCrLf &
                                   "       isnull(ControlarNumeroDoLote,0) AS ControlarNumeroDoLote, isnull(ProprietarioDaMercadoria,0) AS ProprietarioDaMercadoria, " & vbCrLf &
                                   "       SO.NotaDebito, SO.NotaCredito, " & vbCrLf &
                                   "       SO.UsuarioInclusao, SO.UsuarioAlteracao " & vbCrLf &
                                   "  FROM SubOperacoes SO " & vbCrLf &
                                   " Inner Join Operacoes OP" & vbCrLf &
                                   "    on SO.Operacao_id = OP.Operacao_Id " & vbCrLf

            If Operacao > 0 Then
                strSQL &= " WHERE SO.Operacao_Id = " & Operacao.ToString() & " "
                If Where.Length > 0 Then
                    strSQL &= " and " & Where
                End If
            ElseIf Where.Length > 0 Then
                strSQL &= " Where " & Where
            End If
            strSQL &= " ORDER BY SO.Operacao_id, SO.SubOperacoes_Id"

            Dim dsSubOperacoes As DataSet = objDados.ConsultaDataSet(strSQL, "SubOperacoes")

            For Each drSubOperacao As DataRow In dsSubOperacoes.Tables(0).Rows
                Dim objSubOperacao As New SubOperacao(CInt(drSubOperacao("Operacao_Id")))

                objSubOperacao.CodigoOperacao = drSubOperacao("Operacao_Id")
                objSubOperacao.Codigo = drSubOperacao("SubOperacoes_Id")
                objSubOperacao.Descricao = drSubOperacao("Descricao")
                objSubOperacao.EntradaSaida = IIf(drSubOperacao("EntradaSaida") = "E", eEntradaSaida.Entrada, eEntradaSaida.Saida)
                objSubOperacao.Devolucao = drSubOperacao("Devolucao") = "S"
                objSubOperacao.Classe = Conversoes.ConverterClasseOperacao(drSubOperacao("Classe"))
                objSubOperacao.PrecoFixo = drSubOperacao("PrecoFixo") = "S"
                objSubOperacao.Laudo = drSubOperacao("Laudo") = "S"
                objSubOperacao.EstoqueInicial = drSubOperacao("EstoqueInicial") = "S"
                objSubOperacao.EstoqueFisico = drSubOperacao("EstoqueFisico") = "S"
                objSubOperacao.EstoqueFiscal = drSubOperacao("EstoqueFiscal") = "S"
                objSubOperacao.QuantidadeFisico = drSubOperacao("QuantidadeFisico") = "S"
                objSubOperacao.QuantidadeFiscal = drSubOperacao("QuantidadeFiscal") = "S"
                objSubOperacao.QuantidadePedido = drSubOperacao("QuantidadePedido") = "S"
                objSubOperacao.UnitarioPedido = drSubOperacao("UnitarioPedido") = "S"
                objSubOperacao.CodigoGrupoContas = drSubOperacao("GrupoDeContas")
                objSubOperacao.Financeiro = drSubOperacao("Financeiro") = "S"
                objSubOperacao.Contabil = drSubOperacao("Contabil") = "S"
                objSubOperacao.ApuracaoCustos = drSubOperacao("ApuracaoDeCustos")
                If Not drSubOperacao.IsNull("ApuracaodeCustosContraPartida") Then objSubOperacao.ApuracaoCustosContraPartida = drSubOperacao("ApuracaodeCustosContraPartida")
                objSubOperacao.Deposito = drSubOperacao("Deposito") = "S"
                objSubOperacao.CodigoSituacao = drSubOperacao("Situacao")
                objSubOperacao.ControlarPecas = drSubOperacao("ControlarPecas")

                objSubOperacao.Memorando = drSubOperacao("Memorando")
                objSubOperacao.Consignacao = drSubOperacao("Consignacao")
                objSubOperacao.AmostraGratis = drSubOperacao("AmostraGratis")
                objSubOperacao.CodigoOperacaoDestino = drSubOperacao("OperacaoDestino")
                objSubOperacao.CodigoSuboperacaoDestino = drSubOperacao("SubOperacaoDestino")
                objSubOperacao.Pedido = drSubOperacao("Pedido")
                objSubOperacao.CobraServico = drSubOperacao("CobraServico") = "S"
                objSubOperacao.Liminar = drSubOperacao("Liminar")
                objSubOperacao.ProdutoDeTerceiro = drSubOperacao("ProdutoDeTerceiro")
                objSubOperacao.FinalidadeDaNota = drSubOperacao("FinalidadeDaNota")
                objSubOperacao.Representante = drSubOperacao("Representante")
                If Not drSubOperacao.IsNull("ContaDeAdiantamento") Then objSubOperacao.CodigoContaAdiantamento = drSubOperacao("ContaDeAdiantamento")
                objSubOperacao.ControlarNumeroDoLote = drSubOperacao("ControlarNumeroDoLote")
                objSubOperacao.ProprietarioDaMercadoria = drSubOperacao("ProprietarioDaMercadoria")
                objSubOperacao.NotaDebito = drSubOperacao("NotaDebito") = "S"
                objSubOperacao.NotaCredito = drSubOperacao("NotaCredito") = "S"
                If Not drSubOperacao.IsNull("UsuarioInclusao") Then objSubOperacao.UsuarioInclusao = drSubOperacao("UsuarioInclusao")
                If Not drSubOperacao.IsNull("UsuarioAlteracao") Then objSubOperacao.UsuarioAlteracao = drSubOperacao("UsuarioAlteracao")




                Me.Add(objSubOperacao)
            Next

            Return True
        Catch ex As Exception
            Me.Erro = ex
            Return False
        Finally
            objDados = Nothing
        End Try
    End Function

End Class

<Serializable()> _
Public Class SubOperacao
    Implements IBaseEntity

    Public Erro As Exception

#Region "Construtor"
    Public Sub New()

    End Sub

    Public Sub New(ByVal Operacao As Integer)
        _CodigoOperacao = Operacao
        _Operacao = New Operacao(Operacao)
    End Sub

    Public Sub New(ByVal Operacao As Integer, ByVal SubOperacao As Integer)
        _CodigoOperacao = Operacao
        _Operacao = New Operacao(Operacao)
        Me.Codigo = SubOperacao
        Selecionar(Operacao, SubOperacao)
    End Sub

    Public Sub New(ByVal Operacao As Operacao)
        _CodigoOperacao = Operacao.Codigo
        _Operacao = Operacao
    End Sub

    Public Sub New(ByVal Operacao As Operacao, ByVal SubOperacao As Integer)
        _CodigoOperacao = Operacao.Codigo
        _Operacao = Operacao
        Me.Codigo = SubOperacao
        Selecionar(Operacao.Codigo, SubOperacao)
    End Sub

    Public Sub New(ByVal Operacao As Integer, ByVal SubOperacao As Integer, ByVal CodigoFiscal As Integer)
        _CodigoOperacao = Operacao
        _Operacao = New Operacao(Operacao)
        Me.Codigo = SubOperacao
        Me.CodigoFiscal = CodigoFiscal
        Selecionar(Operacao, SubOperacao)
    End Sub

#End Region

#Region "Fields"
    Private _IUD As String
    Private _CodigoOperacao As Integer
    Private _Operacao As Operacao
    Private _Codigo As Integer
    Private _Descricao As String = ""
    Private _EntradaSaida As eEntradaSaida
    Private _Devolucao As Boolean
    Private _Classe As eClassesOperacoes
    Private _PrecoFixo As Boolean
    Private _Laudo As Boolean
    Private _EstoqueInicial As Boolean
    Private _EstoqueFisico As Boolean
    Private _EstoqueFiscal As Boolean
    Private _QuantidadeFisico As Boolean
    Private _QuantidadeFiscal As Boolean
    Private _QuantidadePedido As Boolean
    Private _UnitarioPedido As Boolean
    Private _CodigoGrupoDeContas As String
    Private _GrupoDeConta As PlanoDeConta
    Private _Financeiro As Boolean
    Private _Contabil As Boolean
    Private _ApuracaoDeCustos As Integer
    Private _ApuracaodeCustosContraPartida As Integer
    Private _Deposito As Boolean
    Private _codigoSituacao As Integer
    Private _Situacao As Situacao
    Private _ControlarPecas As Boolean

    Private _Memorando As Boolean
    Private _Consignacao As Boolean
    Private _AmostraGratis As Boolean
    Private _CodigoOperacaoDestino As Integer
    Private _CodigoSuboperacaoDestino As Integer
    Private _OperacaoDestino As SubOperacao
    Private _Pedido As Boolean
    Private _CobraServico As Boolean
    Private _Liminar As Boolean
    Private _ProdutoDeTerceiro As Boolean
    Private _ParaFinsDeExportacao As Boolean
    Private _FinalidadeDaNota As Integer

    Private _CodigoContaAdiantamento As String
    Private _ContaAdiantamento As PlanoDeConta

    Private _Representante As Boolean

    Private _ControlarNumeroDoLote As Boolean 'Para produto Quimico

    Private _ProprietarioDaMercadoria As Boolean

    Private _CodigoFiscal As Integer
    Private _NotaDebito As Boolean
    Private _NotaCredito As Boolean
    Private _UsuarioInclusao As String
    Private _UsuarioAlteracao As String
    Private _Selecionado As Boolean = False




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

    Public Property CodigoOperacao() As Integer
        Get
            Return _CodigoOperacao
        End Get
        Set(ByVal value As Integer)
            _CodigoOperacao = value
        End Set
    End Property

    Public ReadOnly Property Operacao() As Operacao
        Get
            If _Operacao Is Nothing And _CodigoOperacao > 0 Then _Operacao = New Operacao(_CodigoOperacao)
            Return _Operacao
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

    Public Property Descricao() As String
        Get
            Return _Descricao
        End Get
        Set(ByVal value As String)
            _Descricao = value
        End Set
    End Property

    Public Property EntradaSaida() As eEntradaSaida
        Get
            Return _EntradaSaida
        End Get
        Set(ByVal value As eEntradaSaida)
            _EntradaSaida = value
        End Set
    End Property

    Public Property Devolucao() As Boolean
        Get
            Return _Devolucao
        End Get
        Set(ByVal value As Boolean)
            _Devolucao = value
        End Set
    End Property

    Public Property Classe() As eClassesOperacoes
        Get
            Return _Classe
        End Get
        Set(ByVal value As eClassesOperacoes)
            _Classe = value
        End Set
    End Property


    Public Property PrecoFixo() As Boolean
        Get
            Return _PrecoFixo
        End Get
        Set(ByVal value As Boolean)
            _PrecoFixo = value
        End Set
    End Property

    Public Property Laudo() As Boolean
        Get
            Return _Laudo
        End Get
        Set(ByVal value As Boolean)
            _Laudo = value
        End Set
    End Property

    Public Property EstoqueInicial() As Boolean
        Get
            Return _EstoqueInicial
        End Get
        Set(ByVal value As Boolean)
            _EstoqueInicial = value
        End Set
    End Property

    Public Property EstoqueFisico() As Boolean
        Get
            Return _EstoqueFisico
        End Get
        Set(ByVal value As Boolean)
            _EstoqueFisico = value
        End Set
    End Property

    Public Property EstoqueFiscal() As Boolean
        Get
            Return _EstoqueFiscal
        End Get
        Set(ByVal value As Boolean)
            _EstoqueFiscal = value
        End Set
    End Property

    Public Property QuantidadeFisico() As Boolean
        Get
            Return _QuantidadeFisico
        End Get
        Set(ByVal value As Boolean)
            _QuantidadeFisico = value
        End Set
    End Property

    Public Property QuantidadeFiscal() As Boolean
        Get
            Return _QuantidadeFiscal
        End Get
        Set(ByVal value As Boolean)
            _QuantidadeFiscal = value
        End Set
    End Property

    Public Property QuantidadePedido() As Boolean
        Get
            Return _QuantidadePedido
        End Get
        Set(ByVal value As Boolean)
            _QuantidadePedido = value
        End Set
    End Property

    Public Property UnitarioPedido() As Boolean
        Get
            Return _UnitarioPedido
        End Get
        Set(ByVal value As Boolean)
            _UnitarioPedido = value
        End Set
    End Property

    Public Property CodigoGrupoContas() As String
        Get
            Return _CodigoGrupoDeContas
        End Get
        Set(ByVal value As String)
            _CodigoGrupoDeContas = value
            _GrupoDeConta = Nothing
        End Set
    End Property

    Public Property GrupoDeConta() As PlanoDeConta
        Get
            If _GrupoDeConta Is Nothing And _CodigoGrupoDeContas.Length > 0 Then _GrupoDeConta = New PlanoDeConta("", 0, _CodigoGrupoDeContas)
            Return _GrupoDeConta
        End Get
        Set(ByVal value As PlanoDeConta)
            _GrupoDeConta = value
        End Set
    End Property

    Public Property Financeiro() As Boolean
        Get
            Return _Financeiro
        End Get
        Set(ByVal value As Boolean)
            _Financeiro = value
        End Set
    End Property

    Public Property Contabil() As Boolean
        Get
            Return _Contabil
        End Get
        Set(ByVal value As Boolean)
            _Contabil = value
        End Set
    End Property

    Public Property ApuracaoCustos() As Integer
        Get
            Return _ApuracaoDeCustos
        End Get
        Set(ByVal value As Integer)
            _ApuracaoDeCustos = value
        End Set
    End Property

    Public Property ApuracaoCustosContraPartida() As Integer
        Get
            Return _ApuracaodeCustosContraPartida
        End Get
        Set(ByVal value As Integer)
            _ApuracaodeCustosContraPartida = value
        End Set
    End Property

    Public Property Deposito() As Boolean
        Get
            Return _Deposito
        End Get
        Set(ByVal value As Boolean)
            _Deposito = value
        End Set
    End Property

    Public Property CodigoSituacao() As Integer
        Get
            Return _codigoSituacao
        End Get
        Set(ByVal value As Integer)
            _codigoSituacao = value
            _Situacao = Nothing
        End Set
    End Property

    Public ReadOnly Property Situacao() As Situacao
        Get
            If _Situacao Is Nothing And _codigoSituacao > 0 Then _Situacao = New Situacao(_codigoSituacao)
            Return _Situacao
        End Get
    End Property

    Public Property ControlarPecas() As Boolean
        Get
            Return _ControlarPecas
        End Get
        Set(ByVal value As Boolean)
            _ControlarPecas = value
        End Set
    End Property

    Public Property Memorando() As Boolean
        Get
            Return _Memorando
        End Get
        Set(ByVal value As Boolean)
            _Memorando = value
        End Set
    End Property

    Public Property Consignacao() As Boolean
        Get
            Return _Consignacao
        End Get
        Set(ByVal value As Boolean)
            _Consignacao = value
        End Set
    End Property

    Public Property AmostraGratis() As Boolean
        Get
            Return _AmostraGratis
        End Get
        Set(ByVal value As Boolean)
            _AmostraGratis = value
        End Set
    End Property

    Public Property CodigoOperacaoDestino() As Integer
        Get
            Return _CodigoOperacaoDestino
        End Get
        Set(ByVal value As Integer)
            _CodigoOperacaoDestino = value
            _OperacaoDestino = Nothing
        End Set
    End Property

    Public Property CodigoSuboperacaoDestino() As Integer
        Get
            Return _CodigoSuboperacaoDestino
        End Get
        Set(ByVal value As Integer)
            _CodigoSuboperacaoDestino = value
            _OperacaoDestino = Nothing
        End Set
    End Property

    Public ReadOnly Property OperacaoDestino() As SubOperacao
        Get
            If _OperacaoDestino Is Nothing And _CodigoOperacaoDestino > 0 Then _OperacaoDestino = New SubOperacao(_CodigoOperacaoDestino, _CodigoSuboperacaoDestino)
            Return _OperacaoDestino
        End Get
    End Property

    Public Property Pedido() As Boolean
        Get
            Return _Pedido
        End Get
        Set(ByVal value As Boolean)
            _Pedido = value
        End Set
    End Property

    Public Property CobraServico() As Boolean
        Get
            Return _CobraServico
        End Get
        Set(ByVal value As Boolean)
            _CobraServico = value
        End Set
    End Property

    Public Property Liminar() As Boolean
        Get
            Return _Liminar
        End Get
        Set(ByVal value As Boolean)
            _Liminar = value
        End Set
    End Property

    Public Property ProdutoDeTerceiro() As Boolean
        Get
            Return _ProdutoDeTerceiro
        End Get
        Set(ByVal value As Boolean)
            _ProdutoDeTerceiro = value
        End Set
    End Property

    Public Property ParaFinsDeExportacao As Boolean
        Get
            Return _ParaFinsDeExportacao
        End Get
        Set(ByVal value As Boolean)
            _ParaFinsDeExportacao = value
        End Set
    End Property

    Public Property FinalidadeDaNota() As Integer
        Get
            Return _FinalidadeDaNota
        End Get
        Set(ByVal value As Integer)
            _FinalidadeDaNota = value
        End Set
    End Property

    Public Property CodigoContaAdiantamento() As String
        Get
            Return _CodigoContaAdiantamento
        End Get
        Set(ByVal value As String)
            _CodigoContaAdiantamento = value
            _ContaAdiantamento = Nothing
        End Set
    End Property

    Public Property ContaAdiantamento() As PlanoDeConta
        Get
            If _ContaAdiantamento Is Nothing And _CodigoContaAdiantamento.Length > 0 Then _ContaAdiantamento = New PlanoDeConta("", 0, _CodigoContaAdiantamento)
            Return _ContaAdiantamento
        End Get
        Set(ByVal value As PlanoDeConta)
            _ContaAdiantamento = value
        End Set
    End Property

    Public Property Representante() As Boolean
        Get
            Return _Representante
        End Get
        Set(ByVal value As Boolean)
            _Representante = value
        End Set
    End Property

    Public Property ControlarNumeroDoLote() As Boolean
        Get
            Return _ControlarNumeroDoLote
        End Get
        Set(ByVal value As Boolean)
            _ControlarNumeroDoLote = value
        End Set
    End Property

    Public Property ProprietarioDaMercadoria() As Boolean
        Get
            Return _ProprietarioDaMercadoria
        End Get
        Set(ByVal value As Boolean)
            _ProprietarioDaMercadoria = value
        End Set
    End Property

    Public Property CodigoFiscal() As Integer
        Get
            Return _CodigoFiscal
        End Get
        Set(ByVal value As Integer)
            _CodigoFiscal = value
        End Set
    End Property

    Public Property NotaDebito() As Boolean
        Get
            Return _NotaDebito
        End Get
        Set(ByVal value As Boolean)
            _NotaDebito = value
        End Set
    End Property

    Public Property NotaCredito() As Boolean
        Get
            Return _NotaCredito
        End Get
        Set(ByVal value As Boolean)
            _NotaCredito = value
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

    Public Property UsuarioAlteracao() As String
        Get
            Return _UsuarioAlteracao
        End Get
        Set(ByVal value As String)
            _UsuarioAlteracao = value
        End Set
    End Property

    Public Property Selecionado() As Boolean
        Get
            Return _Selecionado
        End Get
        Set(ByVal value As Boolean)
            _Selecionado = value
        End Set
    End Property

#End Region

#Region "Methods"
    Public Function Selecionar(ByVal Operacao As Integer, ByVal SubOperacao As Integer) As Boolean
        Dim objDados As New AcessaBanco()

        Try
            Dim strSQL As String = "SELECT SubOperacoes_Id, Descricao, EntradaSaida, Devolucao, Classe, PrecoFixo, Laudo, " &
                                   "       EstoqueInicial, EstoqueFisico, EstoqueFiscal, QuantidadeFisico, QuantidadeFiscal, " &
                                   "       QuantidadePedido, UnitarioPedido, GrupoDeContas, Financeiro, Contabil, " &
                                   "       ApuracaoDeCustos, isnull(ApuracaodeCustosContraPartida,0) as ApuracaodeCustosContraPartida, Deposito, " &
                                   "       isnull(Situacao,1) as Situacao, isnull(ControlarPecas,0) as ControlarPecas, " &
                                   "       isnull(Memorando,0) as Memorando, isnull(consignacao,0) as consignacao, isnull(AmostraGratis, 0) as AmostraGratis," &
                                   "       isnull(OperacaoDestino,0) as OperacaoDestino, isnull(SubOperacaoDestino,0) as SubOperacaoDestino," &
                                   "       isnull(CobraServico,0) as CobraServico, isnull(Liminar,0) as Liminar," &
                                   "       isnull(Pedido,0) as Pedido, isnull(ProdutoDeTerceiro,0) as ProdutoDeTerceiro," &
                                   "       isnull(ParaFinsDeExportacao,0) as ParaFinsDeExportacao, isnull(FinalidadeDaNota,1) as FinalidadeDaNota ," &
                                   "       isnull(ContaDeAdiantamento,'') as ContaDeAdiantamento, isnull(Representante,0) AS Representante," &
                                   "       isnull(ControlarNumeroDoLote,0) AS ControlarNumeroDoLote, isnull(ProprietarioDaMercadoria,0) AS ProprietarioDaMercadoria, " &
                                   "       NotaDebito, NotaCredito, " &
                                   "       UsuarioInclusao, UsuarioAlteracao " &
                                   "  FROM SubOperacoes " &
                                   " WHERE Operacao_Id     = " & Operacao.ToString() &
                                   "   AND SubOperacoes_Id = " & SubOperacao.ToString()

            Dim dsSubOperacoes As DataSet = objDados.ConsultaDataSet(strSQL, "SubOperacoes")

            If dsSubOperacoes.Tables(0).Rows.Count > 0 Then
                Dim drSubOperacao As DataRow = dsSubOperacoes.Tables(0).Rows(0)

                Me.Codigo = drSubOperacao("SubOperacoes_Id")
                Me.Descricao = drSubOperacao("Descricao")
                Me.EntradaSaida = IIf(drSubOperacao("EntradaSaida") = "E", eEntradaSaida.Entrada, eEntradaSaida.Saida)
                Me.Devolucao = drSubOperacao("Devolucao") = "S"
                Me.Classe = Conversoes.ConverterClasseOperacao(drSubOperacao("Classe"))
                Me.PrecoFixo = drSubOperacao("PrecoFixo") = "S"
                Me.Laudo = drSubOperacao("Laudo") = "S"
                Me.EstoqueInicial = drSubOperacao("EstoqueInicial") = "S"
                Me.EstoqueFisico = drSubOperacao("EstoqueFisico") = "S"
                Me.EstoqueFiscal = drSubOperacao("EstoqueFiscal") = "S"
                Me.QuantidadeFisico = drSubOperacao("QuantidadeFisico") = "S"
                Me.QuantidadeFiscal = drSubOperacao("QuantidadeFiscal") = "S"
                Me.QuantidadePedido = drSubOperacao("QuantidadePedido") = "S"
                Me.UnitarioPedido = drSubOperacao("UnitarioPedido") = "S"
                Me.CodigoGrupoContas = drSubOperacao("GrupoDeContas")
                Me.Financeiro = drSubOperacao("Financeiro") = "S"
                Me.Contabil = drSubOperacao("Contabil") = "S"
                Me.ApuracaoCustos = drSubOperacao("ApuracaoDeCustos")
                Me.ApuracaoCustosContraPartida = drSubOperacao("ApuracaodeCustosContraPartida")
                Me.Deposito = drSubOperacao("Deposito") = "S"

                Me.CodigoSituacao = drSubOperacao("Situacao")
                Me.ControlarPecas = drSubOperacao("ControlarPecas")

                Me.Memorando = drSubOperacao("Memorando")
                Me.Consignacao = drSubOperacao("Consignacao")
                Me.AmostraGratis = drSubOperacao("AmostraGratis")
                Me.CodigoOperacaoDestino = drSubOperacao("OperacaoDestino")
                Me.CodigoSuboperacaoDestino = drSubOperacao("SubOperacaoDestino")
                Me.Pedido = drSubOperacao("Pedido")
                Me.CobraServico = drSubOperacao("CobraServico") = "S"
                Me.Liminar = drSubOperacao("Liminar")
                Me.ProdutoDeTerceiro = drSubOperacao("ProdutoDeTerceiro")
                Me.ParaFinsDeExportacao = drSubOperacao("ParaFinsDeExportacao")
                Me.FinalidadeDaNota = drSubOperacao("FinalidadeDaNota")
                Me.CodigoContaAdiantamento = drSubOperacao("ContaDeAdiantamento")
                Me.Representante = drSubOperacao("Representante")
                Me.ControlarNumeroDoLote = drSubOperacao("ControlarNumeroDoLote")
                Me.ProprietarioDaMercadoria = drSubOperacao("ProprietarioDaMercadoria")
                Me.NotaDebito = drSubOperacao("NotaDebito") = "S"
                Me.NotaCredito = drSubOperacao("NotaCredito") = "S"
                Me.UsuarioInclusao = drSubOperacao("UsuarioInclusao")
                Me.UsuarioAlteracao = drSubOperacao("UsuarioAlteracao")
            End If

            Return True
        Catch ex As Exception
            Me.Erro = ex
            Return False
        Finally
            objDados = Nothing
        End Try
    End Function

    Public Function Salvar() As Boolean
        Dim objBanco As New AcessaBanco()
        Dim sqls As New ArrayList

        salvarSql(sqls)
        Return objBanco.GravaBanco(sqls)
    End Function

    Public Sub salvarSql(ByRef sqls As ArrayList)
        Dim strSql As String = ""

        Select Case IUD
            Case "I"
                strSql = " Insert into SubOperacoes (operacao_id , SubOperacoes_id , Descricao ," & vbCrLf &
                      " EntradaSaida , Devolucao , PrecoFixo ,Laudo , EstoqueInicial , EstoqueFisico ," & vbCrLf &
                      " EstoqueFiscal , QuantidadeFisico , QuantidadeFiscal ,QuantidadePedido ," & vbCrLf &
                      " UnitarioPedido , GrupoDeContas , Financeiro , Contabil , " & vbCrLf &
                      " ApuracaoDeCustos , ApuracaodeCustosContraPartida, Classe, Deposito," & vbCrLf &
                      " situacao, ControlarPecas, OperacaoDestino, SubOperacaoDestino," & vbCrLf &
                      " consignacao, AmostraGratis, Memorando, Pedido, CobraServico, " & vbCrLf &
                      " Liminar, ProdutoDeTerceiro, ParaFinsDeExportacao, FinalidadeDaNota, ContaDeAdiantamento, Representante, " & vbCrLf &
                      " UsuarioInclusao, UsuarioInclusaoData, ControlarNumeroDoLote, ProprietarioDaMercadoria, NotaDebito, NotaCredito ) " & vbCrLf &
                      " values (" & Me.CodigoOperacao & "," & Me.Codigo & ",'" & Me.Descricao & "'" & vbCrLf &
                      ",'" & IIf(Me.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "','" & IIf(Me.Devolucao, "S", "N") & "','" & IIf(Me.PrecoFixo, "S", "N") & "','" & IIf(Me.Laudo, "S", "N") & "','" & IIf(Me.EstoqueInicial, "S", "N") & "','" & IIf(Me.EstoqueFisico, "S", "N") & "'" & vbCrLf &
                      ",'" & IIf(Me.EstoqueFiscal, "S", "N") & "','" & IIf(Me.QuantidadeFisico, "S", "N") & "','" & IIf(Me.QuantidadeFiscal, "S", "N") & "'" & ",'" & IIf(Me.QuantidadePedido, "S", "N") & "'" & vbCrLf &
                      ",'" & IIf(Me.UnitarioPedido, "S", "N") & "','" & Me.CodigoGrupoContas & "','" & IIf(Me.Financeiro, "S", "N") & "','" & IIf(Me.Contabil, "S", "N") & "'" & vbCrLf &
                      "," & Me.ApuracaoCustos & "," & Me.ApuracaoCustosContraPartida & ",'" & Me.Classe.ToString & "','" & IIf(Me.Deposito, "S", "N") & "'" & vbCrLf &
                      "," & Me.CodigoSituacao & "," & CByte(Me.ControlarPecas) & "," & Me.CodigoOperacaoDestino & "," & Me.CodigoSuboperacaoDestino & vbCrLf &
                      "," & CByte(Me.Consignacao) & "," & CByte(Me.AmostraGratis) & "," & CByte(Me.Memorando) & "," & CByte(Me.Pedido) & ",'" & IIf(Me.CobraServico, "S", "N") & "'" & vbCrLf &
                      "," & CByte(Me.Liminar) & "," & CByte(Me.ProdutoDeTerceiro) & "," & CByte(Me.ParaFinsDeExportacao) & "," & CByte(Me.FinalidadeDaNota) & ",'" & Me.CodigoContaAdiantamento & "'" & vbCrLf &
                      "," & CByte(Me.Representante) & vbCrLf &
                      ",'" & UsuarioServidor.NomeUsuario & "','" & DateTime.Now.ToString("yyyy-MM-dd") & "'," & CByte(Me.ControlarNumeroDoLote) & "," & CByte(Me.ProprietarioDaMercadoria) & ",'" & IIf(Me.NotaDebito, "S", "N") & "','" & IIf(Me.NotaCredito, "S", "N") & "')"
                sqls.Add(strSql)
            Case "U"
                strSql = "UPDATE SubOperacoes set Descricao = '" & Me.Descricao & "', EntradaSaida = '" & IIf(Me.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "'" & vbCrLf &
                      ", Devolucao = '" & IIf(Me.Devolucao, "S", "N") & "', PrecoFixo = '" & IIf(Me.PrecoFixo, "S", "N") & "'" & vbCrLf &
                      ", Laudo = '" & IIf(Me.Laudo, "S", "N") & "', EstoqueInicial = '" & IIf(Me.EstoqueInicial, "S", "N") & "'" & vbCrLf &
                      ", EstoqueFisico = '" & IIf(Me.EstoqueFisico, "S", "N") & "', EstoqueFiscal = '" & IIf(Me.EstoqueFiscal, "S", "N") & "'" & vbCrLf &
                      ", QuantidadeFisico = '" & IIf(Me.QuantidadeFisico, "S", "N") & "', QuantidadeFiscal = '" & IIf(Me.QuantidadeFiscal, "S", "N") & "'" & vbCrLf &
                      ", QuantidadePedido = '" & IIf(Me.QuantidadePedido, "S", "N") & "', UnitarioPedido = '" & IIf(Me.UnitarioPedido, "S", "N") & "'" & vbCrLf &
                      ", GrupoDeContas = '" & Me.CodigoGrupoContas & "', Financeiro = '" & IIf(Me.Financeiro, "S", "N") & "', Classe = '" & Me.Classe.ToString & "'" & vbCrLf &
                      ", Contabil = '" & IIf(Me.Contabil, "S", "N") & "', Deposito = '" & IIf(Me.Deposito, "S", "N") & "'" & vbCrLf &
                      ", ApuracaoDeCustos = " & Me.ApuracaoCustos & vbCrLf &
                      ", ApuracaodeCustosContraPartida = " & Me.ApuracaoCustosContraPartida & vbCrLf &
                      " ,Situacao                 = " & Me.CodigoSituacao & vbCrLf &
                      " ,ControlarPecas           = " & CByte(Me.ControlarPecas) & vbCrLf &
                      " ,Consignacao              = " & CByte(Me.Consignacao) & vbCrLf &
                      " ,AmostraGratis            = " & CByte(Me.AmostraGratis) & vbCrLf &
                      " ,Memorando                = " & CByte(Me.Memorando) & vbCrLf &
                      " ,Pedido                   = " & CByte(Me.Pedido) & vbCrLf &
                      " ,CobraServico             ='" & IIf(Me.CobraServico, "S", "N") & "'" & vbCrLf &
                      " ,Liminar                  = " & CByte(Me.Liminar) & vbCrLf &
                      " ,ProdutoDeTerceiro        = " & CByte(Me.ProdutoDeTerceiro) & vbCrLf &
                      " ,ParaFinsDeExportacao     = " & CByte(Me.ParaFinsDeExportacao) & vbCrLf &
                      " ,OperacaoDestino          = " & Me.CodigoOperacaoDestino & vbCrLf &
                      " ,SubOperacaoDestino       = " & Me.CodigoSuboperacaoDestino & vbCrLf &
                      " ,FinalidadeDaNota         = " & Str(Me.FinalidadeDaNota) & vbCrLf &
                      " ,ContaDeAdiantamento      = '" & Me.CodigoContaAdiantamento & "'" & vbCrLf &
                      " ,Representante            = " & CByte(Me.Representante) & vbCrLf &
                      " ,UsuarioAlteracao         = '" & UsuarioServidor.NomeUsuario & "'" & vbCrLf &
                      " ,UsuarioAlteracaoData     = '" & DateTime.Now.ToString("yyyy-MM-dd") & "'" & vbCrLf &
                      " ,ControlarNumeroDoLote    = " & CByte(Me.ControlarNumeroDoLote) & vbCrLf &
                      " ,ProprietarioDaMercadoria = " & CByte(Me.ProprietarioDaMercadoria) & vbCrLf &
                      " ,NotaDebito               ='" & IIf(Me.NotaDebito, "S", "N") & "'" & vbCrLf &
                      " ,NotaCredito              ='" & IIf(Me.NotaCredito, "S", "N") & "'" & vbCrLf &
                      "  WHERE (Operacao_id   = " & Me.CodigoOperacao & " and SubOperacoes_id = " & Me.Codigo & ")" & vbCrLf
                sqls.Add(strSql)
            Case "D"
                strSql = "Delete from SubOperacoes Where (Operacao_id = " & Me.CodigoOperacao & " and SubOperacoes_id = " & Me.Codigo & ")"
                sqls.Add(strSql)
        End Select
    End Sub
#End Region

End Class