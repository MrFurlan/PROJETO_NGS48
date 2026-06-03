Imports Microsoft.VisualBasic
Imports System.Data
Imports NGS.Lib.Uteis
Imports System.Web

<Serializable()>
Public Class MovimentacaoFinanceira
    Implements IBaseEntity

#Region "Variáveis locais"

    Private _IUD As String
    Private _EntradaSaida As eEntradaSaida
    Private _Codigo As Integer
    Private _Sequencia As Integer
    Private _Provisao As eProvisao
    Private _Carteira As String
    Private _Tributo As String
    Private _Indexador As Integer
    Private _CodigoMoeda As Integer
    Private _Moeda As Moeda
    Private _TipoPagamento As Integer
    Private _Situacao As Integer
    Private _Lote As Integer
    Private _Movimento As DateTime
    Private _Vencimento As DateTime
    Private _Prorrogacao As DateTime
    Private _DataMoeda As DateTime
    Private _Baixa As DateTime
    Private _UnidadeNegocio As String
    Private _Empresa As String
    Private _EnderecoEmpresa As Integer
    Private _Cliente As String
    Private _EnderecoCliente As Integer
    Private _BancoCliente As Integer
    Private _AgenciaCliente As String
    Private _DigitoAgenciaCliente As String
    Private _ContaCliente As String
    Private _DigitoContaCliente As String
    Private _ContaContabilCliente As String
    Private _EmpresaPagadora As String
    Private _EnderecoEmpresaPagadora As Integer
    Private _BancoPagador As Integer
    Private _AgenciaPagadora As String
    Private _DigitoAgenciaPagadora As String
    Private _ContaPagadora As String
    Private _DigitoContaPagadora As String
    Private _ContaContabilPagadora As String
    Private _Cheque As Boolean
    Private _Slips As Boolean
    Private _Recibo As Boolean
    Private _Aviso As Boolean
    Private _ReciboDeposito As Boolean
    Private _EmpresaPedido As String
    Private _EnderecoEmpresaPedido As Integer
    Private _Pedido As Integer
    Private _PedidoFixacao As Integer
    Private _Procuracao As Integer
    Private _ValorDocumento As Decimal
    Private _Descontos As Decimal
    Private _Deducoes As Decimal
    Private _Juros As Decimal
    Private _Acrescimos As Decimal
    Private _ValorLiquido As Decimal
    Private _MoedaValorDocumento As Decimal
    Private _MoedaDescontos As Decimal
    Private _MoedaDeducoes As Decimal
    Private _MoedaJuros As Decimal
    Private _MoedaAcrescimos As Decimal
    Private _MoedaValorLiquido As Decimal
    Private _Historico As String
    Private _CodigoBarras As String
    Private _CodigoDigitado As Boolean
    Private _CodigoDestinatario As String
    Private _EnderecoDestinatario As Integer
    Private _Destinatario As Cliente
    Private _Destinacao As String
    Private _Solicitacao As Integer
    Private _UsuarioInclusao As String
    Private _UsuarioInclusaoData As DateTime
    Private _UsuarioAlteracao As String
    Private _UsuarioAlteracaoData As DateTime
    Private _UsuarioCancelamento As String
    Private _UsuarioCancelamentoData As DateTime
    Private _UsuarioLiberacao As String = ""
    Private _UsuarioLiberacaoData As DateTime
    Private _UsuarioBaixa As String
    Private _UsuarioBaixaData As DateTime
    Private _Grupado As Boolean
    Private _RegistroMestre As Integer
    Private _Observacoes As String
    Private _SituacaoBancaria As Integer
    Private _TipoFinanceiro As String = ""  'P - Pagar - R - Receber
    Private _CarteiraDoTitulo As Integer
    Private _ContratoBancario As String = ""

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
    Public Property EntradaSaida() As eEntradaSaida
        Get
            Return _EntradaSaida
        End Get
        Set(ByVal value As eEntradaSaida)
            _EntradaSaida = value
        End Set
    End Property

    Public Property Codigo() As Integer
        Get
            Return _Codigo
        End Get
        Set(ByVal value As Integer)
            _Codigo = value
        End Set
    End Property

    Public Property Sequencia() As Integer
        Get
            Return _Sequencia
        End Get
        Set(ByVal value As Integer)
            _Sequencia = value
        End Set
    End Property

    Public Property Provisao() As eProvisao
        Get
            Return _Provisao
        End Get
        Set(ByVal value As eProvisao)
            _Provisao = value
        End Set
    End Property

    Public Property Carteira() As String
        Get
            Return _Carteira
        End Get
        Set(ByVal value As String)
            _Carteira = value
        End Set
    End Property

    Public Property Tributo() As String
        Get
            Return _Tributo
        End Get
        Set(ByVal value As String)
            _Tributo = value
        End Set
    End Property

    Public Property Indexador() As Integer
        Get
            Return _Indexador
        End Get
        Set(ByVal value As Integer)
            _Indexador = value
        End Set
    End Property

    Public Property CodigoMoeda() As Integer
        Get
            Return _CodigoMoeda
        End Get
        Set(ByVal value As Integer)
            _CodigoMoeda = value
            _Moeda = Nothing
        End Set
    End Property

    Public ReadOnly Property Moeda As Moeda
        Get
            If _Moeda Is Nothing And Me.CodigoMoeda > 0 Then _Moeda = New Moeda(Me.CodigoMoeda)
            Return _Moeda
        End Get
    End Property

    Public Property TipoPagamento() As Integer
        Get
            Return _TipoPagamento
        End Get
        Set(ByVal value As Integer)
            _TipoPagamento = value
        End Set
    End Property

    Public Property Situacao() As Integer
        Get
            Return _Situacao
        End Get
        Set(ByVal value As Integer)
            _Situacao = value
        End Set
    End Property

    Public Property Lote() As Integer
        Get
            Return _Lote
        End Get
        Set(ByVal value As Integer)
            _Lote = value
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

    Public Property Vencimento() As DateTime
        Get
            Return _Vencimento
        End Get
        Set(ByVal value As DateTime)
            _Vencimento = value
        End Set
    End Property

    Public Property DataProrrogacao() As DateTime
        Get
            Return _Prorrogacao
        End Get
        Set(ByVal value As DateTime)
            _Prorrogacao = value
        End Set
    End Property

    Public Property DataMoeda() As DateTime
        Get
            Return _DataMoeda
        End Get
        Set(ByVal value As DateTime)
            _DataMoeda = value
        End Set
    End Property

    Public Property DataBaixa() As DateTime
        Get
            Return _Baixa
        End Get
        Set(ByVal value As DateTime)
            _Baixa = value
        End Set
    End Property

    Public Property UnidadeNegocio() As String
        Get
            Return _UnidadeNegocio
        End Get
        Set(ByVal value As String)
            _UnidadeNegocio = value
        End Set
    End Property

    Public Property CodigoEmpresa() As String
        Get
            Return _Empresa
        End Get
        Set(ByVal value As String)
            _Empresa = value
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
            Return _Cliente
        End Get
        Set(ByVal value As String)
            _Cliente = value
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

    Public Property CodigoBancoCliente() As Integer
        Get
            Return _BancoCliente
        End Get
        Set(ByVal value As Integer)
            _BancoCliente = value
        End Set
    End Property

    Public Property AgenciaCliente() As String
        Get
            Return _AgenciaCliente
        End Get
        Set(ByVal value As String)
            _AgenciaCliente = value
        End Set
    End Property

    Public Property DigitoAgenciaCliente() As String
        Get
            Return _DigitoAgenciaCliente
        End Get
        Set(ByVal value As String)
            _DigitoAgenciaCliente = value
        End Set
    End Property

    Public Property ContaCliente() As String
        Get
            Return _ContaCliente
        End Get
        Set(ByVal value As String)
            _ContaCliente = value
        End Set
    End Property

    Public Property DigitoContaCliente() As String
        Get
            Return _DigitoContaCliente
        End Get
        Set(ByVal value As String)
            _DigitoContaCliente = value
        End Set
    End Property

    Public Property ContaContabilCliente() As String
        Get
            Return _ContaContabilCliente
        End Get
        Set(ByVal value As String)
            _ContaContabilCliente = value
        End Set
    End Property

    Public Property CodigoEmpresaPagadora() As String
        Get
            Return _EmpresaPagadora
        End Get
        Set(ByVal value As String)
            _EmpresaPagadora = value
        End Set
    End Property

    Public Property EnderecoEmpresaPagadora() As Integer
        Get
            Return _EnderecoEmpresaPagadora
        End Get
        Set(ByVal value As Integer)
            _EnderecoEmpresaPagadora = value
        End Set
    End Property

    Public Property BancoPagador() As Integer
        Get
            Return _BancoPagador
        End Get
        Set(ByVal value As Integer)
            _BancoPagador = value
        End Set
    End Property

    Public Property AgenciaPagadora() As String
        Get
            Return _AgenciaPagadora
        End Get
        Set(ByVal value As String)
            _AgenciaPagadora = value
        End Set
    End Property

    Public Property DigitoAgenciaPagadora() As String
        Get
            Return _DigitoAgenciaPagadora
        End Get
        Set(ByVal value As String)
            _DigitoAgenciaPagadora = value
        End Set
    End Property

    Public Property ContaPagadora() As String
        Get
            Return _ContaPagadora
        End Get
        Set(ByVal value As String)
            _ContaPagadora = value
        End Set
    End Property

    Public Property DigitoContaPagadora() As String
        Get
            Return _DigitoContaPagadora
        End Get
        Set(ByVal value As String)
            _DigitoContaPagadora = value
        End Set
    End Property

    Public Property ContaContabilPagadora() As String
        Get
            Return _ContaContabilPagadora
        End Get
        Set(ByVal value As String)
            _ContaContabilPagadora = value
        End Set
    End Property

    Public Property Cheque() As Boolean
        Get
            Return _Cheque
        End Get
        Set(ByVal value As Boolean)
            _Cheque = value
        End Set
    End Property

    Public Property Slips() As Boolean
        Get
            Return _Slips
        End Get
        Set(ByVal value As Boolean)
            _Slips = value
        End Set
    End Property

    Public Property Recibo() As Boolean
        Get
            Return _Recibo
        End Get
        Set(ByVal value As Boolean)
            _Recibo = value
        End Set
    End Property

    Public Property Aviso() As Boolean
        Get
            Return _Aviso
        End Get
        Set(ByVal value As Boolean)
            _Aviso = value
        End Set
    End Property

    Public Property ReciboDeposito() As Boolean
        Get
            Return _ReciboDeposito
        End Get
        Set(ByVal value As Boolean)
            _ReciboDeposito = value
        End Set
    End Property

    Public Property CodigoEmpresaPedido() As String
        Get
            Return _EmpresaPedido
        End Get
        Set(ByVal value As String)
            _EmpresaPedido = value
        End Set
    End Property

    Public Property EnderecoEmpresaPedido() As Integer
        Get
            Return _EnderecoEmpresaPedido
        End Get
        Set(ByVal value As Integer)
            _EnderecoEmpresaPedido = value
        End Set
    End Property

    Public Property CodigoPedido() As Integer
        Get
            Return _Pedido
        End Get
        Set(ByVal value As Integer)
            _Pedido = value
        End Set
    End Property

    Public Property PedidoFixacao() As Integer
        Get
            Return _PedidoFixacao
        End Get
        Set(ByVal value As Integer)
            _PedidoFixacao = value
        End Set
    End Property

    Public Property Procuracao() As Integer
        Get
            Return _Procuracao
        End Get
        Set(ByVal value As Integer)
            _Procuracao = value
        End Set
    End Property

    Public Property ValorDocumentoOficial() As Decimal
        Get
            Return _ValorDocumento
        End Get
        Set(ByVal value As Decimal)
            _ValorDocumento = value
        End Set
    End Property

    Public Property DescontosOficial() As Decimal
        Get
            Return _Descontos
        End Get
        Set(ByVal value As Decimal)
            _Descontos = value
        End Set
    End Property

    Public Property DeducoesOficial() As Decimal
        Get
            Return _Deducoes
        End Get
        Set(ByVal value As Decimal)
            _Deducoes = value
        End Set
    End Property

    Public Property JurosOficial() As Decimal
        Get
            Return _Juros
        End Get
        Set(ByVal value As Decimal)
            _Juros = value
        End Set
    End Property

    Public Property AcrescimosOficial() As Decimal
        Get
            Return _Acrescimos
        End Get
        Set(ByVal value As Decimal)
            _Acrescimos = value
        End Set
    End Property

    Public Property ValorLiquidoOficial() As Decimal
        Get
            Return _ValorLiquido
        End Get
        Set(ByVal value As Decimal)
            _ValorLiquido = value
        End Set
    End Property

    Public Property ValorDocumentoMoeda() As Decimal
        Get
            Return _MoedaValorDocumento
        End Get
        Set(ByVal value As Decimal)
            _MoedaValorDocumento = value
        End Set
    End Property

    Public Property DescontosMoeda() As Decimal
        Get
            Return _MoedaDescontos
        End Get
        Set(ByVal value As Decimal)
            _MoedaDescontos = value
        End Set
    End Property

    Public Property DeducoesMoeda() As Decimal
        Get
            Return _MoedaDeducoes
        End Get
        Set(ByVal value As Decimal)
            _MoedaDeducoes = value
        End Set
    End Property

    Public Property JurosMoeda() As Decimal
        Get
            Return _MoedaJuros
        End Get
        Set(ByVal value As Decimal)
            _MoedaJuros = value
        End Set
    End Property

    Public Property AcrescimosMoeda() As Decimal
        Get
            Return _MoedaAcrescimos
        End Get
        Set(ByVal value As Decimal)
            _MoedaAcrescimos = value
        End Set
    End Property

    Public Property ValorLiquidoMoeda() As Decimal
        Get
            Return _MoedaValorLiquido
        End Get
        Set(ByVal value As Decimal)
            _MoedaValorLiquido = value
        End Set
    End Property

    Public Property Historico() As String
        Get
            Return _Historico
        End Get
        Set(ByVal value As String)
            _Historico = value
        End Set
    End Property

    Public Property CodigoBarras() As String
        Get
            Return _CodigoBarras
        End Get
        Set(ByVal value As String)
            _CodigoBarras = value
        End Set
    End Property

    Public Property CodigoDigitado() As Boolean
        Get
            Return _CodigoDigitado
        End Get
        Set(ByVal value As Boolean)
            _CodigoDigitado = value
        End Set
    End Property

    Public Property CodigoDestinatario() As String
        Get
            Return _CodigoDestinatario
        End Get
        Set(ByVal value As String)
            _CodigoDestinatario = value
        End Set
    End Property

    Public Property EnderecoDestinatario() As Integer
        Get
            Return _EnderecoDestinatario
        End Get
        Set(ByVal value As Integer)
            _EnderecoDestinatario = value
        End Set
    End Property

    Public Property Destinatario() As Cliente
        Get
            If _Destinatario Is Nothing And Me.CodigoDestinatario.Length > 0 Then _Destinatario = New Cliente(Me.CodigoDestinatario, Me.EnderecoDestinatario)
            Return _Destinatario
        End Get
        Set(ByVal value As Cliente)
            _Destinatario = value
        End Set
    End Property

    Public Property Destinacao() As String
        Get
            Return _Destinacao
        End Get
        Set(ByVal value As String)
            _Destinacao = value
        End Set
    End Property

    Public Property Solicitacao() As Integer
        Get
            Return _Solicitacao
        End Get
        Set(ByVal value As Integer)
            _Solicitacao = value
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

    Public Property UsuarioAlteracao() As String
        Get
            Return _UsuarioAlteracao
        End Get
        Set(ByVal value As String)
            _UsuarioAlteracao = value
        End Set
    End Property

    Public Property DataAlteracao() As DateTime
        Get
            Return _UsuarioAlteracaoData
        End Get
        Set(ByVal value As DateTime)
            _UsuarioAlteracaoData = value
        End Set
    End Property

    Public Property UsuarioCancelamento() As String
        Get
            Return _UsuarioCancelamento
        End Get
        Set(ByVal value As String)
            _UsuarioCancelamento = value
        End Set
    End Property

    Public Property DataCancelamento() As DateTime
        Get
            Return _UsuarioCancelamentoData
        End Get
        Set(ByVal value As DateTime)
            _UsuarioCancelamentoData = value
        End Set
    End Property

    Public Property UsuarioLiberacao() As String
        Get
            Return _UsuarioLiberacao
        End Get
        Set(ByVal value As String)
            _UsuarioLiberacao = value
        End Set
    End Property

    Public Property DataLiberacao() As DateTime
        Get
            Return _UsuarioLiberacaoData
        End Get
        Set(ByVal value As DateTime)
            _UsuarioLiberacaoData = value
        End Set
    End Property

    Public Property UsuarioBaixa() As String
        Get
            Return _UsuarioBaixa
        End Get
        Set(ByVal value As String)
            _UsuarioBaixa = value
        End Set
    End Property

    Public Property DataUsuarioBaixa() As DateTime
        Get
            Return _UsuarioBaixaData
        End Get
        Set(ByVal value As DateTime)
            _UsuarioBaixaData = value
        End Set
    End Property

    Public Property Grupado() As Boolean
        Get
            Return _Grupado
        End Get
        Set(ByVal value As Boolean)
            _Grupado = value
        End Set
    End Property

    Public Property RegistroMestre() As Integer
        Get
            Return _RegistroMestre
        End Get
        Set(ByVal value As Integer)
            _RegistroMestre = value
        End Set
    End Property

    Public Property Observacoes() As String
        Get
            Return _Observacoes
        End Get
        Set(ByVal value As String)
            _Observacoes = value
        End Set
    End Property

    Public Property SituacaoBancaria() As Integer
        Get
            Return _SituacaoBancaria
        End Get
        Set(ByVal value As Integer)
            _SituacaoBancaria = value
        End Set
    End Property

    Public Property TipoFinanceiro() As String
        Get
            Return _TipoFinanceiro
        End Get
        Set(ByVal value As String)
            _TipoFinanceiro = value
        End Set
    End Property

    Public Property CarteiraDoTitulo() As Integer
        Get
            Return _CarteiraDoTitulo
        End Get
        Set(ByVal value As Integer)
            _CarteiraDoTitulo = value
        End Set
    End Property

    Public Property ContratoBancario() As String
        Get
            Return _ContratoBancario
        End Get
        Set(value As String)
            _ContratoBancario = value
        End Set
    End Property

#End Region

#Region "Métodos"

    Public Function Incluir(ByVal Servidor As String, ByVal Codigo As String, Optional ByVal pMomentoFinanceiro As Integer = 0) As String
        Dim strNomeTabela As String = "ContasA" & IIf(Me.EntradaSaida = eEntradaSaida.Saida, "Receber", "Pagar")
        Dim intTitulo As Integer = Me.Codigo

        If Codigo > 0 Then
            intTitulo = Codigo
        Else
            If Me.Codigo = 0 Then intTitulo = Numerador.PegarNumero(Servidor.ToUpper(), eTiposNumerador.Titulo)
        End If

        Dim strSQL As String = "INSERT INTO " & strNomeTabela & " " &
                               "(Registro_Id, Sequencia_Id, Provisao, Carteira, Tributo, Indexador, Moeda, TipoPagto, " &
                               "Situacao, Lote, Movimento, Vencimento, Prorrogacao, DataMoeda, Baixa, UnidadeDeNegocio, " &
                               "Empresa, EndEmpresa, Cliente, EndCliente, BancoCliente, AgenciaCliente, DigitoAgenciaCliente, " &
                               "ContaCliente, DigitoContaCliente, ContaContabilCliente, EmpresaPagadora, EndEmpresaPagadora, " &
                               "BancoPagador, AgenciaPagadora, DigitoAgenciaPagadora, ContaPagadora, DigitoContaPagadora, " &
                               "ContaContabilPagadora, Cheque, Slips, Recibo, Aviso, ReciboDeposito, EmpresaPedido, EndEmpresaPedido, " &
                               "Pedido, PedidoFixacao, Procuracao, ValorDoDocumento, Descontos, Deducoes, Juros, Acrescimos," &
                               "ValorLiquido, MoedaValorDoDocumento, MoedaDescontos, MoedaDeducoes, MoedaJuros, MoedaAcrescimos, " &
                               "MoedaValorLiquido, Historico, CodigoDeBarras, CodigoDigitado, Destinatario, EndDestinatario, " &
                               "NomeDoDestinatario, Destinacao, Solicitacao, UsuarioInclusao, UsuarioInclusaoData, " &
                               "UsuarioAlteracao, UsuarioAlteracaoData, UsuarioCancelamento, UsuarioCancelamentoData, " &
                               "UsuarioLiberacao, UsuarioLiberacaoData, UsuarioBaixa, UsuarioBaixaData, Grupado, " &
                               "RegistroMestre, Observacoes, SituacaoBancaria, CarteiraDoTitulo, ContratoBancario) "

        'Registro_Id, Sequencia_Id
        strSQL &= "VALUES(" & intTitulo & ", " & Me.Sequencia.ToString() & ", "

        'Provisao
        'If pMomentoFinanceiro = 0 Then
        strSQL &= Convert.ToInt32(Me.Provisao).ToString() & ", "
        'ElseIf pMomentoFinanceiro = 3 Then

        '    strSQL &= "3,"
        'Else
        '    strSQL &= "2,"
        'End If

        'Carteira, Tributo, Indexador
        strSQL &= "'" & Me.Carteira & "', '" & Me.Tributo & "', " & Me.Indexador.ToString() & ", "
        'Moeda, TipoPagto, Situacao
        strSQL &= Me.CodigoMoeda.ToString() & ", " & Me.TipoPagamento.ToString() & ", " & Me.Situacao.ToString() & ", "
        'Lote, Movimento, Vencimento
        strSQL &= Me.Lote.ToString() & ", '" & Me.Movimento.ToString("yyyy-MM-dd") & "', '" & Me.Vencimento.ToString("yyyy-MM-dd") & "', "
        'Prorrogacao, DataMeda
        strSQL &= "'" & Me.DataProrrogacao.ToString("yyyy-MM-dd") & "', '" & Me.DataMoeda.ToString("yyyy-MM-dd") & "', "
        'Baixa, UnidadeDeNegocio
        strSQL &= "'" & Me.DataBaixa.ToString("yyyy-MM-dd") & "', '" & Me.UnidadeNegocio & "', "
        'Empresa, EndEmpresa
        strSQL &= "'" & Me.CodigoEmpresa & "', " & Me.EnderecoEmpresa.ToString() & ", "
        'Cliente, EndCliente
        strSQL &= "'" & Me.CodigoCliente & "', " & Me.EnderecoCliente.ToString() & ", "

        If Me.CodigoBancoCliente > 0 Then
            'BancoCliente, AgenciaCliente
            strSQL &= Me.CodigoBancoCliente.ToString() & ", '" & Me.AgenciaCliente & "', "
            'DigitoAgenciaCliente, ContaCliente
            strSQL &= "'" & Me.DigitoAgenciaCliente & "', '" & Me.ContaCliente & "', "
            'DigitoContaCliente
            strSQL &= "'" & Me.DigitoContaCliente & "', "
        Else
            'BancoCliente, AgenciaCliente, DigitoAgenciaCliente, ContaCliente, DigitoContaCliente
            strSQL &= "0, '', '', '', '', "
        End If

        'ContaContabilCliente, EmpresaPagadora, EndEmpresaPagadora
        strSQL &= "'" & Me.ContaContabilCliente & "', '" & Me.CodigoEmpresaPagadora & "', " & Me.EnderecoEmpresaPagadora.ToString() & ", "

        If Me.BancoPagador > 0 Then
            'BancoPagador, AgenciaPagadora
            strSQL &= Me.BancoPagador.ToString() & ", '" & Me.AgenciaPagadora & "', "
            'DigitoAgenciaPagadora, ContaPagadora
            strSQL &= "'" & Me.DigitoAgenciaPagadora & "', '" & Me.ContaPagadora & "', "
            'DigitoContaPagadora
            strSQL &= "'" & Me.DigitoContaPagadora & "', "
        Else
            'BancoPagador, AgenciaPagadora, DigitoAgenciaPagadora, ContaPagadora, DigitoContaPagadora
            strSQL &= "0, '', '', '', '', "
        End If

        'ContaContabilPagadora, Cheque, Slips
        strSQL &= "'" & Me.ContaContabilPagadora & "', '" & IIf(Me.Cheque, "S", "N") & "', '" & IIf(Me.Slips, "S", "N") & "', "
        'Recibo, Aviso, ReciboDeposito
        strSQL &= "'" & IIf(Me.Recibo, "S", "N") & "', '" & IIf(Me.Aviso, "S", "N") & "', '" & IIf(Me.ReciboDeposito, "S", "N") & "', "

        If Me.CodigoEmpresaPedido = "" Then
            'EmpresaPedido, EndEmpresaPedido
            strSQL &= "NULL, NULL, "
        Else
            'EmpresaPedido, EndEmpresaPedido
            strSQL &= "'" & Me.CodigoEmpresaPedido & "', " & Me.EnderecoEmpresaPedido & ", "
        End If

        'Pedido, PedidoFixacao
        If Me.CodigoPedido = 0 Then
            strSQL &= "NULL, "
        Else
            strSQL &= Me.CodigoPedido.ToString() & ", "
        End If

        'PedidoFixacao
        If Me.PedidoFixacao = 0 Then
            strSQL &= "NULL, "
        Else
            strSQL &= Me.PedidoFixacao.ToString() & ", "
        End If

        'Procuracao
        If Me.Procuracao = 0 Then
            strSQL &= "NULL, "
        Else
            strSQL &= Me.Procuracao.ToString() & ", "
        End If

        'ValorDoDocumento
        strSQL &= Me.ValorDocumentoOficial.ToString().Replace(".", "").Replace(",", ".") & ", "
        'Descontos
        strSQL &= Me.DescontosOficial.ToString().Replace(".", "").Replace(",", ".") & ", "
        'Deducoes
        strSQL &= Me.DeducoesOficial.ToString().Replace(".", "").Replace(",", ".") & ", "
        'Juros
        strSQL &= Me.JurosOficial.ToString().Replace(".", "").Replace(",", ".") & ", "
        'Acrescimos
        strSQL &= Me.AcrescimosOficial.ToString().Replace(".", "").Replace(",", ".") & ", "
        'ValorLiquido
        strSQL &= Me.ValorLiquidoOficial.ToString().Replace(".", "").Replace(",", ".") & ", "
        'MoedaValorDoDocumento
        strSQL &= Me.ValorDocumentoMoeda.ToString().Replace(".", "").Replace(",", ".") & ", "
        'MoedaDescontos
        strSQL &= Me.DescontosMoeda.ToString().Replace(".", "").Replace(",", ".") & ", "
        'MoedaDeducoes
        strSQL &= Me.DeducoesMoeda.ToString().Replace(".", "").Replace(",", ".") & ", "
        'MoedaJuros
        strSQL &= Me.JurosMoeda.ToString().Replace(".", "").Replace(",", ".") & ", "
        'MoedaAcrescimos
        strSQL &= Me.AcrescimosMoeda.ToString().Replace(".", "").Replace(",", ".") & ", "
        'MoedaValorLiquido
        strSQL &= Me.ValorLiquidoMoeda.ToString().Replace(".", "").Replace(",", ".") & ", "
        'Historico, CodigoDeBarras, CodigoDigitado
        strSQL &= "'" & Historico & "', '', 'N', "
        'Destinatario, EndDestinatario
        strSQL &= "'" & Me.CodigoDestinatario & "', " & Me.EnderecoDestinatario.ToString() & ", "
        'NomeDoDestinatario
        If Me.Destinatario Is Nothing Then
            strSQL &= "NULL, "
        ElseIf Me.Destinatario.Nome.Length > 50 Then
            strSQL &= "'" & Me.Destinatario.Nome.Substring(0, 50) & "', "
        Else
            strSQL &= "'" & Me.Destinatario.Nome & "', "
        End If
        'Destinacao
        strSQL &= "'" & Me.Destinacao & "', "
        'Solicitacao, UsuarioInclusao
        strSQL &= Me.Solicitacao.ToString() & ", '" & Me.UsuarioInclusao.ToString() & "', "
        'UsuarioInclusaoData, UsuarioAlteracao, UsuarioAlteracaoData
        strSQL &= "'" & Me.DataInclusao.ToString("yyyy-MM-dd") & "', NULL, NULL, "
        'UsuarioCancelamento, UsuarioCancelamentoData
        strSQL &= "NULL, NULL, "
        'UsuarioLiberacao, UsuarioLiberacaoData
        strSQL &= "'', NULL, "
        'UsuarioBaixa, UsuarioBaixaData
        strSQL &= "NULL, NULL, "
        'Grupado, RegistroMestre
        strSQL &= "'" & IIf(Me.Grupado, "S", "N") & "', " & Me.RegistroMestre.ToString() & ", "
        'Observacoes, SituacaoBancaria
        Dim OBS As String = ""
        If Me.Observacoes.Length = 0 Then
            Dim Cliente As New Cliente(Me.CodigoCliente, Me.EnderecoCliente)
            OBS = Cliente.Nome
        Else
            OBS = Me.Observacoes
        End If
        strSQL &= "'" & OBS & "', " & Me.SituacaoBancaria.ToString() & ", " & Me.CarteiraDoTitulo.ToString() & ","
        'Contratio Bancario
        strSQL &= "'" & Me.ContratoBancario & "')"

        Return strSQL
    End Function

    Public Function Alterar() As String
        Dim strNomeTabela As String = IIf(Me.EntradaSaida = eEntradaSaida.Saida, "ContasAReceber", "ContasAPagar")
        Dim strSQL As String = "UPDATE " & strNomeTabela & " SET Prorrogacao = '" & Me.DataProrrogacao.ToString("yyyy-MM-dd") & "', UsuarioAlteracao = '" & HttpContext.Current.Session("ssNomeUsuario") & "', UsuarioAlteracaoData = '" & Now().ToString("yyyy-MM-dd") & "'" &
                               "WHERE Registro_Id = " & Me.Codigo.ToString() & " " &
                               "AND Sequencia_Id = " & Me.Sequencia.ToString() & " " &
                               "AND Provisao = " & Me.Provisao

        Return strSQL
    End Function

    Public Function Excluir() As String
        Dim strNomeTabela As String = IIf(Me.EntradaSaida = eEntradaSaida.Saida, "ContasAReceber", "ContasAPagar")
        Dim strSQL As String = "DELETE " & strNomeTabela & " " &
                               "WHERE Registro_Id = " & Me.Codigo.ToString() & " " &
                               "AND Sequencia_Id = " & Me.Sequencia.ToString() & " " &
                               "AND Provisao = " & Me.Provisao

        Return strSQL
    End Function

    Public Function Excluido() As String
        Dim strNomeTabela As String = IIf(Me.EntradaSaida = eEntradaSaida.Saida, "ContasAReceber", "ContasAPagar")
        Dim strSQL As String = "UPDATE " & strNomeTabela & " SET Situacao = 3 " &
                               "WHERE Registro_Id = " & Me.Codigo.ToString() & " " &
                               "AND Sequencia_Id = " & Me.Sequencia.ToString() & " " &
                               "AND Provisao = " & Me.Provisao

        Return strSQL
    End Function

#End Region

End Class

<Serializable()> _
Public Class MovimentacoesFinanceiras
    Inherits List(Of MovimentacaoFinanceira)

    Public EntradaSaida As eEntradaSaida
    Public OficialAMovimentar As Decimal = 0
    Public OficialMovimentado As Decimal = 0
    Public MoedaAMovimentar As Decimal = 0
    Public MoedaMovimentado As Decimal = 0

    Public OficialADevolver As Decimal = 0
    Public OficialDevolvido As Decimal = 0
    Public MoedaADevolver As Decimal = 0
    Public MoedaDevolvido As Decimal = 0

#Region "Construtores"

    Public Sub New()

    End Sub

    Public Sub New(ByVal EntradaSaida As eEntradaSaida)
        Me.EntradaSaida = EntradaSaida
    End Sub

#End Region

#Region "Métodos"

    Protected Function AtribuirDados(ByVal Provisao As eProvisao, ByVal Codigo As Integer, ByVal Carteira As String, ByVal Indexador As Integer, _
                                     ByVal Moeda As Integer, ByVal Movimento As DateTime, ByVal Vencimento As DateTime, _
                                     ByVal Unidade As Cliente, ByVal Empresa As Cliente, ByVal Cliente As Cliente, _
                                     ByVal SubOperacao As SubOperacao, ByVal ValorOficial As Decimal, ByVal ValorMoeda As Decimal, _
                                     ByVal Historico As String, ByVal Pedido As Pedido) As MovimentacaoFinanceira

        Return AtribuirDados(Nothing, Provisao, Codigo, Carteira, Indexador, Moeda, Movimento, Vencimento, _
                             Unidade, Empresa, Cliente, SubOperacao, ValorOficial, ValorMoeda, _
                             Historico, Pedido, Nothing, Nothing)
    End Function

    Protected Function AtribuirDados(ByVal Provisao As eProvisao, ByVal Codigo As Integer, ByVal Carteira As String, ByVal Indexador As Integer, _
                                     ByVal Moeda As Integer, ByVal Movimento As DateTime, ByVal Vencimento As DateTime, _
                                     ByVal Unidade As Cliente, ByVal Empresa As Cliente, ByVal Cliente As Cliente, _
                                     ByVal SubOperacao As SubOperacao, ByVal ValorOficial As Decimal, ByVal ValorMoeda As Decimal, _
                                     ByVal Historico As String, ByVal ContaBancaria As ClienteXContaBancaria, ByVal CarteiraDoTitulo As Integer) As MovimentacaoFinanceira

        Return AtribuirDados(Nothing, Provisao, Codigo, Carteira, Indexador, Moeda, Movimento, Vencimento, _
                             Unidade, Empresa, Cliente, SubOperacao, ValorOficial, ValorMoeda, _
                             Historico, Nothing, ContaBancaria, CarteiraDoTitulo)
    End Function

    Protected Function AtribuirDados(ByVal Provisao As eProvisao, ByVal Codigo As Integer, ByVal Carteira As String, ByVal Indexador As Integer, _
                                     ByVal Moeda As Integer, ByVal Movimento As DateTime, ByVal Vencimento As DateTime, _
                                     ByVal Unidade As Cliente, ByVal Empresa As Cliente, ByVal Cliente As Cliente, _
                                     ByVal SubOperacao As SubOperacao, ByVal ValorOficial As Decimal, ByVal ValorMoeda As Decimal, _
                                     ByVal Historico As String, ByVal Pedido As Pedido, _
                                     ByVal ContaBancaria As ClienteXContaBancaria, ByVal CarteiraDoTitulo As Integer) As MovimentacaoFinanceira

        Return AtribuirDados(Nothing, Provisao, Codigo, Carteira, Indexador, Moeda, Movimento, Vencimento, _
                             Unidade, Empresa, Cliente, SubOperacao, ValorOficial, ValorMoeda, _
                             Historico, Pedido, ContaBancaria, CarteiraDoTitulo)
    End Function

    Protected Function AtribuirDados(ByVal ObjetoMovimentacao As MovimentacaoFinanceira, ByVal Provisao As eProvisao, _
                                     ByVal Codigo As Integer, ByVal Carteira As String, ByVal Indexador As Integer, _
                                     ByVal Moeda As Integer, ByVal Movimento As DateTime, ByVal Vencimento As DateTime, _
                                     ByVal Unidade As Cliente, ByVal Empresa As Cliente, ByVal Cliente As Cliente, _
                                     ByVal SubOperacao As SubOperacao, ByVal ValorOficial As Decimal, ByVal ValorMoeda As Decimal, _
                                     ByVal Historico As String, ByVal Pedido As Pedido) As MovimentacaoFinanceira

        Return AtribuirDados(ObjetoMovimentacao, Provisao, Codigo, Carteira, Indexador, Moeda, Movimento, Vencimento, _
                             Unidade, Empresa, Cliente, SubOperacao, ValorOficial, ValorMoeda, _
                             Historico, Pedido, Nothing, Nothing)
    End Function

    Protected Function AtribuirDados(ByVal ObjetoMovimentacao As MovimentacaoFinanceira, ByVal Provisao As eProvisao, _
                                     ByVal Codigo As Integer, ByVal Carteira As String, ByVal Indexador As Integer, _
                                     ByVal Moeda As Integer, ByVal Movimento As DateTime, ByVal Vencimento As DateTime, _
                                     ByVal Unidade As Cliente, ByVal Empresa As Cliente, ByVal Cliente As Cliente, _
                                     ByVal SubOperacao As SubOperacao, ByVal ValorOficial As Decimal, ByVal ValorMoeda As Decimal, _
                                     ByVal Historico As String, ByVal ContaBancaria As ClienteXContaBancaria, ByVal CarteiraDoTitulo As Integer) As MovimentacaoFinanceira

        Return AtribuirDados(ObjetoMovimentacao, Provisao, Codigo, Carteira, Indexador, Moeda, Movimento, Vencimento, _
                             Unidade, Empresa, Cliente, SubOperacao, ValorOficial, ValorMoeda, _
                             Historico, Nothing, ContaBancaria, CarteiraDoTitulo)
    End Function

    Protected Function AtribuirDados(ByVal ObjetoMovimentacao As MovimentacaoFinanceira, ByVal Provisao As eProvisao, _
                                     ByVal Codigo As Integer, ByVal Carteira As String, ByVal Indexador As Integer, _
                                     ByVal Moeda As Integer, ByVal Movimento As DateTime, ByVal Vencimento As DateTime, _
                                     ByVal Unidade As Cliente, ByVal Empresa As Cliente, ByVal Cliente As Cliente, _
                                     ByVal SubOperacao As SubOperacao, ByVal ValorOficial As Decimal, ByVal ValorMoeda As Decimal, _
                                     ByVal Historico As String, ByVal Pedido As Pedido, _
                                     ByVal ContaBancaria As ClienteXContaBancaria, ByVal CarteiraDoTitulo As Integer) As MovimentacaoFinanceira

        Dim objParcelamento As New MovimentacaoFinanceira()

        If ObjetoMovimentacao IsNot Nothing Then objParcelamento = ObjetoMovimentacao

        With objParcelamento
            If Codigo > 0 Then .Codigo = Codigo

            .EntradaSaida = Me.EntradaSaida
            .Provisao = Provisao
            .Carteira = Carteira
            .Tributo = ""
            .Indexador = Indexador
            .CodigoMoeda = Moeda
            .TipoPagamento = 0
            .Situacao = 1
            .Lote = 70
            .Movimento = Movimento
            .Vencimento = Vencimento
            .DataProrrogacao = Vencimento
            .DataMoeda = Vencimento
            .DataBaixa = Vencimento
            .UnidadeNegocio = Unidade.Codigo
            .CodigoEmpresa = Empresa.Codigo
            .EnderecoEmpresa = Empresa.CodigoEndereco
            .CodigoCliente = Cliente.Codigo
            .EnderecoCliente = Cliente.CodigoEndereco

            If Not ContaBancaria Is Nothing Then
                .CodigoBancoCliente = ContaBancaria.CodigoBanco
                .AgenciaCliente = ContaBancaria.CodigoAgencia.ToString()
                .DigitoAgenciaCliente = ContaBancaria.DigitoAgencia
                .ContaCliente = ContaBancaria.ContaCorrente
                .DigitoContaCliente = ContaBancaria.DigitoConta
            Else
                .CodigoBancoCliente = 0
                .AgenciaCliente = ""
                .DigitoAgenciaCliente = ""
                .ContaCliente = ""
                .DigitoContaCliente = ""
            End If

            .ContaContabilCliente = SubOperacao.CodigoGrupoContas
            .CodigoEmpresaPagadora = ""
            .EnderecoEmpresaPagadora = 0
            .BancoPagador = 0
            .AgenciaPagadora = ""
            .DigitoAgenciaPagadora = ""
            .ContaPagadora = ""
            .DigitoContaPagadora = ""
            .ContaContabilPagadora = ""
            .Cheque = False
            .Slips = False
            .Recibo = False
            .Aviso = False
            .ReciboDeposito = False
            .CodigoEmpresaPedido = Pedido.CodigoEmpresa
            .EnderecoEmpresaPedido = Pedido.EnderecoEmpresa
            .CodigoPedido = Pedido.Codigo
            'Fixacao só deve Gravar no Afixar - Furlan - 09/10/2014
            '.PedidoFixacao = IIf(SubOperacao.PrecoFixo, Pedido.Codigo, 0)
            .PedidoFixacao = 0
            .Procuracao = 0

            .DescontosOficial = 0
            .DeducoesOficial = 0
            .JurosOficial = 0
            .AcrescimosOficial = 0
            .DescontosMoeda = 0
            .DeducoesMoeda = 0
            .JurosMoeda = 0
            .AcrescimosMoeda = 0
            .ValorDocumentoOficial = ValorOficial
            .ValorLiquidoOficial = ValorOficial
            .ValorDocumentoMoeda = ValorMoeda
            .ValorLiquidoMoeda = ValorMoeda
            .Historico = Historico
            .CodigoBarras = ""
            .CodigoDigitado = False
            .CodigoDestinatario = Cliente.Codigo
            .EnderecoDestinatario = Cliente.CodigoEndereco
            .Destinacao = ""
            .Solicitacao = 0
            .UsuarioInclusao = Pedido.UsuarioInclusao
            .DataInclusao = DateTime.Now
            .UsuarioAlteracao = ""
            .DataAlteracao = Nothing
            .UsuarioCancelamento = ""
            .DataCancelamento = Nothing
            .UsuarioLiberacao = ""
            .DataLiberacao = Nothing
            .UsuarioBaixa = ""
            .DataUsuarioBaixa = Nothing
            .Grupado = False
            .RegistroMestre = 0
            .Observacoes = ""
            .SituacaoBancaria = 0
            .TipoFinanceiro = ""
            .CarteiraDoTitulo = CarteiraDoTitulo
            .ContratoBancario = ""
            .Situacao = IIf(ValorOficial = 0, eSituacao.Cancelado, eSituacao.Normal)
        End With

        Return objParcelamento
    End Function

    Public Function GetSQL(ByVal Modo As eModoAlteracao, Optional ByVal Servidor As String = "") As ArrayList
        Dim arrSQL As New ArrayList()

        For Each objItem As MovimentacaoFinanceira In Me
            Select Case Modo
                Case eModoAlteracao.Inclusao : arrSQL.Add(objItem.Incluir(Servidor, objItem.Codigo))
                Case eModoAlteracao.Exclusao : If objItem.Provisao = eProvisao.Previsao Then arrSQL.Add(objItem.Excluir())
            End Select
        Next

        Return arrSQL
    End Function

    Public Sub ModificarHistorico(ByVal Tipo As eTabelas, ByVal Parametros As String())
        For Each objConta As MovimentacaoFinanceira In Me
            Select Case Tipo
                Case eTabelas.Pedido
                    objConta.Historico = objConta.Historico.Replace("{PEDIDO}", Parametros(0))
                    objConta.CodigoPedido = Convert.ToInt32(Parametros(0))
            End Select
        Next
    End Sub

    Public Sub TotalDasParcelas()
        Me.OficialAMovimentar = 0
        Me.OficialMovimentado = 0
        Me.MoedaAMovimentar = 0
        Me.MoedaMovimentado = 0

        Me.OficialADevolver = 0
        Me.OficialDevolvido = 0
        Me.MoedaADevolver = 0
        Me.MoedaDevolvido = 0

        For Each objItem As MovimentacaoFinanceira In Me.Where(Function(s) Not s.Situacao = eSituacao.Excluido)
            Dim cart As New CarteiraFinanceira(objItem.Carteira)
            Dim cop As New PlanoDeConta("", 0, objItem.ContaContabilPagadora)

            If Not (cart.BaixaAdiantamento Or cop.Adiantamento) And _
                ((Me.EntradaSaida = eEntradaSaida.Saida And objItem.TipoFinanceiro = "P") OrElse _
               (Me.EntradaSaida = eEntradaSaida.Entrada And objItem.TipoFinanceiro = "R")) Then
                Me.OficialADevolver += objItem.ValorDocumentoOficial
                Me.OficialDevolvido += IIf(objItem.Provisao = eProvisao.Baixa, objItem.ValorDocumentoOficial, 0)
                Me.MoedaADevolver += objItem.ValorDocumentoMoeda
                Me.MoedaDevolvido += IIf(objItem.Provisao = eProvisao.Baixa, objItem.ValorDocumentoMoeda, 0)
            Else
                Me.OficialAMovimentar += objItem.ValorDocumentoOficial
                Me.OficialMovimentado += IIf(objItem.Provisao = eProvisao.Baixa, objItem.ValorDocumentoOficial, 0)
                Me.MoedaAMovimentar += objItem.ValorDocumentoMoeda
                Me.MoedaMovimentado += IIf(objItem.Provisao = eProvisao.Baixa, objItem.ValorDocumentoMoeda, 0)
            End If
        Next
    End Sub

    Public Function Selecionar(Optional ByVal Empresa As String = "", Optional ByVal EndEmpresa As Integer = -1, Optional ByVal Pedido As Integer = 0, Optional ByVal Situacao As Integer = 0, Optional ByVal objPedido As Pedido = Nothing) As Boolean
        Dim objBanco As New AcessaBanco()

        Me.OficialAMovimentar = 0
        Me.OficialMovimentado = 0
        Me.MoedaAMovimentar = 0
        Me.MoedaMovimentado = 0

        Me.OficialADevolver = 0
        Me.OficialDevolvido = 0
        Me.MoedaADevolver = 0
        Me.MoedaDevolvido = 0

        Try
            Dim strSQL As String
            strSQL = "SELECT T.Registro_Id, T.Sequencia_Id, T.Provisao, T.Carteira, T.Tributo, T.Indexador, T.Moeda, T.TipoPagto, " & vbCrLf & _
                     "       T.Situacao, T.Lote, T.Movimento, T.Vencimento, T.Prorrogacao, T.DataMoeda, T.Baixa, T.UnidadeDeNegocio, T.Empresa, " & vbCrLf & _
                     "       T.EndEmpresa, T.Cliente, T.EndCliente, T.BancoCliente, T.AgenciaCliente, T.DigitoAgenciaCliente, T.ContaCliente, " & vbCrLf & _
                     "       T.DigitoContaCliente, T.ContaContabilCliente, T.EmpresaPagadora, T.EndEmpresaPagadora, T.BancoPagador, " & vbCrLf & _
                     "       T.AgenciaPagadora, T.DigitoAgenciaPagadora, T.ContaPagadora, T.DigitoContaPagadora, T.ContaContabilPagadora, " & vbCrLf & _
                     "       T.Cheque, T.Slips, T.Recibo, T.Aviso, T.ReciboDeposito, T.EmpresaPedido, T.EndEmpresaPedido, T.Pedido, " & vbCrLf & _
                     "       T.PedidoFixacao, T.Procuracao, T.ValorDoDocumento, T.Descontos, T.Deducoes, T.Juros, T.Acrescimos, " & vbCrLf & _
                     "       T.ValorLiquido, T.MoedaValorDoDocumento, T.MoedaDescontos, T.MoedaDeducoes, T.MoedaJuros, T.MoedaAcrescimos, " & vbCrLf & _
                     "       T.MoedaValorLiquido, T.Historico, T.CodigoDeBarras, T.CodigoDigitado, T.Destinatario, T.EndDestinatario, " & vbCrLf & _
                     "       T.NomeDoDestinatario, T.Destinacao, T.Solicitacao, T.UsuarioInclusao, T.UsuarioInclusaoData, " & vbCrLf & _
                     "       T.UsuarioAlteracao, T.UsuarioAlteracaoData, T.UsuarioCancelamento, T.UsuarioCancelamentoData, " & vbCrLf & _
                     "       isnull(T.UsuarioLiberacao,'') AS UsuarioLiberacao, T.UsuarioLiberacaoData, isnull(T.UsuarioBaixa,'') AS UsuarioBaixa, T.UsuarioBaixaData, T.Grupado, T.RegistroMestre, " & vbCrLf & _
                     "       T.Observacoes, T.SituacaoBancaria, 'P' AS TipoFinanceiro, isnull(T.CarteiraDoTitulo,0) AS CarteiraDoTitulo, isnull(T.ContratoBancario,'') AS ContratoBancario " & vbCrLf & _
                     "  FROM ContasAPagar T " & vbCrLf & _
                     " Left Join ComprasxProdutos cart" & vbCrLf & _
                     "   on T.carteira = cart.Produto_id" & vbCrLf & _
                     " Where isnull(Grupado,'N') <> 'M' " & vbCrLf & _
                     "   and (cart.adiantamento = 'N' or (cart.adiantamento = 'S' and cart.baixaadiantamento = 0 and provisao <> 1))"

            If Empresa <> "" Then strSQL &= " and T.Empresa = '" & Empresa & "'" & vbCrLf
            If EndEmpresa > -1 Then strSQL &= " and T.EndEmpresa = " & EndEmpresa.ToString() & vbCrLf
            If Pedido > 0 Then strSQL &= " and T.Pedido = " & Pedido.ToString() & vbCrLf

            If Situacao > 0 Then
                If Situacao = 1 Then
                    strSQL &= " and T.Situacao IN(1,101,102) " & vbCrLf
                Else
                    strSQL &= " and T.Situacao = " & Situacao.ToString() & vbCrLf
                End If
            End If

            strSQL &= " Union All " & vbCrLf & _
                      "SELECT T.Registro_Id, T.Sequencia_Id, T.Provisao, T.Carteira, T.Tributo, T.Indexador, T.Moeda, T.TipoPagto, " & vbCrLf & _
                      "       T.Situacao, T.Lote, T.Movimento, T.Vencimento, T.Prorrogacao, T.DataMoeda, T.Baixa, T.UnidadeDeNegocio, T.Empresa, " & vbCrLf & _
                      "       T.EndEmpresa, T.Cliente, T.EndCliente, T.BancoCliente, T.AgenciaCliente, T.DigitoAgenciaCliente, T.ContaCliente, " & vbCrLf & _
                      "       T.DigitoContaCliente, T.ContaContabilCliente, T.EmpresaPagadora, T.EndEmpresaPagadora, T.BancoPagador, " & vbCrLf & _
                      "       T.AgenciaPagadora, T.DigitoAgenciaPagadora, T.ContaPagadora, T.DigitoContaPagadora, T.ContaContabilPagadora, " & vbCrLf & _
                      "       T.Cheque, T.Slips, T.Recibo, T.Aviso, T.ReciboDeposito, T.EmpresaPedido, T.EndEmpresaPedido, T.Pedido, " & vbCrLf & _
                      "       T.PedidoFixacao, T.Procuracao, T.ValorDoDocumento, T.Descontos, T.Deducoes, T.Juros, T.Acrescimos, " & vbCrLf & _
                      "       T.ValorLiquido, T.MoedaValorDoDocumento, T.MoedaDescontos, T.MoedaDeducoes, T.MoedaJuros, T.MoedaAcrescimos, " & vbCrLf & _
                      "       T.MoedaValorLiquido, T.Historico, T.CodigoDeBarras, T.CodigoDigitado, T.Destinatario, T.EndDestinatario, " & vbCrLf & _
                      "       T.NomeDoDestinatario, T.Destinacao, T.Solicitacao, T.UsuarioInclusao, T.UsuarioInclusaoData, " & vbCrLf & _
                      "       T.UsuarioAlteracao, T.UsuarioAlteracaoData, T.UsuarioCancelamento, T.UsuarioCancelamentoData, " & vbCrLf & _
                      "       isnull(T.UsuarioLiberacao,'') AS UsuarioLiberacao, T.UsuarioLiberacaoData, isnull(T.UsuarioBaixa,'') AS UsuarioBaixa, T.UsuarioBaixaData, T.Grupado, T.RegistroMestre, " & vbCrLf & _
                      "       T.Observacoes, T.SituacaoBancaria, 'R' AS TipoFinanceiro, isnull(T.CarteiraDoTitulo,0) AS CarteiraDoTitulo, isnull(T.ContratoBancario,'') AS ContratoBancario " & vbCrLf & _
                      "  FROM ContasAReceber T " & vbCrLf & _
                      " Left Join ComprasxProdutos cart" & vbCrLf & _
                      "   on T.carteira = cart.Produto_id" & vbCrLf & _
                      " Where isnull(Grupado,'N') <> 'M' " & vbCrLf & _
                      "   and (cart.adiantamento = 'N' or (cart.adiantamento = 'S' and cart.baixaadiantamento = 0 and provisao <> 1))"

            If Empresa <> "" Then strSQL &= " and T.Empresa = '" & Empresa & "'" & vbCrLf
            If EndEmpresa > -1 Then strSQL &= " and T.EndEmpresa = " & EndEmpresa.ToString() & vbCrLf
            If Pedido > 0 Then strSQL &= " and T.Pedido = " & Pedido.ToString() & vbCrLf

            If Situacao > 0 Then
                If Situacao = 1 Then
                    strSQL &= " and T.Situacao IN(1,101,102) " & vbCrLf
                Else
                    strSQL &= " and T.Situacao = " & Situacao.ToString() & vbCrLf
                End If
            End If

            strSQL &= " order By T.Provisao, T.Registro_Id "

            Dim dsContas As DataSet = objBanco.ConsultaDataSet(strSQL, "Titulos")

            For Each drConta As DataRow In dsContas.Tables(0).Rows
                Dim objConta As New MovimentacaoFinanceira()

                objConta.EntradaSaida = Me.EntradaSaida
                objConta.Codigo = Convert.ToInt32(drConta("Registro_Id"))
                objConta.Sequencia = Convert.ToInt32(drConta("Sequencia_Id"))
                objConta.Provisao = CType(Convert.ToInt32(drConta("Provisao")), eProvisao)
                objConta.Carteira = drConta("Carteira").ToString()
                objConta.Tributo = drConta("Tributo").ToString()
                objConta.Indexador = Convert.ToInt32(drConta("Indexador"))
                objConta.CodigoMoeda = Convert.ToInt32(drConta("Moeda"))
                objConta.TipoPagamento = Convert.ToInt32(drConta("TipoPagto"))
                objConta.Situacao = Convert.ToInt32(drConta("Situacao"))
                objConta.Lote = Convert.ToInt32(drConta("Lote"))
                objConta.Movimento = Convert.ToDateTime(drConta("Movimento"))
                objConta.Vencimento = Convert.ToDateTime(drConta("Vencimento"))
                objConta.DataProrrogacao = Convert.ToDateTime(drConta("Prorrogacao")) 'Mudei para vencimento, estava com erro
                objConta.DataMoeda = Convert.ToDateTime(drConta("DataMoeda"))
                objConta.DataBaixa = Convert.ToDateTime(drConta("Baixa"))
                objConta.UnidadeNegocio = drConta("UnidadeDeNegocio").ToString()
                objConta.CodigoEmpresa = drConta("Empresa").ToString()
                objConta.EnderecoEmpresa = Convert.ToInt32(drConta("EndEmpresa"))
                objConta.CodigoCliente = drConta("Cliente").ToString()
                objConta.EnderecoCliente = Convert.ToInt32(drConta("EndCliente"))
                objConta.CodigoBancoCliente = Convert.ToInt32(drConta("BancoCliente"))
                objConta.AgenciaCliente = drConta("AgenciaCliente").ToString()
                objConta.DigitoAgenciaCliente = drConta("DigitoAgenciaCliente").ToString()
                objConta.ContaCliente = drConta("ContaCliente").ToString()
                objConta.DigitoContaCliente = drConta("DigitoContaCliente").ToString()
                objConta.ContaContabilCliente = drConta("ContaContabilCliente").ToString()
                objConta.CodigoEmpresaPagadora = drConta("EmpresaPagadora").ToString()
                objConta.EnderecoEmpresaPagadora = Convert.ToInt32(drConta("EndEmpresaPagadora"))
                objConta.BancoPagador = Convert.ToInt32(drConta("BancoPagador"))
                objConta.AgenciaPagadora = drConta("AgenciaPagadora").ToString()
                objConta.DigitoAgenciaPagadora = drConta("DigitoAgenciaPagadora").ToString()
                objConta.ContaPagadora = drConta("ContaPagadora").ToString()
                objConta.DigitoContaPagadora = drConta("DigitoContaPagadora").ToString()
                objConta.ContaContabilPagadora = drConta("ContaContabilPagadora").ToString()
                objConta.Cheque = drConta("Cheque").ToString() = "S"
                objConta.Slips = drConta("Slips").ToString() = "S"
                objConta.Recibo = drConta("Recibo").ToString() = "S"
                objConta.Aviso = drConta("Aviso").ToString() = "S"
                objConta.ReciboDeposito = drConta("ReciboDeposito").ToString() = "S"
                If Not drConta.IsNull("EmpresaPedido") Then objConta.CodigoEmpresaPedido = drConta("EmpresaPedido").ToString() Else objConta.CodigoEmpresaPedido = ""
                If Not drConta.IsNull("EndEmpresaPedido") Then objConta.EnderecoEmpresaPedido = Convert.ToInt32(drConta("EndEmpresaPedido")) Else objConta.EnderecoEmpresaPedido = 0
                objConta.CodigoPedido = Convert.ToInt32(drConta("Pedido"))
                If Not drConta.IsNull("PedidoFixacao") Then objConta.PedidoFixacao = Convert.ToInt32(drConta("PedidoFixacao")) Else objConta.PedidoFixacao = 0
                If Not drConta.IsNull("Procuracao") Then objConta.Procuracao = Convert.ToInt32(drConta("Procuracao")) Else objConta.Procuracao = Nothing
                objConta.ValorDocumentoOficial = Convert.ToDecimal(drConta("ValorDoDocumento"))
                objConta.DescontosOficial = Convert.ToDecimal(drConta("Descontos"))
                objConta.DeducoesOficial = Convert.ToDecimal(drConta("Deducoes"))
                objConta.JurosOficial = Convert.ToDecimal(drConta("Juros"))
                objConta.AcrescimosOficial = Convert.ToDecimal(drConta("Acrescimos"))
                objConta.ValorLiquidoOficial = Convert.ToDecimal(drConta("ValorLiquido"))
                If Not drConta.IsNull("MoedaValorDoDocumento") Then objConta.ValorDocumentoMoeda = Convert.ToDecimal(drConta("MoedaValorDoDocumento")) Else objConta.ValorDocumentoMoeda = 0
                If Not drConta.IsNull("MoedaDescontos") Then objConta.DescontosMoeda = Convert.ToDecimal(drConta("MoedaDescontos")) Else objConta.DescontosMoeda = 0
                If Not drConta.IsNull("MoedaDeducoes") Then objConta.DeducoesMoeda = Convert.ToDecimal(drConta("MoedaDeducoes")) Else objConta.DeducoesMoeda = 0
                If Not drConta.IsNull("MoedaJuros") Then objConta.JurosMoeda = Convert.ToDecimal(drConta("MoedaJuros")) Else objConta.JurosMoeda = 0
                If Not drConta.IsNull("MoedaAcrescimos") Then objConta.AcrescimosMoeda = Convert.ToDecimal(drConta("MoedaAcrescimos")) Else objConta.AcrescimosMoeda = 0
                If Not drConta.IsNull("MoedaValorLiquido") Then objConta.ValorLiquidoMoeda = Convert.ToDecimal(drConta("MoedaValorLiquido")) Else objConta.ValorLiquidoMoeda = 0
                objConta.Historico = drConta("Historico").ToString()
                objConta.CodigoBarras = drConta("CodigoDeBarras").ToString()
                objConta.CodigoDigitado = drConta("CodigoDigitado").ToString() = "S"
                objConta.CodigoDestinatario = drConta("Destinatario").ToString()
                objConta.EnderecoDestinatario = Convert.ToInt32(drConta("EndDestinatario"))
                objConta.Destinacao = drConta("Destinacao").ToString()
                objConta.Solicitacao = Convert.ToInt32(drConta("Solicitacao"))
                objConta.UsuarioInclusao = drConta("UsuarioInclusao").ToString()
                If Not drConta.IsNull("UsuarioInclusaoData") Then objConta.DataInclusao = Convert.ToDateTime(drConta("UsuarioInclusaoData")) Else objConta.DataInclusao = Nothing
                If Not drConta.IsNull("UsuarioAlteracao") Then objConta.UsuarioAlteracao = drConta("UsuarioAlteracao").ToString() Else objConta.UsuarioAlteracao = Nothing
                If Not drConta.IsNull("UsuarioAlteracaoData") Then objConta.DataAlteracao = Convert.ToDateTime(drConta("UsuarioAlteracaoData")) Else objConta.DataAlteracao = Nothing
                If Not drConta.IsNull("UsuarioCancelamento") Then objConta.UsuarioCancelamento = drConta("UsuarioCancelamento").ToString() Else objConta.UsuarioCancelamento = Nothing
                If Not drConta.IsNull("UsuarioCancelamentoData") Then objConta.DataCancelamento = Convert.ToDateTime(drConta("UsuarioCancelamentoData")) Else objConta.DataCancelamento = Nothing
                If Not drConta.IsNull("UsuarioLiberacao") Then objConta.UsuarioLiberacao = drConta("UsuarioLiberacao").ToString() Else objConta.UsuarioLiberacao = Nothing
                If Not drConta.IsNull("UsuarioLiberacaoData") Then objConta.DataLiberacao = Convert.ToDateTime(drConta("UsuarioLiberacaoData")) Else objConta.DataLiberacao = Nothing
                If Not drConta.IsNull("UsuarioBaixa") Then objConta.UsuarioBaixa = drConta("UsuarioBaixa").ToString() Else objConta.UsuarioBaixa = Nothing
                If Not drConta.IsNull("UsuarioBaixaData") Then objConta.DataUsuarioBaixa = Convert.ToDateTime(drConta("UsuarioBaixaData")) Else objConta.DataUsuarioBaixa = Nothing
                objConta.Grupado = drConta("Grupado").ToString() = "S"
                If Not drConta.IsNull("RegistroMestre") Then objConta.RegistroMestre = Convert.ToInt32(drConta("RegistroMestre")) Else objConta.RegistroMestre = 0
                objConta.Observacoes = drConta("Observacoes").ToString()
                If Not drConta.IsNull("SituacaoBancaria") Then objConta.SituacaoBancaria = Convert.ToInt32(drConta("SituacaoBancaria")) Else objConta.SituacaoBancaria = 0
                objConta.TipoFinanceiro = drConta("TipoFinanceiro")
                objConta.CarteiraDoTitulo = drConta("CarteiraDoTitulo").ToString()
                objConta.ContratoBancario = drConta("ContratoBancario").ToString()

                Dim cart As New CarteiraFinanceira(objConta.Carteira)
                Dim cop As New PlanoDeConta("", 0, objConta.ContaContabilPagadora)

                If Not (cart.BaixaAdiantamento Or cop.Adiantamento) And _
                    (Me.EntradaSaida = eEntradaSaida.Saida And drConta("TipoFinanceiro") = "P" OrElse _
                   Me.EntradaSaida = eEntradaSaida.Entrada And drConta("TipoFinanceiro") = "R") Then
                    Me.OficialADevolver += objConta.ValorDocumentoOficial
                    Me.OficialDevolvido += IIf(objConta.Provisao = eProvisao.Baixa, objConta.ValorDocumentoOficial, 0)
                    Me.MoedaADevolver += objConta.ValorDocumentoMoeda
                    Me.MoedaDevolvido += IIf(objConta.Provisao = eProvisao.Baixa, objConta.ValorDocumentoMoeda, 0)
                    objConta.Provisao = CType(Convert.ToInt32(1), eProvisao)
                Else
                    Me.OficialAMovimentar += objConta.ValorDocumentoOficial
                    Me.OficialMovimentado += IIf(objConta.Provisao = eProvisao.Baixa, objConta.ValorDocumentoOficial, 0)
                    Me.MoedaAMovimentar += objConta.ValorDocumentoMoeda
                    Me.MoedaMovimentado += IIf(objConta.Provisao = eProvisao.Baixa, objConta.ValorDocumentoMoeda, 0)
                End If

                'If Not objPedido Is Nothing Then
                '    If objPedido.MomentoFinanceiro = 3 And objConta.Provisao = 2 Then
                '        objConta.Provisao = CType(Convert.ToInt32(1), eProvisao)
                '    End If
                'End If

                Me.Add(objConta)
            Next

            Return True
        Catch ex As Exception
            Throw New Exception(ex.Message)
            Return False
        Finally
            objBanco = Nothing
        End Try
    End Function

#End Region

End Class