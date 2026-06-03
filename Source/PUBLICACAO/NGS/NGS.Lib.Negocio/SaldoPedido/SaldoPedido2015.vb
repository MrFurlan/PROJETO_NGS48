Imports Microsoft.VisualBasic
Imports System.Data
Imports NGS.Lib.Uteis

'*******************************************************************************************************************
'***********************************  LISTA SALDO PEDIDO  **********************************************************
'*******************************************************************************************************************
Public Class ListSaldoPedido2015
    Inherits List(Of SaldoPedido2015)
    Implements IBaseEntity

#Region "Construtor"
    Public Sub New()

    End Sub

    Public Sub New(ByVal Parametros As Hashtable)
        '*******************************
        '***** PARAMETROS TRATATOS *****
        '*******************************
        'TipoApuracao      integer      = 0,     0 Sintetico - 1 Analitico - 2 - Analalitico sem laudo   
        'TipoPedido        nvarchar(5)  = NULL,  1 - global/Direto 2 - afixar  3 - Deposito ex: 1,2 ou 2,3 ou 1,3 ou 1,2,3 que é o mesmo q nao passar o parametro q no caso seria o NULL
        'Empresa           nvarchar(18) = NULL,
        'EndEmpresa        int          = NULL, 
        'Cliente           nvarchar(18) = NULL,        
        'EndCliente        int          = NULL, 
        'FilialDev         nvarchar(18) = NULL, 
        'EndFilialDev      int          = NULL, 
        'DataReferencia    datetime     = NULL, 
        'Safra             nvarchar(50) = NULL, 
        'Pedido            int          = NULL, 
        'Saldo             int          = NULL, -- 0 Sem Saldo, 1 Com Saldo, NULL Todos 
        'Fiscal            int          = NULL, -- 0 Fechado  , 1 Aberto   , NULL Todos
        'PeriodoInicial    datetime     = NULL,
        'PeriodoFinal      datetime     = NULL,
        'Produto           int          = NULL
        '************************************************** Da Suboperacao
        'Classe            nvarchar(20)  = Classe da operacao do pedido
        'Devolucao         bit           = 0    0 Nao - 1 sim
        'SoValor           bit           = 0    0 Nao - 1 sim  consiste e nos campo qtde fisica e fiscal da suboperacao serem false
        'Nota              int          = NULL, Caso seja uma alteração na NF então pode ser enviada a NF e a serie para que o saldo do pedido não a leve em conta
        'Serie             nvarchar(3)  = NULL

        Dim Sql As String = "spSaldoPedido "
        Dim SqlParametros As New System.Data.SqlClient.SqlCommand
        Dim DataRef As Date = Now.Date

        If Parametros.Contains("TipoApuracao") Then SqlParametros.Parameters.AddWithValue("@TipoApuracao", Parametros("TipoApuracao"))
        If Parametros.Contains("TipoPedido") Then SqlParametros.Parameters.AddWithValue("@TipoPedido", Parametros("TipoPedido"))
        If Parametros.Contains("Empresa") Then SqlParametros.Parameters.AddWithValue("@Empresa", Parametros("Empresa"))
        If Parametros.Contains("EndEmpresa") Then SqlParametros.Parameters.AddWithValue("@EndEmpresa", Parametros("EndEmpresa"))
        If Parametros.Contains("Cliente") Then SqlParametros.Parameters.AddWithValue("@Cliente", Parametros("Cliente"))
        If Parametros.Contains("EndCliente") Then SqlParametros.Parameters.AddWithValue("@EndCliente", Parametros("EndCliente"))
        If Parametros.Contains("FilialDev") Then SqlParametros.Parameters.AddWithValue("@FilialDev", Parametros("FilialDev"))
        If Parametros.Contains("EndFilialDev") Then SqlParametros.Parameters.AddWithValue("@EndFilialDev", Parametros("EndFilialDev"))

        If Parametros.Contains("DataReferencia") Then
            SqlParametros.Parameters.AddWithValue("@DataReferencia", Parametros("DataReferencia"))
            DataRef = CDate(Parametros("DataReferencia"))
        End If

        If Parametros.Contains("Safra") Then SqlParametros.Parameters.AddWithValue("@Safra", Parametros("Safra"))
        If Parametros.Contains("Pedido") Then SqlParametros.Parameters.AddWithValue("@Pedido", Parametros("Pedido"))
        If Parametros.Contains("PedidoEfetivo") Then SqlParametros.Parameters.AddWithValue("@PedidoEfetivo", Parametros("PedidoEfetivo"))
        If Parametros.Contains("Contrato") Then SqlParametros.Parameters.AddWithValue("@Contrato", Parametros("Contrato"))
        If Parametros.Contains("Saldo") Then SqlParametros.Parameters.AddWithValue("@Saldo", Parametros("Saldo"))
        If Parametros.Contains("Fiscal") Then SqlParametros.Parameters.AddWithValue("@Fiscal", Parametros("Fiscal"))
        If Parametros.Contains("PeriodoInicial") Then SqlParametros.Parameters.AddWithValue("@PeriodoInicial", Parametros("PeriodoInicial"))
        If Parametros.Contains("PeriodoFinal") Then SqlParametros.Parameters.AddWithValue("@PeriodoFinal", Parametros("PeriodoFinal"))
        If Parametros.Contains("Operacao") Then SqlParametros.Parameters.AddWithValue("@Operacao", Parametros("Operacao"))
        If Parametros.Contains("SubOperacao") Then SqlParametros.Parameters.AddWithValue("@SubOperacao", Parametros("SubOperacao"))
        If Parametros.Contains("Produto") Then SqlParametros.Parameters.AddWithValue("@Produto", Parametros("Produto"))

        If Parametros.Contains("Classe") Then SqlParametros.Parameters.AddWithValue("@Classe", Parametros("Classe"))
        If Parametros.Contains("Devolucao") Then SqlParametros.Parameters.AddWithValue("@Devolucao", Parametros("Devolucao"))
        If Parametros.Contains("SoValor") Then SqlParametros.Parameters.AddWithValue("@SoValor", Parametros("SoValor"))

        If Parametros.Contains("ExcetoNota") Then SqlParametros.Parameters.AddWithValue("@ExcetoNota", Parametros("ExcetoNota"))
        If Parametros.Contains("ExcetoSerie") Then SqlParametros.Parameters.AddWithValue("@ExcetoSerie", Parametros("ExcetoSerie"))
        If Parametros.Contains("ExcetoCliente") Then SqlParametros.Parameters.AddWithValue("@ExcetoCliente", Parametros("ExcetoCliente"))
        If Parametros.Contains("ExcetoEndCliente") Then SqlParametros.Parameters.AddWithValue("@ExcetoEndCliente", Parametros("ExcetoEndCliente"))
        If Parametros.Contains("ExcetoEntradaSaida") Then SqlParametros.Parameters.AddWithValue("@ExcetoEntradaSaida", Parametros("ExcetoEntradaSaida"))
        If Parametros.Contains("NCM") Then SqlParametros.Parameters.AddWithValue("@NCM", Parametros("NCM"))
        If Parametros.Contains("ParteNomeDoProduto") Then SqlParametros.Parameters.AddWithValue("@ParteNomeDoProduto", Parametros("ParteNomeDoProduto"))
        'If Parametros.Contains("TipoApuracao") Then SqlParametros.Parameters.AddWithValue("@TipoApuracao", Parametros("TipoApuracao"))

        '' Montar o script
        'Dim scriptBuilder As New System.Text.StringBuilder()

        'For Each param As System.Data.SqlClient.SqlParameter In SqlParametros.Parameters
        '    Dim paramName As String = param.ParameterName
        '    Dim paramValue As String

        '    ' Verificar se o valor é DBNull
        '    If param.Value Is DBNull.Value Then
        '        paramValue = "NULL"
        '    Else
        '        ' Converter o valor para string
        '        ' Colocar valores string entre aspas simples
        '        If TypeOf param.Value Is String Then
        '            paramValue = $"'{param.Value.ToString().Replace("'", "''")}'"
        '        Else
        '            paramValue = param.Value.ToString()
        '        End If
        '    End If

        '    ' Adicionar a linha ao script com vírgula separando
        '    scriptBuilder.Append($"{paramName} = {paramValue}, ")
        'Next

        '' Remover a última vírgula e espaço
        'If scriptBuilder.Length > 2 Then
        '    scriptBuilder.Length -= 2
        'End If

        '' Mostrar o script gerado
        'Dim script As String = scriptBuilder.ToString()

        Dim ds As DataSet
        Dim banco As New AcessaBanco
        ds = banco.ConsultaDataSet(Sql, "Saldo", CommandType.StoredProcedure, SqlParametros.Parameters)

        If ds IsNot Nothing AndAlso ds.Tables.Count > 0 Then
            For Each row As DataRow In ds.Tables(0).Rows

                Dim SP As New SaldoPedido2015
                SP.Tipo = row("Tipo")
                SP.DataReferencia = DataRef

                SP.CodigoEmpresa = row("Empresa_Id")
                SP.EndEmpresa = row("EndEmpresa_Id")

                SP.CodigoCliente = row("Cliente")
                SP.EndCliente = row("EndCliente")
                SP.DescricaoCliente = row("DescricaoCliente")

                SP.CodigoPedido = row("Pedido_Id")
                SP.PedidoEfetivo = row("PedidoEfetivo")
                SP.Contrato = row("Contrato")

                SP.DataPedido = row("DataPedido")
                SP.FiscalAberto = row("FiscalAberto")
                SP.CodigoSafra = row("Safra")

                SP.CodigoProduto = row("Produto_id")

                If SP.Empresa.Empresa.UsarRegistroMinAgr Then
                    SP.NomeProduto = SP.Produto.Nome & "-" & SP.Produto.Descricao & "(" & SP.Produto.RegistroMinisterioAgricultura & ")"
                ElseIf SP.Empresa.Empresa.UsarDescricaoProduto Then
                    SP.NomeProduto = SP.Produto.Nome & "-" & SP.Produto.Descricao
                Else
                    SP.NomeProduto = SP.Produto.Nome
                End If

                SP.Lote = row("Lote")
                SP.Classificacao = row("Classificacao")
                SP.CodigoEmbalagem = row("Embalagem")
                SP.TipoDeEmbalagem = row("TipoDeEmbalagem")
                SP.CapacidadeEmbalagem = row("CapacidadeEmbalagem")
                SP.DescricaoEmbalagem = row("DescricaoEmbalagem")

                SP.Unidade = row("Unidade")
                SP.UnidadeComercializacao = row("UnidadeComercializacao")
                SP.FatorConversao = row("FatorConversao")

                SP.CodigoOperacao = row("Operacao")
                SP.CodigoSubOperacao = row("SubOperacao")
                SP.DescricaoSuboperacao = row("Descricao")

                SP.PrecoFixo = row("PrecoFixo") = "S" ' boolean
                SP.CodigoMoeda = row("Moeda")
                SP.DescricaoMoeda = row("DescricaoMoeda")
                SP.CifraoOficial = row("CifraoOficial")
                SP.CifraoPedido = row("CifraoPedido")


                '***********************************************************************************************
                '***********  PEDIDO  **************************************************************************
                '***********************************************************************************************
                SP.QtdeProgramada = row("QtdeProgramada")
                SP.QtdeProgramadaComercializacao = row("QtdeProgramadaComercializacao")

                SP.UnitarioOficial = row("UnitarioOficial")
                SP.UnitarioMoeda = row("UnitarioMoeda")

                SP.UnitarioComercializacaoOficial = row("unitarioComercializacaoOficial")
                SP.UnitarioComercializacaoMoeda = row("unitarioComercializacaoMoeda")

                SP.VlrPedidoOficial = row("VlrPedidoOficial")
                SP.VlrPedidoMoeda = row("VlrPedidoMoeda")

                '***********************************************************************************************
                '***********  FIXACAO  *************************************************************************
                '***********************************************************************************************
                SP.QtdeFixacao = row("QtdeFixacao")
                SP.UntFixacaoOficial = row("UntFixacaoOficial")
                SP.VlrFixacaoOficial = row("VlrFixacaoOficial")
                SP.UntFixacaoMoeda = row("UntFixacaoMoeda")
                SP.VlrFixacaoMoeda = row("VlrFixacaoMoeda")

                SP.VlrFixacaoNF = row("VlrFixacaoNF")
                '***********************************************************************************************
                '***********  CONTRATADO  **********************************************************************
                '***********************************************************************************************
                SP.QtdeContratadoFiscal = row("QtdeContratadoFiscal")
                SP.QtdeContratadoFisico = row("QtdeContratadoFisico")

                '***********************************************************************************************
                '***********  ENTREGUE FISCAL  *****************************************************************
                '***********************************************************************************************
                SP.QtdeEntregueFiscalGlobal = row("QtdeEntregueFiscalGlobal")
                SP.QtdeEntregueFiscalRemessa = row("QtdeEntregueFiscalRemessa")
                SP.QtdeEntregueFiscalAFixar = row("QtdeEntregueFiscalAFixar")
                SP.QtdeEntregueFiscalDeposito = row("QtdeEntregueFiscalDeposito")
                SP.QtdeEntregueFiscalDireta = row("QtdeEntregueFiscalDireta")
                SP.QtdeComercializacaoEntregue = row("QtdeComercializacaoEntregue")

                '***********************************************************************************************
                '***********  ENTREGUE FISICO  *****************************************************************
                '***********************************************************************************************
                SP.QtdeEntregueFisicoGlobal = row("QtdeEntregueFisicoGlobal")
                SP.QtdeEntregueFisicoRemessa = row("QtdeEntregueFisicoRemessa")
                SP.QtdeEntregueFisicoAFixar = row("QtdeEntregueFisicoAFixar")
                SP.QtdeEntregueFisicoDeposito = row("QtdeEntregueFisicoDeposito")
                SP.QtdeEntregueFisicoDireta = row("QtdeEntregueFisicoDireta")

                '***********************************************************************************************
                '***********  ENTREGUE VALOR  ******************************************************************
                '***********************************************************************************************
                SP.VlrNotaOficialGlobalBruto = row("VlrNotaOficialGlobalBruto")
                SP.VlrNotaOficialRemessaBruto = row("VlrNotaOficialRemessaBruto")
                SP.VlrNotaOficialAFixarBruto = row("VlrNotaOficialAFixarBruto")
                SP.VlrNotaOficialDepositoBruto = row("VlrNotaOficialDepositoBruto")
                SP.VlrNotaOficialDiretaBruto = row("VlrNotaOficialDiretaBruto")

                SP.VlrNotaOficialGlobalLiquido = row("VlrNotaOficialGlobalLiquido")
                SP.VlrNotaOficialRemessaLiquido = row("VlrNotaOficialRemessaLiquido")
                SP.VlrNotaOficialAFixarLiquido = row("VlrNotaOficialAFixarLiquido")
                SP.VlrNotaOficialDepositoLiquido = row("VlrNotaOficialDepositoLiquido")
                SP.VlrNotaOficialDiretaLiquido = row("VlrNotaOficialDiretaLiquido")

                SP.VlrNotaMoedaGlobalBruto = row("VlrNotaMoedaGlobalBruto")
                SP.VlrNotaMoedaRemessaBruto = row("VlrNotaMoedaRemessaBruto")
                SP.VlrNotaMoedaAFixarBruto = row("VlrNotaMoedaAFixarBruto")
                SP.VlrNotaMoedaDepositoBruto = row("VlrNotaMoedaDepositoBruto")
                SP.VlrNotaMoedaDiretaBruto = row("VlrNotaMoedaDiretaBruto")

                SP.VlrNotaMoedaGlobalLiquido = row("VlrNotaMoedaGlobalLiquido")
                SP.VlrNotaMoedaRemessaLiquido = row("VlrNotaMoedaRemessaLiquido")
                SP.VlrNotaMoedaAFixarLiquido = row("VlrNotaMoedaAFixarLiquido")
                SP.VlrNotaMoedaDepositoLiquido = row("VlrNotaMoedaDepositoLiquido")
                SP.VlrNotaMoedaDiretaLiquido = row("VlrNotaMoedaDiretaLiquido")

                '***********************************************************************************************
                '***********  SALDO  ***************************************************************************
                '***********************************************************************************************
                SP.SaldoQtdeGlobalFiscal = row("SaldoQtdeGlobalFiscal")
                SP.SaldoQtdeRemessaFiscal = row("SaldoQtdeRemessaFiscal")
                SP.SaldoQtdeRemessaFisica = row("SaldoQtdeRemessaFisica")
                SP.SaldoQtdeDiretoFiscal = row("SaldoQtdeDiretoFiscal")
                SP.SaldoQtdeDiretoFisica = row("SaldoQtdeDiretoFisica")
                SP.SaldoQtdeComercializacao = row("SaldoQtdeComercializacao")
                SP.SaldoValorOficialGlobalDireto = row("SaldoValorOficialGlobalDireto")
                SP.SaldoValorMoedaGlobalDireto = row("SaldoValorMoedaGlobalDireto")
                SP.SaldoValorOficialRemessa = row("SaldoValorOficialRemessa")
                SP.SaldoValorMoedaRemessa = row("SaldoValorMoedaRemessa")

                'SP.SaldoValorOficialGlobalDiretoFilial = row("SaldoValorOficialGlobalDiretoFilial")
                'SP.SaldoValorOficialRemessaFilial = row("SaldoValorOficialRemessaFilial")

                SP.SaldoQtdeAFixar = row("SaldoQtdeAFixar")

                '***********************************************************************************************
                '***********  XML  *****************************************************************************
                '***********************************************************************************************


                Me.Add(SP)
            Next
        End If
    End Sub
#End Region

End Class

'*******************************************************************************************************************
'***************************************  SALDO PEDIDO  ************************************************************
'*******************************************************************************************************************
Public Class SaldoPedido2015
    Implements IBaseEntity

#Region "Contrutor"
    Public Sub New()

    End Sub

    Public Sub New(pCodigoEmpresa As String, pEndEmpresa As Integer, pCodigoPedido As Integer, pCodigoProduto As Integer)
        Dim par As New Hashtable
        par.Add("Tipo", 1)
        par.Add("Empresa", pCodigoEmpresa)
        par.Add("Pedido", pCodigoPedido)
        par.Add("Produto", pCodigoProduto)
        Dim list As New ListSaldoPedido2015(par)
        Dim SD As SaldoPedido2015 = list(0)

        Me.Tipo = SD.Tipo

        Me.CodigoEmpresa = SD.CodigoEmpresa
        Me.EndEmpresa = SD.EndEmpresa

        Me.CodigoCliente = SD.CodigoCliente
        Me.EndCliente = SD.EndCliente
        Me.DescricaoCliente = SD.DescricaoCliente

        Me.CodigoPedido = SD.CodigoPedido
        Me.PedidoEfetivo = SD.PedidoEfetivo
        Me.DataPedido = SD.DataPedido
        Me.FiscalAberto = SD.FiscalAberto
        Me.CodigoSafra = SD.CodigoSafra

        Me.CodigoProduto = SD.CodigoProduto

        If Me.Empresa.Empresa.UsarRegistroMinAgr Then
            Me.NomeProduto = Me.Produto.Nome & "-" & Me.Produto.Descricao & "(" & Me.Produto.RegistroMinisterioAgricultura & ")"
        ElseIf Me.Empresa.Empresa.UsarDescricaoProduto Then
            Me.NomeProduto = Me.Produto.Nome & "-" & Me.Produto.Descricao
        Else
            Me.NomeProduto = Me.Produto.Nome
        End If

        Me.Lote = SD.Lote
        Me.Classificacao = SD.Classificacao
        Me.CodigoEmbalagem = SD.CodigoEmbalagem
        Me.TipoDeEmbalagem = SD.TipoDeEmbalagem
        Me.CapacidadeEmbalagem = SD.CapacidadeEmbalagem
        Me.DescricaoEmbalagem = SD.DescricaoEmbalagem
        Me.Unidade = SD.Unidade
        Me.UnidadeComercializacao = SD.UnidadeComercializacao
        Me.FatorConversao = SD.FatorConversao

        Me.CodigoOperacao = SD.CodigoOperacao
        Me.CodigoSubOperacao = SD.CodigoSubOperacao
        Me.DescricaoSuboperacao = SD.DescricaoSuboperacao
        Me.CifraoOficial = SD.CifraoOficial
        Me.CifraoPedido = SD.CifraoPedido

        Me.PrecoFixo = SD.PrecoFixo
        Me.CodigoMoeda = SD.CodigoMoeda
        Me.DescricaoMoeda = SD.DescricaoMoeda

        '***********************************************************************************************
        '***********  PEDIDO  **************************************************************************
        '***********************************************************************************************
        Me.QtdeProgramada = SD.QtdeProgramada
        Me.QtdeProgramadaComercializacao = SD.QtdeProgramadaComercializacao

        Me.UnitarioOficial = SD.UnitarioOficial
        Me.UnitarioMoeda = SD.UnitarioMoeda

        Me.UnitarioComercializacaoOficial = SD.UnitarioComercializacaoOficial
        Me.UnitarioComercializacaoMoeda = SD.UnitarioComercializacaoMoeda

        Me.VlrPedidoOficial = SD.VlrPedidoOficial
        Me.VlrPedidoMoeda = SD.VlrPedidoMoeda

        '***********************************************************************************************
        '***********  FIXACAO  *************************************************************************
        '***********************************************************************************************
        Me.QtdeFixacao = SD.QtdeFixacao
        Me.UntFixacaoOficial = SD.UntFixacaoOficial
        Me.VlrFixacaoOficial = SD.VlrFixacaoOficial
        Me.UntFixacaoMoeda = SD.UntFixacaoMoeda
        Me.VlrFixacaoMoeda = SD.VlrFixacaoMoeda

        '***********************************************************************************************
        '***********  CONTRATADO  **********************************************************************
        '***********************************************************************************************
        Me.QtdeContratadoFiscal = SD.QtdeContratadoFiscal
        Me.QtdeContratadoFisico = SD.QtdeContratadoFisico

        '***********************************************************************************************
        '***********  ENTREGUE FISCAL  *****************************************************************
        '***********************************************************************************************
        Me.QtdeEntregueFiscalGlobal = SD.QtdeEntregueFiscalGlobal
        Me.QtdeEntregueFiscalRemessa = SD.QtdeEntregueFiscalRemessa
        Me.QtdeEntregueFiscalAFixar = SD.QtdeEntregueFiscalAFixar
        Me.QtdeEntregueFiscalDeposito = SD.QtdeEntregueFiscalDeposito
        Me.QtdeEntregueFiscalDireta = SD.QtdeEntregueFiscalDireta
        Me.QtdeComercializacaoEntregue = SD.QtdeComercializacaoEntregue

        '***********************************************************************************************
        '***********  ENTREGUE FISICO  *****************************************************************
        '***********************************************************************************************
        Me.QtdeEntregueFisicoGlobal = SD.QtdeEntregueFisicoGlobal
        Me.QtdeEntregueFisicoRemessa = SD.QtdeEntregueFisicoRemessa
        Me.QtdeEntregueFisicoAFixar = SD.QtdeEntregueFisicoAFixar
        Me.QtdeEntregueFisicoDeposito = SD.QtdeEntregueFisicoDeposito
        Me.QtdeEntregueFisicoDireta = SD.QtdeEntregueFisicoDireta

        '***********************************************************************************************
        '***********  ENTREGUE VALOR  ******************************************************************
        '***********************************************************************************************
        Me.VlrNotaOficialGlobalBruto = SD.VlrNotaOficialGlobalBruto
        Me.VlrNotaOficialRemessaBruto = SD.VlrNotaOficialRemessaBruto
        Me.VlrNotaOficialAFixarBruto = SD.VlrNotaOficialAFixarBruto
        Me.VlrNotaOficialDepositoBruto = SD.VlrNotaOficialDepositoBruto
        Me.VlrNotaOficialDiretaBruto = SD.VlrNotaOficialDiretaBruto

        Me.VlrNotaOficialGlobalLiquido = SD.VlrNotaOficialGlobalLiquido
        Me.VlrNotaOficialRemessaLiquido = SD.VlrNotaOficialRemessaLiquido
        Me.VlrNotaOficialAFixarLiquido = SD.VlrNotaOficialAFixarLiquido
        Me.VlrNotaOficialDepositoLiquido = SD.VlrNotaOficialDepositoLiquido
        Me.VlrNotaOficialDiretaLiquido = SD.VlrNotaOficialDiretaLiquido

        Me.VlrNotaMoedaGlobalBruto = SD.VlrNotaMoedaGlobalBruto
        Me.VlrNotaMoedaRemessaBruto = SD.VlrNotaMoedaRemessaBruto
        Me.VlrNotaMoedaAFixarBruto = SD.VlrNotaMoedaAFixarBruto
        Me.VlrNotaMoedaDepositoBruto = SD.VlrNotaMoedaDepositoBruto
        Me.VlrNotaMoedaDiretaBruto = SD.VlrNotaMoedaDiretaBruto

        Me.VlrNotaMoedaGlobalLiquido = SD.VlrNotaMoedaGlobalLiquido
        Me.VlrNotaMoedaRemessaLiquido = SD.VlrNotaMoedaRemessaLiquido
        Me.VlrNotaMoedaAFixarLiquido = SD.VlrNotaMoedaAFixarLiquido
        Me.VlrNotaMoedaDepositoLiquido = SD.VlrNotaMoedaDepositoLiquido
        Me.VlrNotaMoedaDiretaLiquido = SD.VlrNotaMoedaDiretaLiquido

        '***********************************************************************************************
        '***********  SALDO  ***************************************************************************
        '***********************************************************************************************
        Me.SaldoQtdeGlobalFiscal = SD.SaldoQtdeGlobalFiscal
        Me.SaldoQtdeRemessaFiscal = SD.SaldoQtdeRemessaFiscal
        Me.SaldoQtdeRemessaFisica = SD.SaldoQtdeRemessaFisica
        Me.SaldoQtdeDiretoFiscal = SD.SaldoQtdeDiretoFiscal
        Me.SaldoQtdeDiretoFisica = SD.SaldoQtdeDiretoFisica
        Me.SaldoQtdeComercializacao = SD.SaldoQtdeComercializacao
        Me.SaldoValorOficialGlobalDireto = SD.SaldoValorOficialGlobalDireto
        Me.SaldoValorMoedaGlobalDireto = SD.SaldoValorMoedaGlobalDireto
        Me.SaldoValorOficialRemessa = SD.SaldoValorOficialRemessa
        Me.SaldoValorMoedaRemessa = SD.SaldoValorMoedaRemessa

        'Me.SaldoValorOficialGlobalDiretoFilial = SD.SaldoValorOficialGlobalDiretoFilial
        'Me.SaldoValorOficialRemessaFilial = SD.SaldoValorOficialRemessaFilial

        Me.SaldoQtdeAFixar = SD.SaldoQtdeAFixar
    End Sub
#End Region

#Region "Fields"
    Private _Selecionado As Boolean = False
    Private _Tipo As Integer  '1 Global/Remessa/Direta   2 AFIXAR   3 DEPOSITO 
    Private _DataReferencia As Date

    Private _Empresa As Cliente
    Private _CodigoEmpresa As String
    Private _EndEmpresa As Integer
    Private _DescricaoEmpresa As String

    Private _Cliente As Cliente
    Private _CodigoCliente As String
    Private _EndCliente As Integer
    Private _DescricaoCliente As String

    Private _CodigoPedido As String
    Private _Pedido As Pedido

    Private _PedidoEfetivo As String
    Private _Contrato As String

    Private _DataPedido As Date
    Private _FiscalAberto As Boolean
    Private _CodigoSafra As String

    Private _CodigoProduto As Integer
    Private _Produto As Produto
    Private _NomeProduto As String

    Private _Lote As String
    Private _Classificacao As String
    Private _CodigoEmbalagem As Integer
    Private _Embalagem As Embalagem
    Private _TipoDeEmbalagem As String
    Private _CapacidadeEmbalagem As Decimal
    Private _DescricaoEmbalagem As String
    Private _Unidade As String
    Private _UnidadeComercializacao As String
    Private _FatorConversao As Decimal

    Private _CodigoOperacao As Integer
    Private _CodigoSubOperacao As Integer
    Private _SubOperacao As SubOperacao
    Private _DescricaoSuboperacao As String

    Private _PrecoFixo As Boolean
    Private _CodigoMoeda As Integer
    Private _DescricaoMoeda As String
    Private _CifraoPedido As String
    Private _CifraoOficial As String

    '***********************************************************************************************
    '***********  PEDIDO  **************************************************************************
    '***********************************************************************************************
    Private _QtdeProgramada As Decimal
    Private _QtdeProgramadaComercializacao As Decimal

    Private _UnitarioOficial As Decimal
    Private _UnitarioMoeda As Decimal

    Private _UnitarioComercializacaoOficial As Decimal
    Private _UnitarioComercializacaoMoeda As Decimal

    Private _VlrPedidoOficial As Decimal
    Private _VlrPedidoMoeda As Decimal

    '***********************************************************************************************
    '***********  FIXACAO  *************************************************************************
    '***********************************************************************************************
    Private _QtdeFixacao As Decimal
    Private _UntFixacaoOficial As Decimal
    Private _VlrFixacaoOficial As Decimal
    Private _UntFixacaoMoeda As Decimal
    Private _VlrFixacaoMoeda As Decimal

    Private _VlrFixacaoNF As Decimal

    '***********************************************************************************************
    '***********  CONTRATADO  **********************************************************************
    '***********************************************************************************************
    Private _QtdeContratadoFiscal As Decimal
    Private _QtdeContratadoFisico As Decimal

    '***********************************************************************************************
    '***********  ENTREGUE FISCAL  *****************************************************************
    '***********************************************************************************************
    Private _QtdeEntregueFiscalGlobal As Decimal
    Private _QtdeEntregueFiscalRemessa As Decimal
    Private _QtdeEntregueFiscalAFixar As Decimal
    Private _QtdeEntregueFiscalDeposito As Decimal
    Private _QtdeEntregueFiscalDireta As Decimal
    Private _QtdeComercializacaoEntregue As Decimal

    '***********************************************************************************************
    '***********  ENTREGUE FISICO  *****************************************************************
    '***********************************************************************************************
    Private _QtdeEntregueFisicoGlobal As Decimal
    Private _QtdeEntregueFisicoRemessa As Decimal
    Private _QtdeEntregueFisicoAFixar As Decimal
    Private _QtdeEntregueFisicoDeposito As Decimal
    Private _QtdeEntregueFisicoDireta As Decimal

    '***********************************************************************************************
    '***********  ENTREGUE VALOR  ******************************************************************
    '***********************************************************************************************
    Private _VlrNotaOficialGlobalBruto As Decimal
    Private _VlrNotaOficialRemessaBruto As Decimal
    Private _VlrNotaOficialAFixarBruto As Decimal
    Private _VlrNotaOficialDepositoBruto As Decimal
    Private _VlrNotaOficialDiretaBruto As Decimal
    Private _VlrNotaOficialGlobalLiquido As Decimal
    Private _VlrNotaOficialRemessaLiquido As Decimal
    Private _VlrNotaOficialAFixarLiquido As Decimal
    Private _VlrNotaOficialDepositoLiquido As Decimal
    Private _VlrNotaOficialDiretaLiquido As Decimal
    Private _VlrNotaMoedaGlobalBruto As Decimal
    Private _VlrNotaMoedaRemessaBruto As Decimal
    Private _VlrNotaMoedaAFixarBruto As Decimal
    Private _VlrNotaMoedaDepositoBruto As Decimal
    Private _VlrNotaMoedaDiretaBruto As Decimal
    Private _VlrNotaMoedaGlobalLiquido As Decimal
    Private _VlrNotaMoedaRemessaLiquido As Decimal
    Private _VlrNotaMoedaAFixarLiquido As Decimal
    Private _VlrNotaMoedaDepositoLiquido As Decimal
    Private _VlrNotaMoedaDiretaLiquido As Decimal

    '***********************************************************************************************
    '***********  SALDO  ***************************************************************************
    '***********************************************************************************************
    Private _SaldoQtdeGlobalFiscal As Decimal
    Private _SaldoQtdeRemessaFiscal As Decimal
    Private _SaldoQtdeRemessaFisica As Decimal
    Private _SaldoQtdeDiretoFiscal As Decimal
    Private _SaldoQtdeDiretoFisica As Decimal
    Private _SaldoQtdeComercializacao As Decimal
    Private _SaldoValorOficialGlobalDireto As Decimal
    Private _SaldoValorMoedaGlobalDireto As Decimal
    Private _SaldoValorOficialRemessa As Decimal
    Private _SaldoValorMoedaRemessa As Decimal


    'Private _SaldoValorOficialGlobalDiretoFilial As Decimal
    'Private _SaldoValorOficialRemessaFilial As Decimal

    Private _SaldoQtdeAFixar As Decimal

    '***********************************************************************************************
    '***********  LISTA DE SALDO POR PRODUTO  ******************************************************
    '***********************************************************************************************
    Private _Itens As ListSaldoPedido2015


    '***********************************************************************************************
    '***********  XML   ****************************************************************************
    '***********************************************************************************************
    Private _CodigoXmlProdutoXDePara As Integer
    Private _XmlProdutoXDePara As XmlProdutoXDePara
    Private _NomeXmlProdutoXDePara As String




#End Region

#Region "Property"
    Public Property Selecionado As Boolean
        Get
            Return _Selecionado
        End Get
        Set(value As Boolean)
            _Selecionado = value
        End Set
    End Property
    Public Property DataReferencia As Date
        Get
            Return _DataReferencia
        End Get
        Set(value As Date)
            _DataReferencia = value
        End Set
    End Property

    Public Property Tipo() As Integer
        Get
            Return _Tipo
        End Get
        Set(ByVal value As Integer)
            _Tipo = value
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
    Public Property EndEmpresa As Integer
        Get
            Return _EndEmpresa
        End Get
        Set(value As Integer)
            _EndEmpresa = value
        End Set
    End Property
    Public ReadOnly Property Empresa As Cliente
        Get
            If _Empresa Is Nothing And Me.CodigoEmpresa.Length > 0 Then _Empresa = New Cliente(Me.CodigoEmpresa, Me.EndEmpresa)
            Return _Empresa
        End Get
    End Property
    Public Property DescricaoEmpresa As String
        Get
            If String.IsNullOrWhiteSpace(_DescricaoEmpresa) Then _DescricaoEmpresa = Empresa.Nome
            Return _DescricaoEmpresa
        End Get
        Set(value As String)
            _DescricaoEmpresa = value
        End Set
    End Property

    Public Property Cliente As Cliente
        Get
            If _Cliente Is Nothing Then _Cliente = New Cliente(Me.CodigoCliente, Me.EndCliente)
            Return _Cliente
        End Get
        Set(value As Cliente)
            _Cliente = value
        End Set
    End Property
    Public Property CodigoCliente As String
        Get
            Return _CodigoCliente
        End Get
        Set(value As String)
            _CodigoCliente = value
        End Set
    End Property
    Public Property EndCliente As Integer
        Get
            Return _EndCliente
        End Get
        Set(value As Integer)
            _EndCliente = value
        End Set
    End Property
    Public Property DescricaoCliente As String
        Get
            Return _DescricaoCliente
        End Get
        Set(value As String)
            _DescricaoCliente = value
        End Set
    End Property

    Public Property CodigoPedido As String
        Get
            Return _CodigoPedido
        End Get
        Set(value As String)
            _CodigoPedido = value
        End Set
    End Property
    Public ReadOnly Property Pedido As Pedido
        Get
            If _Pedido Is Nothing Then _Pedido = New Pedido(Me.CodigoEmpresa, Me.EndEmpresa, Me.CodigoPedido)
            Return _Pedido
        End Get
    End Property

    Public Property PedidoEfetivo As String
        Get
            Return _PedidoEfetivo
        End Get
        Set(value As String)
            _PedidoEfetivo = value
        End Set
    End Property

    Public Property DataPedido As Date
        Get
            Return _DataPedido
        End Get
        Set(value As Date)
            _DataPedido = value
        End Set
    End Property
    Public Property FiscalAberto As Boolean
        Get
            Return _FiscalAberto
        End Get
        Set(value As Boolean)
            _FiscalAberto = value
        End Set
    End Property
    Public Property CodigoSafra As String
        Get
            Return _CodigoSafra
        End Get
        Set(value As String)
            _CodigoSafra = value
        End Set
    End Property

    Public Property CodigoProduto As Integer
        Get
            Return _CodigoProduto
        End Get
        Set(value As Integer)
            _CodigoProduto = value
        End Set
    End Property
    Public ReadOnly Property Produto As Produto
        Get
            If _Produto Is Nothing Then _Produto = New Produto(Me.CodigoProduto)
            Return _Produto
        End Get
    End Property
    Public Property NomeProduto As String
        Get
            If _Empresa.Empresa.UsarRegistroMinAgr Then
                If String.IsNullOrWhiteSpace(_NomeProduto) Then _NomeProduto = (Produto.Nome & "-" & Produto.Descricao & "(" & Produto.RegistroMinisterioAgricultura & ")")
            ElseIf _Empresa.Empresa.UsarDescricaoProduto Then
                If String.IsNullOrWhiteSpace(_NomeProduto) Then _NomeProduto = (Produto.Nome & "-" & Produto.Descricao)
            Else
                If String.IsNullOrWhiteSpace(_NomeProduto) Then _NomeProduto = Produto.Nome
            End If

            Return _NomeProduto
        End Get
        Set(value As String)
            _NomeProduto = value
        End Set
    End Property



    Public Property Lote As String
        Get
            Return _Lote
        End Get
        Set(value As String)
            _Lote = value
        End Set
    End Property
    Public Property Classificacao As String
        Get
            Return _Classificacao
        End Get
        Set(value As String)
            _Classificacao = value
        End Set
    End Property

    Public Property CodigoEmbalagem As Integer
        Get
            Return _CodigoEmbalagem
        End Get
        Set(value As Integer)
            _CodigoEmbalagem = value
        End Set
    End Property
    Public ReadOnly Property Embalagem() As Embalagem
        Get
            If _Embalagem Is Nothing And _CodigoEmbalagem > 0 Then _Embalagem = New Embalagem(_CodigoEmbalagem)
            Return _Embalagem
        End Get
    End Property
    Public ReadOnly Property EmbalagemIndea() As String
        Get
            If Embalagem Is Nothing Then
                Return ""
            Else
                Return Embalagem.EmbalagemIndea
            End If
        End Get
    End Property
    Public Property DescricaoEmbalagem() As String
        Get
            Return _DescricaoEmbalagem
        End Get
        Set(value As String)
            _DescricaoEmbalagem = value
        End Set
    End Property
    Public Property TipoDeEmbalagem As String
        Get
            Return _TipoDeEmbalagem
        End Get
        Set(value As String)
            _TipoDeEmbalagem = value
        End Set
    End Property
    Public Property CapacidadeEmbalagem As Decimal
        Get
            Return _CapacidadeEmbalagem
        End Get
        Set(value As Decimal)
            _CapacidadeEmbalagem = value
        End Set
    End Property

    Public Property Unidade As String
        Get
            Return _Unidade
        End Get
        Set(value As String)
            _Unidade = value
        End Set
    End Property
    Public Property UnidadeComercializacao As String
        Get
            Return _UnidadeComercializacao
        End Get
        Set(value As String)
            _UnidadeComercializacao = value
        End Set
    End Property
    Public Property FatorConversao As Decimal
        Get
            Return _FatorConversao
        End Get
        Set(value As Decimal)
            _FatorConversao = value
        End Set
    End Property

    Public Property CodigoOperacao As Integer
        Get
            Return _CodigoOperacao
        End Get
        Set(value As Integer)
            _CodigoOperacao = value
        End Set
    End Property
    Public Property CodigoSubOperacao As Integer
        Get
            Return _CodigoSubOperacao
        End Get
        Set(value As Integer)
            _CodigoSubOperacao = value
        End Set
    End Property
    Public ReadOnly Property SubOperacao As SubOperacao
        Get
            If _SubOperacao Is Nothing Then _SubOperacao = New SubOperacao(Me.CodigoOperacao, Me.CodigoSubOperacao)
            Return _SubOperacao
        End Get
    End Property
    Public Property DescricaoSuboperacao As String
        Get
            If String.IsNullOrWhiteSpace(_DescricaoSuboperacao) Then _DescricaoSuboperacao = Me.CodigoOperacao + " - " + Me.CodigoSubOperacao + ": " + SubOperacao.Descricao
            Return _DescricaoSuboperacao
        End Get
        Set(value As String)
            _DescricaoSuboperacao = value
        End Set
    End Property
    Public Property PrecoFixo As Boolean
        Get
            Return _PrecoFixo
        End Get
        Set(value As Boolean)
            _PrecoFixo = value
        End Set
    End Property
    Public Property CodigoMoeda As Integer
        Get
            Return _CodigoMoeda
        End Get
        Set(value As Integer)
            _CodigoMoeda = value
        End Set
    End Property
    Public Property DescricaoMoeda As String
        Get
            Return _DescricaoMoeda
        End Get
        Set(value As String)
            _DescricaoMoeda = value
        End Set
    End Property

    Public Property CifraoPedido As String
        Get
            Return _CifraoPedido
        End Get
        Set(value As String)
            _CifraoPedido = value
        End Set
    End Property

    Public Property CifraoOficial As String
        Get
            Return _CifraoOficial
        End Get
        Set(value As String)
            _CifraoOficial = value
        End Set
    End Property

    '***********************************************************************************************
    '***********  PEDIDO  **************************************************************************
    '***********************************************************************************************
    Public Property QtdeProgramada As Decimal
        Get
            Return _QtdeProgramada
        End Get
        Set(value As Decimal)
            _QtdeProgramada = value
        End Set
    End Property
    Public Property QtdeProgramadaComercializacao As Decimal
        Get
            Return _QtdeProgramadaComercializacao
        End Get
        Set(value As Decimal)
            _QtdeProgramadaComercializacao = value
        End Set
    End Property
    Public Property UnitarioOficial As Decimal
        Get
            Return _UnitarioOficial
        End Get
        Set(value As Decimal)
            _UnitarioOficial = value
        End Set
    End Property
    Public Property UnitarioMoeda As Decimal
        Get
            Return _UnitarioMoeda
        End Get
        Set(value As Decimal)
            _UnitarioMoeda = value
        End Set
    End Property
    Public Property UnitarioComercializacaoOficial As Decimal
        Get
            Return _UnitarioComercializacaoOficial
        End Get
        Set(value As Decimal)
            _UnitarioComercializacaoOficial = value
        End Set
    End Property
    Public Property UnitarioComercializacaoMoeda As Decimal
        Get
            Return _UnitarioComercializacaoMoeda
        End Get
        Set(value As Decimal)
            _UnitarioComercializacaoMoeda = value
        End Set
    End Property
    Public Property VlrPedidoOficial As Decimal
        Get
            Return _VlrPedidoOficial
        End Get
        Set(value As Decimal)
            _VlrPedidoOficial = value
        End Set
    End Property
    Public Property VlrPedidoMoeda As Decimal
        Get
            Return _VlrPedidoMoeda
        End Get
        Set(value As Decimal)
            _VlrPedidoMoeda = value
        End Set
    End Property

    '***********************************************************************************************
    '***********  FIXACAO  *************************************************************************
    '***********************************************************************************************
    Public Property QtdeFixacao As Decimal
        Get
            Return _QtdeFixacao
        End Get
        Set(value As Decimal)
            _QtdeFixacao = value
        End Set
    End Property
    Public Property UntFixacaoOficial As Decimal
        Get
            Return _UntFixacaoOficial
        End Get
        Set(value As Decimal)
            _UntFixacaoOficial = value
        End Set
    End Property
    Public Property VlrFixacaoOficial As Decimal
        Get
            Return _VlrFixacaoOficial
        End Get
        Set(value As Decimal)
            _VlrFixacaoOficial = value
        End Set
    End Property

    Public Property UntFixacaoMoeda As Decimal
        Get
            Return _UntFixacaoMoeda
        End Get
        Set(value As Decimal)
            _UntFixacaoMoeda = value
        End Set
    End Property
    Public Property VlrFixacaoMoeda As Decimal
        Get
            Return _VlrFixacaoMoeda
        End Get
        Set(value As Decimal)
            _VlrFixacaoMoeda = value
        End Set
    End Property

    Public Property VlrFixacaoNF As Decimal
        Get
            Return _VlrFixacaoNF
        End Get
        Set(value As Decimal)
            _VlrFixacaoNF = value
        End Set
    End Property
    '***********************************************************************************************
    '***********  CONTRATADO  **********************************************************************
    '***********************************************************************************************
    Public Property QtdeContratadoFiscal As Decimal
        Get
            Return _QtdeContratadoFiscal
        End Get
        Set(value As Decimal)
            _QtdeContratadoFiscal = value
        End Set
    End Property
    Public Property QtdeContratadoFisico As Decimal
        Get
            Return _QtdeContratadoFisico
        End Get
        Set(value As Decimal)
            _QtdeContratadoFisico = value
        End Set
    End Property

    '***********************************************************************************************
    '***********  ENTREGUE FISCAL  *****************************************************************
    '***********************************************************************************************
    Public Property QtdeEntregueFiscalGlobal As Decimal
        Get
            Return _QtdeEntregueFiscalGlobal
        End Get
        Set(value As Decimal)
            _QtdeEntregueFiscalGlobal = value
        End Set
    End Property
    Public Property QtdeEntregueFiscalRemessa As Decimal
        Get
            Return _QtdeEntregueFiscalRemessa
        End Get
        Set(value As Decimal)
            _QtdeEntregueFiscalRemessa = value
        End Set
    End Property
    Public Property QtdeEntregueFiscalAFixar As Decimal
        Get
            Return _QtdeEntregueFiscalAFixar
        End Get
        Set(value As Decimal)
            _QtdeEntregueFiscalAFixar = value
        End Set
    End Property
    Public Property QtdeEntregueFiscalDeposito As Decimal
        Get
            Return _QtdeEntregueFiscalDeposito
        End Get
        Set(value As Decimal)
            _QtdeEntregueFiscalDeposito = value
        End Set
    End Property
    Public Property QtdeEntregueFiscalDireta As Decimal
        Get
            Return _QtdeEntregueFiscalDireta
        End Get
        Set(value As Decimal)
            _QtdeEntregueFiscalDireta = value
        End Set
    End Property
    Public Property QtdeComercializacaoEntregue As Decimal
        Get
            Return _QtdeComercializacaoEntregue
        End Get
        Set(value As Decimal)
            _QtdeComercializacaoEntregue = value
        End Set
    End Property

    '***********************************************************************************************
    '***********  ENTREGUE FISICO  *****************************************************************
    '***********************************************************************************************
    Public Property QtdeEntregueFisicoGlobal As Decimal
        Get
            Return _QtdeEntregueFisicoGlobal
        End Get
        Set(value As Decimal)
            _QtdeEntregueFisicoGlobal = value
        End Set
    End Property
    Public Property QtdeEntregueFisicoRemessa As Decimal
        Get
            Return _QtdeEntregueFisicoRemessa
        End Get
        Set(value As Decimal)
            _QtdeEntregueFisicoRemessa = value
        End Set
    End Property
    Public Property QtdeEntregueFisicoAFixar As Decimal
        Get
            Return _QtdeEntregueFisicoAFixar
        End Get
        Set(value As Decimal)
            _QtdeEntregueFisicoAFixar = value
        End Set
    End Property
    Public Property QtdeEntregueFisicoDeposito As Decimal
        Get
            Return _QtdeEntregueFisicoDeposito
        End Get
        Set(value As Decimal)
            _QtdeEntregueFisicoDeposito = value
        End Set
    End Property
    Public Property QtdeEntregueFisicoDireta As Decimal
        Get
            Return _QtdeEntregueFisicoDireta
        End Get
        Set(value As Decimal)
            _QtdeEntregueFisicoDireta = value
        End Set
    End Property

    '***********************************************************************************************
    '***********  ENTREGUE VALOR  ******************************************************************
    '***********************************************************************************************
    Public Property VlrNotaOficialGlobalBruto As Decimal
        Get
            Return _VlrNotaOficialGlobalBruto
        End Get
        Set(value As Decimal)
            _VlrNotaOficialGlobalBruto = value
        End Set
    End Property
    Public Property VlrNotaOficialRemessaBruto As Decimal
        Get
            Return _VlrNotaOficialRemessaBruto
        End Get
        Set(value As Decimal)
            _VlrNotaOficialRemessaBruto = value
        End Set
    End Property
    Public Property VlrNotaOficialAFixarBruto As Decimal
        Get
            Return _VlrNotaOficialAFixarBruto
        End Get
        Set(value As Decimal)
            _VlrNotaOficialAFixarBruto = value
        End Set
    End Property
    Public Property VlrNotaOficialDepositoBruto As Decimal
        Get
            Return _VlrNotaOficialDepositoBruto
        End Get
        Set(value As Decimal)
            _VlrNotaOficialDepositoBruto = value
        End Set
    End Property
    Public Property VlrNotaOficialDiretaBruto As Decimal
        Get
            Return _VlrNotaOficialDiretaBruto
        End Get
        Set(value As Decimal)
            _VlrNotaOficialDiretaBruto = value
        End Set
    End Property

    Public Property VlrNotaOficialGlobalLiquido As Decimal
        Get
            Return _VlrNotaOficialGlobalLiquido
        End Get
        Set(value As Decimal)
            _VlrNotaOficialGlobalLiquido = value
        End Set
    End Property
    Public Property VlrNotaOficialRemessaLiquido As Decimal
        Get
            Return _VlrNotaOficialRemessaLiquido
        End Get
        Set(value As Decimal)
            _VlrNotaOficialRemessaLiquido = value
        End Set
    End Property
    Public Property VlrNotaOficialAFixarLiquido As Decimal
        Get
            Return _VlrNotaOficialAFixarLiquido
        End Get
        Set(value As Decimal)
            _VlrNotaOficialAFixarLiquido = value
        End Set
    End Property
    Public Property VlrNotaOficialDepositoLiquido As Decimal
        Get
            Return _VlrNotaOficialDepositoLiquido
        End Get
        Set(value As Decimal)
            _VlrNotaOficialDepositoLiquido = value
        End Set
    End Property
    Public Property VlrNotaOficialDiretaLiquido As Decimal
        Get
            Return _VlrNotaOficialDiretaLiquido
        End Get
        Set(value As Decimal)
            _VlrNotaOficialDiretaLiquido = value
        End Set
    End Property

    Public Property VlrNotaMoedaGlobalBruto As Decimal
        Get
            Return _VlrNotaMoedaGlobalBruto
        End Get
        Set(value As Decimal)
            _VlrNotaMoedaGlobalBruto = value
        End Set
    End Property
    Public Property VlrNotaMoedaRemessaBruto As Decimal
        Get
            Return _VlrNotaMoedaRemessaBruto
        End Get
        Set(value As Decimal)
            _VlrNotaMoedaRemessaBruto = value
        End Set
    End Property
    Public Property VlrNotaMoedaAFixarBruto As Decimal
        Get
            Return _VlrNotaMoedaAFixarBruto
        End Get
        Set(value As Decimal)
            _VlrNotaMoedaAFixarBruto = value
        End Set
    End Property
    Public Property VlrNotaMoedaDepositoBruto As Decimal
        Get
            Return _VlrNotaMoedaDepositoBruto
        End Get
        Set(value As Decimal)
            _VlrNotaMoedaDepositoBruto = value
        End Set
    End Property
    Public Property VlrNotaMoedaDiretaBruto As Decimal
        Get
            Return _VlrNotaMoedaDiretaBruto
        End Get
        Set(value As Decimal)
            _VlrNotaMoedaDiretaBruto = value
        End Set
    End Property

    Public Property VlrNotaMoedaGlobalLiquido As Decimal
        Get
            Return _VlrNotaMoedaGlobalLiquido
        End Get
        Set(value As Decimal)
            _VlrNotaMoedaGlobalLiquido = value
        End Set
    End Property
    Public Property VlrNotaMoedaRemessaLiquido As Decimal
        Get
            Return _VlrNotaMoedaRemessaLiquido
        End Get
        Set(value As Decimal)
            _VlrNotaMoedaRemessaLiquido = value
        End Set
    End Property
    Public Property VlrNotaMoedaAFixarLiquido As Decimal
        Get
            Return _VlrNotaMoedaAFixarLiquido
        End Get
        Set(value As Decimal)
            _VlrNotaMoedaAFixarLiquido = value
        End Set
    End Property
    Public Property VlrNotaMoedaDepositoLiquido As Decimal
        Get
            Return _VlrNotaMoedaDepositoLiquido
        End Get
        Set(value As Decimal)
            _VlrNotaMoedaDepositoLiquido = value
        End Set
    End Property
    Public Property VlrNotaMoedaDiretaLiquido As Decimal
        Get
            Return _VlrNotaMoedaDiretaLiquido
        End Get
        Set(value As Decimal)
            _VlrNotaMoedaDiretaLiquido = value
        End Set
    End Property

    '***********************************************************************************************
    '***********  SALDO  ***************************************************************************
    '***********************************************************************************************
    Public Property SaldoQtdeGlobalFiscal As Decimal
        Get
            Return _SaldoQtdeGlobalFiscal
        End Get
        Set(value As Decimal)
            _SaldoQtdeGlobalFiscal = value
        End Set
    End Property
    Public Property SaldoQtdeRemessaFiscal As Decimal
        Get
            Return _SaldoQtdeRemessaFiscal
        End Get
        Set(value As Decimal)
            _SaldoQtdeRemessaFiscal = value
        End Set
    End Property
    Public Property SaldoQtdeRemessaFisica As Decimal
        Get
            Return _SaldoQtdeRemessaFisica
        End Get
        Set(value As Decimal)
            _SaldoQtdeRemessaFisica = value
        End Set
    End Property

    Public Property SaldoQtdeDiretoFiscal As Decimal
        Get
            Return _SaldoQtdeDiretoFiscal
        End Get
        Set(value As Decimal)
            _SaldoQtdeDiretoFiscal = value
        End Set
    End Property

    Public Property SaldoQtdeDiretoFisica As Decimal
        Get
            Return _SaldoQtdeDiretoFisica
        End Get
        Set(value As Decimal)
            _SaldoQtdeDiretoFisica = value
        End Set
    End Property

    Public Property SaldoQtdeComercializacao As Decimal
        Get
            Return _SaldoQtdeComercializacao
        End Get
        Set(value As Decimal)
            _SaldoQtdeComercializacao = value
        End Set
    End Property
    Public Property SaldoValorOficialGlobalDireto As Decimal
        Get
            Return _SaldoValorOficialGlobalDireto
        End Get
        Set(value As Decimal)
            _SaldoValorOficialGlobalDireto = value
        End Set
    End Property
    Public Property SaldoValorMoedaGlobalDireto As Decimal
        Get
            Return _SaldoValorMoedaGlobalDireto
        End Get
        Set(value As Decimal)
            _SaldoValorMoedaGlobalDireto = value
        End Set
    End Property
    Public Property SaldoValorOficialRemessa As Decimal
        Get
            Return _SaldoValorOficialRemessa
        End Get
        Set(value As Decimal)
            _SaldoValorOficialRemessa = value
        End Set
    End Property
    Public Property SaldoValorMoedaRemessa As Decimal
        Get
            Return _SaldoValorMoedaRemessa
        End Get
        Set(value As Decimal)
            _SaldoValorMoedaRemessa = value
        End Set
    End Property


    'Public Property SaldoValorOficialGlobalDiretoFilial As Decimal
    '    Get
    '        Return _SaldoValorOficialGlobalDiretoFilial
    '    End Get
    '    Set(value As Decimal)
    '        _SaldoValorOficialGlobalDiretoFilial = value
    '    End Set
    'End Property

    'Public Property SaldoValorOficialRemessaFilial As Decimal
    '    Get
    '        Return _SaldoValorOficialRemessaFilial
    '    End Get
    '    Set(value As Decimal)
    '        _SaldoValorOficialRemessaFilial = value
    '    End Set
    'End Property



    Public Property SaldoQtdeAFixar As Decimal
        Get
            Return _SaldoQtdeAFixar
        End Get
        Set(value As Decimal)
            _SaldoQtdeAFixar = value
        End Set
    End Property

    Public ReadOnly Property XmlProdutoXDePara() As XmlProdutoXDePara
        Get
            If CodigoProduto > 0 Then _XmlProdutoXDePara = New XmlProdutoXDePara(CodigoCliente, EndCliente, "", Me.CodigoProduto)
            Return _XmlProdutoXDePara
        End Get
    End Property

    Private _XmlCFOP As Integer
    Public Property XmlCFOP() As Integer
        Get
            Return _XmlCFOP
        End Get
        Set(ByVal value As Integer)
            _XmlCFOP = value
        End Set
    End Property


    Private _XmluCom As String
    Public Property XmluCom() As String
        Get
            Return _XmluCom
        End Get
        Set(ByVal value As String)
            _XmluCom = value
        End Set
    End Property


    Private _XmlqCom As Decimal
    Public Property XmlqCom() As Decimal
        Get
            Return _XmlqCom
        End Get
        Set(ByVal value As Decimal)
            _XmlqCom = value
        End Set
    End Property


    Private _XmlvUnCom As Decimal
    Public Property XmlvUnCom() As Decimal
        Get
            Return _XmlvUnCom
        End Get
        Set(ByVal value As Decimal)
            _XmlvUnCom = value
        End Set
    End Property

    Private _XmlvProd As Decimal
    Public Property XmlvProd() As Decimal
        Get
            Return _XmlvProd
        End Get
        Set(ByVal value As Decimal)
            _XmlvProd = value
        End Set
    End Property

    Private _XmluTrib As String
    Public Property XmluTrib() As String
        Get
            Return _XmluTrib
        End Get
        Set(ByVal value As String)
            _XmluTrib = value
        End Set
    End Property

    Private _XmlqTrib As Decimal
    Public Property XmlqTrib() As Decimal
        Get
            Return _XmlqTrib
        End Get
        Set(ByVal value As Decimal)
            _XmlqTrib = value
        End Set
    End Property

    Private _XmlvUnTrib As Decimal
    Public Property XmlvUnTrib() As Decimal
        Get
            Return _XmlvUnTrib
        End Get
        Set(ByVal value As Decimal)
            _XmlvUnTrib = value
        End Set
    End Property

    Private _XmlinfAdProd As String
    Public Property XmlinfAdProd() As String
        Get
            Return _XmlinfAdProd
        End Get
        Set(ByVal value As String)
            _XmlinfAdProd = value
        End Set
    End Property


    '***********************************************************************************************
    '***********  LISTA DE SALDO POR PRODUTO  ******************************************************
    '***********************************************************************************************
    Public Property Itens As ListSaldoPedido2015
        Get
            If _Itens Is Nothing Then
                Dim Par As New Hashtable
                Par.Add("TipoApuracao", 1)
                Par.Add("Empresa", Me.CodigoEmpresa)
                Par.Add("Pedido", Me.CodigoPedido)
                Par.Add("Classe", Me.SubOperacao.Classe.ToString)
                Par.Add("Devolucao", IIf(Me.SubOperacao.Devolucao, 1, 0))
                Par.Add("SoValor", IIf(Me.SubOperacao.QuantidadeFiscal = False And Me.SubOperacao.QuantidadeFisico = False, 1, 0))

                _Itens = New ListSaldoPedido2015(Par)
            End If
            Return _Itens
        End Get
        Set(value As ListSaldoPedido2015)
            _Itens = value
        End Set
    End Property

    Public Property Contrato As String
        Get
            Return _Contrato
        End Get
        Set(value As String)
            _Contrato = value
        End Set
    End Property

#End Region

#Region "Methods"
    Public Sub RecarregarItens(pClasse As String, pDevolucao As Boolean, pSoValor As Boolean, Optional ByVal pCodigoProduto As Integer = 0, Optional ByVal pDataReferencia As String = "", Optional ByVal pCodEmpresaDev As String = "", Optional ByVal pEndEmpresaDev As Integer = 0)
        Dim Par As New Hashtable
        Par.Add("TipoApuracao", 1)
        Par.Add("Empresa", Me.CodigoEmpresa)
        Par.Add("EndEmpresa", Me.EndEmpresa)
        Par.Add("Pedido", Me.CodigoPedido)

        Par.Add("Classe", pClasse)
        Par.Add("Devolucao", IIf(pDevolucao, 1, 0))
        Par.Add("SoValor", IIf(pSoValor, 1, 0))

        If pCodigoProduto > 0 Then Par.Add("Produto", pCodigoProduto)

        If pDataReferencia <> "" Then
            Par.Add("DataReferencia", pDataReferencia)
        Else
            Par.Add("DataReferencia", Me.DataReferencia)
        End If

        If pCodEmpresaDev <> "" Then
            Par.Add("FilialDev", pCodEmpresaDev)
            Par.Add("EndFilialDev", pEndEmpresaDev)
        End If

        _Itens = New ListSaldoPedido2015(Par)
    End Sub

    Public Sub PreencherXmlProdutoXDePara(ByVal IndiceDoProduto As Integer, ByVal ProdutoDe As String)
        Try

            Dim ProdutoPara As String = Me.Itens(IndiceDoProduto).CodigoProduto
            'Dim Sqls As New ArrayList
            Dim objProd As New XmlProdutoXDePara(Me.CodigoCliente, Me.EndCliente, ProdutoPara, ProdutoDe)
            objProd.IUD = "I"
            objProd.CodigoCliente = Me.CodigoCliente
            objProd.EndCliente = Me.EndCliente
            objProd.CodigoProdutoXML = ProdutoDe
            objProd.CodigoProduto = ProdutoPara
            objProd.Salvar()
            'If Not Banco.GravaBanco(Sqls) Then
            '    MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
            'End If
            Me.Itens(IndiceDoProduto).XmlProdutoXDePara.CodigoProdutoXML = ProdutoDe

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

#End Region

End Class
