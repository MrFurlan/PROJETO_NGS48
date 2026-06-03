Imports System.Data.SqlClient
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Threading.Tasks
Imports Newtonsoft.Json
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

<Serializable()>
Public Class ListWsPedido
    Inherits List(Of WsPedido)
    Public Sub New(ByVal DataInicio As Date, ByVal DataFim As Date, ByVal ClienteId As String, Optional ByVal NumPedido As Integer = 0)
        Dim sql As String =
         "      SELECT [VendedorCod] " & vbCrLf &
         "          ,[PedidoNumPedCli] " & vbCrLf &
         "          ,[PedidoNum] " & vbCrLf &
         "          ,[PedidoData] " & vbCrLf &
         "          ,p.ClienteCod " & vbCrLf &
         "          ,c.Nome AS Cliente " & vbCrLf &
         "          ,p.VendedorCod " & vbCrLf &
         "          ,v.Nome AS Vendedor " & vbCrLf &
         "          ,p.PedidoVrpagar " & vbCrLf &
         "      FROM [dbo].[PedidoIntegracaoOnSoft] p " & vbCrLf &
         "      INNER JOIN Clientes c " & vbCrLf &
         "        ON p.ClienteCod = c.Cliente_Id " & vbCrLf &
         "        --AND ClienteEnderecoId = c.Endereco_Id " & vbCrLf &
         "      INNER JOIN Clientes v " & vbCrLf &
         "        ON p.VendedorCod = v.Cliente_Id " & vbCrLf &
         "     WHERE PedidoData BETWEEN '" & CDate(DataInicio).ToString("yyyy/MM/dd") & "' AND '" & CDate(DataFim).ToString("yyyy/MM/dd") & "'" & vbCrLf

        If Not String.IsNullOrEmpty(ClienteId) Then
            sql += " And p.ClienteCod = '" & ClienteId & "' "
        End If

        If NumPedido > 0 Then
            sql += " And p.PedidoNumPedCli = " & NumPedido
        End If

        Dim Banco As New AcessaBancoOnMobile
        Dim ds As DataSet

        ds = Banco.ConsultaDataSet(sql, "WsPedido")

        For Each row As DataRow In ds.Tables(0).Rows
            Dim ws As New WsPedido

            ws.PedidoNum = row("PedidoNum")
            ws.PedidoData = row("PedidoData")
            ws.PedidoVrpagar = row("PedidoVrpagar")
            ws.VendedorCod = row("VendedorCod")
            ws.Vendedor = row("Vendedor")
            ws.ClienteCod = row("ClienteCod")
            ws.Cliente = row("Cliente")
            ws.PedidoNumPedCli = row("PedidoNumPedCli")

            Me.Add(ws)
        Next
    End Sub

End Class

<Serializable()>
Public Class WsPedido

#Region "Property"
    <JsonProperty("vendedorCod")>
    Public Property VendedorCod As String

    <JsonProperty("pedidoNum")>
    Public Property PedidoNum As String

    <JsonProperty("tipooperacaoCod")>
    Public Property TipoOperacaoCod As String

    <JsonProperty("pedidoNumpedcli")>
    Public Property PedidoNumPedCli As String

    <JsonProperty("clienteCod")>
    Public Property ClienteCod As String

    <JsonProperty("tabprecoCod")>
    Public Property TabPrecoCod As String

    <JsonProperty("tabcondicaoCod")>
    Public Property TabCondicaoCod As String

    <JsonProperty("condicaoCod")>
    Public Property CondicaoCod As String

    <JsonProperty("condicaoFator")>
    Public Property CondicaoFator As String

    <JsonProperty("condicaoRnMinUe")>
    Public Property CondicaoRnMinUe As String

    <JsonProperty("condicaoRnMinVr")>
    Public Property CondicaoRnMinVr As String

    <JsonProperty("condicaoRnMinModo")>
    Public Property CondicaoRnMinModo As String

    <JsonProperty("tabdescontoCod")>
    Public Property TabdescontoCod As String

    <JsonProperty("tabdescontoMinUe")>
    Public Property TabDescontoMinUe As String

    <JsonProperty("tabdescontoMinVr")>
    Public Property TabDescontoMinVr As String

    <JsonProperty("tabdescontoDesconto")>
    Public Property TabDescontoDesconto As String

    <JsonProperty("tabdescontoTipo")>
    Public Property TabdescontoTipo As String

    <JsonProperty("tabdescontoDistribuido")>
    Public Property TabdescontoDistribuido As String

    <JsonProperty("enderecoCobranca")>
    Public Property EnderecoCobranca As String

    <JsonProperty("enderecoEntrega")>
    Public Property EnderecoEntrega As String

    <JsonProperty("pedidoData")>
    Public Property PedidoData As String

    <JsonProperty("pedidoDataEntrega")>
    Public Property PedidoDataEntrega As String

    <JsonProperty("pedidoUrgente")>
    Public Property PedidoUrgente As String

    <JsonProperty("pedidoTipofrete")>
    Public Property PedidoTipofrete As String

    <JsonProperty("pedidoTotalUe")>
    Public Property PedidoTotalUe As String

    <JsonProperty("pedidoVrminimo")>
    Public Property PedidoVrminimo As String

    <JsonProperty("pedidoVrbase")>
    Public Property PedidoVrbase As String

    <JsonProperty("pedidoVrbase2")>
    Public Property PedidoVrbase2 As String

    <JsonProperty("pedidoVritens")>
    Public Property PedidoVritens As String

    <JsonProperty("pedidoVritens2")>
    Public Property PedidoVritens2 As String

    <JsonProperty("pedidoTipodesconto")>
    Public Property PedidoTipodesconto As String

    <JsonProperty("pedidoDesconto")>
    Public Property PedidoDesconto As String

    <JsonProperty("pedidoVrpagar")>
    Public Property PedidoVrpagar As String

    <JsonProperty("pedidoVrpagar2")>
    Public Property PedidoVrpagar2 As String

    <JsonProperty("pedidoEnviaEmail")>
    Public Property PedidoEnviaEmail As String

    <JsonProperty("pedidoEmail")>
    Public Property PedidoEmail As String

    <JsonProperty("pedidoObs")>
    Public Property PedidoObs As String

    <JsonProperty("pedidoRinicio")>
    Public Property PedidoRinicio As String

    <JsonProperty("pedidoRtermino")>
    Public Property PedidoRtermino As String

    <JsonProperty("pedidoDtenvio")>
    Public Property PedidoDtenvio As String

    <JsonProperty("pedidoDtexportacao")>
    Public Property PedidoDtexportacao As String

    <JsonProperty("pedidoDtimpressao")>
    Public Property PedidoDtimpressao As String

    <JsonProperty("clienteMargemPedido")>
    Public Property ClienteMargemPedido As String

    <JsonProperty("tipooperacaoCusto")>
    Public Property TipoOperacaoCusto As String

    <JsonProperty("tipooperacaoVerba")>
    Public Property TipoOperacaoVerba As String

    <JsonProperty("pedidoTotalCustoMin")>
    Public Property PedidoTotalCustoMin As String

    <JsonProperty("pedidoTotalCusto")>
    Public Property PedidoTotalCusto As String

    <JsonProperty("pedidoTotalPackCustoMin")>
    Public Property PedidoTotalPackCustoMin As String

    <JsonProperty("pedidoTotalPackCusto")>
    Public Property PedidoTotalPackCusto As String

    <JsonProperty("pedidoTotalVerba")>
    Public Property PedidoTotalVerba As String

    <JsonProperty("pedidoitems")>
    Public Property PedidoItems As List(Of WsPedidoItem)

    <JsonProperty("pedidoObss")>
    Public Property PedidoObss As List(Of WsPedidoObs)

    <JsonProperty("pedidoBonificacoes")>
    Public Property PedidoBonificacoes As List(Of WsPedidoBonificacao)

    Public Property Vendedor As String
    Public Property Cliente As String
#End Region

#Region "Methods"
    Public Sub SalvarSql(ByRef Sqls As ArrayList)

        Dim strSql As String = "INSERT INTO dbo.PedidoIntegracaoOnSoft" & vbCrLf &
                   "(VendedorCod" & vbCrLf &
                   ",PedidoNum" & vbCrLf &
                   ",TipoOperacaoCod" & vbCrLf &
                   ",PedidoNumPedCli" & vbCrLf &
                   ",ClienteCod" & vbCrLf &
                   ",TabPrecoCod" & vbCrLf &
                   ",TabCondicaoCod" & vbCrLf &
                   ",CondicaoCod" & vbCrLf &
                   ",CondicaoFator" & vbCrLf &
                   ",CondicaoRnMinUe" & vbCrLf &
                   ",CondicaoRnMinVr" & vbCrLf &
                   ",CondicaoRnMinModo" & vbCrLf &
                   ",TabDescontoCod" & vbCrLf &
                   ",TabDescontoMinUe" & vbCrLf &
                   ",TabDescontoMinVr" & vbCrLf &
                   ",TabDescontoDesconto" & vbCrLf &
                   ",TabDescontoTipo" & vbCrLf &
                   ",TabDescontoDistribuido" & vbCrLf &
                   ",EnderecoCobranca" & vbCrLf &
                   ",EnderecoEntrega" & vbCrLf &
                   ",PedidoData" & vbCrLf &
                   ",PedidoDataEntrega" & vbCrLf &
                   ",PedidoUrgente" & vbCrLf &
                   ",PedidoTipofrete" & vbCrLf &
                   ",PedidoTotalUe" & vbCrLf &
                   ",PedidoVrminimo" & vbCrLf &
                   ",PedidoVrbase" & vbCrLf &
                   ",PedidoVrbase2" & vbCrLf &
                   ",PedidoVritens" & vbCrLf &
                   ",PedidoVritens2" & vbCrLf &
                   ",PedidoTipodesconto" & vbCrLf &
                   ",PedidoDesconto" & vbCrLf &
                   ",PedidoVrpagar" & vbCrLf &
                   ",PedidoVrpagar2" & vbCrLf &
                   ",PedidoEnviaEmail" & vbCrLf &
                   ",PedidoEmail" & vbCrLf &
                   ",PedidoObs" & vbCrLf &
                   ",PedidoRinicio" & vbCrLf &
                   ",PedidoRtermino" & vbCrLf &
                   ",PedidoDtenvio" & vbCrLf &
                   ",ClienteMargemPedido" & vbCrLf &
                   ",TipooperacaoCusto" & vbCrLf &
                   ",TipooperacaoVerba" & vbCrLf &
                   ",PedidoTotalCustoMin" & vbCrLf &
                   ",PedidoTotalCusto" & vbCrLf &
                   ",PedidoTotalPackCustoMin" & vbCrLf &
                   ",PedidoTotalPackCusto" & vbCrLf &
                   ",PedidoTotalVerba)" & vbCrLf &
                 "VALUES" & vbCrLf &
                    "('" & Me.VendedorCod & "'" & vbCrLf &
                    ",'" & Me.PedidoNum & "'" & vbCrLf &
                    ",'" & Me.TipoOperacaoCod & "'" & vbCrLf &
                    ",'" & Me.PedidoNumPedCli & "'" & vbCrLf &
                    ",'" & Me.ClienteCod & "'" & vbCrLf &
                    ",'" & Me.TabPrecoCod & "'" & vbCrLf &
                    ",'" & Me.TabCondicaoCod & "'" & vbCrLf &
                    ",'" & Me.CondicaoCod & "'" & vbCrLf &
                    "," & Me.CondicaoFator & vbCrLf &
                    "," & Me.CondicaoRnMinUe & vbCrLf &
                    "," & Me.CondicaoRnMinVr & vbCrLf &
                    ",'" & Me.CondicaoRnMinModo & "'" & vbCrLf &
                    ",'" & Me.TabdescontoCod & "'" & vbCrLf &
                    "," & Me.TabDescontoMinUe & vbCrLf &
                    "," & Me.TabDescontoMinVr & vbCrLf &
                    "," & Me.TabDescontoDesconto & vbCrLf &
                    ",'" & Me.TabdescontoTipo & "'" & vbCrLf &
                    "," & Me.TabdescontoDistribuido & vbCrLf &
                    ",'" & Me.EnderecoCobranca & "'" & vbCrLf &
                    ",'" & Me.EnderecoEntrega & "'" & vbCrLf &
                    ",'" & Me.PedidoData & "'" & vbCrLf &
                    ",'" & Me.PedidoDataEntrega & "'" & vbCrLf &
                    "," & Me.PedidoUrgente & vbCrLf &
                    ",'" & Me.PedidoTipofrete & "'" & vbCrLf &
                    "," & Me.PedidoTotalUe & vbCrLf &
                    "," & Me.PedidoVrminimo & vbCrLf &
                    "," & Me.PedidoVrbase & vbCrLf &
                    "," & Me.PedidoVrbase2 & vbCrLf &
                    "," & Me.PedidoVritens & vbCrLf &
                    "," & Me.PedidoVritens2 & vbCrLf &
                    ",'" & Me.PedidoTipodesconto & "'" & vbCrLf &
                    "," & Me.PedidoDesconto & vbCrLf &
                    "," & Me.PedidoVrpagar & vbCrLf &
                    "," & Me.PedidoVrpagar2 & vbCrLf &
                    "," & Me.PedidoEnviaEmail & vbCrLf &
                    ",'" & Me.PedidoEmail & "'" & vbCrLf &
                    ",'" & Me.PedidoObs & "'" & vbCrLf &
                    ",'" & Me.PedidoRinicio & "'" & vbCrLf &
                    ",'" & Me.PedidoRtermino & "'" & vbCrLf &
                    ",'" & Me.PedidoDtenvio & "'" & vbCrLf &
                    "," & Me.ClienteMargemPedido & vbCrLf &
                    "," & Me.TipoOperacaoCusto & vbCrLf &
                    "," & Me.TipoOperacaoVerba & vbCrLf &
                    "," & Me.PedidoTotalCustoMin & vbCrLf &
                    "," & Me.PedidoTotalCusto & vbCrLf &
                    "," & Me.PedidoTotalPackCustoMin & vbCrLf &
                    "," & Me.PedidoTotalPackCusto & vbCrLf &
                    "," & Me.PedidoTotalVerba & ");"

        Sqls.Add(strSql)

        strSql = "DELETE PedidoNAOIntegracaoOnSoft " &
                 " WHERE VendedorCod = '" & Me.VendedorCod & "'" & vbCrLf &
                 "   AND PedidoNum   = '" & Me.PedidoNum & "'" & vbCrLf &
                 "   AND ClienteCod  = '" & Me.ClienteCod & "';"

        Sqls.Add(strSql)

    End Sub

    Public Function VerOrigemEmpresa() As Cliente

        Dim nCliente As New Cliente()

        Dim Banco As New AcessaBancoOnMobile

        Dim cmd As New SqlCommand With
        {
            .CommandText = String.Concat("
                SELECT top 1 pXp.Cliente_Id, pXp.EndCliente_Id As Endereco_Id, gxe.Empresa_Id As UnidadeDeNegocio
                FROM ProdutosXPrecos pXp
                INNER JOIN GruposXEmpresas gxe
                    ON gxe.Cliente_Id    = pXp.Cliente_Id
                   AND gxe.EndCliente_Id = pXp.EndCliente_Id
                WHERE pXp.Tabela_id = @Tabela_id")}

        cmd.Parameters.Add("@Tabela_id", SqlDbType.VarChar)
        cmd.Parameters("@Tabela_id").Value = Me.TabPrecoCod

        Dim ds = Banco.ConsultaDataSet(DataSQLExtensao.CommandAsSqlText(cmd), "Empresa")

        nCliente = New Cliente(ds.Tables(0).Rows(0).Item("Cliente_Id"), ds.Tables(0).Rows(0).Item("Endereco_Id"))

        Return nCliente

    End Function

    Public Function PreenchePedidoNGS() As Pedido
        Dim pedidoNGS As New Pedido

        Dim Banco As New AcessaBancoOnMobile

        Dim cmd As New SqlCommand With
        {
            .CommandText = String.Concat("
                SELECT top 1 pXp.Cliente_Id, pXp.EndCliente_Id As Endereco_Id, gxe.Empresa_Id As UnidadeDeNegocio
                FROM ProdutosXPrecos pXp
                INNER JOIN GruposXEmpresas gxe
                    ON gxe.Cliente_Id    = pXp.Cliente_Id
                   AND gxe.EndCliente_Id = pXp.EndCliente_Id
                WHERE pXp.Tabela_id = @Tabela_id")}

        cmd.Parameters.Add("@Tabela_id", SqlDbType.VarChar)
        cmd.Parameters("@Tabela_id").Value = Me.TabPrecoCod

        Dim ds = Banco.ConsultaDataSet(DataSQLExtensao.CommandAsSqlText(cmd), "Empresa")

        pedidoNGS.IUD = "I"
        pedidoNGS.CodigoUnidadeNegocio = ds.Tables(0).Rows(0).Item("UnidadeDeNegocio")
        pedidoNGS.CodigoEmpresa = ds.Tables(0).Rows(0).Item("Cliente_Id")
        pedidoNGS.EnderecoEmpresa = ds.Tables(0).Rows(0).Item("Endereco_Id")

        pedidoNGS.CodigoCliente = Me.ClienteCod
        pedidoNGS.EnderecoCliente = 0
        pedidoNGS.CodigoPraca = Me.ClienteCod
        pedidoNGS.EnderecoPraca = 0
        pedidoNGS.CodigoSafra = "NENHUMA"
        pedidoNGS.CodigoMoeda = 1
        pedidoNGS.TemVariacao = False
        pedidoNGS.CodigoIndexador = 99
        pedidoNGS.IndiceFixado = 0
        pedidoNGS.MomentoFinanceiro = 3

        Dim op = Me.TipoOperacaoCod.Split(".").ToArray()
        pedidoNGS.CodigoOperacao = Integer.Parse(op(0))
        pedidoNGS.CodigoSubOperacao = Integer.Parse(op(1))

        pedidoNGS.CodigoSituacao = 1
        pedidoNGS.PedidoBloqueado = True
        pedidoNGS.DataPedido = Me.PedidoData

        pedidoNGS.DataEntregaInicial = Me.PedidoDataEntrega
        pedidoNGS.DataEntregaFinal = Me.PedidoDataEntrega

        If Me.PedidoTipofrete = "C" Then
            pedidoNGS.FreteCIFFOB = eTiposFrete.CIF
        ElseIf Me.PedidoTipofrete = "F" Then
            pedidoNGS.FreteCIFFOB = eTiposFrete.FOB
        Else
            pedidoNGS.FreteCIFFOB = eTiposFrete.NEN
        End If

        pedidoNGS.OrigemDestino = "0"
        pedidoNGS.CodigoFinalidade = 2
        pedidoNGS.DataVencimentoPedido = Me.PedidoData
        pedidoNGS.UsuarioInclusao = "INTEGRAÇÃO ONSOFT"
        pedidoNGS.DataInclusao = DateTime.Now()

        'PREENCHE OBSERVAÇÃO NO PEDIDO CASO VENDENDOR TENHA INFORMADO
        If Not String.IsNullOrWhiteSpace(Me.PedidoObs) Then
            pedidoNGS.Observacoes = Me.PedidoObs
        End If

        'PREENCHE OS ITENS DO PEDIDO
        For Each Item In Me.PedidoItems
            Dim Pitem As New PedidoXItem(pedidoNGS)
            Pitem.IUD = "I"

            Dim produto As New Produto(Item.ProdutoCod)
            Pitem.CodigoProduto = produto.Codigo
            Pitem.CodigoClassificacao = 1
            Pitem.Descricao = produto.Nome
            Pitem.Retencao = False
            Pitem.CodigoUnidadeComercializacao = produto.Unidade

            'PREENCHE O LANCAMENTO NORMAL
            Dim Lan As New LancamentoItemPedido(Pitem)
            Lan.TipoLancamento = eTiposLancamentosPedidos.Normal
            Lan.CodigoPedidoItem = 0
            Lan.Movimento = Me.PedidoData
            Lan.DataEntrega = Me.PedidoDataEntrega
            Lan.QuantidadeFaturamento = Item.PedidoItemQtde

            Dim _grupoMaisDe10Kg = produto.UnidadesDeComercializacao.First().FatorConversao > 10

            Dim temIPI As Boolean = False

            If Mid(Pitem.CodigoProduto, 5, 1).Contains("1") OrElse
                Mid(Pitem.CodigoProduto, 5, 1).Contains("3") OrElse
                Mid(Pitem.CodigoProduto, 5, 1).Contains("5") OrElse
                Mid(Pitem.CodigoProduto, 5, 1).Contains("7") OrElse
                Mid(Pitem.CodigoProduto, 5, 1).Contains("9") Then
                temIPI = True
            End If

            If Left(pedidoNGS.CodigoEmpresa, 8) = "49673784" Then 'DISTRIBUIDORA NÃO TEM IPI - FURLAN 13/05/2024
                temIPI = False
            End If

            Dim impostoReverso As New CalculoImpostosOnSoft(pedidoNGS.Cliente.CodigoEstado, _grupoMaisDe10Kg, temIPI)
            Dim indiceImpostoReverso = IIf(impostoReverso.Indice > 0, impostoReverso.Indice, 1)

            If Left(Pitem.CodigoProduto, 3) = "801" Then
                indiceImpostoReverso = 1
            ElseIf pedidoNGS.CodigoOperacao = 21 AndAlso pedidoNGS.CodigoSubOperacao = 82 Then
                If temIPI Then
                    indiceImpostoReverso = 1.065
                Else
                    indiceImpostoReverso = 1
                End If
            End If

            If Left(pedidoNGS.CodigoEmpresa, 8) = "49673784" Then 'DISTRIBUIDORA NÃO TEM ICMS-ST - FURLAN 13/05/2024
                indiceImpostoReverso = 1
            End If

            Dim vlrUnitario As Decimal = CDec(Item.PedidoItemVrunit.Replace(".", ","))
            Dim valorUnitarioFinal As Decimal = (vlrUnitario / indiceImpostoReverso)

            'Lan.UnitarioOficial = Item.PedidoItemVrunit
            Lan.UnitarioOficial = valorUnitarioFinal

            Lan.TotalOficial = (valorUnitarioFinal * Item.PedidoItemQtde)
            Lan.QuantidadeComercializacao = Item.PedidoItemQtde

            Pitem.Lancamentos.Add(Lan)
            Pitem.Encargos = Nothing
            Pitem.Encargos.CriaListar()

            pedidoNGS.Itens.Add(Pitem)
        Next

        'PREENCHE DEPOSITO
        Dim DepositoES As New PedidoXDeposito(pedidoNGS)
        DepositoES.Codigo = pedidoNGS.CodigoEmpresa
        DepositoES.CodigoEndereco = pedidoNGS.EnderecoEmpresa
        DepositoES.Principal = True
        DepositoES.Quantidade = 0
        DepositoES.Tipo = "DE"
        pedidoNGS.Depositos.Add(DepositoES)

        'PREENCHE ORIGEM DESTINO
        Dim DepositoOD As New PedidoXDeposito(pedidoNGS)
        DepositoOD.Codigo = pedidoNGS.CodigoCliente
        DepositoOD.CodigoEndereco = pedidoNGS.EnderecoCliente
        DepositoOD.Principal = True
        DepositoOD.Quantidade = 0
        DepositoOD.Tipo = "OD"
        pedidoNGS.Depositos.Add(DepositoOD)

        'Comissões
        Dim comissao = New PedidoXRepresentante(pedidoNGS)
        comissao.IUD = "I"
        comissao.CodigoRepresentante = Me.VendedorCod
        comissao.CodigoEnderecoRepresentante = 0
        comissao.Principal = True
        comissao.Percentual = 1
        comissao.PercentualFixo = True
        comissao.ValorComissao = Math.Round((Decimal.Parse(Me.PedidoVrpagar) * comissao.Percentual) / 100, 2)

        pedidoNGS.Representantes = New ListPedidoXRepresentante(pedidoNGS)
        pedidoNGS.Representantes.Add(comissao)

        Dim condicaoPgto = Me.CondicaoCod.Split("|").ToArray()
        pedidoNGS.CodigoCondicaoPagamento = condicaoPgto(1)

        pedidoNGS.Vencimentos.CriarParcelamentoOnMobile(pedidoNGS.Itens.LiquidoOficial)

        Return pedidoNGS
    End Function

    Public Function Copy() As WsPedido
        Return DirectCast(Me.MemberwiseClone(), WsPedido)
    End Function



#End Region
End Class
