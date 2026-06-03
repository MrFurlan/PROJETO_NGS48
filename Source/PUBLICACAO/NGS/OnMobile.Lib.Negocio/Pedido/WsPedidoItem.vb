Imports System.Data.SqlClient
Imports System.Threading.Tasks
Imports Newtonsoft.Json
Imports NGS.Lib.Uteis

<Serializable()>
Public Class ListWsPedidoItem
    Inherits List(Of WsPedidoItem)
    Public Sub New(Optional ByVal CarregarDados As Boolean = False)
        Dim sql As String = "Select * from Clientes"
        Dim Banco As New AcessaBancoOnMobile
        Dim ds As DataSet

        ds = Banco.ConsultaDataSet(sql, "Contrato")

        For Each row As DataRow In ds.Tables(0).Rows
            Dim ws As New WsPedidoItem

            Me.Add(ws)
        Next
    End Sub

    Public Function Salvar(ByVal sqls As ArrayList) As Boolean
        For Each p In Me
            p.SalvarSql(sqls)
        Next

        Return True
    End Function
End Class

<Serializable()>
Public Class WsPedidoItem

#Region "Property"

    <JsonProperty("vendedorCod")>
    Public Property VendedorCod As String

    <JsonProperty("pedidoNum")>
    Public Property PedidoNum As String

    <JsonProperty("produtoCod")>
    Public Property ProdutoCod As String

    <JsonProperty("produtoStatus")>
    Public Property ProdutoStatus As String

    <JsonProperty("produtoUe")>
    Public Property ProdutoUe As String

    <JsonProperty("tabprecoFrete")>
    Public Property TabprecoFrete As String

    <JsonProperty("pedidoitemNumitem")>
    Public Property PedidoItemNumitem As String

    <JsonProperty("pedidoitemQtde")>
    Public Property PedidoItemQtde As String

    <JsonProperty("pedidoitemVrunitmin")>
    Public Property PedidoItemVrunitmin As String

    <JsonProperty("pedidoitemVrunitmax")>
    Public Property PedidoItemVrunitmax As String

    <JsonProperty("pedidoitemVrunitbase")>
    Public Property PedidoItemVrunitbase As String

    <JsonProperty("pedidoitemVrunitbase2")>
    Public Property PedidoItemVrunitbase2 As String

    <JsonProperty("pedidoitemVrunitOri")>
    Public Property PedidoItemVrunitOri As String

    <JsonProperty("pedidoitemVrunitOri2")>
    Public Property PedidoItemVrunitOri2 As String

    <JsonProperty("pedidoitemVrunit")>
    Public Property PedidoItemVrunit As String

    <JsonProperty("pedidoitemVrunit2")>
    Public Property PedidoItemVrunit2 As String

    <JsonProperty("pedidoitemDesconto")>
    Public Property PedidoItemDesconto As String

    <JsonProperty("produtoPackUe")>
    Public Property ProdutoPackUe As String

    <JsonProperty("produtoPackQtde")>
    Public Property ProdutoPackQtde As String

    <JsonProperty("tabprecoPackFrete")>
    Public Property TabprecoPackFrete As String

    <JsonProperty("pedidoitemPackQtde")>
    Public Property PedidoItemPackQtde As String

    <JsonProperty("pedidoitemPackVrunitmin")>
    Public Property PedidoItemPackVrunitmin As String

    <JsonProperty("pedidoitemPackVrunitmax")>
    Public Property PedidoItemPackVrunitmax As String

    <JsonProperty("pedidoitemPackVrunitbase")>
    Public Property PedidoItemPackVrunitbase As String

    <JsonProperty("pedidoitemPackVrunitbase2")>
    Public Property PedidoItemPackVrunitbase2 As String

    <JsonProperty("pedidoitemPackVrunitOri")>
    Public Property PedidoItemPackVrunitOri As String

    <JsonProperty("pedidoitemPackVrunitOri2")>
    Public Property PedidoItemPackVrunitOri2 As String

    <JsonProperty("pedidoitemPackVrunit")>
    Public Property PedidoItemPackVrunit As String

    <JsonProperty("pedidoitemPackVrunit2")>
    Public Property PedidoItemPackVrunit2 As String

    <JsonProperty("pedidoitemSubtotalVr")>
    Public Property PedidoItemSubtotalVr As String

    <JsonProperty("pedidoitemSubtotalVr2")>
    Public Property PedidoItemSubtotalVr2 As String

    <JsonProperty("tabprecoVrminCusto")>
    Public Property TabprecoVrminCusto As String

    <JsonProperty("tabprecoVrCusto")>
    Public Property TabprecoVrCusto As String

    <JsonProperty("tabprecoPackVrminCusto")>
    Public Property TabprecoPackVrminCusto As String

    <JsonProperty("tabprecoPackVrCusto")>
    Public Property TabprecoPackVrCusto As String

    <JsonProperty("tabprecoCustoData")>
    Public Property TabprecoCustoData As String

    <JsonProperty("pedidoitemVrCusto")>
    Public Property PedidoItemVrCusto As String

    <JsonProperty("pedidoitemPackVrCusto")>
    Public Property PedidoItemPackVrCusto As String

    <JsonProperty("tabprecoVrminVerba")>
    Public Property TabprecoVrminVerba As String

    <JsonProperty("tabprecoPackVrminVerba")>
    Public Property TabprecoPackVrminVerba As String

    <JsonProperty("pedidoitemVrunitminVerba")>
    Public Property PedidoItemVrunitminVerba As String

    <JsonProperty("pedidoitemPackVrunitminVerba")>
    Public Property PedidoItemPackVrunitminVerba As String

    <JsonProperty("pedidoitemVerbaUtilizada")>
    Public Property PedidoItemVerbaUtilizada As String

    <JsonProperty("pedidoitemVerbaCp")>
    Public Property PedidoItemVerbaCp As String

    <JsonProperty("pedidoitemDescontoOri")>
    Public Property PedidoItemDescontoOri As String

    <JsonProperty("pedidoitemPackDescontoOri")>
    Public Property PedidoItemPackDescontoOri As String

    <JsonProperty("pedidoitemObs")>
    Public Property PedidoItemObs As String

    <JsonProperty("precificacaoCod")>
    Public Property PrecificacaoCod As String

    <JsonProperty("precificacaoComissao")>
    Public Property PrecificacaoComissao As String

    <JsonProperty("precificacaoAcrescimo")>
    Public Property PrecificacaoAcrescimo As String

    <JsonProperty("precificacaoDesconto")>
    Public Property PrecificacaoDesconto As String

    <JsonProperty("precificacaoBonificacao")>
    Public Property PrecificacaoBonificacao As String

#End Region

#Region "Methods"
    ''Public Sub SalvarSql(ByRef Sqls As ArrayList)

    ''    Dim cmd As New SqlCommand With
    ''    {
    ''        .CommandText = String.Concat("
    ''            INSERT INTO [dbo].[PedidoXItemIntegracaoOnSoft]
    ''               ([VendedorCod]
    ''               ,[PedidoNum]
    ''               ,[ProdutoCod]
    ''               ,[ProdutoStatus]
    ''               ,[ProdutoUe]
    ''               ,[TabPrecoFrete]
    ''               ,[PedidoItemNumitem]
    ''               ,[PedidoItemQtde]
    ''               ,[PedidoItemVrunitmin]
    ''               ,[PedidoItemVrunitmax]
    ''               ,[PedidoItemVrunitbase]
    ''               ,[PedidoItemVrunitbase2]
    ''               ,[PedidoItemVrunitOri]
    ''               ,[PedidoItemVrunitOri2]
    ''               ,[PedidoItemVrunit]
    ''               ,[PedidoItemVrunit2]
    ''               ,[PedidoItemDesconto]
    ''               ,[ProdutoPackUe]
    ''               ,[ProdutoPackQtde]
    ''               ,[TabPrecoPackFrete]
    ''               ,[PedidoItemPackQtde]
    ''               ,[PedidoItemPackVrunitmin]
    ''               ,[PedidoItemPackVrunitmax]
    ''               ,[PedidoItemPackVrunitbase]
    ''               ,[PedidoItemPackVrunitbase2]
    ''               ,[PedidoItemPackVrunitOri]
    ''               ,[PedidoItemPackVrunitOri2]
    ''               ,[PedidoItemPackVrunit]
    ''               ,[PedidoItemPackVrunit2]
    ''               ,[PedidoItemSubtotalVr]
    ''               ,[PedidoItemSubtotalVr2]
    ''               ,[TabPrecoVrminCusto]
    ''               ,[TabPrecoVrCusto]
    ''               ,[TabPrecoPackVrminCusto]
    ''               ,[TabPrecoPackVrCusto]
    ''               ,[PedidoItemVrCusto]
    ''               ,[PedidoItemPackVrCusto]
    ''               ,[TabPrecoVrminVerba]
    ''               ,[TabPrecoPackVrminVerba]
    ''               ,[PPedidoItemVrunitminVerba]
    ''               ,[PedidoItemPackVrunitminVerba]
    ''               ,[PedidoItemVerbaUtilizada]
    ''               ,[PedidoItemVerbaCp]
    ''               ,[PedidoItemDescontoOri]
    ''               ,[PedidoItemPackDescontoOri]
    ''               ,[PedidoItemObs]
    ''               ,[PrecificacaoCod]
    ''               ,[PrecificacaoComissao]
    ''               ,[PrecificacaoAcrescimo]
    ''               ,[PrecificacaoDesconto]
    ''               ,[PrecificacaoBonificacao])
    ''            VALUES
    ''               (@VendedorCod
    ''               ,@PedidoNum
    ''               ,@ProdutoCod
    ''               ,@ProdutoStatus
    ''               ,@ProdutoUe
    ''               ,@TabPrecoFrete
    ''               ,@PedidoItemNumitem
    ''               ,@PedidoItemQtde
    ''               ,@PedidoItemVrunitmin
    ''               ,@PedidoItemVrunitmax
    ''               ,@PedidoItemVrunitbase
    ''               ,@PedidoItemVrunitbase2
    ''               ,@PedidoItemVrunitOri
    ''               ,@PedidoItemVrunitOri2
    ''               ,@PedidoItemVrUnitario
    ''               ,@PedidoItemVrunit2
    ''               ,@PedidoItemDesconto
    ''               ,@ProdutoPackUe
    ''               ,@ProdutoPackQtde
    ''               ,@TabPrecoPackFrete
    ''               ,@PedidoItemPackQtde
    ''               ,@PedidoItemPackVrunitmin
    ''               ,@PedidoItemPackVrunitmax
    ''               ,@PedidoItemPackVrunitbase
    ''               ,@PedidoItemPackVrunitbase2
    ''               ,@PedidoItemPackVrunitOri
    ''               ,@PedidoItemPackVrunitOri2
    ''               ,@PedidoItemPackVrunit
    ''               ,@PedidoItemPackVrunit2
    ''               ,@PedidoItemSubtotalVr
    ''               ,@PedidoItemSubtotalVr2
    ''               ,@TabPrecoVrminCusto
    ''               ,@TabPrecoVrCusto
    ''               ,@TabPrecoPackVrminCusto
    ''               ,@TabPrecoPackVrCusto
    ''               ,@PedidoItemVrCusto
    ''               ,@PedidoItemPackVrCusto
    ''               ,@TabPrecoVrminVerba
    ''               ,@TabPrecoPackVrminVerba
    ''               ,@PedidoItemVrUnitMinVerba
    ''               ,@PedidoItemPackVrUnitMinVerba
    ''               ,@PedidoItemVerbaUtilizada
    ''               ,@PedidoItemVerbaCp
    ''               ,@PedidoItem_DescontoOri
    ''               ,@PedidoItemPackDescontoOri
    ''               ,@PedidoItemObs
    ''               ,@PrecificacaoCod
    ''               ,@PrecificacaoComissao
    ''               ,@PrecificacaoAcrescimo
    ''               ,@PrecificacaoDesconto
    ''               ,@PrecificacaoBonificacao)")}

    ''    cmd.Parameters.Add("@VendedorCod", SqlDbType.VarChar)
    ''    cmd.Parameters.Add("@PedidoNum", SqlDbType.VarChar)
    ''    cmd.Parameters.Add("@ProdutoCod", SqlDbType.VarChar)
    ''    cmd.Parameters.Add("@ProdutoStatus", SqlDbType.VarChar)
    ''    cmd.Parameters.Add("@ProdutoUe", SqlDbType.Decimal)
    ''    cmd.Parameters.Add("@TabPrecoFrete", SqlDbType.Decimal)
    ''    cmd.Parameters.Add("@PedidoItemNumitem", SqlDbType.Int)
    ''    cmd.Parameters.Add("@PedidoItemQtde", SqlDbType.Int)
    ''    cmd.Parameters.Add("@PedidoItemVrunitmin", SqlDbType.Decimal)
    ''    cmd.Parameters.Add("@PedidoItemVrunitmax", SqlDbType.Decimal)
    ''    cmd.Parameters.Add("@PedidoItemVrunitbase", SqlDbType.Decimal)
    ''    cmd.Parameters.Add("@PedidoItemVrunitbase2", SqlDbType.Decimal)
    ''    cmd.Parameters.Add("@PedidoItemVrunitOri", SqlDbType.Decimal)
    ''    cmd.Parameters.Add("@PedidoItemVrunitOri2", SqlDbType.Decimal)
    ''    cmd.Parameters.Add("@PedidoItemVrUnitario", SqlDbType.Decimal)
    ''    cmd.Parameters.Add("@PedidoItemVrunit2", SqlDbType.Decimal)
    ''    cmd.Parameters.Add("@PedidoItemDesconto", SqlDbType.Decimal)
    ''    cmd.Parameters.Add("@ProdutoPackUe", SqlDbType.Decimal)
    ''    cmd.Parameters.Add("@ProdutoPackQtde", SqlDbType.Decimal)
    ''    cmd.Parameters.Add("@TabPrecoPackFrete", SqlDbType.Decimal)
    ''    cmd.Parameters.Add("@PedidoItemPackQtde", SqlDbType.Decimal)
    ''    cmd.Parameters.Add("@PedidoItemPackVrunitmin", SqlDbType.Decimal)
    ''    cmd.Parameters.Add("@PedidoItemPackVrunitmax", SqlDbType.Decimal)
    ''    cmd.Parameters.Add("@PedidoItemPackVrunitbase", SqlDbType.Decimal)
    ''    cmd.Parameters.Add("@PedidoItemPackVrunitbase2", SqlDbType.Decimal)
    ''    cmd.Parameters.Add("@PedidoItemPackVrunitOri", SqlDbType.Decimal)
    ''    cmd.Parameters.Add("@PedidoItemPackVrunitOri2", SqlDbType.Decimal)
    ''    cmd.Parameters.Add("@PedidoItemPackVrunit", SqlDbType.Decimal)
    ''    cmd.Parameters.Add("@PedidoItemPackVrunit2", SqlDbType.Decimal)
    ''    cmd.Parameters.Add("@PedidoItemSubtotalVr", SqlDbType.Decimal)
    ''    cmd.Parameters.Add("@PedidoItemSubtotalVr2", SqlDbType.Decimal)
    ''    cmd.Parameters.Add("@TabPrecoVrminCusto", SqlDbType.Decimal)
    ''    cmd.Parameters.Add("@TabPrecoVrCusto", SqlDbType.Decimal)
    ''    cmd.Parameters.Add("@TabPrecoPackVrminCusto", SqlDbType.Decimal)
    ''    cmd.Parameters.Add("@TabPrecoPackVrCusto", SqlDbType.Decimal)
    ''    cmd.Parameters.Add("@PedidoItemVrCusto", SqlDbType.Decimal)
    ''    cmd.Parameters.Add("@PedidoItemPackVrCusto", SqlDbType.Decimal)
    ''    cmd.Parameters.Add("@TabPrecoVrminVerba", SqlDbType.Decimal)
    ''    cmd.Parameters.Add("@TabPrecoPackVrminVerba", SqlDbType.Decimal)
    ''    cmd.Parameters.Add("@PedidoItemVrUnitMinVerba", SqlDbType.Decimal)
    ''    cmd.Parameters.Add("@PedidoItemPackVrUnitMinVerba", SqlDbType.Decimal)
    ''    cmd.Parameters.Add("@PedidoItemVerbaUtilizada", SqlDbType.Decimal)
    ''    cmd.Parameters.Add("@PedidoItemVerbaCp", SqlDbType.VarChar)
    ''    cmd.Parameters.Add("@PedidoItem_DescontoOri", SqlDbType.Decimal)
    ''    cmd.Parameters.Add("@PedidoItemPackDescontoOri", SqlDbType.Decimal)
    ''    cmd.Parameters.Add("@PedidoItemObs", SqlDbType.VarChar)
    ''    cmd.Parameters.Add("@PrecificacaoCod", SqlDbType.VarChar)
    ''    cmd.Parameters.Add("@PrecificacaoComissao", SqlDbType.Decimal)
    ''    cmd.Parameters.Add("@PrecificacaoAcrescimo", SqlDbType.Decimal)
    ''    cmd.Parameters.Add("@PrecificacaoDesconto", SqlDbType.Decimal)
    ''    cmd.Parameters.Add("@PrecificacaoBonificacao", SqlDbType.Decimal)

    ''    cmd.Parameters("@VendedorCod").Value = Me.VendedorCod
    ''    cmd.Parameters("@PedidoNum").Value = Me.PedidoNum
    ''    cmd.Parameters("@ProdutoCod").Value = Me.ProdutoCod
    ''    cmd.Parameters("@ProdutoStatus").Value = Me.ProdutoStatus
    ''    cmd.Parameters("@ProdutoUe").Value = Me.ProdutoUe
    ''    cmd.Parameters("@TabPrecoFrete").Value = Me.TabprecoFrete
    ''    cmd.Parameters("@PedidoItemNumitem").Value = Me.PedidoItemNumitem
    ''    cmd.Parameters("@PedidoItemQtde").Value = Me.PedidoItemQtde
    ''    cmd.Parameters("@PedidoItemVrunitmin").Value = Me.PedidoItemVrunitmin
    ''    cmd.Parameters("@PedidoItemVrunitmax").Value = Me.PedidoItemVrunitmax
    ''    cmd.Parameters("@PedidoItemVrunitbase").Value = Me.PedidoItemVrunitbase
    ''    cmd.Parameters("@PedidoItemVrunitbase2").Value = Me.PedidoItemVrunitbase2
    ''    cmd.Parameters("@PedidoItemVrunitOri").Value = Me.PedidoItemVrunitOri
    ''    cmd.Parameters("@PedidoItemVrunitOri2").Value = Me.PedidoItemVrunitOri2
    ''    cmd.Parameters("@PedidoItemVrUnitario").Value = Me.PedidoItemVrunit
    ''    cmd.Parameters("@PedidoItemVrunit2").Value = Me.PedidoItemVrunit2
    ''    cmd.Parameters("@PedidoItemDesconto").Value = Me.PedidoItemDesconto
    ''    cmd.Parameters("@ProdutoPackUe").Value = Me.ProdutoPackUe
    ''    cmd.Parameters("@ProdutoPackQtde").Value = Me.ProdutoPackQtde
    ''    cmd.Parameters("@TabPrecoPackFrete").Value = Me.TabprecoPackFrete
    ''    cmd.Parameters("@PedidoItemPackQtde").Value = Me.PedidoItemPackQtde
    ''    cmd.Parameters("@PedidoItemPackVrunitmin").Value = Me.PedidoItemPackVrunitmin
    ''    cmd.Parameters("@PedidoItemPackVrunitmax").Value = Me.PedidoItemPackVrunitmax
    ''    cmd.Parameters("@PedidoItemPackVrunitbase").Value = Me.PedidoItemPackVrunitbase
    ''    cmd.Parameters("@PedidoItemPackVrunitbase2").Value = Me.PedidoItemPackVrunitbase2
    ''    cmd.Parameters("@PedidoItemPackVrunitOri").Value = Me.PedidoItemPackVrunitOri
    ''    cmd.Parameters("@PedidoItemPackVrunitOri2").Value = Me.PedidoItemPackVrunitOri2
    ''    cmd.Parameters("@PedidoItemPackVrunit").Value = Me.PedidoItemPackVrunit
    ''    cmd.Parameters("@PedidoItemPackVrunit2").Value = Me.PedidoItemPackVrunit2
    ''    cmd.Parameters("@PedidoItemSubtotalVr").Value = Me.PedidoItemSubtotalVr
    ''    cmd.Parameters("@PedidoItemSubtotalVr2").Value = Me.PedidoItemSubtotalVr2
    ''    cmd.Parameters("@TabPrecoVrminCusto").Value = Me.TabprecoVrminCusto
    ''    cmd.Parameters("@TabPrecoVrCusto").Value = Me.TabprecoVrCusto
    ''    cmd.Parameters("@TabPrecoPackVrminCusto").Value = Me.TabprecoPackVrminCusto
    ''    cmd.Parameters("@TabPrecoPackVrCusto").Value = Me.TabprecoPackVrCusto
    ''    cmd.Parameters("@PedidoItemVrCusto").Value = Me.PedidoItemVrCusto
    ''    cmd.Parameters("@PedidoItemPackVrCusto").Value = Me.PedidoItemPackVrCusto
    ''    cmd.Parameters("@TabPrecoVrminVerba").Value = Me.TabprecoVrminVerba
    ''    cmd.Parameters("@TabPrecoPackVrminVerba").Value = Me.TabprecoPackVrminVerba
    ''    cmd.Parameters("@PedidoItemVrUnitMinVerba").Value = Me.PedidoItemVrunitminVerba
    ''    cmd.Parameters("@PedidoItemPackVrUnitMinVerba").Value = Me.PedidoItemPackVrunitminVerba
    ''    cmd.Parameters("@PedidoItemVerbaUtilizada").Value = Me.PedidoItemVerbaUtilizada
    ''    cmd.Parameters("@PedidoItemVerbaCp").Value = Me.PedidoItemVerbaCp
    ''    cmd.Parameters("@PedidoItem_DescontoOri").Value = Me.PedidoItemDescontoOri
    ''    cmd.Parameters("@PedidoItemPackDescontoOri").Value = Me.PedidoItemPackDescontoOri
    ''    cmd.Parameters("@PedidoItemObs").Value = Me.PedidoItemObs
    ''    cmd.Parameters("@PrecificacaoCod").Value = Me.PrecificacaoCod
    ''    cmd.Parameters("@PrecificacaoComissao").Value = Me.PrecificacaoComissao
    ''    cmd.Parameters("@PrecificacaoAcrescimo").Value = Me.PrecificacaoAcrescimo
    ''    cmd.Parameters("@PrecificacaoDesconto").Value = Me.PrecificacaoDesconto
    ''    cmd.Parameters("@PrecificacaoBonificacao").Value = Me.PrecificacaoBonificacao

    ''    Sqls.Add(DataSQLExtensao.CommandAsSqlText(cmd))
    ''End Sub

    Public Sub SalvarSql(ByRef Sqls As ArrayList)

        Dim strSql As String = "INSERT INTO [dbo].[PedidoXItemIntegracaoOnSoft]" & vbCrLf &
                   "([VendedorCod]" & vbCrLf &
                   ",[PedidoNum]" & vbCrLf &
                   ",[ProdutoCod]" & vbCrLf &
                   ",[ProdutoStatus]" & vbCrLf &
                   ",[ProdutoUe]" & vbCrLf &
                   ",[TabPrecoFrete]" & vbCrLf &
                   ",[PedidoItemNumitem]" & vbCrLf &
                   ",[PedidoItemQtde]" & vbCrLf &
                   ",[PedidoItemVrunitmin]" & vbCrLf &
                   ",[PedidoItemVrunitmax]" & vbCrLf &
                   ",[PedidoItemVrunitbase]" & vbCrLf &
                   ",[PedidoItemVrunitbase2]" & vbCrLf &
                   ",[PedidoItemVrunitOri]" & vbCrLf &
                   ",[PedidoItemVrunitOri2]" & vbCrLf &
                   ",[PedidoItemVrunit]" & vbCrLf &
                   ",[PedidoItemVrunit2]" & vbCrLf &
                   ",[PedidoItemDesconto]" & vbCrLf &
                   ",[ProdutoPackUe]" & vbCrLf &
                   ",[ProdutoPackQtde]" & vbCrLf &
                   ",[TabPrecoPackFrete]" & vbCrLf &
                   ",[PedidoItemPackQtde]" & vbCrLf &
                   ",[PedidoItemPackVrunitmin]" & vbCrLf &
                   ",[PedidoItemPackVrunitmax]" & vbCrLf &
                   ",[PedidoItemPackVrunitbase]" & vbCrLf &
                   ",[PedidoItemPackVrunitbase2]" & vbCrLf &
                   ",[PedidoItemPackVrunitOri]" & vbCrLf &
                   ",[PedidoItemPackVrunitOri2]" & vbCrLf &
                   ",[PedidoItemPackVrunit]" & vbCrLf &
                   ",[PedidoItemPackVrunit2]" & vbCrLf &
                   ",[PedidoItemSubtotalVr]" & vbCrLf &
                   ",[PedidoItemSubtotalVr2]" & vbCrLf &
                   ",[TabPrecoVrminCusto]" & vbCrLf &
                   ",[TabPrecoVrCusto]" & vbCrLf &
                   ",[TabPrecoPackVrminCusto]" & vbCrLf &
                   ",[TabPrecoPackVrCusto]" & vbCrLf &
                   ",[PedidoItemVrCusto]" & vbCrLf &
                   ",[PedidoItemPackVrCusto]" & vbCrLf &
                   ",[TabPrecoVrminVerba]" & vbCrLf &
                   ",[TabPrecoPackVrminVerba]" & vbCrLf &
                   ",[PPedidoItemVrunitminVerba]" & vbCrLf &
                   ",[PedidoItemPackVrunitminVerba]" & vbCrLf &
                   ",[PedidoItemVerbaUtilizada]" & vbCrLf &
                   ",[PedidoItemVerbaCp]" & vbCrLf &
                   ",[PedidoItemDescontoOri]" & vbCrLf &
                   ",[PedidoItemPackDescontoOri]" & vbCrLf &
                   ",[PedidoItemObs]" & vbCrLf &
                   ",[PrecificacaoCod]" & vbCrLf &
                   ",[PrecificacaoComissao]" & vbCrLf &
                   ",[PrecificacaoAcrescimo]" & vbCrLf &
                   ",[PrecificacaoDesconto]" & vbCrLf &
                   ",[PrecificacaoBonificacao])" & vbCrLf &
                "VALUES" & vbCrLf &
                   "('" & Me.VendedorCod & "'" & vbCrLf &
                   ",'" & Me.PedidoNum & "'" & vbCrLf &
                   ",'" & Me.ProdutoCod & "'" & vbCrLf &
                   ",'" & Me.ProdutoStatus & "'" & vbCrLf &
                   "," & Me.ProdutoUe & vbCrLf &
                   "," & Me.TabprecoFrete & vbCrLf &
                   "," & Me.PedidoItemNumitem & vbCrLf &
                   "," & Me.PedidoItemQtde & vbCrLf &
                   "," & Me.PedidoItemVrunitmin & vbCrLf &
                   "," & Me.PedidoItemVrunitmax & vbCrLf &
                   "," & Me.PedidoItemVrunitbase & vbCrLf &
                   "," & Me.PedidoItemVrunitbase2 & vbCrLf &
                   "," & Me.PedidoItemVrunitOri & vbCrLf &
                   "," & Me.PedidoItemVrunitOri2 & vbCrLf &
                   "," & Me.PedidoItemVrunit & vbCrLf &
                   "," & Me.PedidoItemVrunit2 & vbCrLf &
                   "," & Me.PedidoItemDesconto & vbCrLf &
                   "," & Me.ProdutoPackUe & vbCrLf &
                   "," & Me.ProdutoPackQtde & vbCrLf &
                   "," & Me.TabprecoPackFrete & vbCrLf &
                   "," & Me.PedidoItemPackQtde & vbCrLf &
                   "," & Me.PedidoItemPackVrunitmin & vbCrLf &
                   "," & Me.PedidoItemPackVrunitmax & vbCrLf &
                   "," & Me.PedidoItemPackVrunitbase & vbCrLf &
                   "," & Me.PedidoItemPackVrunitbase2 & vbCrLf &
                   "," & Me.PedidoItemPackVrunitOri & vbCrLf &
                   "," & Me.PedidoItemPackVrunitOri2 & vbCrLf &
                   "," & Me.PedidoItemPackVrunit & vbCrLf &
                   "," & Me.PedidoItemPackVrunit2 & vbCrLf &
                   "," & Me.PedidoItemSubtotalVr & vbCrLf &
                   "," & Me.PedidoItemSubtotalVr2 & vbCrLf &
                   "," & Me.TabprecoVrminCusto & vbCrLf &
                   "," & Me.TabprecoVrCusto & vbCrLf &
                   "," & Me.TabprecoPackVrminCusto & vbCrLf &
                   "," & Me.TabprecoPackVrCusto & vbCrLf &
                   "," & Me.PedidoItemVrCusto & vbCrLf &
                   "," & Me.PedidoItemPackVrCusto & vbCrLf &
                   "," & Me.TabprecoVrminVerba & vbCrLf &
                   "," & Me.TabprecoPackVrminVerba & vbCrLf &
                   "," & Me.PedidoItemVrunitminVerba & vbCrLf &
                   "," & Me.PedidoItemPackVrunitminVerba & vbCrLf &
                   "," & Me.PedidoItemVerbaUtilizada & vbCrLf &
                   ",'" & Me.PedidoItemVerbaCp & "'" & vbCrLf &
                   "," & Me.PedidoItemDescontoOri & vbCrLf &
                   "," & Me.PedidoItemPackDescontoOri & vbCrLf &
                   ",'" & Me.PedidoItemObs & "'" & vbCrLf &
                   ",'" & Me.PrecificacaoCod & "'" & vbCrLf &
                   "," & Me.PrecificacaoComissao & vbCrLf &
                   "," & Me.PrecificacaoAcrescimo & vbCrLf &
                   "," & Me.PrecificacaoDesconto & vbCrLf &
                   "," & Me.PrecificacaoBonificacao & ");"

        Sqls.Add(strSql)

    End Sub

#End Region

End Class
