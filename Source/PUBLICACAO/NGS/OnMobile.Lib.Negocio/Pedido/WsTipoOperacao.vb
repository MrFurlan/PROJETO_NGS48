Imports System.Configuration
Imports Newtonsoft.Json

<Serializable()>
Public Class ListWsTipoOperacao
    Inherits List(Of WsTipoOperacao)
    Public Sub New(Optional ByVal CarregarDados As Boolean = False)
        Dim sql As String = String.Concat("
            SELECT 
	            p.Operacao_Id, 
	            so.SubOperacoes_Id, 
	            so.Descricao 
            FROM Operacoes p
            INNER JOIN SubOperacoes so on p.Operacao_Id = so.Operacao_Id
            WHERE p.Operacao_Id = 21 and so.SubOperacoes_Id = 1")

        Dim Banco As New AcessaBancoOnMobile
        Dim ds As DataSet

        ds = Banco.ConsultaDataSet(sql, "TipoOperacaoes")

        For Each row As DataRow In ds.Tables(0).Rows
            Dim ws As New WsTipoOperacao
            ws.TipoOperacaoCod = String.Concat(row("Operacao_Id"), ".", row("SubOperacoes_Id"))
            ws.TipoOperacaoDescricao = row("Descricao")
            Me.Add(ws)
        Next
    End Sub
End Class

<Serializable()>
Public Class WsTipoOperacao
    <JsonProperty("tipooperacaoCod")>
    Public Property TipoOperacaoCod As String

    <JsonProperty("tipooperacaoDescricao")>
    Public Property TipoOperacaoDescricao As String

    <JsonProperty("tipooperacaoRn")>
    Public Property TipoOperacaoRn As String

    <JsonProperty("tipooperacaoObs")>
    Public Property TipoOperacaoObs As String

    <JsonProperty("tipooperacaoObsPedido")>
    Public Property TipoOperacaoObsPedido As String

    <JsonProperty("tipooperacaoRecNo")>
    Public Property TipoOperacaoRecNo As Integer
End Class
