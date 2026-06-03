Imports Newtonsoft.Json

<Serializable()>
Public Class ListPedidoOnMobile
    Inherits List(Of PedidoOnMobile)
    Public Sub New(ByVal DataInicio As Date, ByVal DataFim As Date, ByVal ClienteId As String, Optional ByVal NumPedido As Integer = 0, Optional ByVal Situacao As String = "")
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
         "          ,p_ngs.PedidoBloqueado " & vbCrLf &
         "      FROM [dbo].[PedidoIntegracaoOnSoft] p " & vbCrLf &
         "      INNER JOIN Clientes c " & vbCrLf &
         "        ON p.ClienteCod = c.Cliente_Id " & vbCrLf &
         "      INNER JOIN [dbo].[Pedidos] p_ngs " & vbCrLf &
         "        ON p.PedidoNumPedCli = p_ngs.Pedido_Id " & vbCrLf &
         "        --AND ClienteEnderecoId = c.Endereco_Id " & vbCrLf &
         "      INNER JOIN Clientes v " & vbCrLf &
         "        ON p.VendedorCod = v.Cliente_Id " & vbCrLf &
         "      WHERE PedidoData BETWEEN '" & CDate(DataInicio).ToString("yyyy/MM/dd") & "' AND '" & CDate(DataFim).ToString("yyyy/MM/dd") & "'" & vbCrLf

        If Situacao = "L" Then
            sql += " And p_ngs.PedidoBloqueado = 0" & vbCrLf
        ElseIf Situacao = "B" Then
            sql += " And p_ngs.PedidoBloqueado = 1" & vbCrLf
        End If

        If Not String.IsNullOrEmpty(ClienteId) Then
            sql += " And p.ClienteCod = '" & ClienteId & "' " & vbCrLf
        End If

        If NumPedido > 0 Then
            sql += " And p.PedidoNumPedCli = " & NumPedido
        End If

        Dim Banco As New AcessaBanco
        Dim ds As DataSet

        ds = Banco.ConsultaDataSet(sql, "WsPedido")

        For Each row As DataRow In ds.Tables(0).Rows
            Dim ws As New PedidoOnMobile

            ws.PedidoNum = row("PedidoNum")
            ws.PedidoData = row("PedidoData")
            ws.PedidoVrpagar = row("PedidoVrpagar")
            ws.VendedorCod = row("VendedorCod")
            ws.Vendedor = row("Vendedor")
            ws.ClienteCod = row("ClienteCod")
            ws.Cliente = row("Cliente")
            ws.PedidoNumPedCli = row("PedidoNumPedCli")
            ws.Situacao = row("PedidoBloqueado")

            Me.Add(ws)
        Next
    End Sub

End Class

<Serializable()>
Public Class PedidoOnMobile

#Region "Property"
    Public Property VendedorCod As String
    Public Property PedidoNum As String
    Public Property PedidoNumPedCli As String
    Public Property ClienteCod As String
    Public Property PedidoData As String
    Public Property Vendedor As String
    Public Property Cliente As String
    Public Property PedidoVrpagar As Decimal
    Public Property Situacao As Boolean

#End Region

End Class
