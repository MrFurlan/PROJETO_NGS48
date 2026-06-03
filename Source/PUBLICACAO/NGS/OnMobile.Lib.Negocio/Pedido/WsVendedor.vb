Imports System.Configuration
Imports Newtonsoft.Json

<Serializable()>
Public Class ListWsVendedor
    Inherits List(Of WsVendedor)

    Public Sub New()
        Dim sql As String
        sql = String.Concat("select c.Cliente_Id AS Representante_Id, c.Endereco_Id, c.Nome, c.Endereco, c.Cidade, c.Estado 
                            from clientes c
	                            inner join ClientesXTipos cXt
			                            on cXt.Cliente_Id   = c.Cliente_Id
			                            and cXt.Endereco_Id = c.Endereco_Id
			                            and Tipo_Id = 6 --Representantes (ETIPOS)
	                            inner join ClienteXRepresentante cXr  -- Representante
			                            on cXr.Representante_Id     = c.Cliente_Id
			                            and cXr.EndRepresentante_Id = c.Endereco_Id
	                            inner join ClientesXTabelasDePrecos cXtp
			                            on cXtp.Cliente_Id     = c.Cliente_Id
			                            and cXtp.Endereco_Id = c.Endereco_Id
                            group by c.Cliente_Id, c.Endereco_Id, c.Nome, c.Endereco, c.Cidade, c.Estado")

        Dim Banco As New AcessaBancoOnMobile

        Dim ds = Banco.ConsultaDataSet(sql, "ClienteXRepresentante")

        For Each row As DataRow In ds.Tables(0).Rows
            Dim ws As New WsVendedor
            ws.VendedorCod = row("Representante_Id")
            ws.VendedorNome = Left(row("Nome"), 30)
            Me.Add(ws)
        Next
    End Sub
End Class

<Serializable()>
Public Class WsVendedor

    <JsonProperty("vendedorCod")>
    Public Property VendedorCod As String

    <JsonProperty("vendedorNome")>
    Public Property VendedorNome As String

End Class

