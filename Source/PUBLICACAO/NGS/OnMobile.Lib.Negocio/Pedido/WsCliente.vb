Imports System.Configuration
Imports Newtonsoft.Json
Imports NGS.Lib.Negocio

<Serializable()>
Public Class ListWsCliente
    Inherits List(Of WsCliente)

    Public Sub New()
        Dim Banco As New AcessaBancoOnMobile

        Dim Sql As String

        Sql = String.Concat("
            select c.Cliente_Id AS Cliente_Id, 
		            c.Endereco_Id, c.Nome, c.Endereco, c.Bairro, c.Complemento, c.Cep, c.Cidade, c.Estado, 
	                c.Telefone, c.Fax, c.Email, c.Fantasia, c.Inscricao, c.Site, c.Categoria, c.Situacao, s.Descricao as DescricaoSituacao, c.ClienteDesde,
		            repre.Cliente_Id AS Representante_Id, repre.Endereco_Id AS EnderecoRepresentante, repre.Nome AS NomeRepresentante, repre.Cidade AS CidadeRepre, repre.Estado AS EstadoRepre,
		            cXtp.Tabela_Id
            from clientes c
	            inner join ClienteXRepresentante cXr  -- Representante
			            on cXr.Cliente_Id     = c.Cliente_Id
			            and cXr.EndCliente_Id = c.Endereco_Id
	            inner join ClientesXTipos cXtR
			            on cXtR.Cliente_Id   = c.Cliente_Id
			            and cXtR.Endereco_Id = c.Endereco_Id
			            and Tipo_Id in(4) --Clientes (ETIPOS)
	            inner join Clientes repre
			            on repre.Cliente_Id   = cXr.Representante_Id
			            and repre.Endereco_Id = cXr.EndRepresentante_Id
	            inner join ClienteXRepresentante cXrCD  -- CD/Representante
			            on cXrCD.Representante_Id     = repre.Cliente_Id
			            and cXr.EndRepresentante_Id   = repre.Endereco_Id
	            inner join Clientes cd
			            on cd.Cliente_Id   = cXrCD.Cliente_Id
			            and cd.Endereco_Id = cXrCD.EndCliente_Id
	            inner join ClientesXTipos cdXt
			            on cdXt.Cliente_Id   = cd.Cliente_Id
			            and cdXt.Endereco_Id = cd.Endereco_Id
			            and cdXt.Tipo_Id = 19 --CentroDeDistribuicao (ETIPOS)
	            inner join ProdutosXPrecos pXp --Lista de Preços ref. o CD
			            on pXp.Cliente_Id     = cd.Cliente_Id
			            and pXp.EndCliente_Id = cd.Endereco_Id
			            and left(pXp.Produto_Id, 1) = '8'
	            inner join ClientesXTabelasDePrecos cXtp
			            on cXtp.Cliente_Id  = repre.Cliente_Id
			            and cXtp.Endereco_Id = repre.Endereco_Id 
	            inner join Situacoes s
			            on c.Situacao = s.Situacao_Id
            where c.Cliente_Id not in(select Cliente_id from ClientesXTipos where Tipo_Id = 19) --Clientes (ETIPOS)
            group by c.Cliente_Id, c.Endereco_Id, c.Nome, c.Endereco, c.Bairro, c.Complemento, c.Cep, c.Cidade, c.Estado, 
	                c.Telefone, c.Fax, c.Email, c.Fantasia, c.Inscricao, c.Site, c.Categoria, c.Situacao, s.Descricao, c.ClienteDesde,
		            repre.Cliente_Id, repre.Endereco_Id, repre.Nome, repre.Cidade, repre.Estado, cXtp.Tabela_Id
            order by Nome asc")

        Dim ds As DataSet = Banco.ConsultaDataSet(Sql, "Clientes")
        Dim sequencia = 1

        For Each row As DataRow In ds.Tables(0).Rows
            Dim ws As New WsCliente
            ws.VendedorCod = row("Representante_Id")
            ws.ClienteCod = row("Cliente_Id")
            ws.CategoriaCod = "C"
            'ws.GrupoclienteCod
            'ws.SubgrupoClienteCod
            ws.ClienteOrdem = sequencia
            ws.ClienteNomefantasia = row("Fantasia").ToString()
            ws.ClienteRazaosocial = row("Nome").ToString()
            ws.ClienteLogradouro = row("Endereco").ToString()
            ws.ClienteBairro = Left(row("Bairro").ToString(), 40)
            ws.ClienteComplemento = Left(row("Complemento").ToString(), 40)
            ws.ClienteCidade = row("Cidade").ToString()
            ws.ClienteCep = row("Cep").ToString()
            ws.ClienteUf = row("Estado").ToString()
            ws.ClienteFax = row("Fax").ToString()
            ws.ClienteEmail = row("Email").ToString()
            ws.ClientePessoa = IIf(row("Cliente_Id").ToString().Length = 14, "J", "F")
            ws.ClienteCnpjcpf = row("Cliente_Id").ToString()
            ws.ClienteIe = row("Inscricao").ToString()
            'ws.ClienteIm = row("").ToString()
            'ws.ClienteContato1 = row("").ToString()
            'ws.ClienteContato2 = row("").ToString()
            ws.ClienteTel1 = Left(row("Telefone").ToString(), 11)
            'ws.ClienteTel2 = row("").ToString()
            ws.ClienteHomepage = row("Site").ToString()
            'ws.ClienteLimitecredito = row("").ToString()
            ws.ClienteStatus = "A"
            ws.ClienteStatusDescricao = row("DescricaoSituacao").ToString()
            ws.TabPrecoCod = row("Tabela_Id")
            ws.TabCondicaoCod = "001"
            'ws.TabDescontoCod = row("").ToString() - FURLAN
            ws.ClienteRnCrPzoAlerta = 0
            ws.ClienteRnCrPzoBloq = 0
            ws.ClienteRInicio = Nothing
            ws.ClienteRTermino = Nothing
            ws.ClienteDtEnvio = Nothing

            Me.Add(ws)
            sequencia = (sequencia + 1)
        Next

    End Sub
End Class

<Serializable()>
Public Class WsCliente
    <JsonProperty("vendedorCod")>
    Public Property VendedorCod As String

    <JsonProperty("clienteCod")>
    Public Property ClienteCod As String

    <JsonProperty("categoriaCod")>
    Public Property CategoriaCod As String

    <JsonProperty("grupoclienteCod")>
    Public Property GrupoclienteCod As String

    <JsonProperty("subgrupoClienteCod")>
    Public Property SubgrupoClienteCod As String

    <JsonProperty("clienteOrdem")>
    Public Property ClienteOrdem As Integer

    <JsonProperty("clienteNomefantasia")>
    Public Property ClienteNomefantasia As String

    <JsonProperty("clienteRazaosocial")>
    Public Property ClienteRazaosocial As String

    <JsonProperty("clienteLogradouro")>
    Public Property ClienteLogradouro As String

    <JsonProperty("clienteBairro")>
    Public Property ClienteBairro As String

    <JsonProperty("clienteComplemento")>
    Public Property ClienteComplemento As String

    <JsonProperty("clienteCidade")>
    Public Property ClienteCidade As String

    <JsonProperty("clienteCep")>
    Public Property ClienteCep As String

    <JsonProperty("clienteUf")>
    Public Property ClienteUf As String

    <JsonProperty("clienteFax")>
    Public Property ClienteFax As String

    <JsonProperty("clienteEmail")>
    Public Property ClienteEmail As String

    <JsonProperty("clientePessoa")>
    Public Property ClientePessoa As String

    <JsonProperty("clienteCnpjcpf")>
    Public Property ClienteCnpjcpf As String

    <JsonProperty("clienteIe")>
    Public Property ClienteIe As String

    <JsonProperty("clienteIm")>
    Public Property ClienteIm As String

    <JsonProperty("clienteContato1")>
    Public Property ClienteContato1 As String

    <JsonProperty("clienteContato2")>
    Public Property ClienteContato2 As String

    <JsonProperty("clienteTel1")>
    Public Property ClienteTel1 As String

    <JsonProperty("clienteTel2")>
    Public Property ClienteTel2 As String

    <JsonProperty("clienteHomepage")>
    Public Property ClienteHomepage As String

    <JsonProperty("clienteLimitecredito")>
    Public Property ClienteLimitecredito As Double

    <JsonProperty("clienteStatus")>
    Public Property ClienteStatus As String

    <JsonProperty("clienteStatusdescricao")>
    Public Property ClienteStatusDescricao As String

    <JsonProperty("tabprecoCod")>
    Public Property TabPrecoCod As String

    <JsonProperty("tabcondicaoCod")>
    Public Property TabCondicaoCod As String

    <JsonProperty("tabdescontoCod")>
    Public Property TabDescontoCod As String

    <JsonProperty("clienteRnCrPzoAlerta")>
    Public Property ClienteRnCrPzoAlerta As String

    <JsonProperty("clienteRnCrPzoBloq")>
    Public Property ClienteRnCrPzoBloq As String

    <JsonProperty("clienteRinicio")>
    Public Property ClienteRInicio As DateTime?

    <JsonProperty("clienteRtermino")>
    Public Property ClienteRTermino As DateTime?

    <JsonProperty("clienteDtenvio")>
    Public Property ClienteDtEnvio As DateTime?
    Public Property CodigoEndereco As String

    Public Sub New()

    End Sub
    Public Sub New(ByVal pCodigo As String, ByVal pCodigoEndereco As Integer)
        Dim Banco As New AcessaBancoOnMobile()

        Try
            Dim Sql As String

            Sql = "SELECT Cliente_Id, Endereco_Id " & vbCrLf &
                  "  FROM Clientes" & vbCrLf &
                  " WHERE Cliente_Id  = '" & pCodigo & "'" & vbCrLf &
                  "   AND Endereco_Id = " & pCodigoEndereco

            Dim ds As DataSet = Banco.ConsultaDataSet(Sql, "Clientes")
            If ds.Tables(0).Rows.Count = 0 Then Exit Sub
            Dim row As DataRow = ds.Tables(0).Rows(0)

            Me.ClienteCod = row("Cliente_Id")
            Me.CodigoEndereco = row("Cliente_Id")

        Catch ex As Exception
            Throw New Exception(ex.Message)
        Finally
            Banco = Nothing
        End Try
    End Sub

    Public Sub PreencheCliente(ByVal sqls As ArrayList, ByVal codigoCliente As String, ByVal codigoVendedor As String)
        Dim clienteNGS = New Cliente(codigoCliente)
        clienteNGS.IUD = "I"
        clienteNGS.CEP = Funcoes.RemoverLetrasDoNumero(clienteNGS.CEP)
        clienteNGS.ClienteDesde = Today()
        clienteNGS.Tipos = New ListClientexTipo(clienteNGS)

        Dim clienteXTipos As New ClientexTipo
        clienteXTipos.IUD = "I"
        clienteXTipos.CodigoTipo = 4 'CLIENTES
        clienteXTipos.Cliente = clienteNGS
        clienteNGS.Tipos.Add(clienteXTipos)

        Dim objClienteXRepresentante = New ClienteXRepresentante()
        objClienteXRepresentante.IUD = "I"
        objClienteXRepresentante.Cliente = clienteNGS
        objClienteXRepresentante.CodigoRepresentante = codigoVendedor
        objClienteXRepresentante.EndRepresentante = 0
        objClienteXRepresentante.Principal = False

        clienteNGS.SalvarSql(sqls)
    End Sub
End Class
