Imports System.Security.Claims
Imports System.Text
Imports System.Web.Http
Imports System.Web
Imports System.Net.Http
Imports System.Net
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis
Imports NGS.Web.Api.Models

<Authorize>
Public Class ProdutoDashboardRepositorio
    Implements IProdutoDashboardRepositorio

#Region "Construtor"

    Public Sub New()
    End Sub

#End Region

#Region "Processos"

    'Consultar Banco de dados
    Function ConsultarBancoDeDados() As List(Of SelectListItem) Implements IProdutoDashboardRepositorio.ConsultarBancoDeDados

        Try

            Conexao()
            Dim lstBancos As List(Of BancoDados) = BancoDados.Listar()
            Dim listaBanco As New List(Of SelectListItem)

            For Each Dr In lstBancos

                Dim textoExibido As String = String.Format("{0} | {1}", Dr.Banco_Id.Trim().PadRight(20 - Dr.Banco_Id.Length, "."), Dr.HostServidor)

                Dim valor As String = String.Format("{0}|{1}", Dr.Banco_Id.Trim(), Dr.HostServidor)
                listaBanco.Add(New SelectListItem With {.Text = textoExibido, .Value = valor})

            Next

            Return listaBanco

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try

    End Function

    'Empresas 
    Public Function ConsultarEmpresas() As List(Of SelectListItem) Implements IProdutoDashboardRepositorio.ConsultarEmpresas

        Try

            Conexao()
            Return CarregarEmpresas()

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try

    End Function

    'Seguimentos 
    Public Function ConsultarSeguimentos() As List(Of SelectListItem) Implements IProdutoDashboardRepositorio.ConsultarSeguimentos

        Try

            Conexao()
            Return CarregarSeguimentos()

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try

    End Function

    'Grafico Produtos
    Public Function TotalProdutoPorProduto(ByVal filtro As ProdutoDashboardFiltro) As ProdutoDashboard Implements IProdutoDashboardRepositorio.TotalProdutoPorProduto

        Try

            Conexao()
            Return CarregarTotalProdutoProduto(filtro)

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try

    End Function

    Public Function FaturamentoPorProduto(ByVal filtro As ProdutoDashboardFiltro) As IEnumerable(Of ProdutoDashboard) Implements IProdutoDashboardRepositorio.FaturamentoPorProduto

        Try

            Conexao()
            Return CarregarFaturamentoPorProduto(filtro)

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try

    End Function

    Public Function DistribuicaoDeFaturamentoPorProduto(ByVal filtro As ProdutoDashboardFiltro) As IEnumerable(Of ProdutoDashboard) Implements IProdutoDashboardRepositorio.DistribuicaoDeFaturamentoPorProduto

        Try

            Conexao()
            Return CarregarDistribuicaoDeFaturamentoPorProduto(filtro)

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try

    End Function

    Public Function DetalhamentoProdutos(ByVal filtro As ProdutoDashboardFiltro) As IEnumerable(Of ProdutoDashboard) Implements IProdutoDashboardRepositorio.DetalhamentoProdutos

        Try

            Conexao()
            Return CarregarDetalhamentoProdutos(filtro)

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try

    End Function

    'Clientes
    Public Function TotalProdutoPorClientes(ByVal filtro As ProdutoDashboardFiltro) As ProdutoDashboard Implements IProdutoDashboardRepositorio.TotalProdutoPorClientes

        Try

            Conexao()
            Return CarregarTotalProdutoPorClientes(filtro)

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try

    End Function

    Public Function FaturamentoPorClientes(ByVal filtro As ProdutoDashboardFiltro) As IEnumerable(Of ProdutoDashboard) Implements IProdutoDashboardRepositorio.FaturamentoPorClientes

        Try

            Conexao()
            Return CarregarFaturamentoPorClientes(filtro)

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try

    End Function

    Public Function DetalhamentoClientes(ByVal filtro As ProdutoDashboardFiltro) As IEnumerable(Of ProdutoDashboard) Implements IProdutoDashboardRepositorio.DetalhamentoClientes

        Try

            Conexao()
            Return CarregarDetalhamentoClientes(filtro)

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try

    End Function

    'Representantes
    Public Function TotalProdutoPorRepresentantes(ByVal filtro As ProdutoDashboardFiltro) As ProdutoDashboard Implements IProdutoDashboardRepositorio.TotalProdutoPorRepresentantes

        Try

            Conexao()
            Return CarregarTotalProdutoPorRepresentantes(filtro)

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try

    End Function

    Public Function FaturamentoPorRepresentantes(ByVal filtro As ProdutoDashboardFiltro) As IEnumerable(Of ProdutoDashboard) Implements IProdutoDashboardRepositorio.FaturamentoPorRepresentantes

        Try

            Conexao()
            Return CarregarFaturamentoPorRepresentantes(filtro)

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try

    End Function

    Public Function DetalhamentoRepresentantes(ByVal filtro As ProdutoDashboardFiltro) As IEnumerable(Of ProdutoDashboard) Implements IProdutoDashboardRepositorio.DetalhamentoRepresentantes

        Try

            Conexao()
            Return CarregarDetalhamentoRepresentantes(filtro)

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try

    End Function

    'Cidades
    Public Function TotalProdutoPorCidades(ByVal filtro As ProdutoDashboardFiltro) As ProdutoDashboard Implements IProdutoDashboardRepositorio.TotalProdutoPorCidades

        Try

            Conexao()
            Return CarregarTotalProdutoPorCidades(filtro)

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try

    End Function

    Public Function FaturamentoPorCidades(ByVal filtro As ProdutoDashboardFiltro) As IEnumerable(Of ProdutoDashboard) Implements IProdutoDashboardRepositorio.FaturamentoPorCidades

        Try

            Conexao()
            Return CarregarFaturamentoPorCidades(filtro)

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try

    End Function

    Public Function DetalhamentoCidades(ByVal filtro As ProdutoDashboardFiltro) As IEnumerable(Of ProdutoDashboard) Implements IProdutoDashboardRepositorio.DetalhamentoCidades

        Try

            Conexao()
            Return CarregarDetalhamentoCidades(filtro)

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try

    End Function

    'Estados
    Public Function TotalProdutoPorEstados(ByVal filtro As ProdutoDashboardFiltro) As ProdutoDashboard Implements IProdutoDashboardRepositorio.TotalProdutoPorEstados

        Try

            Conexao()
            Return CarregarTotalProdutoPorEstados(filtro)

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try

    End Function

    Public Function FaturamentoPorEstados(ByVal filtro As ProdutoDashboardFiltro) As IEnumerable(Of ProdutoDashboard) Implements IProdutoDashboardRepositorio.FaturamentoPorEstados

        Try

            Conexao()
            Return CarregarFaturamentoPorEstados(filtro)

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try

    End Function

    Public Function DetalhamentoEstados(ByVal filtro As ProdutoDashboardFiltro) As IEnumerable(Of ProdutoDashboard) Implements IProdutoDashboardRepositorio.DetalhamentoEstados

        Try

            Conexao()
            Return CarregarDetalhamentoEstados(filtro)

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try

    End Function

    'Pedidos
    Public Function TotalProdutoPorPedidos(ByVal filtro As ProdutoDashboardFiltro) As ProdutoDashboard Implements IProdutoDashboardRepositorio.TotalProdutoPorPedidos

        Try

            Conexao()
            Return CarregarTotalProdutoPorPedidos(filtro)

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try

    End Function

    Public Function FaturamentoPorPedidos(ByVal filtro As ProdutoDashboardFiltro) As IEnumerable(Of ProdutoDashboard) Implements IProdutoDashboardRepositorio.FaturamentoPorPedidos

        Try

            Conexao()
            Return CarregarFaturamentoPorPedidos(filtro)

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try

    End Function

    Public Function DetalhamentoPedidos(ByVal filtro As ProdutoDashboardFiltro) As IEnumerable(Of ProdutoDashboard) Implements IProdutoDashboardRepositorio.DetalhamentoPedidos

        Try

            Conexao()
            Return CarregarDetalhamentoPedidos(filtro)

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try

    End Function

    'Vendas Diárias
    Public Function TotalVendasDiarias(ByVal filtro As ProdutoDashboardFiltro) As ProdutoDashboard Implements IProdutoDashboardRepositorio.TotalVendasDiarias

        Try

            Conexao()
            Return CarregarTotalVendasDiarias(filtro)

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try

    End Function

    Public Function EvolucaoVendasDiarias(ByVal filtro As ProdutoDashboardFiltro) As IEnumerable(Of ProdutoDashboard) Implements IProdutoDashboardRepositorio.EvolucaoVendasDiarias

        Try

            Conexao()
            Return CarregarEvolucaoVendasDiarias(filtro)

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try

    End Function

    Public Function DetalhamentoVendasDiarias(ByVal filtro As ProdutoDashboardFiltro) As IEnumerable(Of ProdutoDashboard) Implements IProdutoDashboardRepositorio.DetalhamentoVendasDiarias

        Try

            Conexao()
            Return CarregarDetalhamentoVendasDiarias(filtro)

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try

    End Function

    'Vendas Mensais
    Public Function TotalVendasMensais(ByVal filtro As ProdutoDashboardFiltro) As ProdutoDashboard Implements IProdutoDashboardRepositorio.TotalVendasMensais

        Try

            Conexao()
            Return CarregarTotalVendasMensais(filtro)

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try

    End Function

    Public Function EvolucaoVendasMensais(ByVal filtro As ProdutoDashboardFiltro) As IEnumerable(Of ProdutoDashboard) Implements IProdutoDashboardRepositorio.EvolucaoVendasMensais

        Try

            Conexao()
            Return CarregarEvolucaoVendasMensais(filtro)

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try

    End Function

    Public Function DetalhamentoVendasMensais(ByVal filtro As ProdutoDashboardFiltro) As IEnumerable(Of ProdutoDashboard) Implements IProdutoDashboardRepositorio.DetalhamentoVendasMensais

        Try

            Conexao()
            Return CarregarDetalhamentoVendasMensais(filtro)

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try

    End Function

    'Consolidada
    Public Function TotalConsolidado(ByVal filtro As ProdutoDashboardFiltro) As ProdutoDashboard Implements IProdutoDashboardRepositorio.TotalConsolidado

        Try

            Conexao()
            Return CarregarTotalConsolidado(filtro)

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try

    End Function

    Public Function EvolucaoConsolidado(ByVal filtro As ProdutoDashboardFiltro) As IEnumerable(Of ProdutoDashboard) Implements IProdutoDashboardRepositorio.EvolucaoConsolidado

        Try

            Conexao()
            Return CarregarEvolucaoConsolidado(filtro)

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try

    End Function

    Public Function DetalhamentoConsolidado(ByVal filtro As ProdutoDashboardFiltro) As IEnumerable(Of ProdutoDashboard) Implements IProdutoDashboardRepositorio.DetalhamentoConsolidado

        Try

            Conexao()
            Return CarregarDetalhamentoConsolidado(filtro)

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try

    End Function

    Public Function ComparacaoVolumeXFaturamentoConsolidado(ByVal filtro As ProdutoDashboardFiltro) As IEnumerable(Of ProdutoDashboard) Implements IProdutoDashboardRepositorio.ComparacaoVolumeXFaturamentoConsolidado

        Try

            Conexao()
            Return CarregarComparacaoVolumeXFaturamentoConsolidado(filtro)

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try

    End Function


#End Region

#Region "Conexao"

    Public Sub Conexao()

        Dim data As DataModel
        data = DataAcess.ReadDataFromXmlFile()

        If UsuarioServidor.Conexao Is Nothing Then
            UsuarioServidor.Conexao = UsuarioLocal.GetStringConexao(data.HashedUsuario, data.NomeServidor, data.Banco, data.Host)
        End If

        Dim strConexao As String() = UsuarioServidor.Conexao.ToString.Split(New Char() {";"}, StringSplitOptions.RemoveEmptyEntries)


        If strConexao.Count > 0 Then

            UsuarioServidor.NomeServidor = strConexao(0).Replace("Data Source=", String.Empty).Trim()
            UsuarioServidor.EnderecoLocal = UsuarioServidor.Conexao.ToString()
            UsuarioServidor.BancoDeDados = strConexao(1).Replace("Initial Catalog=", String.Empty).Trim()

            'response.Cookies("conexao").Value = FuncoesStrings.CodificarPara64Bits(UsuarioServidor.Conexao)
            'Dim response = New HttpResponseMessage(HttpStatusCode.OK) ' Cria uma resposta HTTP com status 200 OK
            Dim cookie As New HttpCookie("conexao", FuncoesStrings.CodificarPara64Bits(UsuarioServidor.Conexao)) ' Cria o cookie com o nome "conexao" e valor "some_value"
            cookie.Expires = DateTime.Now.AddDays(1) ' Define a expiração do cookie para um dia
            HttpContext.Current.Response.Cookies.Add(cookie) ' Adiciona o cookie à resposta

            If String.IsNullOrEmpty(HttpContext.Current.Response.Cookies("conexao").Value) Then

                Throw New Exception("Usuário não cadastrado para nenhum banco de dados!")

            End If

        Else

            Throw New Exception("Este usuário não existe no banco de dados de Usuários!")

        End If

    End Sub

#End Region

#Region "Repositorios e Querys"

    'CarregarEmpresas
    Private Function CarregarEmpresas() As List(Of SelectListItem)

        Dim Sql As String = " SELECT DISTINCT Clientes.Cliente_Id AS Codigo, 
		                            Clientes.Endereco_Id, 
		                            Clientes.Reduzido, 
		                            Clientes.Nome, 
		                            Clientes.Cidade, 
		                            Clientes.Estado, 
		                            CASE WHEN ISNULL(cxe.matriz,'N') = 'S' THEN 'Matriz' ELSE 'Filial' END MF
                               FROM GruposXEmpresas 
                              INNER JOIN Clientes
                                 ON GruposXEmpresas.Cliente_Id    = Clientes.Cliente_Id 
                                AND GruposXEmpresas.EndCliente_Id = Clientes.Endereco_Id
                              INNER Join ClientesXEmpresas cxe
                                 ON cxe.Empresa_Id    = Clientes.Cliente_Id
                                AND cxe.EndEmpresa_Id = Clientes.Endereco_Id;"

        Dim banco As New AcessaBanco
        Dim dsDashboard As DataSet = banco.ConsultaDataSet(Sql, "Empresas")

        Dim listaEmpresas As New List(Of SelectListItem)

        If dsDashboard IsNot Nothing AndAlso dsDashboard.Tables IsNot Nothing AndAlso dsDashboard.Tables.Count > 0 AndAlso dsDashboard.Tables(0).Rows.Count > 0 Then
            For Each Dr As DataRow In dsDashboard.Tables(0).Rows
                Dim textoExibido As String = Funcoes.AlinharEsquerda(Dr("Nome"), 28, ".") & " - " &
                                     Funcoes.AlinharEsquerda(Dr("Cidade"), 20, ".") & " " &
                                     Dr("Estado") & " " &
                                     Funcoes.AlinharEsquerda(Funcoes.FormatarCpfCnpj(Dr("Codigo")), 18, ".") & "-" &
                                     CStr(Dr("Endereco_Id")) & "-" & Dr("Reduzido") & "-" & Dr("MF")

                Dim valor As String = Dr("Codigo").ToString() & "|" & Dr("Endereco_Id").ToString()
                listaEmpresas.Add(New SelectListItem With {.Text = textoExibido, .Value = valor})
            Next
        End If


        Return listaEmpresas

    End Function

    'CarregarSeguimentos
    Private Function CarregarSeguimentos() As List(Of SelectListItem)

        Dim Sql As String = "   SELECT Seguimento_Id, Descricao FROM Seguimentos
                                UNION
                                SELECT 0 as Seguimento_Id, 'TODOS' AS Descricao; "

        Dim banco As New AcessaBanco
        Dim dsDashboard As DataSet = banco.ConsultaDataSet(Sql, "Seguimentos")

        Dim listaSeguimentos As New List(Of SelectListItem)

        If dsDashboard IsNot Nothing AndAlso dsDashboard.Tables IsNot Nothing AndAlso dsDashboard.Tables.Count > 0 AndAlso dsDashboard.Tables(0).Rows.Count > 0 Then
            For Each Dr As DataRow In dsDashboard.Tables(0).Rows
                Dim textoExibido As String = Dr("Descricao")
                Dim valor As String = Dr("Seguimento_Id").ToString()
                listaSeguimentos.Add(New SelectListItem With {.Text = textoExibido, .Value = valor})
            Next
        End If

        Return listaSeguimentos

    End Function

    'Grafico Produtos
    Public Function CarregarTotalProdutoProduto(ByVal filtro As ProdutoDashboardFiltro) As ProdutoDashboard

        Dim sql As String = String.Empty

        Dim sqlCondicaoIN As String = ""

        If Not String.IsNullOrEmpty(filtro.EmpresasSelecionada) Then
            Dim listaFormatada = filtro.EmpresasSelecionada.Split(","c).
                                    Where(Function(s) Not String.IsNullOrWhiteSpace(s)).
                                    Select(Function(par) $"'{par.Trim()}'")

            sqlCondicaoIN = String.Join(",", listaFormatada)

        End If

        Dim filtroSeguimento As String = IIf(filtro.SeguimentoSelecionado <> "0", String.Format("AND PRO.Seguimento = {0} ", filtro.SeguimentoSelecionado), "")

        sql = " WITH total_vendas AS 
	                (
		                SELECT 	
			                NF.Empresa_Id,
			                NF.Cliente_Id,
			                NF.Serie_Id,
			                NF.Nota_Id,
			                FORMAT(NF.Movimento, 'dd/MM/yyyy') AS Data,
			                NFI.Produto_Id,
			                PRO.Nome AS ProdutoNome,
			                SEG.Descricao AS Seguimento,
			                NF.Operacao AS OperacaoNota,
			                NF.SubOperacao AS SubOperacaoNota,
			                NFI.QuantidadeFiscal,
			                NFI.PesoFiscal,
			                NFI.Unitario,
			                NFE_LIQUIDO.Valor AS ValorLiquido, --FORMAT(NFE_LIQUIDO.Valor, 'N2', 'pt-br') AS ValorLiquido,
			                NFE_PRO.Valor AS ValorTotal, --FORMAT(NFE_PRO.Valor, 'N2', 'pt-br') AS ValorTotal
                            SO.Devolucao
		                FROM NotasFiscais NF
		                INNER JOIN NotasFiscaisXItens NFI 
			                ON NF.Empresa_Id				= NFI.Empresa_Id 
			                AND NF.EndEmpresa_Id			= NFI.EndEmpresa_Id 
			                AND NF.Cliente_Id				= NFI.Cliente_Id 
			                AND NF.EndCliente_Id			= NFI.EndCliente_Id 
			                AND NF.EntradaSaida_Id			= NFI.EntradaSaida_Id 
			                AND NF.Serie_Id					= NFI.Serie_Id 
			                AND NF.Nota_Id					= NFI.Nota_Id
		                INNER JOIN NotasFiscaisXEncargos NFE_LIQUIDO
			                ON NF.Empresa_Id				= NFE_LIQUIDO.Empresa_Id 
			                AND NF.EndEmpresa_Id			= NFE_LIQUIDO.EndEmpresa_Id 
			                AND NF.Cliente_Id				= NFE_LIQUIDO.Cliente_Id 
			                AND NF.EndCliente_Id			= NFE_LIQUIDO.EndCliente_Id 
			                AND NF.EntradaSaida_Id			= NFE_LIQUIDO.EntradaSaida_Id 
			                AND NF.Serie_Id					= NFE_LIQUIDO.Serie_Id 
			                AND NF.Nota_Id					= NFE_LIQUIDO.Nota_Id
			                AND NFI.Produto_Id				= NFE_LIQUIDO.Produto_Id
			                AND NFI.Sequencia_Id			= NFE_LIQUIDO.Sequencia_id
			                AND NFE_LIQUIDO.Encargo_Id		= 'LIQUIDO'
		                INNER JOIN NotasFiscaisXEncargos NFE_PRO
			                ON NF.Empresa_Id				= NFE_PRO.Empresa_Id 
			                AND NF.EndEmpresa_Id			= NFE_PRO.EndEmpresa_Id 
			                AND NF.Cliente_Id				= NFE_PRO.Cliente_Id 
			                AND NF.EndCliente_Id			= NFE_PRO.EndCliente_Id 
			                AND NF.EntradaSaida_Id			= NFE_PRO.EntradaSaida_Id 
			                AND NF.Serie_Id					= NFE_PRO.Serie_Id 
			                AND NF.Nota_Id					= NFE_PRO.Nota_Id
			                AND NFI.Produto_Id				= NFE_PRO.Produto_Id
			                AND NFI.Sequencia_Id			= NFE_PRO.Sequencia_id
			                AND NFE_PRO.Encargo_Id			= 'PRODUTO'
		                INNER JOIN SubOperacoes SO
			                ON NF.Operacao = SO.Operacao_Id 
			                AND NF.SubOperacao = SO.SubOperacoes_Id
		                INNER JOIN Produtos PRO
			                ON NFI.Produto_Id = PRO.Produto_Id
		                INNER JOIN Seguimentos SEG
			                ON PRO.Seguimento = SEG.Seguimento_Id
                            AND PRO.Dashboard = 1
		                WHERE 
			                ISNULL(NF.Situacao, 1)  = 1
			                AND (NF.Empresa_Id + '|' + CONVERT(VARCHAR, NF.EndEmpresa_Id)) IN (" & IIf(String.IsNullOrEmpty(filtro.EmpresasSelecionada), "(NF.Empresa_Id + '|' + NF.EndEmpresa_Id)", sqlCondicaoIN) & ")
			                AND NF.Movimento        >= '" & filtro.DataInicio.ToString("yyyy-MM-dd") & "'
			                AND NF.Movimento        <= '" & filtro.DataFim.ToString("yyyy-MM-dd") & "'
                            AND NCM IN ('23099010', '23091000')
			                AND SO.Classe IN ('VENDAS', 'VENDASAORDEM', 'EXPORTACOES', 'REMESSAS')
			                " & filtroSeguimento & "
                            AND LEFT(NF.Cliente_Id, 8) NOT IN ('49673784', '40938762')
	                )  

                    SELECT  ISNULL(SUM(CASE WHEN DEVOLUCAO = 'S' THEN -1 ELSE 1 END * PesoFiscal), 0) AS Quantidade,
                            ISNULL(SUM(CASE WHEN DEVOLUCAO = 'S' THEN -1 ELSE 1 END * ValorLiquido), 0) AS ValorTotal
                    FROM total_vendas; "

        Dim banco As New AcessaBanco
        Dim dsDashboard As DataSet = banco.ConsultaDataSet(sql, "NotasFiscais")

        Dim produto As New ProdutoDashboard

        For Each row As DataRow In dsDashboard.Tables(0).Rows

            produto.QuantidadeFiscal = row("Quantidade")
            produto.Valor = row("ValorTotal")

        Next

        Return produto

    End Function

    Public Function CarregarFaturamentoPorProduto(ByVal filtro As ProdutoDashboardFiltro) As IEnumerable(Of ProdutoDashboard)

        Dim sql As String = String.Empty

        Dim sqlCondicaoIN As String = ""

        If Not String.IsNullOrEmpty(filtro.EmpresasSelecionada) Then
            Dim listaFormatada = filtro.EmpresasSelecionada.Split(","c).
                                    Where(Function(s) Not String.IsNullOrWhiteSpace(s)).
                                    Select(Function(par) $"'{par.Trim()}'")

            sqlCondicaoIN = String.Join(",", listaFormatada)

        End If

        Dim filtroSeguimento As String = IIf(filtro.SeguimentoSelecionado <> "0", String.Format("AND PRO.Seguimento = {0} ", filtro.SeguimentoSelecionado), "")

        sql = " WITH total_vendas AS 
	                (
		                SELECT 	
			                NF.Empresa_Id,
			                NF.Cliente_Id,
			                NF.Serie_Id,
			                NF.Nota_Id,
			                FORMAT(NF.Movimento, 'dd/MM/yyyy') AS Data,
			                NFI.Produto_Id,
			                PRO.Nome AS ProdutoNome,
			                SEG.Descricao AS Seguimento,
			                NF.Operacao AS OperacaoNota,
			                NF.SubOperacao AS SubOperacaoNota,
			                NFI.QuantidadeFiscal,
			                NFI.PesoFiscal,
			                NFI.Unitario,
			                NFE_LIQUIDO.Valor AS ValorLiquido, --FORMAT(NFE_LIQUIDO.Valor, 'N2', 'pt-br') AS ValorLiquido,
			                NFE_PRO.Valor AS ValorTotal, --FORMAT(NFE_PRO.Valor, 'N2', 'pt-br') AS ValorTotal
                            SO.Devolucao
		                FROM NotasFiscais NF
		                INNER JOIN NotasFiscaisXItens NFI 
			                ON NF.Empresa_Id				= NFI.Empresa_Id 
			                AND NF.EndEmpresa_Id			= NFI.EndEmpresa_Id 
			                AND NF.Cliente_Id				= NFI.Cliente_Id 
			                AND NF.EndCliente_Id			= NFI.EndCliente_Id 
			                AND NF.EntradaSaida_Id			= NFI.EntradaSaida_Id 
			                AND NF.Serie_Id					= NFI.Serie_Id 
			                AND NF.Nota_Id					= NFI.Nota_Id
		                INNER JOIN NotasFiscaisXEncargos NFE_LIQUIDO
			                ON NF.Empresa_Id				= NFE_LIQUIDO.Empresa_Id 
			                AND NF.EndEmpresa_Id			= NFE_LIQUIDO.EndEmpresa_Id 
			                AND NF.Cliente_Id				= NFE_LIQUIDO.Cliente_Id 
			                AND NF.EndCliente_Id			= NFE_LIQUIDO.EndCliente_Id 
			                AND NF.EntradaSaida_Id			= NFE_LIQUIDO.EntradaSaida_Id 
			                AND NF.Serie_Id					= NFE_LIQUIDO.Serie_Id 
			                AND NF.Nota_Id					= NFE_LIQUIDO.Nota_Id
			                AND NFI.Produto_Id				= NFE_LIQUIDO.Produto_Id
			                AND NFI.Sequencia_Id			= NFE_LIQUIDO.Sequencia_id
			                AND NFE_LIQUIDO.Encargo_Id		= 'LIQUIDO'
		                INNER JOIN NotasFiscaisXEncargos NFE_PRO
			                ON NF.Empresa_Id				= NFE_PRO.Empresa_Id 
			                AND NF.EndEmpresa_Id			= NFE_PRO.EndEmpresa_Id 
			                AND NF.Cliente_Id				= NFE_PRO.Cliente_Id 
			                AND NF.EndCliente_Id			= NFE_PRO.EndCliente_Id 
			                AND NF.EntradaSaida_Id			= NFE_PRO.EntradaSaida_Id 
			                AND NF.Serie_Id					= NFE_PRO.Serie_Id 
			                AND NF.Nota_Id					= NFE_PRO.Nota_Id
			                AND NFI.Produto_Id				= NFE_PRO.Produto_Id
			                AND NFI.Sequencia_Id			= NFE_PRO.Sequencia_id
			                AND NFE_PRO.Encargo_Id			= 'PRODUTO'
		                INNER JOIN SubOperacoes SO
			                ON NF.Operacao = SO.Operacao_Id 
			                AND NF.SubOperacao = SO.SubOperacoes_Id
		                INNER JOIN Produtos PRO
			                ON NFI.Produto_Id = PRO.Produto_Id
		                INNER JOIN Seguimentos SEG
			                ON PRO.Seguimento = SEG.Seguimento_Id
                            AND PRO.Dashboard = 1
		                WHERE 
			                ISNULL(NF.Situacao, 1)  = 1
			                AND (NF.Empresa_Id + '|' + CONVERT(VARCHAR, NF.EndEmpresa_Id)) IN (" & IIf(String.IsNullOrEmpty(filtro.EmpresasSelecionada), "(NF.Empresa_Id + '|' + NF.EndEmpresa_Id)", sqlCondicaoIN) & ")
			                AND NF.Movimento        >= '" & filtro.DataInicio.ToString("yyyy-MM-dd") & "'
			                AND NF.Movimento        <= '" & filtro.DataFim.ToString("yyyy-MM-dd") & "'
                            AND NCM IN ('23099010', '23091000')
			                AND SO.Classe IN ('VENDAS', 'VENDASAORDEM', 'EXPORTACOES', 'REMESSAS')
			                " & filtroSeguimento & "
                            AND LEFT(NF.Cliente_Id, 8) NOT IN ('49673784', '40938762')
	                )  

                    SELECT  ProdutoNome,
                            Seguimento,
                            ISNULL(SUM(CASE WHEN DEVOLUCAO = 'S' THEN -1 ELSE 1 END * ValorLiquido), 0) AS ValorTotal
                    FROM total_vendas
                    GROUP BY
                        Empresa_Id,
                        Produto_Id,
                        ProdutoNome,
                        Seguimento
                    ORDER BY
                        SUM(CASE WHEN DEVOLUCAO = 'S' THEN -1 ELSE 1 END * ValorLiquido) DESC; "

        Dim banco As New AcessaBanco
        Dim dsDashboard As DataSet = banco.ConsultaDataSet(sql, "NotasFiscais")

        Dim listaProdutos As New List(Of ProdutoDashboard)

        For Each row As DataRow In dsDashboard.Tables(0).Rows

            Dim produto As New ProdutoDashboard

            produto.Nome = row("ProdutoNome")
            produto.Seguimento = row("Seguimento")
            produto.Valor = row("ValorTotal")

            listaProdutos.Add(produto)

        Next

        Return listaProdutos

    End Function

    Public Function CarregarDistribuicaoDeFaturamentoPorProduto(ByVal filtro As ProdutoDashboardFiltro) As IEnumerable(Of ProdutoDashboard)

        Dim sql As String = String.Empty

        Dim sqlCondicaoIN As String = ""

        If Not String.IsNullOrEmpty(filtro.EmpresasSelecionada) Then
            Dim listaFormatada = filtro.EmpresasSelecionada.Split(","c).
                                    Where(Function(s) Not String.IsNullOrWhiteSpace(s)).
                                    Select(Function(par) $"'{par.Trim()}'")

            sqlCondicaoIN = String.Join(",", listaFormatada)

        End If

        Dim filtroSeguimento As String = IIf(filtro.SeguimentoSelecionado <> "0", String.Format("AND PRO.Seguimento = {0} ", filtro.SeguimentoSelecionado), "")

        sql = " WITH total_vendas AS 
	                (
		                SELECT 	
			                NF.Empresa_Id,
			                NF.Cliente_Id,
			                NF.Serie_Id,
			                NF.Nota_Id,
			                FORMAT(NF.Movimento, 'dd/MM/yyyy') AS Data,
			                NFI.Produto_Id,
			                PRO.Nome AS ProdutoNome,
			                SEG.Descricao AS Seguimento,
			                NF.Operacao AS OperacaoNota,
			                NF.SubOperacao AS SubOperacaoNota,
			                NFI.QuantidadeFiscal,
			                NFI.PesoFiscal,
			                NFI.Unitario,
			                NFE_LIQUIDO.Valor AS ValorLiquido, --FORMAT(NFE_LIQUIDO.Valor, 'N2', 'pt-br') AS ValorLiquido,
			                NFE_PRO.Valor AS ValorTotal, --FORMAT(NFE_PRO.Valor, 'N2', 'pt-br') AS ValorTotal
                            SO.Devolucao
		                FROM NotasFiscais NF
		                INNER JOIN NotasFiscaisXItens NFI 
			                ON NF.Empresa_Id				= NFI.Empresa_Id 
			                AND NF.EndEmpresa_Id			= NFI.EndEmpresa_Id 
			                AND NF.Cliente_Id				= NFI.Cliente_Id 
			                AND NF.EndCliente_Id			= NFI.EndCliente_Id 
			                AND NF.EntradaSaida_Id			= NFI.EntradaSaida_Id 
			                AND NF.Serie_Id					= NFI.Serie_Id 
			                AND NF.Nota_Id					= NFI.Nota_Id
		                INNER JOIN NotasFiscaisXEncargos NFE_LIQUIDO
			                ON NF.Empresa_Id				= NFE_LIQUIDO.Empresa_Id 
			                AND NF.EndEmpresa_Id			= NFE_LIQUIDO.EndEmpresa_Id 
			                AND NF.Cliente_Id				= NFE_LIQUIDO.Cliente_Id 
			                AND NF.EndCliente_Id			= NFE_LIQUIDO.EndCliente_Id 
			                AND NF.EntradaSaida_Id			= NFE_LIQUIDO.EntradaSaida_Id 
			                AND NF.Serie_Id					= NFE_LIQUIDO.Serie_Id 
			                AND NF.Nota_Id					= NFE_LIQUIDO.Nota_Id
			                AND NFI.Produto_Id				= NFE_LIQUIDO.Produto_Id
			                AND NFI.Sequencia_Id			= NFE_LIQUIDO.Sequencia_id
			                AND NFE_LIQUIDO.Encargo_Id		= 'LIQUIDO'
		                INNER JOIN NotasFiscaisXEncargos NFE_PRO
			                ON NF.Empresa_Id				= NFE_PRO.Empresa_Id 
			                AND NF.EndEmpresa_Id			= NFE_PRO.EndEmpresa_Id 
			                AND NF.Cliente_Id				= NFE_PRO.Cliente_Id 
			                AND NF.EndCliente_Id			= NFE_PRO.EndCliente_Id 
			                AND NF.EntradaSaida_Id			= NFE_PRO.EntradaSaida_Id 
			                AND NF.Serie_Id					= NFE_PRO.Serie_Id 
			                AND NF.Nota_Id					= NFE_PRO.Nota_Id
			                AND NFI.Produto_Id				= NFE_PRO.Produto_Id
			                AND NFI.Sequencia_Id			= NFE_PRO.Sequencia_id
			                AND NFE_PRO.Encargo_Id			= 'PRODUTO'
		                INNER JOIN SubOperacoes SO
			                ON NF.Operacao = SO.Operacao_Id 
			                AND NF.SubOperacao = SO.SubOperacoes_Id
		                INNER JOIN Produtos PRO
			                ON NFI.Produto_Id = PRO.Produto_Id
		                INNER JOIN Seguimentos SEG
			                ON PRO.Seguimento = SEG.Seguimento_Id
                            AND PRO.Dashboard = 1
		                WHERE 
			                ISNULL(NF.Situacao, 1)  = 1
			                AND (NF.Empresa_Id + '|' + CONVERT(VARCHAR, NF.EndEmpresa_Id)) IN (" & IIf(String.IsNullOrEmpty(filtro.EmpresasSelecionada), "(NF.Empresa_Id + '|' + NF.EndEmpresa_Id)", sqlCondicaoIN) & ")
			                AND NF.Movimento        >= '" & filtro.DataInicio.ToString("yyyy-MM-dd") & "'
			                AND NF.Movimento        <= '" & filtro.DataFim.ToString("yyyy-MM-dd") & "'
                            AND NCM IN ('23099010', '23091000')
			                AND SO.Classe IN ('VENDAS', 'VENDASAORDEM', 'EXPORTACOES', 'REMESSAS')
			                " & filtroSeguimento & "
                            AND LEFT(NF.Cliente_Id, 8) NOT IN ('49673784', '40938762')
	                )  

                SELECT  Seguimento AS Seguimento,
	                    ISNULL(SUM(CASE WHEN DEVOLUCAO = 'S' THEN -1 ELSE 1 END * PesoFiscal), 0) AS Quantidade,
                        ISNULL(SUM(CASE WHEN DEVOLUCAO = 'S' THEN -1 ELSE 1 END * PesoFiscal), 0) AS PesoTotal,
                        ISNULL(SUM(CASE WHEN DEVOLUCAO = 'S' THEN -1 ELSE 1 END * ValorLiquido), 0) AS Valor
                FROM total_vendas
                GROUP BY
                    Seguimento
                ORDER BY
                    SUM(CASE WHEN DEVOLUCAO = 'S' THEN -1 ELSE 1 END * PesoFiscal) DESC; "

        Dim banco As New AcessaBanco
        Dim dsDashboard As DataSet = banco.ConsultaDataSet(sql, "NotasFiscais")

        Dim listaProdutos As New List(Of ProdutoDashboard)

        For Each row As DataRow In dsDashboard.Tables(0).Rows

            Dim produto As New ProdutoDashboard

            produto.Seguimento = row("Seguimento")
            produto.Valor = row("Valor")

            listaProdutos.Add(produto)

        Next

        Return listaProdutos

    End Function

    Public Function CarregarDetalhamentoProdutos(ByVal filtro As ProdutoDashboardFiltro) As IEnumerable(Of ProdutoDashboard)

        Dim sql As String = String.Empty

        Dim sqlCondicaoIN As String = ""

        If Not String.IsNullOrEmpty(filtro.EmpresasSelecionada) Then
            Dim listaFormatada = filtro.EmpresasSelecionada.Split(","c).
                                    Where(Function(s) Not String.IsNullOrWhiteSpace(s)).
                                    Select(Function(par) $"'{par.Trim()}'")

            sqlCondicaoIN = String.Join(",", listaFormatada)

        End If

        Dim filtroSeguimento As String = IIf(filtro.SeguimentoSelecionado <> "0", String.Format("AND PRO.Seguimento = {0} ", filtro.SeguimentoSelecionado), "")

        sql = " WITH total_vendas AS 
	                (
		                SELECT 	
			                NF.Empresa_Id,
			                NF.Cliente_Id,
			                NF.Serie_Id,
			                NF.Nota_Id,
			                FORMAT(NF.Movimento, 'dd/MM/yyyy') AS Data,
			                NFI.Produto_Id,
			                PRO.Nome AS ProdutoNome,
			                SEG.Descricao AS Seguimento,
			                NF.Operacao AS OperacaoNota,
			                NF.SubOperacao AS SubOperacaoNota,
			                NFI.QuantidadeFiscal,
			                NFI.PesoFiscal,
			                NFI.Unitario,
			                NFE_LIQUIDO.Valor AS ValorLiquido, --FORMAT(NFE_LIQUIDO.Valor, 'N2', 'pt-br') AS ValorLiquido,
			                NFE_PRO.Valor AS ValorTotal, --FORMAT(NFE_PRO.Valor, 'N2', 'pt-br') AS ValorTotal
                            SO.Devolucao
		                FROM NotasFiscais NF
		                INNER JOIN NotasFiscaisXItens NFI 
			                ON NF.Empresa_Id				= NFI.Empresa_Id 
			                AND NF.EndEmpresa_Id			= NFI.EndEmpresa_Id 
			                AND NF.Cliente_Id				= NFI.Cliente_Id 
			                AND NF.EndCliente_Id			= NFI.EndCliente_Id 
			                AND NF.EntradaSaida_Id			= NFI.EntradaSaida_Id 
			                AND NF.Serie_Id					= NFI.Serie_Id 
			                AND NF.Nota_Id					= NFI.Nota_Id
		                INNER JOIN NotasFiscaisXEncargos NFE_LIQUIDO
			                ON NF.Empresa_Id				= NFE_LIQUIDO.Empresa_Id 
			                AND NF.EndEmpresa_Id			= NFE_LIQUIDO.EndEmpresa_Id 
			                AND NF.Cliente_Id				= NFE_LIQUIDO.Cliente_Id 
			                AND NF.EndCliente_Id			= NFE_LIQUIDO.EndCliente_Id 
			                AND NF.EntradaSaida_Id			= NFE_LIQUIDO.EntradaSaida_Id 
			                AND NF.Serie_Id					= NFE_LIQUIDO.Serie_Id 
			                AND NF.Nota_Id					= NFE_LIQUIDO.Nota_Id
			                AND NFI.Produto_Id				= NFE_LIQUIDO.Produto_Id
			                AND NFI.Sequencia_Id			= NFE_LIQUIDO.Sequencia_id
			                AND NFE_LIQUIDO.Encargo_Id		= 'LIQUIDO'
		                INNER JOIN NotasFiscaisXEncargos NFE_PRO
			                ON NF.Empresa_Id				= NFE_PRO.Empresa_Id 
			                AND NF.EndEmpresa_Id			= NFE_PRO.EndEmpresa_Id 
			                AND NF.Cliente_Id				= NFE_PRO.Cliente_Id 
			                AND NF.EndCliente_Id			= NFE_PRO.EndCliente_Id 
			                AND NF.EntradaSaida_Id			= NFE_PRO.EntradaSaida_Id 
			                AND NF.Serie_Id					= NFE_PRO.Serie_Id 
			                AND NF.Nota_Id					= NFE_PRO.Nota_Id
			                AND NFI.Produto_Id				= NFE_PRO.Produto_Id
			                AND NFI.Sequencia_Id			= NFE_PRO.Sequencia_id
			                AND NFE_PRO.Encargo_Id			= 'PRODUTO'
		                INNER JOIN SubOperacoes SO
			                ON NF.Operacao = SO.Operacao_Id 
			                AND NF.SubOperacao = SO.SubOperacoes_Id
		                INNER JOIN Produtos PRO
			                ON NFI.Produto_Id = PRO.Produto_Id
		                INNER JOIN Seguimentos SEG
			                ON PRO.Seguimento = SEG.Seguimento_Id
                            AND PRO.Dashboard = 1
		                WHERE 
			                ISNULL(NF.Situacao, 1)  = 1
			                AND (NF.Empresa_Id + '|' + CONVERT(VARCHAR, NF.EndEmpresa_Id)) IN (" & IIf(String.IsNullOrEmpty(filtro.EmpresasSelecionada), "(NF.Empresa_Id + '|' + NF.EndEmpresa_Id)", sqlCondicaoIN) & ")
			                AND NF.Movimento        >= '" & filtro.DataInicio.ToString("yyyy-MM-dd") & "'
			                AND NF.Movimento        <= '" & filtro.DataFim.ToString("yyyy-MM-dd") & "'
                            AND NCM IN ('23099010', '23091000')
			                AND SO.Classe IN ('VENDAS', 'VENDASAORDEM', 'EXPORTACOES', 'REMESSAS')
			                " & filtroSeguimento & "
                            AND LEFT(NF.Cliente_Id, 8) NOT IN ('49673784', '40938762')
	                )  

                    SELECT  Empresa_Id,
                            Produto_Id,
                            ProdutoNome,
                            Seguimento,
	                        ISNULL(SUM(CASE WHEN DEVOLUCAO = 'S' THEN -1 ELSE 1 END * PesoFiscal), 0) AS Quantidade,
                            ISNULL(SUM(CASE WHEN DEVOLUCAO = 'S' THEN -1 ELSE 1 END * PesoFiscal), 0) AS PesoTotal,
                            ISNULL(SUM(CASE WHEN DEVOLUCAO = 'S' THEN -1 ELSE 1 END * ValorLiquido), 0) AS ValorTotal
                    FROM total_vendas
                    GROUP BY
                        Empresa_Id,
                        Produto_Id,
                        ProdutoNome,
                        Seguimento
                    ORDER BY
                        SUM(CASE WHEN DEVOLUCAO = 'S' THEN -1 ELSE 1 END * PesoFiscal) DESC; "

        Dim banco As New AcessaBanco
        Dim dsDashboard As DataSet = banco.ConsultaDataSet(sql, "NotasFiscais")

        Dim listaProdutos As New List(Of ProdutoDashboard)

        For Each row As DataRow In dsDashboard.Tables(0).Rows

            Dim produto As New ProdutoDashboard

            produto.EmpresaId = row("Empresa_Id")
            produto.ProdutoId = row("Produto_Id")
            produto.Nome = row("ProdutoNome")
            produto.Seguimento = row("Seguimento")
            produto.QuantidadeFiscal = row("Quantidade")
            produto.PesoFiscal = row("PesoTotal")
            produto.Valor = row("ValorTotal")

            listaProdutos.Add(produto)

        Next

        Return listaProdutos

    End Function

    'Clientes
    Public Function CarregarTotalProdutoPorClientes(ByVal filtro As ProdutoDashboardFiltro) As ProdutoDashboard

        Dim sql As String = String.Empty

        Dim sqlCondicaoIN As String = ""

        If Not String.IsNullOrEmpty(filtro.EmpresasSelecionada) Then
            Dim listaFormatada = filtro.EmpresasSelecionada.Split(","c).
                                    Where(Function(s) Not String.IsNullOrWhiteSpace(s)).
                                    Select(Function(par) $"'{par.Trim()}'")

            sqlCondicaoIN = String.Join(",", listaFormatada)

        End If

        Dim filtroSeguimento As String = IIf(filtro.SeguimentoSelecionado <> "0", String.Format("AND PRO.Seguimento = {0} ", filtro.SeguimentoSelecionado), "")

        sql = "WITH total_vendas AS 
                (
	                SELECT 	
		                NF.Empresa_Id,
		                NF.Cliente_Id,
		                NF.Serie_Id,
		                NF.Nota_Id,
		                FORMAT(NF.Movimento, 'dd/MM/yyyy') AS Data,
		                NFI.Produto_Id,
		                NFI.QuantidadeFiscal,
		                NFI.PesoFiscal,
		                NFI.Unitario,
		                NFE_LIQUIDO.Valor AS ValorLiquido,
		                NFE_PRO.Valor AS ValorTotal,
                        SO.Devolucao,
		                CLI.Cidade,
		                CLI.Estado
	                FROM NotasFiscais NF
	                INNER JOIN NotasFiscaisXItens NFI 
		                ON NF.Empresa_Id		= NFI.Empresa_Id 
		                AND NF.EndEmpresa_Id	= NFI.EndEmpresa_Id 
		                AND NF.Cliente_Id		= NFI.Cliente_Id 
		                AND NF.EndCliente_Id	= NFI.EndCliente_Id 
		                AND NF.EntradaSaida_Id	= NFI.EntradaSaida_Id 
		                AND NF.Serie_Id			= NFI.Serie_Id 
		                AND NF.Nota_Id			= NFI.Nota_Id
	                INNER JOIN NotasFiscaisXEncargos NFE_LIQUIDO
		                ON NF.Empresa_Id		= NFE_LIQUIDO.Empresa_Id 
		                AND NF.EndEmpresa_Id	= NFE_LIQUIDO.EndEmpresa_Id 
		                AND NF.Cliente_Id		= NFE_LIQUIDO.Cliente_Id 
		                AND NF.EndCliente_Id	= NFE_LIQUIDO.EndCliente_Id 
		                AND NF.EntradaSaida_Id	= NFE_LIQUIDO.EntradaSaida_Id 
		                AND NF.Serie_Id			= NFE_LIQUIDO.Serie_Id 
		                AND NF.Nota_Id			= NFE_LIQUIDO.Nota_Id
		                AND NFI.Produto_Id		= NFE_LIQUIDO.Produto_Id
		                AND NFI.Sequencia_Id	= NFE_LIQUIDO.Sequencia_id
		                AND NFE_LIQUIDO.Encargo_Id = 'LIQUIDO'
	                INNER JOIN NotasFiscaisXEncargos NFE_PRO
		                ON NF.Empresa_Id		= NFE_PRO.Empresa_Id 
		                AND NF.EndEmpresa_Id	= NFE_PRO.EndEmpresa_Id 
		                AND NF.Cliente_Id		= NFE_PRO.Cliente_Id 
		                AND NF.EndCliente_Id	= NFE_PRO.EndCliente_Id 
		                AND NF.EntradaSaida_Id	= NFE_PRO.EntradaSaida_Id 
		                AND NF.Serie_Id			= NFE_PRO.Serie_Id 
		                AND NF.Nota_Id			= NFE_PRO.Nota_Id
		                AND NFI.Produto_Id		= NFE_PRO.Produto_Id
		                AND NFI.Sequencia_Id	= NFE_PRO.Sequencia_id
		                AND NFE_PRO.Encargo_Id	= 'PRODUTO'
	                INNER JOIN SubOperacoes SO
		                ON NF.Operacao = SO.Operacao_Id 
		                AND NF.SubOperacao = SO.SubOperacoes_Id
	                INNER JOIN Produtos PRO
		                ON NFI.Produto_Id = PRO.Produto_Id
	                INNER JOIN Seguimentos SEG
			            ON PRO.Seguimento = SEG.Seguimento_Id
                        AND PRO.Dashboard = 1
	                INNER JOIN Clientes CLI
		                ON NF.Cliente_Id = CLI.Cliente_Id 
		                AND NF.EndCliente_Id = CLI.Endereco_Id
	                WHERE 
		                ISNULL(NF.Situacao, 1)  = 1
		                AND (NF.Empresa_Id + '|' + CONVERT(VARCHAR, NF.EndEmpresa_Id)) IN (" & IIf(String.IsNullOrEmpty(filtro.EmpresasSelecionada), "(NF.Empresa_Id + '|' + NF.EndEmpresa_Id)", sqlCondicaoIN) & ")
		                AND NF.Movimento        >= '" & filtro.DataInicio.ToString("yyyy-MM-dd") & "'
		                AND NF.Movimento        <= '" & filtro.DataFim.ToString("yyyy-MM-dd") & "'
                        AND NCM IN ('23099010', '23091000')
		                AND SO.Classe IN ('VENDAS', 'VENDASAORDEM', 'EXPORTACOES', 'REMESSAS')
		                " & filtroSeguimento & "
                        AND LEFT(NF.Cliente_Id, 8) NOT IN ('49673784', '40938762')
                )

                SELECT  ISNULL(SUM(CASE WHEN DEVOLUCAO = 'S' THEN -1 ELSE 1 END * PesoFiscal), 0) AS Quantidade,
                        ISNULL(SUM(CASE WHEN DEVOLUCAO = 'S' THEN -1 ELSE 1 END * ValorLiquido), 0) AS ValorTotal
                FROM total_vendas; "

        Dim banco As New AcessaBanco
        Dim dsDashboard As DataSet = banco.ConsultaDataSet(sql, "NotasFiscais")

        Dim produto As New ProdutoDashboard

        For Each row As DataRow In dsDashboard.Tables(0).Rows

            produto.QuantidadeFiscal = row("Quantidade")
            produto.Valor = row("ValorTotal")

        Next

        Return produto

    End Function

    Public Function CarregarFaturamentoPorClientes(ByVal filtro As ProdutoDashboardFiltro) As IEnumerable(Of ProdutoDashboard)

        Dim sql As String = String.Empty

        Dim sqlCondicaoIN As String = ""

        If Not String.IsNullOrEmpty(filtro.EmpresasSelecionada) Then
            Dim listaFormatada = filtro.EmpresasSelecionada.Split(","c).
                                    Where(Function(s) Not String.IsNullOrWhiteSpace(s)).
                                    Select(Function(par) $"'{par.Trim()}'")

            sqlCondicaoIN = String.Join(",", listaFormatada)

        End If

        Dim filtroSeguimento As String = IIf(filtro.SeguimentoSelecionado <> "0", String.Format("AND PRO.Seguimento = {0} ", filtro.SeguimentoSelecionado), "")

        sql = " WITH total_vendas AS 
                (
	                SELECT 	
		                NF.Empresa_Id,
		                NF.Cliente_Id,
		                NF.Serie_Id,
		                NF.Nota_Id,
		                FORMAT(NF.Movimento, 'dd/MM/yyyy') AS Data,
	                    CLI.Cliente_Id AS ClienteId,
	                    CLI.Nome AS ClienteNome,
	                    CLI.Fantasia AS ClienteFantasia,
		                NFI.Produto_Id,
		                NFI.QuantidadeFiscal,
		                NFI.PesoFiscal,
		                NFI.Unitario,
		                NFE_LIQUIDO.Valor AS ValorLiquido,
		                NFE_PRO.Valor AS ValorTotal,
                        SO.Devolucao,
		                CLI.Cidade,
		                CLI.Estado
	                FROM NotasFiscais NF
	                INNER JOIN NotasFiscaisXItens NFI 
		                ON NF.Empresa_Id		= NFI.Empresa_Id 
		                AND NF.EndEmpresa_Id	= NFI.EndEmpresa_Id 
		                AND NF.Cliente_Id		= NFI.Cliente_Id 
		                AND NF.EndCliente_Id	= NFI.EndCliente_Id 
		                AND NF.EntradaSaida_Id	= NFI.EntradaSaida_Id 
		                AND NF.Serie_Id			= NFI.Serie_Id 
		                AND NF.Nota_Id			= NFI.Nota_Id
	                INNER JOIN NotasFiscaisXEncargos NFE_LIQUIDO
		                ON NF.Empresa_Id		= NFE_LIQUIDO.Empresa_Id 
		                AND NF.EndEmpresa_Id	= NFE_LIQUIDO.EndEmpresa_Id 
		                AND NF.Cliente_Id		= NFE_LIQUIDO.Cliente_Id 
		                AND NF.EndCliente_Id	= NFE_LIQUIDO.EndCliente_Id 
		                AND NF.EntradaSaida_Id	= NFE_LIQUIDO.EntradaSaida_Id 
		                AND NF.Serie_Id			= NFE_LIQUIDO.Serie_Id 
		                AND NF.Nota_Id			= NFE_LIQUIDO.Nota_Id
		                AND NFI.Produto_Id		= NFE_LIQUIDO.Produto_Id
		                AND NFI.Sequencia_Id	= NFE_LIQUIDO.Sequencia_id
		                AND NFE_LIQUIDO.Encargo_Id = 'LIQUIDO'
	                INNER JOIN NotasFiscaisXEncargos NFE_PRO
		                ON NF.Empresa_Id		= NFE_PRO.Empresa_Id 
		                AND NF.EndEmpresa_Id	= NFE_PRO.EndEmpresa_Id 
		                AND NF.Cliente_Id		= NFE_PRO.Cliente_Id 
		                AND NF.EndCliente_Id	= NFE_PRO.EndCliente_Id 
		                AND NF.EntradaSaida_Id	= NFE_PRO.EntradaSaida_Id 
		                AND NF.Serie_Id			= NFE_PRO.Serie_Id 
		                AND NF.Nota_Id			= NFE_PRO.Nota_Id
		                AND NFI.Produto_Id		= NFE_PRO.Produto_Id
		                AND NFI.Sequencia_Id	= NFE_PRO.Sequencia_id
		                AND NFE_PRO.Encargo_Id	= 'PRODUTO'
	                INNER JOIN SubOperacoes SO
		                ON NF.Operacao = SO.Operacao_Id 
		                AND NF.SubOperacao = SO.SubOperacoes_Id
	                INNER JOIN Produtos PRO
		                ON NFI.Produto_Id = PRO.Produto_Id
	                INNER JOIN Seguimentos SEG
			            ON PRO.Seguimento = SEG.Seguimento_Id
                        AND PRO.Dashboard = 1
	                INNER JOIN Clientes CLI
		                ON NF.Cliente_Id = CLI.Cliente_Id 
		                AND NF.EndCliente_Id = CLI.Endereco_Id
	                WHERE 
		                ISNULL(NF.Situacao, 1)  = 1
		                AND (NF.Empresa_Id + '|' + CONVERT(VARCHAR, NF.EndEmpresa_Id)) IN (" & IIf(String.IsNullOrEmpty(filtro.EmpresasSelecionada), "(NF.Empresa_Id + '|' + NF.EndEmpresa_Id)", sqlCondicaoIN) & ")
		                AND NF.Movimento        >= '" & filtro.DataInicio.ToString("yyyy-MM-dd") & "'
		                AND NF.Movimento        <= '" & filtro.DataFim.ToString("yyyy-MM-dd") & "'
                        AND NCM IN ('23099010', '23091000')
		                AND SO.Classe IN ('VENDAS', 'VENDASAORDEM', 'EXPORTACOES', 'REMESSAS')
		                " & filtroSeguimento & "
                        AND LEFT(NF.Cliente_Id, 8) NOT IN ('49673784', '40938762')
                )   

                SELECT TOP 10 
                    ClienteId,
                    ClienteNome,
                    ClienteFantasia,
                    Cidade,
                    ISNULL(SUM(CASE WHEN DEVOLUCAO = 'S' THEN -1 ELSE 1 END * PesoFiscal), 0) AS Quantidade,
                    ISNULL(SUM(CASE WHEN DEVOLUCAO = 'S' THEN -1 ELSE 1 END * ValorLiquido), 0) AS ValorTotal
                FROM total_vendas
                GROUP BY
	                Empresa_Id,
                    ClienteId,
                    ClienteNome,
                    ClienteFantasia,
	                Estado,
                    Cidade
                ORDER BY
                    SUM(CASE WHEN DEVOLUCAO = 'S' THEN -1 ELSE 1 END * ValorLiquido) DESC; "

        Dim banco As New AcessaBanco
        Dim dsDashboard As DataSet = banco.ConsultaDataSet(sql, "NotasFiscais")

        Dim listaProdutos As New List(Of ProdutoDashboard)

        For Each row As DataRow In dsDashboard.Tables(0).Rows

            Dim produto As New ProdutoDashboard

            produto.ClienteId = row("ClienteId")
            produto.ClienteNome = row("ClienteNome")
            produto.ClienteFantasia = row("ClienteFantasia")
            produto.Cidade = row("Cidade")
            produto.QuantidadeFiscal = row("Quantidade")
            produto.Valor = row("ValorTotal")

            listaProdutos.Add(produto)

        Next

        Return listaProdutos

    End Function

    Public Function CarregarDetalhamentoClientes(ByVal filtro As ProdutoDashboardFiltro) As IEnumerable(Of ProdutoDashboard)

        Dim sql As String = String.Empty

        Dim sqlCondicaoIN As String = ""

        If Not String.IsNullOrEmpty(filtro.EmpresasSelecionada) Then
            Dim listaFormatada = filtro.EmpresasSelecionada.Split(","c).
                                    Where(Function(s) Not String.IsNullOrWhiteSpace(s)).
                                    Select(Function(par) $"'{par.Trim()}'")

            sqlCondicaoIN = String.Join(",", listaFormatada)

        End If

        Dim filtroSeguimento As String = IIf(filtro.SeguimentoSelecionado <> "0", String.Format("AND PRO.Seguimento = {0} ", filtro.SeguimentoSelecionado), "")

        sql = "WITH total_vendas AS 
                (
	                SELECT 	
		                NF.Empresa_Id,
		                NF.Cliente_Id,
		                NF.Serie_Id,
		                NF.Nota_Id,
		                FORMAT(NF.Movimento, 'dd/MM/yyyy') AS Data,
	                    CLI.Cliente_Id AS ClienteId,
	                    CLI.Nome AS ClienteNome,
	                    CLI.Fantasia AS ClienteFantasia,
		                NFI.Produto_Id,
		                NFI.QuantidadeFiscal,
		                NFI.PesoFiscal,
		                NFI.Unitario,
		                NFE_LIQUIDO.Valor AS ValorLiquido,
		                NFE_PRO.Valor AS ValorTotal,
                        SO.Devolucao,
		                CLI.Cidade,
		                CLI.Estado
	                FROM NotasFiscais NF
	                INNER JOIN NotasFiscaisXItens NFI 
		                ON NF.Empresa_Id		= NFI.Empresa_Id 
		                AND NF.EndEmpresa_Id	= NFI.EndEmpresa_Id 
		                AND NF.Cliente_Id		= NFI.Cliente_Id 
		                AND NF.EndCliente_Id	= NFI.EndCliente_Id 
		                AND NF.EntradaSaida_Id	= NFI.EntradaSaida_Id 
		                AND NF.Serie_Id			= NFI.Serie_Id 
		                AND NF.Nota_Id			= NFI.Nota_Id
	                INNER JOIN NotasFiscaisXEncargos NFE_LIQUIDO
		                ON NF.Empresa_Id		= NFE_LIQUIDO.Empresa_Id 
		                AND NF.EndEmpresa_Id	= NFE_LIQUIDO.EndEmpresa_Id 
		                AND NF.Cliente_Id		= NFE_LIQUIDO.Cliente_Id 
		                AND NF.EndCliente_Id	= NFE_LIQUIDO.EndCliente_Id 
		                AND NF.EntradaSaida_Id	= NFE_LIQUIDO.EntradaSaida_Id 
		                AND NF.Serie_Id			= NFE_LIQUIDO.Serie_Id 
		                AND NF.Nota_Id			= NFE_LIQUIDO.Nota_Id
		                AND NFI.Produto_Id		= NFE_LIQUIDO.Produto_Id
		                AND NFI.Sequencia_Id	= NFE_LIQUIDO.Sequencia_id
		                AND NFE_LIQUIDO.Encargo_Id = 'LIQUIDO'
	                INNER JOIN NotasFiscaisXEncargos NFE_PRO
		                ON NF.Empresa_Id		= NFE_PRO.Empresa_Id 
		                AND NF.EndEmpresa_Id	= NFE_PRO.EndEmpresa_Id 
		                AND NF.Cliente_Id		= NFE_PRO.Cliente_Id 
		                AND NF.EndCliente_Id	= NFE_PRO.EndCliente_Id 
		                AND NF.EntradaSaida_Id	= NFE_PRO.EntradaSaida_Id 
		                AND NF.Serie_Id			= NFE_PRO.Serie_Id 
		                AND NF.Nota_Id			= NFE_PRO.Nota_Id
		                AND NFI.Produto_Id		= NFE_PRO.Produto_Id
		                AND NFI.Sequencia_Id	= NFE_PRO.Sequencia_id
		                AND NFE_PRO.Encargo_Id	= 'PRODUTO'
	                INNER JOIN SubOperacoes SO
		                ON NF.Operacao = SO.Operacao_Id 
		                AND NF.SubOperacao = SO.SubOperacoes_Id
	                INNER JOIN Produtos PRO
		                ON NFI.Produto_Id = PRO.Produto_Id
	                INNER JOIN Seguimentos SEG
			            ON PRO.Seguimento = SEG.Seguimento_Id
                        AND PRO.Dashboard = 1
	                INNER JOIN Clientes CLI
		                ON NF.Cliente_Id = CLI.Cliente_Id 
		                AND NF.EndCliente_Id = CLI.Endereco_Id
	                WHERE 
		                ISNULL(NF.Situacao, 1)  = 1
		                AND (NF.Empresa_Id + '|' + CONVERT(VARCHAR, NF.EndEmpresa_Id)) IN (" & IIf(String.IsNullOrEmpty(filtro.EmpresasSelecionada), "(NF.Empresa_Id + '|' + NF.EndEmpresa_Id)", sqlCondicaoIN) & ")
		                AND NF.Movimento        >= '" & filtro.DataInicio.ToString("yyyy-MM-dd") & "'
		                AND NF.Movimento        <= '" & filtro.DataFim.ToString("yyyy-MM-dd") & "'
                        AND NCM IN ('23099010', '23091000')
		                AND SO.Classe IN ('VENDAS', 'VENDASAORDEM', 'EXPORTACOES', 'REMESSAS')
		                " & filtroSeguimento & "
                        AND LEFT(NF.Cliente_Id, 8) NOT IN ('49673784', '40938762')
                )  

                SELECT  
                    ClienteId,
                    ClienteNome,
                    ClienteFantasia,
	                Cidade,
	                Estado,
                    ISNULL(SUM(CASE WHEN DEVOLUCAO = 'S' THEN -1 ELSE 1 END * PesoFiscal), 0) AS Quantidade,
                    ISNULL(SUM(CASE WHEN DEVOLUCAO = 'S' THEN -1 ELSE 1 END * ValorLiquido), 0) AS ValorTotal
                FROM total_vendas
                GROUP BY
	                Empresa_Id,
                    ClienteId,
                    ClienteNome,
                    ClienteFantasia,
	                Estado,
                    Cidade
                ORDER BY
	                SUM(CASE WHEN DEVOLUCAO = 'S' THEN -1 ELSE 1 END * ValorLiquido) desc; "

        Dim banco As New AcessaBanco
        Dim dsDashboard As DataSet = banco.ConsultaDataSet(sql, "NotasFiscais")

        Dim listaProdutos As New List(Of ProdutoDashboard)

        For Each row As DataRow In dsDashboard.Tables(0).Rows

            Dim produto As New ProdutoDashboard

            'produto.EmpresaId = row("Empresa_Id")
            produto.ClienteId = row("ClienteId")
            produto.ClienteNome = row("ClienteNome")
            produto.ClienteFantasia = row("ClienteFantasia")
            produto.Cidade = row("Cidade")
            produto.Estado = row("Estado")
            produto.QuantidadeFiscal = row("Quantidade")
            produto.Valor = row("ValorTotal")

            listaProdutos.Add(produto)

        Next

        Return listaProdutos

    End Function

    'Representantes
    Public Function CarregarTotalProdutoPorRepresentantes(ByVal filtro As ProdutoDashboardFiltro) As ProdutoDashboard

        Dim sql As String = String.Empty

        Dim sqlCondicaoIN As String = ""

        If Not String.IsNullOrEmpty(filtro.EmpresasSelecionada) Then
            Dim listaFormatada = filtro.EmpresasSelecionada.Split(","c).
                                    Where(Function(s) Not String.IsNullOrWhiteSpace(s)).
                                    Select(Function(par) $"'{par.Trim()}'")

            sqlCondicaoIN = String.Join(",", listaFormatada)

        End If

        Dim filtroSeguimento As String = IIf(filtro.SeguimentoSelecionado <> "0", String.Format("AND PRO.Seguimento = {0} ", filtro.SeguimentoSelecionado), "")

        sql = " WITH total_vendas AS (
                SELECT 	
	                NF.Empresa_Id,
	                NF.Cliente_Id,
	                NF.Serie_Id,
	                NF.Nota_Id,
	                FORMAT(NF.Movimento, 'dd/MM/yyyy') AS Data,
	                COM.Representante_Id AS RepresentanteId,
	                C.Nome AS RepresentanteNome,
	                C.Fantasia AS RepresentanteFantasia,
	                NF.Operacao AS OperacaoNota,
	                NF.SubOperacao AS SubOperacaoNota,
	                NFI.QuantidadeFiscal,
	                NFI.PesoFiscal,
	                NFI.Unitario,
	                NFE_LIQUIDO.Valor AS ValorLiquido, --FORMAT(NFE_LIQUIDO.Valor, 'N2', 'pt-br') AS ValorLiquido,
	                NFE_PRO.Valor AS ValorTotal, --FORMAT(NFE_PRO.Valor, 'N2', 'pt-br') AS ValorTotal
                    SO.Devolucao
                FROM NotasFiscais NF
                INNER JOIN NotasFiscaisXItens NFI 
	                ON NF.Empresa_Id				= NFI.Empresa_Id 
	                AND NF.EndEmpresa_Id			= NFI.EndEmpresa_Id 
	                AND NF.Cliente_Id				= NFI.Cliente_Id 
	                AND NF.EndCliente_Id			= NFI.EndCliente_Id 
	                AND NF.EntradaSaida_Id			= NFI.EntradaSaida_Id 
	                AND NF.Serie_Id					= NFI.Serie_Id 
	                AND NF.Nota_Id					= NFI.Nota_Id
                INNER JOIN NotasFiscaisXEncargos NFE_LIQUIDO
	                ON NF.Empresa_Id				= NFE_LIQUIDO.Empresa_Id 
	                AND NF.EndEmpresa_Id			= NFE_LIQUIDO.EndEmpresa_Id 
	                AND NF.Cliente_Id				= NFE_LIQUIDO.Cliente_Id 
	                AND NF.EndCliente_Id			= NFE_LIQUIDO.EndCliente_Id 
	                AND NF.EntradaSaida_Id			= NFE_LIQUIDO.EntradaSaida_Id 
	                AND NF.Serie_Id					= NFE_LIQUIDO.Serie_Id 
	                AND NF.Nota_Id					= NFE_LIQUIDO.Nota_Id
	                AND NFI.Produto_Id				= NFE_LIQUIDO.Produto_Id
	                AND NFI.Sequencia_Id			= NFE_LIQUIDO.Sequencia_id
	                AND NFE_LIQUIDO.Encargo_Id		= 'LIQUIDO'
                INNER JOIN NotasFiscaisXEncargos NFE_PRO
	                ON NF.Empresa_Id				= NFE_PRO.Empresa_Id 
	                AND NF.EndEmpresa_Id			= NFE_PRO.EndEmpresa_Id 
	                AND NF.Cliente_Id				= NFE_PRO.Cliente_Id 
	                AND NF.EndCliente_Id			= NFE_PRO.EndCliente_Id 
	                AND NF.EntradaSaida_Id			= NFE_PRO.EntradaSaida_Id 
	                AND NF.Serie_Id					= NFE_PRO.Serie_Id 
	                AND NF.Nota_Id					= NFE_PRO.Nota_Id
	                AND NFI.Produto_Id				= NFE_PRO.Produto_Id
	                AND NFI.Sequencia_Id			= NFE_PRO.Sequencia_id
	                AND NFE_PRO.Encargo_Id			= 'PRODUTO'
                INNER JOIN SubOperacoes SO
	                ON NF.Operacao = SO.Operacao_Id 
	                AND NF.SubOperacao = SO.SubOperacoes_Id
                INNER JOIN Produtos PRO
	                ON NFI.Produto_Id = PRO.Produto_Id
                INNER JOIN Seguimentos SEG
			        ON PRO.Seguimento = SEG.Seguimento_Id
                    AND PRO.Dashboard = 1
                INNER JOIN Comissoes COM
	                ON COM.Pedido_Id = NF.Pedido
	                AND COM.Empresa_Id = NF.Empresa_Id
	                AND COM.EndEmpresa_Id = NF.EndEmpresa_Id
                INNER JOIN Clientes C
	                ON C.Cliente_Id = COM.Representante_Id
	                AND C.Endereco_Id = COM.EndRepresentante_Id
                WHERE 
	                ISNULL(NF.Situacao, 1)  = 1
			        AND (NF.Empresa_Id + '|' + CONVERT(VARCHAR, NF.EndEmpresa_Id)) IN (" & IIf(String.IsNullOrEmpty(filtro.EmpresasSelecionada), "(NF.Empresa_Id + '|' + NF.EndEmpresa_Id)", sqlCondicaoIN) & ")
			        AND NF.Movimento        >= '" & filtro.DataInicio.ToString("yyyy-MM-dd") & "'
			        AND NF.Movimento        <= '" & filtro.DataFim.ToString("yyyy-MM-dd") & "'
                    AND NCM IN ('23099010', '23091000')
	                AND SO.Classe IN ('VENDAS', 'VENDASAORDEM', 'EXPORTACOES', 'REMESSAS')
	                " & filtroSeguimento & "
                    AND LEFT(NF.Cliente_Id, 8) NOT IN ('49673784', '40938762')
                )

                SELECT ISNULL(SUM(CASE WHEN DEVOLUCAO = 'S' THEN -1 ELSE 1 END * PesoFiscal), 0) AS Quantidade,
                       ISNULL(SUM(CASE WHEN DEVOLUCAO = 'S' THEN -1 ELSE 1 END * ValorLiquido), 0) AS ValorTotal                    
                FROM total_vendas; "

        Dim banco As New AcessaBanco
        Dim dsDashboard As DataSet = banco.ConsultaDataSet(sql, "NotasFiscais")

        Dim produto As New ProdutoDashboard

        For Each row As DataRow In dsDashboard.Tables(0).Rows
            produto.QuantidadeFiscal = row("Quantidade")
            produto.Valor = row("ValorTotal")

        Next

        Return produto
    End Function

    Public Function CarregarFaturamentoPorRepresentantes(ByVal filtro As ProdutoDashboardFiltro) As IEnumerable(Of ProdutoDashboard)

        Dim sql As String = String.Empty

        Dim sqlCondicaoIN As String = ""

        If Not String.IsNullOrEmpty(filtro.EmpresasSelecionada) Then
            Dim listaFormatada = filtro.EmpresasSelecionada.Split(","c).
                                    Where(Function(s) Not String.IsNullOrWhiteSpace(s)).
                                    Select(Function(par) $"'{par.Trim()}'")

            sqlCondicaoIN = String.Join(",", listaFormatada)

        End If

        Dim filtroSeguimento As String = IIf(filtro.SeguimentoSelecionado <> "0", String.Format("AND PRO.Seguimento = {0} ", filtro.SeguimentoSelecionado), "")

        sql = " WITH total_vendas AS (
                SELECT 	
	                NF.Empresa_Id,
	                NF.Cliente_Id,
	                NF.Serie_Id,
	                NF.Nota_Id,
	                FORMAT(NF.Movimento, 'dd/MM/yyyy') AS Data,
	                COM.Representante_Id AS RepresentanteId,
	                C.Nome AS RepresentanteNome,
	                C.Fantasia AS RepresentanteFantasia,
	                NF.Operacao AS OperacaoNota,
	                NF.SubOperacao AS SubOperacaoNota,
	                NFI.QuantidadeFiscal,
	                NFI.PesoFiscal,
	                NFI.Unitario,
	                NFE_LIQUIDO.Valor AS ValorLiquido, --FORMAT(NFE_LIQUIDO.Valor, 'N2', 'pt-br') AS ValorLiquido,
	                NFE_PRO.Valor AS ValorTotal, --FORMAT(NFE_PRO.Valor, 'N2', 'pt-br') AS ValorTotal
                    SO.Devolucao
                FROM NotasFiscais NF
                INNER JOIN NotasFiscaisXItens NFI 
	                ON NF.Empresa_Id				= NFI.Empresa_Id 
	                AND NF.EndEmpresa_Id			= NFI.EndEmpresa_Id 
	                AND NF.Cliente_Id				= NFI.Cliente_Id 
	                AND NF.EndCliente_Id			= NFI.EndCliente_Id 
	                AND NF.EntradaSaida_Id			= NFI.EntradaSaida_Id 
	                AND NF.Serie_Id					= NFI.Serie_Id 
	                AND NF.Nota_Id					= NFI.Nota_Id
                INNER JOIN NotasFiscaisXEncargos NFE_LIQUIDO
	                ON NF.Empresa_Id				= NFE_LIQUIDO.Empresa_Id 
	                AND NF.EndEmpresa_Id			= NFE_LIQUIDO.EndEmpresa_Id 
	                AND NF.Cliente_Id				= NFE_LIQUIDO.Cliente_Id 
	                AND NF.EndCliente_Id			= NFE_LIQUIDO.EndCliente_Id 
	                AND NF.EntradaSaida_Id			= NFE_LIQUIDO.EntradaSaida_Id 
	                AND NF.Serie_Id					= NFE_LIQUIDO.Serie_Id 
	                AND NF.Nota_Id					= NFE_LIQUIDO.Nota_Id
	                AND NFI.Produto_Id				= NFE_LIQUIDO.Produto_Id
	                AND NFI.Sequencia_Id			= NFE_LIQUIDO.Sequencia_id
	                AND NFE_LIQUIDO.Encargo_Id		= 'LIQUIDO'
                INNER JOIN NotasFiscaisXEncargos NFE_PRO
	                ON NF.Empresa_Id				= NFE_PRO.Empresa_Id 
	                AND NF.EndEmpresa_Id			= NFE_PRO.EndEmpresa_Id 
	                AND NF.Cliente_Id				= NFE_PRO.Cliente_Id 
	                AND NF.EndCliente_Id			= NFE_PRO.EndCliente_Id 
	                AND NF.EntradaSaida_Id			= NFE_PRO.EntradaSaida_Id 
	                AND NF.Serie_Id					= NFE_PRO.Serie_Id 
	                AND NF.Nota_Id					= NFE_PRO.Nota_Id
	                AND NFI.Produto_Id				= NFE_PRO.Produto_Id
	                AND NFI.Sequencia_Id			= NFE_PRO.Sequencia_id
	                AND NFE_PRO.Encargo_Id			= 'PRODUTO'
                INNER JOIN SubOperacoes SO
	                ON NF.Operacao = SO.Operacao_Id 
	                AND NF.SubOperacao = SO.SubOperacoes_Id
                INNER JOIN Produtos PRO
	                ON NFI.Produto_Id = PRO.Produto_Id
                INNER JOIN Seguimentos SEG
			        ON PRO.Seguimento = SEG.Seguimento_Id
                    AND PRO.Dashboard = 1
                INNER JOIN Comissoes COM
	                ON COM.Pedido_Id = NF.Pedido
	                AND COM.Empresa_Id = NF.Empresa_Id
	                AND COM.EndEmpresa_Id = NF.EndEmpresa_Id
                INNER JOIN Clientes C
	                ON C.Cliente_Id = COM.Representante_Id
	                AND C.Endereco_Id = COM.EndRepresentante_Id
                WHERE 
	                ISNULL(NF.Situacao, 1)  = 1
			        AND (NF.Empresa_Id + '|' + CONVERT(VARCHAR, NF.EndEmpresa_Id)) IN (" & IIf(String.IsNullOrEmpty(filtro.EmpresasSelecionada), "(NF.Empresa_Id + '|' + NF.EndEmpresa_Id)", sqlCondicaoIN) & ")
			        AND NF.Movimento        >= '" & filtro.DataInicio.ToString("yyyy-MM-dd") & "'
			        AND NF.Movimento        <= '" & filtro.DataFim.ToString("yyyy-MM-dd") & "'
                    AND NCM IN ('23099010', '23091000')
	                AND SO.Classe IN ('VENDAS', 'VENDASAORDEM', 'EXPORTACOES', 'REMESSAS')
	                " & filtroSeguimento & "
                    AND LEFT(NF.Cliente_Id, 8) NOT IN ('49673784', '40938762')
                )

                SELECT TOP 10 total_vendas.RepresentanteId,
                       total_vendas.RepresentanteNome,
                       total_vendas.RepresentanteFantasia,
                       SUM(CASE WHEN DEVOLUCAO = 'S' THEN -1 ELSE 1 END * PesoFiscal) AS Quantidade,
                       SUM(CASE WHEN DEVOLUCAO = 'S' THEN -1 ELSE 1 END * ValorLiquido) AS ValorTotal                    
                FROM total_vendas
                GROUP BY RepresentanteId, RepresentanteNome, RepresentanteFantasia
                ORDER BY ValorTotal DESC, Quantidade DESC; "

        Dim banco As New AcessaBanco
        Dim dsDashboard As DataSet = banco.ConsultaDataSet(sql, "NotasFiscais")
        Dim listaProdutos As New List(Of ProdutoDashboard)

        For Each row As DataRow In dsDashboard.Tables(0).Rows
            Dim produto As New ProdutoDashboard

            produto.RepresentanteId = row("RepresentanteId")
            produto.RepresentanteNome = row("RepresentanteNome")
            produto.RepresentanteFantasia = row("RepresentanteFantasia")
            produto.QuantidadeFiscal = row("Quantidade")
            produto.Valor = row("ValorTotal")

            listaProdutos.Add(produto)
        Next

        Return listaProdutos

    End Function

    Public Function CarregarDetalhamentoRepresentantes(ByVal filtro As ProdutoDashboardFiltro) As IEnumerable(Of ProdutoDashboard)

        Dim sql As String = String.Empty

        Dim sqlCondicaoIN As String = ""

        If Not String.IsNullOrEmpty(filtro.EmpresasSelecionada) Then
            Dim listaFormatada = filtro.EmpresasSelecionada.Split(","c).
                                    Where(Function(s) Not String.IsNullOrWhiteSpace(s)).
                                    Select(Function(par) $"'{par.Trim()}'")

            sqlCondicaoIN = String.Join(",", listaFormatada)

        End If

        Dim filtroSeguimento As String = IIf(filtro.SeguimentoSelecionado <> "0", String.Format("AND PRO.Seguimento = {0} ", filtro.SeguimentoSelecionado), "")

        sql = " WITH total_vendas AS (
                SELECT 	
	                NF.Empresa_Id,
	                NF.Cliente_Id,
	                NF.Serie_Id,
	                NF.Nota_Id,
	                FORMAT(NF.Movimento, 'dd/MM/yyyy') AS Data,
	                COM.Representante_Id AS RepresentanteId,
	                C.Nome AS RepresentanteNome,
	                C.Fantasia AS RepresentanteFantasia,
	                NF.Operacao AS OperacaoNota,
	                NF.SubOperacao AS SubOperacaoNota,
	                NFI.QuantidadeFiscal,
	                NFI.PesoFiscal,
	                NFI.Unitario,
	                NFE_LIQUIDO.Valor AS ValorLiquido, --FORMAT(NFE_LIQUIDO.Valor, 'N2', 'pt-br') AS ValorLiquido,
	                NFE_PRO.Valor AS ValorTotal, --FORMAT(NFE_PRO.Valor, 'N2', 'pt-br') AS ValorTotal
                    SO.Devolucao
                FROM NotasFiscais NF
                INNER JOIN NotasFiscaisXItens NFI 
	                ON NF.Empresa_Id				= NFI.Empresa_Id 
	                AND NF.EndEmpresa_Id			= NFI.EndEmpresa_Id 
	                AND NF.Cliente_Id				= NFI.Cliente_Id 
	                AND NF.EndCliente_Id			= NFI.EndCliente_Id 
	                AND NF.EntradaSaida_Id			= NFI.EntradaSaida_Id 
	                AND NF.Serie_Id					= NFI.Serie_Id 
	                AND NF.Nota_Id					= NFI.Nota_Id
                INNER JOIN NotasFiscaisXEncargos NFE_LIQUIDO
	                ON NF.Empresa_Id				= NFE_LIQUIDO.Empresa_Id 
	                AND NF.EndEmpresa_Id			= NFE_LIQUIDO.EndEmpresa_Id 
	                AND NF.Cliente_Id				= NFE_LIQUIDO.Cliente_Id 
	                AND NF.EndCliente_Id			= NFE_LIQUIDO.EndCliente_Id 
	                AND NF.EntradaSaida_Id			= NFE_LIQUIDO.EntradaSaida_Id 
	                AND NF.Serie_Id					= NFE_LIQUIDO.Serie_Id 
	                AND NF.Nota_Id					= NFE_LIQUIDO.Nota_Id
	                AND NFI.Produto_Id				= NFE_LIQUIDO.Produto_Id
	                AND NFI.Sequencia_Id			= NFE_LIQUIDO.Sequencia_id
	                AND NFE_LIQUIDO.Encargo_Id		= 'LIQUIDO'
                INNER JOIN NotasFiscaisXEncargos NFE_PRO
	                ON NF.Empresa_Id				= NFE_PRO.Empresa_Id 
	                AND NF.EndEmpresa_Id			= NFE_PRO.EndEmpresa_Id 
	                AND NF.Cliente_Id				= NFE_PRO.Cliente_Id 
	                AND NF.EndCliente_Id			= NFE_PRO.EndCliente_Id 
	                AND NF.EntradaSaida_Id			= NFE_PRO.EntradaSaida_Id 
	                AND NF.Serie_Id					= NFE_PRO.Serie_Id 
	                AND NF.Nota_Id					= NFE_PRO.Nota_Id
	                AND NFI.Produto_Id				= NFE_PRO.Produto_Id
	                AND NFI.Sequencia_Id			= NFE_PRO.Sequencia_id
	                AND NFE_PRO.Encargo_Id			= 'PRODUTO'
                INNER JOIN SubOperacoes SO
	                ON NF.Operacao = SO.Operacao_Id 
	                AND NF.SubOperacao = SO.SubOperacoes_Id
                INNER JOIN Produtos PRO
	                ON NFI.Produto_Id = PRO.Produto_Id
                INNER JOIN Seguimentos SEG
			        ON PRO.Seguimento = SEG.Seguimento_Id
                    AND PRO.Dashboard = 1
                INNER JOIN Comissoes COM
	                ON COM.Pedido_Id = NF.Pedido
	                AND COM.Empresa_Id = NF.Empresa_Id
	                AND COM.EndEmpresa_Id = NF.EndEmpresa_Id
                INNER JOIN Clientes C
	                ON C.Cliente_Id = COM.Representante_Id
	                AND C.Endereco_Id = COM.EndRepresentante_Id
                WHERE 
	                ISNULL(NF.Situacao, 1)  = 1
			        AND (NF.Empresa_Id + '|' + CONVERT(VARCHAR, NF.EndEmpresa_Id)) IN (" & IIf(String.IsNullOrEmpty(filtro.EmpresasSelecionada), "(NF.Empresa_Id + '|' + NF.EndEmpresa_Id)", sqlCondicaoIN) & ")
			        AND NF.Movimento        >= '" & filtro.DataInicio.ToString("yyyy-MM-dd") & "'
			        AND NF.Movimento        <= '" & filtro.DataFim.ToString("yyyy-MM-dd") & "'
                    AND NCM IN ('23099010', '23091000')
	                AND SO.Classe IN ('VENDAS', 'VENDASAORDEM', 'EXPORTACOES', 'REMESSAS')
	                " & filtroSeguimento & "
                    AND LEFT(NF.Cliente_Id, 8) NOT IN ('49673784', '40938762')
                )

                SELECT total_vendas.RepresentanteId,
                       total_vendas.RepresentanteFantasia,
                       SUM(CASE WHEN DEVOLUCAO = 'S' THEN -1 ELSE 1 END * PesoFiscal) AS Quantidade,
                       SUM(CASE WHEN DEVOLUCAO = 'S' THEN -1 ELSE 1 END * PesoFiscal) AS PesoTotal,
                       SUM(CASE WHEN DEVOLUCAO = 'S' THEN -1 ELSE 1 END * ValorLiquido) AS ValorTotal                    
                FROM total_vendas
                GROUP BY RepresentanteId, RepresentanteNome, RepresentanteFantasia
                ORDER BY RepresentanteNome; "

        Dim banco As New AcessaBanco
        Dim dsDashboard As DataSet = banco.ConsultaDataSet(sql, "NotasFiscais")

        Dim listaProdutos As New List(Of ProdutoDashboard)

        For Each row As DataRow In dsDashboard.Tables(0).Rows

            Dim produto As New ProdutoDashboard

            produto.RepresentanteId = Funcoes.FormatarCpfCnpj(row("RepresentanteId"))
            produto.RepresentanteFantasia = row("RepresentanteFantasia")
            produto.QuantidadeFiscal = row("Quantidade")
            produto.PesoFiscal = row("PesoTotal")
            produto.Valor = row("ValorTotal")

            listaProdutos.Add(produto)

        Next

        Return listaProdutos

    End Function

    'Cidades
    Public Function CarregarTotalProdutoPorCidades(ByVal filtro As ProdutoDashboardFiltro) As ProdutoDashboard

        Dim sql As String = String.Empty

        Dim sqlCondicaoIN As String = ""

        If Not String.IsNullOrEmpty(filtro.EmpresasSelecionada) Then
            Dim listaFormatada = filtro.EmpresasSelecionada.Split(","c).
                                    Where(Function(s) Not String.IsNullOrWhiteSpace(s)).
                                    Select(Function(par) $"'{par.Trim()}'")

            sqlCondicaoIN = String.Join(",", listaFormatada)

        End If

        Dim filtroSeguimento As String = IIf(filtro.SeguimentoSelecionado <> "0", String.Format("AND PRO.Seguimento = {0} ", filtro.SeguimentoSelecionado), "")

        sql = "WITH total_vendas AS 
                (
	                SELECT 	
		                NF.Empresa_Id,
		                NF.Cliente_Id,
		                NF.Serie_Id,
		                NF.Nota_Id,
		                FORMAT(NF.Movimento, 'dd/MM/yyyy') AS Data,
		                NFI.Produto_Id,
		                NFI.QuantidadeFiscal,
		                NFI.PesoFiscal,
		                NFI.Unitario,
		                NFE_LIQUIDO.Valor AS ValorLiquido,
		                NFE_PRO.Valor AS ValorTotal,
                        SO.Devolucao,
		                CLI.Cidade,
		                CLI.Estado
	                FROM NotasFiscais NF
	                INNER JOIN NotasFiscaisXItens NFI 
		                ON NF.Empresa_Id		= NFI.Empresa_Id 
		                AND NF.EndEmpresa_Id	= NFI.EndEmpresa_Id 
		                AND NF.Cliente_Id		= NFI.Cliente_Id 
		                AND NF.EndCliente_Id	= NFI.EndCliente_Id 
		                AND NF.EntradaSaida_Id	= NFI.EntradaSaida_Id 
		                AND NF.Serie_Id			= NFI.Serie_Id 
		                AND NF.Nota_Id			= NFI.Nota_Id
	                INNER JOIN NotasFiscaisXEncargos NFE_LIQUIDO
		                ON NF.Empresa_Id		= NFE_LIQUIDO.Empresa_Id 
		                AND NF.EndEmpresa_Id	= NFE_LIQUIDO.EndEmpresa_Id 
		                AND NF.Cliente_Id		= NFE_LIQUIDO.Cliente_Id 
		                AND NF.EndCliente_Id	= NFE_LIQUIDO.EndCliente_Id 
		                AND NF.EntradaSaida_Id	= NFE_LIQUIDO.EntradaSaida_Id 
		                AND NF.Serie_Id			= NFE_LIQUIDO.Serie_Id 
		                AND NF.Nota_Id			= NFE_LIQUIDO.Nota_Id
		                AND NFI.Produto_Id		= NFE_LIQUIDO.Produto_Id
		                AND NFI.Sequencia_Id	= NFE_LIQUIDO.Sequencia_id
		                AND NFE_LIQUIDO.Encargo_Id = 'LIQUIDO'
	                INNER JOIN NotasFiscaisXEncargos NFE_PRO
		                ON NF.Empresa_Id		= NFE_PRO.Empresa_Id 
		                AND NF.EndEmpresa_Id	= NFE_PRO.EndEmpresa_Id 
		                AND NF.Cliente_Id		= NFE_PRO.Cliente_Id 
		                AND NF.EndCliente_Id	= NFE_PRO.EndCliente_Id 
		                AND NF.EntradaSaida_Id	= NFE_PRO.EntradaSaida_Id 
		                AND NF.Serie_Id			= NFE_PRO.Serie_Id 
		                AND NF.Nota_Id			= NFE_PRO.Nota_Id
		                AND NFI.Produto_Id		= NFE_PRO.Produto_Id
		                AND NFI.Sequencia_Id	= NFE_PRO.Sequencia_id
		                AND NFE_PRO.Encargo_Id	= 'PRODUTO'
	                INNER JOIN SubOperacoes SO
		                ON NF.Operacao = SO.Operacao_Id 
		                AND NF.SubOperacao = SO.SubOperacoes_Id
	                INNER JOIN Produtos PRO
		                ON NFI.Produto_Id = PRO.Produto_Id
	                INNER JOIN Seguimentos SEG
			            ON PRO.Seguimento = SEG.Seguimento_Id
                        AND PRO.Dashboard = 1
	                INNER JOIN Clientes CLI
		                ON NF.Cliente_Id = CLI.Cliente_Id 
		                AND NF.EndCliente_Id = CLI.Endereco_Id
	                WHERE 
		                ISNULL(NF.Situacao, 1)  = 1
		                AND (NF.Empresa_Id + '|' + CONVERT(VARCHAR, NF.EndEmpresa_Id)) IN (" & IIf(String.IsNullOrEmpty(filtro.EmpresasSelecionada), "(NF.Empresa_Id + '|' + NF.EndEmpresa_Id)", sqlCondicaoIN) & ")
		                AND NF.Movimento        >= '" & filtro.DataInicio.ToString("yyyy-MM-dd") & "'
		                AND NF.Movimento        <= '" & filtro.DataFim.ToString("yyyy-MM-dd") & "'
                        AND NCM IN ('23099010', '23091000')
		                AND SO.Classe IN ('VENDAS', 'VENDASAORDEM', 'EXPORTACOES', 'REMESSAS')
		                " & filtroSeguimento & "
                        AND LEFT(NF.Cliente_Id, 8) NOT IN ('49673784', '40938762')
                )

                SELECT  ISNULL(SUM(CASE WHEN DEVOLUCAO = 'S' THEN -1 ELSE 1 END * PesoFiscal), 0) AS Quantidade,
                        ISNULL(SUM(CASE WHEN DEVOLUCAO = 'S' THEN -1 ELSE 1 END * ValorLiquido), 0) AS ValorTotal
                FROM total_vendas; "

        Dim banco As New AcessaBanco
        Dim dsDashboard As DataSet = banco.ConsultaDataSet(sql, "NotasFiscais")

        Dim produto As New ProdutoDashboard

        For Each row As DataRow In dsDashboard.Tables(0).Rows

            produto.QuantidadeFiscal = row("Quantidade")
            produto.Valor = row("ValorTotal")

        Next

        Return produto

    End Function

    Public Function CarregarFaturamentoPorCidades(ByVal filtro As ProdutoDashboardFiltro) As IEnumerable(Of ProdutoDashboard)

        Dim sql As String = String.Empty

        Dim sqlCondicaoIN As String = ""

        If Not String.IsNullOrEmpty(filtro.EmpresasSelecionada) Then
            Dim listaFormatada = filtro.EmpresasSelecionada.Split(","c).
                                    Where(Function(s) Not String.IsNullOrWhiteSpace(s)).
                                    Select(Function(par) $"'{par.Trim()}'")

            sqlCondicaoIN = String.Join(",", listaFormatada)

        End If

        Dim filtroSeguimento As String = IIf(filtro.SeguimentoSelecionado <> "0", String.Format("AND PRO.Seguimento = {0} ", filtro.SeguimentoSelecionado), "")

        sql = " WITH total_vendas AS 
                (
	                SELECT 	
		                NF.Empresa_Id,
		                NF.Cliente_Id,
		                NF.Serie_Id,
		                NF.Nota_Id,
		                FORMAT(NF.Movimento, 'dd/MM/yyyy') AS Data,
		                NFI.Produto_Id,
		                NFI.QuantidadeFiscal,
		                NFI.PesoFiscal,
		                NFI.Unitario,
		                NFE_LIQUIDO.Valor AS ValorLiquido,
		                NFE_PRO.Valor AS ValorTotal,
                        SO.Devolucao,
		                CLI.Cidade,
		                CLI.Estado
	                FROM NotasFiscais NF
	                INNER JOIN NotasFiscaisXItens NFI 
		                ON NF.Empresa_Id		= NFI.Empresa_Id 
		                AND NF.EndEmpresa_Id	= NFI.EndEmpresa_Id 
		                AND NF.Cliente_Id		= NFI.Cliente_Id 
		                AND NF.EndCliente_Id	= NFI.EndCliente_Id 
		                AND NF.EntradaSaida_Id	= NFI.EntradaSaida_Id 
		                AND NF.Serie_Id			= NFI.Serie_Id 
		                AND NF.Nota_Id			= NFI.Nota_Id
	                INNER JOIN NotasFiscaisXEncargos NFE_LIQUIDO
		                ON NF.Empresa_Id		= NFE_LIQUIDO.Empresa_Id 
		                AND NF.EndEmpresa_Id	= NFE_LIQUIDO.EndEmpresa_Id 
		                AND NF.Cliente_Id		= NFE_LIQUIDO.Cliente_Id 
		                AND NF.EndCliente_Id	= NFE_LIQUIDO.EndCliente_Id 
		                AND NF.EntradaSaida_Id	= NFE_LIQUIDO.EntradaSaida_Id 
		                AND NF.Serie_Id			= NFE_LIQUIDO.Serie_Id 
		                AND NF.Nota_Id			= NFE_LIQUIDO.Nota_Id
		                AND NFI.Produto_Id		= NFE_LIQUIDO.Produto_Id
		                AND NFI.Sequencia_Id	= NFE_LIQUIDO.Sequencia_id
		                AND NFE_LIQUIDO.Encargo_Id = 'LIQUIDO'
	                INNER JOIN NotasFiscaisXEncargos NFE_PRO
		                ON NF.Empresa_Id		= NFE_PRO.Empresa_Id 
		                AND NF.EndEmpresa_Id	= NFE_PRO.EndEmpresa_Id 
		                AND NF.Cliente_Id		= NFE_PRO.Cliente_Id 
		                AND NF.EndCliente_Id	= NFE_PRO.EndCliente_Id 
		                AND NF.EntradaSaida_Id	= NFE_PRO.EntradaSaida_Id 
		                AND NF.Serie_Id			= NFE_PRO.Serie_Id 
		                AND NF.Nota_Id			= NFE_PRO.Nota_Id
		                AND NFI.Produto_Id		= NFE_PRO.Produto_Id
		                AND NFI.Sequencia_Id	= NFE_PRO.Sequencia_id
		                AND NFE_PRO.Encargo_Id	= 'PRODUTO'
	                INNER JOIN SubOperacoes SO
		                ON NF.Operacao = SO.Operacao_Id 
		                AND NF.SubOperacao = SO.SubOperacoes_Id
	                INNER JOIN Produtos PRO
		                ON NFI.Produto_Id = PRO.Produto_Id
	                INNER JOIN Seguimentos SEG
			            ON PRO.Seguimento = SEG.Seguimento_Id
                        AND PRO.Dashboard = 1
	                INNER JOIN Clientes CLI
		                ON NF.Cliente_Id = CLI.Cliente_Id 
		                AND NF.EndCliente_Id = CLI.Endereco_Id
	                WHERE 
		                ISNULL(NF.Situacao, 1)  = 1
		                AND (NF.Empresa_Id + '|' + CONVERT(VARCHAR, NF.EndEmpresa_Id)) IN (" & IIf(String.IsNullOrEmpty(filtro.EmpresasSelecionada), "(NF.Empresa_Id + '|' + NF.EndEmpresa_Id)", sqlCondicaoIN) & ")
		                AND NF.Movimento        >= '" & filtro.DataInicio.ToString("yyyy-MM-dd") & "'
		                AND NF.Movimento        <= '" & filtro.DataFim.ToString("yyyy-MM-dd") & "'
                        AND NCM IN ('23099010', '23091000')
		                AND SO.Classe IN ('VENDAS', 'VENDASAORDEM', 'EXPORTACOES', 'REMESSAS')
		                " & filtroSeguimento & "
                        AND LEFT(NF.Cliente_Id, 8) NOT IN ('49673784', '40938762')
                )   

                SELECT TOP 10 
                    Cidade,
                    SUM(CASE WHEN DEVOLUCAO = 'S' THEN -1 ELSE 1 END * PesoFiscal) AS PesoLiquidoKG,
                    SUM(CASE WHEN DEVOLUCAO = 'S' THEN -1 ELSE 1 END * ValorLiquido) AS ValorTotal
                FROM total_vendas
                GROUP BY
                    Empresa_Id,
                    Cidade
                ORDER BY
                    SUM(CASE WHEN DEVOLUCAO = 'S' THEN -1 ELSE 1 END * ValorLiquido) DESC; "

        Dim banco As New AcessaBanco
        Dim dsDashboard As DataSet = banco.ConsultaDataSet(sql, "NotasFiscais")

        Dim listaProdutos As New List(Of ProdutoDashboard)

        For Each row As DataRow In dsDashboard.Tables(0).Rows

            Dim produto As New ProdutoDashboard

            produto.Cidade = row("Cidade")
            produto.PesoFiscal = row("PesoLiquidoKG")
            produto.ValorTotal = row("ValorTotal")

            listaProdutos.Add(produto)

        Next

        Return listaProdutos

    End Function

    Public Function CarregarDetalhamentoCidades(ByVal filtro As ProdutoDashboardFiltro) As IEnumerable(Of ProdutoDashboard)

        Dim sql As String = String.Empty

        Dim sqlCondicaoIN As String = ""

        If Not String.IsNullOrEmpty(filtro.EmpresasSelecionada) Then
            Dim listaFormatada = filtro.EmpresasSelecionada.Split(","c).
                                    Where(Function(s) Not String.IsNullOrWhiteSpace(s)).
                                    Select(Function(par) $"'{par.Trim()}'")

            sqlCondicaoIN = String.Join(",", listaFormatada)

        End If

        Dim filtroSeguimento As String = IIf(filtro.SeguimentoSelecionado <> "0", String.Format("AND PRO.Seguimento = {0} ", filtro.SeguimentoSelecionado), "")

        sql = "WITH total_vendas AS 
                (
	                SELECT 	
		                NF.Empresa_Id,
		                NF.Cliente_Id,
		                NF.Serie_Id,
		                NF.Nota_Id,
		                FORMAT(NF.Movimento, 'dd/MM/yyyy') AS Data,
		                NFI.Produto_Id,
		                NFI.QuantidadeFiscal,
		                NFI.PesoFiscal,
		                NFI.Unitario,
		                NFE_LIQUIDO.Valor AS ValorLiquido,
		                NFE_PRO.Valor AS ValorTotal,
                        SO.Devolucao,
		                CLI.Cidade,
		                CLI.Estado
	                FROM NotasFiscais NF
	                INNER JOIN NotasFiscaisXItens NFI 
		                ON NF.Empresa_Id		= NFI.Empresa_Id 
		                AND NF.EndEmpresa_Id	= NFI.EndEmpresa_Id 
		                AND NF.Cliente_Id		= NFI.Cliente_Id 
		                AND NF.EndCliente_Id	= NFI.EndCliente_Id 
		                AND NF.EntradaSaida_Id	= NFI.EntradaSaida_Id 
		                AND NF.Serie_Id			= NFI.Serie_Id 
		                AND NF.Nota_Id			= NFI.Nota_Id
	                INNER JOIN NotasFiscaisXEncargos NFE_LIQUIDO
		                ON NF.Empresa_Id		= NFE_LIQUIDO.Empresa_Id 
		                AND NF.EndEmpresa_Id	= NFE_LIQUIDO.EndEmpresa_Id 
		                AND NF.Cliente_Id		= NFE_LIQUIDO.Cliente_Id 
		                AND NF.EndCliente_Id	= NFE_LIQUIDO.EndCliente_Id 
		                AND NF.EntradaSaida_Id	= NFE_LIQUIDO.EntradaSaida_Id 
		                AND NF.Serie_Id			= NFE_LIQUIDO.Serie_Id 
		                AND NF.Nota_Id			= NFE_LIQUIDO.Nota_Id
		                AND NFI.Produto_Id		= NFE_LIQUIDO.Produto_Id
		                AND NFI.Sequencia_Id	= NFE_LIQUIDO.Sequencia_id
		                AND NFE_LIQUIDO.Encargo_Id = 'LIQUIDO'
	                INNER JOIN NotasFiscaisXEncargos NFE_PRO
		                ON NF.Empresa_Id		= NFE_PRO.Empresa_Id 
		                AND NF.EndEmpresa_Id	= NFE_PRO.EndEmpresa_Id 
		                AND NF.Cliente_Id		= NFE_PRO.Cliente_Id 
		                AND NF.EndCliente_Id	= NFE_PRO.EndCliente_Id 
		                AND NF.EntradaSaida_Id	= NFE_PRO.EntradaSaida_Id 
		                AND NF.Serie_Id			= NFE_PRO.Serie_Id 
		                AND NF.Nota_Id			= NFE_PRO.Nota_Id
		                AND NFI.Produto_Id		= NFE_PRO.Produto_Id
		                AND NFI.Sequencia_Id	= NFE_PRO.Sequencia_id
		                AND NFE_PRO.Encargo_Id	= 'PRODUTO'
	                INNER JOIN SubOperacoes SO
		                ON NF.Operacao = SO.Operacao_Id 
		                AND NF.SubOperacao = SO.SubOperacoes_Id
	                INNER JOIN Produtos PRO
		                ON NFI.Produto_Id = PRO.Produto_Id
	                INNER JOIN Seguimentos SEG
			            ON PRO.Seguimento = SEG.Seguimento_Id
                        AND PRO.Dashboard = 1
	                INNER JOIN Clientes CLI
		                ON NF.Cliente_Id = CLI.Cliente_Id 
		                AND NF.EndCliente_Id = CLI.Endereco_Id
	                WHERE 
		                ISNULL(NF.Situacao, 1)  = 1
		                AND (NF.Empresa_Id + '|' + CONVERT(VARCHAR, NF.EndEmpresa_Id)) IN (" & IIf(String.IsNullOrEmpty(filtro.EmpresasSelecionada), "(NF.Empresa_Id + '|' + NF.EndEmpresa_Id)", sqlCondicaoIN) & ")
		                AND NF.Movimento        >= '" & filtro.DataInicio.ToString("yyyy-MM-dd") & "'
		                AND NF.Movimento        <= '" & filtro.DataFim.ToString("yyyy-MM-dd") & "'
                        AND NCM IN ('23099010', '23091000')
		                AND SO.Classe IN ('VENDAS', 'VENDASAORDEM', 'EXPORTACOES', 'REMESSAS')
		                " & filtroSeguimento & "
                        AND LEFT(NF.Cliente_Id, 8) NOT IN ('49673784', '40938762')
                )  

                SELECT  
	                Cidade,
	                Estado,
                    ISNULL(SUM(CASE WHEN DEVOLUCAO = 'S' THEN -1 ELSE 1 END * PesoFiscal), 0) AS Quantidade,
                    ISNULL(SUM(CASE WHEN DEVOLUCAO = 'S' THEN -1 ELSE 1 END * ValorLiquido), 0) AS ValorTotal
                FROM total_vendas
                GROUP BY
	                Empresa_Id,
	                Estado,
                    Cidade
                ORDER BY
	                SUM(CASE WHEN DEVOLUCAO = 'S' THEN -1 ELSE 1 END * ValorLiquido) desc; "

        Dim banco As New AcessaBanco
        Dim dsDashboard As DataSet = banco.ConsultaDataSet(sql, "NotasFiscais")

        Dim listaProdutos As New List(Of ProdutoDashboard)

        For Each row As DataRow In dsDashboard.Tables(0).Rows

            Dim produto As New ProdutoDashboard

            'produto.EmpresaId = row("Empresa_Id")
            produto.Cidade = row("Cidade")
            produto.Estado = row("Estado")
            produto.QuantidadeFiscal = row("Quantidade")
            produto.Valor = row("ValorTotal")

            listaProdutos.Add(produto)

        Next

        Return listaProdutos

    End Function

    'Estados
    Public Function CarregarTotalProdutoPorEstados(ByVal filtro As ProdutoDashboardFiltro) As ProdutoDashboard

        Dim sql As String = String.Empty

        Dim sqlCondicaoIN As String = ""

        If Not String.IsNullOrEmpty(filtro.EmpresasSelecionada) Then
            Dim listaFormatada = filtro.EmpresasSelecionada.Split(","c).
                                    Where(Function(s) Not String.IsNullOrWhiteSpace(s)).
                                    Select(Function(par) $"'{par.Trim()}'")

            sqlCondicaoIN = String.Join(",", listaFormatada)

        End If

        Dim filtroSeguimento As String = IIf(filtro.SeguimentoSelecionado <> "0", String.Format("AND PRO.Seguimento = {0} ", filtro.SeguimentoSelecionado), "")

        sql = "WITH total_vendas AS 
                (
	                SELECT 	
		                NF.Empresa_Id,
		                NF.Cliente_Id,
		                NF.Serie_Id,
		                NF.Nota_Id,
		                FORMAT(NF.Movimento, 'dd/MM/yyyy') AS Data,
		                NFI.Produto_Id,
		                NFI.QuantidadeFiscal,
		                NFI.PesoFiscal,
		                NFI.Unitario,
		                NFE_LIQUIDO.Valor AS ValorLiquido,
		                NFE_PRO.Valor AS ValorTotal,
                        SO.Devolucao,
		                CLI.Cidade,
		                CLI.Estado
	                FROM NotasFiscais NF
	                INNER JOIN NotasFiscaisXItens NFI 
		                ON NF.Empresa_Id		= NFI.Empresa_Id 
		                AND NF.EndEmpresa_Id	= NFI.EndEmpresa_Id 
		                AND NF.Cliente_Id		= NFI.Cliente_Id 
		                AND NF.EndCliente_Id	= NFI.EndCliente_Id 
		                AND NF.EntradaSaida_Id	= NFI.EntradaSaida_Id 
		                AND NF.Serie_Id			= NFI.Serie_Id 
		                AND NF.Nota_Id			= NFI.Nota_Id
	                INNER JOIN NotasFiscaisXEncargos NFE_LIQUIDO
		                ON NF.Empresa_Id		= NFE_LIQUIDO.Empresa_Id 
		                AND NF.EndEmpresa_Id	= NFE_LIQUIDO.EndEmpresa_Id 
		                AND NF.Cliente_Id		= NFE_LIQUIDO.Cliente_Id 
		                AND NF.EndCliente_Id	= NFE_LIQUIDO.EndCliente_Id 
		                AND NF.EntradaSaida_Id	= NFE_LIQUIDO.EntradaSaida_Id 
		                AND NF.Serie_Id			= NFE_LIQUIDO.Serie_Id 
		                AND NF.Nota_Id			= NFE_LIQUIDO.Nota_Id
		                AND NFI.Produto_Id		= NFE_LIQUIDO.Produto_Id
		                AND NFI.Sequencia_Id	= NFE_LIQUIDO.Sequencia_id
		                AND NFE_LIQUIDO.Encargo_Id = 'LIQUIDO'
	                INNER JOIN NotasFiscaisXEncargos NFE_PRO
		                ON NF.Empresa_Id		= NFE_PRO.Empresa_Id 
		                AND NF.EndEmpresa_Id	= NFE_PRO.EndEmpresa_Id 
		                AND NF.Cliente_Id		= NFE_PRO.Cliente_Id 
		                AND NF.EndCliente_Id	= NFE_PRO.EndCliente_Id 
		                AND NF.EntradaSaida_Id	= NFE_PRO.EntradaSaida_Id 
		                AND NF.Serie_Id			= NFE_PRO.Serie_Id 
		                AND NF.Nota_Id			= NFE_PRO.Nota_Id
		                AND NFI.Produto_Id		= NFE_PRO.Produto_Id
		                AND NFI.Sequencia_Id	= NFE_PRO.Sequencia_id
		                AND NFE_PRO.Encargo_Id	= 'PRODUTO'
	                INNER JOIN SubOperacoes SO
		                ON NF.Operacao = SO.Operacao_Id 
		                AND NF.SubOperacao = SO.SubOperacoes_Id
	                INNER JOIN Produtos PRO
		                ON NFI.Produto_Id = PRO.Produto_Id
	                INNER JOIN Seguimentos SEG
			            ON PRO.Seguimento = SEG.Seguimento_Id
                        AND PRO.Dashboard = 1
	                INNER JOIN Clientes CLI
		                ON NF.Cliente_Id = CLI.Cliente_Id 
		                AND NF.EndCliente_Id = CLI.Endereco_Id
	                WHERE 
		                ISNULL(NF.Situacao, 1)  = 1
		                AND (NF.Empresa_Id + '|' + CONVERT(VARCHAR, NF.EndEmpresa_Id)) IN (" & IIf(String.IsNullOrEmpty(filtro.EmpresasSelecionada), "(NF.Empresa_Id + '|' + NF.EndEmpresa_Id)", sqlCondicaoIN) & ")
		                AND NF.Movimento        >= '" & filtro.DataInicio.ToString("yyyy-MM-dd") & "'
		                AND NF.Movimento        <= '" & filtro.DataFim.ToString("yyyy-MM-dd") & "'
                        AND NCM IN ('23099010', '23091000')
		                AND SO.Classe IN ('VENDAS', 'VENDASAORDEM', 'EXPORTACOES', 'REMESSAS')
		                " & filtroSeguimento & "
                        AND LEFT(NF.Cliente_Id, 8) NOT IN ('49673784', '40938762')
                )

                SELECT  ISNULL(SUM(CASE WHEN DEVOLUCAO = 'S' THEN -1 ELSE 1 END * PesoFiscal), 0) AS Quantidade,
                        ISNULL(SUM(CASE WHEN DEVOLUCAO = 'S' THEN -1 ELSE 1 END * ValorLiquido), 0) AS ValorTotal
                FROM total_vendas;"

        Dim banco As New AcessaBanco
        Dim dsDashboard As DataSet = banco.ConsultaDataSet(sql, "NotasFiscais")

        Dim produto As New ProdutoDashboard

        For Each row As DataRow In dsDashboard.Tables(0).Rows

            produto.QuantidadeFiscal = row("Quantidade")
            produto.Valor = row("ValorTotal")

        Next

        Return produto

    End Function

    Public Function CarregarFaturamentoPorEstados(ByVal filtro As ProdutoDashboardFiltro) As IEnumerable(Of ProdutoDashboard)

        Dim sql As String = String.Empty

        Dim sqlCondicaoIN As String = ""

        If Not String.IsNullOrEmpty(filtro.EmpresasSelecionada) Then
            Dim listaFormatada = filtro.EmpresasSelecionada.Split(","c).
                                    Where(Function(s) Not String.IsNullOrWhiteSpace(s)).
                                    Select(Function(par) $"'{par.Trim()}'")

            sqlCondicaoIN = String.Join(",", listaFormatada)

        End If

        Dim filtroSeguimento As String = IIf(filtro.SeguimentoSelecionado <> "0", String.Format("AND PRO.Seguimento = {0} ", filtro.SeguimentoSelecionado), "")

        sql = "WITH total_vendas AS 
                (
	                SELECT 	
		                NF.Empresa_Id,
		                NF.Cliente_Id,
		                NF.Serie_Id,
		                NF.Nota_Id,
		                FORMAT(NF.Movimento, 'dd/MM/yyyy') AS Data,
		                NFI.Produto_Id,
		                NFI.QuantidadeFiscal,
		                NFI.PesoFiscal,
		                NFI.Unitario,
		                NFE_LIQUIDO.Valor AS ValorLiquido,
		                NFE_PRO.Valor AS ValorTotal,
                        SO.Devolucao,
		                CLI.Cidade,
		                CLI.Estado
	                FROM NotasFiscais NF
	                INNER JOIN NotasFiscaisXItens NFI 
		                ON NF.Empresa_Id		= NFI.Empresa_Id 
		                AND NF.EndEmpresa_Id	= NFI.EndEmpresa_Id 
		                AND NF.Cliente_Id		= NFI.Cliente_Id 
		                AND NF.EndCliente_Id	= NFI.EndCliente_Id 
		                AND NF.EntradaSaida_Id	= NFI.EntradaSaida_Id 
		                AND NF.Serie_Id			= NFI.Serie_Id 
		                AND NF.Nota_Id			= NFI.Nota_Id
	                INNER JOIN NotasFiscaisXEncargos NFE_LIQUIDO
		                ON NF.Empresa_Id		= NFE_LIQUIDO.Empresa_Id 
		                AND NF.EndEmpresa_Id	= NFE_LIQUIDO.EndEmpresa_Id 
		                AND NF.Cliente_Id		= NFE_LIQUIDO.Cliente_Id 
		                AND NF.EndCliente_Id	= NFE_LIQUIDO.EndCliente_Id 
		                AND NF.EntradaSaida_Id	= NFE_LIQUIDO.EntradaSaida_Id 
		                AND NF.Serie_Id			= NFE_LIQUIDO.Serie_Id 
		                AND NF.Nota_Id			= NFE_LIQUIDO.Nota_Id
		                AND NFI.Produto_Id		= NFE_LIQUIDO.Produto_Id
		                AND NFI.Sequencia_Id	= NFE_LIQUIDO.Sequencia_id
		                AND NFE_LIQUIDO.Encargo_Id = 'LIQUIDO'
	                INNER JOIN NotasFiscaisXEncargos NFE_PRO
		                ON NF.Empresa_Id		= NFE_PRO.Empresa_Id 
		                AND NF.EndEmpresa_Id	= NFE_PRO.EndEmpresa_Id 
		                AND NF.Cliente_Id		= NFE_PRO.Cliente_Id 
		                AND NF.EndCliente_Id	= NFE_PRO.EndCliente_Id 
		                AND NF.EntradaSaida_Id	= NFE_PRO.EntradaSaida_Id 
		                AND NF.Serie_Id			= NFE_PRO.Serie_Id 
		                AND NF.Nota_Id			= NFE_PRO.Nota_Id
		                AND NFI.Produto_Id		= NFE_PRO.Produto_Id
		                AND NFI.Sequencia_Id	= NFE_PRO.Sequencia_id
		                AND NFE_PRO.Encargo_Id	= 'PRODUTO'
	                INNER JOIN SubOperacoes SO
		                ON NF.Operacao = SO.Operacao_Id 
		                AND NF.SubOperacao = SO.SubOperacoes_Id
	                INNER JOIN Produtos PRO
		                ON NFI.Produto_Id = PRO.Produto_Id
	                INNER JOIN Seguimentos SEG
			            ON PRO.Seguimento = SEG.Seguimento_Id
                        AND PRO.Dashboard = 1
	                INNER JOIN Clientes CLI
		                ON NF.Cliente_Id = CLI.Cliente_Id 
		                AND NF.EndCliente_Id = CLI.Endereco_Id
	                WHERE 
		                ISNULL(NF.Situacao, 1)  = 1
		                AND (NF.Empresa_Id + '|' + CONVERT(VARCHAR, NF.EndEmpresa_Id)) IN (" & IIf(String.IsNullOrEmpty(filtro.EmpresasSelecionada), "(NF.Empresa_Id + '|' + NF.EndEmpresa_Id)", sqlCondicaoIN) & ")
		                AND NF.Movimento        >= '" & filtro.DataInicio.ToString("yyyy-MM-dd") & "'
		                AND NF.Movimento        <= '" & filtro.DataFim.ToString("yyyy-MM-dd") & "'
                        AND NCM IN ('23099010', '23091000')
		                AND SO.Classe IN ('VENDAS', 'VENDASAORDEM', 'EXPORTACOES', 'REMESSAS')
		                " & filtroSeguimento & "
                        AND LEFT(NF.Cliente_Id, 8) NOT IN ('49673784', '40938762')
                )   

                SELECT TOP 10 
                    Estado,
                    SUM(CASE WHEN DEVOLUCAO = 'S' THEN -1 ELSE 1 END * PesoFiscal) AS PesoLiquidoKG,
                    SUM(CASE WHEN DEVOLUCAO = 'S' THEN -1 ELSE 1 END * ValorLiquido) AS ValorTotal
                FROM total_vendas
                GROUP BY
                    Empresa_Id,
                    Estado
                ORDER BY
                    SUM(CASE WHEN DEVOLUCAO = 'S' THEN -1 ELSE 1 END * ValorLiquido) DESC;  "

        Dim banco As New AcessaBanco
        Dim dsDashboard As DataSet = banco.ConsultaDataSet(sql, "NotasFiscais")

        Dim listaProdutos As New List(Of ProdutoDashboard)

        For Each row As DataRow In dsDashboard.Tables(0).Rows

            Dim produto As New ProdutoDashboard

            produto.Estado = row("Estado")
            produto.PesoFiscal = row("PesoLiquidoKG")
            produto.ValorTotal = row("ValorTotal")

            listaProdutos.Add(produto)

        Next

        Return listaProdutos

    End Function

    Public Function CarregarDetalhamentoEstados(ByVal filtro As ProdutoDashboardFiltro) As IEnumerable(Of ProdutoDashboard)

        Dim sql As String = String.Empty

        Dim sqlCondicaoIN As String = ""

        If Not String.IsNullOrEmpty(filtro.EmpresasSelecionada) Then
            Dim listaFormatada = filtro.EmpresasSelecionada.Split(","c).
                                    Where(Function(s) Not String.IsNullOrWhiteSpace(s)).
                                    Select(Function(par) $"'{par.Trim()}'")

            sqlCondicaoIN = String.Join(",", listaFormatada)

        End If

        Dim filtroSeguimento As String = IIf(filtro.SeguimentoSelecionado <> "0", String.Format("AND PRO.Seguimento = {0} ", filtro.SeguimentoSelecionado), "")

        sql = "WITH total_vendas AS 
                (
	                SELECT 	
		                NF.Empresa_Id,
		                NF.Cliente_Id,
		                NF.Serie_Id,
		                NF.Nota_Id,
		                FORMAT(NF.Movimento, 'dd/MM/yyyy') AS Data,
		                NFI.Produto_Id,
		                NFI.QuantidadeFiscal,
		                NFI.PesoFiscal,
		                NFI.Unitario,
		                NFE_LIQUIDO.Valor AS ValorLiquido,
		                NFE_PRO.Valor AS ValorTotal,
                        SO.Devolucao,
		                CLI.Cidade,
		                CLI.Estado
	                FROM NotasFiscais NF
	                INNER JOIN NotasFiscaisXItens NFI 
		                ON NF.Empresa_Id		= NFI.Empresa_Id 
		                AND NF.EndEmpresa_Id	= NFI.EndEmpresa_Id 
		                AND NF.Cliente_Id		= NFI.Cliente_Id 
		                AND NF.EndCliente_Id	= NFI.EndCliente_Id 
		                AND NF.EntradaSaida_Id	= NFI.EntradaSaida_Id 
		                AND NF.Serie_Id			= NFI.Serie_Id 
		                AND NF.Nota_Id			= NFI.Nota_Id
	                INNER JOIN NotasFiscaisXEncargos NFE_LIQUIDO
		                ON NF.Empresa_Id		= NFE_LIQUIDO.Empresa_Id 
		                AND NF.EndEmpresa_Id	= NFE_LIQUIDO.EndEmpresa_Id 
		                AND NF.Cliente_Id		= NFE_LIQUIDO.Cliente_Id 
		                AND NF.EndCliente_Id	= NFE_LIQUIDO.EndCliente_Id 
		                AND NF.EntradaSaida_Id	= NFE_LIQUIDO.EntradaSaida_Id 
		                AND NF.Serie_Id			= NFE_LIQUIDO.Serie_Id 
		                AND NF.Nota_Id			= NFE_LIQUIDO.Nota_Id
		                AND NFI.Produto_Id		= NFE_LIQUIDO.Produto_Id
		                AND NFI.Sequencia_Id	= NFE_LIQUIDO.Sequencia_id
		                AND NFE_LIQUIDO.Encargo_Id = 'LIQUIDO'
	                INNER JOIN NotasFiscaisXEncargos NFE_PRO
		                ON NF.Empresa_Id		= NFE_PRO.Empresa_Id 
		                AND NF.EndEmpresa_Id	= NFE_PRO.EndEmpresa_Id 
		                AND NF.Cliente_Id		= NFE_PRO.Cliente_Id 
		                AND NF.EndCliente_Id	= NFE_PRO.EndCliente_Id 
		                AND NF.EntradaSaida_Id	= NFE_PRO.EntradaSaida_Id 
		                AND NF.Serie_Id			= NFE_PRO.Serie_Id 
		                AND NF.Nota_Id			= NFE_PRO.Nota_Id
		                AND NFI.Produto_Id		= NFE_PRO.Produto_Id
		                AND NFI.Sequencia_Id	= NFE_PRO.Sequencia_id
		                AND NFE_PRO.Encargo_Id	= 'PRODUTO'
	                INNER JOIN SubOperacoes SO
		                ON NF.Operacao = SO.Operacao_Id 
		                AND NF.SubOperacao = SO.SubOperacoes_Id
	                INNER JOIN Produtos PRO
		                ON NFI.Produto_Id = PRO.Produto_Id
	                INNER JOIN Seguimentos SEG
			            ON PRO.Seguimento = SEG.Seguimento_Id
                        AND PRO.Dashboard = 1
	                INNER JOIN Clientes CLI
		                ON NF.Cliente_Id = CLI.Cliente_Id 
		                AND NF.EndCliente_Id = CLI.Endereco_Id
	                WHERE 
		                ISNULL(NF.Situacao, 1)  = 1
		                AND (NF.Empresa_Id + '|' + CONVERT(VARCHAR, NF.EndEmpresa_Id)) IN (" & IIf(String.IsNullOrEmpty(filtro.EmpresasSelecionada), "(NF.Empresa_Id + '|' + NF.EndEmpresa_Id)", sqlCondicaoIN) & ")
		                AND NF.Movimento        >= '" & filtro.DataInicio.ToString("yyyy-MM-dd") & "'
		                AND NF.Movimento        <= '" & filtro.DataFim.ToString("yyyy-MM-dd") & "'
                        AND NCM IN ('23099010', '23091000')
		                AND SO.Classe IN ('VENDAS', 'VENDASAORDEM', 'EXPORTACOES', 'REMESSAS')
		                " & filtroSeguimento & "
                        AND LEFT(NF.Cliente_Id, 8) NOT IN ('49673784', '40938762')
                )  

                SELECT  
	                Estado,
                    ISNULL(SUM(CASE WHEN DEVOLUCAO = 'S' THEN -1 ELSE 1 END * PesoFiscal), 0) AS Quantidade,
                    ISNULL(SUM(CASE WHEN DEVOLUCAO = 'S' THEN -1 ELSE 1 END * ValorLiquido), 0) AS ValorTotal
                FROM total_vendas
                GROUP BY
	                Estado
                ORDER BY
	                SUM(CASE WHEN DEVOLUCAO = 'S' THEN -1 ELSE 1 END * ValorLiquido) DESC;"

        Dim banco As New AcessaBanco
        Dim dsDashboard As DataSet = banco.ConsultaDataSet(sql, "NotasFiscais")

        Dim listaProdutos As New List(Of ProdutoDashboard)

        For Each row As DataRow In dsDashboard.Tables(0).Rows

            Dim produto As New ProdutoDashboard

            'produto.EmpresaId = row("Empresa_Id")
            'produto.ProdutoId = row("Produto_Id")
            'produto.Nome = row("ProdutoNome")
            'produto.Seguimento = row("Seguimento")
            'produto.QuantidadeFiscal = row("Quantidade")
            produto.Estado = row("Estado")
            produto.QuantidadeFiscal = row("Quantidade")
            produto.Valor = row("ValorTotal")

            listaProdutos.Add(produto)

        Next

        Return listaProdutos

    End Function

    'Pedidos
    Public Function CarregarTotalProdutoPorPedidos(ByVal filtro As ProdutoDashboardFiltro) As ProdutoDashboard

        Dim sql As String = String.Empty

        Dim sqlCondicaoIN As String = ""

        If Not String.IsNullOrEmpty(filtro.EmpresasSelecionada) Then
            Dim listaFormatada = filtro.EmpresasSelecionada.Split(","c).
                                    Where(Function(s) Not String.IsNullOrWhiteSpace(s)).
                                    Select(Function(par) $"'{par.Trim()}'")

            sqlCondicaoIN = String.Join(",", listaFormatada)

        End If

        Dim filtroSeguimento As String = IIf(filtro.SeguimentoSelecionado <> "0", String.Format("AND PRO.Seguimento = {0} ", filtro.SeguimentoSelecionado), "")

        sql = "WITH total_vendas AS 
                (
	                SELECT 	
		                NF.Empresa_Id,
		                NF.Cliente_Id,
		                NF.Serie_Id,
		                NF.Nota_Id,
                        NF.Pedido,
		                FORMAT(NF.Movimento, 'dd/MM/yyyy') AS Data,
		                NFI.Produto_Id,
		                NFI.QuantidadeFiscal,
		                NFI.PesoFiscal,
		                NFI.Unitario,
		                NFE_LIQUIDO.Valor AS ValorLiquido,
		                NFE_PRO.Valor AS ValorTotal,
                        SO.Devolucao,
		                CLI.Cidade,
		                CLI.Estado
	                FROM NotasFiscais NF
	                INNER JOIN NotasFiscaisXItens NFI 
		                ON NF.Empresa_Id		= NFI.Empresa_Id 
		                AND NF.EndEmpresa_Id	= NFI.EndEmpresa_Id 
		                AND NF.Cliente_Id		= NFI.Cliente_Id 
		                AND NF.EndCliente_Id	= NFI.EndCliente_Id 
		                AND NF.EntradaSaida_Id	= NFI.EntradaSaida_Id 
		                AND NF.Serie_Id			= NFI.Serie_Id 
		                AND NF.Nota_Id			= NFI.Nota_Id
	                INNER JOIN NotasFiscaisXEncargos NFE_LIQUIDO
		                ON NF.Empresa_Id		= NFE_LIQUIDO.Empresa_Id 
		                AND NF.EndEmpresa_Id	= NFE_LIQUIDO.EndEmpresa_Id 
		                AND NF.Cliente_Id		= NFE_LIQUIDO.Cliente_Id 
		                AND NF.EndCliente_Id	= NFE_LIQUIDO.EndCliente_Id 
		                AND NF.EntradaSaida_Id	= NFE_LIQUIDO.EntradaSaida_Id 
		                AND NF.Serie_Id			= NFE_LIQUIDO.Serie_Id 
		                AND NF.Nota_Id			= NFE_LIQUIDO.Nota_Id
		                AND NFI.Produto_Id		= NFE_LIQUIDO.Produto_Id
		                AND NFI.Sequencia_Id	= NFE_LIQUIDO.Sequencia_id
		                AND NFE_LIQUIDO.Encargo_Id = 'LIQUIDO'
	                INNER JOIN NotasFiscaisXEncargos NFE_PRO
		                ON NF.Empresa_Id		= NFE_PRO.Empresa_Id 
		                AND NF.EndEmpresa_Id	= NFE_PRO.EndEmpresa_Id 
		                AND NF.Cliente_Id		= NFE_PRO.Cliente_Id 
		                AND NF.EndCliente_Id	= NFE_PRO.EndCliente_Id 
		                AND NF.EntradaSaida_Id	= NFE_PRO.EntradaSaida_Id 
		                AND NF.Serie_Id			= NFE_PRO.Serie_Id 
		                AND NF.Nota_Id			= NFE_PRO.Nota_Id
		                AND NFI.Produto_Id		= NFE_PRO.Produto_Id
		                AND NFI.Sequencia_Id	= NFE_PRO.Sequencia_id
		                AND NFE_PRO.Encargo_Id	= 'PRODUTO'
	                INNER JOIN SubOperacoes SO
		                ON NF.Operacao = SO.Operacao_Id 
		                AND NF.SubOperacao = SO.SubOperacoes_Id
	                INNER JOIN Produtos PRO
		                ON NFI.Produto_Id = PRO.Produto_Id
	                INNER JOIN Seguimentos SEG
			            ON PRO.Seguimento = SEG.Seguimento_Id
                        AND PRO.Dashboard = 1
	                INNER JOIN Clientes CLI
		                ON NF.Cliente_Id = CLI.Cliente_Id 
		                AND NF.EndCliente_Id = CLI.Endereco_Id
	                WHERE 
		                ISNULL(NF.Situacao, 1)  = 1
		                AND (NF.Empresa_Id + '|' + CONVERT(VARCHAR, NF.EndEmpresa_Id)) IN (" & IIf(String.IsNullOrEmpty(filtro.EmpresasSelecionada), "(NF.Empresa_Id + '|' + NF.EndEmpresa_Id)", sqlCondicaoIN) & ")
		                AND NF.Movimento        >= '" & filtro.DataInicio.ToString("yyyy-MM-dd") & "'
		                AND NF.Movimento        <= '" & filtro.DataFim.ToString("yyyy-MM-dd") & "'
                        AND NCM IN ('23099010', '23091000')
		                AND SO.Classe IN ('VENDAS', 'VENDASAORDEM', 'EXPORTACOES', 'REMESSAS')
		                " & filtroSeguimento & "
                        AND LEFT(NF.Cliente_Id, 8) NOT IN ('49673784', '40938762')
                ),
                total_vendas_10 AS
                (
                    SELECT TOP (10)
                        Cliente_Id,
                        Pedido,
                        ISNULL(SUM(CASE WHEN Devolucao = 'S' THEN -PesoFiscal       ELSE PesoFiscal       END), 0) AS PesoFiscal,
                        ISNULL(SUM(CASE WHEN Devolucao = 'S' THEN -QuantidadeFiscal ELSE QuantidadeFiscal END), 0) AS Quantidade,
                        ISNULL(SUM(CASE WHEN Devolucao = 'S' THEN -ValorLiquido     ELSE ValorLiquido     END), 0) AS ValorTotal
                    FROM total_vendas
                    GROUP BY
                        Cliente_Id,
                        Pedido
                    ORDER BY
                        SUM(CASE WHEN Devolucao = 'S' THEN -ValorLiquido ELSE ValorLiquido END) DESC
                )
                SELECT
                    ISNULL(SUM(PesoFiscal), 0) AS Quantidade,   -- se preferir, troque para SUM(Quantidade)
                    ISNULL(SUM(ValorTotal), 0) AS ValorTotal
                FROM total_vendas_10; "

        Dim banco As New AcessaBanco
        Dim dsDashboard As DataSet = banco.ConsultaDataSet(sql, "NotasFiscais")

        Dim produto As New ProdutoDashboard

        For Each row As DataRow In dsDashboard.Tables(0).Rows

            produto.QuantidadeFiscal = row("Quantidade")
            produto.Valor = row("ValorTotal")

        Next

        Return produto

    End Function

    Public Function CarregarFaturamentoPorPedidos(ByVal filtro As ProdutoDashboardFiltro) As IEnumerable(Of ProdutoDashboard)

        Dim sql As String = String.Empty

        Dim sqlCondicaoIN As String = ""

        If Not String.IsNullOrEmpty(filtro.EmpresasSelecionada) Then
            Dim listaFormatada = filtro.EmpresasSelecionada.Split(","c).
                                    Where(Function(s) Not String.IsNullOrWhiteSpace(s)).
                                    Select(Function(par) $"'{par.Trim()}'")

            sqlCondicaoIN = String.Join(",", listaFormatada)

        End If

        Dim filtroSeguimento As String = IIf(filtro.SeguimentoSelecionado <> "0", String.Format("AND PRO.Seguimento = {0} ", filtro.SeguimentoSelecionado), "")

        sql = "WITH total_vendas AS 
                (
	                SELECT 	
		                NF.Empresa_Id,
		                NF.Cliente_Id,
		                NF.Serie_Id,
		                NF.Nota_Id,
                        NF.Pedido,
                        PE.DataPedido,
		                FORMAT(NF.Movimento, 'dd/MM/yyyy') AS Data,
	                    CLI.Cliente_Id AS ClienteId,
	                    CLI.Nome AS ClienteNome,
	                    CLI.Fantasia AS ClienteFantasia,
		                NFI.Produto_Id,
		                NFI.QuantidadeFiscal,
		                NFI.PesoFiscal,
		                NFI.Unitario,
		                NFE_LIQUIDO.Valor AS ValorLiquido,
		                NFE_PRO.Valor AS ValorTotal,
                        SO.Devolucao,
		                CLI.Cidade,
		                CLI.Estado
	                FROM NotasFiscais NF
	                INNER JOIN NotasFiscaisXItens NFI 
		                ON NF.Empresa_Id		= NFI.Empresa_Id 
		                AND NF.EndEmpresa_Id	= NFI.EndEmpresa_Id 
		                AND NF.Cliente_Id		= NFI.Cliente_Id 
		                AND NF.EndCliente_Id	= NFI.EndCliente_Id 
		                AND NF.EntradaSaida_Id	= NFI.EntradaSaida_Id 
		                AND NF.Serie_Id			= NFI.Serie_Id 
		                AND NF.Nota_Id			= NFI.Nota_Id
	                INNER JOIN NotasFiscaisXEncargos NFE_LIQUIDO
		                ON NF.Empresa_Id		= NFE_LIQUIDO.Empresa_Id 
		                AND NF.EndEmpresa_Id	= NFE_LIQUIDO.EndEmpresa_Id 
		                AND NF.Cliente_Id		= NFE_LIQUIDO.Cliente_Id 
		                AND NF.EndCliente_Id	= NFE_LIQUIDO.EndCliente_Id 
		                AND NF.EntradaSaida_Id	= NFE_LIQUIDO.EntradaSaida_Id 
		                AND NF.Serie_Id			= NFE_LIQUIDO.Serie_Id 
		                AND NF.Nota_Id			= NFE_LIQUIDO.Nota_Id
		                AND NFI.Produto_Id		= NFE_LIQUIDO.Produto_Id
		                AND NFI.Sequencia_Id	= NFE_LIQUIDO.Sequencia_id
		                AND NFE_LIQUIDO.Encargo_Id = 'LIQUIDO'
	                INNER JOIN NotasFiscaisXEncargos NFE_PRO
		                ON NF.Empresa_Id		= NFE_PRO.Empresa_Id 
		                AND NF.EndEmpresa_Id	= NFE_PRO.EndEmpresa_Id 
		                AND NF.Cliente_Id		= NFE_PRO.Cliente_Id 
		                AND NF.EndCliente_Id	= NFE_PRO.EndCliente_Id 
		                AND NF.EntradaSaida_Id	= NFE_PRO.EntradaSaida_Id 
		                AND NF.Serie_Id			= NFE_PRO.Serie_Id 
		                AND NF.Nota_Id			= NFE_PRO.Nota_Id
		                AND NFI.Produto_Id		= NFE_PRO.Produto_Id
		                AND NFI.Sequencia_Id	= NFE_PRO.Sequencia_id
		                AND NFE_PRO.Encargo_Id	= 'PRODUTO'
	                INNER JOIN SubOperacoes SO
		                ON NF.Operacao          = SO.Operacao_Id 
		                AND NF.SubOperacao      = SO.SubOperacoes_Id
	                INNER JOIN Produtos PRO
		                ON NFI.Produto_Id       = PRO.Produto_Id
	                INNER JOIN Seguimentos SEG
			            ON PRO.Seguimento = SEG.Seguimento_Id
                        AND PRO.Dashboard = 1
	                INNER JOIN Clientes CLI
		                ON NF.Cliente_Id        = CLI.Cliente_Id 
		                AND NF.EndCliente_Id    = CLI.Endereco_Id
	                INNER JOIN Pedidos PE
		                ON NF.Empresa_Id        = PE.Empresa_Id 
		                AND NF.EndEmpresa_Id    = PE.EndEmpresa_Id
                        AND NF.Pedido           = PE.Pedido_Id
	                WHERE 
		                ISNULL(NF.Situacao, 1)  = 1
		                AND (NF.Empresa_Id + '|' + CONVERT(VARCHAR, NF.EndEmpresa_Id)) IN (" & IIf(String.IsNullOrEmpty(filtro.EmpresasSelecionada), "(NF.Empresa_Id + '|' + NF.EndEmpresa_Id)", sqlCondicaoIN) & ")
		                AND NF.Movimento        >= '" & filtro.DataInicio.ToString("yyyy-MM-dd") & "'
		                AND NF.Movimento        <= '" & filtro.DataFim.ToString("yyyy-MM-dd") & "'
                        AND NCM IN ('23099010', '23091000')
		                AND SO.Classe IN ('VENDAS', 'VENDASAORDEM', 'EXPORTACOES', 'REMESSAS')
		                " & filtroSeguimento & "
                        AND LEFT(NF.Cliente_Id, 8) NOT IN ('49673784', '40938762')
                )  

                SELECT TOP 10
                ClienteId,
                ClienteNome,
                ClienteFantasia,
                FORMAT(DataPedido, 'dd/MM/yyyy') AS Data,
                Pedido,
                ISNULL(SUM(CASE WHEN DEVOLUCAO = 'S' THEN -1 ELSE 1 END * PesoFiscal), 0) AS PesoFiscal,
                ISNULL(SUM(CASE WHEN DEVOLUCAO = 'S' THEN -1 ELSE 1 END * PesoFiscal), 0) AS Quantidade,
                ISNULL(SUM(CASE WHEN DEVOLUCAO = 'S' THEN -1 ELSE 1 END * ValorLiquido), 0) AS ValorTotal
            FROM total_vendas
            GROUP BY
                ClienteId,
                ClienteNome,
                ClienteFantasia,
                DataPedido,
                Pedido
            ORDER BY
                SUM(CASE WHEN DEVOLUCAO = 'S' THEN -1 ELSE 1 END * ValorLiquido) DESC; "

        Dim banco As New AcessaBanco
        Dim dsDashboard As DataSet = banco.ConsultaDataSet(sql, "NotasFiscais")

        Dim listaProdutos As New List(Of ProdutoDashboard)

        For Each row As DataRow In dsDashboard.Tables(0).Rows

            Dim produto As New ProdutoDashboard

            'produto.EmpresaId = row("Empresa_Id")
            produto.ClienteId = row("ClienteId")
            produto.ClienteNome = row("ClienteNome")
            produto.ClienteFantasia = row("ClienteFantasia")
            produto.Pedidos = row("Pedido")
            produto.Data = row("Data")
            produto.PesoFiscal = row("PesoFiscal")
            produto.QuantidadeFiscal = row("Quantidade")
            produto.Valor = row("ValorTotal")

            listaProdutos.Add(produto)

        Next

        Return listaProdutos

    End Function

    Public Function CarregarDetalhamentoPedidos(ByVal filtro As ProdutoDashboardFiltro) As IEnumerable(Of ProdutoDashboard)

        Dim sql As String = String.Empty

        Dim sqlCondicaoIN As String = ""

        If Not String.IsNullOrEmpty(filtro.EmpresasSelecionada) Then
            Dim listaFormatada = filtro.EmpresasSelecionada.Split(","c).
                                    Where(Function(s) Not String.IsNullOrWhiteSpace(s)).
                                    Select(Function(par) $"'{par.Trim()}'")

            sqlCondicaoIN = String.Join(",", listaFormatada)

        End If

        Dim filtroSeguimento As String = IIf(filtro.SeguimentoSelecionado <> "0", String.Format("AND PRO.Seguimento = {0} ", filtro.SeguimentoSelecionado), "")

        sql = "WITH total_vendas AS 
                (
	                SELECT 	
		                NF.Empresa_Id,
		                NF.Cliente_Id,
		                NF.Serie_Id,
		                NF.Nota_Id,
                        NF.Pedido,
                        PE.DataPedido,
		                FORMAT(NF.Movimento, 'dd/MM/yyyy') AS Data,
	                    CLI.Cliente_Id AS ClienteId,
	                    CLI.Nome AS ClienteNome,
	                    CLI.Fantasia AS ClienteFantasia,
		                NFI.Produto_Id,
		                NFI.QuantidadeFiscal,
		                NFI.PesoFiscal,
		                NFI.Unitario,
		                NFE_LIQUIDO.Valor AS ValorLiquido,
		                NFE_PRO.Valor AS ValorTotal,
                        SO.Devolucao,
		                CLI.Cidade,
		                CLI.Estado
	                FROM NotasFiscais NF
	                INNER JOIN NotasFiscaisXItens NFI 
		                ON NF.Empresa_Id		= NFI.Empresa_Id 
		                AND NF.EndEmpresa_Id	= NFI.EndEmpresa_Id 
		                AND NF.Cliente_Id		= NFI.Cliente_Id 
		                AND NF.EndCliente_Id	= NFI.EndCliente_Id 
		                AND NF.EntradaSaida_Id	= NFI.EntradaSaida_Id 
		                AND NF.Serie_Id			= NFI.Serie_Id 
		                AND NF.Nota_Id			= NFI.Nota_Id
	                INNER JOIN NotasFiscaisXEncargos NFE_LIQUIDO
		                ON NF.Empresa_Id		= NFE_LIQUIDO.Empresa_Id 
		                AND NF.EndEmpresa_Id	= NFE_LIQUIDO.EndEmpresa_Id 
		                AND NF.Cliente_Id		= NFE_LIQUIDO.Cliente_Id 
		                AND NF.EndCliente_Id	= NFE_LIQUIDO.EndCliente_Id 
		                AND NF.EntradaSaida_Id	= NFE_LIQUIDO.EntradaSaida_Id 
		                AND NF.Serie_Id			= NFE_LIQUIDO.Serie_Id 
		                AND NF.Nota_Id			= NFE_LIQUIDO.Nota_Id
		                AND NFI.Produto_Id		= NFE_LIQUIDO.Produto_Id
		                AND NFI.Sequencia_Id	= NFE_LIQUIDO.Sequencia_id
		                AND NFE_LIQUIDO.Encargo_Id = 'LIQUIDO'
	                INNER JOIN NotasFiscaisXEncargos NFE_PRO
		                ON NF.Empresa_Id		= NFE_PRO.Empresa_Id 
		                AND NF.EndEmpresa_Id	= NFE_PRO.EndEmpresa_Id 
		                AND NF.Cliente_Id		= NFE_PRO.Cliente_Id 
		                AND NF.EndCliente_Id	= NFE_PRO.EndCliente_Id 
		                AND NF.EntradaSaida_Id	= NFE_PRO.EntradaSaida_Id 
		                AND NF.Serie_Id			= NFE_PRO.Serie_Id 
		                AND NF.Nota_Id			= NFE_PRO.Nota_Id
		                AND NFI.Produto_Id		= NFE_PRO.Produto_Id
		                AND NFI.Sequencia_Id	= NFE_PRO.Sequencia_id
		                AND NFE_PRO.Encargo_Id	= 'PRODUTO'
	                INNER JOIN SubOperacoes SO
		                ON NF.Operacao          = SO.Operacao_Id 
		                AND NF.SubOperacao      = SO.SubOperacoes_Id
	                INNER JOIN Produtos PRO
		                ON NFI.Produto_Id       = PRO.Produto_Id
	                INNER JOIN Seguimentos SEG
			            ON PRO.Seguimento = SEG.Seguimento_Id
                        AND PRO.Dashboard = 1
	                INNER JOIN Clientes CLI
		                ON NF.Cliente_Id        = CLI.Cliente_Id 
		                AND NF.EndCliente_Id    = CLI.Endereco_Id
	                INNER JOIN Pedidos PE
		                ON NF.Empresa_Id        = PE.Empresa_Id 
		                AND NF.EndEmpresa_Id    = PE.EndEmpresa_Id
                        AND NF.Pedido           = PE.Pedido_Id
	                WHERE 
		                ISNULL(NF.Situacao, 1)  = 1
		                AND (NF.Empresa_Id + '|' + CONVERT(VARCHAR, NF.EndEmpresa_Id)) IN (" & IIf(String.IsNullOrEmpty(filtro.EmpresasSelecionada), "(NF.Empresa_Id + '|' + NF.EndEmpresa_Id)", sqlCondicaoIN) & ")
		                AND NF.Movimento        >= '" & filtro.DataInicio.ToString("yyyy-MM-dd") & "'
		                AND NF.Movimento        <= '" & filtro.DataFim.ToString("yyyy-MM-dd") & "'
                        AND NCM IN ('23099010', '23091000')
		                AND SO.Classe IN ('VENDAS', 'VENDASAORDEM', 'EXPORTACOES', 'REMESSAS')
		                " & filtroSeguimento & "
                        AND LEFT(NF.Cliente_Id, 8) NOT IN ('49673784', '40938762')
                )  

            SELECT TOP 10
                ClienteId,
                ClienteNome,
                ClienteFantasia,
                FORMAT(DataPedido, 'dd/MM/yyyy') AS Data,
                Pedido,
                ISNULL(SUM(CASE WHEN DEVOLUCAO = 'S' THEN -1 ELSE 1 END * PesoFiscal), 0) AS PesoFiscal,
                ISNULL(SUM(CASE WHEN DEVOLUCAO = 'S' THEN -1 ELSE 1 END * PesoFiscal), 0) AS Quantidade,
                ISNULL(SUM(CASE WHEN DEVOLUCAO = 'S' THEN -1 ELSE 1 END * ValorLiquido), 0) AS ValorTotal
            FROM total_vendas
            GROUP BY
                ClienteId,
                ClienteNome,
                ClienteFantasia,
                DataPedido,
                Pedido
            ORDER BY
                SUM(CASE WHEN DEVOLUCAO = 'S' THEN -1 ELSE 1 END * ValorLiquido) DESC; "

        Dim banco As New AcessaBanco
        Dim dsDashboard As DataSet = banco.ConsultaDataSet(sql, "NotasFiscais")

        Dim listaProdutos As New List(Of ProdutoDashboard)

        For Each row As DataRow In dsDashboard.Tables(0).Rows

            Dim produto As New ProdutoDashboard

            'produto.EmpresaId = row("Empresa_Id")
            produto.ClienteId = row("ClienteId")
            produto.ClienteNome = row("ClienteNome")
            produto.ClienteFantasia = row("ClienteFantasia")
            produto.Pedidos = row("Pedido")
            produto.Data = row("Data")
            produto.PesoFiscal = row("PesoFiscal")
            produto.QuantidadeFiscal = row("Quantidade")
            produto.Valor = row("ValorTotal")

            listaProdutos.Add(produto)

        Next

        Return listaProdutos

    End Function

    'Vendas Diárias
    Public Function CarregarEvolucaoVendasDiarias(ByVal filtro As ProdutoDashboardFiltro) As IEnumerable(Of ProdutoDashboard)

        Dim sql As String = String.Empty

        Dim sqlCondicaoIN As String = ""

        If Not String.IsNullOrEmpty(filtro.EmpresasSelecionada) Then
            Dim listaFormatada = filtro.EmpresasSelecionada.Split(","c).
                                    Where(Function(s) Not String.IsNullOrWhiteSpace(s)).
                                    Select(Function(par) $"'{par.Trim()}'")

            sqlCondicaoIN = String.Join(",", listaFormatada)

        End If

        Dim filtroSeguimento As String = IIf(filtro.SeguimentoSelecionado <> "0", String.Format("AND PRO.Seguimento = {0} ", filtro.SeguimentoSelecionado), "")

        sql = " WITH total_vendas AS 
	                (
		                SELECT 	
			                NF.Empresa_Id,
			                NF.Cliente_Id,
			                NF.Serie_Id,
			                NF.Nota_Id,
			                NF.Movimento AS Data,
			                NFI.Produto_Id,
			                PRO.Nome AS ProdutoNome,
			                SEG.Descricao AS Seguimento,
			                NF.Operacao AS OperacaoNota,
			                NF.SubOperacao AS SubOperacaoNota,
			                NFI.QuantidadeFiscal,
			                NFI.PesoFiscal,
			                NFI.Unitario,
			                NFE_LIQUIDO.Valor AS ValorLiquido, --FORMAT(NFE_LIQUIDO.Valor, 'N2', 'pt-br') AS ValorLiquido,
			                NFE_PRO.Valor AS ValorTotal, --FORMAT(NFE_PRO.Valor, 'N2', 'pt-br') AS ValorTotal
                            SO.Devolucao
		                FROM NotasFiscais NF
		                INNER JOIN NotasFiscaisXItens NFI 
			                ON NF.Empresa_Id				= NFI.Empresa_Id 
			                AND NF.EndEmpresa_Id			= NFI.EndEmpresa_Id 
			                AND NF.Cliente_Id				= NFI.Cliente_Id 
			                AND NF.EndCliente_Id			= NFI.EndCliente_Id 
			                AND NF.EntradaSaida_Id			= NFI.EntradaSaida_Id 
			                AND NF.Serie_Id					= NFI.Serie_Id 
			                AND NF.Nota_Id					= NFI.Nota_Id
		                INNER JOIN NotasFiscaisXEncargos NFE_LIQUIDO
			                ON NF.Empresa_Id				= NFE_LIQUIDO.Empresa_Id 
			                AND NF.EndEmpresa_Id			= NFE_LIQUIDO.EndEmpresa_Id 
			                AND NF.Cliente_Id				= NFE_LIQUIDO.Cliente_Id 
			                AND NF.EndCliente_Id			= NFE_LIQUIDO.EndCliente_Id 
			                AND NF.EntradaSaida_Id			= NFE_LIQUIDO.EntradaSaida_Id 
			                AND NF.Serie_Id					= NFE_LIQUIDO.Serie_Id 
			                AND NF.Nota_Id					= NFE_LIQUIDO.Nota_Id
			                AND NFI.Produto_Id				= NFE_LIQUIDO.Produto_Id
			                AND NFI.Sequencia_Id			= NFE_LIQUIDO.Sequencia_id
			                AND NFE_LIQUIDO.Encargo_Id		= 'LIQUIDO'
		                INNER JOIN NotasFiscaisXEncargos NFE_PRO
			                ON NF.Empresa_Id				= NFE_PRO.Empresa_Id 
			                AND NF.EndEmpresa_Id			= NFE_PRO.EndEmpresa_Id 
			                AND NF.Cliente_Id				= NFE_PRO.Cliente_Id 
			                AND NF.EndCliente_Id			= NFE_PRO.EndCliente_Id 
			                AND NF.EntradaSaida_Id			= NFE_PRO.EntradaSaida_Id 
			                AND NF.Serie_Id					= NFE_PRO.Serie_Id 
			                AND NF.Nota_Id					= NFE_PRO.Nota_Id
			                AND NFI.Produto_Id				= NFE_PRO.Produto_Id
			                AND NFI.Sequencia_Id			= NFE_PRO.Sequencia_id
			                AND NFE_PRO.Encargo_Id			= 'PRODUTO'
		                INNER JOIN SubOperacoes SO
			                ON NF.Operacao = SO.Operacao_Id 
			                AND NF.SubOperacao = SO.SubOperacoes_Id
		                INNER JOIN Produtos PRO
			                ON NFI.Produto_Id = PRO.Produto_Id
		                INNER JOIN Seguimentos SEG
			                ON PRO.Seguimento = SEG.Seguimento_Id
                            AND PRO.Dashboard = 1
		                WHERE 
			                ISNULL(NF.Situacao, 1)  = 1
			                AND (NF.Empresa_Id + '|' + CONVERT(VARCHAR, NF.EndEmpresa_Id)) IN (" & IIf(String.IsNullOrEmpty(filtro.EmpresasSelecionada), "(NF.Empresa_Id + '|' + NF.EndEmpresa_Id)", sqlCondicaoIN) & ")
			                AND NF.Movimento        >= '" & filtro.DataInicio.ToString("yyyy-MM-dd") & "'
			                AND NF.Movimento        <= '" & filtro.DataFim.ToString("yyyy-MM-dd") & "'
                            AND NCM IN ('23099010', '23091000')
			                AND SO.Classe IN ('VENDAS', 'VENDASAORDEM', 'EXPORTACOES', 'REMESSAS')
			                " & filtroSeguimento & "
                            AND LEFT(NF.Cliente_Id, 8) NOT IN ('49673784', '40938762')
	                )  


                SELECT	CONVERT(VARCHAR(10), Data, 103) AS DataFormatada, -- 'dd/MM/yyyy' AS Data,
		                ISNULL(SUM(CASE WHEN DEVOLUCAO = 'S' THEN -1 ELSE 1 END * PesoFiscal), 0) AS Quantidade,
		                ISNULL(SUM(CASE WHEN DEVOLUCAO = 'S' THEN -1 ELSE 1 END * ValorLiquido), 0) AS ValorTotal
                FROM total_vendas
                GROUP BY
                    Data
                ORDER BY
                    Data DESC; "

        Dim banco As New AcessaBanco
        Dim dsDashboard As DataSet = banco.ConsultaDataSet(sql, "NotasFiscais")

        Dim listaProdutos As New List(Of ProdutoDashboard)

        For Each row As DataRow In dsDashboard.Tables(0).Rows

            Dim produto As New ProdutoDashboard

            produto.Data = row("DataFormatada")
            produto.QuantidadeFiscal = row("Quantidade")
            produto.Valor = row("ValorTotal")

            listaProdutos.Add(produto)

        Next

        Return listaProdutos

    End Function

    Public Function CarregarDetalhamentoVendasDiarias(ByVal filtro As ProdutoDashboardFiltro) As IEnumerable(Of ProdutoDashboard)

        Dim sql As String = String.Empty

        Dim sqlCondicaoIN As String = ""

        If Not String.IsNullOrEmpty(filtro.EmpresasSelecionada) Then
            Dim listaFormatada = filtro.EmpresasSelecionada.Split(","c).
                                    Where(Function(s) Not String.IsNullOrWhiteSpace(s)).
                                    Select(Function(par) $"'{par.Trim()}'")

            sqlCondicaoIN = String.Join(",", listaFormatada)

        End If

        Dim filtroSeguimento As String = IIf(filtro.SeguimentoSelecionado <> "0", String.Format("AND PRO.Seguimento = {0} ", filtro.SeguimentoSelecionado), "")

        sql = " WITH total_vendas AS 
	                (
		                SELECT 	
			                NF.Empresa_Id,
			                NF.Cliente_Id,
			                NF.Serie_Id,
			                NF.Nota_Id,
                            NF.Pedido,
			                NF.Movimento AS Data,
			                NFI.Produto_Id,
			                PRO.Nome AS ProdutoNome,
			                SEG.Descricao AS Seguimento,
			                NF.Operacao AS OperacaoNota,
			                NF.SubOperacao AS SubOperacaoNota,
			                NFI.QuantidadeFiscal,
			                NFI.PesoFiscal,
			                NFI.Unitario,
			                NFE_LIQUIDO.Valor AS ValorLiquido, --FORMAT(NFE_LIQUIDO.Valor, 'N2', 'pt-br') AS ValorLiquido,
			                NFE_PRO.Valor AS ValorTotal, --FORMAT(NFE_PRO.Valor, 'N2', 'pt-br') AS ValorTotal
                            SO.Devolucao
		                FROM NotasFiscais NF
		                INNER JOIN NotasFiscaisXItens NFI 
			                ON NF.Empresa_Id				= NFI.Empresa_Id 
			                AND NF.EndEmpresa_Id			= NFI.EndEmpresa_Id 
			                AND NF.Cliente_Id				= NFI.Cliente_Id 
			                AND NF.EndCliente_Id			= NFI.EndCliente_Id 
			                AND NF.EntradaSaida_Id			= NFI.EntradaSaida_Id 
			                AND NF.Serie_Id					= NFI.Serie_Id 
			                AND NF.Nota_Id					= NFI.Nota_Id
		                INNER JOIN NotasFiscaisXEncargos NFE_LIQUIDO
			                ON NF.Empresa_Id				= NFE_LIQUIDO.Empresa_Id 
			                AND NF.EndEmpresa_Id			= NFE_LIQUIDO.EndEmpresa_Id 
			                AND NF.Cliente_Id				= NFE_LIQUIDO.Cliente_Id 
			                AND NF.EndCliente_Id			= NFE_LIQUIDO.EndCliente_Id 
			                AND NF.EntradaSaida_Id			= NFE_LIQUIDO.EntradaSaida_Id 
			                AND NF.Serie_Id					= NFE_LIQUIDO.Serie_Id 
			                AND NF.Nota_Id					= NFE_LIQUIDO.Nota_Id
			                AND NFI.Produto_Id				= NFE_LIQUIDO.Produto_Id
			                AND NFI.Sequencia_Id			= NFE_LIQUIDO.Sequencia_id
			                AND NFE_LIQUIDO.Encargo_Id		= 'LIQUIDO'
		                INNER JOIN NotasFiscaisXEncargos NFE_PRO
			                ON NF.Empresa_Id				= NFE_PRO.Empresa_Id 
			                AND NF.EndEmpresa_Id			= NFE_PRO.EndEmpresa_Id 
			                AND NF.Cliente_Id				= NFE_PRO.Cliente_Id 
			                AND NF.EndCliente_Id			= NFE_PRO.EndCliente_Id 
			                AND NF.EntradaSaida_Id			= NFE_PRO.EntradaSaida_Id 
			                AND NF.Serie_Id					= NFE_PRO.Serie_Id 
			                AND NF.Nota_Id					= NFE_PRO.Nota_Id
			                AND NFI.Produto_Id				= NFE_PRO.Produto_Id
			                AND NFI.Sequencia_Id			= NFE_PRO.Sequencia_id
			                AND NFE_PRO.Encargo_Id			= 'PRODUTO'
		                INNER JOIN SubOperacoes SO
			                ON NF.Operacao = SO.Operacao_Id 
			                AND NF.SubOperacao = SO.SubOperacoes_Id
		                INNER JOIN Produtos PRO
			                ON NFI.Produto_Id = PRO.Produto_Id
		                INNER JOIN Seguimentos SEG
			                ON PRO.Seguimento = SEG.Seguimento_Id
                            AND PRO.Dashboard = 1
		                WHERE 
			                ISNULL(NF.Situacao, 1)  = 1
			                AND (NF.Empresa_Id + '|' + CONVERT(VARCHAR, NF.EndEmpresa_Id)) IN (" & IIf(String.IsNullOrEmpty(filtro.EmpresasSelecionada), "(NF.Empresa_Id + '|' + NF.EndEmpresa_Id)", sqlCondicaoIN) & ")
			                AND NF.Movimento        >= '" & filtro.DataInicio.ToString("yyyy-MM-dd") & "'
			                AND NF.Movimento        <= '" & filtro.DataFim.ToString("yyyy-MM-dd") & "'
                            AND NCM IN ('23099010', '23091000')
			                AND SO.Classe IN ('VENDAS', 'VENDASAORDEM', 'EXPORTACOES', 'REMESSAS')
			                " & filtroSeguimento & "
                            AND LEFT(NF.Cliente_Id, 8) NOT IN ('49673784', '40938762')
	                )  

                SELECT	CONVERT(VARCHAR(10), Data, 103) AS DataFormatada, 
						COUNT(Pedido) AS Pedido,
		                ISNULL(SUM(CASE WHEN DEVOLUCAO = 'S' THEN -1 ELSE 1 END * PesoFiscal), 0) AS Quantidade,
		                ISNULL(SUM(CASE WHEN DEVOLUCAO = 'S' THEN -1 ELSE 1 END * ValorLiquido), 0) AS ValorTotal
                FROM total_vendas
                GROUP BY
                    Data
                ORDER BY
                    Data DESC; "

        Dim banco As New AcessaBanco
        Dim dsDashboard As DataSet = banco.ConsultaDataSet(sql, "NotasFiscais")

        Dim listaProdutos As New List(Of ProdutoDashboard)

        For Each row As DataRow In dsDashboard.Tables(0).Rows

            Dim produto As New ProdutoDashboard

            produto.Data = row("DataFormatada")
            produto.QuantidadePedido = row("Pedido")
            produto.QuantidadeFiscal = row("Quantidade")
            produto.Valor = row("ValorTotal")

            listaProdutos.Add(produto)

        Next

        Return listaProdutos

    End Function

    Public Function CarregarTotalVendasDiarias(ByVal filtro As ProdutoDashboardFiltro) As ProdutoDashboard

        Dim sql As String = String.Empty

        Dim sqlCondicaoIN As String = ""

        If Not String.IsNullOrEmpty(filtro.EmpresasSelecionada) Then
            Dim listaFormatada = filtro.EmpresasSelecionada.Split(","c).
                                    Where(Function(s) Not String.IsNullOrWhiteSpace(s)).
                                    Select(Function(par) $"'{par.Trim()}'")

            sqlCondicaoIN = String.Join(",", listaFormatada)

        End If

        Dim filtroSeguimento As String = IIf(filtro.SeguimentoSelecionado <> "0", String.Format("AND PRO.Seguimento = {0} ", filtro.SeguimentoSelecionado), "")

        sql = " WITH total_vendas AS 
	                (
		                SELECT 	
			                NF.Empresa_Id,
			                NF.Cliente_Id,
			                NF.Serie_Id,
			                NF.Nota_Id,
			                NF.Movimento AS Data,
			                NFI.Produto_Id,
			                PRO.Nome AS ProdutoNome,
			                SEG.Descricao AS Seguimento,
			                NF.Operacao AS OperacaoNota,
			                NF.SubOperacao AS SubOperacaoNota,
			                NFI.QuantidadeFiscal,
			                NFI.PesoFiscal,
			                NFI.Unitario,
			                NFE_LIQUIDO.Valor AS ValorLiquido, --FORMAT(NFE_LIQUIDO.Valor, 'N2', 'pt-br') AS ValorLiquido,
			                NFE_PRO.Valor AS ValorTotal, --FORMAT(NFE_PRO.Valor, 'N2', 'pt-br') AS ValorTotal
                            SO.Devolucao
		                FROM NotasFiscais NF
		                INNER JOIN NotasFiscaisXItens NFI 
			                ON NF.Empresa_Id				= NFI.Empresa_Id 
			                AND NF.EndEmpresa_Id			= NFI.EndEmpresa_Id 
			                AND NF.Cliente_Id				= NFI.Cliente_Id 
			                AND NF.EndCliente_Id			= NFI.EndCliente_Id 
			                AND NF.EntradaSaida_Id			= NFI.EntradaSaida_Id 
			                AND NF.Serie_Id					= NFI.Serie_Id 
			                AND NF.Nota_Id					= NFI.Nota_Id
		                INNER JOIN NotasFiscaisXEncargos NFE_LIQUIDO
			                ON NF.Empresa_Id				= NFE_LIQUIDO.Empresa_Id 
			                AND NF.EndEmpresa_Id			= NFE_LIQUIDO.EndEmpresa_Id 
			                AND NF.Cliente_Id				= NFE_LIQUIDO.Cliente_Id 
			                AND NF.EndCliente_Id			= NFE_LIQUIDO.EndCliente_Id 
			                AND NF.EntradaSaida_Id			= NFE_LIQUIDO.EntradaSaida_Id 
			                AND NF.Serie_Id					= NFE_LIQUIDO.Serie_Id 
			                AND NF.Nota_Id					= NFE_LIQUIDO.Nota_Id
			                AND NFI.Produto_Id				= NFE_LIQUIDO.Produto_Id
			                AND NFI.Sequencia_Id			= NFE_LIQUIDO.Sequencia_id
			                AND NFE_LIQUIDO.Encargo_Id		= 'LIQUIDO'
		                INNER JOIN NotasFiscaisXEncargos NFE_PRO
			                ON NF.Empresa_Id				= NFE_PRO.Empresa_Id 
			                AND NF.EndEmpresa_Id			= NFE_PRO.EndEmpresa_Id 
			                AND NF.Cliente_Id				= NFE_PRO.Cliente_Id 
			                AND NF.EndCliente_Id			= NFE_PRO.EndCliente_Id 
			                AND NF.EntradaSaida_Id			= NFE_PRO.EntradaSaida_Id 
			                AND NF.Serie_Id					= NFE_PRO.Serie_Id 
			                AND NF.Nota_Id					= NFE_PRO.Nota_Id
			                AND NFI.Produto_Id				= NFE_PRO.Produto_Id
			                AND NFI.Sequencia_Id			= NFE_PRO.Sequencia_id
			                AND NFE_PRO.Encargo_Id			= 'PRODUTO'
		                INNER JOIN SubOperacoes SO
			                ON NF.Operacao = SO.Operacao_Id 
			                AND NF.SubOperacao = SO.SubOperacoes_Id
		                INNER JOIN Produtos PRO
			                ON NFI.Produto_Id = PRO.Produto_Id
		                INNER JOIN Seguimentos SEG
			                ON PRO.Seguimento = SEG.Seguimento_Id
                            AND PRO.Dashboard = 1
		                WHERE 
			                ISNULL(NF.Situacao, 1)  = 1
			                AND (NF.Empresa_Id + '|' + CONVERT(VARCHAR, NF.EndEmpresa_Id)) IN (" & IIf(String.IsNullOrEmpty(filtro.EmpresasSelecionada), "(NF.Empresa_Id + '|' + NF.EndEmpresa_Id)", sqlCondicaoIN) & ")
			                AND NF.Movimento        >= '" & filtro.DataInicio.ToString("yyyy-MM-dd") & "'
			                AND NF.Movimento        <= '" & filtro.DataFim.ToString("yyyy-MM-dd") & "'
                            AND NCM IN ('23099010', '23091000')
			                AND SO.Classe IN ('VENDAS', 'VENDASAORDEM', 'EXPORTACOES', 'REMESSAS')
			                " & filtroSeguimento & "
                            AND LEFT(NF.Cliente_Id, 8) NOT IN ('49673784', '40938762')
	                )  

                    SELECT  ISNULL(SUM(CASE WHEN DEVOLUCAO = 'S' THEN -1 ELSE 1 END * PesoFiscal), 0) AS Quantidade,
                            ISNULL(SUM(CASE WHEN DEVOLUCAO = 'S' THEN -1 ELSE 1 END * ValorLiquido), 0) AS ValorTotal
                    FROM total_vendas; "

        Dim banco As New AcessaBanco
        Dim dsDashboard As DataSet = banco.ConsultaDataSet(sql, "NotasFiscais")

        Dim produto As New ProdutoDashboard

        For Each row As DataRow In dsDashboard.Tables(0).Rows

            produto.QuantidadeFiscal = row("Quantidade")
            produto.Valor = row("ValorTotal")

        Next

        Return produto

    End Function

    'Vendas Mensais
    Public Function CarregarEvolucaoVendasMensais(ByVal filtro As ProdutoDashboardFiltro) As IEnumerable(Of ProdutoDashboard)
        Dim sql As String = String.Empty

        Dim sqlCondicaoIN As String = ""

        If Not String.IsNullOrEmpty(filtro.EmpresasSelecionada) Then
            Dim listaFormatada = filtro.EmpresasSelecionada.Split(","c).
                                    Where(Function(s) Not String.IsNullOrWhiteSpace(s)).
                                    Select(Function(par) $"'{par.Trim()}'")

            sqlCondicaoIN = String.Join(",", listaFormatada)

        End If

        Dim filtroSeguimento As String = IIf(filtro.SeguimentoSelecionado <> "0", String.Format("AND PRO.Seguimento = {0} ", filtro.SeguimentoSelecionado), "")

        sql = " WITH Meses AS (
                    SELECT CAST('" & filtro.DataInicio.ToString("yyyy-MM-dd") & "' AS DATE) AS PrimeiroDiaMes
                    UNION ALL
                    SELECT DATEADD(MONTH, 1, PrimeiroDiaMes)
                    FROM Meses
                    WHERE DATEADD(MONTH, 1, PrimeiroDiaMes) <= '" & filtro.DataFim.ToString("yyyy-MM-dd") & "'
                ),

                total_vendas AS (

                    SELECT 
                        NF.Movimento AS Data,
                        NFI.QuantidadeFiscal,
                        NFI.PesoFiscal,
                        NFE_LIQUIDO.Valor AS ValorLiquido,
                        SO.Devolucao
                    FROM NotasFiscais NF
                    INNER JOIN NotasFiscaisXItens NFI 
                        ON NF.Empresa_Id = NFI.Empresa_Id 
                        AND NF.EndEmpresa_Id = NFI.EndEmpresa_Id 
                        AND NF.Cliente_Id = NFI.Cliente_Id 
                        AND NF.EndCliente_Id = NFI.EndCliente_Id 
                        AND NF.EntradaSaida_Id = NFI.EntradaSaida_Id 
                        AND NF.Serie_Id = NFI.Serie_Id 
                        AND NF.Nota_Id = NFI.Nota_Id
                    INNER JOIN NotasFiscaisXEncargos NFE_LIQUIDO
                        ON NF.Empresa_Id = NFE_LIQUIDO.Empresa_Id 
                        AND NF.EndEmpresa_Id = NFE_LIQUIDO.EndEmpresa_Id 
                        AND NF.Cliente_Id = NFE_LIQUIDO.Cliente_Id 
                        AND NF.EndCliente_Id = NFE_LIQUIDO.EndCliente_Id 
                        AND NF.EntradaSaida_Id = NFE_LIQUIDO.EntradaSaida_Id 
                        AND NF.Serie_Id = NFE_LIQUIDO.Serie_Id 
                        AND NF.Nota_Id = NFE_LIQUIDO.Nota_Id
                        AND NFI.Produto_Id = NFE_LIQUIDO.Produto_Id
                        AND NFI.Sequencia_Id = NFE_LIQUIDO.Sequencia_id
                        AND NFE_LIQUIDO.Encargo_Id = 'LIQUIDO'
                    INNER JOIN SubOperacoes SO
                        ON NF.Operacao = SO.Operacao_Id 
                        AND NF.SubOperacao = SO.SubOperacoes_Id
                    INNER JOIN Produtos PRO
                        ON NFI.Produto_Id = PRO.Produto_Id
                    INNER JOIN Seguimentos SEG
			            ON PRO.Seguimento = SEG.Seguimento_Id
                        AND PRO.Dashboard = 1
                    WHERE 
                        ISNULL(NF.Situacao, 1)  = 1
                        AND (NF.Empresa_Id + '|' + CONVERT(VARCHAR, NF.EndEmpresa_Id)) IN (" & IIf(String.IsNullOrEmpty(filtro.EmpresasSelecionada), "(NF.Empresa_Id + '|' + NF.EndEmpresa_Id)", sqlCondicaoIN) & ")
                        AND NF.Movimento        >= '" & filtro.DataInicio.ToString("yyyy-MM-dd") & "'
                        AND NF.Movimento        <= '" & filtro.DataFim.ToString("yyyy-MM-dd") & "'
                        AND NCM IN ('23099010', '23091000')
                        AND SO.Classe IN ('VENDAS', 'VENDASAORDEM', 'EXPORTACOES', 'REMESSAS')
                        " & filtroSeguimento & "
                        AND LEFT(NF.Cliente_Id, 8) NOT IN ('49673784', '40938762')
                )

                SELECT 
                    FORMAT(M.PrimeiroDiaMes, 'MM-yyyy') AS DataFormatada,
                    ISNULL(SUM(CASE WHEN DEVOLUCAO = 'S' THEN -1 ELSE 1 END * total_vendas.PesoFiscal), 0) AS Quantidade,
                    ISNULL(SUM(CASE WHEN DEVOLUCAO = 'S' THEN -1 ELSE 1 END * total_vendas.ValorLiquido), 0) AS ValorTotal
                FROM Meses M
                LEFT JOIN total_vendas
                    ON YEAR(total_vendas.Data) = YEAR(M.PrimeiroDiaMes)
                    AND MONTH(total_vendas.Data) = MONTH(M.PrimeiroDiaMes)
                GROUP BY M.PrimeiroDiaMes
                ORDER BY M.PrimeiroDiaMes DESC "

        Dim banco As New AcessaBanco
        Dim dsDashboard As DataSet = banco.ConsultaDataSet(sql, "NotasFiscais")

        Dim listaProdutos As New List(Of ProdutoDashboard)

        For Each row As DataRow In dsDashboard.Tables(0).Rows

            Dim produto As New ProdutoDashboard

            produto.Data = row("DataFormatada")
            produto.QuantidadeFiscal = row("Quantidade")
            produto.Valor = row("ValorTotal")

            listaProdutos.Add(produto)

        Next

        Return listaProdutos
    End Function

    Public Function CarregarDetalhamentoVendasMensais(ByVal filtro As ProdutoDashboardFiltro) As IEnumerable(Of ProdutoDashboard)
        Dim sql As String = String.Empty

        Dim sqlCondicaoIN As String = ""

        If Not String.IsNullOrEmpty(filtro.EmpresasSelecionada) Then
            Dim listaFormatada = filtro.EmpresasSelecionada.Split(","c).
                                    Where(Function(s) Not String.IsNullOrWhiteSpace(s)).
                                    Select(Function(par) $"'{par.Trim()}'")

            sqlCondicaoIN = String.Join(",", listaFormatada)

        End If

        Dim filtroSeguimento As String = IIf(filtro.SeguimentoSelecionado <> "0", String.Format("AND PRO.Seguimento = {0} ", filtro.SeguimentoSelecionado), "")

        sql = " WITH Meses AS (
                    SELECT CAST('" & filtro.DataInicio.ToString("yyyy-MM-dd") & "' AS DATE) AS PrimeiroDiaMes
                    UNION ALL
                    SELECT DATEADD(MONTH, 1, PrimeiroDiaMes)
                    FROM Meses
                    WHERE DATEADD(MONTH, 1, PrimeiroDiaMes) <= '" & filtro.DataFim.ToString("yyyy-MM-dd") & "'
                ),

                total_vendas AS (

                    SELECT 
                        NF.Movimento AS Data,
                        NFI.QuantidadeFiscal,
                        NFI.PesoFiscal,
                        NF.Pedido,
                        NFE_LIQUIDO.Valor AS ValorLiquido,
                        SO.Devolucao
                    FROM NotasFiscais NF
                    INNER JOIN NotasFiscaisXItens NFI 
                        ON NF.Empresa_Id = NFI.Empresa_Id 
                        AND NF.EndEmpresa_Id = NFI.EndEmpresa_Id 
                        AND NF.Cliente_Id = NFI.Cliente_Id 
                        AND NF.EndCliente_Id = NFI.EndCliente_Id 
                        AND NF.EntradaSaida_Id = NFI.EntradaSaida_Id 
                        AND NF.Serie_Id = NFI.Serie_Id 
                        AND NF.Nota_Id = NFI.Nota_Id
                    INNER JOIN NotasFiscaisXEncargos NFE_LIQUIDO
                        ON NF.Empresa_Id = NFE_LIQUIDO.Empresa_Id 
                        AND NF.EndEmpresa_Id = NFE_LIQUIDO.EndEmpresa_Id 
                        AND NF.Cliente_Id = NFE_LIQUIDO.Cliente_Id 
                        AND NF.EndCliente_Id = NFE_LIQUIDO.EndCliente_Id 
                        AND NF.EntradaSaida_Id = NFE_LIQUIDO.EntradaSaida_Id 
                        AND NF.Serie_Id = NFE_LIQUIDO.Serie_Id 
                        AND NF.Nota_Id = NFE_LIQUIDO.Nota_Id
                        AND NFI.Produto_Id = NFE_LIQUIDO.Produto_Id
                        AND NFI.Sequencia_Id = NFE_LIQUIDO.Sequencia_id
                        AND NFE_LIQUIDO.Encargo_Id = 'LIQUIDO'
                    INNER JOIN SubOperacoes SO
                        ON NF.Operacao = SO.Operacao_Id 
                        AND NF.SubOperacao = SO.SubOperacoes_Id
                    INNER JOIN Produtos PRO
                        ON NFI.Produto_Id = PRO.Produto_Id
                    INNER JOIN Seguimentos SEG
			            ON PRO.Seguimento = SEG.Seguimento_Id
                        AND PRO.Dashboard = 1
                    WHERE 
                        ISNULL(NF.Situacao, 1)  = 1
                        AND (NF.Empresa_Id + '|' + CONVERT(VARCHAR, NF.EndEmpresa_Id)) IN (" & IIf(String.IsNullOrEmpty(filtro.EmpresasSelecionada), "(NF.Empresa_Id + '|' + NF.EndEmpresa_Id)", sqlCondicaoIN) & ")
                        AND NF.Movimento        >= '" & filtro.DataInicio.ToString("yyyy-MM-dd") & "'
                        AND NF.Movimento        <= '" & filtro.DataFim.ToString("yyyy-MM-dd") & "'
                        AND NCM IN ('23099010', '23091000')
                        AND SO.Classe IN ('VENDAS', 'VENDASAORDEM', 'EXPORTACOES', 'REMESSAS')
                        " & filtroSeguimento & "
                        AND LEFT(NF.Cliente_Id, 8) NOT IN ('49673784', '40938762')
                )

                SELECT 
                    FORMAT(M.PrimeiroDiaMes, 'MM-yyyy') AS DataFormatada,
                    COUNT(total_vendas.Pedido) AS Pedido,
                    ISNULL(SUM(CASE WHEN DEVOLUCAO = 'S' THEN -1 ELSE 1 END * total_vendas.PesoFiscal), 0) AS Quantidade,
                    ISNULL(SUM(CASE WHEN DEVOLUCAO = 'S' THEN -1 ELSE 1 END * total_vendas.ValorLiquido), 0) AS ValorTotal
                FROM Meses M
                LEFT JOIN total_vendas
                    ON YEAR(total_vendas.Data) = YEAR(M.PrimeiroDiaMes)
                    AND MONTH(total_vendas.Data) = MONTH(M.PrimeiroDiaMes)
                GROUP BY M.PrimeiroDiaMes
                ORDER BY M.PrimeiroDiaMes DESC "

        Dim banco As New AcessaBanco
        Dim dsDashboard As DataSet = banco.ConsultaDataSet(sql, "NotasFiscais")

        Dim listaProdutos As New List(Of ProdutoDashboard)

        For Each row As DataRow In dsDashboard.Tables(0).Rows

            Dim produto As New ProdutoDashboard

            produto.Data = row("DataFormatada")
            produto.QuantidadePedido = row("Pedido")
            produto.QuantidadeFiscal = row("Quantidade")
            produto.Valor = row("ValorTotal")

            listaProdutos.Add(produto)

        Next

        Return listaProdutos
    End Function

    Public Function CarregarTotalVendasMensais(ByVal filtro As ProdutoDashboardFiltro) As ProdutoDashboard
        Dim sql As String = String.Empty

        Dim sqlCondicaoIN As String = ""

        If Not String.IsNullOrEmpty(filtro.EmpresasSelecionada) Then
            Dim listaFormatada = filtro.EmpresasSelecionada.Split(","c).
                                    Where(Function(s) Not String.IsNullOrWhiteSpace(s)).
                                    Select(Function(par) $"'{par.Trim()}'")

            sqlCondicaoIN = String.Join(",", listaFormatada)

        End If

        Dim filtroSeguimento As String = IIf(filtro.SeguimentoSelecionado <> "0", String.Format("AND PRO.Seguimento = {0} ", filtro.SeguimentoSelecionado), "")

        sql = " WITH total_vendas AS 
	                (
		                SELECT 	
			                NF.Empresa_Id,
			                NF.Cliente_Id,
			                NF.Serie_Id,
			                NF.Nota_Id,
			                NF.Movimento AS Data,
			                NFI.Produto_Id,
			                PRO.Nome AS ProdutoNome,
			                SEG.Descricao AS Seguimento,
			                NF.Operacao AS OperacaoNota,
			                NF.SubOperacao AS SubOperacaoNota,
			                NFI.QuantidadeFiscal,
			                NFI.PesoFiscal,
			                NFI.Unitario,
			                NFE_LIQUIDO.Valor AS ValorLiquido, --FORMAT(NFE_LIQUIDO.Valor, 'N2', 'pt-br') AS ValorLiquido,
			                NFE_PRO.Valor AS ValorTotal, --FORMAT(NFE_PRO.Valor, 'N2', 'pt-br') AS ValorTotal
                            SO.Devolucao
		                FROM NotasFiscais NF
		                INNER JOIN NotasFiscaisXItens NFI 
			                ON NF.Empresa_Id				= NFI.Empresa_Id 
			                AND NF.EndEmpresa_Id			= NFI.EndEmpresa_Id 
			                AND NF.Cliente_Id				= NFI.Cliente_Id 
			                AND NF.EndCliente_Id			= NFI.EndCliente_Id 
			                AND NF.EntradaSaida_Id			= NFI.EntradaSaida_Id 
			                AND NF.Serie_Id					= NFI.Serie_Id 
			                AND NF.Nota_Id					= NFI.Nota_Id
		                INNER JOIN NotasFiscaisXEncargos NFE_LIQUIDO
			                ON NF.Empresa_Id				= NFE_LIQUIDO.Empresa_Id 
			                AND NF.EndEmpresa_Id			= NFE_LIQUIDO.EndEmpresa_Id 
			                AND NF.Cliente_Id				= NFE_LIQUIDO.Cliente_Id 
			                AND NF.EndCliente_Id			= NFE_LIQUIDO.EndCliente_Id 
			                AND NF.EntradaSaida_Id			= NFE_LIQUIDO.EntradaSaida_Id 
			                AND NF.Serie_Id					= NFE_LIQUIDO.Serie_Id 
			                AND NF.Nota_Id					= NFE_LIQUIDO.Nota_Id
			                AND NFI.Produto_Id				= NFE_LIQUIDO.Produto_Id
			                AND NFI.Sequencia_Id			= NFE_LIQUIDO.Sequencia_id
			                AND NFE_LIQUIDO.Encargo_Id		= 'LIQUIDO'
		                INNER JOIN NotasFiscaisXEncargos NFE_PRO
			                ON NF.Empresa_Id				= NFE_PRO.Empresa_Id 
			                AND NF.EndEmpresa_Id			= NFE_PRO.EndEmpresa_Id 
			                AND NF.Cliente_Id				= NFE_PRO.Cliente_Id 
			                AND NF.EndCliente_Id			= NFE_PRO.EndCliente_Id 
			                AND NF.EntradaSaida_Id			= NFE_PRO.EntradaSaida_Id 
			                AND NF.Serie_Id					= NFE_PRO.Serie_Id 
			                AND NF.Nota_Id					= NFE_PRO.Nota_Id
			                AND NFI.Produto_Id				= NFE_PRO.Produto_Id
			                AND NFI.Sequencia_Id			= NFE_PRO.Sequencia_id
			                AND NFE_PRO.Encargo_Id			= 'PRODUTO'
		                INNER JOIN SubOperacoes SO
			                ON NF.Operacao = SO.Operacao_Id 
			                AND NF.SubOperacao = SO.SubOperacoes_Id
		                INNER JOIN Produtos PRO
			                ON NFI.Produto_Id = PRO.Produto_Id
		                INNER JOIN Seguimentos SEG
			                ON PRO.Seguimento = SEG.Seguimento_Id
                            AND PRO.Dashboard = 1
		                WHERE 
			                ISNULL(NF.Situacao, 1)  = 1
			                AND (NF.Empresa_Id + '|' + CONVERT(VARCHAR, NF.EndEmpresa_Id)) IN (" & IIf(String.IsNullOrEmpty(filtro.EmpresasSelecionada), "(NF.Empresa_Id + '|' + NF.EndEmpresa_Id)", sqlCondicaoIN) & ")
			                AND NF.Movimento        >= '" & filtro.DataInicio.ToString("yyyy-MM-dd") & "'
			                AND NF.Movimento        <= '" & filtro.DataFim.ToString("yyyy-MM-dd") & "'
                            AND NCM IN ('23099010', '23091000')
			                AND SO.Classe IN ('VENDAS', 'VENDASAORDEM', 'EXPORTACOES', 'REMESSAS')
			                " & filtroSeguimento & "
                            AND LEFT(NF.Cliente_Id, 8) NOT IN ('49673784', '40938762')
	                )  

                    SELECT  ISNULL(SUM(CASE WHEN DEVOLUCAO = 'S' THEN -1 ELSE 1 END * PesoFiscal), 0) AS Quantidade,
                            ISNULL(SUM(CASE WHEN DEVOLUCAO = 'S' THEN -1 ELSE 1 END * ValorLiquido), 0) AS ValorTotal
                    FROM total_vendas; "

        Dim banco As New AcessaBanco
        Dim dsDashboard As DataSet = banco.ConsultaDataSet(sql, "NotasFiscais")

        Dim produto As New ProdutoDashboard

        For Each row As DataRow In dsDashboard.Tables(0).Rows

            produto.QuantidadeFiscal = row("Quantidade")
            produto.Valor = row("ValorTotal")

        Next

        Return produto
    End Function

    'Consolidada
    Public Function CarregarTotalConsolidado(ByVal filtro As ProdutoDashboardFiltro) As ProdutoDashboard

        Dim sql As String = String.Empty

        Dim sqlCondicaoIN As String = ""

        If Not String.IsNullOrEmpty(filtro.EmpresasSelecionada) Then
            Dim listaFormatada = filtro.EmpresasSelecionada.Split(","c).
                                    Where(Function(s) Not String.IsNullOrWhiteSpace(s)).
                                    Select(Function(par) $"'{par.Trim()}'")

            sqlCondicaoIN = String.Join(",", listaFormatada)

        End If

        Dim filtroSeguimento As String = IIf(filtro.SeguimentoSelecionado <> "0", String.Format("AND PRO.Seguimento = {0} ", filtro.SeguimentoSelecionado), "")

        sql = " WITH total_vendas AS 
	                (
		                SELECT 	
			                NF.Empresa_Id,
			                NF.Cliente_Id,
			                NF.Serie_Id,
			                NF.Nota_Id,
			                FORMAT(NF.Movimento, 'dd/MM/yyyy') AS Data,
			                NFI.Produto_Id,
			                PRO.Nome AS ProdutoNome,
			                SEG.Descricao AS Seguimento,
			                NF.Operacao AS OperacaoNota,
			                NF.SubOperacao AS SubOperacaoNota,
			                NFI.QuantidadeFiscal,
			                NFI.PesoFiscal,
			                NFI.Unitario,
			                NFE_LIQUIDO.Valor AS ValorLiquido, --FORMAT(NFE_LIQUIDO.Valor, 'N2', 'pt-br') AS ValorLiquido,
			                NFE_PRO.Valor AS ValorTotal, --FORMAT(NFE_PRO.Valor, 'N2', 'pt-br') AS ValorTotal
                            SO.Devolucao
		                FROM NotasFiscais NF
		                INNER JOIN NotasFiscaisXItens NFI 
			                ON NF.Empresa_Id				= NFI.Empresa_Id 
			                AND NF.EndEmpresa_Id			= NFI.EndEmpresa_Id 
			                AND NF.Cliente_Id				= NFI.Cliente_Id 
			                AND NF.EndCliente_Id			= NFI.EndCliente_Id 
			                AND NF.EntradaSaida_Id			= NFI.EntradaSaida_Id 
			                AND NF.Serie_Id					= NFI.Serie_Id 
			                AND NF.Nota_Id					= NFI.Nota_Id
		                INNER JOIN NotasFiscaisXEncargos NFE_LIQUIDO
			                ON NF.Empresa_Id				= NFE_LIQUIDO.Empresa_Id 
			                AND NF.EndEmpresa_Id			= NFE_LIQUIDO.EndEmpresa_Id 
			                AND NF.Cliente_Id				= NFE_LIQUIDO.Cliente_Id 
			                AND NF.EndCliente_Id			= NFE_LIQUIDO.EndCliente_Id 
			                AND NF.EntradaSaida_Id			= NFE_LIQUIDO.EntradaSaida_Id 
			                AND NF.Serie_Id					= NFE_LIQUIDO.Serie_Id 
			                AND NF.Nota_Id					= NFE_LIQUIDO.Nota_Id
			                AND NFI.Produto_Id				= NFE_LIQUIDO.Produto_Id
			                AND NFI.Sequencia_Id			= NFE_LIQUIDO.Sequencia_id
			                AND NFE_LIQUIDO.Encargo_Id		= 'LIQUIDO'
		                INNER JOIN NotasFiscaisXEncargos NFE_PRO
			                ON NF.Empresa_Id				= NFE_PRO.Empresa_Id 
			                AND NF.EndEmpresa_Id			= NFE_PRO.EndEmpresa_Id 
			                AND NF.Cliente_Id				= NFE_PRO.Cliente_Id 
			                AND NF.EndCliente_Id			= NFE_PRO.EndCliente_Id 
			                AND NF.EntradaSaida_Id			= NFE_PRO.EntradaSaida_Id 
			                AND NF.Serie_Id					= NFE_PRO.Serie_Id 
			                AND NF.Nota_Id					= NFE_PRO.Nota_Id
			                AND NFI.Produto_Id				= NFE_PRO.Produto_Id
			                AND NFI.Sequencia_Id			= NFE_PRO.Sequencia_id
			                AND NFE_PRO.Encargo_Id			= 'PRODUTO'
		                INNER JOIN SubOperacoes SO
			                ON NF.Operacao = SO.Operacao_Id 
			                AND NF.SubOperacao = SO.SubOperacoes_Id
		                INNER JOIN Produtos PRO
			                ON NFI.Produto_Id = PRO.Produto_Id
		                INNER JOIN Seguimentos SEG
			                ON PRO.Seguimento = SEG.Seguimento_Id
                            AND PRO.Dashboard = 1
		                WHERE 
			                ISNULL(NF.Situacao, 1)  = 1
			                AND (NF.Empresa_Id + '|' + CONVERT(VARCHAR, NF.EndEmpresa_Id)) IN (" & IIf(String.IsNullOrEmpty(filtro.EmpresasSelecionada), "(NF.Empresa_Id + '|' + NF.EndEmpresa_Id)", sqlCondicaoIN) & ")
			                AND NF.Movimento        >= '" & filtro.DataInicio.ToString("yyyy-MM-dd") & "'
			                AND NF.Movimento        <= '" & filtro.DataFim.ToString("yyyy-MM-dd") & "'
                            AND NCM IN ('23099010', '23091000')
			                AND SO.Classe IN ('VENDAS', 'VENDASAORDEM', 'EXPORTACOES', 'REMESSAS')
			                " & filtroSeguimento & "
                            AND LEFT(NF.Cliente_Id, 8) NOT IN ('49673784', '40938762')
	                )  

                    SELECT  ISNULL(SUM(CASE WHEN DEVOLUCAO = 'S' THEN -1 ELSE 1 END * PesoFiscal), 0) AS Quantidade,
                            ISNULL(SUM(CASE WHEN DEVOLUCAO = 'S' THEN -1 ELSE 1 END * ValorLiquido), 0) AS ValorTotal
                    FROM total_vendas; "

        Dim banco As New AcessaBanco
        Dim dsDashboard As DataSet = banco.ConsultaDataSet(sql, "NotasFiscais")

        Dim produto As New ProdutoDashboard

        For Each row As DataRow In dsDashboard.Tables(0).Rows

            produto.QuantidadeFiscal = row("Quantidade")
            produto.Valor = row("ValorTotal")

        Next

        Return produto

    End Function

    Public Function CarregarEvolucaoConsolidado(ByVal filtro As ProdutoDashboardFiltro) As IEnumerable(Of ProdutoDashboard)

        Dim sql As String = String.Empty

        Dim sqlCondicaoIN As String = ""

        If Not String.IsNullOrEmpty(filtro.EmpresasSelecionada) Then
            Dim listaFormatada = filtro.EmpresasSelecionada.Split(","c).
                                    Where(Function(s) Not String.IsNullOrWhiteSpace(s)).
                                    Select(Function(par) $"'{par.Trim()}'")

            sqlCondicaoIN = String.Join(",", listaFormatada)

        End If

        Dim filtroSeguimento As String = IIf(filtro.SeguimentoSelecionado <> "0", String.Format("AND PRO.Seguimento = {0} ", filtro.SeguimentoSelecionado), "")

        sql = " WITH total_vendas AS 
	                (
		                SELECT 	
			                NF.Empresa_Id,
			                NF.Cliente_Id,
			                NF.Serie_Id,
			                NF.Nota_Id,
			                NF.Movimento AS Data,
							NF.Pedido,
			                NFI.Produto_Id,
			                PRO.Nome AS ProdutoNome,
			                SEG.Descricao AS Seguimento,
			                NF.Operacao AS OperacaoNota,
			                NF.SubOperacao AS SubOperacaoNota,
			                NFI.PesoFiscal AS QuantidadeFiscal,
			                NFI.PesoFiscal,
			                NFI.Unitario,
			                NFE_LIQUIDO.Valor AS ValorLiquido, --FORMAT(NFE_LIQUIDO.Valor, 'N2', 'pt-br') AS ValorLiquido,
			                NFE_PRO.Valor AS ValorTotal, --FORMAT(NFE_PRO.Valor, 'N2', 'pt-br') AS ValorTotal
                            SO.Devolucao
		                FROM NotasFiscais NF
		                INNER JOIN NotasFiscaisXItens NFI 
			                ON NF.Empresa_Id				= NFI.Empresa_Id 
			                AND NF.EndEmpresa_Id			= NFI.EndEmpresa_Id 
			                AND NF.Cliente_Id				= NFI.Cliente_Id 
			                AND NF.EndCliente_Id			= NFI.EndCliente_Id 
			                AND NF.EntradaSaida_Id			= NFI.EntradaSaida_Id 
			                AND NF.Serie_Id					= NFI.Serie_Id 
			                AND NF.Nota_Id					= NFI.Nota_Id
		                INNER JOIN NotasFiscaisXEncargos NFE_LIQUIDO
			                ON NF.Empresa_Id				= NFE_LIQUIDO.Empresa_Id 
			                AND NF.EndEmpresa_Id			= NFE_LIQUIDO.EndEmpresa_Id 
			                AND NF.Cliente_Id				= NFE_LIQUIDO.Cliente_Id 
			                AND NF.EndCliente_Id			= NFE_LIQUIDO.EndCliente_Id 
			                AND NF.EntradaSaida_Id			= NFE_LIQUIDO.EntradaSaida_Id 
			                AND NF.Serie_Id					= NFE_LIQUIDO.Serie_Id 
			                AND NF.Nota_Id					= NFE_LIQUIDO.Nota_Id
			                AND NFI.Produto_Id				= NFE_LIQUIDO.Produto_Id
			                AND NFI.Sequencia_Id			= NFE_LIQUIDO.Sequencia_id
			                AND NFE_LIQUIDO.Encargo_Id		= 'LIQUIDO'
		                INNER JOIN NotasFiscaisXEncargos NFE_PRO
			                ON NF.Empresa_Id				= NFE_PRO.Empresa_Id 
			                AND NF.EndEmpresa_Id			= NFE_PRO.EndEmpresa_Id 
			                AND NF.Cliente_Id				= NFE_PRO.Cliente_Id 
			                AND NF.EndCliente_Id			= NFE_PRO.EndCliente_Id 
			                AND NF.EntradaSaida_Id			= NFE_PRO.EntradaSaida_Id 
			                AND NF.Serie_Id					= NFE_PRO.Serie_Id 
			                AND NF.Nota_Id					= NFE_PRO.Nota_Id
			                AND NFI.Produto_Id				= NFE_PRO.Produto_Id
			                AND NFI.Sequencia_Id			= NFE_PRO.Sequencia_id
			                AND NFE_PRO.Encargo_Id			= 'PRODUTO'
		                INNER JOIN SubOperacoes SO
			                ON NF.Operacao = SO.Operacao_Id 
			                AND NF.SubOperacao = SO.SubOperacoes_Id
		                INNER JOIN Produtos PRO
			                ON NFI.Produto_Id = PRO.Produto_Id
		                INNER JOIN Seguimentos SEG
			                ON PRO.Seguimento = SEG.Seguimento_Id
                            AND PRO.Dashboard = 1
		                WHERE 
			                ISNULL(NF.Situacao, 1)  = 1
			                AND (NF.Empresa_Id + '|' + CONVERT(VARCHAR, NF.EndEmpresa_Id)) IN (" & IIf(String.IsNullOrEmpty(filtro.EmpresasSelecionada), "(NF.Empresa_Id + '|' + NF.EndEmpresa_Id)", sqlCondicaoIN) & ")
			                AND NF.Movimento        >= '" & filtro.DataInicio.ToString("yyyy-MM-dd") & "'
			                AND NF.Movimento        <= '" & filtro.DataFim.ToString("yyyy-MM-dd") & "'
                            AND NCM IN ('23099010', '23091000')
			                AND SO.Classe IN ('VENDAS', 'VENDASAORDEM', 'EXPORTACOES', 'REMESSAS')
			                " & filtroSeguimento & "
                            AND LEFT(NF.Cliente_Id, 8) NOT IN ('49673784', '40938762')
	                )  

                    SELECT 
                        CONCAT('Q', DATEPART(QUARTER, Data), ' ', DATEPART(YEAR, Data)) AS Periodo,
                        SUM(CASE WHEN DEVOLUCAO = 'S' THEN -1 ELSE 1 END * QuantidadeFiscal) AS VolumeKg,
                        SUM(CASE WHEN DEVOLUCAO = 'S' THEN -1 ELSE 1 END * ValorTotal) AS Faturamento
                    FROM total_vendas
                    GROUP BY 
                        DATEPART(YEAR, Data), 
                        DATEPART(QUARTER, Data)
                    ORDER BY 
                        DATEPART(YEAR, Data), 
                        DATEPART(QUARTER, Data); "

        Dim banco As New AcessaBanco
        Dim dsDashboard As DataSet = banco.ConsultaDataSet(sql, "NotasFiscais")

        Dim listaProdutos As New List(Of ProdutoDashboard)

        For Each row As DataRow In dsDashboard.Tables(0).Rows

            Dim produto As New ProdutoDashboard

            produto.Periodo = row("Periodo")
            produto.VolumeKg = row("VolumeKg")
            produto.Faturamento = row("Faturamento")

            listaProdutos.Add(produto)

        Next

        Return listaProdutos

    End Function

    Public Function CarregarDetalhamentoConsolidado(ByVal filtro As ProdutoDashboardFiltro) As IEnumerable(Of ProdutoDashboard)

        Dim sql As String = String.Empty

        Dim sqlCondicaoIN As String = ""

        If Not String.IsNullOrEmpty(filtro.EmpresasSelecionada) Then
            Dim listaFormatada = filtro.EmpresasSelecionada.Split(","c).
                                    Where(Function(s) Not String.IsNullOrWhiteSpace(s)).
                                    Select(Function(par) $"'{par.Trim()}'")

            sqlCondicaoIN = String.Join(",", listaFormatada)

        End If

        Dim filtroSeguimento As String = IIf(filtro.SeguimentoSelecionado <> "0", String.Format("AND PRO.Seguimento = {0} ", filtro.SeguimentoSelecionado), "")

        sql = " WITH total_vendas AS 
	                (
		                SELECT 	
			                NF.Empresa_Id,
			                NF.Cliente_Id,
			                NF.Serie_Id,
			                NF.Nota_Id,
			                NF.Movimento AS Data,
							NF.Pedido,
			                NFI.Produto_Id,
			                PRO.Nome AS ProdutoNome,
			                SEG.Descricao AS Seguimento,
			                NF.Operacao AS OperacaoNota,
			                NF.SubOperacao AS SubOperacaoNota,
			                NFI.PesoFiscal AS QuantidadeFiscal,
			                NFI.PesoFiscal,
			                NFI.Unitario,
			                NFE_LIQUIDO.Valor AS ValorLiquido, --FORMAT(NFE_LIQUIDO.Valor, 'N2', 'pt-br') AS ValorLiquido,
			                NFE_PRO.Valor AS ValorTotal, --FORMAT(NFE_PRO.Valor, 'N2', 'pt-br') AS ValorTotal
                            SO.Devolucao
		                FROM NotasFiscais NF
		                INNER JOIN NotasFiscaisXItens NFI 
			                ON NF.Empresa_Id				= NFI.Empresa_Id 
			                AND NF.EndEmpresa_Id			= NFI.EndEmpresa_Id 
			                AND NF.Cliente_Id				= NFI.Cliente_Id 
			                AND NF.EndCliente_Id			= NFI.EndCliente_Id 
			                AND NF.EntradaSaida_Id			= NFI.EntradaSaida_Id 
			                AND NF.Serie_Id					= NFI.Serie_Id 
			                AND NF.Nota_Id					= NFI.Nota_Id
		                INNER JOIN NotasFiscaisXEncargos NFE_LIQUIDO
			                ON NF.Empresa_Id				= NFE_LIQUIDO.Empresa_Id 
			                AND NF.EndEmpresa_Id			= NFE_LIQUIDO.EndEmpresa_Id 
			                AND NF.Cliente_Id				= NFE_LIQUIDO.Cliente_Id 
			                AND NF.EndCliente_Id			= NFE_LIQUIDO.EndCliente_Id 
			                AND NF.EntradaSaida_Id			= NFE_LIQUIDO.EntradaSaida_Id 
			                AND NF.Serie_Id					= NFE_LIQUIDO.Serie_Id 
			                AND NF.Nota_Id					= NFE_LIQUIDO.Nota_Id
			                AND NFI.Produto_Id				= NFE_LIQUIDO.Produto_Id
			                AND NFI.Sequencia_Id			= NFE_LIQUIDO.Sequencia_id
			                AND NFE_LIQUIDO.Encargo_Id		= 'LIQUIDO'
		                INNER JOIN NotasFiscaisXEncargos NFE_PRO
			                ON NF.Empresa_Id				= NFE_PRO.Empresa_Id 
			                AND NF.EndEmpresa_Id			= NFE_PRO.EndEmpresa_Id 
			                AND NF.Cliente_Id				= NFE_PRO.Cliente_Id 
			                AND NF.EndCliente_Id			= NFE_PRO.EndCliente_Id 
			                AND NF.EntradaSaida_Id			= NFE_PRO.EntradaSaida_Id 
			                AND NF.Serie_Id					= NFE_PRO.Serie_Id 
			                AND NF.Nota_Id					= NFE_PRO.Nota_Id
			                AND NFI.Produto_Id				= NFE_PRO.Produto_Id
			                AND NFI.Sequencia_Id			= NFE_PRO.Sequencia_id
			                AND NFE_PRO.Encargo_Id			= 'PRODUTO'
		                INNER JOIN SubOperacoes SO
			                ON NF.Operacao = SO.Operacao_Id 
			                AND NF.SubOperacao = SO.SubOperacoes_Id
		                INNER JOIN Produtos PRO
			                ON NFI.Produto_Id = PRO.Produto_Id
		                INNER JOIN Seguimentos SEG
			                ON PRO.Seguimento = SEG.Seguimento_Id
                            AND PRO.Dashboard = 1
		                WHERE 
			                ISNULL(NF.Situacao, 1)  = 1
			                AND (NF.Empresa_Id + '|' + CONVERT(VARCHAR, NF.EndEmpresa_Id)) IN (" & IIf(String.IsNullOrEmpty(filtro.EmpresasSelecionada), "(NF.Empresa_Id + '|' + NF.EndEmpresa_Id)", sqlCondicaoIN) & ")
			                AND NF.Movimento        >= '" & filtro.DataInicio.ToString("yyyy-MM-dd") & "'
			                AND NF.Movimento        <= '" & filtro.DataFim.ToString("yyyy-MM-dd") & "'
                            AND NCM IN ('23099010', '23091000')
			                AND SO.Classe IN ('VENDAS', 'VENDASAORDEM', 'EXPORTACOES', 'REMESSAS')
			                " & filtroSeguimento & "
                            AND LEFT(NF.Cliente_Id, 8) NOT IN ('49673784', '40938762')
	                ),
                    resumo AS (
                    SELECT 
                        1 AS Ordem,
                        'Hoje' AS Periodo,
                        COUNT(DISTINCT Pedido) AS Pedidos,
                        ISNULL(SUM(CASE WHEN DEVOLUCAO = 'S' THEN -1 ELSE 1 END * QuantidadeFiscal), 0) AS VolumeKg,
                        ISNULL(SUM(CASE WHEN DEVOLUCAO = 'S' THEN -1 ELSE 1 END * ValorTotal), 0) AS Faturamento,
                        ISNULL(AVG(CASE WHEN DEVOLUCAO = 'S' THEN -1 ELSE 1 END * QuantidadeFiscal), 0) AS VolumeMedio,
                        ISNULL(AVG(CASE WHEN DEVOLUCAO = 'S' THEN -1 ELSE 1 END * ValorTotal), 0) AS FaturamentoMedio
                    FROM total_vendas
                    WHERE CAST(Data AS DATE) = CAST(GETDATE() AS DATE)

                    UNION ALL

                    SELECT 
                        2,
                        'Última Semana',
                        COUNT(DISTINCT Pedido),
                        ISNULL(SUM(CASE WHEN DEVOLUCAO = 'S' THEN -1 ELSE 1 END * QuantidadeFiscal), 0),
                        ISNULL(SUM(CASE WHEN DEVOLUCAO = 'S' THEN -1 ELSE 1 END * ValorTotal), 0),
                        ISNULL(AVG(CASE WHEN DEVOLUCAO = 'S' THEN -1 ELSE 1 END * QuantidadeFiscal), 0),
                        ISNULL(AVG(CASE WHEN DEVOLUCAO = 'S' THEN -1 ELSE 1 END * ValorTotal), 0)
                    FROM total_vendas
                    WHERE Data >= DATEADD(DAY, -7, GETDATE())

                    UNION ALL

                    SELECT 
                        3,
                        'Último Mês',
                        COUNT(DISTINCT Pedido),
                        ISNULL(SUM(CASE WHEN DEVOLUCAO = 'S' THEN -1 ELSE 1 END * QuantidadeFiscal), 0),
                        ISNULL(SUM(CASE WHEN DEVOLUCAO = 'S' THEN -1 ELSE 1 END * ValorTotal), 0),
                        ISNULL(AVG(CASE WHEN DEVOLUCAO = 'S' THEN -1 ELSE 1 END * QuantidadeFiscal), 0),
                        ISNULL(AVG(CASE WHEN DEVOLUCAO = 'S' THEN -1 ELSE 1 END * ValorTotal), 0)
                    FROM total_vendas
                    WHERE Data >= DATEADD(MONTH, -1, GETDATE())

                    UNION ALL

                    SELECT 
                        4,
                        'Último Trimestre',
                        COUNT(DISTINCT Pedido),
                        ISNULL(SUM(CASE WHEN DEVOLUCAO = 'S' THEN -1 ELSE 1 END * QuantidadeFiscal), 0),
                        ISNULL(SUM(CASE WHEN DEVOLUCAO = 'S' THEN -1 ELSE 1 END * ValorTotal), 0),
                        ISNULL(AVG(CASE WHEN DEVOLUCAO = 'S' THEN -1 ELSE 1 END * QuantidadeFiscal), 0),
                        ISNULL(AVG(CASE WHEN DEVOLUCAO = 'S' THEN -1 ELSE 1 END * ValorTotal), 0)
                    FROM total_vendas
                    WHERE Data >= DATEADD(QUARTER, -1, GETDATE())

                    UNION ALL

                    SELECT 
                        5,
                        'Ano Atual',
                        COUNT(DISTINCT Pedido),
                        ISNULL(SUM(CASE WHEN DEVOLUCAO = 'S' THEN -1 ELSE 1 END * QuantidadeFiscal), 0),
                        ISNULL(SUM(CASE WHEN DEVOLUCAO = 'S' THEN -1 ELSE 1 END * ValorTotal), 0),
                        ISNULL(AVG(CASE WHEN DEVOLUCAO = 'S' THEN -1 ELSE 1 END * QuantidadeFiscal), 0),
                        ISNULL(AVG(CASE WHEN DEVOLUCAO = 'S' THEN -1 ELSE 1 END * ValorTotal), 0)
                    FROM total_vendas
                    WHERE YEAR(Data) = YEAR(GETDATE())
                )

                SELECT 
                    *,
                    RANK() OVER (ORDER BY Faturamento DESC) AS Ranking
                FROM resumo
                ORDER BY Ordem; "

        Dim banco As New AcessaBanco
        Dim dsDashboard As DataSet = banco.ConsultaDataSet(sql, "NotasFiscais")

        Dim listaProdutos As New List(Of ProdutoDashboard)

        For Each row As DataRow In dsDashboard.Tables(0).Rows

            Dim produto As New ProdutoDashboard

            produto.Periodo = row("Periodo")
            produto.Pedidos = row("Pedidos")
            produto.VolumeKg = row("VolumeKg")
            produto.Faturamento = row("Faturamento")
            produto.VolumeMedio = row("VolumeMedio")
            produto.FaturamentoMedio = row("FaturamentoMedio")

            listaProdutos.Add(produto)

        Next

        Return listaProdutos

    End Function

    Public Function CarregarComparacaoVolumeXFaturamentoConsolidado(ByVal filtro As ProdutoDashboardFiltro) As IEnumerable(Of ProdutoDashboard)

        Dim sql As String = String.Empty

        Dim sqlCondicaoIN As String = ""

        If Not String.IsNullOrEmpty(filtro.EmpresasSelecionada) Then
            Dim listaFormatada = filtro.EmpresasSelecionada.Split(","c).
                                    Where(Function(s) Not String.IsNullOrWhiteSpace(s)).
                                    Select(Function(par) $"'{par.Trim()}'")

            sqlCondicaoIN = String.Join(",", listaFormatada)

        End If

        Dim filtroSeguimento As String = IIf(filtro.SeguimentoSelecionado <> "0", String.Format("AND PRO.Seguimento = {0} ", filtro.SeguimentoSelecionado), "")

        sql = " WITH total_vendas AS 
	                (
		                SELECT 	
			                NF.Empresa_Id,
			                NF.Cliente_Id,
			                NF.Serie_Id,
			                NF.Nota_Id,
			                NF.Movimento AS Data,
							NF.Pedido,
			                NFI.Produto_Id,
			                PRO.Nome AS ProdutoNome,
			                SEG.Descricao AS Seguimento,
			                NF.Operacao AS OperacaoNota,
			                NF.SubOperacao AS SubOperacaoNota,
			                NFI.PesoFiscal AS QuantidadeFiscal,
			                NFI.PesoFiscal,
			                NFI.Unitario,
			                NFE_LIQUIDO.Valor AS ValorLiquido, --FORMAT(NFE_LIQUIDO.Valor, 'N2', 'pt-br') AS ValorLiquido,
			                NFE_PRO.Valor AS ValorTotal, --FORMAT(NFE_PRO.Valor, 'N2', 'pt-br') AS ValorTotal
                            SO.Devolucao
		                FROM NotasFiscais NF
		                INNER JOIN NotasFiscaisXItens NFI 
			                ON NF.Empresa_Id				= NFI.Empresa_Id 
			                AND NF.EndEmpresa_Id			= NFI.EndEmpresa_Id 
			                AND NF.Cliente_Id				= NFI.Cliente_Id 
			                AND NF.EndCliente_Id			= NFI.EndCliente_Id 
			                AND NF.EntradaSaida_Id			= NFI.EntradaSaida_Id 
			                AND NF.Serie_Id					= NFI.Serie_Id 
			                AND NF.Nota_Id					= NFI.Nota_Id
		                INNER JOIN NotasFiscaisXEncargos NFE_LIQUIDO
			                ON NF.Empresa_Id				= NFE_LIQUIDO.Empresa_Id 
			                AND NF.EndEmpresa_Id			= NFE_LIQUIDO.EndEmpresa_Id 
			                AND NF.Cliente_Id				= NFE_LIQUIDO.Cliente_Id 
			                AND NF.EndCliente_Id			= NFE_LIQUIDO.EndCliente_Id 
			                AND NF.EntradaSaida_Id			= NFE_LIQUIDO.EntradaSaida_Id 
			                AND NF.Serie_Id					= NFE_LIQUIDO.Serie_Id 
			                AND NF.Nota_Id					= NFE_LIQUIDO.Nota_Id
			                AND NFI.Produto_Id				= NFE_LIQUIDO.Produto_Id
			                AND NFI.Sequencia_Id			= NFE_LIQUIDO.Sequencia_id
			                AND NFE_LIQUIDO.Encargo_Id		= 'LIQUIDO'
		                INNER JOIN NotasFiscaisXEncargos NFE_PRO
			                ON NF.Empresa_Id				= NFE_PRO.Empresa_Id 
			                AND NF.EndEmpresa_Id			= NFE_PRO.EndEmpresa_Id 
			                AND NF.Cliente_Id				= NFE_PRO.Cliente_Id 
			                AND NF.EndCliente_Id			= NFE_PRO.EndCliente_Id 
			                AND NF.EntradaSaida_Id			= NFE_PRO.EntradaSaida_Id 
			                AND NF.Serie_Id					= NFE_PRO.Serie_Id 
			                AND NF.Nota_Id					= NFE_PRO.Nota_Id
			                AND NFI.Produto_Id				= NFE_PRO.Produto_Id
			                AND NFI.Sequencia_Id			= NFE_PRO.Sequencia_id
			                AND NFE_PRO.Encargo_Id			= 'PRODUTO'
		                INNER JOIN SubOperacoes SO
			                ON NF.Operacao = SO.Operacao_Id 
			                AND NF.SubOperacao = SO.SubOperacoes_Id
		                INNER JOIN Produtos PRO
			                ON NFI.Produto_Id = PRO.Produto_Id
		                INNER JOIN Seguimentos SEG
			                ON PRO.Seguimento = SEG.Seguimento_Id
                            AND PRO.Dashboard = 1
		                WHERE 
			                ISNULL(NF.Situacao, 1)  = 1
			                AND (NF.Empresa_Id + '|' + CONVERT(VARCHAR, NF.EndEmpresa_Id)) IN (" & IIf(String.IsNullOrEmpty(filtro.EmpresasSelecionada), "(NF.Empresa_Id + '|' + NF.EndEmpresa_Id)", sqlCondicaoIN) & ")
			                AND NF.Movimento        >= '" & filtro.DataInicio.ToString("yyyy-MM-dd") & "'
			                AND NF.Movimento        <= '" & filtro.DataFim.ToString("yyyy-MM-dd") & "'
                            AND NCM IN ('23099010', '23091000')
			                AND SO.Classe IN ('VENDAS', 'VENDASAORDEM', 'EXPORTACOES', 'REMESSAS')
			                " & filtroSeguimento & "
                            AND LEFT(NF.Cliente_Id, 8) NOT IN ('49673784', '40938762')
	                )  

                    SELECT 
                        CONCAT('Q', DATEPART(QUARTER, Data), ' ', DATEPART(YEAR, Data)) AS Periodo,
                        SUM(CASE WHEN DEVOLUCAO = 'S' THEN -1 ELSE 1 END * QuantidadeFiscal) AS VolumeKg,
                        SUM(CASE WHEN DEVOLUCAO = 'S' THEN -1 ELSE 1 END * ValorTotal) AS Faturamento
                    FROM total_vendas
                    GROUP BY 
                        DATEPART(YEAR, Data), 
                        DATEPART(QUARTER, Data)
                    ORDER BY 
                        DATEPART(YEAR, Data), 
                        DATEPART(QUARTER, Data); "

        Dim banco As New AcessaBanco
        Dim dsDashboard As DataSet = banco.ConsultaDataSet(sql, "NotasFiscais")

        Dim listaProdutos As New List(Of ProdutoDashboard)

        For Each row As DataRow In dsDashboard.Tables(0).Rows

            Dim produto As New ProdutoDashboard

            produto.Periodo = row("Periodo")
            produto.VolumeKg = row("VolumeKg")
            produto.Faturamento = row("Faturamento")

            listaProdutos.Add(produto)

        Next

        Return listaProdutos

    End Function

#End Region

End Class
