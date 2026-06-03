Imports System.Data
Imports System.Data.OleDb
Imports System.Data.SqlClient
Imports System.Data.SqlTypes
Imports System.Drawing
Imports System.IO
Imports System.Security.Cryptography
Imports System.Security.Principal
Imports BoletoNet
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.[Shared].Json
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis
Imports OfficeOpenXml
Imports OfficeOpenXml.Style
Imports ThoughtWorks.QRCode.Codec

Public Class panteste
    Inherits BasePage

    Dim objCliente As New Cliente
    Dim objListaCliente As New ListCliente

    Private ObjOxE As [Lib].Negocio.OperacaoXEstado

    Private SqlArray As New ArrayList
    Private sql As String = String.Empty
    'Private sConnStringDestino As String = "Provider=SqlOleDb;" &
    '                       "Data Source=MrFurlan;" &
    '                       "Initial Catalog=Horus;" &
    '                       "User Id=sa;" &
    '                       "Password=pwd_ngs123"

    'Private sConnStringDestino As String = "Provider=SqlOleDb;" & _
    '                       "Data Source=MrFurlan;" & _
    '                       "Initial Catalog=Insol;" & _
    '                       "User Id=sa;" & _
    '                       "Password=pwd_ngs123"

    'Private sConnStringDestino As String = "Provider=SqlOleDb;" & _
    '                       "Data Source=MrFurlan;" & _
    '                       "Initial Catalog=Gerencial;" & _
    '                       "User Id=sa;" & _
    '                       "Password=pwd_ngs123"

    'Private sConnStringDestino As String = "Provider=SqlOleDb;" &
    '                       "Data Source=SRVNGS;" &
    '                       "Initial Catalog=Horus;" &
    '                       "User Id=n@$;" &
    '                       "Password=pwd_ngs123"

    'CORRETO - USAR ESSE
    Private sConnStringDestino As String = "Provider=SqlOleDb;" &
                           "Data Source=DESKTOP-TSAGGN1;" &
                           "Initial Catalog=Gerencial;" &
                           "User Id=n@$;" &
                           "Password=pwd_curtume"

    'Private sConnStringDestino As String = "Provider=SqlOleDb;" & _
    '                       "Data Source=SrvQuimica;" & _
    '                       "Initial Catalog=Quimica;" & _
    '                       "User Id=sa;" & _
    '                       "Password=QuI21q#p11Xxa"

    'Private sConnStringDestino As String = "Provider=SqlOleDb;" & _
    '                   "Data Source=AlvoradaSrv;" & _
    '                   "Initial Catalog=Alvorada;" & _
    '                   "User Id=sa;" & _
    '                   "Password=YKb21q#p11Xxa"

    'Private sConnStringDestino As String = "Provider=SqlOleDb;" &
    '                   "Data Source=Fronteira;" &
    '                   "Initial Catalog=Fronteira;" &
    '                   "User Id=sa;" &
    '                   "Password=FronteIra42q#f22Xxn"

    'Private sConnStringDestino As String = "Provider=SqlOleDb;" &
    '                       "Data Source=SRVNGS;" &
    '                       "Initial Catalog=Orix;" &
    '                       "User Id=n@$;" &
    '                       "Password=pwd_ngs123"

    'Private sConnStringDestino As String = "Provider=SqlOleDb;" &
    '                       "Data Source=SRVNUTRI;" &
    '                       "Initial Catalog=Nutri;" &
    '                       "User Id=sa;" &
    '                       "Password=NutrI20q#p11Xxa"

    'Private sConnStringDestino As String = "Provider=SqlOleDb;" &
    '                       "Data Source=SRVBAXI;" &
    '                       "Initial Catalog=Baxi;" &
    '                       "User Id=sa;" &
    '                       "Password=BaxiH24q#p11Zxt"

    'Private sConnStringDestino As String = "Provider=SqlOleDb;" &
    '                       "Data Source=SRVNUBA;" &
    '                       "Initial Catalog=Nuba;" &
    '                       "User Id=n@$;" &
    '                       "Password=NbrBe15q#A33Uxn"


    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not IsPostBack And IsConnect Then
            If Session("ssNomeUsuario") = "JABER" OrElse Session("ssNomeUsuario") = "FURLAN" OrElse Session("ssNomeUsuario") = "AUTOMATICO" Then
                ddl.Carregar(ddlAno, CarregarDDL.Tabela.Ano, "2016;10;C", False)
                ddlAno.SelectedValue = Now.Year
                ddlMes.SelectedValue = IIf(Now.Month = 1, 12, Now.Month - 1)

                chkTNotas.Visible = False

                If Session("ssNomeUsuario") = "FURLAN" OrElse Session("ssNomeUsuario") = "AUTOMATICO" Then chkTNotas.Visible = True
                If Session("ssNomeUsuario") = "FURLAN" OrElse Session("ssNomeUsuario") = "AUTOMATICO" Then chkLerPlanilha.Visible = True
                If Session("ssNomeUsuario") = "FURLAN" OrElse Session("ssNomeUsuario") = "AUTOMATICO" Then rExcel.Visible = True

            Else
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página", "~/Gestao.aspx")
                Exit Sub
            End If
        End If
    End Sub

    Private Function GravaBancoDestino(ByVal sql As ArrayList, ByVal strConexao As String) As Boolean
        Dim SqlTrans As OleDb.OleDbTransaction
        Dim Sqlcommand As New OleDb.OleDbCommand
        Dim SQlConnectionDestino As New OleDb.OleDbConnection

        'Dim strm As New StreamWriter(HttpContext.Current.Server.MapPath("~/Files/tabelas.txt"))

        Try
            If (SQlConnectionDestino.State = ConnectionState.Closed) Then
                SQlConnectionDestino = New OleDb.OleDbConnection(strConexao)
                SQlConnectionDestino.Open()
            End If
            SqlTrans = SQlConnectionDestino.BeginTransaction(IsolationLevel.ReadCommitted)
            Sqlcommand.Connection = SQlConnectionDestino
            Sqlcommand.Transaction = SqlTrans
            For index = 0 To sql.Count - 1
                Sqlcommand.CommandText = CStr(sql.Item(index))
                'strm.WriteLine(CStr(sql.Item(index)))
                Sqlcommand.ExecuteNonQuery()
            Next
            'strm.Close()
            SqlTrans.Commit()
            'SQlConnectionDestino.Close()
        Catch SqlException As OleDb.OleDbException
            'strm.Close()
            If SqlTrans IsNot Nothing Then
                SqlTrans.Rollback()
            End If

            'SqlTrans.Rollback()

            Select Case SqlException.ErrorCode
                Case 2627
                    Throw New Exception("Registro já cadastrado!")
                Case 647
                    Throw New Exception("Registro esta sendo utilizado em outras tabelas!")
                Case Else
                    Throw New Exception(SqlException.Message)
            End Select
        Catch ex As Exception
            If SqlTrans IsNot Nothing Then
                SqlTrans.Rollback()
            End If

            Throw New Exception(ex.Message)
        Finally
            SQlConnectionDestino.Close()
        End Try

        Return True
    End Function

    Private Function gravaPlanoDeContas(ByRef mensagem As String) As Boolean
        SqlArray.Clear()

        Try
            sql = "DELETE FROM Gerencial.dbo.PlanoDeContas"
            SqlArray.Add(sql)

            sql = "insert into Gerencial.dbo.PlanoDeContas (Empresa_Id, EndEmpresa_Id, Conta_Id, Titulo, Cliente, Produto, CentroDeCusto, TipodeCliente, Responsabilidade, ContaOrcamentaria, ContaAnterior, Dacon, Pagar," & vbCrLf &
                  "                                         Receber, TipoDeConta, Adiantamento, Pedido, BaixaAdiantamento, ContaBaixaAdiantamento, Encargo, TemEncargo, TipoDeCusto, AdiantamentoSoContabil, Ativo," & vbCrLf &
                  "                                         UsuarioInclusao, UsuarioInclusaoData, UsuarioAlteracao, UsuarioAlteracaoData, Referencial, CodigoDeCustoECF)" & vbCrLf &
                  "SELECT Empresa_Id, EndEmpresa_Id, Conta_Id, Titulo, Cliente, Produto, CentroDeCusto, TipodeCliente, Responsabilidade, ContaOrcamentaria, ContaAnterior, Dacon, Pagar," & vbCrLf &
                  "       Receber, TipoDeConta, Adiantamento, Pedido, BaixaAdiantamento, ContaBaixaAdiantamento, Encargo, TemEncargo, TipoDeCusto, AdiantamentoSoContabil, Ativo," & vbCrLf &
                  "       UsuarioInclusao, UsuarioInclusaoData, UsuarioAlteracao, UsuarioAlteracaoData, Referencial, CodigoDeCustoECF" & vbCrLf &
                  "FROM   Panorama.dbo.PlanoDeContas"
            SqlArray.Add(sql)

            If SqlArray.Count > 0 Then
                If GravaBancoDestino(SqlArray, sConnStringDestino) Then
                    Return True
                Else
                    Return False
                End If
            Else
                mensagem = "Plano de Contas sem registros para Transferência"
                Return True
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
            Return False
        End Try
    End Function

    Private Function gravaProduto(ByRef mensagem As String) As Boolean
        SqlArray.Clear()

        Try
            sql = "DELETE FROM ProdutoXEmbalagem"
            SqlArray.Add(sql)

            sql = "DELETE FROM ProdutosxUnidadeDeComercializacao"
            SqlArray.Add(sql)

            sql = "DELETE FROM ProdutosAgrupados"
            SqlArray.Add(sql)

            sql = "DELETE FROM Produtos"
            SqlArray.Add(sql)

            sql = "DELETE FROM PlanoDeContasXReferencialBacen"
            SqlArray.Add(sql)

            sql = "DELETE FROM GruposDeEstoques"
            SqlArray.Add(sql)

            sql = "insert into Gerencial.dbo.GruposDeEstoques (Grupo_Id, Descricao, Custo, AgruparFinanceiro, MapaDeEstoque)" & vbCrLf &
                  "SELECT     Grupo_Id, Descricao, Custo, AgruparFinanceiro, MapaDeEstoque" & vbCrLf &
                  "FROM         Panorama.dbo.GruposDeEstoques"
            SqlArray.Add(sql)

            sql = "insert into Gerencial.dbo.Produtos (Produto_Id, Grupo, Unidade, Etapa, Situacao, Embalagem, NCM, Nome, Descricao, DescricaoMapa, BaseDeCalculo, PesoQuantidade, EstoqueMinimo, Agrupar, " & vbCrLf &
                  "                      QuantidadeNaCaixa, Qualidade, IPI, IPITributado, PisCofinsIntegral, PisCofinsPresumido, CarteiraDeCompras, CarteiraDeVendas, ICMS, TipoDoItem, " & vbCrLf &
                  "                      CodigoDoGenero, CodigoEX, CodigoDoServico, ControlarEstoque, ControlarLote, Fitossanitario, ControlarEmbalagem, ProdutoIndea, Marca, ControlarPecas, " & vbCrLf &
                  "                      ControlarPrecoDePauta, SubCodigoDoGenero, ControlarRomaneio, ControlarPesagem, ControlarDecimais)" & vbCrLf &
                  "SELECT     Produto_Id, Grupo, Unidade, Etapa, Situacao, Embalagem, NCM, Nome, Descricao, DescricaoMapa, BaseDeCalculo, PesoQuantidade, EstoqueMinimo, Agrupar, " & vbCrLf &
                  "                      QuantidadeNaCaixa, Qualidade, IPI, IPITributado, PisCofinsIntegral, PisCofinsPresumido, CarteiraDeCompras, CarteiraDeVendas, ICMS, TipoDoItem, " & vbCrLf &
                  "                      CodigoDoGenero, CodigoEX, CodigoDoServico, ControlarEstoque, ControlarLote, Fitossanitario, ControlarEmbalagem, ProdutoIndea, Marca, ControlarPecas, " & vbCrLf &
                  "                      ControlarPrecoDePauta, SubCodigoDoGenero, ControlarRomaneio, ControlarPesagem, ControlarDecimais" & vbCrLf &
                  "FROM         Panorama.dbo.Produtos"
            SqlArray.Add(sql)

            sql = "insert into Gerencial.dbo.ProdutosxUnidadeDeComercializacao (Produto_id, Unidade_id, FatorConversao_id, Ativo, PesoDaEmbalagem)" & vbCrLf &
                  "SELECT     Produto_id, Unidade_id, FatorConversao_id, Ativo, 0" & vbCrLf &
                  "FROM         Panorama.dbo.ProdutosxUnidadeDeComercializacao"
            SqlArray.Add(sql)

            sql = "insert into Gerencial.dbo.ProdutoXEmbalagem (Produto_Id, Embalagem_Id, TipoDeEmbalagem_Id, CapacidadeEmbalagem_Id, PesoBruto, PesoLiquido, PesoVariavel)" & vbCrLf &
                  "SELECT     Produto_Id, Embalagem_Id, TipoDeEmbalagem_Id, CapacidadeEmbalagem_Id, PesoBruto, PesoLiquido, PesoVariavel" & vbCrLf &
                  "FROM         Panorama.dbo.ProdutoXEmbalagem"
            SqlArray.Add(sql)

            If SqlArray.Count > 0 Then
                If GravaBancoDestino(SqlArray, sConnStringDestino) Then
                    Return True
                Else
                    Return False
                End If
            Else
                mensagem = "Produtos sem registros para Transferência"
                Return True
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
            Return False
        End Try
    End Function

    Private Function gravaEncargos(ByRef mensagem As String) As Boolean
        SqlArray.Clear()

        Try
            sql = "DELETE FROM Gerencial.dbo.Encargos"
            SqlArray.Add(sql)

            sql = "INSERT INTO Gerencial.dbo.Encargos (Encargo_id, Descricao, ContaDebito, ContaCredito, BaseCalculo, Aliquota, OperacaoXEncargo, Etapa, Atualizacao, GravaCentroDeCusto, ImprimirNFE," & vbCrLf &
                  "            EncargoAgrupador, Operador, ValorOuPeso, PodeSofreRetencao, TipoDePessoa, VerificaEmpresa, TipoDePessoaRetencao)" & vbCrLf &
                  "SELECT     Encargo_id, Descricao, ContaDebito, ContaCredito, BaseCalculo, Aliquota, OperacaoXEncargo, Etapa, Atualizacao, GravaCentroDeCusto, ImprimirNFE," & vbCrLf &
                  "           EncargoAgrupador, Operador, ValorOuPeso, PodeSofreRetencao, TipoDePessoa, VerificaEmpresa, TipoDePessoaRetencao" & vbCrLf &
                  "FROM         Panorama.dbo.Encargos"
            SqlArray.Add(sql)

            If SqlArray.Count > 0 Then
                If GravaBancoDestino(SqlArray, sConnStringDestino) Then
                    Return True
                Else
                    Return False
                End If
            Else
                mensagem = "Encargos sem registros para Transferência"
                Return True
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
            Return False
        End Try
    End Function

    Private Function gravaOperacoes(ByRef mensagem As String) As Boolean
        SqlArray.Clear()

        Try
            sql = "DELETE FROM SubOperacoes"
            SqlArray.Add(sql)

            sql = "DELETE FROM Operacoes"
            SqlArray.Add(sql)

            sql = "insert into Gerencial.dbo.Operacoes (Operacao_Id, Descricao, Producao, Classe, UsuarioInclusao, UsuarioInclusaoData, UsuarioAlteracao, UsuarioAlteracaoData, UFDepositoDestino)" & vbCrLf &
                  "SELECT     Operacao_Id, Descricao, Producao, Classe, UsuarioInclusao, UsuarioInclusaoData, UsuarioAlteracao, UsuarioAlteracaoData, UFDepositoDestino" & vbCrLf &
                  "FROM         Panorama.dbo.Operacoes"
            SqlArray.Add(sql)

            sql = "insert into Gerencial.dbo.SubOperacoes (Operacao_Id, SubOperacoes_Id, Descricao, EntradaSaida, Devolucao, Classe, PrecoFixo, Laudo, EstoqueInicial, EstoqueFisico, EstoqueFiscal, QuantidadeFisico, " & vbCrLf &
                  "                      QuantidadeFiscal, QuantidadePedido, UnitarioPedido, GrupoDeContas, Financeiro, Contabil, ApuracaoDeCustos, ApuracaodeCustosContraPartida, Deposito, Situacao, " & vbCrLf &
                  "                      ControlarPecas, CobraServico, OperacaoDestino, SubOperacaoDestino, consignacao, Memorando, Pedido, Liminar, ProdutoDeTerceiro, ParaFinsDeExportacao, " & vbCrLf &
                  "                      FinalidadeDaNota, ContaDeAdiantamento, UsuarioInclusao, UsuarioInclusaoData, UsuarioAlteracao, UsuarioAlteracaoData, AmostraGratis, Representante) " & vbCrLf &
                  "SELECT     Operacao_Id, SubOperacoes_Id, Descricao, EntradaSaida, Devolucao, Classe, PrecoFixo, Laudo, EstoqueInicial, EstoqueFisico, EstoqueFiscal, QuantidadeFisico, " & vbCrLf &
                  "                      QuantidadeFiscal, QuantidadePedido, UnitarioPedido, GrupoDeContas, Financeiro, Contabil, ApuracaoDeCustos, ApuracaodeCustosContraPartida, Deposito, Situacao, " & vbCrLf &
                  "                      ControlarPecas, CobraServico, OperacaoDestino, SubOperacaoDestino, consignacao, Memorando, Pedido, Liminar, ProdutoDeTerceiro, ParaFinsDeExportacao, " & vbCrLf &
                  "                      FinalidadeDaNota, ContaDeAdiantamento, UsuarioInclusao, UsuarioInclusaoData, UsuarioAlteracao, UsuarioAlteracaoData, AmostraGratis, Representante" & vbCrLf &
                  "FROM         Panorama.dbo.SubOperacoes"
            SqlArray.Add(sql)

            If SqlArray.Count > 0 Then
                If GravaBancoDestino(SqlArray, sConnStringDestino) Then
                    Return True
                Else
                    Return False
                End If
            Else
                mensagem = "Operações sem registros para Transferência"
                Return True
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
            Return False
        End Try
    End Function

    Private Function gravaOpeXEncargos(ByRef mensagem As String) As Boolean
        SqlArray.Clear()

        Try
            sql = "DELETE FROM OperacaoXEstadoXEncargo"
            SqlArray.Add(sql)

            sql = "DELETE FROM OperacaoXEstado"
            SqlArray.Add(sql)

            sql = "insert into Gerencial.dbo.OperacaoXEstado (Codigo_Id, Ativo, InicioVigencia, GrupoProduto, Produto, Operacao, SubOperacao, EstadoOrigem, " & vbCrLf &
                   "            EstadoDestino, GrupoFiscal, CodigoFiscal, STICMS, ObsICMS, STIPI, STPISCOFINS, ObsPISCOFINS, UsuarioInclusao, UsuarioInclusaoData, Empresa) " & vbCrLf &
                   "SELECT Codigo_Id, Ativo, InicioVigencia, GrupoProduto, Produto, Operacao, SubOperacao, EstadoOrigem, EstadoDestino, GrupoFiscal, " & vbCrLf &
                   "       CodigoFiscal, STICMS, ObsICMS, STIPI, STPISCOFINS, ObsPISCOFINS, UsuarioInclusao, UsuarioInclusaoData, Empresa" & vbCrLf &
                   "FROM Panorama.dbo.OperacaoXEstado"
            SqlArray.Add(sql)

            sql = "insert into Gerencial.dbo.OperacaoXEstadoXEncargo (Codigo_Id, Encargo_Id, Sinal, DebitaConta, CreditaConta, " & vbCrLf &
                   "            AliquotaBase, Aliquota, AliquotaExibicao, AliquotaLimite) " & vbCrLf &
                   "SELECT     Codigo_Id, Encargo_Id, Sinal, DebitaConta, CreditaConta, AliquotaBase, Aliquota, AliquotaExibicao, AliquotaLimite " & vbCrLf &
                   "FROM         Panorama.dbo.OperacaoXEstadoXEncargo"
            SqlArray.Add(sql)

            If SqlArray.Count > 0 Then
                If GravaBancoDestino(SqlArray, sConnStringDestino) Then
                    Return True
                Else
                    Return False
                End If
            Else
                mensagem = "OperaçõesXEncargos sem registros para Transferência"
                Return True
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
            Return False
        End Try
    End Function

    Private Function gravaCarteiras(ByRef mensagem As String) As Boolean
        SqlArray.Clear()

        Try
            Dim listCareiras As New [Lib].Negocio.ListCarteiraFinanceira(True)

            sql = "DELETE FROM ComprasXProdutos"
            SqlArray.Add(sql)

            For Each carteira As [Lib].Negocio.CarteiraFinanceira In listCareiras
                carteira.IUD = "I"
                carteira.SalvarSql(SqlArray)
            Next

            If SqlArray.Count > 0 Then
                If GravaBancoDestino(SqlArray, sConnStringDestino) Then
                    Return True
                Else
                    Return False
                End If
            Else
                mensagem = "Carteiras sem registros para Transferência"
                Return True
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
            Return False
        End Try
    End Function

    Private Function gravaCarteirasXTributos(ByRef mensagem As String) As Boolean
        SqlArray.Clear()

        Try
            sql = "DELETE FROM CarteirasXTributos"
            SqlArray.Add(sql)

            sql = "SELECT  CarteirasXTributos.Carteira_Id AS Codigo, ComprasXProdutos.Descricao, CarteirasXTributos.Tributo_ID as Tributo" & vbCrLf &
                  " FROM CarteirasXTributos INNER JOIN" & vbCrLf &
                  " ComprasXProdutos ON CarteirasXTributos.Carteira_Id = ComprasXProdutos.Produto_Id" & vbCrLf &
                  " ORDER BY Codigo" & vbCrLf

            Dim ds As DataSet = Banco.ConsultaDataSet(sql, "CarteirasXTributos")

            If ds IsNot Nothing AndAlso ds.Tables IsNot Nothing AndAlso ds.Tables.Count > 0 AndAlso ds.Tables(0).Rows.Count > 0 Then
                For Each row As DataRow In ds.Tables(0).Rows
                    sql = "Insert into CarteirasXTributos (Carteira_Id, Tributo_ID)" & vbCrLf &
                          "values ('" & row("Codigo") & "', '" & row("Tributo") & "')"
                    SqlArray.Add(sql)
                Next
            End If

            If SqlArray.Count > 0 Then
                If GravaBancoDestino(SqlArray, sConnStringDestino) Then
                    Return True
                Else
                    Return False
                End If
            Else
                mensagem = "CarteirasXTributos sem registros para Transferência"
                Return True
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
            Return False
        End Try
    End Function

    Private Function gravaClientes(ByRef mensagem As String, ByRef dataini As String, ByRef datafin As String) As Boolean
        SqlArray.Clear()

        Try
            sql = "select n.Cliente_id, n.EndCliente_id, " & vbCrLf &
                  "  	case " & vbCrLf &
                  "  		when len(isnull(cDes.Cliente_id,'')) > 0 " & vbCrLf &
                  "  			then 1 " & vbCrLf &
                  "  			else 0 " & vbCrLf &
                  "  	end as temCli " & vbCrLf &
                  "from Panorama.dbo.NotasFiscais n " & vbCrLf &
                  "  	left join Gerencial.dbo.Clientes cDes " & vbCrLf &
                  "  			on cDes.Cliente_id   = n.Cliente_id " & vbCrLf &
                  "  			and cDes.Endereco_id = n.EndCliente_id " & vbCrLf &
                  "where n.Movimento between '" & dataini & "' and '" & datafin & "'" & vbCrLf &
                  "  and isnull(n.Pedido,0) > 0 " & vbCrLf &
                  "  and n.cliente_id not in('03189063000398')" & vbCrLf &
                  "group by n.Cliente_id, n.EndCliente_id,isnull(cDes.Cliente_id,'')"

            Dim ds As DataSet = Banco.ConsultaDataSet(sql, "Clientes")

            If ds IsNot Nothing AndAlso ds.Tables(0).Rows.Count > 0 Then
                For Each row As DataRow In ds.Tables(0).Rows
                    Dim cli As New [Lib].Negocio.Cliente(row("Cliente_id"), row("EndCliente_id"))
                    If row("temCli") = "0" Then
                        cli.IUD = "I"
                    Else
                        cli.IUD = "U"
                    End If

                    cli.SalvarSql(SqlArray)
                Next
            End If

            sql = "update clientesXempresas set nossaemissao = 0, notafiscaleletronica = 'N'"
            SqlArray.Add(sql)
            If SqlArray.Count > 0 Then
                If GravaBancoDestino(SqlArray, sConnStringDestino) Then
                    Return True
                Else
                    Return False
                End If
            Else
                mensagem = "Clientes sem registros para Transferência"
                Return True
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
            Return False
        End Try
    End Function


    Private Function ClientesFaltantes(ByRef mensagem As String) As Boolean
        SqlArray.Clear()

        Try
            sql = "insert into Gerencial.dbo.Clientes (Cliente_Id, Endereco_Id, Regiao, Categoria, Estado, Pais, Nome, Fantasia, Endereco, Numero, Complemento, Bairro, Cep, Cidade, " & vbCrLf &
                    "			Inscricao, Telefone, Fax, Email, Imagem, Reduzido, CodigoDoMunicipio, Situacao, Habilitacao, EmailNFE, OrgaoRegCategoria, UsuarioInclusao, UsuarioInclusaoData, " & vbCrLf &
                    "			UsuarioAlteracao, UsuarioAlteracaoData, OutrosTelefones, Rg, Site, EstadoCivil, Sexo, NascimentoConstituicao, NaturalidadeCidade, NaturalidadeEstado, " & vbCrLf &
                    "			ClienteDesde, Foto, Suframa, DesdobrarFornecedor, RNTRCTransportador, ClienteCorrespondencia, EndClienteCorrespondencia, MicroRegiao) " & vbCrLf &
                    "SELECT c.Cliente_Id, c.Endereco_Id, c.Regiao, c.Categoria, c.Estado, c.Pais, c.Nome, c.Fantasia, c.Endereco, c.Numero, c.Complemento, c.Bairro, c.Cep, c.Cidade, " & vbCrLf &
                    "		c.Inscricao, c.Telefone, c.Fax, c.Email, c.Imagem, c.Reduzido, c.CodigoDoMunicipio, c.Situacao, c.Habilitacao, c.EmailNFE, c.OrgaoRegCategoria, " & vbCrLf &
                    "		c.UsuarioInclusao, c.UsuarioInclusaoData, c.UsuarioAlteracao, c.UsuarioAlteracaoData, c.OutrosTelefones, c.Rg, c.Site, c.EstadoCivil, c.Sexo, c.NascimentoConstituicao, " & vbCrLf &
                    "		c.NaturalidadeCidade, c.NaturalidadeEstado, c.ClienteDesde, c.Foto, c.Suframa, c.DesdobrarFornecedor, c.RNTRCTransportador, " & vbCrLf &
                    "		c.ClienteCorrespondencia, c.EndClienteCorrespondencia, c.MicroRegiao " & vbCrLf &
                    "FROM Panorama.dbo.Clientes AS c " & vbCrLf &
                    "	left join gerencial.dbo.clientes cg " & vbCrLf &
                    "			on cg.Cliente_Id = c.Cliente_Id " & vbCrLf &
                    "			and cg.Endereco_Id = c.Endereco_Id " & vbCrLf &
                    "where len(isnull(cg.Cliente_Id,'')) = 0"

            SqlArray.Add(sql)

            sql = "insert into Gerencial.dbo.ClientesXTipos (Tipo_id, Cliente_id, Endereco_id) " & vbCrLf &
                    "select ct.Tipo_id, ct.Cliente_id, ct.Endereco_id " & vbCrLf &
                    "from Panorama.dbo.ClientesXTipos ct " & vbCrLf &
                    "	left join Gerencial.dbo.ClientesXTipos ctg " & vbCrLf &
                    "			on ctg.Tipo_id = ct.Tipo_Id " & vbCrLf &
                    "			and ctg.Cliente_Id = ct.Cliente_Id " & vbCrLf &
                    "			and ctg.Endereco_Id = ct.Endereco_Id " & vbCrLf &
                    "where len(isnull(ctg.Cliente_Id,'')) = 0"

            SqlArray.Add(sql)

            If SqlArray.Count > 0 Then
                If GravaBancoDestino(SqlArray, sConnStringDestino) Then
                    Return True
                Else
                    Return False
                End If
            Else
                mensagem = "Clientes Faltantes sem registros para Transferência"
                Return True
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
            Return False
        End Try
    End Function

    Private Function gravaProducao(ByRef mensagem As String) As Boolean
        SqlArray.Clear()

        Try
            sql = "DELETE FROM Gerencial.dbo.ProducaoXAnalises where Empresa_id not in ('03189063000398') and year(Movimento_Id) = '" & ddlAno.SelectedValue & "' and month(Movimento_Id) = " & ddlMes.SelectedValue
            SqlArray.Add(sql)

            sql = "DELETE FROM Gerencial.dbo.Producao where Empresa_id not in ('03189063000398') and year(Movimento_Id) = '" & ddlAno.SelectedValue & "' and month(Movimento_Id) = " & ddlMes.SelectedValue
            SqlArray.Add(sql)

            sql = "insert into  Gerencial.dbo.Producao (Empresa_Id, EndEmpresa_Id, Deposito_Id, EndDeposito_Id, Produto_Id, Operacao_Id, SubOperacao_Id, Movimento_Id, FisicoFiscal_Id, ProdutoDerivado_Id, Etapa," & vbCrLf &
                  "                                         Safra, Entradas, Saidas, UnidadeDeNegocio, Lote_Id, Classificacao_Id, Embalagem_Id, TipoDeEmbalagem_Id, CapacidadeEmbalagem_Id, Observacao," & vbCrLf &
                  "                                         UsuarioInclusao, UsuarioInclusaoData)" & vbCrLf &
                  "SELECT     Empresa_Id, EndEmpresa_Id, Deposito_Id, EndDeposito_Id, Produto_Id, Operacao_Id, SubOperacao_Id, Movimento_Id, FisicoFiscal_Id, ProdutoDerivado_Id, Etapa," & vbCrLf &
                  "           Safra, Entradas, Saidas, UnidadeDeNegocio, Lote_Id, Classificacao_Id, Embalagem_Id, TipoDeEmbalagem_Id, CapacidadeEmbalagem_Id, Observacao," & vbCrLf &
                  "           UsuarioInclusao, UsuarioInclusaoData" & vbCrLf &
                  "FROM  Panorama.dbo.Producao" & vbCrLf &
                  "where year(Movimento_Id) = '" & ddlAno.SelectedValue & "' and month(Movimento_Id) = " & ddlMes.SelectedValue
            SqlArray.Add(sql)


            sql = "insert into Gerencial.dbo.ProducaoXAnalises (Empresa_Id, EndEmpresa_Id, Deposito_Id, EndDeposito_Id, Produto_Id, Operacao_Id, SubOperacao_Id, Movimento_Id, FisicoFiscal_Id, ProdutoDerivado_Id," & vbCrLf &
                  "                                                 Analise_Id, Etapa, Quantidade, Indice, Lote_Id, Classificacao_Id, Embalagem_Id, TipoDeEmbalagem_Id, CapacidadeEmbalagem_Id)" & vbCrLf &
                  "SELECT     Empresa_Id, EndEmpresa_Id, Deposito_Id, EndDeposito_Id, Produto_Id, Operacao_Id, SubOperacao_Id, Movimento_Id, FisicoFiscal_Id, ProdutoDerivado_Id," & vbCrLf &
                  "           Analise_Id, Etapa, Quantidade, Indice, Lote_Id, Classificacao_Id, Embalagem_Id, TipoDeEmbalagem_Id, CapacidadeEmbalagem_Id" & vbCrLf &
                  "FROM  Panorama.dbo.ProducaoXAnalises" & vbCrLf &
                  "where year(Movimento_Id) = '" & ddlAno.SelectedValue & "' and month(Movimento_Id) = " & ddlMes.SelectedValue
            SqlArray.Add(sql)

            If SqlArray.Count > 0 Then
                If GravaBancoDestino(SqlArray, sConnStringDestino) Then
                    Return True
                Else
                    Return False
                End If
            Else
                mensagem = "Produção sem registros para Transferência"
                Return True
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
            Return False
        End Try
    End Function

    Private Function passo1(ByRef mensagem As String, ByRef dataini As String, ByRef datafin As String) As Boolean
        SqlArray.Clear()

        Try
            sql = "delete nd" & vbCrLf &
                  " from Panorama.dbo.NotaFiscalDevolucaoXNotaFiscal nd" & vbCrLf &
                  "	inner join Panorama.dbo.NotasFiscais n" & vbCrLf &
                  "			ON n.Empresa_Id      = nd.EmpresaDevolucao_Id " & vbCrLf &
                  "		   AND n.EndEmpresa_Id   = nd.EndEmpresaDevolucao_Id " & vbCrLf &
                  "		   AND n.Cliente_Id      = nd.ClienteDevolucao_Id " & vbCrLf &
                  "		   AND n.EndCliente_Id   = nd.EndClienteDevolucao_Id " & vbCrLf &
                  "		   AND n.EntradaSaida_Id = nd.EntradaSaidaDevolucao_Id " & vbCrLf &
                  "		   AND n.Serie_Id        = nd.SerieDevolucao_Id " & vbCrLf &
                  "		   AND n.Nota_Id         = nd.NotaDevolucao_Id " & vbCrLf &
                  " WHERE not n.Empresa_Id = '03189063000398'" & vbCrLf &
                  "   and YEAR(MOVIMENTO) = '" & ddlAno.SelectedValue & "' AND MONTH(MOVIMENTO) = '" & ddlMes.SelectedValue & "'"
            SqlArray.Add(sql)

            sql = "delete nd" & vbCrLf &
                  " from NotaFiscalDevolucaoXNotaFiscal nd" & vbCrLf &
                  "	inner join NotasFiscais n" & vbCrLf &
                  "			ON n.Empresa_Id      = nd.EmpresaDevolucao_Id " & vbCrLf &
                  "		   AND n.EndEmpresa_Id   = nd.EndEmpresaDevolucao_Id " & vbCrLf &
                  "		   AND n.Cliente_Id      = nd.ClienteDevolucao_Id " & vbCrLf &
                  "		   AND n.EndCliente_Id   = nd.EndClienteDevolucao_Id " & vbCrLf &
                  "		   AND n.EntradaSaida_Id = nd.EntradaSaidaDevolucao_Id " & vbCrLf &
                  "		   AND n.Serie_Id        = nd.SerieDevolucao_Id " & vbCrLf &
                  "		   AND n.Nota_Id         = nd.NotaDevolucao_Id " & vbCrLf &
                  " WHERE not n.Empresa_Id = '03189063000398'" & vbCrLf &
                  "   and YEAR(MOVIMENTO) = '" & ddlAno.SelectedValue & "' AND MONTH(MOVIMENTO) = '" & ddlMes.SelectedValue & "'"
            SqlArray.Add(sql)

            sql = "delete nd" & vbCrLf &
                  " from NotaFiscalReferencial nd" & vbCrLf &
                  "	inner join NotasFiscais n" & vbCrLf &
                  "			ON n.Empresa_Id      = nd.EmpresaReferencial_Id " & vbCrLf &
                  "		   AND n.EndEmpresa_Id   = nd.EndEmpresaReferencial_Id " & vbCrLf &
                  "		   AND n.Cliente_Id      = nd.ClienteReferencial_Id " & vbCrLf &
                  "		   AND n.EndCliente_Id   = nd.EndClienteReferencial_Id " & vbCrLf &
                  "		   AND n.EntradaSaida_Id = nd.EntradaSaidaReferencial_Id " & vbCrLf &
                  "		   AND n.Serie_Id        = nd.SerieReferencial_Id " & vbCrLf &
                  "		   AND n.Nota_Id         = nd.NotaReferencial_Id " & vbCrLf &
                  " WHERE not n.Empresa_Id = '03189063000398'" & vbCrLf &
                  "   and YEAR(MOVIMENTO) = '" & ddlAno.SelectedValue & "' AND MONTH(MOVIMENTO) = '" & ddlMes.SelectedValue & "'"
            SqlArray.Add(sql)

            sql = "delete nd" & vbCrLf &
                  " from NotaFiscalXLote nd" & vbCrLf &
                  "	inner join NotasFiscais n" & vbCrLf &
                  "			ON n.Empresa_Id      = nd.Empresa_Id " & vbCrLf &
                  "		   AND n.EndEmpresa_Id   = nd.EndEmpresa_Id " & vbCrLf &
                  "		   AND n.Cliente_Id      = nd.Cliente_Id " & vbCrLf &
                  "		   AND n.EndCliente_Id   = nd.EndCliente_Id " & vbCrLf &
                  "		   AND n.EntradaSaida_Id = nd.EntradaSaida_Id " & vbCrLf &
                  "		   AND n.Serie_Id        = nd.Serie_Id " & vbCrLf &
                  "		   AND n.Nota_Id         = nd.Nota_Id " & vbCrLf &
                  " WHERE not n.Empresa_Id = '03189063000398'" & vbCrLf &
                  "   and YEAR(MOVIMENTO) = '" & ddlAno.SelectedValue & "' AND MONTH(MOVIMENTO) = '" & ddlMes.SelectedValue & "'"
            SqlArray.Add(sql)

            sql = "delete nd" & vbCrLf &
                  " from NotasXEmbalagens nd" & vbCrLf &
                  "	inner join NotasFiscais n" & vbCrLf &
                  "			ON n.Empresa_Id      = nd.Empresa_Id " & vbCrLf &
                  "		   AND n.EndEmpresa_Id   = nd.EndEmpresa_Id " & vbCrLf &
                  "		   AND n.Cliente_Id      = nd.Cliente_Id " & vbCrLf &
                  "		   AND n.EndCliente_Id   = nd.EndCliente_Id " & vbCrLf &
                  "		   AND n.EntradaSaida_Id = nd.EntradaSaida_Id " & vbCrLf &
                  "		   AND n.Serie_Id        = nd.Serie_Id " & vbCrLf &
                  "		   AND n.Nota_Id         = nd.Nota_Id " & vbCrLf &
                  " WHERE not n.Empresa_Id = '03189063000398'" & vbCrLf &
                  "   and YEAR(MOVIMENTO) = '" & ddlAno.SelectedValue & "' AND MONTH(MOVIMENTO) = '" & ddlMes.SelectedValue & "'"
            SqlArray.Add(sql)

            sql = "delete nd" & vbCrLf &
                  " from NotasXNotas nd" & vbCrLf &
                  "	inner join NotasFiscais n" & vbCrLf &
                  "			ON n.Empresa_Id      = nd.Empresa_Id " & vbCrLf &
                  "		   AND n.EndEmpresa_Id   = nd.EndEmpresa_Id " & vbCrLf &
                  "		   AND n.Cliente_Id      = nd.Cliente_Id " & vbCrLf &
                  "		   AND n.EndCliente_Id   = nd.EndCliente_Id " & vbCrLf &
                  "		   AND n.EntradaSaida_Id = nd.EntradaSaida_Id " & vbCrLf &
                  "		   AND n.Serie_Id        = nd.Serie_Id " & vbCrLf &
                  "		   AND n.Nota_Id         = nd.Nota_Id " & vbCrLf &
                  " WHERE not n.Empresa_Id = '03189063000398'" & vbCrLf &
                  "   and YEAR(MOVIMENTO) = '" & ddlAno.SelectedValue & "' AND MONTH(MOVIMENTO) = '" & ddlMes.SelectedValue & "'"
            SqlArray.Add(sql)

            sql = "delete nd" & vbCrLf &
                  " from NotasFiscaisXPercursos nd" & vbCrLf &
                  "	inner join NotasFiscais n" & vbCrLf &
                  "			ON n.Empresa_Id      = nd.Empresa_Id " & vbCrLf &
                  "		   AND n.EndEmpresa_Id   = nd.EndEmpresa_Id " & vbCrLf &
                  "		   AND n.Cliente_Id      = nd.Cliente_Id " & vbCrLf &
                  "		   AND n.EndCliente_Id   = nd.EndCliente_Id " & vbCrLf &
                  "		   AND n.EntradaSaida_Id = nd.EntradaSaida_Id " & vbCrLf &
                  "		   AND n.Serie_Id        = nd.Serie_Id " & vbCrLf &
                  "		   AND n.Nota_Id         = nd.Nota_Id " & vbCrLf &
                  " WHERE not n.Empresa_Id = '03189063000398'" & vbCrLf &
                  "   and YEAR(MOVIMENTO) = '" & ddlAno.SelectedValue & "' AND MONTH(MOVIMENTO) = '" & ddlMes.SelectedValue & "'"
            SqlArray.Add(sql)

            sql = "delete nd" & vbCrLf &
                  " from NotasFiscaisXTransportadores nd" & vbCrLf &
                  "	inner join NotasFiscais n" & vbCrLf &
                  "			ON n.Empresa_Id      = nd.Empresa_Id " & vbCrLf &
                  "		   AND n.EndEmpresa_Id   = nd.EndEmpresa_Id " & vbCrLf &
                  "		   AND n.Cliente_Id      = nd.Cliente_Id " & vbCrLf &
                  "		   AND n.EndCliente_Id   = nd.EndCliente_Id " & vbCrLf &
                  "		   AND n.EntradaSaida_Id = nd.EntradaSaida_Id " & vbCrLf &
                  "		   AND n.Serie_Id        = nd.Serie_Id " & vbCrLf &
                  "		   AND n.Nota_Id         = nd.Nota_Id " & vbCrLf &
                  " WHERE not n.Empresa_Id = '03189063000398'" & vbCrLf &
                  "   and YEAR(MOVIMENTO) = '" & ddlAno.SelectedValue & "' AND MONTH(MOVIMENTO) = '" & ddlMes.SelectedValue & "'"
            SqlArray.Add(sql)

            sql = "delete nd" & vbCrLf &
                  " from NotaFiscalXImportacao nd" & vbCrLf &
                  "	inner join NotasFiscais n" & vbCrLf &
                  "			ON n.Empresa_Id      = nd.Empresa_Id " & vbCrLf &
                  "		   AND n.EndEmpresa_Id   = nd.EndEmpresa_Id " & vbCrLf &
                  "		   AND n.Cliente_Id      = nd.Cliente_Id " & vbCrLf &
                  "		   AND n.EndCliente_Id   = nd.EndCliente_Id " & vbCrLf &
                  "		   AND n.EntradaSaida_Id = nd.EntradaSaida_Id " & vbCrLf &
                  "		   AND n.Serie_Id        = nd.Serie_Id " & vbCrLf &
                  "		   AND n.Nota_Id         = nd.Nota_Id " & vbCrLf &
                  " WHERE not n.Empresa_Id = '03189063000398'" & vbCrLf &
                  "   and YEAR(MOVIMENTO) = '" & ddlAno.SelectedValue & "' AND MONTH(MOVIMENTO) = '" & ddlMes.SelectedValue & "'"
            SqlArray.Add(sql)

            sql = "delete nd" & vbCrLf &
                  " from NotaFiscalXExportacao nd" & vbCrLf &
                  "	inner join NotasFiscais n" & vbCrLf &
                  "			ON n.Empresa_Id      = nd.Empresa_Id " & vbCrLf &
                  "		   AND n.EndEmpresa_Id   = nd.EndEmpresa_Id " & vbCrLf &
                  "		   AND n.Cliente_Id      = nd.Cliente_Id " & vbCrLf &
                  "		   AND n.EndCliente_Id   = nd.EndCliente_Id " & vbCrLf &
                  "		   AND n.EntradaSaida_Id = nd.EntradaSaida_Id " & vbCrLf &
                  "		   AND n.Serie_Id        = nd.Serie_Id " & vbCrLf &
                  "		   AND n.Nota_Id         = nd.Nota_Id " & vbCrLf &
                  " WHERE not n.Empresa_Id = '03189063000398'" & vbCrLf &
                  "   and YEAR(MOVIMENTO) = '" & ddlAno.SelectedValue & "' AND MONTH(MOVIMENTO) = '" & ddlMes.SelectedValue & "'"
            SqlArray.Add(sql)

            sql = "delete nd" & vbCrLf &
                  " from NotaFiscalXRE nd" & vbCrLf &
                  "	inner join NotasFiscais n" & vbCrLf &
                  "			ON n.Empresa_Id      = nd.Empresa_Id " & vbCrLf &
                  "		   AND n.EndEmpresa_Id   = nd.EndEmpresa_Id " & vbCrLf &
                  "		   AND n.Cliente_Id      = nd.Cliente_Id " & vbCrLf &
                  "		   AND n.EndCliente_Id   = nd.EndCliente_Id " & vbCrLf &
                  "		   AND n.EntradaSaida_Id = nd.EntradaSaida_Id " & vbCrLf &
                  "		   AND n.Serie_Id        = nd.Serie_Id " & vbCrLf &
                  "		   AND n.Nota_Id         = nd.Nota_Id " & vbCrLf &
                  " WHERE not n.Empresa_Id = '03189063000398'" & vbCrLf &
                  "   and YEAR(MOVIMENTO) = '" & ddlAno.SelectedValue & "' AND MONTH(MOVIMENTO) = '" & ddlMes.SelectedValue & "'"
            SqlArray.Add(sql)

            sql = "delete nd" & vbCrLf &
                  " from NotaFiscalXTitulo nd" & vbCrLf &
                  "	inner join NotasFiscais n" & vbCrLf &
                  "			ON n.Empresa_Id      = nd.Empresa_Id " & vbCrLf &
                  "		   AND n.EndEmpresa_Id   = nd.EndEmpresa_Id " & vbCrLf &
                  "		   AND n.Cliente_Id      = nd.Cliente_Id " & vbCrLf &
                  "		   AND n.EndCliente_Id   = nd.EndCliente_Id " & vbCrLf &
                  "		   AND n.EntradaSaida_Id = nd.EntradaSaida_Id " & vbCrLf &
                  "		   AND n.Serie_Id        = nd.Serie_Id " & vbCrLf &
                  "		   AND n.Nota_Id         = nd.Nota_Id " & vbCrLf &
                  " WHERE not n.Empresa_Id = '03189063000398'" & vbCrLf &
                  "   and YEAR(MOVIMENTO) = '" & ddlAno.SelectedValue & "' AND MONTH(MOVIMENTO) = '" & ddlMes.SelectedValue & "'"
            SqlArray.Add(sql)

            sql = "delete nd" & vbCrLf &
                  " from NFERealizadas nd" & vbCrLf &
                  "	inner join NotasFiscais n" & vbCrLf &
                  "			ON n.Empresa_Id      = nd.Empresa_Id " & vbCrLf &
                  "		   AND n.EndEmpresa_Id   = nd.EndEmpresa_Id " & vbCrLf &
                  "		   AND n.Cliente_Id      = nd.Cliente_Id " & vbCrLf &
                  "		   AND n.EndCliente_Id   = nd.EndCliente_Id " & vbCrLf &
                  "		   AND n.EntradaSaida_Id = nd.EntradaSaida_Id " & vbCrLf &
                  "		   AND n.Serie_Id        = nd.Serie_Id " & vbCrLf &
                  "		   AND n.Nota_Id         = nd.Nota_Id " & vbCrLf &
                  " WHERE not n.Empresa_Id = '03189063000398'" & vbCrLf &
                  "   and YEAR(MOVIMENTO) = '" & ddlAno.SelectedValue & "' AND MONTH(MOVIMENTO) = '" & ddlMes.SelectedValue & "'"
            SqlArray.Add(sql)

            sql = "delete nd" & vbCrLf &
                  " from NotasFiscaisXEncargos nd" & vbCrLf &
                  "	inner join NotasFiscais n" & vbCrLf &
                  "			ON n.Empresa_Id      = nd.Empresa_Id " & vbCrLf &
                  "		   AND n.EndEmpresa_Id   = nd.EndEmpresa_Id " & vbCrLf &
                  "		   AND n.Cliente_Id      = nd.Cliente_Id " & vbCrLf &
                  "		   AND n.EndCliente_Id   = nd.EndCliente_Id " & vbCrLf &
                  "		   AND n.EntradaSaida_Id = nd.EntradaSaida_Id " & vbCrLf &
                  "		   AND n.Serie_Id        = nd.Serie_Id " & vbCrLf &
                  "		   AND n.Nota_Id         = nd.Nota_Id " & vbCrLf &
                  " WHERE not n.Empresa_Id = '03189063000398'" & vbCrLf &
                  "   and YEAR(MOVIMENTO) = '" & ddlAno.SelectedValue & "' AND MONTH(MOVIMENTO) = '" & ddlMes.SelectedValue & "'"
            SqlArray.Add(sql)

            sql = "delete nd" & vbCrLf &
                  " from NotasFiscaisXItensXRateio nd" & vbCrLf &
                  "	inner join NotasFiscais n" & vbCrLf &
                  "			ON n.Empresa_Id      = nd.Empresa_Id " & vbCrLf &
                  "		   AND n.EndEmpresa_Id   = nd.EndEmpresa_Id " & vbCrLf &
                  "		   AND n.Cliente_Id      = nd.Cliente_Id " & vbCrLf &
                  "		   AND n.EndCliente_Id   = nd.EndCliente_Id " & vbCrLf &
                  "		   AND n.EntradaSaida_Id = nd.EntradaSaida_Id " & vbCrLf &
                  "		   AND n.Serie_Id        = nd.Serie_Id " & vbCrLf &
                  "		   AND n.Nota_Id         = nd.Nota_Id " & vbCrLf &
                  " WHERE not n.Empresa_Id = '03189063000398'" & vbCrLf &
                  "   and YEAR(MOVIMENTO) = '" & ddlAno.SelectedValue & "' AND MONTH(MOVIMENTO) = '" & ddlMes.SelectedValue & "'"
            SqlArray.Add(sql)

            sql = "delete nd" & vbCrLf &
                  " from NotasFiscaisXitens nd" & vbCrLf &
                  "	inner join NotasFiscais n" & vbCrLf &
                  "			ON n.Empresa_Id      = nd.Empresa_Id " & vbCrLf &
                  "		   AND n.EndEmpresa_Id   = nd.EndEmpresa_Id " & vbCrLf &
                  "		   AND n.Cliente_Id      = nd.Cliente_Id " & vbCrLf &
                  "		   AND n.EndCliente_Id   = nd.EndCliente_Id " & vbCrLf &
                  "		   AND n.EntradaSaida_Id = nd.EntradaSaida_Id " & vbCrLf &
                  "		   AND n.Serie_Id        = nd.Serie_Id " & vbCrLf &
                  "		   AND n.Nota_Id         = nd.Nota_Id " & vbCrLf &
                  " WHERE not n.Empresa_Id = '03189063000398'" & vbCrLf &
                  "   and YEAR(MOVIMENTO) = '" & ddlAno.SelectedValue & "' AND MONTH(MOVIMENTO) = '" & ddlMes.SelectedValue & "'"
            SqlArray.Add(sql)

            sql = "delete from NotasFiscais" & vbCrLf &
                  " WHERE not Empresa_Id = '03189063000398'" & vbCrLf &
                  "   and YEAR(MOVIMENTO) = '" & ddlAno.SelectedValue & "' AND MONTH(MOVIMENTO) = '" & ddlMes.SelectedValue & "'"
            SqlArray.Add(sql)

            sql = "Delete from Razao " & vbCrLf &
                  "where empresa_id not in ('03189063000398') and year(Movimento_Id) = '" & ddlAno.SelectedValue & "' and month(Movimento_Id) = '" & ddlMes.SelectedValue & "' and Lote_id in(9,10,11,21)"
            SqlArray.Add(sql)

            sql = "Update Panorama.dbo.NotasFiscais set Conferencia = 0 Where Conferencia = 1"
            SqlArray.Add(sql)

            If SqlArray.Count > 0 Then
                If GravaBancoDestino(SqlArray, sConnStringDestino) Then
                    Return True
                Else
                    Return False
                End If
            Else
                mensagem = "Notas/Razão sem registros para Deleção"
                Return True
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
            Return False
        End Try
    End Function

    Private Sub Deletar()
        SqlArray.Clear()

        Try
            'Razão
            'sql = "delete from razao" & vbCrLf & _
            '      "where titulo in(74503,74504,74505,74506,74507,76760,76761)"
            'SqlArray.Add(sql)


            ''Financeiro
            'sql = "Delete from contasAreceber" & vbCrLf & _
            '      "where cliente = '45512332000193'"
            'SqlArray.Add(sql)


            'Notas Fiscais
            Dim listNotas As New [Lib].Negocio.ListNotasFiscais("05272759000147", "0", "2015-01-01", "2015-12-31", "45512332000193")

            For Each iNota As [Lib].Negocio.NotaFiscal In listNotas
                If iNota.Codigo = 9441 Or iNota.Codigo = 9454 Or iNota.Codigo = 9572 Or iNota.Codigo = 9714 Or iNota.Codigo = 9716 Or iNota.Codigo = 9717 Then
                    iNota.EnderecoCliente = 0
                    iNota.IUD = "D"
                    iNota.SalvarSql(SqlArray)
                End If
            Next


            ''Pedidos
            'sql = "select Empresa_Id, EndEmpresa_Id, Pedido_Id from pedidos" & vbCrLf & _
            '      "where cliente = '45512332000193'"
            'Dim dsPed As DataSet = Banco.ConsultaDataSet(sql, "Pedidos")

            'If dsPed IsNot Nothing AndAlso dsPed.Tables(0).Rows.Count > 0 Then
            '    For Each row As DataRow In dsPed.Tables(0).Rows
            '        Dim ped As New [Lib].Negocio.Pedido(row("Empresa_Id"), row("EndEmpresa_Id"), row("Pedido_Id"))
            '        ped.IUD = "D"
            '        ped.SalvarSql(SqlArray)
            '    Next
            'End If


            If SqlArray.Count > 0 Then
                If GravaBancoDestino(SqlArray, sConnStringDestino) Then
                    MsgBox(Me.Page, "registros deletados")
                Else
                    MsgBox(Me.Page, "erro")
                End If
            Else
                MsgBox(Me.Page, "Sem registros")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Private Sub GravarRegistros()
        Try
            ''Pedidos
            'sql = "select Empresa_Id, EndEmpresa_Id, Pedido_Id from pedidos" & vbCrLf & _
            '      "where cliente = '45512332000193'"
            'Dim dsPed As DataSet = Banco.ConsultaDataSet(sql, "Pedidos")

            'If dsPed IsNot Nothing AndAlso dsPed.Tables(0).Rows.Count > 0 Then
            '    For Each row As DataRow In dsPed.Tables(0).Rows
            '        Dim ped As New [Lib].Negocio.Pedido(row("Empresa_Id"), row("EndEmpresa_Id"), row("Pedido_Id"))
            '        ped.IUD = "I"
            '        ped.EnderecoCliente = 0
            '        ped.SalvarSql(SqlArray)
            '    Next
            'End If


            'Notas Fiscais
            Dim listNotas As New [Lib].Negocio.ListNotasFiscais("04854422000185", "0", "2010-02-28", "2010-02-28", "36820989920", "2")

            For Each iNota As [Lib].Negocio.NotaFiscal In listNotas
                If iNota.Codigo = 26466 Then
                    'Private objNotaFiscal As [Lib].Negocio.NotaFiscal
                    Dim objNotaFiscal = New [Lib].Negocio.NotaFiscal(iNota)

                    objNotaFiscal.IUD = "I"

                    objNotaFiscal.EnderecoCliente = 0
                    objNotaFiscal.EnderecoDestino = 0

                    If objNotaFiscal.Cliente Is Nothing Then
                        MsgBox(Me.Page, "cliente não encontrado")
                        Exit Sub
                    End If

                    If objNotaFiscal.Itens.Count = 0 Then
                        MsgBox(Me.Page, "Sem itens")
                        Exit Sub
                    Else
                        For Each item In objNotaFiscal.Itens
                            If item.Encargos.Count = 0 Then
                                MsgBox(Me.Page, "Sem encargos")
                                Exit Sub
                            End If
                        Next
                    End If

                    objNotaFiscal.Serie = "0"

                    objNotaFiscal.SalvarSql(SqlArray)
                End If
            Next

            If SqlArray.Count > 0 Then
                If GravaBancoDestino(SqlArray, sConnStringDestino) Then
                    MsgBox(Me.Page, "registros gravados")
                End If
            Else
                MsgBox(Me.Page, "Sem registros")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Private Sub transfNotas(ByRef mensagem As String, ByRef dataini As String, ByRef datafin As String)
        SqlArray.Clear()

        Try
            Dim listNotas As New [Lib].Negocio.ListNotasFiscais("44005444000276", 0, dataini, datafin, "", "", "", "", 1, 39)

            For Each iNota As [Lib].Negocio.NotaFiscal In listNotas
                iNota.IUD = "I"

                For Each tit In iNota.VencimentosNota
                    tit.IUD = "I"
                Next

                iNota.SalvarSql(SqlArray)
            Next

            If SqlArray.Count > 0 Then
                If GravaBancoDestino(SqlArray, sConnStringDestino) Then
                    MsgBox(Me.Page, "registros transferidos")
                Else
                    MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                End If
            Else
                MsgBox(Me.Page, "Notas sem registros para Transferência")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub


    Private Function gravaNotas(ByRef mensagem As String, ByRef dataini As String, ByRef datafin As String) As Boolean
        SqlArray.Clear()

        Dim strm As New StreamWriter(HttpContext.Current.Server.MapPath("~/Files/notas.txt"))

        Try
            sql = "SELECT Empresa_Id, EndEmpresa_Id " & vbCrLf &
                  " FROM clientesXempresas " & vbCrLf

            If Left(Session("ssEmpresa"), 8) = "05272759" Then
                sql &= "where left(Empresa_id,8) = '05272759'"
            Else
                sql &= "where left(Empresa_id,8) = '03189063'"
            End If

            Dim ds As DataSet = Banco.ConsultaDataSet(sql, "clientesXempresas")

            sql = "SELECT     Operacao_Id, SubOperacao_Id, OperacaoDestino_Id, SubOperacaoDestino_Id" & vbCrLf &
                  "FROM         DeParaOperacao"
            Dim dsDepara As DataSet = Banco.ConsultaDataSet(sql, "DeParaOperacao")

            If ds IsNot Nothing AndAlso ds.Tables IsNot Nothing AndAlso ds.Tables.Count > 0 AndAlso ds.Tables(0).Rows.Count > 0 Then
                For Each row As DataRow In ds.Tables(0).Rows
                    Dim listNotas As New [Lib].Negocio.ListNotasFiscais(row("Empresa_Id"), row("EndEmpresa_Id"), dataini, datafin, "", "", "", "")

                    For Each iNota As [Lib].Negocio.NotaFiscal In listNotas
                        strm.WriteLine(iNota.Codigo)

                        If iNota.Codigo = 40794 Then
                            Dim teste As String = ""
                        End If


                        Dim ope As String = String.Empty
                        Dim sop As String = String.Empty

                        If Not Left(Session("ssEmpresa"), 8) = "05272759" AndAlso iNota.CodigoCliente = "05841957000184" Then

                            For Each rOpe As DataRow In dsDepara.Tables(0).Rows
                                If rOpe("Operacao_Id") = iNota.CodigoOperacao AndAlso rOpe("SubOperacao_Id") = iNota.CodigoSubOperacao Then
                                    ope = rOpe("OperacaoDestino_Id")
                                    sop = rOpe("SubOperacaoDestino_Id")
                                End If
                            Next

                            If Not String.IsNullOrWhiteSpace(ope) Then
                                For x As Integer = 0 To iNota.Itens.Count - 1

                                    If Not iNota.CodigoOperacao = ope Then
                                        iNota.CodigoOperacao = ope
                                        iNota.CodigoSubOperacao = sop
                                    End If

                                    iNota.CodigoCliente = "03189063000398"

                                    If iNota.CodigoDeposito = "05841957000184" Then
                                        iNota.CodigoDeposito = "03189063000398"
                                    End If

                                    If iNota.CodigoDestino = "05841957000184" Then
                                        iNota.CodigoDestino = "03189063000398"
                                    End If

                                    iNota.Itens(x).CodigoOperacao = ope
                                    iNota.Itens(x).CodigoSubOperacao = sop
                                    iNota.Itens(x).CodigoOperacaoEstado = 0
                                    iNota.Itens(x).Encargos.Clear()

                                    Dim Parametros As New OperacaoXEstado
                                    Parametros.Empresa = Left(iNota.CodigoEmpresa, 8)
                                    Parametros.CodigoGrupoProduto = iNota.Itens(x).Produto.CodigoGrupo
                                    Parametros.CodigoProduto = iNota.Itens(x).CodigoProduto
                                    Parametros.CodigoOperacao = iNota.Itens(x).CodigoOperacao
                                    Parametros.CodigoSubOperacao = iNota.Itens(x).CodigoSubOperacao
                                    Parametros.EstadoOrigem = iNota.Empresa.CodigoEstado
                                    Parametros.EstadoDestino = iNota.Cliente.CodigoEstado
                                    Parametros.InicioVigencia = iNota.Movimento
                                    Dim OXE As New OperacaoXEstado(Parametros)
                                    iNota.Itens(x).CodigoOperacaoEstado = OXE.Codigo

                                    iNota.Itens(x).CarregandoEncargos = True
                                    iNota.Itens(x).Encargos = New [Lib].Negocio.ListNotaFiscalXItemXEncargo(iNota.Itens(x), New List(Of eEtapaEncago) From {eEtapaEncago.Normal})
                                    iNota.Itens(x).CarregandoEncargos = False

                                    If iNota.Itens(x).Encargos.Count = 0 Then
                                        mensagem = "Nota Fiscal " & iNota.Codigo & "-" & iNota.Serie & " - Produto " & iNota.Itens(x).CodigoProduto & "-" & iNota.Itens(x).Produto.Nome & ", não tem encargos cadastrados na Operação:" & iNota.Itens(x).CodigoOperacao & "-" & iNota.Itens(x).CodigoSubOperacao
                                        Return False
                                    End If

                                    Dim codigocfop As Integer = iNota.Itens.OrderByDescending(Function(s) s.ValorTotal).First.OperacaoEstado.CodigoFiscal
                                    iNota.CFOP = New [Lib].Negocio.CFOP(codigocfop)
                                Next
                            End If
                        End If

                        sql = " Insert Into NotasFiscais " & vbCrLf &
                                  " (Empresa_Id, EndEmpresa_Id, " & vbCrLf &
                                  "  Cliente_Id, EndCliente_Id, " & vbCrLf &
                                  "  EntradaSaida_Id, Serie_Id, Nota_Id, TipoDeDocumento, " & vbCrLf &
                                  "  Formulario, Pedido, Procuracao, " & vbCrLf &
                                  "  Operacao, SubOperacao, Finalidade, " & vbCrLf &
                                  "  Movimento, DataDaNota, NossaEmissao, " & vbCrLf &
                                  "  SerieNotadoProdutor, NumeroNotadoProdutor, " & vbCrLf &
                                  "  Deposito, EndDeposito, " & vbCrLf &
                                  "  Destino, EndDestino, " & vbCrLf &
                                  "  Transbordo, EndTransbordo, " & vbCrLf &
                                  "  Agenciador, EndAgenciador, " & vbCrLf &
                                  "  CIFFOB, Observacoes, Eletronica, " & vbCrLf &
                                  "  Autorizacao," & vbCrLf &
                                  "  UsuarioInclusao, UsuarioInclusaoData, " & vbCrLf

                        If iNota.UsuarioAlteracao.Length > 0 Then sql &= "UsuarioAlteracao, UsuarioAlteracaoData, " & vbCrLf

                        sql &= "  Situacao,ObservacoesDoProduto,ObservacoesDeEmbarque,ObservacoesControleInterno,EstadoDoCliente,LocalEmbarque,EndLocalEmbarque, " & vbCrLf &
                                  "  ContratoANTT, ProtocoloANTT, CartaoPgtoFrete, DataTermino, idPamcard, TipoDeDocumentoFrete, Favorecido, EndFavorecido, NFG, Conferencia, UsuarioConferencia, UsuarioConferenciaData,Troca) " & vbCrLf &
                                  "Values ('" & iNota.CodigoEmpresa & "'," & iNota.EnderecoEmpresa & ", " & vbCrLf &
                                                             "'" & iNota.CodigoCliente & "'," & iNota.EnderecoCliente & ", " & vbCrLf &
                                                             "'" & iNota.EntradaSaida.ToString.Substring(0, 1) & "','" & iNota.Serie & "'," & iNota.Codigo & "," & iNota.CodigoTipoDeDocumento & ", " & vbCrLf &
                                                             " " & iNota.Formulario & ", " & iNota.CodigoPedido.ToSqlNULL & ", " & iNota.CodigoProcuracao.ToSqlNULL & ", " & vbCrLf &
                                                             " " & iNota.CodigoOperacao & ", " & iNota.CodigoSubOperacao & ", " & iNota.CodigoFinalidade & ", " & vbCrLf &
                                                             "'" & iNota.Movimento.ToString("yyyy-MM-dd") & "','" & iNota.DataNota.ToString("yyyy-MM-dd") & "','" & IIf(iNota.NossaEmissao, "S", "N") & "' , " & vbCrLf &
                                                             "'" & iNota.SerieNotaProdutor & "', " & IIf(iNota.NotaProdutor = 0, "Null", iNota.NotaProdutor) & ", " & vbCrLf &
                                                             "'" & iNota.CodigoDeposito & "', " & iNota.EnderecoDeposito & ", " & vbCrLf &
                                                             "'" & iNota.CodigoDestino & "', " & iNota.EnderecoDestino & ", " & vbCrLf &
                                                             "'" & iNota.CodigoTransbordo & "', " & iNota.EnderecoTransbordo & ", " & vbCrLf &
                                                             "'" & iNota.CodigoAgenciador & "', " & iNota.EnderecoAgenciador & ", " & vbCrLf &
                                                             "'" & iNota.CIFFOB.ToString & "', '" & iNota.Observacoes & "', '" & IIf(iNota.Eletronica, "S", "N") & "', " & vbCrLf &
                                                             " " & iNota.CodigoAutorizacao & "," & vbCrLf &
                                                             "'" & iNota.UsuarioInclusao & "','" & iNota.DataInclusao.ToString("yyyy-MM-dd HH:mm:ss") & "'," & vbCrLf

                        If iNota.UsuarioAlteracao.Length > 0 Then sql &= "'" & iNota.UsuarioAlteracao & "','" & iNota.DataAlteracao.ToString("yyyy-MM-dd HH:mm:ss") & "'," & vbCrLf

                        sql &= "'" & iNota.CodigoSituacao & "','" & iNota.ObservacoesDoProduto & "','" & iNota.ObservacoesDeEmbarque & "','" & iNota.ObservacoesControleInterno & "'" & vbCrLf &
                                  ",'" & iNota.Cliente.CodigoEstado & "'" & vbCrLf &
                                  "," & IIf(iNota.CodigoLocalEmbarque.Length > 0, "'" & iNota.CodigoLocalEmbarque & "'", "NULL") & vbCrLf &
                                  "," & IIf(iNota.CodigoLocalEmbarque.Length > 0, iNota.EndLocalEmbarque, "NULL") & vbCrLf &
                                  ",'" & iNota.ContratoANTT & "','" & iNota.ProtocoloANTT & "','" & iNota.CartaoPgtoFrete & "','" & iNota.DataTermino.ToString("yyyy-MM-dd") & "','" & iNota.idPamcard & "', "

                        If iNota.TipoDeDocumentoFrete IsNot Nothing Then sql &= "'" & iNota.TipoDeDocumentoFrete & "'," Else sql &= "null,"
                        If Not String.IsNullOrWhiteSpace(iNota.CodigoFavorecido) Then sql &= "'" & iNota.CodigoFavorecido & "'," & iNota.EnderecoFavorecido & "," Else sql &= "null,null,"

                        sql &= IIf(iNota.NFG, 1, 0) & "," & vbCrLf &
                                  IIf(iNota.Conferencia, 1, IIf(iNota.Empresa.Empresa.ConferenciaNFE, 0, "NULL")) & "," & vbCrLf &
                                  IIf(iNota.Conferencia AndAlso Not String.IsNullOrWhiteSpace(iNota.UsuarioConferencia), iNota.UsuarioConferencia & "," & iNota.UsuarioConferenciaData.ToString("yyyy-MM-dd HH:mm:ss"), "NULL,NULL") & vbCrLf &
                                  "," & IIf(iNota.Troca, 1, 0) & ")"
                        SqlArray.Add(sql)

                        iNota.IUD = "I"

                        For j As Integer = 0 To iNota.Itens.Count - 1
                            iNota.Itens(j).IUD = "I"
                            iNota.Itens(j).SalvarSql(SqlArray, iNota.Itens(j).Sequencia)
                        Next
                    Next
                Next
            End If

            strm.Close()

            If SqlArray.Count > 0 Then
                If GravaBancoDestino(SqlArray, sConnStringDestino) Then
                    Return True
                Else
                    Return False
                End If
            Else
                mensagem = "Notas sem registros para Transferência"
                Return True
            End If
        Catch ex As Exception
            strm.Close()
            MsgBox(Me.Page, ex.Message)
            Return False
        End Try
    End Function

    Private Function gravaPedidos(ByRef mensagem As String, ByRef dataini As String, ByRef datafin As String) As Boolean
        SqlArray.Clear()

        'Dim strm As New StreamWriter(HttpContext.Current.Server.MapPath("~/Files/pedidos.txt"))

        Try
            sql = "select n.Empresa_id, n.EndEmpresa_id, n.Pedido, dp.OperacaoDestino_Id, dp.SubOperacaoDestino_Id, " & vbCrLf &
                  "	case" & vbCrLf &
                  "		when isnull(p.Pedido_id,0) > 0" & vbCrLf &
                  "			then 1" & vbCrLf &
                  "			else 0" & vbCrLf &
                  "	end as temPed" & vbCrLf &
                  "from Panorama.dbo.NotasFiscais n" & vbCrLf &
                  "	inner join Pedidos ped" & vbCrLf &
                  "			on ped.Empresa_id     = n.Empresa_id" & vbCrLf &
                  "			and ped.EndEmpresa_id = n.EndEmpresa_id" & vbCrLf &
                  "			and ped.Pedido_id     = n.Pedido" & vbCrLf &
                  "	left join DeParaOperacao dp" & vbCrLf &
                  "			on dp.operacao_id     = ped.operacao" & vbCrLf &
                  "			and dp.suboperacao_id = ped.suboperacao" & vbCrLf &
                  "	left join Gerencial.dbo.Pedidos p" & vbCrLf &
                  "			on p.Empresa_id     = n.Empresa_id" & vbCrLf &
                  "			and p.EndEmpresa_id = n.EndEmpresa_id" & vbCrLf &
                  "			and p.Pedido_id     = n.Pedido" & vbCrLf &
                  "where n.Movimento between '" & dataini & "' and '" & datafin & "'" & vbCrLf &
                  "and isnull(n.Pedido,0) > 0" & vbCrLf &
                  "group by n.Empresa_id, n.EndEmpresa_id, n.Pedido, dp.OperacaoDestino_Id, dp.SubOperacaoDestino_Id, isnull(p.Pedido_id,0)"
            Dim dsPed As DataSet = Banco.ConsultaDataSet(sql, "Pedidos")

            Try
                If dsPed IsNot Nothing AndAlso dsPed.Tables(0).Rows.Count > 0 Then
                    For Each row As DataRow In dsPed.Tables(0).Rows
                        'strm.WriteLine(CStr(row("Pedido")))

                        Dim ped As New [Lib].Negocio.Pedido(row("Empresa_Id"), row("EndEmpresa_Id"), row("Pedido"))
                        Dim ope As String = String.Empty
                        Dim sop As String = String.Empty

                        For Each itemPed In ped.Itens
                            If itemPed.Encargos.Count = 0 Then
                                'strm.Close()
                                mensagem = "Pedido " & ped.Codigo & " - Produto " & itemPed.Produto.Nome & ", não tem encargos cadastrados na Operação:" & ped.CodigoOperacao & "-" & ped.CodigoSubOperacao
                                Return False
                            End If
                        Next

                        If Not Left(Session("ssEmpresa"), 8) = "05272759" AndAlso ped.CodigoCliente = "05841957000184" Then
                            If Not IsDBNull(row("OperacaoDestino_Id")) Then
                                ped.CodigoCliente = "03189063000398"
                                ped.CodigoPraca = "03189063000398"
                                ope = row("OperacaoDestino_Id")
                                sop = row("SubOperacaoDestino_Id")

                                For Each dep In ped.Depositos
                                    If dep.Codigo = "05841957000184" Then
                                        dep.Codigo = "03189063000398"
                                    End If
                                Next
                            End If

                            If Not String.IsNullOrWhiteSpace(ope) Then
                                ped.CodigoOperacao = ope
                                ped.CodigoSubOperacao = sop

                                For Each itemPed In ped.Itens
                                    itemPed.Encargos.Clear()
                                    itemPed.Encargos = Nothing
                                    itemPed.Encargos.CriaListar()

                                    If itemPed.Encargos.Count = 0 Then
                                        'strm.Close()
                                        mensagem = "O Produto " & itemPed.CodigoProduto & "-" & itemPed.Produto.Nome & ", não tem encargos cadastrados na Operação:" & ped.CodigoOperacao & "-" & ped.CodigoSubOperacao & " - Pedido " & ped.Codigo
                                        Return False
                                    End If
                                Next
                            End If
                        End If

                        sql = "delete from pedidoXtabeladecomissao" & vbCrLf &
                              "where empresa_id    = '" & ped.CodigoEmpresa & "'" & vbCrLf &
                              "  and endempresa_id = " & ped.EnderecoEmpresa & vbCrLf &
                              "  and pedido_id     = " & ped.Codigo
                        SqlArray.Add(sql)

                        sql = "delete from comissoes" & vbCrLf &
                              "where empresa_id    = '" & ped.CodigoEmpresa & "'" & vbCrLf &
                              "  and endempresa_id = " & ped.EnderecoEmpresa & vbCrLf &
                              "  and pedido_id     = " & ped.Codigo
                        SqlArray.Add(sql)

                        If row("temPed") = "0" Then
                            ped.IUD = "I"
                        Else
                            ped.IUD = "U"
                        End If

                        ped.SalvarSql(SqlArray)
                    Next
                End If

                'strm.Close()

                If SqlArray.Count > 0 Then
                    If GravaBancoDestino(SqlArray, sConnStringDestino) Then
                        Return True
                    Else
                        Return False
                    End If
                Else
                    mensagem = "Pedidos sem registros para Transferência"
                    Return True
                End If
            Catch ex As Exception
                'strm.Close()
                MsgBox(Me.Page, ex.Message)
                Return False
            End Try
        Catch ex As Exception
            'strm.Close()
            MsgBox(Me.Page, ex.Message)
            Return False
        End Try
    End Function

    Private Function gravaTitulos(ByRef mensagem As String, ByRef dataini As String, ByRef datafin As String) As Boolean
        SqlArray.Clear()

        Try

            sql = "Select Registro_Id from contasApagar " & vbCrLf &
                  "WHERE YEAR(baixa) = '" & ddlAno.SelectedValue & "' AND MONTH(baixa) = '" & ddlMes.SelectedValue & "'" & vbCrLf &
                  "  AND not Empresa in('03189063000398')"
            Dim dsP As DataSet = Banco.ConsultaDataSet(sql, "contasApagar")

            If dsP IsNot Nothing AndAlso dsP.Tables IsNot Nothing AndAlso dsP.Tables.Count > 0 AndAlso dsP.Tables(0).Rows.Count > 0 Then
                For Each row As DataRow In dsP.Tables(0).Rows
                    sql = "Delete from Gerencial.dbo.NotaFiscalXTitulo " & vbCrLf &
                          "WHERE Titulo_Id = " & row("Registro_Id")
                    SqlArray.Add(sql)

                    sql = "Delete from Gerencial.dbo.contasApagar " & vbCrLf &
                          "WHERE Registro_Id = " & row("Registro_Id")
                    SqlArray.Add(sql)
                Next
            End If

            sql = "Select Registro_Id from contasAreceber " & vbCrLf &
                  "WHERE YEAR(baixa) = '" & ddlAno.SelectedValue & "' AND MONTH(baixa) = '" & ddlMes.SelectedValue & "'" & vbCrLf &
                  "  AND not Empresa in('03189063000398')"
            Dim dsR As DataSet = Banco.ConsultaDataSet(sql, "contasAreceber")

            If dsR IsNot Nothing AndAlso dsR.Tables IsNot Nothing AndAlso dsR.Tables.Count > 0 AndAlso dsR.Tables(0).Rows.Count > 0 Then
                For Each row As DataRow In dsR.Tables(0).Rows
                    sql = "Delete from Gerencial.dbo.NotaFiscalXTitulo " & vbCrLf &
                          "WHERE Titulo_Id = " & row("Registro_Id")
                    SqlArray.Add(sql)

                    sql = "Delete from Gerencial.dbo.contasAreceber " & vbCrLf &
                          "WHERE Registro_Id = " & row("Registro_Id")
                    SqlArray.Add(sql)
                Next
            End If

            sql = "Select Registro_Id from MovimentacoesFinanceiras " & vbCrLf &
                  "WHERE YEAR(baixa) = '" & ddlAno.SelectedValue & "' AND MONTH(baixa) = '" & ddlMes.SelectedValue & "'" & vbCrLf &
                  "  AND not Empresa in('03189063000398')"
            Dim dsM As DataSet = Banco.ConsultaDataSet(sql, "MovimentacoesFinanceiras")

            If dsM IsNot Nothing AndAlso dsM.Tables IsNot Nothing AndAlso dsM.Tables.Count > 0 AndAlso dsM.Tables(0).Rows.Count > 0 Then
                For Each row As DataRow In dsM.Tables(0).Rows
                    sql = "Delete from Gerencial.dbo.MovimentacoesFinanceiras " & vbCrLf &
                          "WHERE Registro_Id = " & row("Registro_Id")
                    SqlArray.Add(sql)
                Next
            End If

            sql = "delete nt" & vbCrLf &
                  "FROM  Gerencial.dbo.NotaFiscalXTitulo nt" & vbCrLf &
                  "	inner join Gerencial.dbo.ContasAPagar cp" & vbCrLf &
                  "			on nt.Empresa_Id     = cp.Empresa" & vbCrLf &
                  "			and nt.EndEmpresa_Id = cp.EndEmpresa" & vbCrLf &
                  "			and nt.Titulo_Id     = cp.Registro_id" & vbCrLf &
                  "WHERE YEAR(cp.baixa) = '" & ddlAno.SelectedValue & "' AND MONTH(cp.baixa) = '" & ddlMes.SelectedValue & "'" & vbCrLf &
                  "  AND not cp.Empresa in('03189063000398')"
            SqlArray.Add(sql)

            sql = "Delete from Gerencial.dbo.contasApagar " & vbCrLf &
                  "WHERE YEAR(baixa) = '" & ddlAno.SelectedValue & "' AND MONTH(baixa) = '" & ddlMes.SelectedValue & "'" & vbCrLf &
                  "  AND not Empresa in('03189063000398')"
            SqlArray.Add(sql)

            sql = "delete nt" & vbCrLf &
                  "FROM  Gerencial.dbo.NotaFiscalXTitulo nt" & vbCrLf &
                  "	inner join Gerencial.dbo.contasAreceber cp" & vbCrLf &
                  "			on nt.Empresa_Id     = cp.Empresa" & vbCrLf &
                  "			and nt.EndEmpresa_Id = cp.EndEmpresa" & vbCrLf &
                  "			and nt.Titulo_Id     = cp.Registro_id" & vbCrLf &
                  "WHERE YEAR(cp.baixa) = '" & ddlAno.SelectedValue & "' AND MONTH(cp.baixa) = '" & ddlMes.SelectedValue & "'" & vbCrLf &
                  "  AND not cp.Empresa in('03189063000398')"
            SqlArray.Add(sql)

            sql = "Delete from Gerencial.dbo.contasAreceber " & vbCrLf &
                  "WHERE YEAR(baixa) = '" & ddlAno.SelectedValue & "' AND MONTH(baixa) = '" & ddlMes.SelectedValue & "'" & vbCrLf &
                  "  AND not Empresa in('03189063000398')"
            SqlArray.Add(sql)

            sql = "Delete from Gerencial.dbo.MovimentacoesFinanceiras " & vbCrLf &
                  "WHERE YEAR(baixa) = '" & ddlAno.SelectedValue & "' AND MONTH(baixa) = '" & ddlMes.SelectedValue & "'" & vbCrLf &
                  "  AND not Empresa in('03189063000398')"
            SqlArray.Add(sql)

            sql = "Delete from Gerencial.dbo.TitulosXDesdobrarFornecedor "
            SqlArray.Add(sql)

            sql = "insert into Gerencial.dbo.ContasAPagar (Registro_Id, Sequencia_Id, Provisao, Carteira, Tributo, Indexador, Moeda, TipoPagto, Situacao, Lote, Movimento, Vencimento, Prorrogacao, DataMoeda, Baixa, " & vbCrLf &
                  "                      UnidadeDeNegocio, Empresa, EndEmpresa, Cliente, EndCliente, BancoCliente, AgenciaCliente, DigitoAgenciaCliente, ContaCliente, DigitoContaCliente, " & vbCrLf &
                  "                      ContaContabilCliente, EmpresaPagadora, EndEmpresaPagadora, BancoPagador, AgenciaPagadora, DigitoAgenciaPagadora, ContaPagadora, DigitoContaPagadora, " & vbCrLf &
                  "                      ContaContabilPagadora, Cheque, Slips, Recibo, Aviso, ReciboDeposito, EmpresaPedido, EndEmpresaPedido, Pedido, PedidoFixacao, Procuracao, ValorDoDocumento, " & vbCrLf &
                  "                      Descontos, Deducoes, Juros, Acrescimos, ValorLiquido, MoedaValorDoDocumento, MoedaDescontos, MoedaDeducoes, MoedaJuros, MoedaAcrescimos, " & vbCrLf &
                  "                      MoedaValorLiquido, Historico, CodigoDeBarras, CodigoDigitado, Destinatario, EndDestinatario, NomeDoDestinatario, Destinacao, solicitacao, UsuarioInclusao, " & vbCrLf &
                  "                      UsuarioInclusaoData, UsuarioAlteracao, UsuarioAlteracaoData, UsuarioCancelamento, UsuarioCancelamentoData, UsuarioLiberacao, UsuarioLiberacaoData, " & vbCrLf &
                  "                      UsuarioBaixa, UsuarioBaixaData, Grupado, RegistroMestre, Observacoes, SituacaoBancaria, TaxaAdto, VencimentoAdto, Adiantamento, NumeroDoCheque, " & vbCrLf &
                  "                      UsuarioLiberacaoBloqueio, UsuarioLiberacaoBloqueioDate, UsuarioLiberacaoPedido, UsuarioLiberacaoPedidoDate, UsuarioLiberacaoCheque, " & vbCrLf &
                  "                      UsuarioLiberacaoChequeDate, CarteiraAdto, CarteiraDoTitulo, ContratoDeFinanciamento, DataEnvio, Ocorrencia, motivo, DigitoNossoNumero, NossoNumero, " & vbCrLf &
                  "                      Timbrado, CodigoDeBarra, CodigoDeBarraDigitado, ModalidadeDePagamento, CreditoNaConta, NumeroDaRemessa, SituacaoRemessaBancaria, TipoContaCliente, " & vbCrLf &
                  "                      ContratoBancario, CodigoDeBarraPreImpresso, TituloOrigem, ValorRecompra)" & vbCrLf &
                  "SELECT     Registro_Id, Sequencia_Id, Provisao, Carteira, Tributo, Indexador, Moeda, TipoPagto, Situacao, Lote, Movimento, Vencimento, Prorrogacao, DataMoeda, Baixa, " & vbCrLf &
                  "                      UnidadeDeNegocio, Empresa, EndEmpresa, Cliente, EndCliente, BancoCliente, AgenciaCliente, DigitoAgenciaCliente, ContaCliente, DigitoContaCliente, " & vbCrLf &
                  "                      ContaContabilCliente, EmpresaPagadora, EndEmpresaPagadora, BancoPagador, AgenciaPagadora, DigitoAgenciaPagadora, ContaPagadora, DigitoContaPagadora, " & vbCrLf &
                  "                      ContaContabilPagadora, Cheque, Slips, Recibo, Aviso, ReciboDeposito, EmpresaPedido, EndEmpresaPedido, Pedido, PedidoFixacao, Procuracao, ValorDoDocumento, " & vbCrLf &
                  "                      Descontos, Deducoes, Juros, Acrescimos, ValorLiquido, MoedaValorDoDocumento, MoedaDescontos, MoedaDeducoes, MoedaJuros, MoedaAcrescimos, " & vbCrLf &
                  "                      MoedaValorLiquido, Historico, CodigoDeBarras, CodigoDigitado, Destinatario, EndDestinatario, NomeDoDestinatario, Destinacao, solicitacao, UsuarioInclusao, " & vbCrLf &
                  "                      UsuarioInclusaoData, UsuarioAlteracao, UsuarioAlteracaoData, UsuarioCancelamento, UsuarioCancelamentoData, UsuarioLiberacao, UsuarioLiberacaoData, " & vbCrLf &
                  "                      UsuarioBaixa, UsuarioBaixaData, Grupado, RegistroMestre, Observacoes, SituacaoBancaria, TaxaAdto, VencimentoAdto, Adiantamento, NumeroDoCheque, " & vbCrLf &
                  "                      UsuarioLiberacaoBloqueio, UsuarioLiberacaoBloqueioDate, UsuarioLiberacaoPedido, UsuarioLiberacaoPedidoDate, UsuarioLiberacaoCheque, " & vbCrLf &
                  "                      UsuarioLiberacaoChequeDate, CarteiraAdto, CarteiraDoTitulo, ContratoDeFinanciamento, DataEnvio, Ocorrencia, motivo, DigitoNossoNumero, NossoNumero, " & vbCrLf &
                  "                      Timbrado, CodigoDeBarra, CodigoDeBarraDigitado, ModalidadeDePagamento, CreditoNaConta, NumeroDaRemessa, SituacaoRemessaBancaria, TipoContaCliente, " & vbCrLf &
                  "                      ContratoBancario, CodigoDeBarraPreImpresso, TituloOrigem, ValorRecompra" & vbCrLf &
                  "FROM         Panorama.dbo.ContasAPagar" & vbCrLf &
                  "WHERE YEAR(baixa) = '" & ddlAno.SelectedValue & "' AND MONTH(baixa) = '" & ddlMes.SelectedValue & "'"
            SqlArray.Add(sql)

            sql = "insert into Gerencial.dbo.NotaFiscalXTitulo (Empresa_Id, EndEmpresa_Id, Cliente_Id, EndCliente_Id, EntradaSaida_Id, Serie_Id, Nota_Id, Titulo_Id)" & vbCrLf &
                  "SELECT nt.Empresa_Id, nt.EndEmpresa_Id, " & vbCrLf &
                  "case" & vbCrLf &
                  "	when left(nt.Empresa_Id,8) not in('05272759') and nt.Cliente_Id = '05841957000184' and isnull(dp.OperacaoDestino_Id,0) > 0" & vbCrLf &
                  "		then '03189063000398'" & vbCrLf &
                  "		else nt.Cliente_Id" & vbCrLf &
                  "	end as Cliente_Id," & vbCrLf &
                  "    nt.EndCliente_Id, nt.EntradaSaida_Id, nt.Serie_Id, nt.Nota_Id, nt.Titulo_Id" & vbCrLf &
                  "FROM  Panorama.dbo.NotaFiscalXTitulo nt" & vbCrLf &
                  "    inner join Panorama.dbo.NotasFiscais n" & vbCrLf &
                  "            ON n.Empresa_Id       = nt.Empresa_Id" & vbCrLf &
                  "            AND n.EndEmpresa_Id   = nt.EndEmpresa_Id " & vbCrLf &
                  "            AND n.Cliente_Id      = nt.Cliente_Id " & vbCrLf &
                  "            AND n.EndCliente_Id   = nt.EndCliente_Id " & vbCrLf &
                  "            AND n.EntradaSaida_Id = nt.EntradaSaida_Id " & vbCrLf &
                  "            AND n.Serie_Id        = nt.Serie_Id " & vbCrLf &
                  "            AND n.Nota_Id         = nt.Nota_Id " & vbCrLf &
                  "	inner join Panorama.dbo.ContasAPagar cp" & vbCrLf &
                  "			on nt.Empresa_Id     = cp.Empresa" & vbCrLf &
                  "			and nt.EndEmpresa_Id = cp.EndEmpresa" & vbCrLf &
                  "			and nt.Titulo_Id     = cp.Registro_id" & vbCrLf &
                  "    left join Panorama.dbo.DeParaOperacao dp" & vbCrLf &
                  "            ON dp.Operacao_Id     = n.Operacao " & vbCrLf &
                  "            AND dp.SubOperacao_Id = n.SubOperacao " & vbCrLf &
                  "WHERE YEAR(cp.baixa) = '" & ddlAno.SelectedValue & "' AND MONTH(cp.baixa) = '" & ddlMes.SelectedValue & "'"
            SqlArray.Add(sql)

            sql = "insert into Gerencial.dbo.contasAreceber (Registro_Id, Sequencia_Id, Provisao, Carteira, Tributo, Indexador, Moeda, TipoPagto, Situacao, Lote, Movimento, Vencimento, Prorrogacao, DataMoeda, Baixa, " & vbCrLf &
                  "                      UnidadeDeNegocio, Empresa, EndEmpresa, Cliente, EndCliente, BancoCliente, AgenciaCliente, DigitoAgenciaCliente, ContaCliente, DigitoContaCliente, " & vbCrLf &
                  "                      ContaContabilCliente, EmpresaPagadora, EndEmpresaPagadora, BancoPagador, AgenciaPagadora, DigitoAgenciaPagadora, ContaPagadora, DigitoContaPagadora, " & vbCrLf &
                  "                      ContaContabilPagadora, Cheque, Slips, Recibo, Aviso, ReciboDeposito, EmpresaPedido, EndEmpresaPedido, Pedido, PedidoFixacao, Procuracao, ValorDoDocumento, " & vbCrLf &
                  "                      Descontos, Deducoes, Juros, Acrescimos, ValorLiquido, MoedaValorDoDocumento, MoedaDescontos, MoedaDeducoes, MoedaJuros, MoedaAcrescimos, " & vbCrLf &
                  "                      MoedaValorLiquido, Historico, CodigoDeBarras, CodigoDigitado, Destinatario, EndDestinatario, NomeDoDestinatario, Destinacao, solicitacao, UsuarioInclusao, " & vbCrLf &
                  "                      UsuarioInclusaoData, UsuarioAlteracao, UsuarioAlteracaoData, UsuarioCancelamento, UsuarioCancelamentoData, UsuarioLiberacao, UsuarioLiberacaoData, " & vbCrLf &
                  "                      UsuarioBaixa, UsuarioBaixaData, Grupado, RegistroMestre, Observacoes, SituacaoBancaria, TaxaAdto, VencimentoAdto, Adiantamento, NumeroDoCheque, " & vbCrLf &
                  "                      UsuarioLiberacaoBloqueio, UsuarioLiberacaoBloqueioDate, UsuarioLiberacaoPedido, UsuarioLiberacaoPedidoDate, UsuarioLiberacaoCheque, " & vbCrLf &
                  "                      UsuarioLiberacaoChequeDate, CarteiraAdto, CarteiraDoTitulo, ContratoDeFinanciamento, DataEnvio, Ocorrencia, motivo, DigitoNossoNumero, NossoNumero, " & vbCrLf &
                  "                      Timbrado, CodigoDeBarra, CodigoDeBarraDigitado, ModalidadeDePagamento, CreditoNaConta, NumeroDaRemessa, SituacaoRemessaBancaria, TipoContaCliente, " & vbCrLf &
                  "                      ContratoBancario, CodigoDeBarraPreImpresso, TituloOrigem, ValorRecompra)" & vbCrLf &
                  "SELECT     Registro_Id, Sequencia_Id, Provisao, Carteira, Tributo, Indexador, Moeda, TipoPagto, Situacao, Lote, Movimento, Vencimento, Prorrogacao, DataMoeda, Baixa, " & vbCrLf &
                  "                      UnidadeDeNegocio, Empresa, EndEmpresa, Cliente, EndCliente, BancoCliente, AgenciaCliente, DigitoAgenciaCliente, ContaCliente, DigitoContaCliente, " & vbCrLf &
                  "                      ContaContabilCliente, EmpresaPagadora, EndEmpresaPagadora, BancoPagador, AgenciaPagadora, DigitoAgenciaPagadora, ContaPagadora, DigitoContaPagadora, " & vbCrLf &
                  "                      ContaContabilPagadora, Cheque, Slips, Recibo, Aviso, ReciboDeposito, EmpresaPedido, EndEmpresaPedido, Pedido, PedidoFixacao, Procuracao, ValorDoDocumento, " & vbCrLf &
                  "                      Descontos, Deducoes, Juros, Acrescimos, ValorLiquido, MoedaValorDoDocumento, MoedaDescontos, MoedaDeducoes, MoedaJuros, MoedaAcrescimos, " & vbCrLf &
                  "                      MoedaValorLiquido, Historico, CodigoDeBarras, CodigoDigitado, Destinatario, EndDestinatario, NomeDoDestinatario, Destinacao, solicitacao, UsuarioInclusao, " & vbCrLf &
                  "                      UsuarioInclusaoData, UsuarioAlteracao, UsuarioAlteracaoData, UsuarioCancelamento, UsuarioCancelamentoData, UsuarioLiberacao, UsuarioLiberacaoData, " & vbCrLf &
                  "                      UsuarioBaixa, UsuarioBaixaData, Grupado, RegistroMestre, Observacoes, SituacaoBancaria, TaxaAdto, VencimentoAdto, Adiantamento, NumeroDoCheque, " & vbCrLf &
                  "                      UsuarioLiberacaoBloqueio, UsuarioLiberacaoBloqueioDate, UsuarioLiberacaoPedido, UsuarioLiberacaoPedidoDate, UsuarioLiberacaoCheque, " & vbCrLf &
                  "                      UsuarioLiberacaoChequeDate, CarteiraAdto, CarteiraDoTitulo, ContratoDeFinanciamento, DataEnvio, Ocorrencia, motivo, DigitoNossoNumero, NossoNumero, " & vbCrLf &
                  "                      Timbrado, CodigoDeBarra, CodigoDeBarraDigitado, ModalidadeDePagamento, CreditoNaConta, NumeroDaRemessa, SituacaoRemessaBancaria, TipoContaCliente, " & vbCrLf &
                  "                      ContratoBancario, CodigoDeBarraPreImpresso, TituloOrigem, ValorRecompra" & vbCrLf &
                  "FROM         Panorama.dbo.contasAreceber" & vbCrLf &
                  "WHERE YEAR(baixa) = '" & ddlAno.SelectedValue & "' AND MONTH(baixa) = '" & ddlMes.SelectedValue & "'"
            SqlArray.Add(sql)

            sql = "insert into Gerencial.dbo.NotaFiscalXTitulo (Empresa_Id, EndEmpresa_Id, Cliente_Id, EndCliente_Id, EntradaSaida_Id, Serie_Id, Nota_Id, Titulo_Id)" & vbCrLf &
                  "SELECT nt.Empresa_Id, nt.EndEmpresa_Id, " & vbCrLf &
                  "case" & vbCrLf &
                  "	when left(nt.Empresa_Id,8) not in('05272759') and nt.Cliente_Id = '05841957000184' and isnull(dp.OperacaoDestino_Id,0) > 0" & vbCrLf &
                  "		then '03189063000398'" & vbCrLf &
                  "		else nt.Cliente_Id" & vbCrLf &
                  "	end as Cliente_Id," & vbCrLf &
                  "    nt.EndCliente_Id, nt.EntradaSaida_Id, nt.Serie_Id, nt.Nota_Id, nt.Titulo_Id" & vbCrLf &
                  "FROM  Panorama.dbo.NotaFiscalXTitulo nt" & vbCrLf &
                  "    inner join Panorama.dbo.NotasFiscais n" & vbCrLf &
                  "            ON n.Empresa_Id       = nt.Empresa_Id" & vbCrLf &
                  "            AND n.EndEmpresa_Id   = nt.EndEmpresa_Id " & vbCrLf &
                  "            AND n.Cliente_Id      = nt.Cliente_Id " & vbCrLf &
                  "            AND n.EndCliente_Id   = nt.EndCliente_Id " & vbCrLf &
                  "            AND n.EntradaSaida_Id = nt.EntradaSaida_Id " & vbCrLf &
                  "            AND n.Serie_Id        = nt.Serie_Id " & vbCrLf &
                  "            AND n.Nota_Id         = nt.Nota_Id " & vbCrLf &
                  "	inner join Panorama.dbo.contasAreceber cp" & vbCrLf &
                  "			on nt.Empresa_Id     = cp.Empresa" & vbCrLf &
                  "			and nt.EndEmpresa_Id = cp.EndEmpresa" & vbCrLf &
                  "			and nt.Titulo_Id     = cp.Registro_id" & vbCrLf &
                  "    left join Panorama.dbo.DeParaOperacao dp" & vbCrLf &
                  "            ON dp.Operacao_Id     = n.Operacao " & vbCrLf &
                  "            AND dp.SubOperacao_Id = n.SubOperacao " & vbCrLf &
                  "WHERE YEAR(cp.baixa) = '" & ddlAno.SelectedValue & "' AND MONTH(cp.baixa) = '" & ddlMes.SelectedValue & "'"
            SqlArray.Add(sql)

            sql = "insert into Gerencial.dbo.MovimentacoesFinanceiras (Registro_Id, Sequencia_Id, Provisao, Carteira, Tributo, Indexador, Moeda, TipoPagto, Situacao, Lote, Movimento, Vencimento, Prorrogacao, DataMoeda, Baixa, " & vbCrLf &
                  "                      UnidadeDeNegocio, Empresa, EndEmpresa, Cliente, EndCliente, BancoCliente, AgenciaCliente, DigitoAgenciaCliente, ContaCliente, DigitoContaCliente, " & vbCrLf &
                  "                      ContaContabilCliente, EmpresaPagadora, EndEmpresaPagadora, BancoPagador, AgenciaPagadora, DigitoAgenciaPagadora, ContaPagadora, DigitoContaPagadora, " & vbCrLf &
                  "                      ContaContabilPagadora, Cheque, Slips, Recibo, Aviso, ReciboDeposito, EmpresaPedido, EndEmpresaPedido, Pedido, PedidoFixacao, Procuracao, ValorDoDocumento, " & vbCrLf &
                  "                      Descontos, Deducoes, Juros, Acrescimos, ValorLiquido, MoedaValorDoDocumento, MoedaDescontos, MoedaDeducoes, MoedaJuros, MoedaAcrescimos, " & vbCrLf &
                  "                      MoedaValorLiquido, Historico, CodigoDeBarras, CodigoDigitado, Destinatario, EndDestinatario, NomeDoDestinatario, Destinacao, solicitacao, UsuarioInclusao, " & vbCrLf &
                  "                      UsuarioInclusaoData, UsuarioAlteracao, UsuarioAlteracaoData, UsuarioCancelamento, UsuarioCancelamentoData, UsuarioLiberacao, UsuarioLiberacaoData, " & vbCrLf &
                  "                      UsuarioBaixa, UsuarioBaixaData, Grupado, RegistroMestre, Observacoes, SituacaoBancaria, TaxaAdto, VencimentoAdto, Adiantamento, NumeroDoCheque, " & vbCrLf &
                  "                      UsuarioLiberacaoBloqueio, UsuarioLiberacaoBloqueioDate, UsuarioLiberacaoPedido, UsuarioLiberacaoPedidoDate, UsuarioLiberacaoCheque, " & vbCrLf &
                  "                      UsuarioLiberacaoChequeDate, CarteiraAdto, CarteiraDoTitulo, ContratoDeFinanciamento, DataEnvio, Ocorrencia, motivo, DigitoNossoNumero, NossoNumero, " & vbCrLf &
                  "                      Timbrado, CodigoDeBarra, CodigoDeBarraDigitado, ModalidadeDePagamento, CreditoNaConta, NumeroDaRemessa, SituacaoRemessaBancaria, TipoContaCliente, " & vbCrLf &
                  "                      ContratoBancario, CodigoDeBarraPreImpresso, TituloOrigem, ValorRecompra)" & vbCrLf &
                  "SELECT     Registro_Id, Sequencia_Id, Provisao, Carteira, Tributo, Indexador, Moeda, TipoPagto, Situacao, Lote, Movimento, Vencimento, Prorrogacao, DataMoeda, Baixa, " & vbCrLf &
                  "                      UnidadeDeNegocio, Empresa, EndEmpresa, Cliente, EndCliente, BancoCliente, AgenciaCliente, DigitoAgenciaCliente, ContaCliente, DigitoContaCliente, " & vbCrLf &
                  "                      ContaContabilCliente, EmpresaPagadora, EndEmpresaPagadora, BancoPagador, AgenciaPagadora, DigitoAgenciaPagadora, ContaPagadora, DigitoContaPagadora, " & vbCrLf &
                  "                      ContaContabilPagadora, Cheque, Slips, Recibo, Aviso, ReciboDeposito, EmpresaPedido, EndEmpresaPedido, Pedido, PedidoFixacao, Procuracao, ValorDoDocumento, " & vbCrLf &
                  "                      Descontos, Deducoes, Juros, Acrescimos, ValorLiquido, MoedaValorDoDocumento, MoedaDescontos, MoedaDeducoes, MoedaJuros, MoedaAcrescimos, " & vbCrLf &
                  "                      MoedaValorLiquido, Historico, CodigoDeBarras, CodigoDigitado, Destinatario, EndDestinatario, NomeDoDestinatario, Destinacao, solicitacao, UsuarioInclusao, " & vbCrLf &
                  "                      UsuarioInclusaoData, UsuarioAlteracao, UsuarioAlteracaoData, UsuarioCancelamento, UsuarioCancelamentoData, UsuarioLiberacao, UsuarioLiberacaoData, " & vbCrLf &
                  "                      UsuarioBaixa, UsuarioBaixaData, Grupado, RegistroMestre, Observacoes, SituacaoBancaria, TaxaAdto, VencimentoAdto, Adiantamento, NumeroDoCheque, " & vbCrLf &
                  "                      UsuarioLiberacaoBloqueio, UsuarioLiberacaoBloqueioDate, UsuarioLiberacaoPedido, UsuarioLiberacaoPedidoDate, UsuarioLiberacaoCheque, " & vbCrLf &
                  "                      UsuarioLiberacaoChequeDate, CarteiraAdto, CarteiraDoTitulo, ContratoDeFinanciamento, DataEnvio, Ocorrencia, motivo, DigitoNossoNumero, NossoNumero, " & vbCrLf &
                  "                      Timbrado, CodigoDeBarra, CodigoDeBarraDigitado, ModalidadeDePagamento, CreditoNaConta, NumeroDaRemessa, SituacaoRemessaBancaria, TipoContaCliente, " & vbCrLf &
                  "                      ContratoBancario, CodigoDeBarraPreImpresso, TituloOrigem, ValorRecompra" & vbCrLf &
                  "FROM         Panorama.dbo.MovimentacoesFinanceiras" & vbCrLf &
                  "WHERE YEAR(baixa) = '" & ddlAno.SelectedValue & "' AND MONTH(baixa) = '" & ddlMes.SelectedValue & "'"
            SqlArray.Add(sql)

            sql = "insert into Gerencial.dbo.TitulosXDesdobrarFornecedor (Registro_Id, Cliente, EndCliente, Pedido, Carteira)" & vbCrLf &
                  "SELECT Registro_Id, Cliente, EndCliente, Pedido, Carteira " & vbCrLf &
                  "FROM  Panorama.dbo.TitulosXDesdobrarFornecedor"
            SqlArray.Add(sql)

            If SqlArray.Count > 0 Then
                If GravaBancoDestino(SqlArray, sConnStringDestino) Then
                    Return True
                Else
                    Return False
                End If
            Else
                mensagem = "Financeiro sem registros para Transferência"
                Return True
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
            Return False
        End Try
    End Function

    Private Function gravaRazao(ByRef mensagem As String) As Boolean
        SqlArray.Clear()

        Try
            sql = "Delete Gerencial.dbo.Razao " & vbCrLf &
                  "where empresa_id not in ('03189063000398') and year(Movimento_Id) = '" & ddlAno.SelectedValue & "' and month(Movimento_Id) = '" & ddlMes.SelectedValue & "' and Not Lote_id in(9,10,11,21,123)"
            SqlArray.Add(sql)

            sql = "INSERT INTO Gerencial.dbo.Razao (Empresa_Id, EndEmpresa_Id, Conta_Id, Cliente_Id, EndCliente_Id, Movimento_Id, Lote_Id, Sequencia_Id, UnidadeDeNegocio, Titulo, Pedido, PedidoFixacao, Produto, " & vbCrLf &
                  "                      Custo, Indexador, DataMoeda, DebitoOficial, CreditoOficial, DebitoMoeda, CreditoMoeda, Historico, PrevistoRealizado, Cliente_Nf, EndCliente_Nf, EntradaSaida_Nf, " & vbCrLf &
                  "                      Serie_Nf, Numero_Nf, ChequeEntregue, PagamentoAutorizado, DataDaBaixa, Conciliacao, Deposito, EndDeposito, Rateado, Processo, UsuarioInclusao, " & vbCrLf &
                  "                      UsuarioInclusaoData, DebitoQuantidade, CreditoQuantidade, Situacao, Produto_NF, Cfop_NF, Sequencia_NF, Encargo_NF, AtivoImobilizado, " & vbCrLf &
                  "                      AgrupadorDeLancamento)" & vbCrLf &
                  "SELECT     Empresa_Id, EndEmpresa_Id, Conta_Id, Cliente_Id, EndCliente_Id, Movimento_Id, Lote_Id, Sequencia_Id, UnidadeDeNegocio, Titulo, Pedido, PedidoFixacao, Produto, " & vbCrLf &
                  "                      Custo, Indexador, DataMoeda, DebitoOficial, CreditoOficial, DebitoMoeda, CreditoMoeda, Historico, PrevistoRealizado, Cliente_Nf, EndCliente_Nf, EntradaSaida_Nf, " & vbCrLf &
                  "                      Serie_Nf, Numero_Nf, ChequeEntregue, PagamentoAutorizado, DataDaBaixa, Conciliacao, Deposito, EndDeposito, Rateado, Processo, UsuarioInclusao, " & vbCrLf &
                  "                      UsuarioInclusaoData, DebitoQuantidade, CreditoQuantidade, Situacao, Produto_NF, Cfop_NF, Sequencia_NF, Encargo_NF, AtivoImobilizado, " & vbCrLf &
                  "                      AgrupadorDeLancamento" & vbCrLf &
                  "FROM         Panorama.dbo.Razao" & vbCrLf &
                  "where empresa_id not in ('03189063000398') and year(Movimento_Id) = '" & ddlAno.SelectedValue & "' and month(Movimento_Id) = '" & ddlMes.SelectedValue & "' and Not Lote_id in(9,10,11,21,123) "
            SqlArray.Add(sql)

            If SqlArray.Count > 0 Then
                If GravaBancoDestino(SqlArray, sConnStringDestino) Then
                    Return True
                Else
                    Return False
                End If
            Else
                mensagem = "Razão sem registros para Transferência"
                Return True
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
            Return False
        End Try
    End Function

    Private Function RemoveTitulo(ByRef mensagem As String) As Boolean
        SqlArray.Clear()

        Try
            sql = "SELECT ContasAPagar.EmpresaPedido, ContasAPagar.EndEmpresaPedido, ContasAPagar.Registro_Id " & vbCrLf &
                  "FROM  Gerencial.dbo.ContasAPagar " & vbCrLf &
                  "		INNER JOIN Gerencial.dbo.Pedidos " & vbCrLf &
                  "				ON ContasAPagar.EmpresaPedido    = Pedidos.Empresa_Id " & vbCrLf &
                  "			   AND ContasAPagar.EndEmpresaPedido = Pedidos.EndEmpresa_Id " & vbCrLf &
                  "			   AND ContasAPagar.Pedido           = Pedidos.Pedido_Id " & vbCrLf &
                  "		INNER JOIN Gerencial.dbo.DeParaOperacao " & vbCrLf &
                  "				ON Pedidos.Operacao    = DeParaOperacao.OperacaoDestino_Id " & vbCrLf &
                  "			   AND Pedidos.SubOperacao = DeParaOperacao.SubOperacaoDestino_Id " & vbCrLf &
                  "		INNER JOIN Gerencial.dbo.SubOperacoes " & vbCrLf &
                  "				ON DeParaOperacao.OperacaoDestino_Id    = SubOperacoes.Operacao_Id " & vbCrLf &
                  "			   AND DeParaOperacao.SubOperacaoDestino_Id = SubOperacoes.SubOperacoes_Id " & vbCrLf &
                  "WHERE ContasAPagar.Provisao = 1 AND ContasAPagar.Situacao = 1 AND SubOperacoes.Classe = '" & eClassesOperacoes.TRANSFERENCIAS.ToString & "' " & vbCrLf &
                  "Group by ContasAPagar.EmpresaPedido, ContasAPagar.EndEmpresaPedido, ContasAPagar.Registro_Id "

            Dim dsRemoveTitulosAPagar As DataSet = Banco.ConsultaDataSet(sql, "RomveTitulosAPagar")

            For Each dr As DataRow In dsRemoveTitulosAPagar.Tables(0).Rows
                sql = "delete from razao " & vbCrLf &
                      "where Empresa_id    = '" & dr("EmpresaPedido") & "'" & vbCrLf &
                      "  and EndEmpresa_id = " & dr("EndEmpresaPedido") & "" & vbCrLf &
                      "  and Titulo        = " & dr("Registro_Id") & vbCrLf

                SqlArray.Add(sql)

                sql = "delete from ContasAPagar " & vbCrLf &
                      "where EmpresaPedido    = '" & dr("EmpresaPedido") & "'" & vbCrLf &
                      "  and EndEmpresaPedido = " & dr("EndEmpresaPedido") & "" & vbCrLf &
                      "  and Registro_Id      = " & dr("Registro_Id") & vbCrLf

                SqlArray.Add(sql)

                sql = "delete from NotaFiscalXTitulo " & vbCrLf &
                      "where Empresa_Id    = '" & dr("EmpresaPedido") & "'" & vbCrLf &
                      "  and EndEmpresa_Id = " & dr("EndEmpresaPedido") & "" & vbCrLf &
                      "  and Titulo_Id     = " & dr("Registro_Id") & vbCrLf

                SqlArray.Add(sql)
            Next

            sql = "SELECT ContasAReceber.EmpresaPedido, ContasAReceber.EndEmpresaPedido, ContasAReceber.Registro_Id " & vbCrLf &
                  "FROM  Gerencial.dbo.ContasAReceber " & vbCrLf &
                  "		INNER JOIN Gerencial.dbo.Pedidos " & vbCrLf &
                  "				ON ContasAReceber.EmpresaPedido    = Pedidos.Empresa_Id " & vbCrLf &
                  "			   AND ContasAReceber.EndEmpresaPedido = Pedidos.EndEmpresa_Id " & vbCrLf &
                  "			   AND ContasAReceber.Pedido           = Pedidos.Pedido_Id " & vbCrLf &
                  "		INNER JOIN Gerencial.dbo.DeParaOperacao " & vbCrLf &
                  "				ON Pedidos.Operacao    = DeParaOperacao.OperacaoDestino_Id " & vbCrLf &
                  "			   AND Pedidos.SubOperacao = DeParaOperacao.SubOperacaoDestino_Id " & vbCrLf &
                  "		INNER JOIN Gerencial.dbo.SubOperacoes " & vbCrLf &
                  "				ON DeParaOperacao.OperacaoDestino_Id    = SubOperacoes.Operacao_Id " & vbCrLf &
                  "			   AND DeParaOperacao.SubOperacaoDestino_Id = SubOperacoes.SubOperacoes_Id " & vbCrLf &
                  "WHERE ContasAReceber.Provisao = 1 AND ContasAReceber.Situacao = 1 AND SubOperacoes.Classe = '" & eClassesOperacoes.TRANSFERENCIAS.ToString & "' " & vbCrLf &
                  "Group by ContasAReceber.EmpresaPedido, ContasAReceber.EndEmpresaPedido, ContasAReceber.Registro_Id"

            Dim dsRemoveTitulosAReceber As DataSet = Banco.ConsultaDataSet(sql, "RomveTitulosAReceber")

            For Each dr As DataRow In dsRemoveTitulosAReceber.Tables(0).Rows
                sql = "delete from razao " & vbCrLf &
                      "where Empresa_id    = '" & dr("EmpresaPedido") & "'" & vbCrLf &
                      "  and EndEmpresa_id = " & dr("EndEmpresaPedido") & "" & vbCrLf &
                      "  and Titulo        = " & dr("Registro_Id") & vbCrLf

                SqlArray.Add(sql)

                sql = "delete from ContasAReceber " & vbCrLf &
                      "where EmpresaPedido    = '" & dr("EmpresaPedido") & "'" & vbCrLf &
                      "  and EndEmpresaPedido = " & dr("EndEmpresaPedido") & "" & vbCrLf &
                      "  and Registro_Id      = " & dr("Registro_Id") & vbCrLf

                SqlArray.Add(sql)

                sql = "delete from NotaFiscalXTitulo " & vbCrLf &
                      "where Empresa_Id    = '" & dr("EmpresaPedido") & "'" & vbCrLf &
                      "  and EndEmpresa_Id = " & dr("EndEmpresaPedido") & "" & vbCrLf &
                      "  and Titulo_Id     = " & dr("Registro_Id") & vbCrLf

                SqlArray.Add(sql)
            Next

            If SqlArray.Count > 0 Then
                If GravaBancoDestino(SqlArray, sConnStringDestino) Then
                    Return True
                Else
                    Return False
                End If
            Else
                mensagem = "Sem registros"
                Return True
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
            Return False
        End Try
    End Function

    Private Sub Limpar()
        chkOperacoes.Enabled = True
        chkOpeXEncargos.Enabled = True
        chkCarteiras.Enabled = True
        chkEncargos.Enabled = True
        chkCartXEnc.Enabled = True
        chkPlanoDeContas.Enabled = True
        chkProduto.Enabled = True
        chkClientes.Enabled = True
        chkClientesFaltantes.Enabled = True
        chkPedidos.Enabled = True
        chkNotas.Enabled = True
        chkProducao.Enabled = True
        chkTitulos.Enabled = True
        chkRazao.Enabled = True
        chkTitTransf.Enabled = True

        chkOperacoes.Checked = True
        chkOpeXEncargos.Checked = False
        chkCarteiras.Checked = False
        chkEncargos.Checked = False
        chkCartXEnc.Checked = False
        chkPlanoDeContas.Checked = False
        chkProduto.Checked = False
        chkClientes.Checked = False
        chkClientesFaltantes.Checked = False
        chkPedidos.Checked = False
        chkNotas.Checked = False
        chkProducao.Checked = False
        chkTitulos.Checked = False
        chkRazao.Checked = False
        chkTitTransf.Checked = False
    End Sub

    Protected Sub lnkLimpar_Click(sender As Object, e As EventArgs) Handles lnkLimpar.Click
        Try
            Limpar()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Protected Sub lnkNovo_Click(sender As Object, e As EventArgs) Handles lnkNovo.Click
        SqlArray.Clear()

        Dim dataini As String = ddlAno.SelectedValue & "-" & ddlMes.SelectedValue & "-01"
        Dim datafin As String = ddlAno.SelectedValue & "-" & ddlMes.SelectedValue & "-" & Day(CDate(dataini).AddMonths(1).AddDays(-1))

        Dim msgAlerta As String = String.Empty

        If chkProduto.Checked Then
            If gravaProduto(msgAlerta) Then
                chkProduto.Checked = False
                chkProduto.Enabled = False
            Else
                If Not String.IsNullOrWhiteSpace(msgAlerta) Then MsgBox(Me.Page, msgAlerta)
                Exit Sub
            End If
        End If

        If chkEncargos.Checked Then
            If gravaEncargos(msgAlerta) Then
                chkEncargos.Checked = False
                chkEncargos.Enabled = False
            Else
                If Not String.IsNullOrWhiteSpace(msgAlerta) Then MsgBox(Me.Page, msgAlerta)
                Exit Sub
            End If
        End If

        If chkOperacoes.Checked Then
            If gravaOperacoes(msgAlerta) Then
                chkOperacoes.Checked = False
                chkOperacoes.Enabled = False
            Else
                If Not String.IsNullOrWhiteSpace(msgAlerta) Then MsgBox(Me.Page, msgAlerta)
                Exit Sub
            End If
        End If

        If chkOpeXEncargos.Checked Then
            If gravaOpeXEncargos(msgAlerta) Then
                chkOpeXEncargos.Checked = False
                chkOpeXEncargos.Enabled = False
            Else
                If Not String.IsNullOrWhiteSpace(msgAlerta) Then MsgBox(Me.Page, msgAlerta)
                Exit Sub
            End If
        End If

        If chkCarteiras.Checked Then
            If gravaCarteiras(msgAlerta) Then
                chkCarteiras.Checked = False
                chkCarteiras.Enabled = False
            Else
                If Not String.IsNullOrWhiteSpace(msgAlerta) Then MsgBox(Me.Page, msgAlerta)
                Exit Sub
            End If
        End If

        If chkCartXEnc.Checked Then
            If gravaCarteirasXTributos(msgAlerta) Then
                chkCartXEnc.Checked = False
                chkCartXEnc.Enabled = False
            Else
                If Not String.IsNullOrWhiteSpace(msgAlerta) Then MsgBox(Me.Page, msgAlerta)
                Exit Sub
            End If
        End If

        If chkClientes.Checked Then
            If gravaClientes(msgAlerta, dataini, datafin) Then
                chkClientes.Checked = False
                chkClientes.Enabled = False
            Else
                If Not String.IsNullOrWhiteSpace(msgAlerta) Then MsgBox(Me.Page, msgAlerta)
                Exit Sub
            End If
        End If

        If chkClientesFaltantes.Checked Then
            If ClientesFaltantes(msgAlerta) Then
                chkClientes.Checked = False
                chkClientes.Enabled = False
            Else
                If Not String.IsNullOrWhiteSpace(msgAlerta) Then MsgBox(Me.Page, msgAlerta)
                Exit Sub
            End If


        End If

        If chkPlanoDeContas.Checked Then
            If gravaPlanoDeContas(msgAlerta) Then
                chkPlanoDeContas.Checked = False
                chkPlanoDeContas.Enabled = False
            Else
                If Not String.IsNullOrWhiteSpace(msgAlerta) Then MsgBox(Me.Page, msgAlerta)
                Exit Sub
            End If
        End If

        If chkProducao.Checked Then
            If gravaProducao(msgAlerta) Then
                chkProducao.Checked = False
                chkProducao.Enabled = False
            Else
                If Not String.IsNullOrWhiteSpace(msgAlerta) Then MsgBox(Me.Page, msgAlerta)
                Exit Sub
            End If
        End If

        If chkPedidos.Checked Then
            If gravaPedidos(msgAlerta, dataini, datafin) Then
                chkPedidos.Checked = False
                chkPedidos.Enabled = False
            Else
                If Not String.IsNullOrWhiteSpace(msgAlerta) Then MsgBox(Me.Page, msgAlerta)
                Exit Sub
            End If
        End If

        If chkPasso1.Checked Then
            If passo1(msgAlerta, dataini, datafin) Then
                chkPasso1.Checked = False
                chkPasso1.Enabled = False
            Else
                If Not String.IsNullOrWhiteSpace(msgAlerta) Then MsgBox(Me.Page, msgAlerta)
                Exit Sub
            End If
        End If

        If chkNotas.Checked Then
            If gravaNotas(msgAlerta, dataini, datafin) Then
                chkNotas.Checked = False
                chkNotas.Enabled = False
            Else
                If Not String.IsNullOrWhiteSpace(msgAlerta) Then MsgBox(Me.Page, msgAlerta)
                Exit Sub
            End If
        End If

        If chkTNotas.Checked Then
            transfNotas(msgAlerta, dataini, datafin)
        End If

        If chkTitulos.Checked Then
            If gravaTitulos(msgAlerta, dataini, datafin) Then
                chkTitulos.Checked = False
                chkTitulos.Enabled = False
            Else
                If Not String.IsNullOrWhiteSpace(msgAlerta) Then MsgBox(Me.Page, msgAlerta)
                Exit Sub
            End If
        End If

        If chkRazao.Checked Then
            If gravaRazao(msgAlerta) Then
                chkRazao.Checked = False
                chkRazao.Enabled = False
            Else
                If Not String.IsNullOrWhiteSpace(msgAlerta) Then MsgBox(Me.Page, msgAlerta)
                Exit Sub
            End If
        End If

        If chkTitTransf.Checked Then
            If RemoveTitulo(msgAlerta) Then
                chkTitTransf.Checked = False
                chkTitTransf.Enabled = False
            Else
                If Not String.IsNullOrWhiteSpace(msgAlerta) Then MsgBox(Me.Page, msgAlerta)
                Exit Sub
            End If
        End If

        If chkLerPlanilha.Checked Then
            Ler_Excel()
        End If

        If Not String.IsNullOrWhiteSpace(msgAlerta) Then
            MsgBox(Me.Page, msgAlerta)
        Else
            MsgBox(Me.Page, "Registros transferidos com Sucesso")
        End If
    End Sub

    Private Sub BaixarArquivo()
        Dim SqlTrans As OleDb.OleDbTransaction
        Dim Sqlcommand As New OleDb.OleDbCommand
        Dim SQlConnectionDestino As New OleDb.OleDbConnection
        Dim adp As New OleDb.OleDbDataAdapter
        Dim ds As New DataSet

        Dim sConnStringDestino As String = "Provider=SqlOleDb;" &
                       "Data Source=SRVMGA;" &
                       "Server=LOCALHOST;" &
                       "Database=Documentos;" &
                       "User Id=sa;" &
                       "Password=Pwd_in$ol@mga;" &
                       "Trusted_Connection=false;" &
                       "Integrated Security=true;"


        Try
            If (SQlConnectionDestino.State = ConnectionState.Closed) Then
                SQlConnectionDestino = New OleDb.OleDbConnection(sConnStringDestino)
                SQlConnectionDestino.Open()
            End If
            SqlTrans = SQlConnectionDestino.BeginTransaction(IsolationLevel.ReadCommitted)
            Sqlcommand.Connection = SQlConnectionDestino
            Sqlcommand.Transaction = SqlTrans

            Sqlcommand.CommandText = "SELECT Arquivo.PathName(), GET_FILESTREAM_TRANSACTION_CONTEXT(), Descricao FROM Arquivo WHERE Arquivo_Id = 'ff6dde92-b7f5-e511-8f3d-005056c00008'"
            Sqlcommand.CommandTimeout = 5000

            adp.SelectCommand = Sqlcommand
            adp.Fill(ds, "Arquivo")

            If ds IsNot Nothing AndAlso ds.Tables(0).Rows.Count > 0 Then

                Dim filePath As String = ds.Tables(0).Rows(0).Item(0).ToString()
                Dim objContext As Byte() = ds.Tables(0).Rows(0).Item(1)
                Dim fName As String = ds.Tables(0).Rows(0).Item(2).ToString()

                Dim sfs As New SqlFileStream(filePath, objContext, System.IO.FileAccess.Read)

                Dim buffer As Byte() = New Byte(CInt(sfs.Length) - 1) {}
                sfs.Read(buffer, 0, buffer.Length)
                sfs.Close()
            End If

            SqlTrans.Commit()
            SQlConnectionDestino.Close()
        Catch SqlException As OleDb.OleDbException
            Select Case SqlException.ErrorCode
                Case 2627
                    Throw New Exception("Registro já cadastrado!")
                Case 647
                    Throw New Exception("Registro esta sendo utilizado em outras tabelas!")
                Case Else
                    Throw New Exception(SqlException.Message)
            End Select

            SqlTrans.Rollback()
            SQlConnectionDestino.Close()
        End Try
    End Sub


    Protected Sub lnkExcel_Click(sender As Object, e As EventArgs) Handles lnkExcel.Click
        Gerar_Excel()
    End Sub


    Private Sub Gerar_Excel()
        Try
            sql = "create table #Temp (Empresa varchar(18), EndEmpresa int, Cliente varchar(18), EndCliente int, Nome varchar(60), " & vbCrLf &
                    "EntradaSaida varchar(1), Serie varchar(3), Nota int, ChaveNfe varchar(50))" & vbCrLf

            sql &= "insert into #Temp (Empresa, EndEmpresa, Cliente, EndCliente, Nome, EntradaSaida, Serie, Nota, ChaveNfe)" & vbCrLf &
                    "SELECT n.Empresa_Id, n.EndEmpresa_Id, n.Cliente_id, n.EndCliente_Id, isnull(c.Nome,'') AS Nome, n.EntradaSaida_Id , n.Serie_Id, n.Nota_Id, nfe.ChaveNfe " & vbCrLf &
                    "FROM Insoja.dbo.NOTASFISCAIS N " & vbCrLf &
                    "	INNER JOIN Insoja.dbo.NFERealizadas NFE " & vbCrLf &
                    "		   ON NFE.Empresa_Id    = n.Empresa_Id" & vbCrLf &
                    "		AND NFE.EndEmpresa_Id   = n.EndEmpresa_Id" & vbCrLf &
                    "		AND NFE.Cliente_Id      = n.Cliente_Id" & vbCrLf &
                    "		AND NFE.EndCliente_Id   = n.EndCliente_Id" & vbCrLf &
                    "		AND NFE.EntradaSaida_Id = n.EntradaSaida_Id" & vbCrLf &
                    "		AND NFE.Serie_Id        = n.Serie_Id" & vbCrLf &
                    "		AND NFE.Nota_Id         = n.Nota_Id" & vbCrLf &
                    "	LEFT JOIN Insoja.dbo.CLIENTES C" & vbCrLf &
                    "			ON C.Cliente_Id = n.Cliente_Id" & vbCrLf &
                    "			AND C.Endereco_Id = n.EndCliente_Id" & vbCrLf &
                    "WHERE N.SITUACAO = 1" & vbCrLf &
                    "AND N.MOVIMENTO between '2020-01-01' and '2020-12-31'" & vbCrLf &
                    "AND LEFT(n.Empresa_Id,8) = '04440724'" & vbCrLf


            sql &= "insert into #Temp (Empresa, EndEmpresa, Cliente, EndCliente, Nome, EntradaSaida, Serie, Nota, ChaveNfe)" & vbCrLf &
                    "SELECT n.Empresa_Id, n.EndEmpresa_Id, n.Cliente_id, n.EndCliente_Id, isnull(c.Nome,'') AS Nome, n.EntradaSaida_Id , n.Serie_Id, n.Nota_Id, nfe.ChaveNfe " & vbCrLf &
                    "FROM Insol.dbo.NOTASFISCAIS N " & vbCrLf &
                    "	INNER JOIN Insol.dbo.NFERealizadas NFE " & vbCrLf &
                    "		   ON NFE.Empresa_Id    = n.Empresa_Id" & vbCrLf &
                    "		AND NFE.EndEmpresa_Id   = n.EndEmpresa_Id" & vbCrLf &
                    "		AND NFE.Cliente_Id      = n.Cliente_Id" & vbCrLf &
                    "		AND NFE.EndCliente_Id   = n.EndCliente_Id" & vbCrLf &
                    "		AND NFE.EntradaSaida_Id = n.EntradaSaida_Id" & vbCrLf &
                    "		AND NFE.Serie_Id        = n.Serie_Id" & vbCrLf &
                    "		AND NFE.Nota_Id         = n.Nota_Id" & vbCrLf &
                    "	LEFT JOIN Insol.dbo.CLIENTES C" & vbCrLf &
                    "			ON C.Cliente_Id = n.Cliente_Id" & vbCrLf &
                    "			AND C.Endereco_Id = n.EndCliente_Id" & vbCrLf &
                    "WHERE N.SITUACAO = 1" & vbCrLf &
                    "AND N.MOVIMENTO between '2020-01-01' and '2020-12-31'" & vbCrLf &
                    "AND LEFT(n.Empresa_Id,8) = '04440724'" & vbCrLf

            sql &= "select Empresa, EndEmpresa, Cliente, EndCliente, Nome, EntradaSaida, Serie, Nota, ChaveNfe from #Temp" & vbCrLf


            Dim ds As DataSet = Banco.ConsultaDataSet(sql, "ConsistenciaDeNotas")

            Dim fileName As String = Server.MapPath("~/PlanilhasExcel/" & Funcoes.GeraNomeArquivo & ".xlsx")

            If File.Exists(fileName) Then
                File.Delete(fileName)
            End If

            Using arquivo As New FileStream(fileName, FileMode.CreateNew)
                Using package As New ExcelPackage(arquivo)
                    'criando planilha títulos
                    Dim worksheet As ExcelWorksheet = package.Workbook.Worksheets.Add("NOTAS")

                    'criando linha com o cabeçalho da planilha
                    Dim rowIndex As Integer = 1
                    Dim columnIndex As Integer = 1

                    'criando linha com o cabeçalho da planilha
                    For Each col As DataColumn In ds.Tables(0).Columns
                        worksheet.Cells(rowIndex, columnIndex).Value = col.ColumnName
                        columnIndex += 1
                    Next

                    ''criando auto filtro na planilha
                    'worksheet.Cells("A5:H" & rowIndex).AutoFilter = True

                    'aplicando formatação nas células do cabeçalho
                    Using range = worksheet.Cells(rowIndex, 1, rowIndex, ds.Tables(0).Columns.Count)
                        range.Style.Font.Bold = True
                        range.Style.Fill.PatternType = ExcelFillStyle.Solid
                        range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
                        range.Style.Font.Color.SetColor(Color.White)
                    End Using
                    rowIndex += 1

                    ' criando conteúdo da planilha com os dados do dataset
                    For Each row As DataRow In ds.Tables(0).Rows
                        columnIndex = 1
                        For Each col As DataColumn In ds.Tables(0).Columns

                            If col.ColumnName = "Empresa" OrElse col.ColumnName = "Cliente" Then
                                worksheet.Cells(rowIndex, columnIndex).Value = Convert.ToDouble(row(col.ColumnName)).ToString()
                            Else
                                worksheet.Cells(rowIndex, columnIndex).Value = row(col.ColumnName)
                            End If

                            columnIndex += 1
                        Next

                        'aplicando formatação nas células do conteúdo
                        If rowIndex Mod 2 = 0 Then
                            Using range = worksheet.Cells(rowIndex, 1, rowIndex, ds.Tables(0).Columns.Count)
                                range.Style.Fill.PatternType = ExcelFillStyle.Solid
                                range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(153, 204, 255))
                            End Using
                        End If
                        rowIndex += 1
                    Next

                    'setando autofit nas células da planilha
                    worksheet.Cells.AutoFitColumns(0)

                    'congelando primeira linha (cabeçalho)
                    worksheet.View.FreezePanes(2, 1)

                    'salvando planilha do excel
                    package.Save()
                End Using
            End Using
            Funcoes.AbrirExcel(Me.Page, "PlanilhasExcel/" & Path.GetFileName(fileName))
            'download do arquivo pelo browser
            'Download(fileName)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub Ler_Excel()

        Dim fileName As String = Server.MapPath("~/PlanilhasExcel/ListaDePrecos.xlsx")

        Using arquivo As New FileStream(fileName, FileMode.Open, FileAccess.Read)
            Using package As New ExcelPackage(arquivo)

                'Aqui vai Ler as Abas 
                Dim i As Integer
                For i = 0 To package.Workbook.Worksheets.Count - 1

                    Dim aba As ExcelWorksheet = package.Workbook.Worksheets.Item(i)


                    'Aqui vair ler as Linhas


                    'Aqui vai ler as colunas


                Next


            End Using
        End Using

        MsgBox(Me.Page, "Planilha lida com sucesso.")

    End Sub

    Protected Sub lnkImporta_Click(sender As Object, e As EventArgs) Handles lnkImporta.Click
        'LerExcelXLS()
        'LerContasBancarias()
        'LerProdutosBaxi()
        'LerPlanilhaClientes()
        RemoverClientes()
    End Sub

    Private Sub LerExcelXLS()

        Dim temClientes As String = String.Empty
        Dim impClientes As String = String.Empty
        Dim valida As String = String.Empty
        Dim duplicados As String = String.Empty


        objListaCliente = New ListCliente()

        Dim caminhoArquivoExcel As String = "C:\NGS\BAXI\Fornecedores.xls"
        Dim nomePlanilhaExcel As String = "Sheet1" & "$"

        Dim conexaoOleDb As OleDbConnection = Nothing
        Dim ds As DataSet
        Dim cmd As OleDbDataAdapter

        Try
            conexaoOleDb = New OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0;Data Source= " & caminhoArquivoExcel & ";Extended Properties='Excel 8.0;HDR=Yes'")
            cmd = New OleDbDataAdapter("Select * from [" & nomePlanilhaExcel & "]", conexaoOleDb)
            cmd.TableMappings.Add("Table", "tabelaExcel")
            ds = New DataSet
            cmd.Fill(ds)

            If Not ds Is Nothing AndAlso ds.Tables(0).Rows.Count > 0 Then

                For Each row As DataRow In ds.Tables(0).Rows

                    Dim verCli As Cliente = New Cliente(row("CNPJ/CPF"), 0)

                    If Not verCli.Nome Is Nothing AndAlso verCli.Nome.Length > 0 Then
                        temClientes += verCli.Codigo & "-" & verCli.Nome & ";"
                        Continue For
                    End If

                    valida = row("CNPJ/CPF") & ";"

                    If impClientes.Contains(valida) Then
                        duplicados += valida
                        Continue For
                    End If

                    impClientes += row("CNPJ/CPF") & ";"

                    objCliente = New Cliente()

                    objCliente.Codigo = row("CNPJ/CPF")

                    objCliente.IUD = "I"

                    objCliente.CodigoEndereco = 0

                    objCliente.CodigoSituacao = 1

                    objCliente.CodigoEstado = row("Estado")

                    objCliente.Nome = Left(Funcoes.EliminarCaracteresEspeciais(row("Nome")), 80).ToUpper
                    objCliente.Fantasia = Left(Funcoes.EliminarCaracteresEspeciais(row("Fantasia")), 60).ToUpper

                    If Not IsDBNull(row("Endereco")) Then
                        objCliente.Endereco = Left(Funcoes.EliminarCaracteresEspeciais(row("Endereco")), 60).ToUpper
                    End If

                    If Not IsDBNull(row("Numero")) Then
                        Dim data = Funcoes.RemoverLetrasDoNumero(row("Numero").ToUpper)
                        data = Funcoes.EliminarCaracteresEspeciais(data)
                        If data.Length > 0 Then
                            objCliente.Numero = data
                        Else
                            objCliente.Numero = 0
                        End If
                    End If

                    If Not IsDBNull(row("Complemento")) Then
                        objCliente.Complemento = Left(Funcoes.EliminarCaracteresEspeciais(row("Complemento")), 60).ToUpper
                    End If

                    If Not IsDBNull(row("Bairro")) Then
                        objCliente.Bairro = Left(row("Bairro"), 60)
                    End If

                    If Not IsDBNull(row("Cep")) Then
                        objCliente.CEP = Funcoes.RemoverLetrasDoNumero(row("Cep"))
                    End If

                    objCliente.Cidade = Funcoes.SubstituirCaracteresEspeciaisMunicipio(row("Cidade")).ToUpper

                    Dim objMunicipio = New Municipio(objCliente.CodigoEstado, objCliente.Cidade)

                    If objMunicipio.CodigoIbge = 0 Then
                        Dim teste As String = "NAO ACHO"
                        objCliente.CodigoMunicipio = 0
                    Else
                        objCliente.CodigoMunicipio = objMunicipio.CodigoIbge
                    End If

                    If Not IsDBNull(row("Inscricao")) AndAlso row("Inscricao").ToString.Length > 0 Then
                        objCliente.InscricaoEstadual = Funcoes.SubstituirCaracteresEspeciaisMunicipio(row("Inscricao")).ToUpper
                    Else
                        objCliente.InscricaoEstadual = "ISENTO"
                    End If

                    objCliente.CodigoPais = 1058

                    If Not IsDBNull(row("Telefone")) AndAlso row("Telefone").ToString.Length > 0 Then
                        objCliente.Telefone = Trim(row("Telefone"))
                    End If

                    If Not IsDBNull(row("Email")) AndAlso row("Email").ToString.Length > 0 Then
                        objCliente.Email = row("Email")
                    End If

                    If Not IsDBNull(row("EmailNFE")) AndAlso row("EmailNFE").ToString.Length > 0 Then
                        objCliente.EmailNFE = row("EmailNFE")
                    Else
                        If Not IsDBNull(row("Email")) AndAlso row("Email").ToString.Length > 0 Then
                            objCliente.EmailNFE = row("Email")
                        End If
                    End If

                    If Not IsDBNull(row("OutrosTelefones")) AndAlso row("OutrosTelefones").ToString.Length > 0 Then
                        objCliente.OutrosTelefones = Trim(row("OutrosTelefones"))
                    End If

                    objCliente.ClienteDesde = CDate(row("FornecedorDesde")).ToString("yyyy/MM/dd hh:mm:ss")
                    objCliente.NascimentoConstituicao = CDate(row("FornecedorDesde")).ToString("yyyy/MM/dd hh:mm:ss")

                    'objCliente.ClienteDesde = Now().ToString("yyyy-MM-dd")
                    'objCliente.NascimentoConstituicao = Now().ToString("yyyy-MM-dd")

                    objCliente.CodigoCategoria = 17

                    If objCliente.CodigoEstado = "PR" Or objCliente.CodigoEstado = "SC" Or objCliente.CodigoEstado = "RS" Then
                        objCliente.CodigoRegiao = 18
                    ElseIf objCliente.CodigoEstado = "SP" Or objCliente.CodigoEstado = "RJ" Or objCliente.CodigoEstado = "MG" Or objCliente.CodigoEstado = "ES" Then
                        objCliente.CodigoRegiao = 17
                    ElseIf objCliente.CodigoEstado = "MS" Or objCliente.CodigoEstado = "MT" Or objCliente.CodigoEstado = "DF" Or objCliente.CodigoEstado = "GO" Then
                        objCliente.CodigoRegiao = 16
                    ElseIf objCliente.CodigoEstado = "BA" Or objCliente.CodigoEstado = "SE" Or objCliente.CodigoEstado = "AL" Or objCliente.CodigoEstado = "PE" Or objCliente.CodigoEstado = "PB" Or objCliente.CodigoEstado = "RN" Or objCliente.CodigoEstado = "CE" Or objCliente.CodigoEstado = "PI" Or objCliente.CodigoEstado = "MA" Then
                        objCliente.CodigoRegiao = 19
                    ElseIf objCliente.CodigoEstado = "TO" Or objCliente.CodigoEstado = "RO" Or objCliente.CodigoEstado = "PA" Or objCliente.CodigoEstado = "AC" Or objCliente.CodigoEstado = "AP" Or objCliente.CodigoEstado = "RR" Or objCliente.CodigoEstado = "AM" Then
                        objCliente.CodigoRegiao = 15
                    Else
                        objCliente.CodigoRegiao = 1
                    End If

                    Dim objTipo As New [Lib].Negocio.ClientexTipo(objCliente)
                    objTipo = New [Lib].Negocio.ClientexTipo(objCliente)
                    objTipo.IUD = "I"
                    objTipo.CodigoTipo = 5
                    objCliente.Tipos.Add(objTipo)

                    'If row("TIPO") = "FFF" Then
                    '    objTipo = New [Lib].Negocio.ClientexTipo(objCliente)
                    '    objTipo.IUD = "I"
                    '    objTipo.CodigoTipo = 5
                    '    objCliente.Tipos.Add(objTipo)
                    'ElseIf row("TIPO") = "TTT" Then
                    '    objTipo = New [Lib].Negocio.ClientexTipo(objCliente)
                    '    objTipo.IUD = "I"
                    '    objTipo.CodigoTipo = 7
                    '    objCliente.Tipos.Add(objTipo)
                    'Else

                    'End If

                    objCliente.UsuarioInclusao = "ADMIN"
                    objCliente.UsuarioInclusaoData = Now().ToString("yyyy-MM-dd hh:mm:ss")

                    objListaCliente.Add(objCliente)

                Next

                If objListaCliente.Count > 0 Then
                    If objListaCliente.Salvar() Then
                        MsgBox(Me.Page, "gravado com sucesso")

                        'EmitirExcel(dsCleintesExcel)
                    Else
                        MsgBox(Me.Page, HttpContext.Current.Session("ssMessage").ToString)
                    End If
                Else
                    MsgBox(Me.Page, "Lista não encontrada")
                End If

            End If

        Catch ex As Exception
            MsgBox(Me.Page, "Erro" & ex.ToString)
        Finally
            conexaoOleDb.Close()
        End Try
    End Sub


    Private Sub LerContasBancarias()

        Dim repetido As Boolean

        Dim naoTemClientes As String = String.Empty
        Dim temConta As String = String.Empty
        SqlArray.Clear()

        Dim caminhoArquivoExcel As String = "C:\NGS\BAXI\ContasBancariasDeFornecedores.xls"
        Dim nomePlanilhaExcel As String = "Sheet1" & "$"

        Dim conexaoOleDb As OleDbConnection = Nothing
        Dim ds As DataSet
        Dim cmd As OleDbDataAdapter

        Try
            conexaoOleDb = New OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0;Data Source= " & caminhoArquivoExcel & ";Extended Properties='Excel 8.0;HDR=Yes'")
            cmd = New OleDbDataAdapter("Select * from [" & nomePlanilhaExcel & "]", conexaoOleDb)
            cmd.TableMappings.Add("Table", "tabelaExcel")
            ds = New DataSet
            cmd.Fill(ds)

            For Each row As DataRow In ds.Tables(0).Rows

                Dim verCli As Cliente = New Cliente(row("CPJ/CNPJ"), 0)

                If verCli.Nome Is Nothing OrElse verCli.Nome.Length = 0 Then
                    naoTemClientes += row("CPJ/CNPJ") & ";"
                    Continue For
                End If

                Dim conta As ClienteXContaBancaria = New ClienteXContaBancaria()
                conta.Cliente = verCli

                Dim nbanco As String()
                nbanco = row("Banco").ToString.Split("-")

                Dim nrConta As String = String.Empty

                If row("ContaCorrente").ToString.Contains("-") Then
                    Dim nConta As String()
                    nConta = row("ContaCorrente").ToString.Split("-")
                    nrConta = nConta(0)
                Else
                    nrConta = row("ContaCorrente")
                End If


                repetido = False

                For Each cTa In verCli.ContasBancarias
                    If cTa.CodigoBanco = CInt(nbanco(0)) AndAlso cTa.CodigoAgencia = CInt(row("Agencia")) AndAlso Trim(cTa.ContaCorrente) = Trim(nrConta) Then
                        temConta += verCli.Codigo & "-" & verCli.Nome & ";"
                        repetido = True
                        Continue For
                    End If
                Next

                If repetido Then
                    Continue For
                End If


                sql = " insert into ClientesXContasBancarias (Cliente_id, Endereco_id, Banco_id, Agencia_id, DigitoAgencia_id, ContaCorrente_id, DigitoConta_id, TipoConta, Ativo) " & vbCrLf &
                      " values ('" & verCli.Codigo & "'," & verCli.CodigoEndereco & ", " & CInt(nbanco(0)) & ", " & CInt(row("Agencia")) & ", '" & row("DigitoAgencia") & "', '" & nrConta & "', '" & row("DigitoConta") & "', 'C',1);"

                SqlArray.Add(sql)

                'conta.IUD = "I"
                'conta.Ativo = True

                'Dim nbanco As String()
                'nbanco = row("Banco").ToString.Split("-")
                'conta.CodigoBanco = nbanco(0)

                'conta.CodigoAgencia = row("Agencia")

                'If Not IsDBNull(row("DigitoAgencia")) AndAlso row("DigitoAgencia").ToString.Length > 0 Then
                '    conta.DigitoAgencia = row("DigitoAgencia")
                'Else
                '    conta.DigitoAgencia = ""
                'End If



                'If Not IsDBNull(row("DigitoConta")) AndAlso row("DigitoConta").ToString.Length > 0 Then
                '    conta.DigitoConta = row("DigitoConta")
                'Else
                '    conta.DigitoConta = ""
                'End If

                'conta.TipoConta = "C"

                'conta.Observacoes = ""

                'conta.SalvarSql(SqlArray)
            Next

            If SqlArray.Count > 0 Then
                If Banco.GravaBanco(SqlArray) Then
                    MsgBox(Me.Page, "gravado com sucesso")
                Else
                    MsgBox(Me.Page, HttpContext.Current.Session("ssMessage").ToString)
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, "Erro" & ex.ToString)
        Finally
            conexaoOleDb.Close()
        End Try

    End Sub


    Private Sub LerProdutosBaxi()

        Dim TESTE As String

        SqlArray.Clear()

        Dim caminhoArquivoExcel As String = "C:\NGS\Baxi\Produtos\PRODUTOSBAXI.xlsx"
        Dim nomePlanilhaExcel As String = "Sheet1" & "$"

        'Dim conexaoOleDb As OleDbConnection = Nothing
        'Dim ds As DataSet
        'Dim cmd As OleDbDataAdapter

        Try

            'conexaoOleDb = New OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0;Data Source= " & caminhoArquivoExcel & ";Extended Properties='Excel 8.0;HDR=Yes'")
            'cmd = New OleDbDataAdapter("Select * from [" & nomePlanilhaExcel & "]", conexaoOleDb)
            'cmd.TableMappings.Add("Table", "tabelaExcel")
            'ds = New DataSet
            'cmd.Fill(ds)

            'Dim i As String = "201010001"

            Dim ds As DataSet
            ds = FuncoesStrings.LerExcelParaDataSet(caminhoArquivoExcel)

            If ds IsNot Nothing AndAlso ds.Tables(0).Rows.Count > 0 Then
                For Each row As DataRow In ds.Tables(0).Rows


                    'If row("CodigoProduto") = "401010135" Then
                    '    TESTE = row("CodigoProduto")
                    'End If


                    Dim prd As New Produto()

                    prd.IUD = "I"
                    TESTE = row("CodigoProduto")
                    prd.Codigo = row("CodigoProduto")
                    TESTE = row("CodigoProduto") & row("Grupo")
                    prd.CodigoGrupo = Trim(row("Grupo"))

                    Dim unid As New ProdutosXUnidadeDeComercializacao()

                    TESTE = row("CodigoProduto") & row("Unidade")

                    If Trim(row("Unidade")) = "KG" Then
                        prd.CodigoEmbalagem = 6
                        prd.PesoQuantidade = "P"

                        prd.Unidade = "KG"
                        unid.IUD = "I"
                        unid.CodigoProduto = prd.Codigo
                        unid.CodigoUnidade = prd.Unidade
                        unid.FatorConversao = row("FatorConversao")
                        unid.PesoDaEmbalagem = row("PesoEmbalagem")
                        unid.Ativo = True
                        prd.UnidadesDeComercializacao.Add(unid)

                    ElseIf Trim(row("Unidade")) = "SC" Then
                        prd.CodigoEmbalagem = 54
                        prd.PesoQuantidade = "Q"

                        prd.Unidade = "SC"
                        unid.IUD = "I"
                        unid.CodigoProduto = prd.Codigo
                        unid.CodigoUnidade = prd.Unidade
                        unid.FatorConversao = row("FatorConversao")
                        unid.PesoDaEmbalagem = row("PesoEmbalagem")
                        unid.Ativo = True
                        prd.UnidadesDeComercializacao.Add(unid)

                    ElseIf Trim(row("Unidade")) = "FRD" Then
                        prd.CodigoEmbalagem = 76
                        prd.PesoQuantidade = "Q"

                        prd.Unidade = "FRD"
                        unid.IUD = "I"
                        unid.CodigoProduto = prd.Codigo
                        unid.CodigoUnidade = prd.Unidade
                        unid.FatorConversao = 1
                        unid.Ativo = True
                        prd.UnidadesDeComercializacao.Add(unid)

                    ElseIf Trim(row("Unidade")) = "GRN" Then
                        prd.CodigoEmbalagem = 1
                        prd.PesoQuantidade = "P"

                        prd.Unidade = "GRN"
                        unid.IUD = "I"
                        unid.CodigoProduto = prd.Codigo
                        unid.CodigoUnidade = prd.Unidade
                        unid.FatorConversao = 1
                        unid.Ativo = True
                        prd.UnidadesDeComercializacao.Add(unid)

                    ElseIf Trim(row("Unidade")) = "L" Then
                        prd.CodigoEmbalagem = 5
                        prd.PesoQuantidade = "P"

                        prd.Unidade = "LTS"
                        unid.IUD = "I"
                        unid.CodigoProduto = prd.Codigo
                        unid.CodigoUnidade = prd.Unidade
                        unid.FatorConversao = 1
                        unid.Ativo = True
                        prd.UnidadesDeComercializacao.Add(unid)

                    ElseIf Trim(row("Unidade")) = "PEÇ" Then
                        prd.CodigoEmbalagem = 77
                        prd.PesoQuantidade = "P"

                        prd.Unidade = "PCS"
                        unid.IUD = "I"
                        unid.CodigoProduto = prd.Codigo
                        unid.CodigoUnidade = prd.Unidade
                        unid.FatorConversao = 1
                        unid.Ativo = True
                        prd.UnidadesDeComercializacao.Add(unid)

                    ElseIf Trim(row("Unidade")) = "UN" Then
                        prd.CodigoEmbalagem = 7
                        prd.PesoQuantidade = "P"

                        prd.Unidade = "UN"
                        unid.IUD = "I"
                        unid.CodigoProduto = prd.Codigo
                        unid.CodigoUnidade = prd.Unidade
                        unid.FatorConversao = 1
                        unid.Ativo = True
                        prd.UnidadesDeComercializacao.Add(unid)

                    ElseIf Trim(row("Unidade")) = "CDA" Then
                        prd.CodigoEmbalagem = 7
                        prd.PesoQuantidade = "P"

                        prd.Unidade = "UN"
                        unid.IUD = "I"
                        unid.CodigoProduto = prd.Codigo
                        unid.CodigoUnidade = prd.Unidade
                        unid.FatorConversao = 1
                        unid.Ativo = True
                        prd.UnidadesDeComercializacao.Add(unid)

                    ElseIf Trim(row("Unidade")) = "M3" Then
                        prd.CodigoEmbalagem = 73
                        prd.PesoQuantidade = "P"

                        prd.Unidade = "M3"
                        unid.IUD = "I"
                        unid.CodigoProduto = prd.Codigo
                        unid.CodigoUnidade = prd.Unidade
                        unid.FatorConversao = 1
                        unid.Ativo = True
                        prd.UnidadesDeComercializacao.Add(unid)

                    ElseIf Trim(row("Unidade")) = "PAR" Then
                        prd.CodigoEmbalagem = 78
                        prd.PesoQuantidade = "P"

                        prd.Unidade = "PR"
                        unid.IUD = "I"
                        unid.CodigoProduto = prd.Codigo
                        unid.CodigoUnidade = prd.Unidade
                        unid.FatorConversao = 1
                        unid.Ativo = True
                        prd.UnidadesDeComercializacao.Add(unid)

                    ElseIf Trim(row("Unidade")) = "M" Then
                        prd.CodigoEmbalagem = 79
                        prd.PesoQuantidade = "P"

                        prd.Unidade = "MTS"
                        unid.IUD = "I"
                        unid.CodigoProduto = prd.Codigo
                        unid.CodigoUnidade = prd.Unidade
                        unid.FatorConversao = 1
                        unid.Ativo = True
                        prd.UnidadesDeComercializacao.Add(unid)

                    ElseIf Trim(row("Unidade")) = "G" Then
                        prd.CodigoEmbalagem = 34
                        prd.PesoQuantidade = "P"

                        prd.Unidade = "GL"
                        unid.IUD = "I"
                        unid.CodigoProduto = prd.Codigo
                        unid.CodigoUnidade = prd.Unidade
                        unid.FatorConversao = 1
                        unid.Ativo = True
                        prd.UnidadesDeComercializacao.Add(unid)

                    ElseIf Trim(row("Unidade")) = "M2" Then
                        prd.CodigoEmbalagem = 80
                        prd.PesoQuantidade = "P"

                        prd.Unidade = "M2"
                        unid.IUD = "I"
                        unid.CodigoProduto = prd.Codigo
                        unid.CodigoUnidade = prd.Unidade
                        unid.FatorConversao = 1
                        unid.Ativo = True
                        prd.UnidadesDeComercializacao.Add(unid)

                    ElseIf Trim(row("Unidade")) = "JG" Then
                        prd.CodigoEmbalagem = 81
                        prd.PesoQuantidade = "P"

                        prd.Unidade = "JG"
                        unid.IUD = "I"
                        unid.CodigoProduto = prd.Codigo
                        unid.CodigoUnidade = prd.Unidade
                        unid.FatorConversao = 1
                        unid.Ativo = True
                        prd.UnidadesDeComercializacao.Add(unid)

                    End If

                    prd.Etapa = 1
                    prd.Qualidade = 1
                    prd.Agrupar = "S"
                    prd.CodigoCarteiraCompra = "001001001"
                    prd.CodigoCarteiraVenda = "002005026"
                    prd.ControlarNumeroDoLote = True

                    TESTE = row("CodigoProduto") & row("Nome")

                    Dim pNome As String = Funcoes.EliminarCaracteresEspeciais(Trim(row("Nome")))

                    prd.Nome = Left(pNome, 100)

                    TESTE = row("CodigoProduto") & row("Descricao")

                    Dim pDescricao As String = Funcoes.EliminarCaracteresEspeciais(Trim(row("Descricao")))
                    prd.Descricao = pDescricao

                    prd.DescricaoMapa = Left(pNome, 20)

                    TESTE = row("CodigoProduto") & row("NCM")
                    prd.NCM = row("NCM").ToString.Replace(".", "")

                    prd.TipoItem = 4 ' 4-PRODUTO ACABADO   06 - Produto Intermediario  7 - Material Uso e Consumo   2 - Embalagem 

                    prd.CodigoSituacao = 1

                    prd.CodigoEstadoFisico = 0

                    prd.ControlarEstoque = True
                    prd.UsuarioInclusao = "ADMIN"
                    prd.UsuarioInclusaoData = Now()
                    prd.EstoqueMinimo = 0
                    prd.QuantidadeCaixa = 0
                    prd.Qualidade = 1
                    prd.IPI = 0
                    prd.CodigoGenero = 0
                    prd.CodigoEX = ""
                    prd.Fitossanitario = False
                    prd.ControlarEmbalagem = False
                    prd.ProdutoIndea = ""
                    prd.ControlarPecas = False
                    prd.ControlarPrecoDePauta = False
                    prd.SubCodigoGenero = 0
                    prd.ControlarRomaneio = False
                    prd.ControlarPesagem = False
                    prd.ControlarDecimais = False
                    prd.Almoxarifado = False
                    prd.PrecoDoProduto = False
                    prd.ControlarLote = False
                    prd.CodigoEstadoFisico = 0
                    prd.InfaDProd = ""
                    prd.AutorizacaoDeRetirada = False
                    prd.RegistroMinisterioAgricultura = ""
                    prd.CodigoProdutoTerceiro = ""
                    prd.CustoIndireto = False
                    prd.Gtin8 = ""
                    prd.Gtin12 = ""
                    prd.Gtin13 = ""
                    prd.Gtin14 = ""

                    prd.salvarSql(SqlArray)

                    'i = i + 1

                Next

                If SqlArray.Count > 0 Then
                    If Banco.GravaBanco(SqlArray) Then
                        MsgBox(Me.Page, "gravado com sucesso")
                    Else
                        MsgBox(Me.Page, HttpContext.Current.Session("ssMessage").ToString)
                    End If
                End If

            Else
                MsgBox(Me.Page, "Lista não encontrada")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try

    End Sub

    Private Sub LerPlanilhaClientes()
        SqlArray.Clear()

        Dim strm As New StreamWriter(HttpContext.Current.Server.MapPath("~/Files/gravaclientes.txt"))

        Dim caminhoArquivoExcel As String = "C:\NGS\Horus\Cadastro\CadastroDeCliFor.xlsx"
        Dim nomePlanilhaExcel As String = "Sheet1" & "$"

        Try
            Dim ds As DataSet
            ds = FuncoesStrings.LerExcelParaDataSet(caminhoArquivoExcel)

            If ds IsNot Nothing AndAlso ds.Tables(0).Rows.Count > 0 Then
                For Each row As DataRow In ds.Tables(0).Rows

                    Dim verCliente As String = row("CPFCNPJ").ToString.Replace(".", "").Replace("-", "").Replace("/", "")

                    If verCliente.Length = 11 AndAlso Not Funcoes.ValidaCPF(verCliente) Then
                        strm.WriteLine("ERRO: CPF  " & verCliente & " inválido.")
                        Continue For
                    ElseIf verCliente.Length = 14 AndAlso Not Funcoes.ValidaCNPJ(verCliente) Then
                        strm.WriteLine("ERRO: CNPJ " & verCliente & " inválido.")
                        Continue For
                    End If

                    Dim objCliente As New Cliente(verCliente, 0)

                    If objCliente.Nome Is Nothing OrElse objCliente.Nome.Length = 0 Then

                        objCliente.IUD = "I"
                        objCliente.Codigo = verCliente
                        objCliente.CodigoEndereco = 0
                        objCliente.Reduzido = Trim(row("REDUZIDO"))
                        objCliente.Nome = UCase(Funcoes.EliminarCaracteresEspeciais(RTrim(row("NOME"))))
                        objCliente.Fantasia = UCase(Funcoes.EliminarCaracteresEspeciais(RTrim(row("FANTASIA"))))
                        objCliente.Endereco = UCase(Funcoes.EliminarCaracteresEspeciais(RTrim(row("ENDERECO"))))

                        If CInt(Trim(row("REDUZIDO"))) > 0 Then
                            objCliente.Numero = CInt(Trim(row("REDUZIDO")))
                        End If

                        If row("COMPLEMENTO").ToString.Length > 0 Then
                            objCliente.Complemento = UCase(Funcoes.EliminarCaracteresEspeciais(RTrim(row("COMPLEMENTO"))))
                        End If

                        objCliente.Bairro = UCase(Funcoes.EliminarCaracteresEspeciais(RTrim(row("BAIRRO"))))
                        objCliente.CEP = Funcoes.EliminarCaracteresEspeciais(RTrim(row("CEP")))

                        objCliente.Cidade = UCase(Funcoes.EliminarCaracteresEspeciais(RTrim(row("CIDADE"))))
                        objCliente.CodigoEstado = UCase(Funcoes.EliminarCaracteresEspeciais(RTrim(row("ESTADO"))))

                        Dim objMunicipio = New Municipio(objCliente.CodigoEstado, objCliente.Cidade)

                        If objMunicipio.CodigoIbge = 0 Then
                            strm.WriteLine("ERRO: NAO ENCONTROU MUNICIPIO - " & verCliente)
                            Continue For
                        Else
                            objCliente.CodigoMunicipio = objMunicipio.CodigoIbge
                        End If

                        If objCliente.CodigoEstado = "PR" Or objCliente.CodigoEstado = "SC" Or objCliente.CodigoEstado = "RS" Then
                            objCliente.CodigoRegiao = 18
                        ElseIf objCliente.CodigoEstado = "SP" Or objCliente.CodigoEstado = "RJ" Or objCliente.CodigoEstado = "MG" Or objCliente.CodigoEstado = "ES" Then
                            objCliente.CodigoRegiao = 17
                        ElseIf objCliente.CodigoEstado = "MS" Or objCliente.CodigoEstado = "MT" Or objCliente.CodigoEstado = "DF" Or objCliente.CodigoEstado = "GO" Then
                            objCliente.CodigoRegiao = 16
                        ElseIf objCliente.CodigoEstado = "BA" Or objCliente.CodigoEstado = "SE" Or objCliente.CodigoEstado = "AL" Or objCliente.CodigoEstado = "PE" Or objCliente.CodigoEstado = "PB" Or objCliente.CodigoEstado = "RN" Or objCliente.CodigoEstado = "CE" Or objCliente.CodigoEstado = "PI" Or objCliente.CodigoEstado = "MA" Then
                            objCliente.CodigoRegiao = 19
                        ElseIf objCliente.CodigoEstado = "TO" Or objCliente.CodigoEstado = "RO" Or objCliente.CodigoEstado = "PA" Or objCliente.CodigoEstado = "AC" Or objCliente.CodigoEstado = "AP" Or objCliente.CodigoEstado = "RR" Or objCliente.CodigoEstado = "AM" Then
                            objCliente.CodigoRegiao = 15
                        Else
                            objCliente.CodigoRegiao = 1
                        End If

                        objCliente.CodigoPais = 1058

                        If row("FONE").ToString.Length > 0 Then
                            objCliente.Telefone = Funcoes.EliminarCaracteresEspeciais(RTrim(row("FONE")))
                        ElseIf row("CELULAR1").ToString.Length > 0 Then
                            objCliente.Telefone = Funcoes.EliminarCaracteresEspeciais(RTrim(row("CELULAR1")))
                        End If

                        objCliente.ClienteDesde = CDate(Trim(row("DATACADASTRO")))
                        objCliente.NascimentoConstituicao = CDate(Trim(row("DATACADASTRO")))

                        If Funcoes.EliminarCaracteresEspeciais(RTrim(row("RGIE"))) = "ISENTO" Then
                            objCliente.InscricaoEstadual = "ISENTO"
                        ElseIf verCliente.Length = 14 Then
                            objCliente.InscricaoEstadual = Funcoes.EliminarCaracteresEspeciais(RTrim(row("RGIE")))
                        ElseIf verCliente.Length = 11 Then
                            objCliente.RG = Funcoes.EliminarCaracteresEspeciais(RTrim(row("RGIE")))
                        End If

                        If verCliente.Length = 11 Then
                            objCliente.Sexo = RTrim(row("SEXO"))

                            If RTrim(row("ESTADOCIVIL")) = "S" Then
                                objCliente.EstadoCivil = "Solteiro"
                            ElseIf RTrim(row("ESTADOCIVIL")) = "C" Then
                                objCliente.EstadoCivil = "Casado"
                            End If
                        End If

                        If RTrim(row("PRODUTORURAL")) = "S" Then
                            objCliente.CodigoCategoria = 1

                            Dim objTipo As New [Lib].Negocio.ClientexTipo(objCliente)

                            objTipo = New [Lib].Negocio.ClientexTipo(objCliente)
                            objTipo.IUD = "I"
                            objTipo.CodigoTipo = 4 'CLIENTE
                            objCliente.Tipos.Add(objTipo)

                            objTipo = New [Lib].Negocio.ClientexTipo(objCliente)
                            objTipo.IUD = "I"
                            objTipo.CodigoTipo = 5 'FORNECEDOR
                            objCliente.Tipos.Add(objTipo)
                        Else

                            Dim objTipo As New [Lib].Negocio.ClientexTipo(objCliente)

                            Dim tipos() As String = RTrim(row("TIPODOCLIENTE")).ToString.Split(",")

                            For n As Integer = 0 To tipos.Count - 1

                                If tipos(n) = 2 Then
                                    objCliente.CodigoCategoria = 4

                                    objTipo = New [Lib].Negocio.ClientexTipo(objCliente)
                                    objTipo.IUD = "I"
                                    objTipo.CodigoTipo = 4 'CLIENTE
                                    objCliente.Tipos.Add(objTipo)
                                End If

                                If tipos(n) = 3 Then
                                    objCliente.CodigoCategoria = 4

                                    objTipo = New [Lib].Negocio.ClientexTipo(objCliente)
                                    objTipo.IUD = "I"
                                    objTipo.CodigoTipo = 5 'FORNECEDOR
                                    objCliente.Tipos.Add(objTipo)
                                End If

                                If tipos(n) = 4 Then
                                    objCliente.CodigoCategoria = 9

                                    objTipo = New [Lib].Negocio.ClientexTipo(objCliente)
                                    objTipo.IUD = "I"
                                    objTipo.CodigoTipo = 7 'TRANSPORTADORES
                                    objCliente.Tipos.Add(objTipo)
                                End If
                            Next

                        End If

                        objCliente.UsuarioInclusao = "ADMIN"
                        objCliente.UsuarioInclusaoData = Now().ToString("yyyy-MM-dd hh:mm:ss")

                    Else

                        objCliente.IUD = "U"
                        objCliente.UsuarioAlteracao = "ADMIN"

                    End If

                    objCliente.SalvarSql(SqlArray)

                Next
            End If

            If SqlArray.Count > 0 Then

                For n As Integer = 0 To SqlArray.Count - 1
                    strm.WriteLine(SqlArray(n))
                Next

                If Banco.GravaBanco(SqlArray) Then
                    MsgBox(Me.Page, "gravado com sucesso")
                Else
                    MsgBox(Me.Page, HttpContext.Current.Session("ssMessage").ToString)
                End If
            End If

        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        Finally
            strm.Close()
            strm.Dispose()
        End Try
    End Sub

    Private Sub RemoverClientes()

        Dim strm As New StreamWriter(HttpContext.Current.Server.MapPath("~/Files/removeclientes.txt"))

        Dim objBanco As New AcessaBanco()

        SqlArray.Clear()

        Try
            sql = "Select Cliente_Id, Endereco_Id From Clientes where UsuarioInclusao IN('FURLAN')"

            Dim dsClientes As DataSet = objBanco.ConsultaDataSet(sql, "Clientes")

            If dsClientes.Tables(0).Rows.Count > 0 Then

                For Each row As DataRow In dsClientes.Tables(0).Rows

                    Dim objCliente As New Cliente(row("Cliente_Id"), row("Endereco_Id"))

                    If objCliente.Nome Is Nothing OrElse objCliente.Nome.Length = 0 Then
                        strm.WriteLine("ERRO: NAO ENCONTROU CLIENTE - " & row("Cliente_Id"))
                        Continue For
                    Else

                        objCliente.IUD = "D"
                        objCliente.UsuarioAlteracao = "ADMIN"

                        objCliente.SalvarSql(SqlArray)
                    End If
                Next

                If SqlArray.Count > 0 Then

                    'For n As Integer = 0 To SqlArray.Count - 1
                    '    strm.WriteLine(SqlArray(n))
                    'Next

                    If Banco.GravaBanco(SqlArray) Then
                        MsgBox(Me.Page, "gravado com sucesso")
                    Else
                        MsgBox(Me.Page, HttpContext.Current.Session("ssMessage").ToString)
                    End If
                End If

            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        Finally
            strm.Close()
            strm.Dispose()
        End Try
    End Sub

    Protected Sub btnRodarExcel_Click(sender As Object, e As EventArgs)
    End Sub

    Protected Sub lnkImportaClientes_Click(sender As Object, e As EventArgs) Handles lnkImportaClientes.Click
        'LerClientesFirebird()

        LerClienteNGScomFirebird()
    End Sub

    Protected Sub lnkImportaProdutos_Click(sender As Object, e As EventArgs) Handles lnkImportaProdutos.Click
        LerProdutoFirebird()

        'LerDadosProdutoFirebird()

    End Sub


    Private Sub LerClienteNGScomFirebird()
        Dim objBanco As New AcessaBanco()

        SqlArray.Clear()

        sql = "Select Cliente_Id From Clientes where len(Cliente_Id) = 14 and Inscricao in('ISENTO','')"

        Dim dsClientes As DataSet = objBanco.ConsultaDataSet(sql, "Clientes")

        If dsClientes.Tables(0).Rows.Count > 0 Then

            Dim objBancoFirebird As New AcessaBancoFirebird

            For Each row As DataRow In dsClientes.Tables(0).Rows

                Dim cFormatado As String = Funcoes.FormatarCpfCnpj(row("Cliente_Id"))

                sql = "Select RG_IE from clientes_forn_tr where CPF_CGC = '" & cFormatado & "';"

                Dim dsC As DataSet = objBancoFirebird.ConsultaDataSet(sql, "tabela")

                If dsC IsNot Nothing AndAlso dsC.Tables(0).Rows.Count > 0 Then

                    For Each rowC As DataRow In dsC.Tables(0).Rows

                        If Not IsDBNull(rowC("RG_IE")) Then

                            Dim lInscr As String = Funcoes.RemoverLetrasDoNumero(rowC("RG_IE").ToUpper)

                            sql = "update Clientes Set Inscricao = '" & lInscr & "' where Cliente_Id = '" & row("Cliente_Id") & "';"

                            SqlArray.Add(sql)

                        End If

                    Next

                End If
            Next

        End If

        If SqlArray.Count > 0 Then
            If objBanco.GravaBanco(SqlArray) Then
                MsgBox(Me.Page, "gravado com sucesso")
            Else
                MsgBox(Me.Page, HttpContext.Current.Session("ssMessage").ToString)
            End If
        End If

    End Sub





    Private Sub LerClientesFirebird()

        Dim teste As String = ""


        Try
            objListaCliente = New ListCliente()
            Dim objListaClienteFaltantes As New ListCliente

            Dim objBancoFirebird As New AcessaBancoFirebird
            sql = "Select '' AS MENSAGEM, c.ID_CLIENTE_FORN_TR, c.CODCLI, c.CPF_CGC, c.RAZAO, c.FANTASIA, e.END_CEP, e.END_ENDERECO, e.END_NUMERO, e.END_COMPL, e.END_BAIRRO, e.END_CIDADE, e.END_UF, " & vbCrLf & _
                    "c.DATA_CADASTRO, c.FONE_1, c.E_MAIL, c.NFE_EMAIL_DANFE, c.ID_GRUPO, c.RG_IE, c.BLOQUEIO_VENDAS, c.TIPO " & vbCrLf & _
                    "from clientes_forn_tr c " & vbCrLf & _
                    "    inner join cad_end e " & vbCrLf & _
                    "        on e.id_end = c.id_endcoml" & vbCrLf & _
                    "where c.id_cliente_forn_tr  not in(134, 249, 254) and c.cpf_cgc not in ('24.450.490/0001-96', '24.450.490/0003-58', '24.450.490/0004-39', '24.450.490/0005-10', '24.450.490/0002-77', '24.450.490/0006-09', '24.450.490/0007-81', '24.450.490/0008-62', '24.450.490/0009-43', '24.450.490/0010-87', '24.450.490/0011-68', '24.450.490/0012-49', '27.153.202/0001-20', '072.775.379-71', '107.114.629-76', '064.135.499-12','01')"

            Dim ds As DataSet = objBancoFirebird.ConsultaDataSet(sql, "tabela")

            Dim dsCleintesExcel As DataTable = New DataTable()
            dsCleintesExcel.Columns.Add("MENSAGEM", Type.GetType("System.String"))
            dsCleintesExcel.Columns.Add("ID_CLIENTE_FORN_TR", Type.GetType("System.String"))
            dsCleintesExcel.Columns.Add("CODCLI", Type.GetType("System.String"))
            dsCleintesExcel.Columns.Add("CPF_CGC", Type.GetType("System.String"))
            dsCleintesExcel.Columns.Add("RG_IE", Type.GetType("System.String"))
            dsCleintesExcel.Columns.Add("RAZAO", Type.GetType("System.String"))
            dsCleintesExcel.Columns.Add("END_ENDERECO", Type.GetType("System.String"))
            dsCleintesExcel.Columns.Add("END_CIDADE", Type.GetType("System.String"))

            If ds IsNot Nothing AndAlso ds.Tables(0).Rows.Count > 0 Then
                For Each row As DataRow In ds.Tables(0).Rows
                    If Not IsDBNull(row("CPF_CGC")) Then
                        Dim objcli As String = Funcoes.RemoverLetrasDoNumero(Trim(row("CPF_CGC")))


                        teste = row("CPF_CGC")

                        If row("CPF_CGC") = "24.535.750/0001-26" Or row("CPF_CGC") = "07.575.651/0015-54" Then
                            Dim d As String = "07.575.651/0015-54"
                        End If


                        If IsDBNull(row("RAZAO")) Then
                            row("RAZAO") = ""
                        End If

                        If IsDBNull(row("END_CIDADE")) Then
                            row("MENSAGEM") = "FORMATO DA CIDADA INVALIDADA"
                            dadosClientesExcel(objListaClienteFaltantes, objcli, dsCleintesExcel, row)
                            Continue For
                        End If

                        If IsDBNull(row("END_UF")) Then
                            row("MENSAGEM") = "FORMATO DA ESTADO INVALIDADO"
                            dadosClientesExcel(objListaClienteFaltantes, objcli, dsCleintesExcel, row)
                            Continue For
                        End If

                        If objcli.Length = 0 Then
                            row("MENSAGEM") = "FORMATO DA CPF/CNPJ INVALIDADO"
                            dadosClientesExcel(objListaClienteFaltantes, objcli, dsCleintesExcel, row)
                            Continue For
                        End If

                        If Not CDbl(objcli) > 0 Then
                            row("MENSAGEM") = "FORMATO DA CPF/CNPJ INVALIDADO"
                            dadosClientesExcel(objListaClienteFaltantes, objcli, dsCleintesExcel, row)
                            Continue For
                        End If

                        If Len(objcli) = 14 Then
                            If Not Funcoes.ValidaCNPJ(objcli) Then
                                row("MENSAGEM") = "FORMATO DA CNPJ INVALIDADO"
                                dadosClientesExcel(objListaClienteFaltantes, objcli, dsCleintesExcel, row)
                                Continue For
                            End If
                        Else
                            If Not Funcoes.ValidaCPF(objcli) Then
                                row("MENSAGEM") = "FORMATO DA CPF INVALIDADO"
                                dadosClientesExcel(objListaClienteFaltantes, objcli, dsCleintesExcel, row)
                                Continue For
                            End If
                        End If

                        Dim cidade As String = Funcoes.SubstituirCaracteresEspeciaisMunicipio(row("END_CIDADE")).ToUpper

                        Dim objMunicipio = New Municipio(row("END_UF"), cidade)

                        If objMunicipio.CodigoIbge = 0 Then
                            row("MENSAGEM") = "CÓDIGO DO IBGE DA CIDADA/ESTADO NÃO ENCONTRADO"
                            dadosClientesExcel(objListaClienteFaltantes, objcli, dsCleintesExcel, row)
                            Continue For
                        End If

                        If IsDBNull(row("END_ENDERECO")) Then
                            row("MENSAGEM") = "ENDEREÇO NÃO FOI INFORMADO"
                            dadosClientesExcel(objListaClienteFaltantes, objcli, dsCleintesExcel, row)
                            Continue For
                        End If

                        objCliente = New Cliente()

                        objCliente.Codigo = objcli

                        objCliente.IUD = "I"

                        objCliente.CodigoEndereco = 0

                        If IsDBNull(row("BLOQUEIO_VENDAS")) Then
                            objCliente.CodigoSituacao = 1
                        Else
                            If row("BLOQUEIO_VENDAS") = 0 Then
                                objCliente.CodigoSituacao = 1
                            Else
                                objCliente.CodigoSituacao = 4
                            End If
                        End If

                        objCliente.Nome = Left(Funcoes.EliminarCaracteresEspeciais(row("RAZAO")), 80).ToUpper
                        objCliente.Fantasia = Left(Funcoes.EliminarCaracteresEspeciais(row("FANTASIA")), 60).ToUpper

                        If Not IsDBNull(row("END_CEP")) Then
                            objCliente.CEP = Funcoes.RemoverLetrasDoNumero(row("END_CEP"))
                        End If

                        If Not IsDBNull(row("END_ENDERECO")) Then
                            objCliente.Endereco = Left(Funcoes.EliminarCaracteresEspeciais(row("END_ENDERECO")), 60).ToUpper
                        End If

                        If Not IsDBNull(row("END_NUMERO")) Then
                            Dim data = Funcoes.RemoverLetrasDoNumero(row("END_NUMERO").ToUpper)
                            data = Funcoes.EliminarCaracteresEspeciais(data)
                            If data.Length > 0 Then
                                objCliente.Numero = data
                            Else
                                objCliente.Numero = 0
                            End If
                        End If

                        If Not IsDBNull(row("END_COMPL")) Then
                            objCliente.Complemento = Left(Funcoes.EliminarCaracteresEspeciais(row("END_COMPL")), 60).ToUpper
                        End If

                        If Not IsDBNull(row("END_BAIRRO")) Then
                            objCliente.Bairro = Left(row("END_BAIRRO"), 60)
                        End If

                        objCliente.Cidade = Funcoes.SubstituirCaracteresEspeciaisMunicipio(row("END_CIDADE")).ToUpper

                        objCliente.CodigoEstado = row("END_UF")

                        objCliente.CodigoMunicipio = objMunicipio.CodigoIbge

                        objCliente.CodigoPais = 1058
                        objCliente.ClienteDesde = CDate(row("DATA_CADASTRO")).ToString("yyyy/MM/dd hh:mm:ss")
                        objCliente.NascimentoConstituicao = CDate(row("DATA_CADASTRO")).ToString("yyyy/MM/dd hh:mm:ss")

                        If Not IsDBNull(row("FONE_1")) Then
                            objCliente.Telefone = row("FONE_1")
                        End If

                        If Not IsDBNull(row("E_MAIL")) Then
                            objCliente.Email = row("E_MAIL")
                        End If
                        If Not IsDBNull(row("NFE_EMAIL_DANFE")) Then
                            objCliente.EmailNFE = row("NFE_EMAIL_DANFE")
                        End If

                        objCliente.CodigoCategoria = 4

                        If Not IsDBNull(row("ID_GRUPO")) Then
                            If row("ID_GRUPO") = 1 Then
                                objCliente.CodigoCategoria = 6
                            ElseIf row("ID_GRUPO") = 2 Then
                                objCliente.CodigoCategoria = 1
                            ElseIf row("ID_GRUPO") = 3 Then
                                objCliente.CodigoCategoria = 18
                            ElseIf row("ID_GRUPO") = 71 Then
                                objCliente.CodigoCategoria = 17
                            ElseIf row("ID_GRUPO") = 72 Then
                                objCliente.CodigoCategoria = 9
                            End If
                        End If

                        If Not IsDBNull(row("RG_IE")) Then
                            If Len(row("CPF_CGC")) = 14 Then
                                objCliente.InscricaoEstadual = Funcoes.RemoverLetrasDoNumero(row("RG_IE").ToUpper)
                            ElseIf Len(row("CPF_CGC")) = 11 Then
                                objCliente.RG = Funcoes.RemoverLetrasDoNumero(row("RG_IE").ToUpper)
                            End If
                        End If

                        If row("TIPO") = "CCC" Then
                            objCliente.CodigoSituacao = 4
                        ElseIf row("TIPO") = "FFF" Then
                            objCliente.CodigoSituacao = 5
                        ElseIf row("TIPO") = "TTT" Then
                            objCliente.CodigoSituacao = 7
                        Else
                            objCliente.CodigoSituacao = 4
                        End If

                        If objCliente.CodigoEstado = "PR" Or objCliente.CodigoEstado = "SC" Or objCliente.CodigoEstado = "RS" Then
                            objCliente.CodigoRegiao = 18
                        ElseIf objCliente.CodigoEstado = "SP" Or objCliente.CodigoEstado = "RJ" Or objCliente.CodigoEstado = "MG" Or objCliente.CodigoEstado = "ES" Then
                            objCliente.CodigoRegiao = 17
                        ElseIf objCliente.CodigoEstado = "MS" Or objCliente.CodigoEstado = "MT" Or objCliente.CodigoEstado = "DF" Or objCliente.CodigoEstado = "GO" Then
                            objCliente.CodigoRegiao = 16
                        ElseIf objCliente.CodigoEstado = "BA" Or objCliente.CodigoEstado = "SE" Or objCliente.CodigoEstado = "AL" Or objCliente.CodigoEstado = "PE" Or objCliente.CodigoEstado = "PB" Or objCliente.CodigoEstado = "RN" Or objCliente.CodigoEstado = "CE" Or objCliente.CodigoEstado = "PI" Or objCliente.CodigoEstado = "MA" Then
                            objCliente.CodigoRegiao = 19
                        ElseIf objCliente.CodigoEstado = "TO" Or objCliente.CodigoEstado = "RO" Or objCliente.CodigoEstado = "PA" Or objCliente.CodigoEstado = "AC" Or objCliente.CodigoEstado = "AP" Or objCliente.CodigoEstado = "RR" Or objCliente.CodigoEstado = "AM" Then
                            objCliente.CodigoRegiao = 15
                        Else
                            objCliente.CodigoRegiao = 1
                        End If

                        Dim objTipo As New [Lib].Negocio.ClientexTipo(objCliente)
                        If row("TIPO") = "FFF" Then
                            objTipo = New [Lib].Negocio.ClientexTipo(objCliente)
                            objTipo.IUD = "I"
                            objTipo.CodigoTipo = 5
                            objCliente.Tipos.Add(objTipo)
                        ElseIf row("TIPO") = "TTT" Then
                            objTipo = New [Lib].Negocio.ClientexTipo(objCliente)
                            objTipo.IUD = "I"
                            objTipo.CodigoTipo = 7
                            objCliente.Tipos.Add(objTipo)
                        Else
                            objTipo = New [Lib].Negocio.ClientexTipo(objCliente)
                            objTipo.IUD = "I"
                            objTipo.CodigoTipo = 4
                            objCliente.Tipos.Add(objTipo)
                        End If

                        objCliente.UsuarioInclusao = "ADMIN"
                        objCliente.UsuarioInclusaoData = Now().ToString("yyyy-MM-dd hh:mm:ss")

                        objListaCliente.Add(objCliente)
                    End If
                Next
            End If

            If objListaClienteFaltantes.Count > 0 Then
                Dim strm As New StreamWriter(HttpContext.Current.Server.MapPath("~/Files/ClientesFaltantes.txt"))

                For Each cFaltante In objListaClienteFaltantes
                    strm.WriteLine(cFaltante.Codigo)
                Next

                strm.Close()
            End If


            If objListaCliente.Count > 0 Then
                If objListaCliente.Salvar() Then
                    MsgBox(Me.Page, "gravado com sucesso")

                    EmitirExcel(dsCleintesExcel)
                Else
                    MsgBox(Me.Page, HttpContext.Current.Session("ssMessage").ToString)
                End If
            Else
                MsgBox(Me.Page, "Lista não encontrada")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message & " cli:" & teste)
        End Try
    End Sub

    Private Sub dadosClientesExcel(ByRef objListaClienteFaltantes As ListCliente, ByRef objcli As String, ByRef dsExcel As DataTable, ByVal row As DataRow)

        If Len(objcli) = 14 Then

            Dim novoCliente As New Cliente()
            novoCliente.Codigo = objcli

            objListaClienteFaltantes.Add(novoCliente)

            Exit Sub
        End If

        Dim newRow As DataRow = dsExcel.NewRow()
        newRow("MENSAGEM") = ""
        newRow("ID_CLIENTE_FORN_TR") = ""
        newRow("CODCLI") = ""
        newRow("CPF_CGC") = ""
        newRow("RG_IE") = ""
        newRow("RAZAO") = ""
        newRow("END_ENDERECO") = ""
        newRow("END_CIDADE") = ""

        newRow("MENSAGEM") = row("MENSAGEM")
        newRow("ID_CLIENTE_FORN_TR") = row("ID_CLIENTE_FORN_TR")
        newRow("CODCLI") = row("CODCLI")
        newRow("CPF_CGC") = row("CPF_CGC")
        newRow("RG_IE") = row("RG_IE")
        newRow("RAZAO") = row("RAZAO")
        newRow("END_ENDERECO") = row("END_ENDERECO")
        newRow("END_CIDADE") = row("END_CIDADE")

        dsExcel.Rows.Add(newRow)
    End Sub

    Private Sub LerProdutoFirebird()
        Dim teste As String

        Try
            Dim objBancoFirebird As New AcessaBancoFirebird
            sql = "Select '' AS MENSAGEM, CODIGO, DESCRICAO, NCM, NUMERO_REGISTRO, ID_STATUS, ID_NATUREZA_FISICA, ID_UNIDADE, GARANTIA from produtos where codigo in('RT-083','RT-099','MOS-076','FTG-012','RT-102','LIN-085','CMO-021','RT-100','RT-103','FTG-076','ORM-018','RT-024','ORM-020','RT-038','FTF-018')"
            Dim ds As DataSet = objBancoFirebird.ConsultaDataSet(sql, "tabela")

            Dim dsProdutosExcel As DataTable = New DataTable()
            dsProdutosExcel.Columns.Add("MENSAGEM", Type.GetType("System.String"))
            dsProdutosExcel.Columns.Add("CODIGO", Type.GetType("System.String"))
            dsProdutosExcel.Columns.Add("DESCRICAO", Type.GetType("System.String"))
            dsProdutosExcel.Columns.Add("NCM", Type.GetType("System.String"))
            dsProdutosExcel.Columns.Add("NUMERO_REGISTRO", Type.GetType("System.String"))
            dsProdutosExcel.Columns.Add("ID_STATUS", Type.GetType("System.String"))
            dsProdutosExcel.Columns.Add("ID_NATUREZA_FISICA", Type.GetType("System.String"))
            dsProdutosExcel.Columns.Add("ID_UNIDADE", Type.GetType("System.String"))
            dsProdutosExcel.Columns.Add("GARANTIA", Type.GetType("System.String"))

            Dim i As String = "201010151"


            If ds IsNot Nothing AndAlso ds.Tables(0).Rows.Count > 0 Then
                For Each row As DataRow In ds.Tables(0).Rows

                    If IsDBNull(row("GARANTIA")) Then
                        row("MENSAGEM") = "GARANTIA DO PRODUTO NAO INFORMADA"
                        dadosProdutosExcel(dsProdutosExcel, row)
                        Continue For
                    End If

                    If IsDBNull(row("DESCRICAO")) Then
                        row("MENSAGEM") = "*DESCRICAO DO PRODUTO NAO INFORMADA"
                        dadosProdutosExcel(dsProdutosExcel, row)
                        Continue For
                    End If

                    teste = row("ID_UNIDADE") & " - " & i

                    If IsDBNull(row("ID_UNIDADE")) Then
                        row("MENSAGEM") = "*FORMATO DA UNIDADE INVALIDADO"
                        dadosProdutosExcel(dsProdutosExcel, row)
                        Continue For
                    ElseIf Not row("ID_UNIDADE").ToString = "46" Then
                        row("MENSAGEM") = "*FORMATO DA UNIDADE INVALIDADO"
                        dadosProdutosExcel(dsProdutosExcel, row)
                        Continue For
                    End If

                    Dim prd As New Produto()

                    prd.IUD = "I"
                    prd.Codigo = i

                    If Not IsDBNull(row("CODIGO")) Then prd.CodigoProdutoTerceiro = RTrim(row("CODIGO"))

                    prd.CodigoEmbalagem = 6
                    prd.Etapa = 1
                    prd.Qualidade = 1
                    prd.PesoQuantidade = "P"
                    prd.Agrupar = "N"
                    prd.CodigoCarteiraCompra = "001001001"
                    prd.CodigoCarteiraVenda = "002005026"
                    prd.CodigoGrupo = "20101"
                    prd.Nome = Left(row("DESCRICAO"), 100)
                    prd.Descricao = row("GARANTIA")
                    prd.DescricaoMapa = Left(row("DESCRICAO"), 20)
                    prd.CodigoGenero = 31
                    prd.SubCodigoGenero = 3
                    prd.CodigoCnae = "0119999"

                    If IsDBNull(row("NCM")) Then
                        row("MENSAGEM") = "NCM NAO INFORMADO"
                        dadosProdutosExcel(dsProdutosExcel, row)
                    Else
                        prd.NCM = row("NCM")
                    End If

                    If IsDBNull(row("NUMERO_REGISTRO")) Then
                        prd.RegistroMinisterioAgricultura = ""
                        row("MENSAGEM") = "NUMERDO DO REGISTRO NAO INFORMADO"
                        dadosProdutosExcel(dsProdutosExcel, row)
                        prd.RegistroMinisterioAgricultura = ""
                    Else
                        prd.RegistroMinisterioAgricultura = row("NUMERO_REGISTRO")
                    End If

                    prd.Unidade = "KG"
                    Dim unid As New ProdutosXUnidadeDeComercializacao()
                    unid.IUD = "I"
                    unid.CodigoProduto = i
                    unid.CodigoUnidade = prd.Unidade
                    unid.FatorConversao = 1
                    unid.Ativo = True
                    prd.UnidadesDeComercializacao.Add(unid)

                    Dim unid2 As New ProdutosXUnidadeDeComercializacao()
                    unid2.IUD = "I"
                    unid2.CodigoProduto = i
                    unid2.CodigoUnidade = "TON"
                    unid2.FatorConversao = 1000
                    unid2.Ativo = True
                    prd.UnidadesDeComercializacao.Add(unid2)

                    If IsDBNull(row("ID_STATUS")) Then
                        row("MENSAGEM") = "STATUS NAO INFORMADO - INSERIDO BLOQUEADO"
                        dadosProdutosExcel(dsProdutosExcel, row)
                        prd.CodigoSituacao = 3
                    ElseIf row("ID_STATUS") = 1 Then
                        prd.CodigoSituacao = 1
                    Else
                        prd.CodigoSituacao = 3
                    End If

                    If IsDBNull(row("ID_NATUREZA_FISICA")) Then
                        prd.CodigoEstadoFisico = 0
                    ElseIf row("ID_NATUREZA_FISICA") = 1 Then
                        prd.CodigoEstadoFisico = 8
                    ElseIf row("ID_NATUREZA_FISICA") = 2 Then
                        prd.CodigoEstadoFisico = 11
                    ElseIf row("ID_NATUREZA_FISICA") = 3 Then
                        prd.CodigoEstadoFisico = 12
                    ElseIf row("ID_NATUREZA_FISICA") = 4 Then
                        prd.CodigoEstadoFisico = 13
                    ElseIf row("ID_NATUREZA_FISICA") = 5 Then
                        prd.CodigoEstadoFisico = 14
                    Else
                        prd.CodigoEstadoFisico = 0
                    End If

                    prd.ControlarEstoque = True
                    prd.UsuarioInclusao = "ADMIN"
                    prd.UsuarioInclusaoData = Now()

                    If Not prd.Salvar Then
                        MsgBox(Me.Page, HttpContext.Current.Session("ssMessage").ToString)
                    End If

                    i = i + 1

                Next

                MsgBox(Me.Page, "gravado com sucesso")

                EmitirExcel(dsProdutosExcel)
            Else
                MsgBox(Me.Page, "Lista não encontrada")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message & " - " & teste)
        End Try
    End Sub

    Private Sub LerDadosProdutoFirebird()
        Dim teste As String = String.Empty
        Dim naoachou As String = String.Empty
        Dim garantiaNula As String = String.Empty

        SqlArray.Clear()

        Try
            Dim objBancoFirebird As New AcessaBancoFirebird
            sql = "Select '' AS MENSAGEM, CODIGO, DESCRICAO, NCM, NUMERO_REGISTRO, ID_STATUS, ID_NATUREZA_FISICA, ID_UNIDADE, GARANTIA from produtos"
            Dim ds As DataSet = objBancoFirebird.ConsultaDataSet(sql, "tabela")

            Dim dsProdutosExcel As DataTable = New DataTable()
            dsProdutosExcel.Columns.Add("MENSAGEM", Type.GetType("System.String"))
            dsProdutosExcel.Columns.Add("CODIGO", Type.GetType("System.String"))
            dsProdutosExcel.Columns.Add("DESCRICAO", Type.GetType("System.String"))
            dsProdutosExcel.Columns.Add("NCM", Type.GetType("System.String"))
            dsProdutosExcel.Columns.Add("NUMERO_REGISTRO", Type.GetType("System.String"))
            dsProdutosExcel.Columns.Add("ID_STATUS", Type.GetType("System.String"))
            dsProdutosExcel.Columns.Add("ID_NATUREZA_FISICA", Type.GetType("System.String"))
            dsProdutosExcel.Columns.Add("ID_UNIDADE", Type.GetType("System.String"))
            dsProdutosExcel.Columns.Add("GARANTIA", Type.GetType("System.String"))

            Dim i As String = "201010001"

            Dim objBanco As New AcessaBanco()

            Dim strSQL As String = String.Empty

            If ds IsNot Nothing AndAlso ds.Tables(0).Rows.Count > 0 Then
                For Each row As DataRow In ds.Tables(0).Rows

                    If Not IsDBNull(row("CODIGO")) Then

                        If IsDBNull(row("GARANTIA")) Then
                            garantiaNula += "G NULL " & RTrim(row("CODIGO")) & " - "
                        Else
                            strSQL = "Select Produto_Id From Produtos where CodigoProdutoTerceiro = '" & RTrim(row("CODIGO")) & "'"

                            Dim dsProdutos As DataSet = objBanco.ConsultaDataSet(strSQL, "Produtos")

                            If dsProdutos.Tables(0).Rows.Count > 0 Then
                                If dsProdutos.Tables(0).Rows.Count = 1 Then
                                    strSQL = "Update  Produtos set Descricao = '" & RTrim(row("GARANTIA")) & "' where Produto_Id = '" & dsProdutos.Tables(0).Rows(0).Item("Produto_Id") & "';"

                                    SqlArray.Add(strSQL)
                                Else
                                    teste += "TEM 2 " & RTrim(row("CODIGO")) & " - "
                                End If
                            Else
                                naoachou += " NAO ACHOU " + RTrim(row("CODIGO")) & " - "
                            End If

                        End If



                    End If
                Next

                If SqlArray.Count > 0 Then
                    If objBanco.GravaBanco(SqlArray) Then
                        MsgBox(Me.Page, "gravado com sucesso")
                    Else
                        MsgBox(Me.Page, HttpContext.Current.Session("ssMessage").ToString)
                    End If
                End If

            Else
                MsgBox(Me.Page, "Lista não encontrada")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message & " - " & teste)
        End Try

    End Sub

    Private Sub dadosProdutosExcel(ByRef dsExcel As DataTable, ByVal row As DataRow)
        Dim newRow As DataRow = dsExcel.NewRow()
        newRow("MENSAGEM") = ""
        newRow("CODIGO") = ""
        newRow("DESCRICAO") = ""
        newRow("NCM") = ""
        newRow("NUMERO_REGISTRO") = ""
        newRow("ID_STATUS") = ""
        newRow("ID_NATUREZA_FISICA") = ""
        newRow("ID_UNIDADE") = ""

        newRow("MENSAGEM") = row("MENSAGEM")
        newRow("CODIGO") = row("CODIGO")
        newRow("DESCRICAO") = row("DESCRICAO")
        newRow("NCM") = row("NCM")
        newRow("NUMERO_REGISTRO") = row("NUMERO_REGISTRO")
        newRow("ID_STATUS") = row("ID_STATUS")
        newRow("ID_NATUREZA_FISICA") = row("ID_NATUREZA_FISICA")
        newRow("ID_UNIDADE") = row("ID_UNIDADE")

        dsExcel.Rows.Add(newRow)
    End Sub

    Private Sub EmitirExcel(ByRef dsExcel As DataTable)
        Try
            Dim objEmpresa As New [Lib].Negocio.Cliente(Session("ssEmpresa"), Session("ssEndEmpresa"))

            Dim fileName As String = Server.MapPath("~/PlanilhasExcel/" & Funcoes.GeraNomeArquivo & ".xlsx")

            If File.Exists(fileName) Then
                File.Delete(fileName)
            End If

            Using arquivo As New FileStream(fileName, FileMode.CreateNew)
                Using package As New ExcelPackage(arquivo)
                    'criando planilha títulos
                    Dim worksheet As ExcelWorksheet = package.Workbook.Worksheets.Add("Registros invalidados.")

                    'criando linha com o cabeçalho da planilha
                    Dim rowIndex As Integer = 1
                    Dim columnIndex As Integer = 1

                    'criando linha que informa o nome da empresa e o cnpj
                    worksheet.Cells(rowIndex, columnIndex).Style.Font.Bold = True
                    worksheet.Cells(rowIndex, columnIndex).Style.Font.Color.SetColor(Color.Black)
                    worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                    worksheet.Cells(rowIndex, columnIndex).Value = String.Format("{0} - {1}", objEmpresa.Nome, Funcoes.FormatarCpfCnpj(objEmpresa.Codigo))
                    worksheet.Cells(String.Format("A{0}:E{0}", rowIndex)).Merge = True
                    worksheet.Cells(String.Format("A{0}:E{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                    rowIndex += 1

                    'criando linha que informa a cidade e o estado da empresa
                    worksheet.Cells(rowIndex, columnIndex).Style.Font.Bold = True
                    worksheet.Cells(rowIndex, columnIndex).Style.Font.Color.SetColor(Color.Black)
                    worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                    worksheet.Cells(rowIndex, columnIndex).Value = String.Format("{0}/{1}", objEmpresa.Cidade, objEmpresa.CodigoEstado)
                    worksheet.Cells(String.Format("A{0}:E{0}", rowIndex)).Merge = True
                    worksheet.Cells(String.Format("A{0}:E{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                    rowIndex += 1

                    'criando linha que informa o título do relatório
                    worksheet.Cells(rowIndex, columnIndex).Style.Font.Bold = True
                    worksheet.Cells(rowIndex, columnIndex).Style.Font.Color.SetColor(Color.Red)
                    worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                    worksheet.Cells(rowIndex, columnIndex).Value = String.Format("{0}", "Produtos que não foram registrados no banco.")
                    worksheet.Cells(String.Format("A{0}:E{0}", rowIndex)).Merge = True
                    worksheet.Cells(String.Format("A{0}:E{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                    rowIndex += 1

                    'criando linha que informa o período selecionado na página
                    worksheet.Cells(rowIndex, columnIndex).Style.Font.Bold = True
                    worksheet.Cells(rowIndex, columnIndex).Style.Font.Color.SetColor(Color.Black)
                    worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                    worksheet.Cells(rowIndex, columnIndex).Value = String.Format("{0}", "DATA : " & String.Format(Now(), "dd/MM/yyyy"))
                    worksheet.Cells(String.Format("A{0}:E{0}", rowIndex)).Merge = True
                    worksheet.Cells(String.Format("A{0}:E{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                    rowIndex += 1

                    'criando linha com o cabeçalho da planilha
                    For Each col As DataColumn In dsExcel.Columns
                        worksheet.Cells(rowIndex, columnIndex).Value = col.ColumnName
                        columnIndex += 1
                    Next

                    'criando auto filtro na planilha
                    worksheet.Cells("A5:H" & rowIndex).AutoFilter = True

                    'aplicando formatação nas células do cabeçalho
                    Using range = worksheet.Cells(rowIndex, 1, rowIndex, dsExcel.Columns.Count)
                        range.Style.Font.Bold = True
                        range.Style.Fill.PatternType = ExcelFillStyle.Solid
                        range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
                        range.Style.Font.Color.SetColor(Color.White)
                    End Using
                    rowIndex += 1

                    ' criando conteúdo da planilha com os dados do dataset
                    For Each row As DataRow In dsExcel.Rows
                        columnIndex = 1
                        For Each col As DataColumn In dsExcel.Columns
                            worksheet.Cells(rowIndex, columnIndex).Value = row(col.ColumnName)
                            columnIndex += 1
                        Next

                        'formatando células datas
                        'worksheet.Cells(String.Format("A{0}", rowIndex)).Style.Numberformat.Format = "dd/MM/yyyy"
                        'formatando células numéricas
                        'worksheet.Cells(String.Format("H{0}", rowIndex)).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"

                        'aplicando formatação nas células do conteúdo
                        If rowIndex Mod 2 = 0 Then
                            Using range = worksheet.Cells(rowIndex, 1, rowIndex, dsExcel.Columns.Count)
                                range.Style.Fill.PatternType = ExcelFillStyle.Solid
                                range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(153, 204, 255))
                            End Using
                        End If
                        rowIndex += 1
                    Next

                    'criando colunas de totalizadores
                    worksheet.Cells(String.Format("G{0}", rowIndex)).Style.Font.Bold = True
                    worksheet.Cells(String.Format("G{0}", rowIndex)).Style.Fill.PatternType = ExcelFillStyle.Solid
                    worksheet.Cells(String.Format("G{0}", rowIndex)).Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
                    worksheet.Cells(String.Format("G{0}", rowIndex)).Style.Font.Color.SetColor(Color.White)
                    worksheet.Cells(String.Format("G{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                    worksheet.Cells(String.Format("G{0}", rowIndex)).Value = String.Format("{0}", "TOTAL")

                    'criando colunas de totalizadores
                    worksheet.Cells(String.Format("H{0}", rowIndex)).Style.Font.Bold = True
                    worksheet.Cells(String.Format("H{0}", rowIndex)).Style.Fill.PatternType = ExcelFillStyle.Solid
                    worksheet.Cells(String.Format("H{0}", rowIndex)).Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
                    worksheet.Cells(String.Format("H{0}", rowIndex)).Style.Font.Color.SetColor(Color.White)
                    worksheet.Cells(String.Format("H{0}", rowIndex)).Formula = String.Format("=SUM(H6:H{0})", rowIndex - 1)
                    worksheet.Cells(String.Format("H{0}", rowIndex)).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"

                    'setando autofit nas células da planilha
                    worksheet.Cells.AutoFitColumns(0)

                    'congelando quinta linha (cabeçalho)
                    worksheet.View.FreezePanes(6, 1)

                    'salvando planilha do excel
                    package.Save()
                End Using
            End Using
            Funcoes.AbrirExcel(Me.Page, "PlanilhasExcel/" & Path.GetFileName(fileName))
            'download do arquivo pelo browser
            'Download(fileName)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub envairepresetnante()
    End Sub

    Protected Sub lnkAPI_Click(sender As Object, e As EventArgs) Handles lnkAPI.Click
        Try
            'Dim WsVendedor = New WsVendedor()
            'Dim listaWsVendedor = WsVendedor.ObterWsVendedor()

            'Dim teste = listaWsVendedor

        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    'Protected Sub lnkverHorario_Click(sender As Object, e As EventArgs) Handles lnkverHorario.Click
    '    Dim dataPc As DateTime = DateTime.Now()

    '    Dim easternZoneMT As TimeZoneInfo = TimeZoneInfo.FindSystemTimeZoneById("Central Brazilian Standard Time")
    '    Dim horaMT As DateTime = TimeZoneInfo.ConvertTimeFromUtc(DateTime.UtcNow, easternZoneMT)

    '    Dim easternZoneBR As TimeZoneInfo = TimeZoneInfo.FindSystemTimeZoneById("E. South America Standard Time")
    '    Dim horaBrasilia As DateTime = TimeZoneInfo.ConvertTimeFromUtc(DateTime.UtcNow, easternZoneBR)

    '    Dim horaLocal As String = dataPc.ToString("HH:mm")
    '    Dim horarioMT As String = horaMT.ToString("HH:mm")
    '    Dim horarioBrasilia As String = horaBrasilia.ToString("HH:mm")

    '    If horaLocal = horarioBrasilia Then
    '        ' lblHorario.Text = "Hora local: " & horaLocal & " - MT: " & horarioMT & " - Brasilia: " & horaBrasilia & " - HORA BRASILIA VERDADEIRO"
    '        MsgBox(Me.Page, "Hora local: " & horaLocal & " - MT: " & horarioMT & " - Brasilia: " & horaBrasilia & " - HORA BRASILIA VERDADEIRO")
    '    Else
    '        '  lblHorario.Text = "Hora local: " & horaLocal & " - MT: " & horarioMT & " - Brasilia: " & horaBrasilia & " - HORA BRASILIA FALSO"
    '        MsgBox(Me.Page, "Hora local: " & horaLocal & " - MT: " & horarioMT & " - Brasilia: " & horaBrasilia & " - HORA BRASILIA FALSO")
    '    End If

    'End Sub


    Protected Sub lnkRecontabilizaTitulo_Click(sender As Object, e As EventArgs) Handles lnkRecontabilizaTitulo.Click
        Dim objBanco As New AcessaBanco()

        SqlArray.Clear()

        Try
            'sql = "Select Registro_Id from ContasAPagar " & vbCrLf &
            '        "where EmpresaPagadora = '05366261000143'" & vbCrLf &
            '        "and Empresa  = '05366261000143'" & vbCrLf &
            '        "and Provisao = 1" & vbCrLf &
            '        "--and Situacao = 1" & vbCrLf &
            '        "and Registro_id in(297196, 297232, 297225)" & vbCrLf &
            '        "order by Baixa"

            'sql = "Select Registro_Id from ContasAPagar " & vbCrLf &
            '        "where registro_id in(400805,400893,400922,400938,400882,400887,400850,400904,400890,400925,400963,400933,400867,400914,400903,400908,400937,400926,400824,400804,400988,400871,400989,400990," & vbCrLf &
            '        "						400977,400976,400907,400795,400917,400899,400958,400883,400794,400851,400886,400929,400909,400924,400835,400811,400872,400957,400813,400866,400836,400931,400982,400823," & vbCrLf &
            '        "						400784,400953,400962,400932,400885,400918,400927,400783,400864,400809,400865,400786,400875,400858,400841,400859,400921,400975,400884,400897,400942,400846,400984,400923," & vbCrLf &
            '        "						400825,400978,400790,400940,400838,400791,400915,400820,400983,400952,399966,400973,400826,400894,400916,400837,400896,400956,400812,400967,400900,400930,400816,400819," & vbCrLf &
            '        "						400941,400895,400789,400964,400913,400971,400934,400868,400959,400818,400961,400806,400817,400928,400912,400847,400965,400845,400970,400803,400856,400972,400985,400960,400844)"

            sql = "Select Registro_Id from ContasAReceber " & vbCrLf &
                    "where carteira = ''"
            '"where registro_id in(451433,484121,484571,486969,487077,489633,489644,489651,489678,489687,489697,490468,491332,492248,492288,492312,492316,492330,492341,492366,492624,494606,495091,497074,501322,502612,502690,504179,504335,504392)"

            Dim dsTitulos As DataSet = objBanco.ConsultaDataSet(sql, "Titulos")

            If dsTitulos IsNot Nothing AndAlso dsTitulos.Tables(0).Rows.Count > 0 Then
                For Each row As DataRow In dsTitulos.Tables(0).Rows

                    Dim tit As New Titulo(row("Registro_Id"))

                    'tit.Prorrogacao = CDate("18-08-2025")
                    'tit.Baixa = CDate("18-08-2025")

                    'tit.CodigoCarteira = ""

                    'sql = "update ContasAReceber set Prorrogacao = '2025-08-18', Baixa = '2025-08-18', ObservacoesControleInterno = 'AJUSTADO BAIXA DE 15/08/2025 PARA 18/08/2025 CFE. SOLICITACAO LUCAS EM 29/08/2025 NO WHATSAPP - FURLAN 29/08/2025' where Registro_Id = " & tit.Codigo
                    sql = "update ContasAReceber set Carteira = '002005002', ObservacoesControleInterno = 'AJUSTADO CARTEIRA DE 001002007 PARA 002005002 - FURLAN 19/12/2025' where Registro_Id = " & tit.Codigo

                    SqlArray.Add(sql)

                    If tit.ValorDoDocumento > 0 AndAlso tit.CodigoProvisao = 1 Then

                        tit.DeletaContabilizacao(SqlArray)

                        tit.Contabilizar(SqlArray)

                    End If

                Next


                If SqlArray.Count > 0 Then
                    If objBanco.GravaBanco(SqlArray) Then
                        MsgBox(Me.Page, "gravado com sucesso")
                    Else
                        MsgBox(Me.Page, HttpContext.Current.Session("ssMessage").ToString)
                    End If
                End If

            Else
                MsgBox(Me.Page, "Lista não encontrada")
            End If

        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)

        End Try

    End Sub

    Protected Sub lnkLerExcelImobilizado_Click(sender As Object, e As EventArgs) Handles lnkLerExcelImobilizado.Click

        Dim caminhoArquivoExcel As String = "C:\NGS\NUTRI\ImobilizadoBAXI.xls"
        Dim nomePlanilhaExcel As String = "Planilha" & "$"

        Dim conexaoOleDb As OleDbConnection = Nothing
        Dim ds As DataSet
        Dim cmd As OleDbDataAdapter

        Dim objBanco As New AcessaBanco()

        SqlArray.Clear()

        Try
            conexaoOleDb = New OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0;Data Source= " & caminhoArquivoExcel & ";Extended Properties='Excel 8.0;HDR=Yes'")
            cmd = New OleDbDataAdapter("Select * from [" & nomePlanilhaExcel & "]", conexaoOleDb)
            cmd.TableMappings.Add("Table", "tabelaExcel")
            ds = New DataSet
            cmd.Fill(ds)

            Dim Descricao As String
            Dim Historico As String
            Dim sequencia As Integer = 68
            'ENP = 14
            'TER = 3
            'MEI = 68

            If ds IsNot Nothing AndAlso ds.Tables(0).Rows.Count > 0 Then

                For Each row As DataRow In ds.Tables(0).Rows

                    If row("GRUPO") = "MEI" Then

                        If row("CENTRODECUSTO").ToString.Length = 0 Then
                            row("CENTRODECUSTO") = "50215"
                        End If

                        Descricao = row("DESCRICAO").ToString.Replace(",", "")
                        Historico = row("HISTORICO").ToString.Replace(",", "")

                        sql = "INSERT Into Ativos(Empresa_Id, Grupo_Id, Codigo_Id, Sequencia_Id, Situacao, " & vbCrLf &
                                           "UnidadeDeNegocio, Empresa, EndEmpresa, CentroDeCusto, Conta, " & vbCrLf &
                                           "DataAquisicao, InicioDeUso, Descricao, Historico, ValorOriginal, " & vbCrLf &
                                           "Depreciar, Atualizacao, QuemLancou, QuandoLancou, Seguro) VALUES " & vbCrLf &
                                           "('" & row("EMPRESA") & "','" & row("GRUPO") & "'," & sequencia & ",0,1" & vbCrLf &
                                           ",'03','" & row("EMPRESALOCACAO") & "'," & row("ENDEMPRESA") & "," & row("CENTRODECUSTO") & ",'" & row("CONTACONTABIL") & "'" & vbCrLf &
                                           ",'" & row("AQUISICAO") & "','" & row("INICIODEUSO") & "','" & Funcoes.EliminarCaracteresEspeciais(Descricao) & "','" & Funcoes.EliminarCaracteresEspeciais(Historico) & "'," & Str(CDec(row("VALOR"))) & vbCrLf &
                                           ",'S'" & ",'" & row("ATUALIZACAO") & "','AUTOMATICO',GETDATE(),'FALSE');"

                        SqlArray.Add(sql)


                        sql = "INSERT Into AtivosXMovimentos (Empresa_Id, Grupo_Id, Codigo_Id, Sequencia_Id, Movimento_Id, Valor, Processo, Empresa, EndEmpresa) VALUES " & vbCrLf &
                                                            "('" & row("EMPRESA") & "','" & row("GRUPO") & "'," & sequencia & ",0,'" & row("INICIODEUSO") & "',0,'INICIAL','" & row("EMPRESALOCACAO") & "'," & row("ENDEMPRESA") & ");"

                        SqlArray.Add(sql)

                        sequencia += 1

                    End If
                Next

                If SqlArray.Count > 0 Then
                    If objBanco.GravaBanco(SqlArray) Then
                        MsgBox(Me.Page, "gravado com sucesso")
                    Else
                        MsgBox(Me.Page, HttpContext.Current.Session("ssMessage").ToString)
                    End If
                End If

            Else
                MsgBox(Me.Page, "Lista não encontrada")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try

    End Sub


    Protected Sub lnkLerExcelEncargos_Click(sender As Object, e As EventArgs) Handles lnkLerExcelEncargos.Click

        Dim strm As New StreamWriter("C:\NGS\RTGRAOS\ICMS2025.txt")
        Dim caminhoArquivoExcel As String = "C:\NGS\RTGRAOS\ICMS2025.xls"
        Dim nomePlanilhaExcel As String = "Planilha" & "$"

        Dim conexaoOleDb As OleDbConnection = Nothing
        Dim ds As DataSet
        Dim cmd As OleDbDataAdapter

        Try
            conexaoOleDb = New OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0;Data Source= " & caminhoArquivoExcel & ";Extended Properties='Excel 8.0;HDR=Yes'")
            cmd = New OleDbDataAdapter("Select * from [" & nomePlanilhaExcel & "]", conexaoOleDb)
            cmd.TableMappings.Add("Table", "tabelaExcel")
            ds = New DataSet
            cmd.Fill(ds)

            If ds IsNot Nothing AndAlso ds.Tables(0).Rows.Count > 0 Then

                For Each row As DataRow In ds.Tables(0).Rows

                    SqlArray.Clear()

                    Dim Parametros As New OperacaoXEstado
                    Parametros.Codigo = row("ID")

                    ObjOxE = New OperacaoXEstado(Parametros)

                    Dim TeveEncargo As Boolean = False

                    If ObjOxE.Encargos.Count > 0 Then

                        For Each enc As OperacaoXEstadoXEncargo In ObjOxE.Encargos

                            If enc.CodigoEncargo = "ICMS" OrElse enc.CodigoEncargo = "ICMS RET SIMB" Then
                                ObjOxE.IUD = "I"

                                ObjOxE.UsuarioInclusao = "AUTOMATICO"
                                ObjOxE.InicioVigencia = CDate("01/01/2025")
                                ObjOxE.Ativo = True

                                enc.AliquotaBase = row("PercentualBaseNovo")
                                enc.Aliquota = row("AliquotaNova")
                                enc.AliquotaExibicao = row("AliquotaNova")

                                TeveEncargo = True
                            End If

                        Next

                    End If

                    'If ObjOxE.Codigo = 119638 OrElse ObjOxE.Codigo = 119639 OrElse ObjOxE.Codigo = 119640 Then
                    'Else
                    '    TeveEncargo = False
                    'End If

                    If TeveEncargo Then
                        ObjOxE.SalvarSql(SqlArray)

                        If Banco.GravaBanco(SqlArray) Then

                        Else
                            MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                            Exit Sub
                        End If
                    Else
                        strm.WriteLine(row("ID"))
                    End If
                Next

                strm.Close()

                MsgBox(Me.Page, "Processo finalizado")

            Else
                MsgBox(Me.Page, "Lista não encontrada")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try

    End Sub

    Protected Sub lnkTransferirCotacoes_Click(sender As Object, e As EventArgs) Handles lnkTransferirCotacoes.Click
        Dim objBanco As New AcessaBanco()

        SqlArray.Clear()

        Try

            sql = "Select Indexador_Id, Data_Id, Indice, Realizado, UsuarioLiberacao, UsuarioLiberacaoData FROM Cotacoes " & vbCrLf &
                    "where Data_Id >= '2026-01-01'"

            Dim ds As DataSet = objBanco.ConsultaDataSet(sql, "Cotacoes")

            If ds IsNot Nothing AndAlso ds.Tables(0).Rows.Count > 0 Then

                For Each row As DataRow In ds.Tables(0).Rows

                    sql = "Insert into Cotacoes (Indexador_Id, Data_Id, Indice, Realizado, UsuarioLiberacao, UsuarioLiberacaoData) " & vbCrLf &
                        " values (" & row("Indexador_Id") & ",'" & CDate(row("Data_Id")).ToString("yyyy-MM-dd") & "'," & Str(row("Indice")) & ",'" & row("Realizado") & "','" & row("UsuarioLiberacao") & "','" & CDate(row("UsuarioLiberacaoData")).ToString("yyyy-MM-dd") & "');"

                    SqlArray.Add(sql)

                Next

                If SqlArray.Count > 0 Then
                    If GravaBancoDestino(SqlArray, sConnStringDestino) Then
                        MsgBox(Me.Page, "gravado com sucesso")
                    Else
                        MsgBox(Me.Page, HttpContext.Current.Session("ssMessage").ToString)
                    End If
                End If

            Else
                MsgBox(Me.Page, "Lista não encontrada")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub


    Protected Sub lnkTransferirEncargos_Click(sender As Object, e As EventArgs) Handles lnkTransferirEncargos.Click
        Dim objBanco As New AcessaBanco()

        Try

            sql = "Select Codigo_Id from OperacaoXEstado " & vbCrLf &
                    "where Empresa = '53267147'" & vbCrLf &
                    "and CODIGO_ID IN(3,4,15,32,33,52,55,88,94,96,154,203,290,442,476,506,508,649,650,651,652,653,654,656,657,658,721,744,759,760,823,960,1431,1618,1664,1906,1920,1928,1982,2010,2093,2163,2164,2277)"

            Dim ds As DataSet = objBanco.ConsultaDataSet(sql, "OperacaoXEstado")

            If ds IsNot Nothing AndAlso ds.Tables(0).Rows.Count > 0 Then

                For Each row As DataRow In ds.Tables(0).Rows

                    SqlArray.Clear()

                    Dim Parametros As New OperacaoXEstado
                    Parametros.Codigo = row("Codigo_Id")

                    ObjOxE = New OperacaoXEstado(Parametros)

                    ObjOxE.IUD = "I"
                    ObjOxE.Empresa = "53267147"
                    ObjOxE.UsuarioInclusao = "AUTOMATICO"
                    ObjOxE.InicioVigencia = CDate("01/04/2026")
                    ObjOxE.Ativo = True

                    For Each enc In ObjOxE.Encargos
                        If enc.CodigoEncargo = "FUNRURAL" Then
                            enc.Aliquota = 1.43
                            enc.AliquotaExibicao = 0
                        End If
                    Next

                    If ObjOxE.Encargos.Count > 0 Then

                        ObjOxE.SalvarSql(SqlArray)

                        If Not GravaBancoDestino(SqlArray, sConnStringDestino) Then
                            MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                            Exit Sub
                        End If
                    End If
                Next

                MsgBox(Me.Page, "Processo finalizado")

            Else
                MsgBox(Me.Page, "Lista não encontrada")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try

    End Sub


    Protected Sub lnkLerIBSCBS_Click(sender As Object, e As EventArgs) Handles lnkLerIBSCBS.Click

        SqlArray.Clear()

        'Dim strm As New StreamWriter(HttpContext.Current.Server.MapPath("~/Files/gravaID.txt"))


        Dim caminhoArquivoExcel As String = "C:\NGS\Nutri\IBSCBS\IBS-CBS-NUTRI.xlsx"
        Dim nomePlanilhaExcel As String = "IBS-CBS-POR-ID" & "$"

        Dim ds As DataSet
        ds = FuncoesStrings.LerExcelParaDataSet(caminhoArquivoExcel)

        Dim ObjOxE As [Lib].Negocio.OperacaoXEstado

        Dim CODIGO As String
        Dim CSTIBS As Integer
        Dim CLACBS As Integer

        Dim ALIQUOTAIBS As String

        Dim PERCENTUAL As String

        Dim ReducaoIBS As String

        Dim comAliquota As Boolean


        If ds IsNot Nothing AndAlso ds.Tables(0).Rows.Count > 0 Then

            For Each row As DataRow In ds.Tables(0).Rows

                CODIGO = row("OPXENCARGO")

                'strm.WriteLine(CODIGO & ",")

                CSTIBS = CInt(row("CST"))

                CLACBS = CInt(row("CCLASSTRIB"))

                If row("ALIQUOTA IBS") = "PADRÃO" Then
                    comAliquota = True
                Else
                    comAliquota = False
                End If

                If row("PERCENTUAL DE REDUÇÃO") = "60%" Then
                    ReducaoIBS = "60"
                Else
                    ReducaoIBS = "0"
                End If


                sql = "update OperacaoXEstado set  STIBSCBS = " & CSTIBS & ", ClassificacaoIBSCBS = " & CLACBS & vbCrLf

                If ReducaoIBS = "60" Then
                    sql &= ", ReducaoIBS_Perc = 60, ReducaoCBS_Perc = 60 " & vbCrLf
                End If

                sql &= "where Codigo_Id = " & CODIGO

                SqlArray.Add(sql)

                If Not comAliquota Then
                    sql = "update OperacaoXEstadoXEncargo set  AliquotaBase = 0, Aliquota = 0, AliquotaExibicao = 0, AliquotaLimite = 0 " & vbCrLf
                    sql &= "where Codigo_Id = " & CODIGO & " AND Encargo_Id = 'CBS'"

                    SqlArray.Add(sql)

                    sql = "update OperacaoXEstadoXEncargo set  AliquotaBase = 0, Aliquota = 0, AliquotaExibicao = 0, AliquotaLimite = 0 " & vbCrLf
                    sql &= "where Codigo_Id = " & CODIGO & " AND Encargo_Id = 'IBS'"

                    SqlArray.Add(sql)
                End If

            Next

            'strm.Close()
            'strm.Dispose()

            If SqlArray.Count > 0 Then
                If Banco.GravaBanco(SqlArray) Then
                    MsgBox(Me.Page, "gravado com sucesso")
                Else
                    MsgBox(Me.Page, HttpContext.Current.Session("ssMessage").ToString)
                End If
            End If

        End If

    End Sub


End Class