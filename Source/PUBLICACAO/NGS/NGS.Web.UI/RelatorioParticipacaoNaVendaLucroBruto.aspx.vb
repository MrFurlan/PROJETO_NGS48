Imports System.Data
Imports System.IO
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class RelatorioParticipacaoNaVendaLucroBruto
    Inherits BasePage

    Private dateEmissao As New DateTime
    Private intPagina As Integer = 0

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Expedicao)
        If Not IsPostBack And IsConnect Then
            If Funcoes.VerificaPermissao("ExtratoDePedido", "ACESSAR") Then
                BuscaEmpresa()
                CarregarSafras()
                Limpar()
                ddl.Carregar(ddlMarca, CarregarDDL.Tabela.Marca, "")
                HID.Value = Guid.NewGuid().ToString
                ucConsultaClientes.SetarHID(HID.Value)
            Else
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Expedicao.aspx")
                Exit Sub
            End If
        End If
    End Sub

    Public Overrides Sub Carregar(obj As IBaseEntity)
        If Not Session("objClienteCarCliRepre" & HID.Value) Is Nothing Then
            Dim itemCliente As ListItem = Funcoes.FormatarListItemCliente(CType(Session("objClienteCarCliRepre" & HID.Value), [Lib].Negocio.Cliente))
            txtClienteRepresentante.Text = itemCliente.Text
            txtCodigoCliRepre.Value = itemCliente.Value
            Session.Remove("objClienteCarCliRepre" & HID.Value)
        End If
    End Sub

    Private Sub Limpar()
        txtClienteRepresentante.Text = ""
        txtCodigoCliRepre.Value = ""
        ddlSafra.SelectedIndex = 0
        cmbEmpresa.SelectedIndex = 0
    End Sub

    Private Sub BuscaEmpresa()
        ddl.Carregar(cmbEmpresa, CarregarDDL.Tabela.ClientesXEmpresas, "", True)
    End Sub

    Private Sub CarregarSafras()
        ddl.Carregar(ddlSafra, CarregarDDL.Tabela.Safra, "", True)
    End Sub

    Private Function ValidarCampos() As Boolean
        If cmbEmpresa.SelectedIndex = 0 Then
            MsgBox(Me.Page, "Empresa não foi selecionada")
            Return False
        End If
        Return True
    End Function

    Public Sub BuscarRegistros()
        Dim sqlProd As String = ""
        Dim DescricaoProduto As String = ""
        Dim dsProduto As New DataSet

        Dim ds As New DataSet
        Dim SQL As String
        Dim Parametros As String = "Parametros:" & vbCrLf
        Dim Cliente As String = ""
        Dim strEmpresa() As String = cmbEmpresa.SelectedValue.Split("-")
        Dim strClienteRepresentante() As String = txtCodigoCliRepre.Value.Split("-")
        Dim RetornoProdutos As ArrayList

        If Not rbPorRepresentante.Checked Then

            SQL = " Declare " & vbCrLf & _
                  "  @ValorTotalVenda  numeric(18,2), " & vbCrLf & _
                  "  @ValorTotalLucro  numeric(18,2) " & vbCrLf & _
                  "                                          " & vbCrLf & _
                  "  select P.Cliente,  " & vbCrLf & _
                  "          P.EndCliente, " & vbCrLf & _
                  "          Cli.Nome, " & vbCrLf & _
                  "          p.Pedido_Id, " & vbCrLf & _
                  "          PxI.Produto_Id, " & vbCrLf & _
                  "          PxI.Descricao, " & vbCrLf & _
                  "          max(isnull(PxI.UnitarioOficialCompra,0)) as UnitarioOficialCompra,  " & vbCrLf & _
                  "          max(isnull(PxI.UnitarioMoedaCompra,0)) as UnitarioMoedaCompra,   " & vbCrLf & _
                  "   		Sum(case      " & vbCrLf & _
                  "   			  When PxI.TipoDeLancamento = 'E'   " & vbCrLf & _
                  "   			    then PxI.Quantidade * - 1    " & vbCrLf & _
                  "   			    else PxI.Quantidade     " & vbCrLf & _
                  "   			end) as Quantidade,     " & vbCrLf & _
                  "   		Sum(case      " & vbCrLf & _
                  "   		  	  When PxI.TipoDeLancamento = 'E'  " & vbCrLf & _
                  "   			    then PxI.TotalOficial * - 1    " & vbCrLf & _
                  "   			    else PxI.TotalOficial     " & vbCrLf & _
                  "   			end) as TotalOficial,    " & vbCrLf & _
                  "   		Sum(case      " & vbCrLf & _
                  "   		      When PxI.TipoDeLancamento = 'E'  " & vbCrLf & _
                  "   			    then PxI.TotalMoeda * - 1     " & vbCrLf & _
                  "   				else PxI.TotalMoeda     " & vbCrLf & _
                  "   		    end) as TotalMoeda   " & vbCrLf & _
                  "   	into #PreTemp     " & vbCrLf & _
                  "   	FROM Pedidos P     " & vbCrLf & _
                  "      Inner Join Suboperacoes SO   " & vbCrLf & _
                  "         on SO.Operacao_Id     = P.Operacao  " & vbCrLf & _
                  "        and SO.Suboperacoes_Id = P.Suboperacao   " & vbCrLf & _
                  "      Inner Join PedidoXItemXLancamento PxI     " & vbCrLf & _
                  "         on P.Empresa_Id    = PxI.Empresa_Id  " & vbCrLf & _
                  "        and P.EndEmpresa_Id = PXI.EndEmpresa_Id    " & vbCrLf & _
                  "        and P.Pedido_Id     = PxI.Pedido_Id     " & vbCrLf & _
                  "      INNER JOIN   Produtos Prod   " & vbCrLf & _
                  "         ON PxI.Produto_Id = Prod.Produto_Id  " & vbCrLf & _
                  "      INNER JOIN Clientes Cli   " & vbCrLf & _
                  "         ON P.Cliente = Cli.Cliente_Id   " & vbCrLf & _
                  "        AND P.EndCliente = Cli.Endereco_Id   " & vbCrLf & _
                  "          WHERE P.Situacao = 1  " & vbCrLf

            If strEmpresa(0).Length > 0 Then
                Parametros &= "Empresa: " & cmbEmpresa.SelectedItem.Text & vbCrLf
                If Not ckConsolidarEmp.Checked Then
                    Parametros &= "Empresa Consolidada" & vbCrLf
                    SQL &= "  AND P.Empresa_id     = '" & strEmpresa(0) & "'" & vbCrLf
                    SQL &= "  AND P.EndEmpresa_Id  = " & strEmpresa(1) & " " & vbCrLf
                Else
                    SQL &= "  AND left(P.Empresa_id,8)  = '" & Left(strEmpresa(0), 8) & "'" & vbCrLf
                    SQL &= "  AND left(P.EndEmpresa_Id,8)  =  " & Left(strEmpresa(1), 8) & " " & vbCrLf
                End If
            End If

            If strClienteRepresentante(0).Length > 0 Then
                Parametros &= "Cliente: " & txtClienteRepresentante.Text & vbCrLf
                SQL &= " AND P.Cliente = '" & strClienteRepresentante(0) & "'" & vbCrLf
                SQL &= " AND P.EndCliente = " & strClienteRepresentante(1) & " " & vbCrLf
            End If

            If ddlSafra.SelectedIndex > 0 Then
                SQL &= " AND P.Safra   = '" & ddlSafra.SelectedValue & "'" & vbCrLf
                Parametros &= "Safra: " & ddlSafra.SelectedItem.Text & vbCrLf
            End If

            If ddlMarca.SelectedIndex > 0 Then
                SQL &= " AND isnull(Prod.Marca,0) =" & ddlMarca.SelectedValue & vbCrLf
                Parametros &= "Marca: " & ddlMarca.SelectedItem.Text & vbCrLf
            End If

            If ucSelecaoProduto.TemSelecionado Then
                RetornoProdutos = ucSelecaoProduto.GetSqlEParametrosRelatorio("Produtos.Grupo", "Produtos.Produto_Id", "")
                Parametros &= " AND " & RetornoProdutos(0)
                Parametros &= RetornoProdutos(1)
            End If

            If ckCPeriodo.Checked = True Then
                SQL &= " AND  (P.DataPedido BETWEEN '" & txtDataInicial.Text.ToSqlDate() & "' AND '" & txtDataFinal.Text.ToSqlDate() & "')  " & vbCrLf
                Parametros &= "Periodo de " & txtDataInicial.Text.ToSqlDate() & " a " & txtDataFinal.Text.ToSqlDate() & vbCrLf
            End If

            SQL &= "        AND So.Entradasaida  = 'S'   " & vbCrLf & _
                   "        AND So.Classe     ('" & eClassesOperacoes.VENDAS.ToString & "','" & eClassesOperacoes.REMESSAS.ToString & "','" & eClassesOperacoes.EXPORTACOES.ToString & "','" & eClassesOperacoes.VENDASAORDEM.ToString & "','" & eClassesOperacoes.GLOBAL.ToString & "')" & vbCrLf & _
                   "      Group by P.Cliente,     " & vbCrLf & _
                   "          P.EndCliente, " & vbCrLf & _
                   "          Cli.Nome, " & vbCrLf & _
                   "          p.Pedido_Id, " & vbCrLf & _
                   "          PxI.Produto_Id, " & vbCrLf & _
                   "          PxI.Descricao" & vbCrLf & _
                   "   having Sum(case      " & vbCrLf & _
                   "                When PxI.TipoDeLancamento = 'E'  " & vbCrLf & _
                   "  			      then PxI.Quantidade * - 1     " & vbCrLf & _
                   "  			      else PxI.Quantidade      " & vbCrLf & _
                   "  		      end) > 0   " & vbCrLf & _
                   "                                                         " & vbCrLf & _
                   "   Select Cliente,      " & vbCrLf & _
                   "          EndCliente, " & vbCrLf & _
                   "          Nome, " & vbCrLf & _
                   "          Produto_Id, " & vbCrLf & _
                   "          Descricao, " & vbCrLf & _
                   "          Sum(TotalOficial) as TotalVenda, " & vbCrLf & _
                   "          Sum(TotalOficial) - sum(UnitarioOficialCompra * Quantidade) as LucroBruto   " & vbCrLf & _
                   "      into #Temp  " & vbCrLf & _
                   "      from #PreTemp   " & vbCrLf & _
                   "     group By Cliente, EndCliente, Nome, Produto_Id, Descricao; " & vbCrLf & _
                   "                                                            " & vbCrLf & _
                   "  (select @ValorTotalLucro = sum(LucroBruto) from #temp) " & vbCrLf & _
                   "  (select @ValorTotalVenda = sum(TotalVenda) from #temp) " & vbCrLf

            If rbPorCliente.Checked Then
                '*** SQL por cliente-----------------------------------------------------------------------------------------------------------------------------
                SQL &= " select " & IIf(ckConsolidarCli.Checked, "left(cliente,8)", "cliente") & " AS cliente,  " & vbCrLf & _
                        "       " & IIf(ckConsolidarCli.Checked, " left(endcliente,8)", "endcliente") & " AS endcliente, " & vbCrLf & _
                        "       nome, " & vbCrLf & _
                        "       sum(TotalVenda) as TotalClienteVenda, " & vbCrLf & _
                        "       sum(LucroBruto) as TotalClienteLucro, " & vbCrLf & _
                        "       sum(LucroBruto)*100/@ValorTotalLucro as ParticipacaoLucro, " & vbCrLf & _
                        "       sum(TotalVenda)*100/@ValorTotalVenda as ParticipacaoVenda " & vbCrLf & _
                        "  from #temp " & vbCrLf & _
                        " group by " & IIf(ckConsolidarCli.Checked, "left(cliente,8)", "cliente") & ",  " & vbCrLf & _
                        "          " & IIf(ckConsolidarCli.Checked, " left(endcliente,8)", "endcliente") & ", " & vbCrLf & _
                        "          Nome " & vbCrLf


            ElseIf rbPorProduto.Checked Then
                '*** SQL por produto-----------------------------------------------------------------------------------------------------------------------------
                SQL &= " select Produto_Id, " & vbCrLf & _
                       "        Descricao, " & vbCrLf & _
                       "        sum(TotalVenda) as TotalProdutoVenda, " & vbCrLf & _
                       "        sum(LucroBruto) as TotalProdutoLucro, " & vbCrLf & _
                       "        sum(LucroBruto)*100/@ValorTotalLucro as ParticipacaoLucro, " & vbCrLf & _
                       "        sum(TotalVenda)*100/@ValorTotalVenda as ParticipacaoVenda " & vbCrLf & _
                       "   from #temp " & vbCrLf & _
                       "  group by Produto_Id,Descricao " & vbCrLf

            End If

            If rbVenda.Checked Then
                SQL &= "   Order By sum(TotalVenda)*100/@ValorTotalVenda desc;" & vbCrLf
            ElseIf rbLucro.Checked Then
                SQL &= "   order By sum(LucroBruto)*100/@ValorTotalLucro desc; " & vbCrLf
            End If

            SQL &= "  drop table #PreTemp " & vbCrLf & _
                   "  drop table #temp " & vbCrLf


        Else
            '*** SQL por representante-----------------------------------------------------------------------------------------------------------------------------
            SQL = " declare " & vbCrLf & _
                  " @ValorTotalVenda numeric(18,2), " & vbCrLf & _
                  " @ValorTotalLucro numeric(18,2) " & vbCrLf & _
                  "                                         " & vbCrLf & _
                  " select  PTC.Representante_Id, " & vbCrLf & _
                  "         PTC.EndRepresentante_Id,    " & vbCrLf & _
                  "         p.Pedido_Id,  " & vbCrLf & _
                  "         PxI.Produto_Id, " & vbCrLf & _
                  "         PxI.Descricao, " & vbCrLf & _
                  "         max(isnull(PxI.UnitarioOficialCompra,0)) as UnitarioOficialCompra,     " & vbCrLf & _
                  "  		Sum(case      " & vbCrLf & _
                  "  			  When PxI.TipoDeLancamento = 'E'  " & vbCrLf & _
                  "  			    then PxI.Quantidade * - 1   " & vbCrLf & _
                  "  			    else PxI.Quantidade      " & vbCrLf & _
                  "  			end) as Quantidade,     " & vbCrLf & _
                  "  		Sum(case      " & vbCrLf & _
                  "  		  	  When PxI.TipoDeLancamento = 'E'    " & vbCrLf & _
                  "  			    then PxI.TotalOficial * - 1     " & vbCrLf & _
                  "  			    else PxI.TotalOficial     " & vbCrLf & _
                  "  			end) as TotalOficial " & vbCrLf & _
                  "  	into #PreTemp     " & vbCrLf & _
                  "  	FROM Pedidos P     " & vbCrLf & _
                  "     Inner Join Suboperacoes SO  " & vbCrLf & _
                  "        on SO.Operacao_Id     = P.Operacao   " & vbCrLf & _
                  "       and SO.Suboperacoes_Id = P.Suboperacao   " & vbCrLf & _
                  "     Inner Join PedidoXItemXLancamento PxI     " & vbCrLf & _
                  "        on P.Empresa_Id    = PxI.Empresa_Id     " & vbCrLf & _
                  "       and P.EndEmpresa_Id = PXI.EndEmpresa_Id     " & vbCrLf & _
                  "       and P.Pedido_Id     = PxI.Pedido_Id     " & vbCrLf & _
                  "     INNER JOIN   Produtos Prod   " & vbCrLf & _
                  "        ON PxI.Produto_Id = Prod.Produto_Id   " & vbCrLf & _
                  "     INNER JOIN Clientes Cli   " & vbCrLf & _
                  "        ON P.Cliente    = Cli.Cliente_Id   " & vbCrLf & _
                  "       AND P.EndCliente = Cli.Endereco_Id   " & vbCrLf & _
                  "     INNER Join PedidoxTabelaDeComissao PTC " & vbCrLf & _
                  "        ON P.Empresa_Id    = PTC.Empresa_Id " & vbCrLf & _
                  "       And P.EndEmpresa_Id = PTC.EndEmpresa_Id " & vbCrLf & _
                  "       And P.Pedido_Id     = PTC.Pedido_Id " & vbCrLf & _
                  "       AND PxI.Produto_Id    = PTC.Produto_Id " & vbCrLf & _
                  "     WHERE  P.Situacao      = 1  " & vbCrLf

            If strEmpresa(0).Length > 0 Then
                Parametros &= "Empresa: " & cmbEmpresa.SelectedItem.Text & vbCrLf
                If Not ckConsolidarEmp.Checked Then
                    SQL &= "  AND P.Empresa_id     = '" & strEmpresa(0) & "'" & vbCrLf
                    SQL &= "  AND P.EndEmpresa_Id  = " & strEmpresa(1) & " " & vbCrLf
                Else
                    Parametros &= "Empresa Consolidada" & vbCrLf
                    SQL &= "  AND left(P.Empresa_id,8)  = '" & Left(strEmpresa(0), 8) & "'" & vbCrLf
                    SQL &= "  AND left(P.EndEmpresa_Id,8)  =  " & Left(strEmpresa(1), 8) & " " & vbCrLf
                End If
            End If

            If strClienteRepresentante(0).Length > 0 Then
                Parametros &= "Cliente: " & txtClienteRepresentante.Text & vbCrLf
                SQL &= " AND PTC.Representante_Id = '" & strClienteRepresentante(0) & "'" & vbCrLf
                SQL &= " AND PTC.EndRepresentante_Id = " & strClienteRepresentante(1) & " " & vbCrLf
            End If

            If ddlSafra.SelectedIndex > 0 Then
                SQL &= " AND P.Safra   = '" & ddlSafra.SelectedValue & "'" & vbCrLf
                Parametros &= "Safra: " & ddlSafra.SelectedItem.Text & vbCrLf
            End If

            If ddlMarca.SelectedIndex > 0 Then
                SQL &= " AND isnull(Prod.Marca,0) =" & ddlMarca.SelectedValue & vbCrLf
                Parametros &= "Marca: " & ddlMarca.SelectedItem.Text & vbCrLf
            End If

            If ucSelecaoProduto.TemSelecionado Then
                RetornoProdutos = ucSelecaoProduto.GetSqlEParametrosRelatorio("Produtos.Grupo", "Produtos.Produto_Id", "")
                Parametros &= " AND " & RetornoProdutos(0)
                Parametros &= RetornoProdutos(1)
            End If

            If ckCPeriodo.Checked = True Then
                SQL &= " AND  (P.DataPedido BETWEEN '" & txtDataInicial.Text.ToSqlDate() & "' AND '" & txtDataFinal.Text.ToSqlDate() & "')  " & vbCrLf
                Parametros &= "Periodo de " & txtDataInicial.Text.ToSqlDate() & " a " & txtDataFinal.Text.ToSqlDate() & vbCrLf
            End If

            SQL &= "       AND So.Entradasaida  = 'S'   " & vbCrLf & _
                   "       AND So.Classe       in ('" & eClassesOperacoes.VENDAS.ToString & "','" & eClassesOperacoes.REMESSAS.ToString & "','" & eClassesOperacoes.EXPORTACOES.ToString & "','" & eClassesOperacoes.VENDASAORDEM.ToString & "','" & eClassesOperacoes.GLOBAL.ToString & "')" & vbCrLf & _
                   "     Group by PTC.Representante_Id, " & vbCrLf & _
                   "             PTC.EndRepresentante_Id, " & vbCrLf & _
                   "             p.Pedido_Id, " & vbCrLf & _
                   "             PxI.Produto_Id, " & vbCrLf & _
                   "             PxI.Descricao" & vbCrLf & _
                   "    having Sum(case      " & vbCrLf & _
                   "   		        When PxI.TipoDeLancamento = 'E'  " & vbCrLf & _
                   " 			      then PxI.Quantidade * - 1    " & vbCrLf & _
                   " 			      else PxI.Quantidade      " & vbCrLf & _
                   " 		      end) > 0   " & vbCrLf & _
                   "                                                      " & vbCrLf & _
                   "  Select Representante_Id, " & vbCrLf & _
                   "         EndRepresentante_Id,  " & vbCrLf & _
                   "         sum(TotalOficial) as TotalVenda,  " & vbCrLf & _
                   "         Sum(TotalOficial) - sum(UnitarioOficialCompra * Quantidade) as LucroBruto   " & vbCrLf & _
                   "     into #Temp  " & vbCrLf & _
                   "     from #PreTemp   " & vbCrLf & _
                   "    group By Representante_Id,  EndRepresentante_Id; " & vbCrLf & _
                   "                                                           " & vbCrLf & _
                   " (select @ValorTotalVenda = sum(TotalVenda) from #temp); " & vbCrLf & _
                   " (select @ValorTotalLucro = sum(LucroBruto) from #temp); " & vbCrLf & _
                   "                                                           " & vbCrLf & _
                   " select " & IIf(ckConsolidarCli.Checked, "left(T.Representante_Id,8)", "T.Representante_Id") & " AS Representante_Id,   " & vbCrLf & _
                   "        " & IIf(ckConsolidarCli.Checked, " left(T.EndRepresentante_Id,8)", "T.EndRepresentante_Id") & " AS EndRepresentante_Id, " & vbCrLf & _
                   "        C.Nome, " & vbCrLf & _
                   "        sum(T.TotalVenda) as TotalRepresentanteVenda, " & vbCrLf & _
                   "        sum(T.LucroBruto) as TotalRepresentanteLucro, " & vbCrLf & _
                   "        sum(T.LucroBruto)*100/@ValorTotalLucro as ParticipacaoLucro, " & vbCrLf & _
                   "        sum(T.TotalVenda)*100/@ValorTotalVenda as ParticipacaoVenda " & vbCrLf & _
                   "   from #temp T " & vbCrLf & _
                   "  Inner Join Clientes C " & vbCrLf & _
                   "     on T.Representante_Id    = C.Cliente_Id " & vbCrLf & _
                   "    and T.EndRepresentante_Id = C.Endereco_Id  " & vbCrLf & _
                   "  Group by " & IIf(ckConsolidarCli.Checked, "left(T.Representante_Id,8)", "T.Representante_Id") & ",  " & vbCrLf & _
                   "           " & IIf(ckConsolidarCli.Checked, " left(T.EndRepresentante_Id,8)", "T.EndRepresentante_Id") & ", " & vbCrLf & _
                   "            C.Nome " & vbCrLf

            If rbVenda.Checked Then
                SQL &= "   Order By sum(T.TotalVenda)*100/@ValorTotalVenda desc;" & vbCrLf
            ElseIf rbLucro.Checked Then
                SQL &= "   order By sum(T.LucroBruto)*100/@ValorTotalLucro desc; " & vbCrLf
            End If

            SQL &= " drop table #PreTemp; " & vbCrLf & _
                   " drop table #temp; " & vbCrLf
        End If


        ds = Banco.ConsultaDataSet(SQL, "RelatorioParticipacaoNaVendaLucroBruto")
        AlimentaCrptRelatorios(ds, IIf(Not rbPorRepresentante.Checked, IIf(rbPorCliente.Checked, "~/Reports/Cr_RelPartNaVendaLucroBrutoCli", "~/Reports/Cr_RelPartNaVendaLucroBrutoProd"), "~/Reports/Cr_RelPartNaVendaLucroBrutoRep"), Parametros)

    End Sub

    Public Sub AlimentaCrptRelatorios(ByVal Ds As DataSet, ByVal Caminho As String, ByVal Parametros As String)
        Dim crptRelatorio As New ReportDocument()

        Try
            crptRelatorio.FileName = Server.MapPath(Caminho & ".rpt")
            crptRelatorio.Load(crptRelatorio.FileName, CrystalDecisions.Shared.OpenReportMethod.OpenReportByDefault)

            Dim NomeArquivo2 As String = "Files/" & Funcoes.GeraNomeArquivo & ".PDF"
            Dim NomeArquivo As String = Server.MapPath(NomeArquivo2)
            Dim arquivo As String = NomeArquivo

            crptRelatorio.SetDataSource(Ds)

            Dim crParameterValues As CrystalDecisions.Shared.ParameterValues
            Dim crParameterDiscreteValue As CrystalDecisions.Shared.ParameterDiscreteValue
            Dim crParameterFieldDefinitions As CrystalDecisions.CrystalReports.Engine.ParameterFieldDefinitions
            Dim crParameterFieldDefinition As CrystalDecisions.CrystalReports.Engine.ParameterFieldDefinition

            crParameterFieldDefinitions = crptRelatorio.DataDefinition.ParameterFields()

            crParameterFieldDefinition = crParameterFieldDefinitions.Item("Parametros")
            crParameterValues = crParameterFieldDefinition.CurrentValues
            crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
            crParameterDiscreteValue.Value = Parametros
            crParameterValues.Add(crParameterDiscreteValue)
            crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

            crParameterFieldDefinition = crParameterFieldDefinitions.Item("Nome")
            crParameterValues = crParameterFieldDefinition.CurrentValues
            crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
            crParameterDiscreteValue.Value = HttpContext.Current.Session("ssNomeEmpresa")
            crParameterValues.Add(crParameterDiscreteValue)
            crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

            crParameterFieldDefinition = crParameterFieldDefinitions.Item("Cidade")
            crParameterValues = crParameterFieldDefinition.CurrentValues
            crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
            crParameterDiscreteValue.Value = HttpContext.Current.Session("ssCidadeEmpresa") & " - " & HttpContext.Current.Session("ssEstadoEmpresa")
            crParameterValues.Add(crParameterDiscreteValue)
            crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

            crParameterFieldDefinition = crParameterFieldDefinitions.Item("Zerado")
            crParameterValues = crParameterFieldDefinition.CurrentValues
            crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
            crParameterDiscreteValue.Value = "True"
            crParameterValues.Add(crParameterDiscreteValue)
            crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

            crParameterFieldDefinition = crParameterFieldDefinitions.Item("Titulo")
            crParameterValues = crParameterFieldDefinition.CurrentValues
            crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
            crParameterDiscreteValue.Value = "Relatório de Participação na Venda Lucro Bruto - " & IIf(Not rbPorRepresentante.Checked, IIf(rbPorCliente.Checked, "Por Cliente", "Por Produto"), "Por Representante") & " "
            crParameterValues.Add(crParameterDiscreteValue)
            crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

            crParameterFieldDefinition = crParameterFieldDefinitions.Item("Pagina")
            crParameterValues = crParameterFieldDefinition.CurrentValues
            crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
            crParameterDiscreteValue.Value = "Página"
            crParameterValues.Add(crParameterDiscreteValue)
            crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

            crParameterFieldDefinition = crParameterFieldDefinitions.Item("Data")
            crParameterValues = crParameterFieldDefinition.CurrentValues
            crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
            crParameterDiscreteValue.Value = "Emissão"
            crParameterValues.Add(crParameterDiscreteValue)
            crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

            crParameterFieldDefinition = crParameterFieldDefinitions.Item("DataOrigem")
            crParameterValues = crParameterFieldDefinition.CurrentValues
            crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
            crParameterDiscreteValue.Value = ""
            crParameterValues.Add(crParameterDiscreteValue)
            crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

            crParameterFieldDefinition = crParameterFieldDefinitions.Item("NomeProdutor")
            crParameterValues = crParameterFieldDefinition.CurrentValues
            crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
            crParameterDiscreteValue.Value = "Descrição"
            crParameterValues.Add(crParameterDiscreteValue)
            crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

            crParameterFieldDefinition = crParameterFieldDefinitions.Item("Complemento")
            crParameterValues = crParameterFieldDefinition.CurrentValues
            crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
            crParameterDiscreteValue.Value = ""
            crParameterValues.Add(crParameterDiscreteValue)
            crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

            crParameterFieldDefinition = crParameterFieldDefinitions.Item("NomeProduto")
            crParameterValues = crParameterFieldDefinition.CurrentValues
            crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
            crParameterDiscreteValue.Value = "T.Prod.Venda"
            crParameterValues.Add(crParameterDiscreteValue)
            crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

            crParameterFieldDefinition = crParameterFieldDefinitions.Item("NotaOrigem")
            crParameterValues = crParameterFieldDefinition.CurrentValues
            crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
            crParameterDiscreteValue.Value = "T.Prod.Lucro"
            crParameterValues.Add(crParameterDiscreteValue)
            crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

            crParameterFieldDefinition = crParameterFieldDefinitions.Item("Operacao")
            crParameterValues = crParameterFieldDefinition.CurrentValues
            crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
            crParameterDiscreteValue.Value = "Partic.Lucro"
            crParameterValues.Add(crParameterDiscreteValue)
            crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

            crParameterFieldDefinition = crParameterFieldDefinitions.Item("PesoBruto")
            crParameterValues = crParameterFieldDefinition.CurrentValues
            crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
            crParameterDiscreteValue.Value = "Partic.Venda"
            crParameterValues.Add(crParameterDiscreteValue)
            crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

            crParameterFieldDefinition = crParameterFieldDefinitions.Item("PesoLiquido")
            crParameterValues = crParameterFieldDefinition.CurrentValues
            crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
            crParameterDiscreteValue.Value = ""
            crParameterValues.Add(crParameterDiscreteValue)
            crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

            crParameterFieldDefinition = crParameterFieldDefinitions.Item("DataDestino")
            crParameterValues = crParameterFieldDefinition.CurrentValues
            crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
            crParameterDiscreteValue.Value = ""
            crParameterValues.Add(crParameterDiscreteValue)
            crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

            crParameterFieldDefinition = crParameterFieldDefinitions.Item("ClienteCnpj")
            crParameterValues = crParameterFieldDefinition.CurrentValues
            crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
            crParameterDiscreteValue.Value = ""
            crParameterValues.Add(crParameterDiscreteValue)
            crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

            crParameterFieldDefinition = crParameterFieldDefinitions.Item("NomeCliente")
            crParameterValues = crParameterFieldDefinition.CurrentValues
            crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
            crParameterDiscreteValue.Value = ""
            crParameterValues.Add(crParameterDiscreteValue)
            crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

            crParameterFieldDefinition = crParameterFieldDefinitions.Item("NotaDestino")
            crParameterValues = crParameterFieldDefinition.CurrentValues
            crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
            crParameterDiscreteValue.Value = ""
            crParameterValues.Add(crParameterDiscreteValue)
            crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

            crParameterFieldDefinition = crParameterFieldDefinitions.Item("Peso")
            crParameterValues = crParameterFieldDefinition.CurrentValues
            crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
            crParameterDiscreteValue.Value = ""
            crParameterValues.Add(crParameterDiscreteValue)
            crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

            crParameterFieldDefinition = crParameterFieldDefinitions.Item("REC")
            crParameterValues = crParameterFieldDefinition.CurrentValues
            crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
            crParameterDiscreteValue.Value = ""
            crParameterValues.Add(crParameterDiscreteValue)
            crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

            crParameterFieldDefinition = crParameterFieldDefinitions.Item("CTRC")
            crParameterValues = crParameterFieldDefinition.CurrentValues
            crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
            crParameterDiscreteValue.Value = ""
            crParameterValues.Add(crParameterDiscreteValue)
            crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

            crParameterFieldDefinition = crParameterFieldDefinitions.Item("Unitario")
            crParameterValues = crParameterFieldDefinition.CurrentValues
            crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
            crParameterDiscreteValue.Value = ""
            crParameterValues.Add(crParameterDiscreteValue)
            crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

            crParameterFieldDefinition = crParameterFieldDefinitions.Item("Valor")
            crParameterValues = crParameterFieldDefinition.CurrentValues
            crParameterDiscreteValue = New CrystalDecisions.Shared.ParameterDiscreteValue
            crParameterDiscreteValue.Value = ""
            crParameterValues.Add(crParameterDiscreteValue)
            crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

            If Dir(NomeArquivo).Length > 0 Then Kill(NomeArquivo)

            crptRelatorio.ExportToDisk(ExportFormatType.PortableDocFormat, arquivo)

            If IO.File.Exists(arquivo) Then
                ScriptManager.RegisterClientScriptBlock(Me.Page, Me.Page.GetType(), Guid.NewGuid().ToString, "window.open('" & NomeArquivo2 & "');", True)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        Finally
            crptRelatorio.Close()
            crptRelatorio.Dispose()
        End Try
    End Sub

    Protected Sub cmdBuscaCliOrigem_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            HttpContext.Current.Session("ssCampo") = "Livre"
            ucConsultaClientes.Limpar()
            Popup.ConsultaDeClientes(Me.Page, "objClienteCarCliRepre" & HID.Value, "txtNome")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkLimpar_Click(sender As Object, e As EventArgs) Handles lnkLimpar.Click
        Try
            Limpar()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkRelatorio_Click(sender As Object, e As EventArgs) Handles lnkRelatorio.Click
        Try
            If ValidarCampos() Then
                BuscarRegistros()
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "RelatorioParticipacaoNaVendaLucroBruto")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

End Class