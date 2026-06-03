Imports System.Data
Imports System.IO
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class RelatorioMargemDeLucro
    Inherits BasePage

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.setMenu(eModulo.Revenda)
            If Not IsPostBack And IsConnect Then
                'If Funcoes.VerificaPermissao("RelatorioMargemDeLucro", "ACESSAR") Then 'criar processo
                BuscaEmpresa()
                    CarregarSafras()
                    ddl.Carregar(ddlUnidade, CarregarDDL.Tabela.UnidadeDeNegocio, "", True)
                    Funcoes.VerificaUnidadeDeNegocio(ddlUnidade, ddlEmpresa, True)
                    ddl.Carregar(ddlMarca, CarregarDDL.Tabela.Marca, "")
                    HID.Value = Guid.NewGuid().ToString
                    ucConsultaClientes.SetarHID(HID.Value)
                    LiberaEmpresa()
                'Else
                '    MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Revenda.aspx")
                '    Exit Sub
                'End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub BuscaEmpresa()
        ddl.Carregar(ddlEmpresa, CarregarDDL.Tabela.ClientesXEmpresas, "", True)
    End Sub

    Private Sub CarregarSafras()
        ddl.Carregar(ddlSafra, CarregarDDL.Tabela.Safra, "", True)
    End Sub

    Private Function ValidarCampos() As Boolean
        If ddlEmpresa.SelectedIndex = 0 Then
            MsgBox(Me.Page, "Informe a empresa.")
            Return False
        End If
        Return True
    End Function

    Private Function getParam() As String
        Dim param As String = ""

        If Not String.IsNullOrWhiteSpace(ddlEmpresa.SelectedItem.Text) Then
            param &= "EMPRESA" & IIf(chkConsolidarEmpresa.Checked, " Consolidada:", ":") & ddlEmpresa.SelectedItem.Text & " - "
        End If
        If Not String.IsNullOrWhiteSpace(txtCliente.Text) Then
            param &= "CLIENTE:" & IIf(chkConsolidarCliente.Checked, " Consolidado:", ":") & txtCliente.Text & vbCrLf
        End If

        If ucSelecaoProduto.TemSelecionado Then
            Dim Retorno As ArrayList = ucSelecaoProduto.GetSqlEParametrosRelatorio("", "", "")
            param &= Retorno(1)
        End If

        If Not String.IsNullOrWhiteSpace(ddlSafra.SelectedItem.Text) Then
            param &= "SAFRA: " & ddlSafra.SelectedItem.Text & " - "
        End If
        If Not String.IsNullOrWhiteSpace(ddlMarca.SelectedItem.Text) Then
            param &= "MARCA: " & ddlMarca.SelectedItem.Text
        End If

        Return param
    End Function

    Private Sub Limpar()
        Funcoes.VerificaUnidadeDeNegocio(ddlUnidade, ddlEmpresa)
        txtCliente.Text = ""
        txtCodigoCliente.Value = ""
        ucSelecaoProduto.Limpar()
        ddlSafra.SelectedIndex = 0
        chkConsolidarEmpresa.Checked = False
        chkConsolidarCliente.Checked = False
        chkVisualizarPedido.Checked = False
        chkVisualizarPedido.Visible = False
        ckAgruparCliente.Checked = False
        LiberaEmpresa()
    End Sub

    Private Sub LiberaEmpresa()
        If Not UsuarioServidor.LiberaEmpresa Then
            ddlUnidade.Enabled = False
            ddlEmpresa.Enabled = False
        End If
    End Sub

    Public Overrides Sub Carregar(obj As IBaseEntity)
        Try
            If Not Session("objClienteRelMargLuc" & HID.Value) Is Nothing Then
                Dim itemCliente As ListItem = Funcoes.FormatarListItemCliente(CType(Session("objClienteRelMargLuc" & HID.Value), [Lib].Negocio.Cliente))
                txtCliente.Text = itemCliente.Text
                txtCodigoCliente.Value = itemCliente.Value
                Session.Remove("objClienteRelMargLuc" & HID.Value)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub cmdBuscaCliente_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            HttpContext.Current.Session("ssCampo") = "Livre"
            ucConsultaClientes.Limpar()
            Popup.ConsultaDeClientes(Me.Page, "objClienteRelMargLuc" & HID.Value, "txtNome")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub ddlUnidade_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddlUnidade.SelectedIndexChanged
        Try
            ddl.Carregar(ddlEmpresa, CarregarDDL.Tabela.Empresas, ddlUnidade.SelectedValue, True)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Public Sub BuscarRegistros()
        Try
            If Funcoes.VerificaPermissao("RelatorioMargemDeLucro", "RELATORIO") Then
                Dim Parametros As String = "Parametros:" & vbCrLf
                Dim ds As New DataSet
                Dim SQL As String

                '*** SQL Margem de lucro -----------------------------------------------------------------------------------------------------------------------------

                SQL = " select  P.Cliente,     " & vbCrLf & _
                      " 		P.EndCliente,    " & vbCrLf & _
                      " 		Cli.Nome,    " & vbCrLf & _
                      "         p.Pedido_Id, " & vbCrLf & _
                      "         PxI.Produto_Id, " & vbCrLf & _
                      "         Prod.Nome as Descricao, " & vbCrLf & _
                      "         M.Cifrao," & vbCrLf & _
                      "         M.Classificacao," & vbCrLf & _
                      "         max(isnull(PxI.UnitarioOficialCompra,0)) as UnitarioOficialCompra,  " & vbCrLf & _
                      "         max(isnull(PxI.UnitarioMoedaCompra,0)) as UnitarioMoedaCompra,  " & vbCrLf & _
                      " 		Sum(case     " & vbCrLf & _
                      " 			  When PxI.TipoDeLancamento = 'E'    " & vbCrLf & _
                      " 			    then PxI.Quantidade * - 1    " & vbCrLf & _
                      " 			    else PxI.Quantidade     " & vbCrLf & _
                      " 			end) as Quantidade,    " & vbCrLf & _
                      " 		Sum(case     " & vbCrLf & _
                      " 		  	  When PxI.TipoDeLancamento = 'E'    " & vbCrLf & _
                      " 			    then PxI.TotalOficial * - 1    " & vbCrLf & _
                      " 			    else PxI.TotalOficial    " & vbCrLf & _
                      " 			end) as TotalOficial,   " & vbCrLf & _
                      " 		Sum(case     " & vbCrLf & _
                      " 		      When PxI.TipoDeLancamento = 'E'    " & vbCrLf & _
                      " 			    then PxI.TotalMoeda * - 1    " & vbCrLf & _
                      " 				else PxI.TotalMoeda    " & vbCrLf & _
                      " 		    end) as TotalMoeda  " & vbCrLf & _
                      " 	into #temp    " & vbCrLf & _
                      " 	FROM Pedidos P    " & vbCrLf & _
                      "    Inner Join Suboperacoes SO  " & vbCrLf & _
                      "       on SO.Operacao_Id     = P.Operacao  " & vbCrLf & _
                      "      and SO.Suboperacoes_Id = P.Suboperacao  " & vbCrLf & _
                      "    Inner Join PedidoXItemxLancamento PxI    " & vbCrLf & _
                      "       on P.Empresa_Id    = PxI.Empresa_Id    " & vbCrLf & _
                      "      and P.EndEmpresa_Id = PXI.EndEmpresa_Id    " & vbCrLf & _
                      "      and P.Pedido_Id     = PxI.Pedido_Id    " & vbCrLf & _
                      "    INNER JOIN   Produtos Prod  " & vbCrLf & _
                      "       ON PxI.Produto_Id = Prod.Produto_Id  " & vbCrLf & _
                      "    INNER JOIN Clientes Cli  " & vbCrLf & _
                      "       ON P.Cliente = Cli.Cliente_Id  " & vbCrLf & _
                      "      AND P.EndCliente = Cli.Endereco_Id  " & vbCrLf & _
                      "    INNER Join Moedas M" & vbCrLf & _
                      "      on M.Moeda_id = P.Moeda" & vbCrLf & _
                      "    WHERE  P.Situacao      = 1 " & vbCrLf

                If Not String.IsNullOrWhiteSpace(ddlEmpresa.SelectedValue) Then
                    If chkConsolidarEmpresa.Checked Then
                        SQL &= " AND left(P.Empresa_id, 8)    = '" & Left(ddlEmpresa.SelectedValue.Split("-")(0), 8) & "'" & vbCrLf
                    Else
                        SQL &= " AND P.Empresa_id    = '" & ddlEmpresa.SelectedValue.Split("-")(0) & "'" & vbCrLf & _
                               " AND P.EndEmpresa_Id = " & ddlEmpresa.SelectedValue.Split("-")(1) & " " & vbCrLf
                        Parametros &= "Empresa: " & ddlEmpresa.SelectedItem.Text & vbCrLf
                    End If
                End If

                If Not String.IsNullOrWhiteSpace(txtCliente.Text) Then
                    If chkConsolidarCliente.Checked Then
                        SQL &= " AND left(P.Cliente, 8) = '" & Left(txtCodigoCliente.Value.Split("-")(0), 8) & "'" & vbCrLf
                    Else
                        SQL &= " AND P.Cliente = '" & txtCodigoCliente.Value.Split("-")(0) & "'" & vbCrLf & _
                               " AND P.EndCliente = " & txtCodigoCliente.Value.Split("-")(1) & " " & vbCrLf
                        Parametros &= "Cliente: " & txtCliente.Text & vbCrLf
                    End If
                End If

                If ddlSafra.SelectedIndex > 0 Then
                    SQL &= " AND P.Safra   = '" & ddlSafra.SelectedValue & "'" & vbCrLf
                    Parametros &= "Safra: " & ddlSafra.SelectedItem.Text & vbCrLf
                End If

                If ddlMarca.SelectedIndex > 0 Then
                    SQL &= " AND isnull(Prod.Marca,0) = " & ddlMarca.SelectedValue & vbCrLf
                    Parametros &= "Marca: " & ddlMarca.SelectedItem.Text & vbCrLf
                End If

                If ucSelecaoProduto.TemSelecionado Then
                    Dim Retorno As ArrayList = ucSelecaoProduto.GetSqlEParametrosRelatorio("Prod.Grupo", "PxI.Produto_Id", "")
                    SQL &= "And " & Retorno(0)
                    Parametros &= Retorno(1)
                End If

                SQL &= "            AND So.Entradasaida = 'S'  " & vbCrLf & _
                       "            AND So.Classe in ('" & eClassesOperacoes.VENDAS.ToString & "','" & eClassesOperacoes.REMESSAS.ToString & "','" & eClassesOperacoes.EXPORTACOES.ToString & "','" & eClassesOperacoes.VENDASAORDEM.ToString & "','" & eClassesOperacoes.GLOBAL.ToString & "')" & vbCrLf & _
                       " 		 group by P.Cliente,     " & vbCrLf & _
                       "         P.EndCliente, " & vbCrLf & _
                       "         Cli.Nome,  " & vbCrLf & _
                       "         p.Pedido_Id, " & vbCrLf & _
                       "         PxI.Produto_Id, " & vbCrLf & _
                       "         Prod.Nome," & vbCrLf & _
                       "         M.Cifrao," & vbCrLf & _
                       "         M.Classificacao" & vbCrLf & _
                       "  having Sum(Case" & vbCrLf & _
                       "               When PxI.TipoDeLancamento = 'E'" & vbCrLf & _
                       " 				 then PxI.Quantidade * - 1" & vbCrLf & _
                       " 				 else PxI.Quantidade" & vbCrLf & _
                       " 			 end) > 0" & vbCrLf

                If chkVisualizarPedido.Checked Then
                    SQL &= " Select cliente," & vbCrLf & _
                         "        EndCliente," & vbCrLf & _
                         "        Pedido_id," & vbCrLf & _
                         "        sum(Case when UnitarioOficialCompra = 0 then 0 else (Totaloficial * 100) / (UnitarioOficialCompra * Quantidade) - 100 end) as PercPedido" & vbCrLf & _
                         "   Into #PartPedido" & vbCrLf & _
                         "   from #temp" & vbCrLf & _
                         "  where UnitarioOficialCompra > 0" & vbCrLf & _
                         "  group by cliente, EndCliente, Pedido_Id;" & vbCrLf

                    SQL &= "select cliente," & vbCrLf & _
                           "       EndCliente," & vbCrLf & _
                           "       sum(Case when UnitarioOficialCompra = 0 then 0 else (Totaloficial * 100) / (UnitarioOficialCompra * Quantidade) - 100 end) as PercCliente" & vbCrLf & _
                           "  Into #PartCliente" & vbCrLf & _
                           "  from #temp" & vbCrLf & _
                           " where UnitarioOficialCompra > 0" & vbCrLf & _
                           " group by cliente, EndCliente;" & vbCrLf
                ElseIf ckAgruparCliente.Checked Then
                    SQL &= "select cliente," & vbCrLf & _
                          "       EndCliente," & vbCrLf & _
                          "       sum(Case when UnitarioOficialCompra = 0 then 0 else (Totaloficial * 100) / (UnitarioOficialCompra * Quantidade) - 100 end) as PercCliente" & vbCrLf & _
                          "  Into #PartCliente" & vbCrLf & _
                          "  from #temp" & vbCrLf & _
                          " where UnitarioOficialCompra > 0" & vbCrLf & _
                          " group by cliente, EndCliente;" & vbCrLf
                End If

                SQL &= " Select "
                If chkVisualizarPedido.Checked Then
                    SQL &= " #Temp.Cliente, #Temp.EndCliente, Nome, #PartCliente.PercCliente, #Temp.Pedido_Id, #PartPedido.PercPedido," & vbCrLf
                ElseIf ckAgruparCliente.Checked Then
                    SQL &= " #Temp.Cliente, #Temp.EndCliente, Nome, #PartCliente.PercCliente, 0 as Pedido_Id, 0 as PercPedido,  " & vbCrLf
                End If

                SQL &= "         Produto_Id, " & vbCrLf & _
                       "         Descricao, " & vbCrLf & _
                       "         sum(UnitarioOficialCompra * Quantidade) / Sum(Quantidade) as UnitarioMedioOficialCompra,  " & vbCrLf & _
                       "         Sum(TotalOficial) / Sum(Quantidade) as UnitarioMedioOficialVenda,    " & vbCrLf & _
                       "         sum(UnitarioMoedaCompra   * Quantidade) / Sum(Quantidade) as UnitarioMedioMoedaCompra,     " & vbCrLf & _
                       "         sum(TotalMoeda) / Sum(Quantidade) as UnitarioMedioMoedaVenda,      " & vbCrLf & _
                       "         Sum(Quantidade) as Quantidade,  " & vbCrLf & _
                       "         convert(numeric(08,2),case   " & vbCrLf & _
                       "             when sum(UnitarioOficialCompra * Quantidade) = 0  " & vbCrLf & _
                       "               then 0  " & vbCrLf & _
                       "               else (Sum(TotalOficial) * 100) / sum(UnitarioOficialCompra * Quantidade) - 100  " & vbCrLf & _
                       "         end) Margem,     " & vbCrLf & _
                       "         ' % em ' + #temp.Cifrao DescMargem" & vbCrLf & _
                       "    from #Temp  " & vbCrLf

                If chkVisualizarPedido.Checked Then
                    SQL &= "   Inner Join #PartPedido " & vbCrLf & _
                           "      on #Temp.cliente    = #PartPedido.cliente" & vbCrLf & _
                           "     and #Temp.EndCliente = #PartPedido.EndCliente" & vbCrLf & _
                           "     And #Temp.Pedido_Id  = #PartPedido.Pedido_Id" & vbCrLf & _
                           "   inner Join #PartCliente" & vbCrLf & _
                           "      on #Temp.cliente    = #PartCliente.cliente" & vbCrLf & _
                           "     and #Temp.EndCliente = #PartCliente.EndCliente" & vbCrLf
                ElseIf ckAgruparCliente.Checked Then
                    SQL &= "   inner Join #PartCliente" & vbCrLf & _
                           "      on #Temp.cliente    = #PartCliente.cliente" & vbCrLf & _
                           "     and #Temp.EndCliente = #PartCliente.EndCliente" & vbCrLf
                End If

                SQL &= "   group By   " & vbCrLf
                If chkVisualizarPedido.Checked Then
                    SQL &= " #Temp.Cliente, #Temp.EndCliente, Nome, #PartCliente.PercCliente, #Temp.Pedido_Id, #PartPedido.PercPedido," & vbCrLf
                ElseIf ckAgruparCliente.Checked Then
                    SQL &= " #Temp.Cliente, #Temp.EndCliente, Nome, #PartCliente.PercCliente, " & vbCrLf
                End If

                SQL &= " Produto_Id, Descricao, #temp.Classificacao, #temp.Cifrao" & vbCrLf

                ds = Banco.ConsultaDataSet(SQL, "RelatorioMargemDeLucro")

                Dim parameters = New Dictionary(Of String, Object)()

                If ckAgruparCliente.Checked Then
                    parameters.Add("ConsultaParametros", getParam())
                    parameters.Add("VisualizarPedido", chkVisualizarPedido.Checked)
                Else
                    parameters.Add("ConsultaParametros", getParam())
                End If
                Funcoes.BindReport(Me.Page, ds, IIf(ckAgruparCliente.Checked, "CrRelatorioMargemDeLucroPorCliente", "CrRelatorioMargemDeLucro"), eExportType.PDF, parameters)
            Else
                MsgBox(Me.Page, "Usuário sem permissão para emitir relatório.", eTitulo.Info)
            End If
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

    Protected Sub lnkLimpar_Click(sender As Object, e As EventArgs) Handles lnkLimpar.Click
        Try
            Limpar()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "RelatorioMargemDeLucro")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub ckAgruparCliente_CheckedChanged(sender As Object, e As EventArgs) Handles ckAgruparCliente.CheckedChanged
        Try
            chkVisualizarPedido.Checked = False
            chkVisualizarPedido.Visible = ckAgruparCliente.Checked
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub
End Class