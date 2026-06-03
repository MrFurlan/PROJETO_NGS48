Imports System.Data
Imports System.IO
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class RelatorioDeComissoes
    Inherits BasePage

    Private dateEmissao As New DateTime
    Private intPagina As Integer = 0

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.setMenu(eModulo.Expedicao)
            If Not IsPostBack And IsConnect Then
                If Funcoes.VerificaPermissao("RelatorioDeComissoes", "ACESSAR") Then
                    BuscaEmpresa()
                    Funcoes.VerificaEmpresa(ddlEmpresa)
                    CarregarSafras()
                    ddl.Carregar(ddlMarca, CarregarDDL.Tabela.Marca, "")
                    ddl.Carregar(ddlMoeda, CarregarDDL.Tabela.Moeda, "")
                    HID.Value = Guid.NewGuid().ToString()
                    ucConsultaClientes.SetarHID(HID.Value)

                    LiberaEmpresa()
                Else
                    MsgBox(Me.Page, "Usuário sem permissão para acessar essa página", "~/Expedição.aspx")
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Public Overrides Sub Carregar(obj As IBaseEntity)
        If Not Session("objClienteRepCom" & HID.Value) Is Nothing Then
            Dim itemCliente As ListItem = Funcoes.FormatarListItemCliente(CType(Session("objClienteRepCom" & HID.Value), [Lib].Negocio.Cliente))
            txtRepresentante.Text = itemCliente.Text
            txtCodigoRepresentante.Value = itemCliente.Value
            Session.Remove("objClienteRepCom" & HID.Value)
        ElseIf Not Session("objClienteCliCom" & HID.Value) Is Nothing Then
            Dim itemCliente As ListItem = Funcoes.FormatarListItemCliente(CType(Session("objClienteCliCom" & HID.Value), [Lib].Negocio.Cliente))
            txtClientes.Text = itemCliente.Text
            txtCodigoCliente.Value = itemCliente.Value
            Session.Remove("objClienteCliCom" & HID.Value)
        End If
    End Sub

    Private Sub BuscaEmpresa()
        ddl.Carregar(ddlEmpresa, CarregarDDL.Tabela.ClientesXEmpresas, "", True)
    End Sub

    Private Sub CarregarSafras()
        ddl.Carregar(ddlSafra, CarregarDDL.Tabela.Safra, "", True)
    End Sub

    Private Function ValidarCampos() As Boolean
        'ElseIf txtRepresentante.Text.Length = 0 Then
        '    MsgBox(Me.Page, "Representante não foi selecionando")
        '    Return False
        'ElseIf txtClientes.Text.Length = 0 Then
        '    MsgBox(Me.Page, "Cliente não foi selecionando")
        '    Return False
        If String.IsNullOrWhiteSpace(ddlEmpresa.SelectedValue) Then
            MsgBox(Me.Page, "Empresa não foi selecionada.")
            Return False
        End If
        If String.IsNullOrWhiteSpace(ddlSafra.SelectedValue) Then
            MsgBox(Me.Page, "Safra não foi selecionada.")
            Return False
        End If
        'If String.IsNullOrWhiteSpace(ddlTroca.SelectedValue) Then
        '    MsgBox(Me.Page, "Troca não foi selecionada")
        '    Return False
        'End If


        Return True
    End Function

    Protected Sub cmdBuscaRepresentante_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdBuscaRepresentante.Click
        Try
            ucConsultaClientes.Limpar()
            Popup.ConsultaDeClientes(Me.Page, "objClienteRepCom" & HID.Value)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub cmdBuscaCliente_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdBuscaCliente.Click
        Try
            ucConsultaClientes.Limpar()
            Popup.ConsultaDeClientes(Me.Page, "objClienteCliCom" & HID.Value)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub EmitirRelatorio(ByVal Pdf As Boolean)
        Try
            If Funcoes.VerificaPermissao("RelatorioDeComissoes", "RELATORIO") Then
                If ValidarCampos() Then
                    Dim sqlProd As String = ""
                    Dim DescricaoProduto As String = ""
                    Dim dsProduto As New DataSet
                    Dim ds As New DataSet
                    Dim SQL As String
                    Dim Parametros As String = "Parametros:" & vbCrLf
                    Dim Cliente As String = ""
                    Dim strEmpresa() As String = ddlEmpresa.SelectedValue.Split("-")
                    Dim strRepresentante() As String = txtCodigoRepresentante.Value.Split("-")
                    Dim strCliente() As String = txtCodigoCliente.Value.Split("-")

                    SQL = " select PC.Representante_Id,   " & vbCrLf & _
                          "        PC.EndRepresentante_Id,   " & vbCrLf & _
                          "        P.Cliente,   " & vbCrLf & _
                          "        P.EndCliente,  " & vbCrLf & _
                          "        isnull(P.troca,0) as troca," & vbCrLf & _
                          "        p.Pedido_Id,  " & vbCrLf & _
                          "        P.DataPedido,  " & vbCrLf & _
                          "        PxI.Produto_Id,  " & vbCrLf & _
                          "        Prd.Descricao,  " & vbCrLf & _
                          "        P.Moeda,    " & vbCrLf & _
                          "        sum(case When PxI.TipoDeLancamento = 'N' then isnull(PxI.UnitarioOficialCompra,0) end) as UnitarioOficialCompra," & vbCrLf & _
                          "        sum(case When PxI.TipoDeLancamento = 'N' then isnull(PxI.UnitarioMoedaCompra,0)   end) as UnitarioMoedaCompra," & vbCrLf & _
                          "        Sum(case When PxI.TipoDeLancamento = 'E' then PxI.TotalOficial * - 1 else PxI.TotalOficial end) / Sum(case When PxI.TipoDeLancamento = 'E' then PxI.Quantidade * - 1 else PxI.Quantidade end) as UnitarioOficial,  " & vbCrLf & _
                          "        Sum(case When PxI.TipoDeLancamento = 'E' then PxI.TotalMoeda   * - 1 else PxI.TotalMoeda end)   / Sum(case When PxI.TipoDeLancamento = 'E' then PxI.Quantidade * - 1 else PxI.Quantidade end) as UnitarioMoeda,  " & vbCrLf & _
                          "        Sum(case   " & vbCrLf & _
                          "               When PxI.TipoDeLancamento = 'E'  " & vbCrLf & _
                          "                 then PxI.Quantidade * - 1  " & vbCrLf & _
                          "                 else PxI.Quantidade   " & vbCrLf & _
                          "            end) as Quantidade,  " & vbCrLf & _
                          "        Sum(case   " & vbCrLf & _
                          "               When PxI.TipoDeLancamento = 'E'  " & vbCrLf & _
                          "                 then PxI.TotalOficial * - 1  " & vbCrLf & _
                          "                 else PxI.TotalOficial   " & vbCrLf & _
                          "            end) as TotalOficial,  " & vbCrLf & _
                          "        Sum(case   " & vbCrLf & _
                          "               When PxI.TipoDeLancamento = 'E'  " & vbCrLf & _
                          "                 then PxI.TotalMoeda * - 1  " & vbCrLf & _
                          "                 else PxI.TotalMoeda   " & vbCrLf & _
                          "            end) as TotalMoeda,  " & vbCrLf & _
                          "         PC.Tabela_Id  " & vbCrLf & _
                          "        into #temp  " & vbCrLf & _
                          "        from Pedidos P  " & vbCrLf & _
                          "        Inner Join PedidoXItemxLancamento PxI  " & vbCrLf & _
                          "          on P.Empresa_Id    = PxI.Empresa_Id  " & vbCrLf & _
                          "         and P.EndEmpresa_Id = PXI.EndEmpresa_Id  " & vbCrLf & _
                          "         and P.Pedido_Id     = PxI.Pedido_Id  " & vbCrLf & _
                          "        Inner Join Produtos Prd " & vbCrLf & _
                          "           on Prd.Produto_id = PxI.Produto_Id " & vbCrLf & _
                          "        Inner join SubOperacoes so" & vbCrLf & _
                          "           on P.operacao    = so.Operacao_id" & vbCrLf & _
                          "          and p.suboperacao = so.suboperacoes_id" & vbCrLf & _
                          "        Inner Join PedidoXTabelaDeComissao PC  " & vbCrLf & _
                          "           on P.Empresa_Id    = PC.Empresa_Id  " & vbCrLf & _
                          "          and P.EndEmpresa_Id = PC.EndEmpresa_Id  " & vbCrLf & _
                          "          and P.Pedido_Id     = PC.Pedido_Id  " & vbCrLf & _
                          "          and PC.Produto_Id   = PxI.Produto_Id   " & vbCrLf & _
                          "        WHERE P.Situacao = 1 " & vbCrLf
                    If ddlTroca.SelectedIndex > 0 Then
                        SQL &= "          and isnull(P.troca,0) = " & ddlTroca.SelectedValue & vbCrLf
                    End If

                    If strEmpresa(0).Length > 0 Then
                        If chkConsolidaEmpresa.Checked Then
                            SQL &= " AND Left(P.Empresa_id,8) = '" & Left(strEmpresa(0), 8) & "'" & vbCrLf
                            Parametros &= "Empresa: " & ddlEmpresa.SelectedItem.Text & vbCrLf & _
                                          "Consolidado Empresa." & vbCrLf
                        Else
                            SQL &= " AND P.Empresa_id= '" & strEmpresa(0) & "'" & vbCrLf
                            SQL &= " AND P.EndEmpresa_Id= " & strEmpresa(1) & " " & vbCrLf
                            Parametros &= "Empresa: " & ddlEmpresa.SelectedItem.Text & vbCrLf
                        End If
                    End If

                    If strRepresentante(0).Length > 0 Then
                        SQL &= " AND PC.Representante_Id= '" & strRepresentante(0) & "'" & vbCrLf
                        SQL &= " AND PC.EndRepresentante_Id= " & strRepresentante(1) & " " & vbCrLf
                        Parametros &= "Representante: " & txtRepresentante.Text & vbCrLf
                    End If
                    If strCliente(0).Length > 0 Then
                        SQL &= " AND P.Cliente= '" & strCliente(0) & "'" & vbCrLf
                        SQL &= " AND P.EndCliente= " & strCliente(1) & " " & vbCrLf
                        Parametros &= "Cliente: " & txtClientes.Text & vbCrLf
                    End If
                    If ddlSafra.SelectedIndex > 0 Then
                        SQL &= " AND P.Safra      = '" & ddlSafra.SelectedValue & "'" & vbCrLf
                        Parametros &= "Safra: " & ddlSafra.SelectedItem.Text & " ... "
                    End If
                    If ddlMarca.SelectedIndex > 0 Then
                        SQL &= " AND isnull(Prd.Marca,0) = " & ddlMarca.SelectedValue & vbCrLf
                        Parametros &= "Marca: " & ddlMarca.SelectedItem.Text & " ... "
                    End If
                    If ucSelecaoProduto.TemSelecionado Then
                        Dim ret As New ArrayList
                        ret = ucSelecaoProduto.GetSqlEParametrosRelatorio("Prd.Grupo", "Prd.Produto_Id", "", True)
                        SQL &= " and " & ret(0) & vbCrLf
                        Parametros &= "Produtos: " & ret(1) & vbCrLf
                    End If
                    If ddlMoeda.SelectedIndex > 0 Then
                        SQL &= " AND P.Moeda = " & ddlMoeda.SelectedValue & vbCrLf
                        Parametros &= "Moeda do Pedido: " & ddlMoeda.SelectedItem.Text & " ... "
                    End If

                    SQL &= " group by PC.Representante_Id,   " & vbCrLf & _
                           " 	      PC.EndRepresentante_Id,   " & vbCrLf & _
                           " 	      P.Cliente,   " & vbCrLf & _
                           " 	      P.EndCliente,  " & vbCrLf & _
                           " 	      p.Pedido_Id,  " & vbCrLf & _
                           "          isnull(P.Troca,0)," & vbCrLf & _
                           "          P.Moeda," & vbCrLf & _
                           " 	      P.DataPedido,  " & vbCrLf & _
                           " 	      PxI.Produto_Id,  " & vbCrLf & _
                           " 	      Prd.Descricao,  " & vbCrLf & _
                           "          PC.Tabela_Id" & vbCrLf & _
                           "   having Sum(case   " & vbCrLf & _
                           "                When PxI.TipoDeLancamento = 'E' " & vbCrLf & _
                           "                  then PxI.Quantidade * - 1 " & vbCrLf & _
                           "                  else PxI.Quantidade " & vbCrLf & _
                           "              end) > 0 " & vbCrLf

                    'SQL &= " select T.Representante_Id," & vbCrLf & _
                    '       "        T.EndRepresentante_Id," & vbCrLf & _
                    '       "        sum(T.TotalOficial)*100/sum(T.UnitarioOficialCompra * T.Quantidade)-100 as PercRep" & vbCrLf & _
                    '       "   into #PercRep" & vbCrLf & _
                    '       "   from #Temp T" & vbCrLf & _
                    '       "  where T.UnitarioOficialCompra > 0" & vbCrLf & _
                    '       "  group by T.Representante_Id, T.EndRepresentante_Id" & vbCrLf

                    SQL &= "select Representante_id," & vbCrLf & _
                           "       EndRepresentante_id," & vbCrLf & _
                           "       sum(case" & vbCrLf & _
                           "             when Classificacao = 'O' and troca = 0" & vbCrLf & _
                           "               then PercRep" & vbCrLf & _
                           "               else 0" & vbCrLf & _
                           "           end) PercRepOficial," & vbCrLf & _
                           "       sum(case" & vbCrLf & _
                           "             when Classificacao <> 'O' and troca = 0" & vbCrLf & _
                           "               then PercRep" & vbCrLf & _
                           "               else 0" & vbCrLf & _
                           "           end) PercRepMoeda," & vbCrLf & _
                           "       sum(case" & vbCrLf & _
                           "             when Classificacao = 'O' and troca = 1" & vbCrLf & _
                           "               then PercRep" & vbCrLf & _
                           "               else 0" & vbCrLf & _
                           "           end) PercRepOficialTroca," & vbCrLf & _
                           "       sum(case" & vbCrLf & _
                           "             when Classificacao <> 'O' and troca = 1" & vbCrLf & _
                           "               then PercRep" & vbCrLf & _
                           "               else 0" & vbCrLf & _
                           "           end) PercRepMoedaTroca" & vbCrLf & _
                           " into #PercRep" & vbCrLf & _
                           " from (" & vbCrLf & _
                           "		 select T.Representante_Id," & vbCrLf & _
                           "				T.EndRepresentante_Id," & vbCrLf & _
                           "				M.Classificacao," & vbCrLf & _
                           "                T.Troca," & vbCrLf & _
                           "				case" & vbCrLf & _
                           "				  when M.Classificacao = 'O'" & vbCrLf & _
                           "				    then sum(T.TotalOficial)*100/sum(T.UnitarioOficialCompra * T.Quantidade)-100" & vbCrLf & _
                           "				    else sum(T.TotalMoeda)*100/sum(T.UnitarioMoedaCompra * T.quantidade)- 100" & vbCrLf & _
                           "				end PercRep" & vbCrLf & _
                           "		   from #Temp T" & vbCrLf & _
                           "		  inner join Moedas M" & vbCrLf & _
                           "		     on T.Moeda = M.Moeda_id" & vbCrLf & _
                           "		  where case" & vbCrLf & _
                           "		          when M.Classificacao = 'O'" & vbCrLf & _
                           "		           then T.UnitarioOficialCompra" & vbCrLf & _
                           "		           else T.UnitarioMoedaCompra" & vbCrLf & _
                           "		        end > 0" & vbCrLf & _
                           "		    and case" & vbCrLf & _
                           "		          when M.Classificacao = 'O'" & vbCrLf & _
                           "		           then T.UnitarioOficialCompra" & vbCrLf & _
                           "		           else T.UnitarioMoedaCompra" & vbCrLf & _
                           "		        end < case" & vbCrLf & _
                           "						  when M.Classificacao = 'O'" & vbCrLf & _
                           "						   then T.UnitarioOficial" & vbCrLf & _
                           "						   else T.UnitarioMoeda" & vbCrLf & _
                           "						end" & vbCrLf & _
                           "		  group by T.Representante_Id, T.EndRepresentante_Id,M.Classificacao,T.Troca" & vbCrLf & _
                           "	) sb" & vbCrLf & _
                           "group by Representante_id," & vbCrLf & _
                           "         EndRepresentante_id;" & vbCrLf

                    SQL &= " select T.Representante_Id,  " & vbCrLf & _
                           "        T.EndRepresentante_Id,  " & vbCrLf & _
                           "        #PercRep.PercRepOficial," & vbCrLf & _
                           "        #PercRep.PercRepMoeda," & vbCrLf & _
                           "        #PercRep.PercRepOficialTroca," & vbCrLf & _
                           "        #PercRep.PercRepMoedaTroca," & vbCrLf & _
                           "        Rep.Nome + ' ' + Rep.cidade + '-' + Rep.Estado as NomeRepresentante, " & vbCrLf & _
                           "        T.Cliente,   " & vbCrLf & _
                           "        T.EndCliente,  " & vbCrLf & _
                           "        Cli.Nome + ' ' + Cli.Cidade + '-' + Cli.Estado as NomeCliente, " & vbCrLf & _
                           "        convert(nvarchar,T.Pedido_Id) + '-' + Case when moeda = 1 then 'R$' else '$' end + case when T.troca = 1 then '-T' else '' end as Pedido_id," & vbCrLf & _
                           "        T.DataPedido,  " & vbCrLf & _
                           "        T.Produto_Id,  " & vbCrLf & _
                           "        T.Descricao,  " & vbCrLf & _
                           "        T.UnitarioOficialCompra,  " & vbCrLf & _
                           "        T.UnitarioOficial,  " & vbCrLf & _
                           "        case  " & vbCrLf & _
                           "           when Moeda = 1  " & vbCrLf & _
                           "             then ((T.UnitarioOficial * 100) / case when T.UnitarioOficialCompra = 0 then T.UnitarioOficial else T.UnitarioOficialCompra end ) -100  " & vbCrLf & _
                           "             else ((T.UnitarioMoeda * 100)  /  case when T.UnitarioMoedaCompra   = 0 then T.UnitarioMoeda   else T.UnitarioMoedaCompra   end ) -100  " & vbCrLf & _
                           "        end as Margem,   " & vbCrLf & _
                           "        ISNULL(FC.Indice,0) AS Indice,  " & vbCrLf & _
                           "        T.UnitarioMoedaCompra,  " & vbCrLf & _
                           "        T.UnitarioMoeda,  " & vbCrLf & _
                           "        T.Quantidade,  " & vbCrLf & _
                           "        T.TotalOficial,  " & vbCrLf & _
                           "        T.TotalMoeda,  " & vbCrLf & _
                           "        case" & vbCrLf & _
                           "          when Moeda = 1" & vbCrLf & _
                           "            then ISNULL(((T.TotalOficial * FC.Indice) / 100),0)" & vbCrLf & _
                           "            else 0" & vbCrLf & _
                           "        end as ComissaoOficial," & vbCrLf & _
                           "        case" & vbCrLf & _
                           "          when Moeda <> 1" & vbCrLf & _
                           "            then ISNULL(((T.TotalMoeda * FC.Indice) / 100),0)" & vbCrLf & _
                           "            else 0" & vbCrLf & _
                           "        end as ComissaoMoeda" & vbCrLf & _
                           "   from #Temp T  " & vbCrLf & _
                           "  Inner Join Clientes as Rep  " & vbCrLf & _
                           "     on Rep.Cliente_Id  = T.Representante_Id   " & vbCrLf & _
                           "    and Rep.Endereco_id = T.EndRepresentante_Id  " & vbCrLf & _
                           "  Inner Join Clientes Cli  " & vbCrLf & _
                           "     on Cli.Cliente_Id  = T.Cliente   " & vbCrLf & _
                           "    and Cli.Endereco_id = T.EndCliente  " & vbCrLf & _
                           "   LEFT Join FaixaDeComissao FC  " & vbCrLf & _
                           "     on FC.Tabela_Id = T.Tabela_Id  " & vbCrLf & _
                           "    and case  " & vbCrLf & _
                           "           when Moeda = 1  " & vbCrLf & _
                           "             then ((T.UnitarioOficial * 100) / case when T.UnitarioOficialCompra = 0 then T.UnitarioOficial else T.UnitarioOficialCompra end ) -100  " & vbCrLf & _
                           "             else ((T.UnitarioMoeda * 100)  /  case when T.UnitarioMoedaCompra   = 0 then T.UnitarioMoeda   else T.UnitarioMoedaCompra   end ) -100  " & vbCrLf & _
                           "        end between FC.FaixaInicial_Id and FC.FaixaFinal  " & vbCrLf & _
                           "  Inner Join #PercRep " & vbCrLf & _
                           "     on T.Representante_Id    = #PercRep.Representante_Id" & vbCrLf & _
                           "    and T.EndRepresentante_Id = #PercRep.EndRepresentante_Id" & vbCrLf

                    If rbRepresentante.Checked Then
                        SQL &= "  order by  T.Representante_Id, T.EndRepresentante_Id, Cli.Nome, T.Cliente, T.EndCliente, T.Pedido_Id,  T.Descricao  " & vbCrLf
                    ElseIf rbProduto.Checked Then
                        SQL &= "  order by  T.Descricao, T.Representante_Id, T.EndRepresentante_Id, Cli.Nome, T.Cliente, T.EndCliente, T.Pedido_Id   " & vbCrLf
                    ElseIf rdPorFaixa.Checked Then
                        SQL &= "  order by  ISNULL(FC.Indice,0),T.Representante_Id, T.EndRepresentante_Id, Cli.Nome, T.Cliente, T.EndCliente, T.Pedido_Id,  T.Descricao  " & vbCrLf
                    End If
                    '"  order by  T.Representante_Id, T.EndRepresentante_Id, #PercRep.PercRep, Cli.Nome, T.Cliente, T.EndCliente, T.Pedido_Id, T.DataPedido, T.Descricao  " & vbCrLf
                    ds = Banco.ConsultaDataSet(SQL, "Comissoes")

                    Dim parameters As New Dictionary(Of String, Object)()

                    parameters.Add("Parametros", Parametros)
                    parameters.Add("Nome", HttpContext.Current.Session("ssNomeEmpresa"))
                    parameters.Add("Cidade", HttpContext.Current.Session("ssCidadeEmpresa") & " - " & HttpContext.Current.Session("ssEstadoEmpresa"))
                    parameters.Add("Zerado", "Comissão R$")
                    parameters.Add("Titulo", "Relatório de Comissões " & IIf(rbRepresentante.Checked, "Por Representante", "Por Produto"))
                    parameters.Add("Pagina", "Página")
                    parameters.Add("Data", "Emissão")
                    parameters.Add("Numero", "Cliente")
                    parameters.Add("Romaneio", "Indice")
                    parameters.Add("Produto", "Descrição")
                    parameters.Add("Origem", "Compra U$")
                    parameters.Add("Destino", "Pedido")
                    parameters.Add("Placa", "Representante")
                    parameters.Add("Transportador", "Compra R$")
                    parameters.Add("Motorista", "Venda R$")
                    parameters.Add("Quantidade", "Margem")
                    parameters.Add("Valor/Ton", "Venda U$")
                    parameters.Add("ValorFrete", "Quantidade")
                    parameters.Add("Adiantamento", "Total R$")
                    parameters.Add("Saldo", "Total U$")
                    parameters.Add("ComissaoMoeda", "Comissão U$")
                    parameters.Add("chkCompraReais", chkCompraReais.Checked)
                    parameters.Add("chkVendaReais", chkVendaReais.Checked)
                    parameters.Add("chkMargem", chkMargem.Checked)
                    parameters.Add("chkIndice", chkIndice.Checked)
                    parameters.Add("chkCompraDolar", chkCompraDolar.Checked)
                    parameters.Add("chkVendaDolar", chkVendaDolar.Checked)
                    parameters.Add("chkQuantidade", chkQuantidade.Checked)
                    parameters.Add("chkTotalReais", chkTotalReais.Checked)
                    parameters.Add("chkTotalDolar", chkTotalDolar.Checked)
                    parameters.Add("chkComissaoReais", chkComissaoReais.Checked)
                    parameters.Add("chkComissaoDolar", chkComissaoDolar.Checked)

                    Funcoes.BindReport(Me.Page, ds, IIf(rbRepresentante.Checked Or rdPorFaixa.Checked, "Cr_RelDeComissoes", "Cr_RelDeComissoesPorProduto"), IIf(Pdf, eExportType.PDF, eExportType.ExcelCrystal), parameters)
                End If
            Else
                MsgBox(Me.Page, "Usuario sem permissao para emitir relatório.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkPdf_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkPdf.Click
        Try
            EmitirRelatorio(True)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkExcel_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkExcel.Click
        Try
            EmitirRelatorio(False)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkLimpar_Click(sender As Object, e As EventArgs) Handles lnkLimpar.Click
        Try
            txtRepresentante.Text = ""
            txtCodigoRepresentante.Value = ""
            txtClientes.Text = ""
            txtCodigoCliente.Value = ""
            ddlSafra.SelectedIndex = 0
            ddlMoeda.SelectedIndex = 0
            ddlTroca.SelectedIndex = 0

            LiberaEmpresa()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub LiberaEmpresa()
        If Not UsuarioServidor.LiberaEmpresa Then
            ddlEmpresa.Enabled = False
        End If
    End Sub

    Protected Sub ddlMarca_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddlMarca.SelectedIndexChanged
        Try
            If ddlMarca.SelectedIndex = 0 Then
                ucSelecaoProduto.WhereProduto = ""
            Else
                ucSelecaoProduto.WhereProduto = "marca = " & ddlMarca.SelectedValue
            End If
            ucSelecaoProduto.CarregarNivel(1)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "RelatorioDeComissoes")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

End Class