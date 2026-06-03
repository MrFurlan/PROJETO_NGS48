Imports System.Data
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class CorrecaoMonetaria
    Inherits BasePage

    Dim Sql As String
    Dim SqlArray As New ArrayList
    Dim Mensagem As String
    Dim Cliente As String
    Dim campo() As String


    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Contabil)
        If Not IsPostBack And IsConnect Then
            'If Funcoes.VerificaPermissao("CorrecaoMonetaria", "ACESSAR") Then
            ddl.Carregar(ddlAno, CarregarDDL.Tabela.Ano, (Today.Year - 1) & ";2;C")
                ddl.Carregar(ddlAnoEncerrado, CarregarDDL.Tabela.Ano, (Today.Year - 1) & ";2;C")

                ddlMes.SelectedValue = Today.AddMonths(-1).Month.ToString("00")
                ddlMesEncerrado.SelectedValue = Today.AddMonths(-1).Month.ToString("00")

                ddlAno.SelectedValue = Today.AddMonths(-1).Year
                ddlAnoEncerrado.SelectedValue = Today.AddMonths(-1).Year

            ddl.Carregar(ddlEmpresa, CarregarDDL.Tabela.Empresas, "")
            Funcoes.VerificaEmpresa(ddlEmpresa)
            ddl.Carregar(ddlIndexador, CarregarDDL.Tabela.Indexador, "", False)
                ddl.Carregar(ddlSafra, CarregarDDL.Tabela.Safra, "")
                txtIndice.Text = New [Lib].Negocio.Cotacao(ddlIndexador.SelectedValue, Today).Indice
                HID.Value = Guid.NewGuid().ToString
                ucConsultaClientes.SetarHID(HID.Value)
                ddlIndexador.Focus()
                LiberaEmpresa()
            'Else
            '    MsgBox(Me.Page, "usuário sem permissão para acessar a página.", "~/contabil.aspx")
            '    Exit Sub
            'End If
        End If
    End Sub

    Public Overrides Sub Carregar(obj As IBaseEntity)
        If Not Session("objClienteVar" & HID.Value) Is Nothing Then
            Dim Cli As [Lib].Negocio.Cliente = CType(Session("objClienteVar" & HID.Value), [Lib].Negocio.Cliente)
            With Cli
                txtNomeCliente.Text = .Nome & " - " & RTrim(.Cidade) & "/" & .CodigoEstado & " - " & .CodigoFormatado
                txtCodigoCliente.Value = .Codigo & " - " & .CodigoEndereco
            End With
            Session.Remove("objClienteVar" & HID.Value)
            limparGrids()
        End If
    End Sub

    Public Sub CarregarGridPedidos()
        Dim sql As String

        If FinanceiroNovo Then
            sql = "   Select sb.Empresa, sb.EndEmpresa, Emp.Fantasia, P.Cliente, P.EndCliente, C.Nome, sb.Pedido, P.Safra" & vbCrLf & _
                  "      from(" & vbCrLf & _
                  "           Select distinct Empresa, EndEmpresa, Pedido, clifor, endclifor" & vbCrLf & _
                  "   	        from Titulos" & vbCrLf & _
                  "   	       where situacao  = 1" & vbCrLf & _
                  "   	         and RecPag    ='R'" & vbCrLf & _
                  "              and Moeda     = 3" & vbCrLf & _
                  "              and case" & vbCrLf & _
                  "		               when provisao = 1 and DataBaixa >= '" & CDate("01-" & ddlMes.SelectedValue & "-" & ddlAno.SelectedValue).AddMonths(1).ToSqlDate & "'" & vbCrLf & _
                  "				         then 2" & vbCrLf & _
                  "				         else provisao" & vbCrLf & _
                  "			         end <> 1" & vbCrLf & _
                  "   	       union all" & vbCrLf & _
                  "   	      Select distinct Empresa, EndEmpresa, Pedido, clifor, endclifor" & vbCrLf & _
                  "   	        from Titulos" & vbCrLf & _
                  "   	       where situacao  = 1" & vbCrLf & _
                  "   	         and RecPag    = 'P'" & vbCrLf & _
                  "              and Moeda     = 3" & vbCrLf & _
                  "              and case" & vbCrLf & _
                  "		               when provisao = 1 and DataBaixa >= '" & CDate("01-" & ddlMes.SelectedValue & "-" & ddlAno.SelectedValue).AddMonths(1).ToSqlDate & "'" & vbCrLf & _
                  "				         then 2" & vbCrLf & _
                  "				         else provisao" & vbCrLf & _
                  "			         end <> 1" & vbCrLf & _
                  "           )sb" & vbCrLf & _
                  "     Inner join Pedidos P" & vbCrLf & _
                  "        on P.Empresa_id    = sb.Empresa" & vbCrLf & _
                  "       and P.EndEmpresa_id = sb.EndEmpresa" & vbCrLf & _
                  "       and P.Pedido_id     = sb.Pedido" & vbCrLf & _
                  "       and P.Cliente       = sb.clifor" & vbCrLf & _
                  "       and P.EndCliente    = sb.endclifor" & vbCrLf & _
                  "     Inner Join Clientes Emp" & vbCrLf & _
                  "        on sb.Empresa    = Emp.Cliente_id" & vbCrLf & _
                  "       and sb.EndEmpresa = Emp.Endereco_id" & vbCrLf & _
                  "     Inner Join Clientes C" & vbCrLf & _
                  "        on P.Cliente    = C.Cliente_id" & vbCrLf & _
                  "       and P.EndCliente = C.Endereco_id" & vbCrLf & _
                  "     where 1=1" & vbCrLf
        Else
            sql = " Select sb.Empresa, sb.EndEmpresa, Emp.Fantasia, P.Cliente, P.EndCliente, C.Nome, sb.Pedido, P.Safra" & vbCrLf & _
                  "   from(" & vbCrLf & _
                  "        Select distinct Empresa, EndEmpresa, Pedido, Cliente, endcliente" & vbCrLf & _
                  "	         from contasareceber" & vbCrLf & _
                  "	        where situacao  = 1" & vbCrLf & _
                  "           and Moeda     = 3" & vbCrLf & _
                  "           and case" & vbCrLf & _
                  "		            when provisao = 1 and Baixa >= '" & CDate("01-" & ddlMes.SelectedValue & "-" & ddlAno.SelectedValue).AddMonths(1).ToSqlDate & "'" & vbCrLf & _
                  "			          then 2" & vbCrLf & _
                  "				      else provisao" & vbCrLf & _
                  "			      end <> 1" & vbCrLf & _
                  "	        union all" & vbCrLf & _
                  "	       Select distinct Empresa, EndEmpresa, Pedido, Cliente, endcliente" & vbCrLf & _
                  "	         from contasapagar" & vbCrLf & _
                  "	        where situacao  = 1" & vbCrLf & _
                  "           and Moeda     = 3" & vbCrLf & _
                  "           and case" & vbCrLf & _
                  "		            when provisao = 1 and Baixa >= '" & CDate("01-" & ddlMes.SelectedValue & "-" & ddlAno.SelectedValue).AddMonths(1).ToSqlDate & "'" & vbCrLf & _
                  "			          then 2" & vbCrLf & _
                  "				      else provisao" & vbCrLf & _
                  "			      end <> 1" & vbCrLf & _
                  "         union all" & vbCrLf & _
                  "        Select distinct Empresa_id, EndEmpresa_id,Pedido_id,cliente,endcliente" & vbCrLf & _
                  "          from VW_TituloVirtual" & vbCrLf & _
                  "         where classificacao = 'M'" & vbCrLf & _
                  "       )sb" & vbCrLf & _
                  "  Inner join Pedidos P" & vbCrLf & _
                  "     on P.Empresa_id    = sb.Empresa" & vbCrLf & _
                  "    and P.EndEmpresa_id = sb.EndEmpresa" & vbCrLf & _
                  "    and P.Pedido_id     = sb.Pedido" & vbCrLf & _
                  "    and P.Cliente       = sb.cliente" & vbCrLf & _
                  "    and P.EndCliente    = sb.endcliente" & vbCrLf & _
                  "  Inner Join Clientes Emp" & vbCrLf & _
                  "     on sb.Empresa    = Emp.Cliente_id" & vbCrLf & _
                  "    and sb.EndEmpresa = Emp.Endereco_id" & vbCrLf & _
                  "  Inner Join Clientes C" & vbCrLf & _
                  "     on P.Cliente    = C.Cliente_id" & vbCrLf & _
                  "    and P.EndCliente = C.Endereco_id" & vbCrLf & _
                  "  where 1=1" & vbCrLf
        End If


        If ddlEmpresa.SelectedIndex > 0 Then
            Dim emp As String() = ddlEmpresa.SelectedValue.Split("-")

            If ChkConsEmpresa.Checked Then
                sql &= "    and left(sb.Empresa,8) ='" & Left(emp(0), 8) & "'" & vbCrLf
            Else
                sql &= "    and sb.Empresa    ='" & emp(0) & "'" & vbCrLf & _
                       "    and Sb.EndEmpresa = " & emp(1) & vbCrLf
            End If
        End If

        If ddlSafra.SelectedIndex > 0 Then
            sql &= "    and P.Safra ='" & ddlSafra.SelectedValue & "'" & vbCrLf
        End If

        If txtNomeCliente.Text.Trim.Length > 0 Then
            Dim cli As String() = txtCodigoCliente.Value.Split("-")
            If chkConsolidarCliente.Checked Then
                sql &= "    and left(P.Cliente,8)  = '" & cli(0).Substring(0, 8) & "'" & vbCrLf
            Else
                sql &= "    and P.Cliente    ='" & cli(0) & "'" & vbCrLf & _
                       "    and P.EndCliente = " & cli(1) & vbCrLf
            End If
        End If

        sql &= "    and Exists (select 1" & vbCrLf & _
               "                  from Notasfiscais NF" & vbCrLf & _
               "                 where NF.situacao                  = 1" & vbCrLf & _
               "                   and isnull(NF.TipoDeDocumento,1) = 1" & vbCrLf & _
               "                   and NF.Pedido                    = P.Pedido_Id" & vbCrLf & _
               "                   and NF.Movimento                 <'" & CDate("01-" & ddlMes.SelectedValue & "-" & ddlAno.SelectedValue).AddMonths(1).ToSqlDate & "')" & vbCrLf

        Dim ds As DataSet = Banco.ConsultaDataSet(sql, "Pedidos")
        If ds.Tables IsNot Nothing AndAlso ds.Tables(0).Rows.Count > 0 Then
            gridPedidos.DataSource = ds.Tables(0)
            gridPedidos.DataBind()
        Else
            MsgBox(Me.Page, "Nenhum resultado encontrado.")
        End If

    End Sub

    Protected Sub btnEmitente_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            ucConsultaClientes.Limpar()
            Popup.ConsultaDeClientes(Me.Page, "objClienteVar" & HID.Value, "")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub imgExtratoPedido_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)
        Try
            Dim img As Image = CType(sender, Image)
            Dim row As GridViewRow = CType(img.NamingContainer, GridViewRow)
            Dim i As Integer = row.RowIndex

            Extrato.Emitir(Me.Page, FinanceiroNovo, gridPedidos.Rows(i).Cells(1).Text, gridPedidos.Rows(i).Cells(2).Text, "T", gridPedidos.Rows(i).Cells(7).Text)

        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Public Function getsqlEncerrado(ByVal strPedidos As String, ByVal tipo As String) As String
        Dim sql As String = ""
        Dim UltimoDiaDoMesSql As String = "'" & CDate("01-" & ddlMes.SelectedValue & "-" & ddlAno.SelectedValue).AddMonths(1).AddDays(-1).ToSqlDate & "'"
        Dim PrimeiroDiaDoMesSql As String = "'" & CDate("01-" & ddlMes.SelectedValue & "-" & ddlAno.SelectedValue).ToSqlDate & "'"

        If tipo = "CONTABILIZACAO" Then
            sql &= "Delete razao" & vbCrLf & _
                   "  from razao R" & vbCrLf & _
                   " where R.lote_id       = 66 " & vbCrLf & _
                   "   and R.movimento_id  =" & UltimoDiaDoMesSql & ";" & vbCrLf
        End If

        sql &= "Select P.Unidade," & vbCrLf & _
               "       P.Empresa_Id    as Empresa ," & vbCrLf & _
               "       P.EndEmpresa_Id as EndEmpresa," & vbCrLf & _
               "	   P.Pedido_Id     as Pedido," & vbCrLf & _
               "	   p.classe," & vbCrLf & _
               "	   P.cliente," & vbCrLf & _
               "	   P.EndCliente," & vbCrLf & _
               "	   P.Nome," & vbCrLf & _
               "       cxe.ContaVariacaoMonetariaPassiva, cxe.ContaVariacaoMonetariaAtiva," & vbCrLf & _
               "       cxe.ContaVariacaoMonetariaCliente, cxe.ContaVariacaoMonetariaFornecedor," & vbCrLf & _
               "	   Sum(case" & vbCrLf & _
               "		     When r.Conta_Id = cxe.ContaVariacaoMonetariaPassiva" & vbCrLf & _
               "		       then Debitooficial - creditooficial" & vbCrLf & _
               "			   else 0" & vbCrLf & _
               "           end) as Passiva," & vbCrLf & _
               "	   Sum(case" & vbCrLf & _
               "		     When r.Conta_Id = cxe.ContaVariacaoMonetariaAtiva" & vbCrLf & _
               "		       then creditooficial - Debitooficial" & vbCrLf & _
               "			   else 0" & vbCrLf & _
               "           end) as Ativa" & vbCrLf & _
               "  into #Temp2" & vbCrLf & _
               "  from (Select p.UnidadeDeNegocio as Unidade, cp.EmpresaPedido as Empresa_Id, cp.EndEmpresaPedido as EndEmpresa_Id, cp.Pedido as Pedido_Id, op.Classe, p.cliente, p.EndCliente, cli.Nome" & vbCrLf & _
               "		  from Contasapagar cp" & vbCrLf & _
               "		 inner join Moedas M" & vbCrLf & _
               "			on cp.Moeda = M.Moeda_Id" & vbCrLf & _
               "         Inner join Pedidos P" & vbCrLf & _
               "		    on Cp.EmpresaPedido    = P.Empresa_Id" & vbCrLf & _
               "		   and Cp.EndEmpresaPedido = P.EndEmpresa_Id" & vbCrLf & _
               "		   and cp.Pedido           = P.Pedido_Id" & vbCrLf & _
               "		 Inner Join Operacoes Op" & vbCrLf & _
               "		    on P.Operacao = OP.Operacao_Id" & vbCrLf & _
               "         Inner join Clientes cli" & vbCrLf & _
               "		    on cli.Cliente_Id  = p.Cliente" & vbCrLf & _
               "		   and cli.Endereco_Id = p.EndCliente" & vbCrLf & _
               "		 Where isnull(p.Troca,0) = 0" & vbCrLf & _
               "		   and M.Classificacao = 'M'" & vbCrLf & _
               "		   and cp.Provisao     = 1" & vbCrLf & _
               "		   and cp.situacao = 1" & vbCrLf & _
               "		   and cp.baixa between " & PrimeiroDiaDoMesSql & " and " & UltimoDiaDoMesSql & vbCrLf & _
               "		   and not exists(select 1 from ContasAPagar   where Situacao = 1 and Provisao <> 1 and Pedido = cp.Pedido" & vbCrLf & _
               "						   union" & vbCrLf & _
               "						  select 1 from ContasAReceber where Situacao = 1 and Provisao <> 1 and Pedido = cp.Pedido)" & vbCrLf & _
               "		 union " & vbCrLf & _
               "		select p.UnidadeDeNegocio, cp.EmpresaPedido, cp.EndEmpresaPedido, cp.Pedido, op.Classe, p.cliente, p.EndCliente, cli.Nome" & vbCrLf & _
               "		  from ContasAReceber cp" & vbCrLf & _
               "		 inner join Moedas M" & vbCrLf & _
               "			on cp.Moeda = M.Moeda_Id" & vbCrLf & _
               "         Inner join Pedidos P" & vbCrLf & _
               "		    on Cp.EmpresaPedido    = P.Empresa_Id" & vbCrLf & _
               "		   and Cp.EndEmpresaPedido = P.EndEmpresa_Id" & vbCrLf & _
               "		   and cp.Pedido           = P.Pedido_Id" & vbCrLf & _
               "         Inner Join Operacoes Op" & vbCrLf & _
               "		    on P.Operacao = OP.Operacao_Id " & vbCrLf & _
               "		 Inner join Clientes cli" & vbCrLf & _
               "		    on cli.Cliente_Id  = p.Cliente" & vbCrLf & _
               "		   and cli.Endereco_Id = p.EndCliente" & vbCrLf & _
               "		 Where isnull(p.Troca,0) = 0" & vbCrLf & _
               "		   and M.Classificacao = 'M'" & vbCrLf & _
               "		   and cp.Provisao     = 1" & vbCrLf & _
               "		   and cp.situacao = 1" & vbCrLf & _
               "		   and cp.baixa between " & PrimeiroDiaDoMesSql & " and " & UltimoDiaDoMesSql & vbCrLf & _
               "		   and not exists(select 1 from ContasAPagar   where Situacao = 1 and Provisao <> 1 and Pedido = cp.Pedido" & vbCrLf & _
               "						   union" & vbCrLf & _
               "						  select 1 from ContasAReceber where Situacao = 1 and Provisao <> 1 and Pedido = cp.Pedido)" & vbCrLf & _
               "		) p" & vbCrLf & _
               " inner join razao r" & vbCrLf & _
               "    on r.Empresa_Id    = p.empresa_Id" & vbCrLf & _
               "   and r.EndEmpresa_id = p.endEmpresa_Id" & vbCrLf & _
               "   and r.pedido        = p.pedido_Id" & vbCrLf & _
               " Inner join clientesxEmpresas cxe" & vbCrLf & _
               "    on cxe.Empresa_id    = p.empresa_Id" & vbCrLf & _
               "   and cxe.endEmpresa_id = p.EndEmpresa_Id" & vbCrLf & _
               " where r.lote_id in (6,66)" & vbCrLf & _
               "   and r.conta_id in (cxe.ContaVariacaoMonetariaPassiva, cxe.ContaVariacaoMonetariaAtiva)" & vbCrLf & _
               IIf(strPedidos.Length > 0, "   and (" & strPedidos & ")", "") & vbCrLf & _
               " Group by P.Unidade," & vbCrLf & _
               "          P.Empresa_Id," & vbCrLf & _
               "          P.EndEmpresa_Id," & vbCrLf & _
               "	      P.Pedido_id," & vbCrLf & _
               "	      p.classe," & vbCrLf & _
               "	      P.cliente," & vbCrLf & _
               "	      P.EndCliente," & vbCrLf & _
               "		  P.Nome," & vbCrLf & _
               "          cxe.ContaVariacaoMonetariaPassiva, cxe.ContaVariacaoMonetariaAtiva," & vbCrLf & _
               "          cxe.ContaVariacaoMonetariaCliente, cxe.ContaVariacaoMonetariaFornecedor;" & vbCrLf

        If tipo.Contains("VISUALIZACAO") Or tipo.Contains("CONTABILIZACAO") Then
            sql &= "Select sb.tipo," & vbCrLf & _
                   "       t.Unidade," & vbCrLf & _
                   "       t.Empresa," & vbCrLf & _
                   "       t.EndEmpresa," & vbCrLf & _
                   "	   t.Pedido," & vbCrLf & _
                   "	   t.classe," & vbCrLf & _
                   "	   t.cliente," & vbCrLf & _
                   "	   t.EndCliente," & vbCrLf & _
                   "	   t.Nome," & vbCrLf & _
                   "       t.ContaVariacaoMonetariaPassiva, t.ContaVariacaoMonetariaAtiva," & vbCrLf & _
                   "       t.ContaVariacaoMonetariaCliente, t.ContaVariacaoMonetariaFornecedor," & vbCrLf & _
                   "       t.passiva - t.ativa as saldo," & vbCrLf & _
                   "	   case" & vbCrLf & _
                   "	      when t.passiva - t.ativa < 0" & vbCrLf & _
                   "		    then t.ContaVariacaoMonetariaPassiva" & vbCrLf & _
                   "			else t.ContaVariacaoMonetariaAtiva" & vbCrLf & _
                   "	   end contaVariacao," & vbCrLf & _
                   "	   case" & vbCrLf & _
                   "	      when t.passiva - t.ativa < 0" & vbCrLf & _
                   "		    then 'D'" & vbCrLf & _
                   "			else 'C'" & vbCrLf & _
                   "	   end VDC," & vbCrLf & _
                   "	   case" & vbCrLf & _
                   "	      when t.classe = '" & eClassesOperacoes.VENDAS.ToString & "'" & vbCrLf & _
                   "		    then c.ContaVariacaoMonetariaCliente" & vbCrLf & _
                   "			else c.ContaVariacaoMonetariaFornecedor" & vbCrLf & _
                   "	   end contaCLIFOR," & vbCrLf & _
                   "	   case" & vbCrLf & _
                   "	      when t.passiva - t.ativa < 0 " & vbCrLf & _
                   "		    then 'C' " & vbCrLf & _
                   "			else 'D' " & vbCrLf & _
                   "	   end CFDC" & vbCrLf & _
                   "  into #Cont" & vbCrLf & _
                   "  from #temp2 t" & vbCrLf & _
                   " inner join(select 'VA' as Tipo" & vbCrLf & _
                   "             union" & vbCrLf & _
                   "			select 'CF'" & vbCrLf & _
                   "           ) sb" & vbCrLf & _
                   "    on 1 = 1" & vbCrLf & _
                   " inner join ClientesXEmpresas c" & vbCrLf & _
                   "    on left(c.Empresa_Id,8) = left(t.empresa,8)" & vbCrLf & _
                   "   and c.Matriz = 'S'" & vbCrLf & _
                   " where t.passiva <> t.ativa;" & vbCrLf
        End If


            If tipo = "VISUALIZACAO" Then
                sql &= "Select Empresa, EndEmpresa,Pedido, Classe, Cliente, endCliente, Nome, Passiva, Ativa" & vbCrLf & _
                       "  from #Temp2" & vbCrLf
            ElseIf tipo = "VISUALIZACAOCONTABILIZACAO" Then
                sql &= "select Pedido," & vbCrLf & _
                       "       case when tipo = 'CF' then contaCliFor else contaVariacao  end as Conta," & vbCrLf & _
                       "	   case when tipo = 'CF' then cliente     else ''             end as Cliente," & vbCrLf & _
                       "	   case when tipo = 'CF' then Endcliente  else ''             end as EndCliente," & vbCrLf & _
                       "	   Nome,    	" & vbCrLf & _
                       "	   case when (tipo = 'CF' and CFDC = 'D') or (tipo = 'VA' and VDC = 'D') then abs(saldo) else 0 end DebitoOficial," & vbCrLf & _
                       "	   case when (tipo = 'CF' and CFDC = 'C') or (tipo = 'VA' and VDC = 'C') then abs(saldo) else 0 end CreditoOficial," & vbCrLf & _
                       "	   case " & vbCrLf & _
                       "	     when contavariacao = ContaVariacaoMonetariaAtiva " & vbCrLf & _
                       "		   then 'Estorno da Provisão da Variacao Passiva - AJUSTE DE FECHAMENTO DE PEDIDO' + convert(varchar,Pedido) + ' - ' + Nome " & vbCrLf & _
                       "		   else 'Estorno da Provisão da Variacao Ativa   - AJUSTE DE FECHAMENTO DE PEDIDO' + convert(varchar,Pedido) + ' - ' + Nome " & vbCrLf & _
                       "	   end Historico" & vbCrLf & _
                       "from #Cont " & vbCrLf & _
                       "order by Nome, Pedido, case when tipo = 'CF' then contaCliFor else contaVariacao  end " & vbCrLf

            ElseIf tipo = "CONTABILIZACAO" Then
                sql &= "INSERT INTO Razao(Empresa_Id, EndEmpresa_Id, Conta_Id, Cliente_Id, EndCliente_Id, Movimento_Id, Lote_Id, Sequencia_Id, UnidadeDeNegocio,  Pedido,  Indexador, DataMoeda, DebitoOficial, " & vbCrLf & _
                       "                         CreditoOficial, DebitoMoeda, CreditoMoeda, Historico, PrevistoRealizado, DataDaBaixa, Rateado, Processo, Situacao)" & vbCrLf & _
                       "select empresa, " & vbCrLf & _
                       "       EndEmpresa," & vbCrLf & _
                       "       case when tipo = 'CF' then contaCliFor else contaVariacao  end," & vbCrLf & _
                       "	   case when tipo = 'CF' then cliente     else ''             end," & vbCrLf & _
                       "	   case when tipo = 'CF' then Endcliente  else ''             end," & vbCrLf & _
                       "	   " & UltimoDiaDoMesSql & "," & vbCrLf & _
                       "	   66, ROW_NUMBER() OVER(ORDER BY empresa,Pedido DESC) AS Row," & vbCrLf & _
                       "	   Unidade, Pedido, 3, " & UltimoDiaDoMesSql & "," & vbCrLf & _
                       "	   case when (tipo = 'CF' and CFDC = 'D') or (tipo = 'VA' and VDC = 'D') then abs(saldo) else 0 end," & vbCrLf & _
                       "	   case when (tipo = 'CF' and CFDC = 'C') or (tipo = 'VA' and VDC = 'C') then abs(saldo) else 0 end," & vbCrLf & _
                       "	   0,0," & vbCrLf & _
                       "	   case " & vbCrLf & _
                       "	     when contavariacao = ContaVariacaoMonetariaAtiva " & vbCrLf & _
                       "		   then 'AJUSTE DE FECHAMENTO DE PEDIDO' + convert(varchar,Pedido) + ' - ' + Nome + ' Estorno da Provisão da Variacao Passiva'" & vbCrLf & _
                       "		   else 'AJUSTE DE FECHAMENTO DE PEDIDO' + convert(varchar,Pedido) + ' - ' + Nome + ' Estorno da Provisão da Variacao Ativa'" & vbCrLf & _
                       "	   end," & vbCrLf & _
                       "	   'R', " & UltimoDiaDoMesSql & ", 0,'VARIACAO ENCERRAMENTO',1" & vbCrLf & _
                       "from #Cont" & vbCrLf
            End If

            Return sql
    End Function

    Public Function getSql(ByVal strPedidos As String, ByVal tipo As String) As String
        Dim emp As String() = ddlEmpresa.SelectedValue.Split("-")
        Dim sql As String

        Dim PrimeiroDiaDoMesSql As String = "'" & CDate("01-" & ddlMes.SelectedValue & "-" & ddlAno.SelectedValue).ToSqlDate & "'"
        Dim UltimoDiaDoMesSql As String = "'" & CDate("01-" & ddlMes.SelectedValue & "-" & ddlAno.SelectedValue).AddMonths(1).AddDays(-1).ToSqlDate & "'"

        Dim PrimeiroDiaDoMesbr As String = CDate("01-" & ddlMes.SelectedValue & "-" & ddlAno.SelectedValue).ToStrDate
        Dim UltimoDiaDoMesbr As String = CDate("01-" & ddlMes.SelectedValue & "-" & ddlAno.SelectedValue).AddMonths(1).AddDays(-1).ToStrDate


        sql = " declare" & vbCrLf & _
              " @Cotacao decimal(18,8)," & vbCrLf & _
              " @Sequencia integer" & vbCrLf & _
              " set @Cotacao = " & Str(CDec(txtIndice.Text)) & vbCrLf & _
              " (select @Sequencia = isnull(max(Sequencia_id),0) + 1 from razao where " & IIf(ChkConsEmpresa.Checked, "Left(Empresa_Id,8)", "empresa_Id") & " = '" & IIf(ChkConsEmpresa.Checked, Left(emp(0), 8), emp(0)) & "' and lote_id = 6 and movimento_Id =" & UltimoDiaDoMesSql & ")" & vbCrLf & _
              " select P.Empresa_Id as Empresa, P.EndEmpresa_Id as EndEmpresa, P.Pedido_Id as Pedido" & vbCrLf & _
              "   into #Pedidos" & vbCrLf & _
              "   from Pedidos P" & vbCrLf & _
              "  Where " & strPedidos & ";" & vbCrLf

        sql &= "SELECT CASE" & vbCrLf & _
               "         WHEN SO.EntradaSaida = 'S'" & vbCrLf & _
               "           THEN 'R'" & vbCrLf & _
               "           ELSE 'P'" & vbCrLf & _
               "       END  as TipoPR," & vbCrLf & _
               "       P.Empresa_Id as Empresa," & vbCrLf & _
               "       P.EndEmpresa_id as EndEmpresa," & vbCrLf & _
               "       P.Pedido_id as Pedido," & vbCrLf & _
               "       isnull(sb.ValorBaixadoDolar,0) as FinanceiroBaixadoEmDolar," & vbCrLf & _
               "       isnull(sb.ValorFechadoReais,0) as FinanceiroBaixadoEmReais," & vbCrLf & _
               "       sum(case" & vbCrLf & _
               "             When SONota.devolucao = 'S'" & vbCrLf & _
               "               then nxi.quantidadefiscal * Pxi.UnitarioMoeda * -1" & vbCrLf & _
               "               else nxi.quantidadefiscal * Pxi.UnitarioMoeda " & vbCrLf & _
               "           end)  -  isnull(sb.ValorBaixadoDolar,0) as ValorDolarPedido," & vbCrLf & _
               "       (sum(case" & vbCrLf & _
               "             When SONota.devolucao = 'S'" & vbCrLf & _
               "               then nxi.quantidadefiscal * Pxi.UnitarioMoeda * -1" & vbCrLf & _
               "               else nxi.quantidadefiscal * Pxi.UnitarioMoeda " & vbCrLf & _
               "            end) -  isnull(sb.ValorBaixadoDolar,0)) * @Cotacao as ValorEmDolarPedidoAtualizadoEmReais," & vbCrLf & _
               "       sum(case" & vbCrLf & _
               "             When SONota.devolucao = 'S'" & vbCrLf & _
               "               then nxe.Valor * -1" & vbCrLf & _
               "               else nxe.Valor" & vbCrLf & _
               "            end) - isnull(sb.ValorFechadoReais,0) as ReaisNotas" & vbCrLf & _
               "   into #temp" & vbCrLf & _
               "   FROM NotasFiscais AS NF" & vbCrLf & _
               "  INNER JOIN NotasFiscaisXItens Nxi" & vbCrLf & _
               "     ON NF.Empresa_Id      = Nxi.Empresa_Id" & vbCrLf & _
               "    AND NF.EndEmpresa_Id   = Nxi.EndEmpresa_Id" & vbCrLf & _
               "    AND NF.Cliente_Id      = Nxi.Cliente_Id" & vbCrLf & _
               "    AND NF.EndCliente_Id   = Nxi.EndCliente_Id" & vbCrLf & _
               "    AND NF.EntradaSaida_Id = Nxi.EntradaSaida_Id" & vbCrLf & _
               "    AND NF.Serie_Id        = Nxi.Serie_Id" & vbCrLf & _
               "    AND NF.Nota_Id         = Nxi.Nota_Id" & vbCrLf & _
               "  Inner Join Produtos Prd" & vbCrLf & _
               "     ON Prd.Produto_id = nxi.Produto_id" & vbCrLf & _
               "  INNER JOIN Pedidos AS P" & vbCrLf & _
               "     ON NF.Empresa_Id    = P.Empresa_Id" & vbCrLf & _
               "    AND NF.EndEmpresa_Id = P.EndEmpresa_Id" & vbCrLf & _
               "    AND NF.Pedido        = P.Pedido_Id" & vbCrLf & _
               "  INNER JOIN (	select Pxi.Empresa_id," & vbCrLf & _
               " 					   Pxi.EndEmpresa_id," & vbCrLf & _
               " 					   Pxi.Pedido_id," & vbCrLf & _
               " 					   Pxi.Produto_id," & vbCrLf & _
               " 					   case" & vbCrLf & _
               " 						  when PxE.ValorMoeda = 0 or sum(case" & vbCrLf & _
               "                                                           when Pxi.TipoDeLancamento = 'E'" & vbCrLf & _
               "                                                            then Pxi.Quantidade * -1" & vbCrLf & _
               "                                                            else Pxi.Quantidade" & vbCrLf & _
               "                                                        end) <= 0" & vbCrLf & _
               " 							then 0" & vbCrLf & _
               " 							else (PxE.ValorMoeda / sum(case" & vbCrLf & _
               "                                                          when Pxi.TipoDeLancamento = 'E'" & vbCrLf & _
               "                                                            then Pxi.Quantidade * -1" & vbCrLf & _
               "                                                            else Pxi.Quantidade" & vbCrLf & _
               "                                                        end))" & vbCrLf & _
               " 					   end UnitarioMoeda" & vbCrLf & _
               " 				  from PedidoxItemxLancamento Pxi" & vbCrLf & _
               "                   left join(SELECT Empresa_Id," & vbCrLf & _
               "                                    EndEmpresa_id," & vbCrLf & _
               "                                    Pedido_Id," & vbCrLf & _
               "                                    Produto_Id," & vbCrLf & _
               "                                    ValorMoeda" & vbCrLf & _
               "                               FROM PedidosXEncargos" & vbCrLf & _
               "                              Where Encargo_Id = 'LIQUIDO'" & vbCrLf & _
               "                              ) PxE" & vbCrLf & _
               "                    on Pxi.Empresa_id    = PxE.Empresa_id" & vbCrLf & _
               " 				   and Pxi.EndEmpresa_id = PxE.EndEmpresa_id" & vbCrLf & _
               " 				   and Pxi.Pedido_id     = PxE.Pedido_id" & vbCrLf & _
               " 				   and Pxi.Produto_id    = PxE.Produto_id" & vbCrLf & _
               " 				 Inner Join Produtos prd" & vbCrLf & _
               " 					on Pxi.Produto_id = prd.Produto_Id" & vbCrLf & _
               " 				group by Pxi.Empresa_id," & vbCrLf & _
               " 						 Pxi.EndEmpresa_id," & vbCrLf & _
               " 						 Pxi.Pedido_id," & vbCrLf & _
               " 						 Pxi.Produto_id," & vbCrLf & _
               "                         PxE.ValorMoeda" & vbCrLf & _
               "             ) PxI" & vbCrLf & _
               "     on PxI.Empresa_id       = Nxi.Empresa_Id" & vbCrLf & _
               "    AND PxI.EndEmpresa_id    = Nxi.EndEmpresa_Id" & vbCrLf & _
               "    AND PxI.Pedido_id        = Nxi.Pedido" & vbCrLf & _
               "    AND PxI.Produto_id       = Nxi.Produto_id" & vbCrLf & _
               "  INNER JOIN NotasFiscaisXEncargos nxe" & vbCrLf & _
               "     ON Nxi.Empresa_Id      = nxe.Empresa_Id" & vbCrLf & _
               "    AND Nxi.EndEmpresa_Id   = nxe.EndEmpresa_Id" & vbCrLf & _
               "    AND Nxi.Cliente_Id      = nxe.Cliente_Id" & vbCrLf & _
               "    AND Nxi.EndCliente_Id   = nxe.EndCliente_Id" & vbCrLf & _
               "    AND Nxi.EntradaSaida_Id = nxe.EntradaSaida_Id" & vbCrLf & _
               "    AND Nxi.Serie_Id        = nxe.Serie_Id" & vbCrLf & _
               "    AND Nxi.Nota_Id         = nxe.Nota_Id" & vbCrLf & _
               "    AND Nxi.Produto_id      = nxe.Produto_Id" & vbCrLf & _
               "    and nxi.cfop_id         = nxe.cfop_id" & vbCrLf & _
               "    and nxi.sequencia_id    = nxe.sequencia_id" & vbCrLf & _
               "    and nxe.Encargo_id      = 'LIQUIDO'" & vbCrLf & _
               " LEFT Join(Select Pedido," & vbCrLf & _
               "                 Cliente," & vbCrLf & _
               "                 endcliente," & vbCrLf & _
               "                 Moeda," & vbCrLf & _
               "                 abs(sum(case" & vbCrLf & _
               "                           when Tipo = 'R'" & vbCrLf & _
               "                            then ValorBaixadoDolar" & vbCrLf & _
               "                            else ValorBaixadoDolar * -1" & vbCrLf & _
               "                         end)) ValorBaixadoDolar," & vbCrLf & _
               "                 abs(sum(case" & vbCrLf & _
               "                           when Tipo = 'R'" & vbCrLf & _
               "                            then ValorAbertoDolar" & vbCrLf & _
               "                            else ValorAbertoDolar * -1" & vbCrLf & _
               "                         end)) ValorAbertoDolar," & vbCrLf & _
               "                 abs(sum(case" & vbCrLf & _
               "                           when Tipo = 'R'" & vbCrLf & _
               "                            then ValorFechadoReais" & vbCrLf & _
               "                            else ValorFechadoReais * -1" & vbCrLf & _
               "                         end)) ValorFechadoReais" & vbCrLf & _
               "           FROM ( " & vbCrLf & _
               "                 select 'R' as Tipo," & vbCrLf & _
               "                         Pedido," & vbCrLf & _
               "                         cliente," & vbCrLf & _
               "                         endcliente, " & vbCrLf & _
               "                         Moeda," & vbCrLf & _
               "                        sum(case" & vbCrLf & _
               "                               when case" & vbCrLf & _
               "                                       when provisao = 1 and baixa > " & UltimoDiaDoMesSql & vbCrLf & _
               "                                        then 2" & vbCrLf & _
               "                                        else provisao" & vbCrLf & _
               "                                    end = 1" & vbCrLf & _
               "                                 then MoedaValorDoDocumento" & vbCrLf & _
               "                                 else 0" & vbCrLf & _
               "                            end) as ValorBaixadoDolar," & vbCrLf & _
               "                        sum(case" & vbCrLf & _
               "                               when case" & vbCrLf & _
               "                                       when provisao = 1 and baixa > " & UltimoDiaDoMesSql & vbCrLf & _
               "                                        then 2" & vbCrLf & _
               "                                        else provisao" & vbCrLf & _
               "                                    end  <> 1" & vbCrLf & _
               "                                 then MoedaValorDoDocumento" & vbCrLf & _
               "                                 else 0" & vbCrLf & _
               "                            end) as ValorAbertoDolar," & vbCrLf & _
               "                         sum(case" & vbCrLf & _
               "                               when case" & vbCrLf & _
               "                                       when provisao = 1 and baixa > " & UltimoDiaDoMesSql & vbCrLf & _
               "                                        then 2" & vbCrLf & _
               "                                        else provisao" & vbCrLf & _
               "                                    end  = 1" & vbCrLf & _
               "                                 then ValorDoDocumento" & vbCrLf & _
               "                                 else 0" & vbCrLf & _
               "                            end) as ValorFechadoReais" & vbCrLf & _
               " 		     	   from contasareceber" & vbCrLf & _
               " 		     	  where situacao = 1" & vbCrLf & _
               " 		     	  Group by Pedido, cliente, endcliente, Moeda" & vbCrLf & _
               " 		     	  union all" & vbCrLf & _
               " 		     	 select 'P' as Tipo," & vbCrLf & _
               "                         Pedido," & vbCrLf & _
               "                         cliente," & vbCrLf & _
               "                         endcliente," & vbCrLf & _
               "                         Moeda," & vbCrLf & _
               "                         sum(case" & vbCrLf & _
               "                               when case" & vbCrLf & _
               "                                       when provisao = 1 and baixa > " & UltimoDiaDoMesSql & vbCrLf & _
               "                                        then 2" & vbCrLf & _
               "                                        else provisao" & vbCrLf & _
               "                                    end = 1" & vbCrLf & _
               "                                 then MoedaValorDoDocumento" & vbCrLf & _
               "                                 else 0" & vbCrLf & _
               "                            end) as ValorBaixadoDolar," & vbCrLf & _
               "                         sum(case" & vbCrLf & _
               "                               when case" & vbCrLf & _
               "                                       when provisao = 1 and baixa > " & UltimoDiaDoMesSql & vbCrLf & _
               "                                        then 2" & vbCrLf & _
               "                                        else provisao" & vbCrLf & _
               "                                    end  <> 1" & vbCrLf & _
               "                                 then MoedaValorDoDocumento" & vbCrLf & _
               "                                 else 0" & vbCrLf & _
               "                            end) as ValorAbertoDolar," & vbCrLf & _
               "                          sum(case" & vbCrLf & _
               "                               when case" & vbCrLf & _
               "                                       when provisao = 1 and baixa > " & UltimoDiaDoMesSql & vbCrLf & _
               "                                        then 2" & vbCrLf & _
               "                                        else provisao" & vbCrLf & _
               "                                    end  = 1" & vbCrLf & _
               "                                 then ValorDoDocumento" & vbCrLf & _
               "                                 else 0" & vbCrLf & _
               "                            end) as ValorFechadoReais" & vbCrLf & _
               " 		     	   from contasapagar" & vbCrLf & _
               " 		     	  where situacao = 1" & vbCrLf & _
               "                    and len(isnull(CarteiraAdto,'')) = 0" & vbCrLf & _
               " 		     	  Group by Pedido, cliente, endcliente, Moeda" & vbCrLf & _
               "				   union all" & vbCrLf & _
               "				  Select CR," & vbCrLf & _
               "                         Pedido_Id," & vbCrLf & _
               "                         cliente," & vbCrLf & _
               "                         endcliente," & vbCrLf & _
               "                         Moeda," & vbCrLf & _
               "                         0 as ValorBaixadoDolar," & vbCrLf & _
               "                         sum(ParcelaMoeda) as ValorAbertoDolar," & vbCrLf & _
               "                         0 as ValorFechadoReais" & vbCrLf & _
               "				    From VW_TituloVirtual" & vbCrLf & _
               " 		     	  Group by CR, Pedido_Id, cliente, endcliente, Moeda" & vbCrLf & _
               " 				) ResumoFinanceiro" & vbCrLf & _
               " 			group by Pedido," & vbCrLf & _
               "                     Cliente," & vbCrLf & _
               "                     endcliente," & vbCrLf & _
               "                     Moeda" & vbCrLf & _
               "         ) Sb" & vbCrLf & _
               "     ON sb.Pedido     = P.Pedido_id" & vbCrLf & _
               "    And sb.Cliente    = P.Cliente" & vbCrLf & _
               "    And Sb.endCliente = P.EndCliente" & vbCrLf & _
               "    AND sb.moeda      = P.moeda" & vbCrLf & _
               "  INNER JOIN SubOperacoes AS SO" & vbCrLf & _
               "     ON P.Operacao    = SO.Operacao_Id" & vbCrLf & _
               "    AND P.SubOperacao = SO.SubOperacoes_Id" & vbCrLf & _
               "  INNER JOIN SubOperacoes AS SONota" & vbCrLf & _
               "     ON NF.Operacao    = SONota.Operacao_Id" & vbCrLf & _
               "    AND NF.SubOperacao = SONota.SubOperacoes_Id" & vbCrLf & _
               "   JOIN #Pedidos Pd" & vbCrLf & _
               "     ON P.Pedido_Id = Pd.Pedido" & vbCrLf & _
               "    AND P.Empresa_Id = pd.Empresa" & vbCrLf & _
               "    AND P.EndEmpresa_Id = Pd.EndEmpresa" & vbCrLf & _
               "  WHERE (NF.Situacao  = 1)" & vbCrLf & _
               "    And NF.Movimento <=" & UltimoDiaDoMesSql & vbCrLf & _
               "  Group By CASE" & vbCrLf & _
               "             WHEN SO.EntradaSaida = 'S'" & vbCrLf & _
               "               THEN 'R'" & vbCrLf & _
               "               ELSE 'P'" & vbCrLf & _
               "           END," & vbCrLf & _
               "           P.Empresa_Id," & vbCrLf & _
               "           P.EndEmpresa_id, " & vbCrLf & _
               "           P.Pedido_id, " & vbCrLf & _
               "           isnull(sb.ValorBaixadoDolar,0)," & vbCrLf & _
               "           isnull(sb.ValorFechadoReais,0);" & vbCrLf

        sql &= " select t.FinanceiroBaixadoEmDolar," & vbCrLf & _
               "        t.FinanceiroBaixadoEmReais," & vbCrLf & _
               "        t.ValorDolarPedido," & vbCrLf & _
               "        t.ValorEmDolarPedidoAtualizadoEmReais," & vbCrLf & _
               "        t.ReaisNotas," & vbCrLf & _
               "        t.TipoPR," & vbCrLf & _
               "        case when t.TipoPR = 'R' then 'Recebiveis' else 'Pagaveis' end DescTipo," & vbCrLf & _
               "        t.Empresa," & vbCrLf & _
               "        T.EndEmpresa," & vbCrLf & _
               "        Emp.fantasia, p.cliente, p.endcliente, c.nome," & vbCrLf & _
               "        t.Pedido," & vbCrLf & _
               "        case when (t.ValorEmDolarPedidoAtualizadoEmReais > t.ReaisNotas and t.TipoPR = 'R') or (t.ValorEmDolarPedidoAtualizadoEmReais < t.ReaisNotas and t.TipoPR = 'P')" & vbCrLf & _
               "           	then 'A'" & vbCrLf & _
               "           	else 'P'" & vbCrLf & _
               "        end as TipoVariacao," & vbCrLf & _
               "        (t.ValorEmDolarPedidoAtualizadoEmReais - t.ReaisNotas) * case when (t.ValorEmDolarPedidoAtualizadoEmReais > t.ReaisNotas and t.TipoPR = 'R') or (t.ValorEmDolarPedidoAtualizadoEmReais < t.ReaisNotas and t.TipoPR = 'P')" & vbCrLf & _
               "                                                                    then 1" & vbCrLf & _
               "                                                                    else -1" & vbCrLf & _
               "                                                                end as VariacaoCambial," & vbCrLf & _
               "        isnull(R.VariacaoPassiva,0) as VariacaoPassivaCorrecao," & vbCrLf & _
               "        isnull(R.VariacaoAtiva,0) as VariacaoAtivaCorrecao," & vbCrLf & _
              "        ((t.ValorEmDolarPedidoAtualizadoEmReais - t.ReaisNotas) * case" & vbCrLf & _
               "		                                                            when (t.ValorEmDolarPedidoAtualizadoEmReais > t.ReaisNotas and t.TipoPR = 'R') or (t.ValorEmDolarPedidoAtualizadoEmReais < t.ReaisNotas and t.TipoPR = 'P')" & vbCrLf & _
               "                                                                      then 1" & vbCrLf & _
               "                                                                      else -1" & vbCrLf & _
               "                                                                  end) -  (isnull(R.VariacaoAtiva,0) - isnull(R.VariacaoPassiva,0))	as VariacaoCorrigida" & vbCrLf

        If tipo = "CONTABILIZACAO" OrElse tipo = "VCONTABILIZACAO" Then
            sql &= "   into #Razao"
        End If

        sql &= "   from #temp t" & vbCrLf & _
               "   left Join (" & vbCrLf & _
               "               select r.Pedido," & vbCrLf & _
               "                      sum(isnull(r.DebitoOficial,0))  as VariacaoPassiva," & vbCrLf & _
               "                      sum(isnull(r.CreditoOficial,0)) as VariacaoAtiva" & vbCrLf & _
               "                 from razao r" & vbCrLf & _
               "                inner join ClientesXEmpresas CxE" & vbCrLf & _
               "                   on r.Empresa_id    = CxE.Empresa_id" & vbCrLf & _
               "                  and r.EndEmpresa_id = CxE.EndEmpresa_Id" & vbCrLf & _
               "                  and(r.Conta_id = CxE.ContaVariacaoMonetariaPassiva or r.Conta_id = CxE.ContaVariacaoMonetariaAtiva)" & vbCrLf & _
               "                where lote_id in (2, 6) " & vbCrLf & _
               "                group by Pedido" & vbCrLf & _
               "              ) R" & vbCrLf & _
               "      on t.Pedido = R.Pedido" & vbCrLf & _
               "  Inner join Pedidos P" & vbCrLf & _
               "     on P.Empresa_id    = t.Empresa" & vbCrLf & _
               "    and P.EndEmpresa_id = t.EndEmpresa" & vbCrLf & _
               "    and P.Pedido_id     = t.Pedido" & vbCrLf & _
               "  Inner Join Clientes Emp" & vbCrLf & _
               "     on t.Empresa    = Emp.Cliente_id" & vbCrLf & _
               "    and t.EndEmpresa = Emp.Endereco_id" & vbCrLf & _
               "  Inner Join Clientes C" & vbCrLf & _
               "     on p.Cliente    = C.Cliente_id" & vbCrLf & _
               "    and p.EndCliente = C.Endereco_id" & vbCrLf
        '"" & vbCrLf & _
        '"update #temp set VariacaoCorrigida = VariacaoCorrigida *-1 where TipoVariacao = 'Passivo'" & vbCrLf & _

        If tipo = "CONTABILIZACAO" OrElse tipo = "VCONTABILIZACAO" Then
            If tipo = "CONTABILIZACAO" Then
                sql &= "delete razao" & vbCrLf & _
                       "  from razao R" & vbCrLf & _
                       " inner Join #Pedidos P" & vbCrLf & _
                       "    on R.Empresa_id    = P.Empresa" & vbCrLf & _
                       "   and R.EndEmpresa_id = P.EndEmpresa" & vbCrLf & _
                       "   and R.Pedido        = P.Pedido" & vbCrLf & _
                       " where R.lote_id       = 6 " & vbCrLf & _
                       "   and R.movimento_id  >= " & UltimoDiaDoMesSql & vbCrLf

                sql &= "Insert into Razao(Empresa_Id, EndEmpresa_Id, Conta_Id, Cliente_Id, EndCliente_Id, Historico, Movimento_Id, Lote_Id, Sequencia_Id, UnidadeDeNegocio, Pedido, DataMoeda, DebitoOficial, CreditoOficial, DebitoMoeda, CreditoMoeda,  PrevistoRealizado, Processo, Situacao)" & vbCrLf & _
                       " (" & vbCrLf
            End If
            sql &= " select r.Empresa," & vbCrLf & _
                   "        r.EndEmpresa," & vbCrLf & _
                   "        case" & vbCrLf & _
                   "           when r.TipoPR = 'P' and VariacaoCorrigida > 0 and TL.TipoLancamento = 'C' then CxE.ContaVariacaoMonetariaAtiva" & vbCrLf & _
                   "           when r.TipoPR = 'P' and VariacaoCorrigida > 0 and TL.TipoLancamento = 'D' then CxE.ContaVariacaoMonetariaFornecedor" & vbCrLf & _
                   "           when r.TipoPR = 'P' and VariacaoCorrigida < 0 and TL.TipoLancamento = 'C' then CxE.ContaVariacaoMonetariaFornecedor" & vbCrLf & _
                   "           when r.TipoPR = 'P' and VariacaoCorrigida < 0 and TL.TipoLancamento = 'D' then CxE.ContaVariacaoMonetariaPassiva" & vbCrLf & _
                   "           when r.TipoPR = 'R' and VariacaoCorrigida > 0 and TL.TipoLancamento = 'D' then CxE.ContaVariacaoMonetariaCliente" & vbCrLf & _
                   "           when r.TipoPR = 'R' and VariacaoCorrigida > 0 and TL.TipoLancamento = 'C' then CxE.ContaVariacaoMonetariaAtiva" & vbCrLf & _
                   "           when r.TipoPR = 'R' and VariacaoCorrigida < 0 and TL.TipoLancamento = 'C' then CxE.ContaVariacaoMonetariaCliente" & vbCrLf & _
                   "           when r.TipoPR = 'R' and VariacaoCorrigida < 0 and TL.TipoLancamento = 'D' then CxE.ContaVariacaoMonetariaPassiva" & vbCrLf & _
                   "        end Conta," & vbCrLf & _
                   "        case" & vbCrLf & _
                   "           when r.TipoPR = 'P' and VariacaoCorrigida > 0 and TL.TipoLancamento = 'D' then P.Cliente" & vbCrLf & _
                   "           when r.TipoPR = 'P' and VariacaoCorrigida < 0 and TL.TipoLancamento = 'C' then P.Cliente" & vbCrLf & _
                   "           when r.TipoPR = 'R' and VariacaoCorrigida > 0 and TL.TipoLancamento = 'D' then P.Cliente" & vbCrLf & _
                   "           when r.TipoPR = 'R' and VariacaoCorrigida < 0 and TL.TipoLancamento = 'C' then P.Cliente" & vbCrLf & _
                   "           else ''" & vbCrLf & _
                   "        end Cliente," & vbCrLf & _
                   "        case" & vbCrLf & _
                   "           when r.TipoPR = 'P' and VariacaoCorrigida > 0 and TL.TipoLancamento = 'D' then P.EndCliente" & vbCrLf & _
                   "           when r.TipoPR = 'P' and VariacaoCorrigida < 0 and TL.TipoLancamento = 'C' then P.EndCliente" & vbCrLf & _
                   "           when r.TipoPR = 'R' and VariacaoCorrigida > 0 and TL.TipoLancamento = 'D' then P.EndCliente" & vbCrLf & _
                   "           when r.TipoPR = 'R' and VariacaoCorrigida < 0 and TL.TipoLancamento = 'C' then P.EndCliente" & vbCrLf & _
                   "           else 0" & vbCrLf & _
                   "        end EndCliente," & vbCrLf

            If tipo = "VCONTABILIZACAO" Then sql &= "        r.nome,"

            sql &= "        'VLR DE VARIAÇÃO CAMBIAL ' + case" & vbCrLf & _
                   "							           when r.TipoPR = 'P' and VariacaoCorrigida > 0 and TL.TipoLancamento = 'C' then 'ATIVA'" & vbCrLf & _
                   "							           when r.TipoPR = 'P' and VariacaoCorrigida > 0 and TL.TipoLancamento = 'D' then 'ATIVA'" & vbCrLf & _
                   "							           when r.TipoPR = 'P' and VariacaoCorrigida < 0 and TL.TipoLancamento = 'C' then 'PASSIVA'" & vbCrLf & _
                   "							           when r.TipoPR = 'P' and VariacaoCorrigida < 0 and TL.TipoLancamento = 'D' then 'PASSIVA'" & vbCrLf & _
                   "							           when r.TipoPR = 'R' and VariacaoCorrigida > 0 and TL.TipoLancamento = 'D' then 'ATIVA'" & vbCrLf & _
                   "							           when r.TipoPR = 'R' and VariacaoCorrigida > 0 and TL.TipoLancamento = 'C' then 'ATIVA'" & vbCrLf & _
                   "							           when r.TipoPR = 'R' and VariacaoCorrigida < 0 and TL.TipoLancamento = 'C' then 'PASSIVA'" & vbCrLf & _
                   "							           when r.TipoPR = 'R' and VariacaoCorrigida < 0 and TL.TipoLancamento = 'D' then 'PASSIVA'" & vbCrLf & _
                   "							         end + ' EM " & UltimoDiaDoMesbr & " PTAX ' + convert(nvarchar,@Cotacao) + ', PEDIDO '+ convert(nvarchar,P.Pedido_id) +' - ' + C.Nome as Historico," & vbCrLf & _
                   "        " & UltimoDiaDoMesSql & " as Movimento," & vbCrLf & _
                   "        6 as lote, " & vbCrLf & _
                   "        @Sequencia + ROW_NUMBER() OVER(ORDER BY P.Pedido_id) as sequencia," & vbCrLf & _
                   "        P.UnidadeDeNegocio," & vbCrLf & _
                   "        P.Pedido_id as Pedido," & vbCrLf & _
                   "        " & UltimoDiaDoMesSql & " as DataMoeda," & vbCrLf & _
                   "        case" & vbCrLf & _
                   "           when TL.TipoLancamento = 'D'" & vbCrLf & _
                   "             then abs(r.VariacaoCorrigida)" & vbCrLf & _
                   "             else 0" & vbCrLf & _
                   "        end DebitoOficial," & vbCrLf & _
                   "        case" & vbCrLf & _
                   "           when TL.TipoLancamento = 'C'" & vbCrLf & _
                   "             then abs(r.VariacaoCorrigida)" & vbCrLf & _
                   "             else 0" & vbCrLf & _
                   "        end CreditoOficial," & vbCrLf & _
                   "        case" & vbCrLf & _
                   "           when TL.TipoLancamento = 'D'" & vbCrLf & _
                   "             then round(abs(r.VariacaoCorrigida) / @Cotacao,2)" & vbCrLf & _
                   "             else 0" & vbCrLf & _
                   "        end DebitoMoeda," & vbCrLf & _
                   "        case" & vbCrLf & _
                   "           when TL.TipoLancamento = 'C'" & vbCrLf & _
                   "             then round(abs(r.VariacaoCorrigida) / @Cotacao,2)" & vbCrLf & _
                   "             else 0" & vbCrLf & _
                   "        end CreditoMoeda," & vbCrLf & _
                   "        'R' as PrevistoRealizado," & vbCrLf & _
                   "        'VARIACAO AUTOMATICA' as processo," & vbCrLf & _
                   "        1 as Situacao" & vbCrLf & _
                   "   from #razao r" & vbCrLf & _
                   "   Inner Join (Select 'C' as TipoLancamento" & vbCrLf & _
                   "                union" & vbCrLf & _
                   "               Select 'D'" & vbCrLf & _
                   "               ) TL" & vbCrLf & _
                   "      on 1 = 1" & vbCrLf & _
                   "  Inner Join Pedidos P" & vbCrLf & _
                   "     on P.Empresa_id    = r.Empresa" & vbCrLf & _
                   "    and P.EndEmpresa_id = r.EndEmpresa" & vbCrLf & _
                   "    and P.Pedido_id     = r.Pedido" & vbCrLf & _
                   "  Inner Join ClientesXEmpresas CxE" & vbCrLf & _
                   "     on P.Empresa_id    = CxE.Empresa_id" & vbCrLf & _
                   "    and P.EndEmpresa_id = CxE.EndEmpresa_Id" & vbCrLf & _
                   "  Inner Join Clientes C" & vbCrLf & _
                   "     on C.Cliente_id  = P.Cliente" & vbCrLf & _
                   "    and C.Endereco_id = P.EndCliente" & vbCrLf & _
                   "  Where VariacaoCorrigida <> 0"
            If tipo = "CONTABILIZACAO" Then
                sql &= ")"
            Else
                sql &= "  order by P.Pedido_id" & vbCrLf
            End If
        End If
        Return sql
    End Function


    Protected Sub ddlEmpresa_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            limparGrids()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub ddlSafra_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            limparGrids()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub


    Public Sub CarregarCotacao()
        Try
            txtIndice.Text = New [Lib].Negocio.Cotacao(ddlIndexador.SelectedValue, Funcoes.ValidaDataUtil("", 0, CDate("01-" & ddlMes.SelectedValue & "-" & ddlAno.SelectedValue).AddMonths(1))).Indice
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub ddlIndexador_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ddlIndexador.SelectedIndexChanged
        Try
            CarregarCotacao()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub chkConsolidarCliente_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            limparGrids()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub imgExtratoPedido2_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)
        Try
            Dim img As Image = CType(sender, Image)
            Dim row As GridViewRow = CType(img.NamingContainer, GridViewRow)
            Dim i As Integer = row.RowIndex

            Extrato.Emitir(Me.Page, FinanceiroNovo, gridPedidos.Rows(i).Cells(2).Text, gridPedidos.Rows(i).Cells(3).Text, "T", gridPedidos.Rows(i).Cells(8).Text)


        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub imgExtratoPedidoEnc_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)
        Try
            Dim img As Image = CType(sender, Image)
            Dim row As GridViewRow = CType(img.NamingContainer, GridViewRow)
            Dim i As Integer = row.RowIndex

            Extrato.Emitir(Me.Page, FinanceiroNovo, gridCorrecaoEncerrados.Rows(i).Cells(1).Text, gridCorrecaoEncerrados.Rows(i).Cells(2).Text, "T", gridCorrecaoEncerrados.Rows(i).Cells(3).Text)

        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Public Sub limparGrids()
        gridPedidos.DataSource = Nothing
        gridPedidos.DataBind()
        chkConsolidarCliente.Checked = False
        gridCorrecao.DataSource = Nothing
        gridContabilizacao.DataSource = Nothing
    End Sub

    Private Sub LiberaEmpresa()
        If Not UsuarioServidor.LiberaEmpresa Then
            ddlEmpresa.Enabled = False
        End If
    End Sub

    Protected Sub lnkConsultar_Click(sender As Object, e As EventArgs) Handles lnkConsultar.Click
        Try
            CarregarGridPedidos()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkRelatorio_Click(sender As Object, e As EventArgs) Handles lnkRelatorio.Click
        Try
            gridCorrecao.DataSource = Nothing
            gridCorrecao.DataBind()

            gridContabilizacao.DataSource = Nothing
            gridContabilizacao.DataBind()

            Dim pedidos As String = ""
            For Each row As GridViewRow In gridPedidos.Rows
                If CType(row.FindControl("chkPedido"), CheckBox).Checked = True Then
                    pedidos &= "or (P.Empresa_id    ='" & row.Cells(1).Text & "' AND P.EndEmpresa_id =" & row.Cells(2).Text & " And P.Pedido_id =" & row.Cells(7).Text & ")"
                End If
            Next

            If pedidos.Length = 0 Then
                MsgBox(Me.Page, "Selecione um ou mais pedidos para correção.")
                Exit Sub
            End If

            Dim ds As DataSet
            If pedidos.Length > 0 Then
                pedidos = pedidos.Substring(2, pedidos.Length - 2)

                ds = Banco.ConsultaDataSet(getSql(pedidos, ""), "tabela")
                gridCorrecao.DataSource = ds
                gridCorrecao.DataBind()
            End If
            TC02.ActiveTab = TPVisualizacao
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkLimpar_Click(sender As Object, e As EventArgs) Handles lnkLimpar.Click
        Try
            ddlSafra.SelectedIndex = 0
            ddlMes.SelectedValue = Now.AddMonths(-1).Month.ToString("00")
            ddlAno.SelectedValue = Now.AddMonths(-1).Year
            txtNomeCliente.Text = ""
            txtCodigoCliente.Value = ""
            CarregarCotacao()
            limparGrids()
            LiberaEmpresa()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    'VISUALIZAR CORREÇÃO

    Protected Sub lnkRelatorioVC_Click(sender As Object, e As EventArgs) Handles lnkRelatorioVC.Click
        Try
            gridContabilizacao.DataSource = Nothing
            gridContabilizacao.DataBind()

            Dim pedidos As String = ""
            For Each row As GridViewRow In gridCorrecao.Rows
                If CType(row.FindControl("chkPedidoCorrecao"), CheckBox).Checked = True Then
                    pedidos &= "or (P.Empresa_id    ='" & row.Cells(2).Text & "' AND P.EndEmpresa_id =" & row.Cells(3).Text & " And P.Pedido_id =" & row.Cells(8).Text & ")"
                End If
            Next

            If pedidos.Length = 0 Then
                MsgBox(Me.Page, "Selecione um ou mais pedidos.")
                Exit Sub
            End If

            Dim ds As DataSet

            pedidos = pedidos.Substring(2, pedidos.Length - 2)

            ds = Banco.ConsultaDataSet(getSql(pedidos, "VCONTABILIZACAO"), "tabela")
            gridContabilizacao.DataSource = ds
            gridContabilizacao.DataBind()
            TC02.ActiveTab = TPContabilizacao
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkNovo_Click(sender As Object, e As EventArgs) Handles lnkNovo.Click
        Try
            Dim pedidos As String = ""
            For Each row As GridViewRow In gridCorrecao.Rows
                If CType(row.FindControl("chkPedidoCorrecao"), CheckBox).Checked = True Then
                    pedidos &= "or (P.Empresa_id    ='" & row.Cells(2).Text & "' AND P.EndEmpresa_id =" & row.Cells(3).Text & " And P.Pedido_id =" & row.Cells(8).Text & ")"
                End If
            Next

            If pedidos.Length = 0 Then
                MsgBox(Me.Page, "Selecione um ou mais pedidos")
                Exit Sub
            End If

            Dim sqls As New ArrayList
            pedidos = pedidos.Substring(2, pedidos.Length - 2)
            'pega os dados para a emissao do relatorio caso ele seja contabilizado com exito
            Dim ds = Banco.ConsultaDataSet(getSql(pedidos, ""), "CorrecaoMonetaria")
            sqls.Add(getSql(pedidos, "CONTABILIZACAO"))

            If Banco.GravaBanco(sqls) Then
                MsgBox(Me.Page, "Correção efetuada com Sucesso.", eTitulo.Sucess)
                Funcoes.BindReport(Me.Page, ds, "Cr_CorrecaoMonetaria", eExportType.PDF)

                Dim param As New Dictionary(Of String, eTipoCampo)
                param.Add("VariacaoCambial", eTipoCampo.ValorComTotalizador)
                param.Add("VariacaoPassivaCorrecao", eTipoCampo.ValorComTotalizador)
                param.Add("VariacaoAtivaCorrecao", eTipoCampo.ValorComTotalizador)
                param.Add("VariacaoCorrigida", eTipoCampo.ValorComTotalizador)
                param.Add("FinanceiroBaixadoEmDolar", eTipoCampo.ValorComTotalizador)
                param.Add("FinanceiroBaixadoEmReais", eTipoCampo.ValorComTotalizador)
                param.Add("ValorDolarPedido", eTipoCampo.ValorComTotalizador)
                param.Add("ValorEmDolarPedidoAtualizadoEmReais", eTipoCampo.ValorComTotalizador)
                param.Add("ReaisNotas", eTipoCampo.ValorComTotalizador)

                Funcoes.BindExcelOffice(Me.Page, ds, "Correção Monetária", param)
            Else
                MsgBox(Me.Page, "Erro durante o processo de correção.", eTitulo.Sucess)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "CotacoesDeMoedas")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub

    Protected Sub lnkConsultaEncerrados_Click(sender As Object, e As EventArgs) Handles lnkConsultaEncerrados.Click
        Dim sql As String = getsqlEncerrado("", "VISUALIZACAO")
        Dim ds As DataSet
        Dim banco As New AcessaBanco
        ds = banco.ConsultaDataSet(sql, "ENCERRADOS")

        gridCorrecaoEncerrados.DataSource = ds.Tables(0)
        gridCorrecaoEncerrados.DataBind()
    End Sub

    Protected Sub lnkVisContEnc_Click(sender As Object, e As EventArgs) Handles lnkVisContEnc.Click
        Try
            gridContEncerrados.DataSource = Nothing
            gridContEncerrados.DataBind()

            Dim pedidos As String = ""
            For Each row As GridViewRow In gridCorrecaoEncerrados.Rows
                If CType(row.FindControl("chkPedidoCorrecao"), CheckBox).Checked = True Then
                    pedidos &= "or (P.Empresa_id    ='" & row.Cells(1).Text & "' AND P.EndEmpresa_id =" & row.Cells(2).Text & " And P.Pedido_id =" & row.Cells(3).Text & ")"
                End If
            Next

            If pedidos.Length = 0 Then
                MsgBox(Me.Page, "Selecione um ou mais pedidos.")
                Exit Sub
            End If

            Dim ds As DataSet

            pedidos = pedidos.Substring(2, pedidos.Length - 2)

            ds = Banco.ConsultaDataSet(getsqlEncerrado(pedidos, "VISUALIZACAOCONTABILIZACAO"), "tabela")
            gridContEncerrados.DataSource = ds
            gridContEncerrados.DataBind()
            TC03.ActiveTab = TPContabilizacaoEncerrados
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkContEnc_Click(sender As Object, e As EventArgs) Handles lnkContEnc.Click
        Try
            gridContEncerrados.DataSource = Nothing
            gridContEncerrados.DataBind()

            Dim pedidos As String = ""
            For Each row As GridViewRow In gridCorrecaoEncerrados.Rows
                If CType(row.FindControl("chkPedidoCorrecao"), CheckBox).Checked = True Then
                    pedidos &= "or (P.Empresa_id    ='" & row.Cells(1).Text & "' AND P.EndEmpresa_id =" & row.Cells(2).Text & " And P.Pedido_id =" & row.Cells(3).Text & ")"
                End If
            Next

            If pedidos.Length = 0 Then
                MsgBox(Me.Page, "Selecione um ou mais pedidos.")
                Exit Sub
            End If

            pedidos = pedidos.Substring(2, pedidos.Length - 2)

            Dim ds As DataSet = Banco.ConsultaDataSet(getsqlEncerrado(pedidos, "VISUALIZACAOCONTABILIZACAO"), "Conta")

            Dim sqls As New ArrayList
            sqls.Add(getsqlEncerrado(pedidos, "CONTABILIZACAO"))

            If Banco.GravaBanco(sqls) Then
                MsgBox(Me.Page, "Correção efetuada com Sucesso.", eTitulo.Sucess)
                'Funcoes.BindReport(Me.Page, ds, "Cr_CorrecaoMonetaria", eExportType.PDF)

                'Dim param As New Dictionary(Of String, eTipoCampo)
                'param.Add("VariacaoCambial", eTipoCampo.ValorComTotalizador)
                'param.Add("VariacaoPassivaCorrecao", eTipoCampo.ValorComTotalizador)
                'param.Add("VariacaoAtivaCorrecao", eTipoCampo.ValorComTotalizador)
                'param.Add("VariacaoCorrigida", eTipoCampo.ValorComTotalizador)
                'param.Add("FinanceiroBaixadoEmDolar", eTipoCampo.ValorComTotalizador)
                'param.Add("FinanceiroBaixadoEmReais", eTipoCampo.ValorComTotalizador)
                'param.Add("ValorDolarPedido", eTipoCampo.ValorComTotalizador)
                'param.Add("ValorEmDolarPedidoAtualizadoEmReais", eTipoCampo.ValorComTotalizador)
                'param.Add("ReaisNotas", eTipoCampo.ValorComTotalizador)

                'Funcoes.BindExcelOffice(Me.Page, ds, "Correção Monetária", param)
            Else
                MsgBox(Me.Page, "Erro durante o processo de correção.", eTitulo.Sucess)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub
End Class