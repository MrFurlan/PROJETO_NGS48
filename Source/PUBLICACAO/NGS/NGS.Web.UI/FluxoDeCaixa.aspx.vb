Imports System.Data
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class FluxoDeCaixa
    Inherits BasePage

#Region "Events"

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.setMenu(eModulo.Financeiro)
            If Not IsPostBack And IsConnect Then
                If Funcoes.VerificaPermissao("FluxoDeCaixa", "ACESSAR") Then
                    CarregarUnidade()
                    Limpar()
                Else
                    MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/financeiro.aspx")
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub ddlUnidadeDeNegocio_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddlUnidadeDeNegocio.SelectedIndexChanged
        Try
            CarregarEmpresas()
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
            Dim ds As DataSet = getDataSet()
            Dim param As New Dictionary(Of String, Object)
            param.Add("DataAbertura", txtData1.Text)
            param.Add("Periodo", "Período: " & txtData1.Text & " á " & txtData2.Text)
            param.Add("PTax", Funcoes.PegarValorConversao(3, CDate(txtData1.Text)).ToString("N8"))
            Funcoes.BindReport(Me.Page, ds, "Cr_FluxoDeCaixa", eExportType.PDF, param)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

#End Region

#Region "Methods"

    Private Sub Limpar()
        Funcoes.VerificaUnidadeDeNegocio(ddlUnidadeDeNegocio, ddlEmpresa)
        txtData1.Text = DateTime.Now.ToString("dd/MM/yyyy")
        txtData2.Text = DateTime.Now.ToString("dd/MM/yyyy")

        LiberaEmpresa()
    End Sub

    Private Sub LiberaEmpresa()
        If Not UsuarioServidor.LiberaEmpresa Then
            ddlUnidadeDeNegocio.Enabled = False
            ddlEmpresa.Enabled = False
        End If
    End Sub

    Private Sub CarregarUnidade()
        ddl.Carregar(ddlUnidadeDeNegocio, CarregarDDL.Tabela.UnidadeDeNegocio, "", True)
    End Sub

    Private Sub CarregarEmpresas()
        ddl.Carregar(ddlEmpresa, CarregarDDL.Tabela.Empresas, ddlUnidadeDeNegocio.SelectedValue, True)
    End Sub

    Private Function getDataSet() As DataSet
        Dim sql As String
        sql = "Declare @data as datetime;" & vbCrLf & _
              "   set @data = '" & CDate(txtData1.Text).ToString("yyyy-MM-dd") & "';" & vbCrLf & _
              "" & vbCrLf & _
              "Select r.Empresa_Id," & vbCrLf & _
              "       r.EndEmpresa_Id as EndEmpresa," & vbCrLf & _
              "       bc.Banco_Id," & vbCrLf & _
              "       bc.Agencia_Id," & vbCrLf & _
              "       replace(bc.Conta_Id, '.','') as Conta_Id," & vbCrLf & _
              "       Sum(r.DebitoOficial - r.CreditoOficial) as SaldoAplicFinanc" & vbCrLf & _
              "  into #ContaAplic" & vbCrLf & _
              "  FROM Razao r" & vbCrLf & _
              " Inner Join PlanoDeContas pc" & vbCrLf & _
              "    ON pc.Conta_Id      = r.Conta_Id" & vbCrLf & _
              "  Inner Join BancosXContas bc" & vbCrLf & _
              "   	ON bc.ContaContabil = pc.Conta_Id" & vbCrLf & _
              "    AND bc.FluxoDeCaixa = 1" & vbCrLf & _
              "    AND bc.Ativo = 1" & vbCrLf & _
              "  WHERE r.Empresa_Id    ='" & ddlEmpresa.SelectedValue.Split("-")(0) & "'" & vbCrLf & _
              "    And r.EndEmpresa_Id = " & ddlEmpresa.SelectedValue.Split("-")(1) & vbCrLf & _
              "    And r.Movimento_Id  < @data" & vbCrLf & _
              "    And bc.TipoConta    = 'A'" & vbCrLf & _
              "  Group by r.Empresa_Id, r.EndEmpresa_Id, bc.Banco_Id,  bc.Agencia_Id, replace(bc.Conta_Id, '.','');" & vbCrLf & _
              "" & vbCrLf & _
              " SELECT r.Empresa_id as Empresa, r.EndEmpresa_Id as EndEmpresa," & vbCrLf & _
              "        bc.Banco_Id, bc.Agencia_Id, bc.Conta_Id," & vbCrLf & _
              "        Emp.Cliente_Id + ' - ' + Emp.Nome + ' - ' + Emp.Cidade + ' / ' + Emp.Estado as DescricaoEmpresa," & vbCrLf & _
              "   	   r.Conta_Id as Conta, pc.Titulo, isnull(bc.LimiteBancario, 0) as LimiteBancario,   0.00 as SaldoAplicFinanc," & vbCrLf & _
              "   	   Case" & vbCrLf & _
              "          when r.Movimento_Id < @data" & vbCrLf & _
              "   		   Then COALESCE(SUM(r.DebitoOficial - r.CreditoOficial), 0)" & vbCrLf & _
              "   		   else 0" & vbCrLf & _
              "   	   end as SaldoExtratoContaAnterior," & vbCrLf & _
              "   	   Case" & vbCrLf & _
              "          when r.Movimento_Id = @data" & vbCrLf & _
              "   		   Then COALESCE(SUM(r.DebitoOficial), 0)" & vbCrLf & _
              "   		   else 0" & vbCrLf & _
              "   	   end as DebitoExtratoContaDoDia," & vbCrLf & _
              "   	   Case" & vbCrLf & _
              "          when r.Movimento_Id = @data" & vbCrLf & _
              "   		   Then COALESCE(SUM(r.CreditoOficial), 0)" & vbCrLf & _
              "   		   else 0" & vbCrLf & _
              "   	   end as CreditoExtratoContaDoDia," & vbCrLf & _
              "   	   Case" & vbCrLf & _
              "          When r.Movimento_Id < @data and (r.Conciliacao <> 'B' or r.Conciliacao is null)" & vbCrLf & _
              "   	       Then COALESCE(SUM(r.DebitoOficial - r.CreditoOficial), 0)" & vbCrLf & _
              "   		   else 0" & vbCrLf & _
              "   	   End as SaldoConciliacaoDeContasAnterior," & vbCrLf & _
              "   	   COALESCE(SUM(r.DebitoOficial - r.CreditoOficial), 0) AS SaldoExtratoConta" & vbCrLf & _
              "   into #Banco" & vbCrLf & _
              "   FROM Razao r" & vbCrLf & _
              "  Inner Join Clientes Emp" & vbCrLf & _
              "     On Emp.Cliente_Id  = r.Empresa_Id" & vbCrLf & _
              "    And Emp.Endereco_Id = r.EndEmpresa_Id" & vbCrLf & _
              "  Inner Join PlanoDeContas pc" & vbCrLf & _
              "     ON pc.Conta_Id      = r.Conta_Id" & vbCrLf & _
              "  Inner Join BancosXContas bc" & vbCrLf & _
              "     ON bc.ContaContabil = pc.Conta_Id" & vbCrLf & _
              "    AND bc.FluxoDeCaixa  = 1" & vbCrLf & _
              "    AND bc.Ativo  = 1" & vbCrLf & _
              "  WHERE r.Empresa_Id = '" & ddlEmpresa.SelectedValue.Split("-")(0) & "'" & vbCrLf & _
              "    And r.EndEmpresa_Id = " & ddlEmpresa.SelectedValue.Split("-")(1) & vbCrLf & _
              "    And r.Movimento_Id <= @data" & vbCrLf & _
              "    And bc.TipoConta <> 'A'" & vbCrLf & _
              "  group by r.Empresa_id, r.EndEmpresa_Id, Emp.Cliente_Id + ' - ' + Emp.Nome + ' - ' + Emp.Cidade + ' / ' + Emp.Estado," & vbCrLf & _
              "           bc.Banco_Id, bc.Agencia_Id, bc.Conta_Id, r.Conta_Id, bc.TipoConta, pc.Titulo, r.Movimento_Id, r.Conciliacao, isnull(bc.LimiteBancario, 0);" & vbCrLf & _
              "" & vbCrLf & _
              "Select Empresa, EndEmpresa, DescricaoEmpresa," & vbCrLf & _
              "       Banco_Id, Agencia_Id, replace(Conta_Id,'.','') as Conta_Id, Conta, Titulo, LimiteBancario," & vbCrLf & _
              "   	  SUM(SaldoExtratoContaAnterior) - SUM(SaldoConciliacaoDeContasAnterior) as SaldoExtratoContaAnterior," & vbCrLf & _
              "       Sum(SaldoAplicFinanc) as SaldoAplicFinanc," & vbCrLf & _
              "   	  SUM(DebitoExtratoContaDoDia) as DebitoExtratoContaDoDia," & vbCrLf & _
              "   	  SUM(CreditoExtratoContaDoDia) as CreditoExtratoContaDoDia," & vbCrLf & _
              "   	  SUM(SaldoConciliacaoDeContasAnterior) as SaldoConciliacaoDeContasAnterior," & vbCrLf & _
              "   	  SUM(SaldoExtratoConta) as SaldoExtratoConta" & vbCrLf & _
              "  Into #Razao" & vbCrLf & _
              "  from #Banco" & vbCrLf & _
              " Group by Empresa, EndEmpresa, DescricaoEmpresa, Banco_Id, Agencia_Id, Conta_Id, Conta, Titulo, LimiteBancario;" & vbCrLf & _
              "" & vbCrLf & _
              "Update #Razao  set SaldoAplicFinanc = c.SaldoAplicFinanc" & vbCrLf & _
              "  From #ContaAplic c" & vbCrLf & _
              " Inner Join #Razao r" & vbCrLf & _
              "	   on r.Empresa    = c.Empresa_Id" & vbCrLf & _
              "   and r.EndEmpresa = c.EndEmpresa" & vbCrLf & _
              "	  and r.Banco_Id   = c.Banco_Id" & vbCrLf & _
              "	  and r.Agencia_Id = c.Agencia_Id" & vbCrLf & _
              "	  and r.Conta_Id   = c.Conta_Id;" & vbCrLf & _
              "" & vbCrLf & _
              "Select Empresa, EndEmpresa, DescricaoEmpresa," & vbCrLf & _
              "       Banco_Id, Agencia_Id, Conta_Id, Conta, Titulo," & vbCrLf & _
              "       LimiteBancario," & vbCrLf & _
              "   	  SaldoExtratoContaAnterior," & vbCrLf & _
              "       SaldoAplicFinanc," & vbCrLf & _
              "       DebitoExtratoContaDoDia," & vbCrLf & _
              "       CreditoExtratoContaDoDia," & vbCrLf & _
              "   	  SaldoConciliacaoDeContasAnterior," & vbCrLf & _
              "       SaldoExtratoConta  + SaldoAplicFinanc as SaldoExtratoConta" & vbCrLf & _
              "  from #Razao;" & vbCrLf & _
              "" & vbCrLf


        '*************** Criação da tabela temporária para a view VW_Titulovirtual por questões de performance
        sql &= "SELECT * " & vbCrLf & _
               "  INTO #TituloVirtual" & vbCrLf & _
               "  FROM VW_TituloVirtual" & vbCrLf & _
               "" & vbCrLf


        '****************************************************************************************************************
        '**************************************** CONTAS A RECEBER ******************************************************
        '****************************************************************************************************************

        sql &= "Select Titulo," & vbCrLf & _
               "       Prorrogacao," & vbCrLf & _
               "	   ValorLiquido," & vbCrLf & _
               "	   MoedaValorLiquido" & vbCrLf & _
               "  into #Titulos" & vbCrLf & _
               "  from (SELECT 'R' as Titulo," & vbCrLf & _
               "           	  Titulos.Prorrogacao," & vbCrLf & _
               "               Case" & vbCrLf & _
               "                  when Titulos.Provisao = 1" & vbCrLf & _
               "           		   then Titulos.ValorLiquido" & vbCrLf & _
               "                    else Case" & vbCrLf & _
               "                           when Titulos.Moeda <> 1" & vbCrLf & _
               "                             then convert(numeric(18,2), Titulos.MoedaValorDoDocumento * Cotacoes.indice)" & vbCrLf & _
               "                             else Titulos.ValorDoDocumento" & vbCrLf & _
               "                         End" & vbCrLf & _
               "               end ValorLiquido," & vbCrLf & _
               "               0.0 as MoedaValorLiquido" & vbCrLf & _
               "               --case" & vbCrLf & _
               "               --  when Titulos.Provisao = 1" & vbCrLf & _
               "           	  --    then Titulos.MoedaValorLiquido" & vbCrLf & _
               "               --    else Case" & vbCrLf & _
               "               --           when Titulos.Moeda = 1" & vbCrLf & _
               "               --             then convert(numeric(18,2),Titulos.ValorDoDocumento / Cotacoes.indice)" & vbCrLf & _
               "               --              else Titulos.MoedaValorDoDocumento" & vbCrLf & _
               "               --         End" & vbCrLf & _
               "               --end MoedaValorLiquido" & vbCrLf & _
               "          FROM ContasAReceber AS Titulos" & vbCrLf & _
               "         LEFT Join cotacoes" & vbCrLf & _
               "            on Cotacoes.Data_id      = Titulos.Prorrogacao" & vbCrLf & _
               "           and Cotacoes.Indexador_Id = Titulos.Indexador" & vbCrLf & _
               "          Left Join Pedidos P" & vbCrLf & _
               "            on Titulos.Empresa    = P.Empresa_Id" & vbCrLf & _
               "           and Titulos.EndEmpresa = P.EndEmpresa_Id" & vbCrLf & _
               "           and Titulos.Pedido     = P.Pedido_Id" & vbCrLf & _
               "         Where Titulos.Situacao in (1,101,102)" & vbCrLf & _
               "           and Titulos.Grupado <> 'M'" & vbCrLf & _
               "           and Titulos.ContaContabilCliente NOT LIKE '1010101%'" & vbCrLf & _
               "           and Titulos.ContaContabilCliente NOT LIKE '1010102%'" & vbCrLf

        If Not chkEmprestimo.Checked Then sql &= "  and isnull(P.finalidade,0) not in (4)"

        If Left(Session("ssEmpresa"), 8) = "04854422" Then
            sql &= "           and ((Titulos.Provisao = 2) or (Titulos.Provisao = 3 and isnull(Titulos.PedidoFixacao,0) > 0 and isnull(Titulos.Pedido,0) <> isnull(Titulos.PedidoFixacao,0)))" & vbCrLf
        Else
            sql &= "           and (Titulos.Provisao = 2 or Titulos.Provisao = 3)   " & vbCrLf
        End If

        sql &= "           and Titulos.Prorrogacao BETWEEN '" & CDate(txtData1.Text).AddDays(1).ToString("yyyy/MM/dd") & "' and '" & CDate(txtData2.Text).ToString("yyyy/MM/dd") & "'" & vbCrLf & _
               "         Union All" & vbCrLf & _
               "		Select 'R',Vencimento, Parcelaoficial, ParcelaMoeda" & vbCrLf & _
               "		  from #TituloVirtual" & vbCrLf & _
               "		 Where CR = 'R'" & vbCrLf & _
               "		   and Vencimento BETWEEN '" & CDate(txtData1.Text).AddDays(1).ToString("yyyy/MM/dd") & "' and '" & CDate(txtData2.Text).ToString("yyyy/MM/dd") & "'" & vbCrLf & _
               "	  ) sb;" & vbCrLf

        '****************************************************************************************************************
        '****************************************  CONTAS A PAGAR  ******************************************************
        '****************************************************************************************************************

        sql &= "Insert Into #Titulos" & vbCrLf & _
               "SELECT 'P' as Titulo," & vbCrLf & _
               "       Titulos.Prorrogacao," & vbCrLf & _
               "       case" & vbCrLf & _
               "         when Titulos.Provisao = 1" & vbCrLf & _
               "		   then Titulos.ValorLiquido" & vbCrLf & _
               "           else Case" & vbCrLf & _
               "                  When Titulos.Moeda <> 1" & vbCrLf & _
               "                    Then convert(numeric(18,2), Titulos.MoedaValorDoDocumento * Cotacoes.indice)" & vbCrLf & _
               "                    Else Titulos.ValorDoDocumento" & vbCrLf & _
               "                End" & vbCrLf & _
               "       end ValorLiquido," & vbCrLf & _
               "       0.0 as MoedaValorLiquido" & vbCrLf & _
               "       --case" & vbCrLf & _
               "       --   when Titulos.Provisao = 1" & vbCrLf & _
               "	   --     then Titulos.MoedaValorLiquido" & vbCrLf & _
               "       --     else Case" & vbCrLf & _
               "       --            when Titulos.Moeda = 1" & vbCrLf & _
               "       --              then convert(numeric(18,2),Titulos.ValorDoDocumento / Cotacoes.indice)" & vbCrLf & _
               "       --              else Titulos.MoedaValorDoDocumento" & vbCrLf & _
               "       --          End" & vbCrLf & _
               "       --end MoedaValorLiquido" & vbCrLf & _
               "  FROM ContasAPagar AS Titulos" & vbCrLf & _
               " LEFT Join cotacoes" & vbCrLf & _
               "    on Cotacoes.Data_id      = Titulos.Prorrogacao" & vbCrLf & _
               "   and Cotacoes.Indexador_Id = Titulos.Indexador" & vbCrLf & _
               " INNER JOIN Clientes AS Empresas" & vbCrLf & _
               "    ON Titulos.Empresa    = Empresas.Cliente_Id" & vbCrLf & _
               "   AND Titulos.EndEmpresa = Empresas.Endereco_Id" & vbCrLf & _
               " INNER JOIN Clientes" & vbCrLf & _
               "    ON Titulos.Cliente    = Clientes.Cliente_Id" & vbCrLf & _
               "   AND Titulos.EndCliente = Clientes.Endereco_Id" & vbCrLf & _
               "  LEFT OUTER JOIN ComprasXProdutos AS Carteiras" & vbCrLf & _
               "    ON Titulos.Carteira = Carteiras.Produto_Id" & vbCrLf & _
               "  LEFT OUTER JOIN Carteira" & vbCrLf & _
               "    ON Titulos.CarteiraDoTitulo = Carteira.Carteira_Id " & vbCrLf & _
               "  Left Join Pedidos P" & vbCrLf & _
               "    on Titulos.Empresa    = P.Empresa_Id" & vbCrLf & _
               "   and Titulos.EndEmpresa = P.EndEmpresa_Id" & vbCrLf & _
               "   and Titulos.Pedido     = P.Pedido_Id" & vbCrLf & _
               " Where Titulos.Situacao in(1,101,102)" & vbCrLf & _
               "   AND Titulos.Grupado <> 'M'" & vbCrLf & _
               "   AND Titulos.ContaContabilCliente NOT LIKE '1010101%'" & vbCrLf & _
               "   AND Titulos.ContaContabilCliente NOT LIKE '1010102%'" & vbCrLf

        If Not chkEmprestimo.Checked Then sql &= "  and isnull(P.finalidade,0) not in (4)"

        If Left(Session("ssEmpresa"), 8) = "04854422" Then
            sql &= "   and ((Titulos.Provisao = 2) or (Titulos.Provisao = 3 and isnull(Titulos.PedidoFixacao,0) > 0 and isnull(Titulos.Pedido,0) <> isnull(Titulos.PedidoFixacao,0)))   " & vbCrLf
        Else
            sql &= "   and (Titulos.Provisao = 2 or Titulos.Provisao = 3)   " & vbCrLf
        End If
        sql &= "    AND Titulos.Prorrogacao BETWEEN '" & CDate(txtData1.Text).AddDays(1).ToString("yyyy/MM/dd") & "' and '" & CDate(txtData2.Text).ToString("yyyy/MM/dd") & "'" & vbCrLf & _
               "  Union All" & vbCrLf & _
               " Select 'P',Vencimento, Parcelaoficial, ParcelaMoeda" & vbCrLf & _
               "   from #TituloVirtual" & vbCrLf & _
               "  Where CR = 'P'" & vbCrLf & _
               "	and Vencimento BETWEEN '" & CDate(txtData1.Text).AddDays(1).ToString("yyyy/MM/dd") & "' and '" & CDate(txtData2.Text).ToString("yyyy/MM/dd") & "';" & vbCrLf & _
               "" & vbCrLf

        '****************************************************************************************************************
        '****************************************************************************************************************
        '****************************************************************************************************************

        sql &= "Select Prorrogacao," & vbCrLf & _
               "   	   Coalesce(SUM(Case When Titulo = 'R'" & vbCrLf & _
               "   			Then ValorLiquido" & vbCrLf & _
               "   			else 0" & vbCrLf & _
               "   		End),0) EntradaValorLiquido," & vbCrLf & _
               "   		Coalesce(SUM(Case When Titulo = 'R'" & vbCrLf & _
               "   			Then MoedaValorLiquido" & vbCrLf & _
               "   			else 0" & vbCrLf & _
               "   		End),0) EntradaValorMoeda," & vbCrLf & _
               "   		Coalesce(SUM(Case When Titulo = 'P'" & vbCrLf & _
               "   			Then ValorLiquido" & vbCrLf & _
               "   			else 0" & vbCrLf & _
               "   		End),0) SaidaValorLiquido," & vbCrLf & _
               "   		Coalesce(SUM(Case When Titulo = 'P'" & vbCrLf & _
               "   			Then MoedaValorLiquido" & vbCrLf & _
               "   			else 0" & vbCrLf & _
               "   		End ),0)SaidaValorMoeda," & vbCrLf & _
               "   	    SUM(ValorLiquido) as ValorLiquido, SUM(MoedaValorLiquido) as MoedaValorLiquido" & vbCrLf & _
               "   Into #t" & vbCrLf & _
               "   from #Titulos" & vbCrLf & _
               "  group by Titulo, Prorrogacao" & vbCrLf & _
               "  Order By Prorrogacao;" & vbCrLf & _
               "" & vbCrLf

        sql &= "Declare @SaldoAnt as decimal(18, 2);" & vbCrLf & _
               "Select @SaldoAnt =  (SUM(SaldoExtratoConta) + SUM(SaldoAplicFinanc))" & vbCrLf & _
               "  From #Razao;" & vbCrLf & _
               "Insert Into #t (Prorrogacao, EntradaValorLiquido, EntradaValorMoeda, SaidaValorLiquido, SaidaValorMoeda) values ('" & CDate(txtData1.Text).ToString("yyyy/MM/dd") & "', 0.00, 0.00, 0.00, 0.00);" & vbCrLf & _
               "Select Prorrogacao," & vbCrLf & _
               "       SUM(EntradaValorLiquido) as EntradaValorLiquido," & vbCrLf & _
               "       SUM(EntradaValorMoeda)   as EntradaValorMoeda," & vbCrLf & _
               "   	   SUM(SaidaValorLiquido)   as SaidaValorLiquido," & vbCrLf & _
               "       SUM(SaidaValorMoeda)     as SaidaValorMoeda," & vbCrLf & _
               "   	   isnull(( Select @SaldoAnt + (SUM(EntradaValorLiquido - SaidaValorLiquido))" & vbCrLf & _
               "                  From #t" & vbCrLf & _
               "                 where Prorrogacao <= t.Prorrogacao), @SaldoAnt) as Saldo" & vbCrLf & _
               "  from #t as t" & vbCrLf & _
               " group by Prorrogacao" & vbCrLf & _
               " order by Prorrogacao;" & vbCrLf & _
               "Drop Table #Banco;" & vbCrLf & _
               "Drop Table #Titulos;" & vbCrLf & _
               "Drop Table #t;" & vbCrLf & _
               "Drop table #Razao;" & vbCrLf

        Dim ds As DataSet = Banco.ConsultaDataSet(sql, "ResumoDeBancos")

        If Deserializar() IsNot Nothing AndAlso ds.Tables IsNot Nothing AndAlso ds.Tables.Count > 1 Then
            ds.Tables(1).TableName = "FluxosDeCaixaDetalhes"
        End If
        Return ds
    End Function

#End Region

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "FluxoDeCaixa")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub
End Class