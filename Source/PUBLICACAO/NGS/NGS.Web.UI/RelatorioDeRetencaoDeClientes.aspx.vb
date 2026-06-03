Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis
Imports System.Data
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared

Public Class RelatorioDeRetencaoDeClientes
    Inherits BasePage

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Revenda)
        If Not IsPostBack And IsConnect Then
            If Funcoes.VerificaPermissao("RelatorioDeRetencaoDeClientes", "ACESSAR") Then
                ddl.Carregar(ddlMarcaProd, CarregarDDL.Tabela.Marca, "")
                ddl.Carregar(ddlEmpresa, CarregarDDL.Tabela.Empresas, "")
                ddl.Carregar(ddlCultura, CarregarDDL.Tabela.Cultura, "")
                ddl.Carregar(ddlSafra01, CarregarDDL.Tabela.Safra, "")
                ddl.Carregar(ddlSafra02, CarregarDDL.Tabela.Safra, "")
                ddl.Carregar(ddlSafra03, CarregarDDL.Tabela.Safra, "")
            Else
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página", "Revenda.aspx")
                Exit Sub
            End If
        End If
    End Sub

    Protected Sub ddlMarcaProd_SelectedIndexChanged(ByVal sender As Object, ByVal e As EventArgs) Handles ddlMarcaProd.SelectedIndexChanged
        Try
            If ddlMarcaProd.SelectedIndex = 0 Then
                ucSelecaoProduto.WhereProduto = ""
            Else
                ucSelecaoProduto.WhereProduto = "marca = " & ddlMarcaProd.SelectedValue
            End If
            ucSelecaoProduto.CarregarNivel(1)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Public Function ValidaRelatorio() As Boolean
        If ddlEmpresa.SelectedIndex = 0 Then
            MsgBox(Me.Page, "Selecione a Empresa")
            Return False
        End If

        If ddlMarcaProd.SelectedIndex = 0 And Not ucSelecaoProduto.TemSelecionado Then
            MsgBox(Me.Page, "Selecione os produtos no grid ou a Marca.")
            Return False
        End If

        If ddlSafra01.SelectedIndex = 0 Then
            MsgBox(Me.Page, "Selecione a Safra 01")
            Return False
        End If

        If ddlSafra02.SelectedIndex = 0 Then
            MsgBox(Me.Page, "Selecione a Safra 02")
            Return False
        End If

        If ddlSafra03.SelectedIndex = 0 Then
            MsgBox(Me.Page, "Selecione a Safra 03")
            Return False
        End If

        Return True
    End Function

    Protected Sub lnkRelatorio_Click(sender As Object, e As EventArgs) Handles lnkRelatorio.Click
        Dim rpt As New ReportDocument()

        Try
            If Funcoes.VerificaPermissao("RelatorioDeRetencaoDeClientes", "RELATORIO") Then
                If Not ValidaRelatorio() Then Exit Sub

                Dim Parametros As String = ""
                Parametros = "Parametros do Relatorio " & vbCrLf & _
                             "Empresa: " & ddlEmpresa.SelectedItem.Text & vbCrLf & _
                             IIf(chkConsolidarEmpresa.Checked, "Empresa Consolidada" & vbCrLf, "") & _
                             "Safra 01: " & ddlSafra01.SelectedItem.Text & vbCrLf & _
                             "Safra 02: " & ddlSafra02.SelectedItem.Text & vbCrLf & _
                             "Safra 03: " & ddlSafra03.SelectedItem.Text & vbCrLf


                Dim sql As String = ""

                sql = " Declare" & vbCrLf & _
                      " @Safra01 nvarchar(50)," & vbCrLf & _
                      " @Safra02 nvarchar(50)," & vbCrLf & _
                      " @Safra03 nvarchar(50)" & vbCrLf & _
                      " Set @Safra01 = '" & ddlSafra01.SelectedValue & "'" & vbCrLf & _
                      " Set @Safra02 = '" & ddlSafra02.SelectedValue & "'" & vbCrLf & _
                      " Set @Safra03 = '" & ddlSafra03.SelectedValue & "'" & vbCrLf

                sql &= "Select Safra_Id" & vbCrLf & _
                       "  into #Safras" & vbCrLf & _
                       "  from Safras" & vbCrLf & _
                       " where safra_id in (@Safra01, @Safra02, @Safra03);" & vbCrLf

                sql &= "Select ISNULL(sbCultura.cliente_id, sbNotas.cliente_id) as cliente_id," & vbCrLf & _
                       "       ISNULL(sbCultura.Safra_Id, sbNotas.Safra_Id) as Safra," & vbCrLf & _
                       "       ISNULL(sbCultura.AreaPlantada, 0) as AreaPlantada," & vbCrLf & _
                       "       ISNULL(sbNotas.QtdeEmb,0) as QtdeEmb" & vbCrLf & _
                       " Into #Temp" & vbCrLf & _
                       " from" & vbCrLf & _
                       "	(" & vbCrLf & _
                       "     Select case" & vbCrLf & _
                       "			   when len(cxs.Cliente_Id) = 14" & vbCrLf & _
                       "				 then left(cxs.Cliente_Id,8)" & vbCrLf & _
                       "				 else cxs.Cliente_Id" & vbCrLf & _
                       "			end Cliente_Id," & vbCrLf & _
                       "			S.Safra_id," & vbCrLf & _
                       "			sum(cxs.AreaPlantada) AreaPlantada," & vbCrLf & _
                       "			CONVERT(numeric(18,4),0) as QtdeEmb" & vbCrLf & _
                       "	   from ClienteXSafra cxs" & vbCrLf & _
                       "	   inner Join #safras s" & vbCrLf & _
                       "		  on cxs.safra_id = s.safra_id" & vbCrLf

                If ddlCultura.SelectedIndex > 0 Then
                    sql &= "	   Where cxs.Cultura_Id = " & ddlCultura.SelectedValue & vbCrLf
                    Parametros &= "Cultura: " & ddlCultura.SelectedItem.Text & vbCrLf
                End If

                sql &= "	   group by case" & vbCrLf & _
                       "				  when len(cxs.Cliente_Id) = 14" & vbCrLf & _
                       "					then left(cxs.Cliente_Id,8)" & vbCrLf & _
                       "					else cxs.Cliente_Id" & vbCrLf & _
                       "				end," & vbCrLf & _
                       "				S.Safra_id" & vbCrLf & _
                       "      ) sbCultura" & vbCrLf & _
                       "   full join " & vbCrLf & _
                       "            (" & vbCrLf & _
                       "             Select case" & vbCrLf & _
                       "                      when Len(NF.Cliente_Id) = 14" & vbCrLf & _
                       "			            then LEFT(NF.Cliente_Id,8)" & vbCrLf & _
                       "			            else NF.Cliente_Id" & vbCrLf & _
                       "			        end Cliente_Id," & vbCrLf & _
                       "			        Ss.Safra_id," & vbCrLf & _
                       "			        CONVERT(numeric(18,4),0) as AreaPlantada," & vbCrLf & _
                       "                    sum(isnull(nfi.QtdeDeEmbalagem,0)) as QtdeEmb" & vbCrLf & _
                       "               from NotasFiscais AS NF" & vbCrLf & _
                       "			  INNER JOIN NotasFiscaisXItens AS NFI" & vbCrLf & _
                       "			 	 ON NF.Empresa_Id      = NFI.Empresa_Id" & vbCrLf & _
                       "			    AND NF.EndEmpresa_Id   = NFI.EndEmpresa_Id" & vbCrLf & _
                       "			    AND NF.Cliente_Id      = NFI.Cliente_Id" & vbCrLf & _
                       "			    AND NF.EndCliente_Id   = NFI.EndCliente_Id" & vbCrLf & _
                       "			    AND NF.EntradaSaida_Id = NFI.EntradaSaida_Id" & vbCrLf & _
                       "			    AND NF.Serie_Id        = NFI.Serie_Id" & vbCrLf & _
                       "			    AND NF.Nota_Id         = NFI.Nota_Id" & vbCrLf & _
                       "			  Inner Join Pedidos P" & vbCrLf & _
                       "				 on P.Empresa_id     = NF.Empresa_id" & vbCrLf & _
                       "			    and P.EndEmpresa_Id  = NF.EndEmpresa_Id" & vbCrLf & _
                       "			    And P.Pedido_Id      = NF.Pedido" & vbCrLf & _
                       "			  Inner Join #Safras Ss" & vbCrLf & _
                       "			     on Ss.safra_id = P.safra" & vbCrLf & _
                       "			   Inner Join Produtos Prd" & vbCrLf & _
                       "				 on NFI.Produto_Id = Prd.Produto_Id" & vbCrLf & _
                       "			  INNER JOIN SubOperacoes SO" & vbCrLf & _
                       "		 		 ON NFI.Operacao    = SO.Operacao_Id" & vbCrLf & _
                       "			    AND NFI.SubOperacao = SO.SubOperacoes_Id" & vbCrLf & _
                       "			  INNER JOIN Operacoes OP" & vbCrLf & _
                       "			 	 ON OP.Operacao_Id = SO.Operacao_Id" & vbCrLf & _
                       "			  Where NF.Situacao      = 1" & vbCrLf & _
                       "			    and SO.Classe   <> '" & eClassesOperacoes.GLOBAL.ToString & "'" & vbCrLf & _
                       "			    and OP.Classe   = '" & eClassesOperacoes.VENDAS.ToString & "'" & vbCrLf

                Dim emp As String() = ddlEmpresa.SelectedValue.Split("-")
                If chkConsolidarEmpresa.Checked Then
                    sql &= "			    and left(NF.Empresa_Id,8) ='" & emp(0).Substring(0, 8) & "'" & vbCrLf
                Else
                    sql &= "			    and NF.Empresa_Id    ='" & emp(0) & "'" & vbCrLf & _
                           "			    and NF.EndEmpresa_id = " & emp(1) & vbCrLf
                End If

                If ddlMarcaProd.SelectedIndex > 0 Then
                    sql &= "			    and Prd.Marca = " & ddlMarcaProd.SelectedValue
                    Parametros &= "Marca: " & ddlMarcaProd.SelectedItem.Text & vbCrLf
                End If

                If ucSelecaoProduto.TemSelecionado Then
                    Dim ret As ArrayList
                    ret = ucSelecaoProduto.GetSqlEParametrosRelatorio("Prd.Grupo", "Prd.Produto_Id")
                    sql &= "			    And " & ret(0)
                    Parametros &= "Produtos: " & ret(1)
                End If

                sql &= "			  group by case" & vbCrLf & _
                       "						when Len(NF.Cliente_Id) = 14" & vbCrLf & _
                       "						 then LEFT(NF.Cliente_Id,8)" & vbCrLf & _
                       "						 else NF.Cliente_Id" & vbCrLf & _
                       "					  end," & vbCrLf & _
                       "					  Ss.Safra_id" & vbCrLf & _
                       "     ) sbNotas" & vbCrLf & _
                       "   on sbCultura.cliente_id = sbNotas.Cliente_Id " & vbCrLf & _
                       "  and sbCultura.Safra_id   = sbNotas.Safra_id;" & vbCrLf


                sql &= "Select Cliente_Id," & vbCrLf & _
                       "       SUM(case" & vbCrLf & _
                       "             when Safra = @Safra01" & vbCrLf & _
                       "              then AreaPlantada" & vbCrLf & _
                       "              else 0" & vbCrLf & _
                       "           end) AreaSafra01," & vbCrLf & _
                       "       SUM(case" & vbCrLf & _
                       "             when Safra = @Safra01" & vbCrLf & _
                       "              then QtdeEmb" & vbCrLf & _
                       "              else 0" & vbCrLf & _
                       "           end) QtdeSafra01," & vbCrLf & _
                       "       SUM(case" & vbCrLf & _
                       "             when Safra = @Safra02" & vbCrLf & _
                       "              then AreaPlantada" & vbCrLf & _
                       "              else 0" & vbCrLf & _
                       "           end) AreaSafra02," & vbCrLf & _
                       "       SUM(case" & vbCrLf & _
                       "             when Safra = @Safra02" & vbCrLf & _
                       "              then QtdeEmb" & vbCrLf & _
                       "              else 0" & vbCrLf & _
                       "           end) QtdeSafra02," & vbCrLf & _
                       "       SUM(case" & vbCrLf & _
                       "             when Safra = @Safra03" & vbCrLf & _
                       "              then AreaPlantada" & vbCrLf & _
                       "              else 0" & vbCrLf & _
                       "           end) AreaSafra03," & vbCrLf & _
                       "       SUM(case" & vbCrLf & _
                       "             when Safra = @Safra03" & vbCrLf & _
                       "              then QtdeEmb" & vbCrLf & _
                       "              else 0" & vbCrLf & _
                       "           end) QtdeSafra03" & vbCrLf & _
                       "  into #Retencao" & vbCrLf & _
                       "  from #temp" & vbCrLf & _
                       " group by Cliente_Id;" & vbCrLf

                sql &= "Select Cliente_id," & vbCrLf & _
                       "	  (select top 1 Nome" & vbCrLf & _
                       "	     from Clientes c" & vbCrLf & _
                       "	    where case" & vbCrLf & _
                       "	            when Len(c.Cliente_Id) = 14" & vbCrLf & _
                       "	              then left(c.cliente_id,8)" & vbCrLf & _
                       "	              else c.Cliente_Id" & vbCrLf & _
                       "	           end        = #Retencao.Cliente_Id" & vbCrLf & _
                       "	   ) as nome," & vbCrLf & _
                       "       AreaSafra01," & vbCrLf & _
                       "       QtdeSafra01," & vbCrLf & _
                       "       case" & vbCrLf & _
                       "         when AreaSafra01 = 0" & vbCrLf & _
                       "          then 0" & vbCrLf & _
                       "          else convert(numeric(18,2),round(QtdeSafra01 * 100 / AreaSafra01,2))" & vbCrLf & _
                       "       End ShareSafra01," & vbCrLf & _
                       "       AreaSafra02," & vbCrLf & _
                       "       QtdeSafra02," & vbCrLf & _
                       "       case" & vbCrLf & _
                       "         when AreaSafra02 = 0" & vbCrLf & _
                       "          then 0" & vbCrLf & _
                       "          else convert(numeric(18,2),round(QtdeSafra02 * 100 / AreaSafra02,2))" & vbCrLf & _
                       "       End ShareSafra02," & vbCrLf & _
                       "       AreaSafra03," & vbCrLf & _
                       "       QtdeSafra03," & vbCrLf & _
                       "       case" & vbCrLf & _
                       "         when AreaSafra03 = 0" & vbCrLf & _
                       "          then 0" & vbCrLf & _
                       "          else convert(numeric(18,2),round(QtdeSafra03 * 100 / AreaSafra03,2))" & vbCrLf & _
                       "       End ShareSafra03" & vbCrLf & _
                       "  from #Retencao" & vbCrLf & _
                       " Order by 2;" & vbCrLf

                sql &= "select Sum(AreaSafra01) as AreaSafra01," & vbCrLf & _
                       "       sum(QtdeSafra01) as QtdeSafra01," & vbCrLf & _
                       "       case" & vbCrLf & _
                       "         when sum(AreaSafra01) = 0" & vbCrLf & _
                       "          then 0" & vbCrLf & _
                       "          else convert(numeric(18,2),round(sum(QtdeSafra01) * 100 / sum(AreaSafra01),2))" & vbCrLf & _
                       "       End CoberturaSafra01," & vbCrLf & _
                       "       sum(case" & vbCrLf & _
                       "             When AreaSafra01 > 0" & vbCrLf & _
                       "               then 1" & vbCrLf & _
                       "               else 0" & vbCrLf & _
                       "           end) PlantioSafra01," & vbCrLf & _
                       "       sum(case" & vbCrLf & _
                       "             When QtdeSafra01 > 0" & vbCrLf & _
                       "               then 1" & vbCrLf & _
                       "               else 0" & vbCrLf & _
                       "           end) VendasSafra01," & vbCrLf & _
                       "       case" & vbCrLf & _
                       "         when sum(AreaSafra01) = 0" & vbCrLf & _
                       "          then 0" & vbCrLf & _
                       "          else convert(numeric(18,2)," & vbCrLf & _
                       "                       (sum(case" & vbCrLf & _
                       "							 When QtdeSafra01 > 0" & vbCrLf & _
                       "							   then 1.00" & vbCrLf & _
                       "							   else 0.00" & vbCrLf & _
                       "						   end)" & vbCrLf & _
                       "					   /" & vbCrLf & _
                       "					   sum(case" & vbCrLf & _
                       "							 When AreaSafra01 > 0" & vbCrLf & _
                       "							   then 1.00" & vbCrLf & _
                       "							   else 0.00" & vbCrLf & _
                       "						   end)" & vbCrLf & _
                       "					   ) * 100" & vbCrLf & _
                       "                       )" & vbCrLf & _
                       "       End AbrangenciaSafra01," & vbCrLf
                '*************************************************************** SAFRA 02 ************************************************************************
                sql &= "       Sum(AreaSafra02) as AreaSafra02," & vbCrLf & _
                       "       sum(QtdeSafra02) as QtdeSafra02," & vbCrLf & _
                       "       case" & vbCrLf & _
                       "         when sum(AreaSafra02) = 0" & vbCrLf & _
                       "          then 0" & vbCrLf & _
                       "          else convert(numeric(18,2),round(sum(QtdeSafra02) * 100 / sum(AreaSafra02),2))" & vbCrLf & _
                       "       End CoberturaSafra02," & vbCrLf & _
                       "       sum(case" & vbCrLf & _
                       "             When AreaSafra02 > 0" & vbCrLf & _
                       "               then 1" & vbCrLf & _
                       "               else 0" & vbCrLf & _
                       "           end) PlantioSafra02," & vbCrLf & _
                       "       sum(case" & vbCrLf & _
                       "             When QtdeSafra02 > 0" & vbCrLf & _
                       "               then 1" & vbCrLf & _
                       "               else 0" & vbCrLf & _
                       "           end) VendasSafra02," & vbCrLf & _
                       "       case" & vbCrLf & _
                       "         when sum(AreaSafra02) = 0" & vbCrLf & _
                       "          then 0" & vbCrLf & _
                       "          else convert(numeric(18,2)," & vbCrLf & _
                       "                       (sum(case" & vbCrLf & _
                       "							 When QtdeSafra02 > 0" & vbCrLf & _
                       "							   then 1.00" & vbCrLf & _
                       "							   else 0.00" & vbCrLf & _
                       "						   end)" & vbCrLf & _
                       "					   /" & vbCrLf & _
                       "					   sum(case" & vbCrLf & _
                       "							 When AreaSafra02 > 0" & vbCrLf & _
                       "							   then 1.00" & vbCrLf & _
                       "							   else 0.00" & vbCrLf & _
                       "						   end)" & vbCrLf & _
                       "					   ) * 100" & vbCrLf & _
                       "                       )" & vbCrLf & _
                       "       End AbrangenciaSafra02," & vbCrLf & _
                       "       sum(case" & vbCrLf & _
                       "             When QtdeSafra01 > 0 and QtdeSafra02 > 0" & vbCrLf & _
                       "               then 1" & vbCrLf & _
                       "               else 0" & vbCrLf & _
                       "           end) RetidosSafra02," & vbCrLf & _
                       "       sum(case" & vbCrLf & _
                       "             When QtdeSafra01 > 0 and QtdeSafra02 = 0" & vbCrLf & _
                       "               then 1" & vbCrLf & _
                       "               else 0" & vbCrLf & _
                       "           end) PerdidosSafra02," & vbCrLf & _
                       "       sum(case" & vbCrLf & _
                       "             When QtdeSafra01 = 0 and QtdeSafra02 > 0" & vbCrLf & _
                       "               then 1" & vbCrLf & _
                       "               else 0" & vbCrLf & _
                       "           end) NovosSafra02," & vbCrLf & _
                       "       sum(case" & vbCrLf & _
                       "             when" & vbCrLf & _
                       "			   case" & vbCrLf & _
                       "				 when AreaSafra01 = 0" & vbCrLf & _
                       "				  then 0" & vbCrLf & _
                       "				  else convert(numeric(18,2),round(QtdeSafra01 * 100 / AreaSafra01,2))" & vbCrLf & _
                       "			   End" & vbCrLf & _
                       "			   >" & vbCrLf & _
                       "			   case" & vbCrLf & _
                       "				 when AreaSafra02 = 0" & vbCrLf & _
                       "				  then 0" & vbCrLf & _
                       "				  else convert(numeric(18,2),round(QtdeSafra02 * 100 / AreaSafra02,2))" & vbCrLf & _
                       "			   End" & vbCrLf & _
                       "		     then 1" & vbCrLf & _
                       "		     else 0" & vbCrLf & _
                       "		   end) DimShare02," & vbCrLf & _
                       "       sum(case" & vbCrLf & _
                       "             when" & vbCrLf & _
                       "			   case" & vbCrLf & _
                       "				 when AreaSafra01 = 0" & vbCrLf & _
                       "				  then 0" & vbCrLf & _
                       "				  else convert(numeric(18,2),round(QtdeSafra01 * 100 / AreaSafra01,2))" & vbCrLf & _
                       "			   End" & vbCrLf & _
                       "			   <" & vbCrLf & _
                       "			   case" & vbCrLf & _
                       "				 when AreaSafra02 = 0" & vbCrLf & _
                       "				  then 0" & vbCrLf & _
                       "				  else convert(numeric(18,2),round(QtdeSafra02 * 100 / AreaSafra02,2))" & vbCrLf & _
                       "			   End" & vbCrLf & _
                       "		     then 1" & vbCrLf & _
                       "		     else 0" & vbCrLf & _
                       "		   end) AumShare02," & vbCrLf & _
                       "	   convert(numeric(5,2), " & vbCrLf & _
                       "                       	   (sum(case" & vbCrLf & _
                       "                                  When QtdeSafra01 > 0 and QtdeSafra02 > 0" & vbCrLf & _
                       "                                    then 1.00" & vbCrLf & _
                       "                                    else 0.00" & vbCrLf & _
                       "                                 end)" & vbCrLf & _
                       "                            /" & vbCrLf & _
                       "              	            sum(case" & vbCrLf & _
                       "                                  When QtdeSafra01 > 0" & vbCrLf & _
                       "                                    then 1.00" & vbCrLf & _
                       "                                    else 0.00" & vbCrLf & _
                       "                                end)" & vbCrLf & _
                       "                            ) * 100) RetencaoPerc02," & vbCrLf

                '*************************************************************** SAFRA 03 ************************************************************************
                sql &= "       Sum(AreaSafra03) as AreaSafra03," & vbCrLf & _
                       "       sum(QtdeSafra03) as QtdeSafra03," & vbCrLf & _
                       "       case" & vbCrLf & _
                       "         when sum(AreaSafra03) = 0" & vbCrLf & _
                       "          then 0" & vbCrLf & _
                       "          else convert(numeric(18,2),round(sum(QtdeSafra03) * 100 / sum(AreaSafra03),2))" & vbCrLf & _
                       "       End CoberturaSafra03," & vbCrLf & _
                       "       sum(case" & vbCrLf & _
                       "             When AreaSafra03 > 0" & vbCrLf & _
                       "               then 1" & vbCrLf & _
                       "               else 0" & vbCrLf & _
                       "           end) PlantioSafra03," & vbCrLf & _
                       "       sum(case" & vbCrLf & _
                       "             When QtdeSafra03 > 0" & vbCrLf & _
                       "               then 1" & vbCrLf & _
                       "               else 0" & vbCrLf & _
                       "           end) VendasSafra03," & vbCrLf & _
                       "       case" & vbCrLf & _
                       "         when sum(AreaSafra03) = 0" & vbCrLf & _
                       "          then 0" & vbCrLf & _
                       "          else convert(numeric(18,2)," & vbCrLf & _
                       "                       (sum(case" & vbCrLf & _
                       "							 When QtdeSafra03 > 0" & vbCrLf & _
                       "							   then 1.00" & vbCrLf & _
                       "							   else 0.00" & vbCrLf & _
                       "						   end)" & vbCrLf & _
                       "					   /" & vbCrLf & _
                       "					   sum(case" & vbCrLf & _
                       "							 When AreaSafra03 > 0" & vbCrLf & _
                       "							   then 1.00" & vbCrLf & _
                       "							   else 0.00" & vbCrLf & _
                       "						   end)" & vbCrLf & _
                       "					   ) * 100" & vbCrLf & _
                       "                       )" & vbCrLf & _
                       "       End AbrangenciaSafra03," & vbCrLf & _
                       "       sum(case" & vbCrLf & _
                       "             When QtdeSafra02 > 0 and QtdeSafra03 > 0" & vbCrLf & _
                       "               then 1" & vbCrLf & _
                       "               else 0" & vbCrLf & _
                       "           end) RetidosSafra03," & vbCrLf & _
                       "       sum(case" & vbCrLf & _
                       "             When QtdeSafra02 > 0 and QtdeSafra03 = 0" & vbCrLf & _
                       "               then 1" & vbCrLf & _
                       "               else 0" & vbCrLf & _
                       "           end) PerdidosSafra03," & vbCrLf & _
                       "       sum(case" & vbCrLf & _
                       "             When QtdeSafra02 = 0 and QtdeSafra03 > 0" & vbCrLf & _
                       "               then 1" & vbCrLf & _
                       "               else 0" & vbCrLf & _
                       "           end) NovosSafra03," & vbCrLf & _
                       "       sum(case" & vbCrLf & _
                       "             when" & vbCrLf & _
                       "			   case" & vbCrLf & _
                       "				 when AreaSafra02 = 0" & vbCrLf & _
                       "				  then 0" & vbCrLf & _
                       "				  else convert(numeric(18,2),round(QtdeSafra02 * 100 / AreaSafra02,2))" & vbCrLf & _
                       "			   End" & vbCrLf & _
                       "			   >" & vbCrLf & _
                       "			   case" & vbCrLf & _
                       "				 when AreaSafra03 = 0" & vbCrLf & _
                       "				  then 0" & vbCrLf & _
                       "				  else convert(numeric(18,2),round(QtdeSafra03 * 100 / AreaSafra03,2))" & vbCrLf & _
                       "			   End" & vbCrLf & _
                       "		     then 1" & vbCrLf & _
                       "		     else 0" & vbCrLf & _
                       "		   end) DimShare03," & vbCrLf & _
                       "       sum(case" & vbCrLf & _
                       "             when" & vbCrLf & _
                       "			   case" & vbCrLf & _
                       "				 when AreaSafra02 = 0" & vbCrLf & _
                       "				  then 0" & vbCrLf & _
                       "				  else convert(numeric(18,2),round(QtdeSafra02 * 100 / AreaSafra02,2))" & vbCrLf & _
                       "			   End" & vbCrLf & _
                       "			   <" & vbCrLf & _
                       "			   case" & vbCrLf & _
                       "				 when AreaSafra03 = 0" & vbCrLf & _
                       "				  then 0" & vbCrLf & _
                       "				  else convert(numeric(18,2),round(QtdeSafra03 * 100 / AreaSafra03,2))" & vbCrLf & _
                       "			   End" & vbCrLf & _
                       "		     then 1" & vbCrLf & _
                       "		     else 0" & vbCrLf & _
                       "		   end) AumShare03," & vbCrLf & _
                       "	   convert(numeric(5,2), " & vbCrLf & _
                       "	   (sum(case" & vbCrLf & _
                       "              When QtdeSafra02 > 0 and QtdeSafra03 > 0" & vbCrLf & _
                       "                then 1.00" & vbCrLf & _
                       "                else 0.00" & vbCrLf & _
                       "            end)" & vbCrLf & _
                       "        /" & vbCrLf & _
                       "	    sum(case" & vbCrLf & _
                       "             When QtdeSafra02 > 0" & vbCrLf & _
                       "               then 1.00" & vbCrLf & _
                       "               else 0.00" & vbCrLf & _
                       "           end)" & vbCrLf & _
                       "        ) * 100) RetencaoPerc03" & vbCrLf & _
                       " from #Retencao" & vbCrLf

                Dim ds As DataSet

                ds = Banco.ConsultaDataSet(sql, "Retencao")
                ds.Tables(0).TableName = "CustomerShare"
                ds.Tables(1).TableName = "Retencao"

                ds.Tables.Add("Logo")

                Dim strCaminhoImagem As String = Server.MapPath("~/Images/" & HttpContext.Current.Session("ssImagemEmpresa"))
                ds.Tables("Logo").Columns.Add("Imagem", GetType(System.Byte()))
                ds.Tables("Logo").Columns.Add("NomeEmp", GetType(String))
                ds.Tables("Logo").Columns.Add("CidadeEstado", GetType(String))

                Dim row As DataRow = ds.Tables("Logo").NewRow()
                row("Imagem") = [Lib].Negocio.Conversoes.ConverterImagemEmByte(strCaminhoImagem)
                row("NomeEmp") = HttpContext.Current.Session("ssNomeEmpresa") & " (" & Funcoes.FormatarCpfCnpj(HttpContext.Current.Session("ssEmpresa")) & ")"
                row("CidadeEstado") = HttpContext.Current.Session("ssCidadeEmpresa") & "/" & HttpContext.Current.Session("ssEstadoEmpresa")
                ds.Tables("Logo").Rows.Add(row)

                rpt.FileName = HttpContext.Current.Server.MapPath("~/Reports/Cr_RelatorioDeRetencaoDeClientes.rpt")
                rpt.Load(rpt.FileName, OpenReportMethod.OpenReportByDefault)

                Dim UrlArquivo As String = "Files/" & Funcoes.GeraNomeArquivo & ".PDF"
                Dim NomeArquivo As String = HttpContext.Current.Server.MapPath(UrlArquivo)

                rpt.SetDataSource(ds)

                Dim param As New Dictionary(Of String, Object)
                param.Add("ParamConsulta", Parametros)
                param.Add("safra01", ddlSafra01.SelectedValue)
                param.Add("safra02", ddlSafra02.SelectedValue)
                param.Add("safra03", ddlSafra03.SelectedValue)
                param.Add("@Safra01", ddlSafra01.SelectedValue)
                param.Add("@Safra02", ddlSafra02.SelectedValue)
                param.Add("@Safra03", ddlSafra03.SelectedValue)
                Funcoes.BindParameters(rpt, param)

                If Dir(NomeArquivo).Length > 0 Then Kill(NomeArquivo)

                rpt.ExportToDisk(ExportFormatType.PortableDocFormat, NomeArquivo)

                If System.IO.File.Exists(NomeArquivo) Then
                    Funcoes.AbrirArquivo(Page, UrlArquivo)
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para emitir relatório.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        Finally
            rpt.Close()
            rpt.Dispose()
        End Try
    End Sub

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "RelatorioDeRetencaoDeClientes")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub
End Class