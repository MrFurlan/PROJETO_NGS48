Imports System.Data
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis
Imports System.IO
Imports OfficeOpenXml

Public Class RelatorioEstoqueRevenda
    Inherits BasePage

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.setMenu(eModulo.Expedicao)
            If Not IsPostBack And IsConnect Then
                If Funcoes.VerificaPermissao("RelatorioEstoqueRevenda", "ACESSAR") Then
                    ddl.Carregar(lstEmpresa, CarregarDDL.Tabela.ClientesXEmpresas, "", False)
                    ddl.Carregar(ddlMarca, CarregarDDL.Tabela.Marca, "", True)
                    ddl.Carregar(ddlSafra, CarregarDDL.Tabela.Safra, "", True)
                    Limpar()
                Else
                    MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Expedicao.aspx")
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Private Sub Limpar()
        ddlMarca.SelectedIndex = 0
        txtDataInicial.Text = "01/01/" & DateTime.Now.Year.ToString()
        txtDataFinal.Text = DateTime.Now.ToShortDateString()
    End Sub

    Private Sub EmitirRelatorio(ByVal Pdf As Boolean)
        Try
            Dim ds As DataSet = Nothing
            Dim sql As String = ""
            Dim Parametros As String = "Parametros:  Empresa(s): " & String.Join(", ", lstEmpresa.GetSelectedValues().ToArray()) & vbCrLf

            Dim objSafra As [Lib].Negocio.Safra = Nothing
            If ddlSafra.SelectedIndex > 0 Then
                objSafra = New [Lib].Negocio.Safra(ddlSafra.SelectedValue)
                Parametros &= "Safra Final: " & ddlSafra.SelectedItem.Text & vbCrLf
            End If


            sql = ""
            If ddlSafra.SelectedIndex > 0 Then
                sql = "select Safra_Id " & vbCrLf & _
                        "  into #Safra" & vbCrLf & _
                        "  from Safras S" & vbCrLf & _
                        " where S.Vencimento <= '" & objSafra.Vencimento.ToString("yyyy-MM-dd") & "'" & vbCrLf
            End If


            sql &= "Select Ge.Descricao as DescGrupo" & vbCrLf & _
                  "       ,P.Produto_Id as CodigoProduto " & vbCrLf & _
                  "       ,P.Nome as NomeProduto " & vbCrLf & _
                  "       ,P.Unidade " & vbCrLf

            If chkMapa.Checked AndAlso chkLoteSemente.Checked Then
                sql &= ""
            ElseIf chkLoteSemente.Checked Then
                sql &= "       ,'' as Lote " & vbCrLf & _
                       "       ,'' as Classificacao " & vbCrLf
            Else
                sql &= "       ,isnull(sb.Lote,'') as Lote " & vbCrLf & _
                       "       ,isnull(sb.Classificacao,'') as Classificacao" & vbCrLf
            End If

            If chkMapa.Checked AndAlso chkEmbalagens.Checked Then
                sql &= ""
            ElseIf chkEmbalagens.Checked Then
                sql &= "       ,0 as Embalagem " & vbCrLf & _
                       "       ,'' TipoDeEmbalagem " & vbCrLf & _
                       "       ,0 as CapacidadeEmbalagem " & vbCrLf
            Else
                sql &= "       ,isnull(sb.Embalagem,0) as Embalagem " & vbCrLf & _
                       "       ,isnull(sb.TipoDeEmbalagem,'') as TipoDeEmbalagem " & vbCrLf & _
                       "       ,ISNULL(sb.CapacidadeEmbalagem, 0) AS CapacidadeEmbalagem  "
            End If

            If chkMapa.Checked Then
                Dim EmpSel As ArrayList = lstEmpresa.GetSelectedValues()
                For i As Integer = 0 To EmpSel.Count - 1
                    Dim Emp() As String = EmpSel(i).Split("-")

                    If chkSaldoAnterior.Checked Then sql &= ",sum(case when Sb.Empresa_ID = '" & Emp(0) & "' and sb.endempresa_id = " & Emp(1) & " then sb.SaldoAnterior else 0 end) 'SaldoAnterior-" & Emp(0) & "'" & vbCrLf
                    If ChkEntradas.Checked Then sql &= ",sum(case when Sb.Empresa_ID = '" & Emp(0) & "' and sb.endempresa_id = " & Emp(1) & " then sb.EQtde else 0 end) 'EQtde-" & Emp(0) & "'" & vbCrLf
                    If ChkSaidas.Checked Then sql &= ",sum(case when Sb.Empresa_ID = '" & Emp(0) & "' and sb.endempresa_id = " & Emp(1) & " then sb.SQtde else 0 end) 'SQtde-" & Emp(0) & "'" & vbCrLf
                    If ChkConsignado.Checked Then sql &= ",sum(case when Sb.Empresa_ID = '" & Emp(0) & "' and sb.endempresa_id = " & Emp(1) & " then sb.SaldoConsignado else 0 end) 'SaldoConsignado-" & Emp(0) & "'" & vbCrLf
                    If ChkAmostraGratis.Checked Then sql &= ",sum(case when Sb.Empresa_ID = '" & Emp(0) & "' and sb.endempresa_id = " & Emp(1) & " then sb.SaldoAmostraGratis else 0 end) 'SaldoAmostraGratis-" & Emp(0) & "'" & vbCrLf
                    If ChkTerceiros.Checked Then sql &= ",sum(case when Sb.Empresa_ID = '" & Emp(0) & "' and sb.endempresa_id = " & Emp(1) & " then sb.ProdDeTerc else 0 end) 'SaldoTerceiros-" & Emp(0) & "'" & vbCrLf
                    If ChkSaldoProprio.Checked Then sql &= ",sum(case when Sb.Empresa_ID = '" & Emp(0) & "' and sb.endempresa_id = " & Emp(1) & " then sb.SaldoProprio else 0 end) 'SaldoProprio-" & Emp(0) & "'" & vbCrLf
                    If ChkSaldoFora.Checked Then sql &= ",abs(sum(case when Sb.Empresa_ID = '" & Emp(0) & "' and sb.endempresa_id = " & Emp(1) & " then sb.SaldoEmPoderDeTerceiros else 0 end)) 'SaldoEmPoderDeTerceiros-" & Emp(0) & "'" & vbCrLf
                    If ChkFatNaoEntregue.Checked Then sql &= ",max(case when Sb.Empresa_ID = '" & Emp(0) & "' and sb.endempresa_id = " & Emp(1) & " then PosicaoPedido.FaturadaNaoEntregue else 0 end) 'FaturadaNaoEntregue-" & Emp(0) & "'" & vbCrLf
                    If ChkCompra.Checked Then sql &= ",max(isnull(case when Sb.Empresa_ID = '" & Emp(0) & "' and sb.endempresa_id = " & Emp(1) & " then PosicaoPedido.QtdCompra else 0 end,0)) 'QtdCompra-" & Emp(0) & "'" & vbCrLf
                    If ChkVenda.Checked Then sql &= ",max(isnull(case when Sb.Empresa_ID = '" & Emp(0) & "' and sb.endempresa_id = " & Emp(1) & " then PosicaoPedido.QtdVenda else 0 end,0)) 'QtdVenda-" & Emp(0) & "'" & vbCrLf
                    If ChkPrevisao.Checked Then sql &= ",sum(isnull(case when Sb.Empresa_ID = '" & Emp(0) & "' and sb.endempresa_id = " & Emp(1) & " then sb.SaldoProprio else 0 end,0)) + abs(sum(isnull(case when Sb.Empresa_ID = '" & Emp(0) & "' and sb.endempresa_id = " & Emp(1) & " then sb.SaldoEmPoderDeTerceiros else 0 end,0)))  + max(isnull(case when Sb.Empresa_ID = '" & Emp(0) & "' and sb.endempresa_id = " & Emp(1) & " then PosicaoPedido.QtdCompra else 0 end,0)) - max(isnull(case when Sb.Empresa_ID = '" & Emp(0) & "' and sb.endempresa_id = " & Emp(1) & " then PosicaoPedido.QtdVenda else 0 end,0)) - max(case when Sb.Empresa_ID = '" & Emp(0) & "' and sb.endempresa_id = " & Emp(1) & " then PosicaoPedido.FaturadaNaoEntregue else 0 end) as Previsao" & vbCrLf
                Next

            Else
                sql &= "      ,sum(sb.SaldoAnterior) as SaldoAnterior " & vbCrLf & _
                       "      ,sum(sb.EQtde) EQtde " & vbCrLf & _
                       "      ,sum(sb.SQtde) SQtde " & vbCrLf & _
                       "      ,sum(sb.SaldoConsignado) SaldoConsignado" & vbCrLf & _
                       "      ,SUM(sb.SaldoAmostraGratis) SaldoAmostraGratis" & vbCrLf & _
                       "      ,sum(sb.ProdDeTerc)  SaldoTerceiros" & vbCrLf & _
                       "      ,sum(sb.SaldoProprio) SaldoProprio" & vbCrLf & _
                       "      ,abs(sum(sb.SaldoEmPoderDeTerceiros)) SaldoEmPoderDeTerceiros" & vbCrLf & _
                       "      ,max(PosicaoPedido.FaturadaNaoEntregue) as FaturadoNaoEntregue" & vbCrLf & _
                       "      ,max(isnull(PosicaoPedido.QtdCompra,0)) PedCompra " & vbCrLf & _
                       "      ,max(isnull(PosicaoPedido.QtdVenda,0)) PedVenda " & vbCrLf & _
                       "      ,sum(isnull(sb.SaldoProprio,0)) + abs(sum(isnull(sb.SaldoEmPoderDeTerceiros,0)))  + max(isnull(PosicaoPedido.QtdCompra,0)) - max(isnull(PosicaoPedido.QtdVenda,0)) - max(PosicaoPedido.FaturadaNaoEntregue) as Previsao" & vbCrLf & _
                       "      ,convert(nvarchar,'')      as NomeMoeda" & vbCrLf & _
                       "      ,convert(nvarchar,'')      as Cifrao" & vbCrLf & _
                       "      ,convert(numeric(18,6),0) as ValorProduto" & vbCrLf & _
                       "      ,convert(numeric(18,6),0) as VTotal" & vbCrLf & _
                       "      ,convert(numeric(18,6),0) as VTotalPrevisao " & vbCrLf
            End If

            If RdEstoqueComPreco.Checked Or rbPreco.Checked Then sql &= " into #Temp " & vbCrLf

            sql &= "  From (select M.Empresa_Id, M.EndEmpresa_Id, M.Produto_id, M.Lote, M.Classificacao, M.Embalagem, M.TipoDeEmbalagem, M.CapacidadeEmbalagem," & vbCrLf & _
                   "               Sum(M.SaldoAnterior) as SaldoAnterior," & vbCrLf & _
                   "               sum(M.EQtde) as EQtde," & vbCrLf & _
                   "               sum(M.SQtde) as SQtde," & vbCrLf & _
                   "               sum(M.ProdDeTerc) as ProdDeTerc," & vbCrLf & _
                   "               sum(M.SaldoProprio) as SaldoProprio," & vbCrLf & _
                   "               sum(M.SaldoEmPoderDeTerceiros) as SaldoEmPoderDeTerceiros," & vbCrLf & _
                   "               sum(M.SaldoConsignado) as SaldoConsignado," & vbCrLf & _
                   "               SUM(M.SaldoAmostraGratis) as SaldoAmostraGratis" & vbCrLf & _
                   "          from ( " & vbCrLf & _
                   "		        SELECT NotasFiscaisXItens.Empresa_Id," & vbCrLf & _
                   "                       NotasFiscaisXItens.EndEmpresa_Id," & vbCrLf & _
                   "                       NotasFiscaisXItens.Produto_Id, " & vbCrLf & _
                   "			           isnull(NotasFiscaisXItens.Lote,'') as Lote, " & vbCrLf & _
                   "			           isnull(NotasFiscaisXItens.Classificacao,'') as Classificacao, " & vbCrLf & _
                   "                       isnull(NotasFiscaisXItens.Embalagem,0) as Embalagem, " & vbCrLf & _
                   "                       isnull(NotasFiscaisXItens.TipoDeEmbalagem,'') as TipoDeEmbalagem, " & vbCrLf
            If chkEmbalagens.Checked Then
                sql &= "       ISNULL(NotasFiscaisXItens.CapacidadeEmbalagem, 0) AS CapacidadeEmbalagem,  " & vbCrLf
            Else
                sql &= "            CASE " & vbCrLf & _
                       "                WHEN PxE.PesoVariavel = 1" & vbCrLf & _
                       "                     THEN ISNULL(LxC.PesoSaco,0)" & vbCrLf & _
                       "                ELSE ISNULL(NotasFiscaisXItens.CapacidadeEmbalagem, 0)" & vbCrLf & _
                       "            END as CapacidadeEmbalagem, " & vbCrLf
            End If

            sql &= "                       sum(case " & vbCrLf & _
                   "                             when NotasFiscais.Movimento < '" & txtDataInicial.Text.ToSqlDate() & "' " & vbCrLf & _
                   "                               then case " & vbCrLf & _
                   "					                  when Suboperacoes.EntradaSaida = 'E' " & vbCrLf & _
                   "					                    then NotasFiscaisXItens.QuantidadeFisica " & vbCrLf & _
                   "					                    else  NotasFiscaisXItens.QuantidadeFisica * - 1 " & vbCrLf & _
                   "				                    end " & vbCrLf & _
                   "                               else 0 " & vbCrLf & _
                   "                           end) SaldoAnterior, " & vbCrLf & _
                   "                       sum(case " & vbCrLf & _
                   "                             when NotasFiscais.Movimento  between '" & txtDataInicial.Text.ToSqlDate() & "' and '" & txtDataFinal.Text.ToSqlDate() & "' " & vbCrLf & _
                   "                               then case " & vbCrLf & _
                   "					                  when Suboperacoes.EntradaSaida = 'E' " & vbCrLf & _
                   "					                    then NotasFiscaisXItens.QuantidadeFisica " & vbCrLf & _
                   "					                    else 0 " & vbCrLf & _
                   "				                    end " & vbCrLf & _
                   "                               else 0 " & vbCrLf & _
                   "                           end) EQtde, " & vbCrLf & _
                   "                       sum(case " & vbCrLf & _
                   "                             when NotasFiscais.Movimento  between '" & txtDataInicial.Text.ToSqlDate() & "' and '" & txtDataFinal.Text.ToSqlDate() & "' " & vbCrLf & _
                   "                               then case " & vbCrLf & _
                   "					                  when Suboperacoes.EntradaSaida = 'S' " & vbCrLf & _
                   " 					                    then NotasFiscaisXItens.QuantidadeFisica " & vbCrLf & _
                   "  					                    else 0 " & vbCrLf & _
                   "				                    end " & vbCrLf & _
                   "                               else 0 " & vbCrLf & _
                   "                           end) SQtde, " & vbCrLf & _
                   "					  sum(case" & vbCrLf & _
                   "							 when NotasFiscais.Movimento <= '" & txtDataFinal.Text.ToSqlDate() & "'" & vbCrLf & _
                   "							   then case" & vbCrLf & _
                   "									 when (isnull(Suboperacoes.Consignacao,0) = 0 and ISNULL(Suboperacoes.AmostraGratis, 0) = 0   and Suboperacoes.EntradaSaida = 'E' and Suboperacoes.classe <> '" & eClassesOperacoes.DEPOSITOS.ToString & "') and Suboperacoes.ProdutoDeTerceiro = 0" & vbCrLf & _
                   "                                       or (isnull(Suboperacoes.Consignacao,0) = 1 and ISNULL(Suboperacoes.Devolucao, 'S')   = 'S' and Suboperacoes.EntradaSaida = 'E' and Suboperacoes.classe <> '" & eClassesOperacoes.DEPOSITOS.ToString & "') and Suboperacoes.ProdutoDeTerceiro = 0" & vbCrLf & _
                   "                                       or (isnull(Suboperacoes.Consignacao,0) = 0 and ISNULL(Suboperacoes.Devolucao, 'S')   = 'S' and Suboperacoes.EntradaSaida = 'E' and Suboperacoes.classe  = '" & eClassesOperacoes.DEPOSITOS.ToString & "') and Suboperacoes.ProdutoDeTerceiro = 0" & vbCrLf & _
                   "									   then NotasFiscaisXItens.QuantidadeFisica" & vbCrLf & _
                   "									 when (isnull(Suboperacoes.Consignacao,0) = 0 and ISNULL(Suboperacoes.AmostraGratis, 0) =  0  and Suboperacoes.EntradaSaida = 'S') and Suboperacoes.ProdutoDeTerceiro = 0" & vbCrLf & _
                   "                                       or (isnull(Suboperacoes.Consignacao,0) = 1 and ISNULL(Suboperacoes.Devolucao, 'S')   = 'N' and Suboperacoes.EntradaSaida = 'S') and Suboperacoes.ProdutoDeTerceiro = 0" & vbCrLf & _
                   "									   then NotasFiscaisXItens.QuantidadeFisica * - 1" & vbCrLf & _
                   "									 else 0" & vbCrLf & _
                   "									end" & vbCrLf & _
                   "							   else 0" & vbCrLf & _
                   "						   end) SaldoProprio," & vbCrLf & _
                   "					  sum(case" & vbCrLf & _
                   "							 when NotasFiscais.Movimento <= '" & txtDataFinal.Text.ToSqlDate() & "'" & vbCrLf & _
                   "							   then case" & vbCrLf & _
                   "									 when (isnull(Suboperacoes.Consignacao,0) = 0 and ISNULL(Suboperacoes.AmostraGratis, 0) = 0   and Suboperacoes.EntradaSaida = 'E' and Suboperacoes.classe = '" & eClassesOperacoes.DEPOSITOS.ToString & "') and Suboperacoes.deposito = 'N' and Suboperacoes.ProdutoDeTerceiro = 1" & vbCrLf & _
                   "                                       or (isnull(Suboperacoes.Consignacao,0) = 1 and ISNULL(Suboperacoes.Devolucao, 'S')   = 'S' and Suboperacoes.EntradaSaida = 'E' and Suboperacoes.classe = '" & eClassesOperacoes.DEPOSITOS.ToString & "') and Suboperacoes.deposito = 'N' and Suboperacoes.ProdutoDeTerceiro = 1" & vbCrLf & _
                   "									   then NotasFiscaisXItens.QuantidadeFisica" & vbCrLf & _
                   "									 when isnull(Suboperacoes.Consignacao,0) = 0 and ISNULL(Suboperacoes.AmostraGratis, 0) = 0 and Suboperacoes.EntradaSaida = 'S' and Suboperacoes.classe = '" & eClassesOperacoes.DEPOSITOS.ToString & "' and Suboperacoes.deposito = 'N' and Suboperacoes.ProdutoDeTerceiro = 1" & vbCrLf & _
                   "									   then NotasFiscaisXItens.QuantidadeFisica * - 1" & vbCrLf & _
                   "									 else 0" & vbCrLf & _
                   "									end" & vbCrLf & _
                   "							   else 0" & vbCrLf & _
                   "						   end) ProdDeTerc," & vbCrLf & _
                   "					  sum(case" & vbCrLf & _
                   "							 when NotasFiscais.Movimento <= '" & txtDataFinal.Text.ToSqlDate() & "'" & vbCrLf & _
                   "							   then case" & vbCrLf & _
                   "									 when (isnull(Suboperacoes.Consignacao,0) = 0 and ISNULL(Suboperacoes.AmostraGratis, 0) = 0 and SubOperacoes.Deposito = 'S' and Suboperacoes.EntradaSaida = 'E')" & vbCrLf & _
                   "                                       or (isnull(Suboperacoes.Consignacao,0) = 1  and SubOperacoes.Devolucao = 'N' and Suboperacoes.EntradaSaida = 'S')" & vbCrLf & _
                   "									   then NotasFiscaisXItens.QuantidadeFisica" & vbCrLf & _
                   "									 when (isnull(Suboperacoes.Consignacao,0) = 0 and ISNULL(Suboperacoes.AmostraGratis, 0) = 0 and SubOperacoes.Deposito = 'S' and Suboperacoes.EntradaSaida = 'S')" & vbCrLf & _
                   "                                       or (isnull(Suboperacoes.Consignacao,0) = 1  and SubOperacoes.Devolucao = 'S' and Suboperacoes.EntradaSaida = 'E')" & vbCrLf & _
                   "									   then NotasFiscaisXItens.QuantidadeFisica * - 1" & vbCrLf & _
                   "									 else 0" & vbCrLf & _
                   "									end" & vbCrLf & _
                   "							   else 0" & vbCrLf & _
                   "						   end) SaldoEmPoderDeTerceiros," & vbCrLf & _
                   "					  sum(case" & vbCrLf & _
                   "							 when NotasFiscais.Movimento <= '" & txtDataFinal.Text.ToSqlDate() & "'" & vbCrLf & _
                   "							   then case" & vbCrLf & _
                   "									 when isnull(Suboperacoes.Consignacao,0) = 1 and Suboperacoes.EntradaSaida = 'E' and Suboperacoes.Devolucao = 'N'" & vbCrLf & _
                   "									   then NotasFiscaisXItens.QuantidadeFisica" & vbCrLf & _
                   "									 when isnull(Suboperacoes.Consignacao,0) = 1 and Suboperacoes.EntradaSaida = 'S' and Suboperacoes.Devolucao = 'S'" & vbCrLf & _
                   "									   then NotasFiscaisXItens.QuantidadeFisica  * - 1" & vbCrLf & _
                   "									 else 0" & vbCrLf & _
                   "									end " & vbCrLf & _
                   "							   else 0 " & vbCrLf & _
                   "						   end) SaldoConsignado," & vbCrLf & _
                   "                      sum(case" & vbCrLf & _
                   "                     		when NotasFiscais.Movimento <= '" & txtDataFinal.Text.ToSqlDate() & "'" & vbCrLf & _
                   "                     		  then case" & vbCrLf & _
                   "                     			     when isnull(Suboperacoes.AmostraGratis,0) = 1 and Suboperacoes.EntradaSaida = 'E'" & vbCrLf & _
                   "                     				   then NotasFiscaisXItens.QuantidadeFisica" & vbCrLf & _
                   "                     				 when isnull(Suboperacoes.AmostraGratis,0) = 1 and Suboperacoes.EntradaSaida = 'S'" & vbCrLf & _
                   "                     				   then NotasFiscaisXItens.QuantidadeFisica  * - 1" & vbCrLf & _
                   "                     				 else 0" & vbCrLf & _
                   "                                   End" & vbCrLf & _
                   "                     		  else 0" & vbCrLf & _
                   "                     	  end) SaldoAmostraGratis" & vbCrLf & _
                   "		         FROM NotasFiscais " & vbCrLf & _
                   "		        INNER JOIN NotasFiscaisXItens " & vbCrLf & _
                   "			       ON NotasFiscais.Empresa_Id      = NotasFiscaisXItens.Empresa_Id " & vbCrLf & _
                   "		          AND NotasFiscais.EndEmpresa_Id   = NotasFiscaisXItens.EndEmpresa_Id " & vbCrLf & _
                   "		          AND NotasFiscais.Cliente_Id      = NotasFiscaisXItens.Cliente_Id " & vbCrLf & _
                   "		          AND NotasFiscais.EndCliente_Id   = NotasFiscaisXItens.EndCliente_Id " & vbCrLf & _
                   "		          AND NotasFiscais.EntradaSaida_Id = NotasFiscaisXItens.EntradaSaida_Id " & vbCrLf & _
                   "		          AND NotasFiscais.Serie_Id        = NotasFiscaisXItens.Serie_Id " & vbCrLf & _
                   "		          AND NotasFiscais.Nota_Id         = NotasFiscaisXItens.Nota_Id " & vbCrLf & _
                   "            --***************************************************************************************************" & vbCrLf & _
                   "                 Left Join Embalagens Emb " & vbCrLf & _
                   "                   on Emb.Embalagem_id = NotasFiscaisXItens.Embalagem " & vbCrLf & _
                   "                 LEFT JOIN ProdutoXEmbalagem PxE" & vbCrLf & _
                   "                   ON NotasFiscaisXItens.Embalagem = PxE.Embalagem_Id" & vbCrLf & _
                   "                  AND NotasFiscaisXItens.Produto_Id = PxE.Produto_Id" & vbCrLf & _
                   "                  AND NotasFiscaisXItens.CapacidadeEmbalagem  = PxE.CapacidadeEmbalagem_Id " & vbCrLf & _
                   "                  AND NotasFiscaisXItens.TipoDeEmbalagem  = PxE.TipoDeEmbalagem_Id " & vbCrLf & _
                   "                 LEFT JOIN LoteXClassificacao LxC" & vbCrLf & _
                   "                   ON NotasFiscaisXItens.Classificacao  = LxC.Classificacao_Id" & vbCrLf & _
                   "                  AND NotasFiscaisXItens.Produto_Id = LxC.Produto_id" & vbCrLf & _
                   "                  AND NotasFiscaisXItens.Lote = LxC.Lote_Id" & vbCrLf & _
                   "           --*************************************************************************************************** " & vbCrLf & _
                   "		        INNER JOIN SubOperacoes " & vbCrLf & _
                   "			       ON NotasFiscais.Operacao    = SubOperacoes.Operacao_Id " & vbCrLf & _
                   "		          AND NotasFiscais.SubOperacao = SubOperacoes.SubOperacoes_Id " & vbCrLf & _
                   "	  	        Where SubOperacoes.EstoqueFisico = 'S' " & vbCrLf & _
                   "                  and SubOperacoes.Classe not in ('" & eClassesOperacoes.GLOBAL.ToString & "','" & eClassesOperacoes.COMPLEMENTACOES.ToString & "','" & eClassesOperacoes.CONTAEORDEM.ToString & "') " & vbCrLf & _
                   "		          and NotasFiscais.situacao IN (1,4,7) " & vbCrLf & _
                   "		          and NotasFiscais.Movimento     <= '" & txtDataFinal.Text.ToSqlDate() & "' " & vbCrLf

            If lstEmpresa.GetSelectedValues().Count > 0 Then
                sql &= "              and NotasFiscais.Empresa_Id + '-' + cast(NotasFiscais.EndEmpresa_Id as varchar) in ('" & String.Join("', '", lstEmpresa.GetSelectedValues().ToArray()) & "')" & vbCrLf
            End If

            sql &= "	            Group by NotasFiscaisXItens.Empresa_id," & vbCrLf & _
                   "                         NotasFiscaisXItens.EndEmpresa_id," & vbCrLf & _
                   "                         NotasFiscaisXItens.Produto_Id," & vbCrLf & _
                   "				         isnull(NotasFiscaisXItens.Lote,''), " & vbCrLf & _
                   "				         isnull(NotasFiscaisXItens.Classificacao,'')," & vbCrLf & _
                   "		   	             isnull(NotasFiscaisXItens.Embalagem,0)," & vbCrLf & _
                   "				         isnull(NotasFiscaisXItens.TipoDeEmbalagem,'')," & vbCrLf & _
                   "				         --isnull(NotasFiscaisXItens.CapacidadeEmbalagem,0) " & vbCrLf

            If chkEmbalagens.Checked Then
                sql &= "            ISNULL(NotasFiscaisXItens.CapacidadeEmbalagem, 0)  " & vbCrLf
            Else
                sql &= "            CASE " & vbCrLf & _
                       "                WHEN PxE.PesoVariavel = 1" & vbCrLf & _
                       "                     THEN ISNULL(LxC.PesoSaco,0)" & vbCrLf & _
                       "                ELSE ISNULL(NotasFiscaisXItens.CapacidadeEmbalagem, 0)" & vbCrLf & _
                       "            END " & vbCrLf
            End If

            sql &= "		        union " & vbCrLf & _
                   "		       SELECT P.Empresa_Id," & vbCrLf & _
                   "                      P.EndEmpresa_Id," & vbCrLf & _
                   "                      P.Produto_Id, " & vbCrLf & _
                   "                      isnull(P.Lote_Id,'') as Lote_Id, " & vbCrLf & _
                   "                      isnull(P.Classificacao_Id, '') as Classificacao_Id, " & vbCrLf & _
                   "                      isnull(P.Embalagem_Id,0) as Embalagem_Id, " & vbCrLf & _
                   "                      isnull(P.TipoDeEmbalagem_Id,'') as TipoDeEmbalagem_Id, " & vbCrLf
            If chkEmbalagens.Checked Then
                sql &= "            ISNULL(P.CapacidadeEmbalagem_Id, 0) AS CapacidadeEmbalagem,  " & vbCrLf
            Else
                sql &= "            CASE " & vbCrLf & _
                       "                WHEN PxE.PesoVariavel = 1" & vbCrLf & _
                       "                     THEN ISNULL(LxC.PesoSaco,0)" & vbCrLf & _
                       "                ELSE ISNULL(P.CapacidadeEmbalagem_Id, 0)" & vbCrLf & _
                       "            END as CapacidadeEmbalagem, "
            End If

            sql &= "                      sum(case " & vbCrLf & _
                   "                            when Movimento_id < '" & txtDataInicial.Text.ToSqlDate() & "' " & vbCrLf & _
                   "                              then Entradas - Saidas " & vbCrLf & _
                   "                              else 0 " & vbCrLf & _
                   "                          end) SaldoAnterior, " & vbCrLf & _
                   "                      sum(case " & vbCrLf & _
                   "                            when Movimento_id between '" & txtDataInicial.Text.ToSqlDate() & "' and '" & txtDataFinal.Text.ToSqlDate() & "' " & vbCrLf & _
                   "                              then Entradas " & vbCrLf & _
                   "                              else 0 " & vbCrLf & _
                   "                          end) EQtde, " & vbCrLf & _
                   "                      sum(case " & vbCrLf & _
                   "                            when Movimento_id between '" & txtDataInicial.Text.ToSqlDate() & "' and '" & txtDataFinal.Text.ToSqlDate() & "' " & vbCrLf & _
                   "                              then Saidas " & vbCrLf & _
                   "                              else 0 " & vbCrLf & _
                   "                          end) SQtde," & vbCrLf & _
                   "					  sum(case " & vbCrLf & _
                   "					 	    when isnull(SO.Consignacao,0) = 0 and ISNULL(so.Deposito, 'N') = 'N' and ISNULL(so.AmostraGratis, 0) = 0  and P.Movimento_id <= '" & txtDataFinal.Text.ToSqlDate() & "' and so.Classe <> '" & eClassesOperacoes.DEPOSITOS.ToString & "'" & vbCrLf & _
                   "							  then P.Entradas - P.Saidas" & vbCrLf & _
                   "							  else 0 " & vbCrLf & _
                   "						  end) SaldoProprio," & vbCrLf & _
                   "					  sum(case " & vbCrLf & _
                   "					 	    when isnull(SO.Consignacao,0) = 0 and ISNULL(so.Deposito, 'N') = 'N' and ISNULL(so.AmostraGratis, 0) = 0  and P.Movimento_id <= '" & txtDataFinal.Text.ToSqlDate() & "' and so.Classe = '" & eClassesOperacoes.DEPOSITOS.ToString & "' and so.deposito = 'N'" & vbCrLf & _
                   "							  then P.Entradas - P.Saidas" & vbCrLf & _
                   "							  else 0 " & vbCrLf & _
                   "						  end) ProdDeTec," & vbCrLf & _
                   "					  sum(case " & vbCrLf & _
                   "					 	    when isnull(SO.Consignacao,0) = 0 and ISNULL(so.Deposito, 'N') = 'S' and ISNULL(so.AmostraGratis, 0) = 0  and P.Movimento_id <= '" & txtDataFinal.Text.ToSqlDate() & "'" & vbCrLf & _
                   "							  then P.Entradas - P.Saidas" & vbCrLf & _
                   "							  else 0 " & vbCrLf & _
                   "						  end) SaldoEmPoderDeTerceiros," & vbCrLf & _
                   "					  sum(case " & vbCrLf & _
                   "							when isnull(SO.Consignacao,0) = 1 and P.Movimento_id <= '" & txtDataFinal.Text.ToSqlDate() & "'" & vbCrLf & _
                   "							  then P.Entradas - P.Saidas " & vbCrLf & _
                   "							  else 0 " & vbCrLf & _
                   "						  end) SaldoConsignado, " & vbCrLf & _
                   "                       SUM(case" & vbCrLf & _
                   "                             when ISNULL(SO.AmostraGratis, 0) = 1 and P.Movimento_Id <= '" & txtDataFinal.Text.ToSqlDate() & "'" & vbCrLf & _
                   "                               then P.Entradas - P.Saidas" & vbCrLf & _
                   "                               else 0" & vbCrLf & _
                   "                           end)SaldoAmostraGratis" & vbCrLf & _
                   "		         FROM Producao P " & vbCrLf & _
                   "				inner Join SubOperacoes SO" & vbCrLf & _
                   "				   on P.Operacao_id    = SO.Operacao_id" & vbCrLf & _
                   "				  and P.Suboperacao_id = SO.Suboperacoes_Id" & vbCrLf & _
                   "             --***************************************************************************************************" & vbCrLf & _
                   "                   Left Join Embalagens Emb " & vbCrLf & _
                   "					  on Emb.Embalagem_id = P.Embalagem_Id " & vbCrLf & _
                   "					LEFT JOIN ProdutoXEmbalagem PxE" & vbCrLf & _
                   "					  ON PxE.Embalagem_Id           = P.Embalagem_id" & vbCrLf & _
                   "					 AND PxE.Produto_Id             = P.Produto_Id " & vbCrLf & _
                   "					 AND PxE.CapacidadeEmbalagem_Id = P.CapacidadeEmbalagem_Id   " & vbCrLf & _
                   "                     AND PxE.TipoDeEmbalagem_Id  = P.TipoDeEmbalagem_Id " & vbCrLf & _
                   "					LEFT JOIN LoteXClassificacao LxC" & vbCrLf & _
                   "					  ON LxC.Classificacao_Id = P.Classificacao_id  " & vbCrLf & _
                   "					 AND LxC.Produto_id       = P.Produto_Id " & vbCrLf & _
                   "					 AND LxC.Lote_Id          = P.Lote_id " & vbCrLf & _
                   "		        --***************************************************************************************************  " & vbCrLf & _
                   "                Where P.FisicoFiscal_Id = 1 " & _
                   "                  and P.Movimento_Id    <='" & txtDataFinal.Text.ToSqlDate() & "' " & vbCrLf

            If lstEmpresa.GetSelectedValues().Count > 0 Then
                sql &= "              and P.Empresa_Id + '-' + cast(P.EndEmpresa_Id as varchar) in ('" & String.Join("', '", lstEmpresa.GetSelectedValues().ToArray()) & "')" & vbCrLf
            End If

            sql &= "		        GROUP BY P.Empresa_Id, P.EndEmpresa_Id, P.Produto_Id, isnull(P.Lote_Id,''), isnull(P.Classificacao_Id,''), isnull(P.Embalagem_Id,0), isnull(P.TipoDeEmbalagem_Id,''), " & vbCrLf

            If chkEmbalagens.Checked Then
                sql &= "            ISNULL(P.CapacidadeEmbalagem_Id, 0)  " & vbCrLf
            Else
                sql &= "            CASE " & vbCrLf & _
                       "                WHEN PxE.PesoVariavel = 1" & vbCrLf & _
                       "                     THEN ISNULL(LxC.PesoSaco,0)" & vbCrLf & _
                       "                ELSE ISNULL(P.CapacidadeEmbalagem_Id, 0)" & vbCrLf & _
                       "            END " & vbCrLf
            End If
            sql &= "               ) M" & vbCrLf & _
                   "         Group by M.Empresa_Id, M.EndEmpresa_Id, M.Produto_id, M.Lote, M.Classificacao, M.Embalagem, M.TipoDeEmbalagem, M.CapacidadeEmbalagem " & vbCrLf & _
                   "	    ) Sb " & vbCrLf

            'Tanto na quantidade de compra como de venda foi adicionado  + ( isnull(PDxNF.QtdNotaRemessa,0) - isnull(PDxNF.QtdNotaGlobal,0) 
            'Para os casos antigos q no pedido tinha venda normal e futura misturado
            sql &= "    Full Join (Select sb2.Produto_Id," & vbCrLf & _
                   "                      sum(isnull(sb2.QtdPedCompra,0)" & vbCrLf & _
                   "                          - case " & vbCrLf & _
                   "                                  when sb2.Entradasaida = 'E'" & vbCrLf & _
                   "                                    then (isnull(PDxNF.QtdNota,0) +  isnull(PDxNF.QtdNotaRemessa,0)) + ( isnull(PDxNF.QtdNotaGlobal,0) - isnull(PDxNF.QtdNotaRemessa,0))  " & vbCrLf & _
                   "                                    else 0" & vbCrLf & _
                   "                                end) as QtdCompra," & vbCrLf & _
                   "                      sum(isnull(sb2.QtdPedVenda,0)" & vbCrLf & _
                   "                          - case " & vbCrLf & _
                   "                                  when sb2.Entradasaida = 'S'" & vbCrLf & _
                   "                                   then (isnull(PDxNF.QtdNota,0) +  isnull(PDxNF.QtdNotaRemessa,0))  + ( isnull(PDxNF.QtdNotaGlobal,0) - isnull(PDxNF.QtdNotaRemessa,0)) " & vbCrLf & _
                   "                                    else 0" & vbCrLf & _
                   "                                end) as QtdVenda," & vbCrLf & _
                   "                      sum(case" & vbCrLf & _
                   "                            when isnull(PDxNF.QtdNotaGlobal,0) > 0" & vbCrLf & _
                   "                              then isnull(PDxNF.QtdNotaGlobal,0) - isnull(PDxNF.QtdNotaRemessa,0)" & vbCrLf & _
                   "                              else 0" & vbCrLf & _
                   "                          end) as FaturadaNaoEntregue" & vbCrLf & _
                   "                 from( SELECT P.Empresa_Id," & vbCrLf & _
                   "                              P.EndEmpresa_Id," & vbCrLf & _
                   "                              p.Pedido_id," & vbCrLf & _
                   "                              PxI.Produto_Id, " & vbCrLf & _
                   "                              PSO.EntradaSaida, " & vbCrLf & _
                   "                              PSO.Classe," & vbCrLf & _
                   "                              sum(case " & vbCrLf & _
                   "                                    when PSO.EntradaSaida = 'E' " & vbCrLf & _
                   "                                      then (case " & vbCrLf & _
                   "                                              when PxI.TipoDelancamento  = 'E' " & vbCrLf & _
                   "                                                then ISNULL(PxI.Quantidade,0) * - 1 " & vbCrLf & _
                   "                                                 else ISNULL(PxI.Quantidade,0) " & vbCrLf & _
                   "                                            end) " & vbCrLf & _
                   "                                  end) AS QtdPedCompra, " & vbCrLf & _
                   "                              sum(case PSO.EntradaSaida " & vbCrLf & _
                   "                                    when 'S' " & vbCrLf & _
                   "                                      then (case PxI.TipoDelancamento" & vbCrLf & _
                   "                                              when 'E' " & vbCrLf & _
                   "                                                then ISNULL(PxI.Quantidade,0) * - 1 " & vbCrLf & _
                   "                                                else ISNULL(PxI.Quantidade,0) " & vbCrLf & _
                   "                                            end) " & vbCrLf & _
                   "                                  end) AS QtdPedVenda " & vbCrLf & _
                   "                         FROM Pedidos AS P " & vbCrLf & _
                   "			            INNER JOIN PedidoXItemXLancamento PxI " & vbCrLf & _
                   "                           ON P.Empresa_Id    = PxI.Empresa_Id " & vbCrLf & _
                   "                          AND P.EndEmpresa_Id = PxI.EndEmpresa_Id " & vbCrLf & _
                   "                          AND P.Pedido_Id     = PxI.Pedido_Id " & vbCrLf & _
                   "                        Inner join Operacoes OP" & vbCrLf & _
                   "                           ON P.Operacao    = op.Operacao_Id" & vbCrLf & _
                   "                        INNER JOIN SubOperacoes as PSO " & vbCrLf & _
                   "                           ON P.Operacao    = PSO.Operacao_Id " & vbCrLf & _
                   "                          AND P.SubOperacao = PSO.SubOperacoes_Id " & vbCrLf

            If ddlSafra.SelectedIndex > 0 Then
                sql &= "                        Inner Join #Safra S" & vbCrLf & _
                       "                           On S.safra_id = P.safra " & vbCrLf
            End If

            sql &= "                        WHERE P.DataPedido   <= '" & txtDataFinal.Text.ToSqlDate() & "' " & vbCrLf & _
                   "                          AND P.Situacao      = 1 " & vbCrLf & _
                   "                          AND PSO.Classe not in ('" & eClassesOperacoes.CONTAEORDEM.ToString & "','" & eClassesOperacoes.DEPOSITOS.ToString & "')" & vbCrLf & _
                   "                          and op.Classe <> '" & eClassesOperacoes.TRANSFERENCIAS.ToString & "'" & vbCrLf & _
                   "                          and pso.Deposito = 'N'" & vbCrLf & _
                   "                        GROUP BY P.Empresa_Id, P.EndEmpresa_Id,p.Pedido_id, PxI.Produto_Id, PSO.EntradaSaida, PSO.Classe" & vbCrLf & _
                   "                       ) sb2" & vbCrLf & _
                   "                  LEFT JOIN (SELECT NF.Pedido, " & vbCrLf & _
                   "                                    NFxI.Produto_Id, " & vbCrLf & _
                   "                                    sum(case " & vbCrLf & _
                   "                                          when SO.Devolucao = 'N' and SO.Classe = '" & eClassesOperacoes.GLOBAL.ToString & "'" & vbCrLf & _
                   "                                            then NFxI.QuantidadeFiscal " & vbCrLf & _
                   "                                          When SO.Devolucao = 'S' and SO.Classe = '" & eClassesOperacoes.GLOBAL.ToString & "' " & vbCrLf & _
                   "                                            then NFxI.QuantidadeFiscal * - 1 " & vbCrLf & _
                   "                                          else 0 " & vbCrLf & _
                   "                                        end) AS QtdNotaGlobal, " & vbCrLf & _
                   "                                    sum(case " & vbCrLf & _
                   "                                          when SO.Devolucao = 'N' and SO.Classe = '" & eClassesOperacoes.REMESSAS.ToString & "' " & vbCrLf & _
                   "                                            then NFxI.QuantidadeFiscal " & vbCrLf & _
                   "                                          when SO.Devolucao = 'S' and SO.Classe = '" & eClassesOperacoes.REMESSAS.ToString & "' " & vbCrLf & _
                   "                                            then NFxI.QuantidadeFiscal * - 1 " & vbCrLf & _
                   "                                          else 0 " & vbCrLf & _
                   "                                        end) AS QtdNotaRemessa, " & vbCrLf & _
                   "                                    sum(case " & vbCrLf & _
                   "                                          when SO.Devolucao = 'N' and SO.Classe <> '" & eClassesOperacoes.REMESSAS.ToString & "' and SO.Classe <> '" & eClassesOperacoes.GLOBAL.ToString & "'" & vbCrLf & _
                   "                                            then NFxI.QuantidadeFiscal " & vbCrLf & _
                   "                                          when SO.Devolucao = 'S' and SO.Classe <> '" & eClassesOperacoes.REMESSAS.ToString & "' and SO.Classe <> '" & eClassesOperacoes.GLOBAL.ToString & "'" & vbCrLf & _
                   "                                            then NFxI.QuantidadeFiscal * - 1 " & vbCrLf & _
                   "                                          else 0 " & vbCrLf & _
                   "                                        end) AS QtdNota " & vbCrLf & _
                   "                               FROM NotasFiscais AS NF" & vbCrLf & _
                   "              	              INNER JOIN NotasFiscaisXItens NFxI" & vbCrLf & _
                   "                                 ON NF.Empresa_Id      = NFxI.Empresa_Id" & vbCrLf & _
                   "                                AND NF.EndEmpresa_Id   = NFxI.EndEmpresa_Id" & vbCrLf & _
                   "                                AND NF.Cliente_Id      = NFxI.Cliente_Id" & vbCrLf & _
                   "                                AND NF.EndCliente_Id   = NFxI.EndCliente_Id" & vbCrLf & _
                   "                                AND NF.Serie_Id        = NFxI.Serie_Id" & vbCrLf & _
                   "                                AND NF.EntradaSaida_Id = NFxI.EntradaSaida_Id" & vbCrLf & _
                   "                                AND NF.Nota_Id         = NFxI.Nota_Id" & vbCrLf & _
                   "                              INNER JOIN SubOperacoes SO" & vbCrLf & _
                   "                                 ON NF.Operacao    = SO.Operacao_Id" & vbCrLf & _
                   "                                AND NF.SubOperacao = SO.SubOperacoes_Id" & vbCrLf
            If ddlSafra.SelectedIndex > 0 Then
                sql &= "                              Inner Join Pedidos P" & vbCrLf & _
                       "                                 ON P.Empresa_Id    = NF.Empresa_Id" & vbCrLf & _
                       "                                And P.EndEmpresa_Id = NF.EndEmpresa_Id" & vbCrLf & _
                       "                                And P.Pedido_id     = NF.Pedido" & vbCrLf & _
                       "                              Inner Join #Safra S" & vbCrLf & _
                       "                                 On S.safra_id = P.safra " & vbCrLf
            End If

            sql &= "		                      WHERE NF.Situacao    in (1,4,7)" & vbCrLf & _
                   "                                AND NF.DataDaNota <= '" & txtDataFinal.Text.ToSqlDate() & "' " & vbCrLf & _
                   "                                AND SO.Classe not in ('" & eClassesOperacoes.COMPLEMENTACOES.ToString & "','" & eClassesOperacoes.CONTAEORDEM.ToString & "','" & eClassesOperacoes.DEPOSITOS.ToString & "') " & vbCrLf & _
                   "                                and so.Deposito = 'N'" & vbCrLf & _
                   "                              GROUP BY NF.Empresa_Id, NF.EndEmpresa_Id, NF.Pedido, NFxI.Produto_Id" & vbCrLf & _
                   "                             ) AS PDxNF" & vbCrLf & _
                   "                          ON sb2.Pedido_Id     = PDxNF.Pedido" & vbCrLf & _
                   "                         AND sb2.Produto_Id    = PDxNF.Produto_Id" & vbCrLf & _
                   "                       Where 1 = 1" & vbCrLf & _
                   "              and sb2.Empresa_Id + '-' + cast(sb2.EndEmpresa_Id as varchar) in ('" & String.Join("', '", lstEmpresa.GetSelectedValues().ToArray()) & "')" & vbCrLf & _
                   "                     group by sb2.Produto_Id" & vbCrLf & _
                   "                     ) AS PosicaoPedido " & vbCrLf & _
                   "      ON PosicaoPedido.Produto_id    = sb.Produto_Id" & vbCrLf & _
                   "   Inner Join Produtos P " & vbCrLf & _
                   "      on P.Produto_Id     = isnull(PosicaoPedido.Produto_Id,Sb.Produto_Id)" & vbCrLf & _
                   "     and (P.ControlarLote = 'S' or P.ControlarEmbalagem = 'S' or P.ControlarEstoque = 'S') " & vbCrLf & _
                   "   Inner Join GruposDeEstoques GE " & vbCrLf & _
                   "      on left(P.Grupo,2) = GE.Grupo_id" & vbCrLf & _
                   "    Left Join Embalagens Emb " & vbCrLf & _
                   "      on Emb.Embalagem_id = Sb.Embalagem " & vbCrLf & _
                   "   Where 1 = 1 " & vbCrLf

            If ddlMarca.SelectedIndex > 0 Then
                sql &= " AND P.Marca = " & ddlMarca.SelectedValue & vbCrLf
                Parametros &= "Marca: " & ddlMarca.SelectedItem.Text & vbCrLf
            End If

            If ucSelecaoProduto.TemSelecionado Then
                Dim ret As New ArrayList
                ret = ucSelecaoProduto.GetSqlEParametrosRelatorio("P.Grupo", "P.Produto_Id", "", True)
                sql &= " and " & ret(0) & vbCrLf
                Parametros &= "Produtos: " & ret(1) & vbCrLf
            End If

            sql &= "   Group by Ge.Descricao, P.Produto_Id, P.Nome, P.Unidade " & vbCrLf


            If Not chkLoteSemente.Checked Then
                sql &= "       ,isnull(sb.Lote,''), " & vbCrLf & _
                       "       isnull(sb.Classificacao,'') " & vbCrLf
            End If

            If Not chkEmbalagens.Checked Then
                sql &= "       ,isnull(sb.Embalagem,0), " & vbCrLf & _
                       "       isnull(sb.TipoDeEmbalagem,''),  " & vbCrLf & _
                       "       ISNULL(sb.CapacidadeEmbalagem, 0)  "
            End If

            If chkSaldoQtde.Checked = True Then
                sql &= " having (sum(isnull(sb.SaldoConsignado,0)) <> 0 or sum(isnull(sb.SaldoProprio,0)) <> 0 OR SUM(ISNULL(sb.SaldoAmostraGratis,0)) <> 0 or SUM(ISNULL(sb.SaldoEmPoderDeTerceiros,0)) <> 0 or sum(sb.ProdDeTerc) <> 0 or abs(sum(sb.SaldoEmPoderDeTerceiros)) <> 0)"
                If chkVendaCompraSaldo.Checked Then
                    sql &= " or ( max(isnull(PosicaoPedido.QtdCompra,0)) > 0 or max(isnull(PosicaoPedido.QtdVenda,0)) > 0 )"
                    Parametros &= "Apenas Com Saldo Quantidade ou Pedidos de Compra/Venda em Aberto "
                Else
                    Parametros &= "Apenas Com Saldo Quantidade "
                End If
            ElseIf chkVendaCompraSaldo.Checked Then
                sql &= " having max(isnull(PosicaoPedido.QtdCompra,0)) > 0 or max(isnull(PosicaoPedido.QtdVenda,0)) > 0"
                Parametros &= "Com Pedidos de Compra/Venda em Aberto "
            End If

            sql &= " Order By P.Nome "

            If RdEstoqueComPreco.Checked Or rbPreco.Checked Then
                If RdEstoqueComPreco.Checked Then
                    sql &= "select sb.Cifrao, t.*, isnull(sb.MargemMenor,0) as MargemMenor, isnull(sb.MargemMaior,0) as MargemMaior" & vbCrLf & _
                           "  from #Temp t" & vbCrLf & _
                           " left Join (" & vbCrLf & _
                           "			Select PxP.Produto_Id," & vbCrLf & _
                           "				   M.Cifrao," & vbCrLf & _
                           "				   (PxP.Valor * (1 + PxP.FixoOperacional /100)) * (1 + isnull(MargemMenor,0) / 100)  as MargemMenor," & vbCrLf & _
                           "				   (PxP.Valor * (1 + PxP.FixoOperacional /100)) * (1 + isnull(MargemMaior,0) / 100)  as MargemMaior" & vbCrLf & _
                           "			  from ProdutosXPrecos PxP" & vbCrLf & _
                           "			 Inner Join (" & vbCrLf & _
                           "						  select Produto_Id, Max(Data_Id) as Data_Id" & vbCrLf & _
                           "							from produtosxprecos" & vbCrLf & _
                           "						   Where Data_Id <= '" & txtDataFinal.Text.ToSqlDate() & "'" & vbCrLf & _
                           "						   group by Produto_Id" & vbCrLf & _
                           "						 ) sb" & vbCrLf & _
                           "			   on PxP.Produto_id = sb.Produto_Id" & vbCrLf & _
                           "			  and PxP.Data_Id    = sb.Data_Id" & vbCrLf & _
                           "			Inner Join Moedas M" & vbCrLf & _
                           "			   on M.Moeda_Id = PxP.Moeda" & vbCrLf & _
                           "            ) sb" & vbCrLf & _
                           "   on t.CodigoProduto = sb.Produto_Id" & vbCrLf
                    ds = Banco.ConsultaDataSet(sql, "RelatorioEstoqueRevenda")
                End If

                If rbPreco.Checked Then
                    sql &= " update #temp set " & vbCrLf & _
                           "    NomeMoeda      = Sb.NomeMoeda " & vbCrLf & _
                           "   ,cifrao         = sb.Cifrao " & vbCrLf & _
                           "   ,ValorProduto   = sb.valor " & vbCrLf & _
                           "   ,VTotal         = (sb.valor * saldoProprio)" & vbCrLf & _
                           "   ,VTotalPrevisao = (sb.valor * previsao)" & vbCrLf & _
                           " from #temp" & vbCrLf & _
                           " left Join (" & vbCrLf & _
                           " 			Select PxP.Produto_Id," & vbCrLf & _
                           "                   M.Descricao as NomeMoeda," & vbCrLf & _
                           " 				   M.Cifrao," & vbCrLf & _
                           " 				   (PxP.Valor * (1 + PxP.FixoOperacional /100)) as Valor" & vbCrLf & _
                           " 			  from ProdutosXPrecos PxP" & vbCrLf & _
                           " 			 Inner Join (" & vbCrLf & _
                           " 						  select Produto_Id, Max(Data_Id) as Data_Id" & vbCrLf & _
                           " 							from produtosxprecos" & vbCrLf & _
                           " 						   Where Data_Id <= '" & txtDataFinal.Text.ToSqlDate() & "'" & vbCrLf & _
                           " 						   group by Produto_Id" & vbCrLf & _
                           " 						 ) sb" & vbCrLf & _
                           " 			   on PxP.Produto_id = sb.Produto_Id" & vbCrLf & _
                           " 			  and PxP.Data_Id    = sb.Data_Id" & vbCrLf & _
                           " 			Inner Join Moedas M" & vbCrLf & _
                           " 			   on M.Moeda_Id = PxP.Moeda" & vbCrLf & _
                           "            Inner Join Produtos Prd" & vbCrLf & _
                           "               on PxP.Produto_Id = prd.Produto_Id" & vbCrLf & _
                           "            ) sb" & vbCrLf & _
                           "    on #temp.CodigoProduto = sb.Produto_Id" & vbCrLf

                    sql &= " Select * from #Temp" & vbCrLf


                    sql &= " select NomeMoeda," & vbCrLf & _
                           "        Cifrao," & vbCrLf & _
                           "        sum(VTotal) as ValorSaldo," & vbCrLf & _
                           "        sum(VTotalPrevisao) as ValorPrevisao" & vbCrLf & _
                           "   from #temp" & vbCrLf & _
                           "  Where NomeMoeda is not null" & vbCrLf & _
                           "  group by NomeMoeda, Cifrao" & vbCrLf

                    ds = Banco.ConsultaDataSet(sql, "Consulta")
                    ds.Tables(0).TableName = "RelatorioEstoqueRevenda"
                    ds.Tables(1).TableName = "ResumoMoeda"

                End If
            Else
                ds = Banco.ConsultaDataSet(sql, "RelatorioEstoqueRevenda")
            End If


            Dim parameters = New Dictionary(Of String, Object)
            parameters.Add("Parametros", Parametros)

            Dim Titulo As String = String.Empty
            If RdEstoqueComPreco.Checked Then
                Titulo = "Estoque e Tabela de Preço Dia " & txtDataFinal.Text
            Else
                Titulo = "" & IIf(rbEstoque.Checked = True, "Estoque", "Preço") & " atual de " & txtDataInicial.Text & " à " & txtDataFinal.Text
            End If
            parameters.Add("Titulo", Titulo)

            If chkMapa.Checked Then
                gerarExcel(ds, "Mapa Comparativo")
            Else
                Dim NomeDoArquivo As String = String.Empty

                If RdEstoqueComPreco.Checked Then
                    NomeDoArquivo = "Cr_RelatorioEstoquePreco"
                ElseIf rbPreco.Checked Then
                    NomeDoArquivo = "Cr_RelatorioPrecoRevenda"
                ElseIf rbEstoque.Checked Then
                    NomeDoArquivo = "Cr_RelatorioEstoqueRevenda"
                    parameters.Add("EsconderSafra", Not ckPorSafra.Checked)
                    ds.Merge(ListaPorSafra(ckPorSafra.Checked))
                End If

                Funcoes.BindReport(Me.Page, ds, NomeDoArquivo, IIf(Pdf, eExportType.PDF, eExportType.ExcelCrystal), parameters)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub gerarExcel(ByVal ds As DataSet, ByVal TituloAba As String)
        If ds Is Nothing OrElse ds.Tables(0) Is Nothing OrElse ds.Tables(0).Rows.Count = 0 Then
            MsgBox(Me.Page, "Nenhum resultado encontrado!")
        Else
            Try
                Dim rowIndex As Integer = 2
                Dim fileName As String = HttpContext.Current.Server.MapPath("~/PlanilhasExcel/" & Funcoes.GeraNomeArquivo & ".xlsx")
                Dim qtdeColunas As Integer = countColunasMapaCorporativo()

                If File.Exists(fileName) Then File.Delete(fileName)

                Using arquivo As New FileStream(fileName, FileMode.CreateNew)
                    Using package As New ExcelPackage(arquivo)
                        'criando aba da planilha
                        Dim worksheet As ExcelWorksheet = package.Workbook.Worksheets.Add(TituloAba)
                        Dim columnIndex As Integer = 1

                        Dim emp As String = String.Empty
                        For Each col As DataColumn In ds.Tables(0).Columns
                            If columnIndex > 9 - IIf(chkLoteSemente.Checked, 2, 0) - IIf(chkEmbalagens.Checked, 3, 0) Then
                                If emp <> col.ColumnName.Split("-")(1) Then
                                    worksheet.Cells(rowIndex - 1, columnIndex, rowIndex - 1, columnIndex - 1 + qtdeColunas).Merge = True
                                    worksheet.Cells(rowIndex - 1, columnIndex).Value = col.ColumnName.Split("-")(1)
                                    worksheet.Cells(rowIndex - 1, columnIndex).Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Center
                                    worksheet.Cells(rowIndex - 1, columnIndex).Style.Border.Left.Style = Style.ExcelBorderStyle.Thin
                                    worksheet.Cells(rowIndex, columnIndex).Style.Border.Left.Style = Style.ExcelBorderStyle.Thin
                                    worksheet.Cells(rowIndex - 1, columnIndex).Style.Font.Bold = True
                                    emp = col.ColumnName.Split("-")(1)
                                End If
                                worksheet.Cells(rowIndex, columnIndex).Value = col.ColumnName.Split("-")(0)
                            Else
                                worksheet.Cells(rowIndex, columnIndex).Value = col.ColumnName
                            End If

                            worksheet.Cells(rowIndex, columnIndex).Style.Font.Bold = True
                            columnIndex += 1
                        Next

                        rowIndex += 1
                        emp = String.Empty
                        For Each row As DataRow In ds.Tables(0).Rows
                            columnIndex = 1
                            For Each col As DataColumn In ds.Tables(0).Columns
                                If IsNumeric(row(col.ColumnName)) AndAlso Not row(col.ColumnName).ToString.Trim.Contains(",") Then
                                    worksheet.Cells(rowIndex, columnIndex).Value = Convert.ToInt64(row(col.ColumnName).ToString().Replace(".", ""))
                                    worksheet.Cells(rowIndex, columnIndex).Style.Numberformat.Format = "0"
                                Else
                                    worksheet.Cells(rowIndex, columnIndex).Value = row(col.ColumnName)
                                End If

                                'Formatando campos com valores numerico, decimais e datas.
                                If IsNumeric(row(col.ColumnName)) AndAlso row(col.ColumnName).ToString.Contains(",") Then
                                    worksheet.Cells(rowIndex, columnIndex).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"
                                ElseIf IsDate(row(col.ColumnName)) Then
                                    worksheet.Cells(rowIndex, columnIndex).Style.Numberformat.Format = "dd/MM/yyyy"
                                End If

                                If columnIndex > 9 - IIf(chkLoteSemente.Checked, 2, 0) - IIf(chkEmbalagens.Checked, 3, 0) AndAlso emp <> col.ColumnName.Split("-")(1) Then
                                    worksheet.Cells(rowIndex, columnIndex).Style.Border.Left.Style = Style.ExcelBorderStyle.Thin
                                    emp = col.ColumnName.Split("-")(1)
                                End If

                                columnIndex += 1
                            Next
                            rowIndex += 1
                        Next

                        worksheet.Cells(1, columnIndex, rowIndex, columnIndex).Style.Border.Left.Style = Style.ExcelBorderStyle.Thin

                        'setando autofit nas células da planilha
                        worksheet.Cells.AutoFitColumns(0)

                        'congelando segunda linha
                        worksheet.View.FreezePanes(3, 1)

                        'salvando planilha do excel
                        package.Save()
                    End Using
                End Using

                'download do arquivo pelo browser
                Funcoes.AbrirExcel(Me.Page, "PlanilhasExcel/" & Path.GetFileName(fileName))
            Catch ex As Exception
                MsgBox(Me.Page, ex.Message)
            End Try
        End If
    End Sub

    Private Function countColunasMapaCorporativo() As Integer
        Dim count As Integer = 0
        If chkSaldoAnterior.Checked Then count += 1
        If ChkEntradas.Checked Then count += 1
        If ChkSaidas.Checked Then count += 1
        If ChkConsignado.Checked Then count += 1
        If ChkAmostraGratis.Checked Then count += 1
        If ChkTerceiros.Checked Then count += 1
        If ChkSaldoProprio.Checked Then count += 1
        If ChkSaldoFora.Checked Then count += 1
        If ChkFatNaoEntregue.Checked Then count += 1
        If ChkCompra.Checked Then count += 1
        If ChkVenda.Checked Then count += 1
        If ChkPrevisao.Checked Then count += 1
        Return count
    End Function

    Public Function ListaPorSafra(Optional ByVal Carregar As Boolean = True) As DataSet
        Dim ds As DataSet
        Dim sql As String

        Dim Safra As [Lib].Negocio.Safra = Nothing
        If ddlSafra.SelectedIndex > 0 Then
            Safra = New [Lib].Negocio.Safra(ddlSafra.SelectedValue)
        End If

        If Not Carregar Then
            sql = "Select '' as Empresa_id," & vbCrLf & _
                  "       0  as EndEmpresa_Id," & vbCrLf & _
                  "       '' as Produto_id," & vbCrLf & _
                  "       '' as Safra," & vbCrLf & _
                  "       0.0 as QtdCompra," & vbCrLf & _
                  "       0.0 as QtdVenda " & vbCrLf
        Else
            sql = " Select sb2.Empresa_Id, " & vbCrLf & _
                  "        sb2.EndEmpresa_Id," & vbCrLf & _
                  "        sb2.Produto_Id," & vbCrLf & _
                  "        sb2.Safra," & vbCrLf & _
                  "        sum(isnull(sb2.QtdCompra,0) - case when sb2.Entradasaida = 'E' then isnull(PDxNF.QtdNota,0) else 0 end) as QtdCompra, " & vbCrLf & _
                  "        sum(isnull(sb2.QtdVenda,0)  - case when sb2.Entradasaida = 'S' then isnull(PDxNF.QtdNota,0) else 0 end) as QtdVenda" & vbCrLf & _
                  "   from(" & vbCrLf & _
                  "        SELECT P.Empresa_Id," & vbCrLf & _
                  "               P.EndEmpresa_Id," & vbCrLf & _
                  "               p.Pedido_id," & vbCrLf & _
                  "               PxI.Produto_Id, " & vbCrLf & _
                  "               P.Safra, " & vbCrLf & _
                  "               PSO.EntradaSaida, " & vbCrLf & _
                  "               PR.Grupo, " & vbCrLf & _
                  "               sum(case " & vbCrLf & _
                  "                     when PSO.EntradaSaida = 'E' " & vbCrLf & _
                  "                       then (case " & vbCrLf & _
                  "                               when PxI.TipoDelancamento  = 'E' " & vbCrLf & _
                  "                                then ISNULL(PxI.Quantidade,0) * - 1 " & vbCrLf & _
                  "                                else ISNULL(PxI.Quantidade,0) " & vbCrLf & _
                  "                             end) " & vbCrLf & _
                  "                   end) AS QtdCompra, " & vbCrLf & _
                  "                sum(case PSO.EntradaSaida " & vbCrLf & _
                  "                      when 'S' " & vbCrLf & _
                  "                        then (case PxI.TipoDelancamento" & vbCrLf & _
                  "                                when 'E' " & vbCrLf & _
                  "                                  then ISNULL(PxI.Quantidade,0) * - 1 " & vbCrLf & _
                  "                                  else ISNULL(PxI.Quantidade,0) " & vbCrLf & _
                  "                              end) " & vbCrLf & _
                  "                    end) AS QtdVenda " & vbCrLf & _
                  "          FROM Pedidos AS P " & vbCrLf & _
                  "	        INNER JOIN PedidoXItemXLancamento PxI " & vbCrLf & _
                  "            ON P.Empresa_Id    = PxI.Empresa_Id " & vbCrLf & _
                  "           AND P.EndEmpresa_Id = PxI.EndEmpresa_Id " & vbCrLf & _
                  "           AND P.Pedido_Id     = PxI.Pedido_Id " & vbCrLf & _
                  "         INNER JOIN SubOperacoes as PSO " & vbCrLf & _
                  "            ON P.Operacao    = PSO.Operacao_Id " & vbCrLf & _
                  "           AND P.SubOperacao = PSO.SubOperacoes_Id " & vbCrLf & _
                  "         INNER JOIN Produtos PR " & vbCrLf & _
                  "            ON PxI.Produto_Id  = Pr.Produto_Id " & vbCrLf & _
                  "         WHERE P.DataPedido   <= '" & txtDataFinal.Text.ToSqlDate() & "' " & vbCrLf & _
                  "           AND P.Situacao      = 1 " & vbCrLf & _
                  "           AND PSO.Classe <> '" & eClassesOperacoes.CONTAEORDEM.ToString & "' " & vbCrLf & _
                  "         GROUP BY P.Empresa_Id, P.EndEmpresa_Id,p.Pedido_id, PxI.Produto_Id, P.Safra, PSO.EntradaSaida, Pr.Grupo" & vbCrLf & _
                  "        ) sb2" & vbCrLf & _
                  "   LEFT JOIN (SELECT NF.Empresa_Id, " & vbCrLf & _
                  "                     NF.EndEmpresa_Id, " & vbCrLf & _
                  "                     NF.Pedido, " & vbCrLf & _
                  "                     NFxI.Produto_Id, " & vbCrLf & _
                  "                     sum(case " & vbCrLf & _
                  "                           when SO.Devolucao = 'N' " & vbCrLf & _
                  "                             then NFxI.QuantidadeFiscal " & vbCrLf & _
                  "                              else NFxI.QuantidadeFiscal * - 1 " & vbCrLf & _
                  "                         end) AS QtdNota" & vbCrLf & _
                  "                FROM NotasFiscais AS NF" & vbCrLf & _
                  "               INNER JOIN NotasFiscaisXItens NFxI" & vbCrLf & _
                  "                  ON NF.Empresa_Id      = NFxI.Empresa_Id" & vbCrLf & _
                  "                 AND NF.EndEmpresa_Id   = NFxI.EndEmpresa_Id" & vbCrLf & _
                  "                 AND NF.Cliente_Id      = NFxI.Cliente_Id" & vbCrLf & _
                  "                 AND NF.EndCliente_Id   = NFxI.EndCliente_Id" & vbCrLf & _
                  "                 AND NF.Serie_Id        = NFxI.Serie_Id" & vbCrLf & _
                  "                 AND NF.EntradaSaida_Id = NFxI.EntradaSaida_Id" & vbCrLf & _
                  "                 AND NF.Nota_Id         = NFxI.Nota_Id" & vbCrLf & _
                  "               INNER JOIN SubOperacoes SO" & vbCrLf & _
                  "                  ON NF.Operacao    = SO.Operacao_Id" & vbCrLf & _
                  "                 AND NF.SubOperacao = SO.SubOperacoes_Id" & vbCrLf & _
                  "		          WHERE NF.Situacao    in (1,4,7)" & vbCrLf & _
                  "                 and NF.DataDaNota  <= '" & txtDataFinal.Text.ToSqlDate() & "' " & vbCrLf & _
                  "                 AND SO.Classe not in ('" & eClassesOperacoes.GLOBAL.ToString & "','" & eClassesOperacoes.COMPLEMENTACOES.ToString & "','" & eClassesOperacoes.CONTAEORDEM.ToString & "') " & vbCrLf

            If lstEmpresa.GetSelectedValues().Count > 0 Then
                sql &= "              and NF.Empresa_Id + '-' + cast(Nf.EndEmpresa_Id as varchar) in ('" & String.Join("', '", lstEmpresa.GetSelectedValues().ToArray()) & "')" & vbCrLf
            End If

            sql &= "               GROUP BY NF.Empresa_Id, NF.EndEmpresa_Id, NF.Pedido, NFxI.Produto_Id" & vbCrLf & _
                  "               ) AS PDxNF" & vbCrLf & _
                  "           ON sb2.Empresa_Id    = PDxNF.Empresa_Id" & vbCrLf & _
                  "          AND sb2.EndEmpresa_Id = PDxNF.EndEmpresa_Id" & vbCrLf & _
                  "          AND sb2.Pedido_Id     = PDxNF.Pedido" & vbCrLf & _
                  "          AND sb2.Produto_Id    = PDxNF.Produto_Id" & vbCrLf

            If ddlSafra.SelectedIndex > 0 Then
                sql &= "  Inner Join #Safra S" & vbCrLf & _
                       "     On S.safra_id = sb2.Safra " & vbCrLf & _
                       "    --and S.Vencimento <= '" & Safra.Vencimento.ToString("yyyy-MM-dd") & "'" & vbCrLf
            End If

            sql &= "              Where sb2.Empresa_Id + '-' + cast(sb2.EndEmpresa_Id as varchar) in ('" & String.Join("', '", lstEmpresa.GetSelectedValues().ToArray()) & "')" & vbCrLf

            If ucSelecaoProduto.TemSelecionado Then
                Dim ret As New ArrayList
                ret = ucSelecaoProduto.GetSqlEParametrosRelatorio("sb2.Grupo", "sb2.Produto_Id", "", True)
                sql &= " and " & ret(0) & vbCrLf
            End If

            sql &= "  group by sb2.Empresa_Id, sb2.EndEmpresa_Id, sb2.Produto_Id, sb2.Safra" & vbCrLf & _
                   " Having sum(isnull(sb2.QtdCompra,0) - case when sb2.Entradasaida = 'E' then isnull(PDxNF.QtdNota,0) else 0 end) <> 0 " & vbCrLf & _
                   "        OR" & vbCrLf & _
                   "        sum(isnull(sb2.QtdVenda,0)  - case when sb2.Entradasaida = 'S' then isnull(PDxNF.QtdNota,0) else 0 end) <> 0" & vbCrLf
        End If

        ds = Banco.ConsultaDataSet(sql, "ListaPorSafra")

        Return ds
    End Function

    Protected Sub ddlMarca_SelectedIndexChanged(ByVal sender As Object, ByVal e As EventArgs) Handles ddlMarca.SelectedIndexChanged
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
            Limpar()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Protected Sub ddlSafra_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddlSafra.SelectedIndexChanged
        Try
            If ddlSafra.SelectedIndex > 0 Then
                ckPorSafra.Checked = False
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub ckPorSafra_CheckedChanged(sender As Object, e As EventArgs) Handles ckPorSafra.CheckedChanged
        Try
            If ckPorSafra.Checked Then ddlSafra.SelectedIndex = 0
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "RelatorioEstoqueRevenda")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub chkMapa_CheckedChanged(sender As Object, e As EventArgs) Handles chkMapa.CheckedChanged
        If CType(sender, CheckBox).Checked Then
            pnlColunas.Style.Clear()
        Else
            pnlColunas.Style.Add("display", "none")
        End If
    End Sub
End Class