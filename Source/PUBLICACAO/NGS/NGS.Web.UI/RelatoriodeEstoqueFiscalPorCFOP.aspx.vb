Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports System.IO
Imports System.Data
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class RelatoriodeEstoqueFiscalPorCFOP
    Inherits BasePage

#Region "Eventos"

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            If Not IsPostBack And IsConnect Then
                Me.setMenu(eModulo.Comercial)
                If Funcoes.VerificaPermissao("RelatoriodeEstoqueFiscalPorCFOP", "ACESSAR") Then
                    ddl.Carregar(ddlEmpresa, CarregarDDL.Tabela.Empresas, "")
                    Funcoes.VerificaEmpresa(ddlEmpresa)
                    ddl.Carregar(ddlEmpresaSI, CarregarDDL.Tabela.Empresas, "")
                    Funcoes.VerificaEmpresa(ddlEmpresaSI)
                    ddl.Carregar(ddlAno, CarregarDDL.Tabela.Ano, "2016;10;C", False)
                    ddl.Carregar(ddlAnoSI, CarregarDDL.Tabela.Ano, "2016;10;C", False)
                    ddl.Carregar(ddlMesInicial, CarregarDDL.Tabela.Mes, "", False)
                    txtDataFinal.Text = Date.Now.AddDays(-1).ToString("dd/MM/yyyy")
                    ucConsultaProduto.SetarHID(HID.Value)
                    LiberaEmpresa()
                Else
                    MsgBox(Me.Page, "Usuário sem permissão para acessar essa página", "~/Comercial.aspx")
                    Exit Sub
                End If
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

    Private Sub EmitirRelatorio(ByVal Pdf As Boolean)
        Try
            If Funcoes.VerificaPermissao("RelatoriodeEstoqueFiscalPorCFOP", "RELATORIO") Then
                If Not valida() Then Exit Sub
                Dim Parametros As String = "PARAMETROS" & vbCrLf

                Dim selprd As ArrayList
                selprd = ucSelecaoProduto.GetSqlEParametrosRelatorio("P.Grupo", "P.Produto_Id", "", True)

                Dim Emp As String() = ddlEmpresa.SelectedValue.Split("-")
                Dim Empresa As New Cliente(Emp(0), Emp(1))
                Parametros &= "Empresa: " & Empresa.Codigo & " - " & Empresa.Nome & " / " & Empresa.Cidade & "-" & Empresa.CodigoEstado & vbCrLf
                If chkConsolidarEmpresa.Checked Then Parametros &= "Consolidar Empresa" & vbCrLf
                Parametros &= "Periodo: Ano " & ddlAno.SelectedValue & " Mes de " & ddlMesInicial.SelectedItem.Text & " a " & txtDataFinal.Text & vbCrLf
                Parametros &= "Produto(s): " & selprd(1)

                Dim sql As String
                sql = "Select TAG, Ano, Mes, EntradaSaida, CFOP, EmNossoPoder, EmPoderDeTerceiros, ParaFinsDeExportacao, TotalEstoqueProprio, EstoqueDeTerceiro, TotalEstoqueGeral" & vbCrLf & _
                      "  from (" & vbCrLf & _
                      "		select '1 - ESTOQUE INICIAL' as TAG," & vbCrLf & _
                      "		       0  as Ano," & vbCrLf & _
                      "			   -1 as Mes," & vbCrLf & _
                      "			   '' as EntradaSaida," & vbCrLf & _
                      "			   0  as CFOP," & vbCrLf & _
                      "			   isnull(sum(case" & vbCrLf & _
                      "					        when E.Tipo_id = 'NP'" & vbCrLf & _
                      "					          then E.Quantidade" & vbCrLf & _
                      "					          else 0" & vbCrLf & _
                      "				          end),0) as EmNossoPoder," & vbCrLf & _
                      "			   isnull(sum(case" & vbCrLf & _
                      "					        when E.Tipo_id = 'PT'" & vbCrLf & _
                      "					          then E.Quantidade" & vbCrLf & _
                      "					          else 0" & vbCrLf & _
                      "				          end),0) as EmPoderDeTerceiros," & vbCrLf & _
                      "			   isnull(sum(case" & vbCrLf & _
                      "					        when E.Tipo_id = 'FE'" & vbCrLf & _
                      "					          then E.Quantidade" & vbCrLf & _
                      "					          else 0" & vbCrLf & _
                      "				          end),0) as ParaFinsDeExportacao," & vbCrLf & _
                      "			   isnull(sum(case" & vbCrLf & _
                      "					        when E.Tipo_id = 'EP'" & vbCrLf & _
                      "					          then E.Quantidade" & vbCrLf & _
                      "					          else 0" & vbCrLf & _
                      "				          end),0) as TotalEstoqueProprio," & vbCrLf & _
                      "			   isnull(sum(case" & vbCrLf & _
                      "					        when E.Tipo_id = 'ET'" & vbCrLf & _
                      "					          then E.Quantidade" & vbCrLf & _
                      "					          else 0" & vbCrLf & _
                      "				          end),0) as EstoqueDeTerceiro," & vbCrLf & _
                      "            isnull(sum(case" & vbCrLf & _
                      "					        when E.Tipo_id = 'EG'" & vbCrLf & _
                      "					         then E.Quantidade" & vbCrLf & _
                      "					         else 0" & vbCrLf & _
                      "				          end),0) as TotalEstoqueGeral" & vbCrLf & _
                      "		  from EstoqueInicialCFOP E" & vbCrLf & _
                      "      inner join Produtos P" & vbCrLf & _
                      "         on E.Produto_id = P.Produto_Id" & vbCrLf & _
                      "		 where year(E.data_Id) = " & ddlAno.SelectedValue & vbCrLf & _
                      "        AND " & selprd(0) & vbCrLf

                If chkConsolidarEmpresa.Checked Then
                    sql &= "		  and left(Empresa_id,8) = '" & Emp(0).Substring(0, 8) & "'" & vbCrLf
                Else
                    sql &= "		  And Empresa_id    ='" & Emp(0) & "'" & vbCrLf & _
                           "          And EndEmpresa_id = " & Emp(1) & vbCrLf
                End If

                sql &= "		 union all" & vbCrLf & _
                       "		SELECT '2 - ESTOQUE ANTERIOR' as TAG," & vbCrLf & _
                       "		       0  as Ano," & vbCrLf & _
                       "			   0  as MES," & vbCrLf & _
                       "			   '' as EntradaSaida," & vbCrLf & _
                       "			   0  as CFOP_Id," & vbCrLf & _
                       "			   isnull(sum(case" & vbCrLf & _
                       "					        when SO.ProdutoDeTerceiro = 0 and SO.Deposito = 'N' and isnull(SO.ParaFinsDeExportacao,0) = 0" & vbCrLf & _
                       "							  then case when nf.entradaSaida_id = 'E' then Nfi.QuantidadeFiscal else Nfi.QuantidadeFiscal * -1 end" & vbCrLf & _
                       "							  else 0" & vbCrLf & _
                       "						   end),0) as EmNossoPoder," & vbCrLf & _
                       "			   isnull(sum(case" & vbCrLf & _
                       "							 when SO.Deposito = 'S' and isnull(SO.ParaFinsDeExportacao,0) = 0" & vbCrLf & _
                       "							   then case when nf.entradaSaida_id = 'E' then Nfi.QuantidadeFiscal * -1 else Nfi.QuantidadeFiscal end" & vbCrLf & _
                       "							   else 0" & vbCrLf & _
                       "						   end),0) as EmPoderDeTerceiros," & vbCrLf & _
                       "			   isnull(sum(case" & vbCrLf & _
                       "							 when isnull(SO.ParaFinsDeExportacao,0) = 1" & vbCrLf & _
                       "							   then case when nf.entradaSaida_id = 'E' then Nfi.QuantidadeFiscal * -1 else Nfi.QuantidadeFiscal  end" & vbCrLf & _
                       "							   else 0" & vbCrLf & _
                       "						   end),0) as ParaFinsDeExportacao, " & vbCrLf & _
                       "			   isnull(sum(case " & vbCrLf & _
                       "							 when SO.ProdutoDeTerceiro = 0" & vbCrLf & _
                       "							   then case when nf.entradaSaida_id = 'E' then Nfi.QuantidadeFiscal else Nfi.QuantidadeFiscal * -1 end" & vbCrLf & _
                       "							   else 0" & vbCrLf & _
                       "						   end),0) as TotalEstoqueProprio," & vbCrLf & _
                       "			   isnull(sum(case" & vbCrLf & _
                       "							 when SO.ProdutoDeTerceiro = 1" & vbCrLf & _
                       "							   then case when nf.entradaSaida_id = 'E' then Nfi.QuantidadeFiscal else Nfi.QuantidadeFiscal * -1 end" & vbCrLf & _
                       "							   else 0" & vbCrLf & _
                       "						   end),0) as EstoqueDeTerceiro," & vbCrLf & _
                       "			   isnull(sum(case when nf.entradaSaida_id = 'E' then Nfi.QuantidadeFiscal else Nfi.QuantidadeFiscal * -1 end),0) as TotalEstoqueGeral" & vbCrLf & _
                       "		  FROM NotasFiscais AS nf" & vbCrLf & _
                       "		 INNER JOIN NotasFiscaisXItens Nfi" & vbCrLf & _
                       "			ON nf.Empresa_Id      = Nfi.Empresa_Id" & vbCrLf & _
                       "		   AND nf.EndEmpresa_Id   = Nfi.EndEmpresa_Id" & vbCrLf & _
                       "		   AND nf.Cliente_Id      = Nfi.Cliente_Id" & vbCrLf & _
                       "		   AND nf.EndCliente_Id   = Nfi.EndCliente_Id" & vbCrLf & _
                       "		   AND nf.EntradaSaida_Id = Nfi.EntradaSaida_Id" & vbCrLf & _
                       "		   AND nf.Serie_Id        = Nfi.Serie_Id" & vbCrLf & _
                       "		   AND nf.Nota_Id         = Nfi.Nota_Id" & vbCrLf & _
                       "		 Inner Join Clientes C" & vbCrLf & _
                       "			on Nf.Cliente_Id    = C.Cliente_Id" & vbCrLf & _
                       "		   and Nf.EndCliente_Id = C.Endereco_Id" & vbCrLf & _
                       "		 Inner join SubOperacoes SO" & vbCrLf & _
                       "			on SO.Operacao_Id     = Nfi.Operacao" & vbCrLf & _
                       "		   and SO.SubOperacoes_Id = Nfi.SubOperacao" & vbCrLf & _
                       "         Inner Join Produtos P" & vbCrLf & _
                       "            on Nfi.produto_id = P.Produto_Id" & vbCrLf & _
                       "		 WHERE NF.TipoDeDocumento = 1" & vbCrLf & _
                       "		   and NF.Situacao        = 1" & vbCrLf & _
                       "           and Nfi.CFOP_Id        <> 5923" & vbCrLf & _
                       "		   and year(nf.movimento) = " & ddlAno.SelectedValue & vbCrLf & _
                       "		   and " & selprd(0) & vbCrLf

                If ddlMesInicial.SelectedValue = 1 Then
                    sql &= "		  and MONTH(nf.movimento) between 0 and 0"
                Else
                    sql &= "		  and MONTH(nf.movimento) between 1 and " & ddlMesInicial.SelectedValue - 1 & vbCrLf
                End If

                If chkConsolidarEmpresa.Checked Then
                    sql &= "		  and left(NF.Empresa_id,8) = '" & Emp(0).Substring(0, 8) & "'" & vbCrLf
                Else
                    sql &= "		  and NF.Empresa_id    ='" & Emp(0) & "'" & vbCrLf & _
                           "          and NF.EndEmpresa_id = " & Emp(1) & vbCrLf
                End If

                sql &= "		  and SO.Classe not in ('" & eClassesOperacoes.GLOBAL.ToString & "','" & eClassesOperacoes.CONTAEORDEM.ToString & "')" & vbCrLf & _
                       "		union all" & vbCrLf & _
                       "		SELECT '3 - MES' as TAG," & vbCrLf & _
                       "		       year(nf.movimento)  as Ano," & vbCrLf & _
                       "			   month(nf.movimento) as MES," & vbCrLf & _
                       "			   SO.EntradaSaida," & vbCrLf & _
                       "			   Nfi.CFOP_Id," & vbCrLf & _
                       "			   sum(case" & vbCrLf & _
                       "					 when SO.ProdutoDeTerceiro = 0 and SO.Deposito = 'N' and isnull(SO.ParaFinsDeExportacao,0) = 0" & vbCrLf & _
                       "					   then Nfi.QuantidadeFiscal" & vbCrLf & _
                       "					   else 0" & vbCrLf & _
                       "				   end) as EmNossoPoder," & vbCrLf & _
                       "			   sum(case" & vbCrLf & _
                       "					 when SO.Deposito = 'S' and isnull(SO.ParaFinsDeExportacao,0) = 0" & vbCrLf & _
                       "					   then Nfi.QuantidadeFiscal" & vbCrLf & _
                       "					   else 0" & vbCrLf & _
                       "				   end) as EmPoderDeTerceiros," & vbCrLf & _
                       "			   	sum(case" & vbCrLf & _
                       "					 when isnull(SO.ParaFinsDeExportacao,0) = 1" & vbCrLf & _
                       "					   then Nfi.QuantidadeFiscal" & vbCrLf & _
                       "					   else 0" & vbCrLf & _
                       "				   end) as ParaFinsDeExportacao, " & vbCrLf & _
                       "				sum(case " & vbCrLf & _
                       "					 when SO.ProdutoDeTerceiro = 0" & vbCrLf & _
                       "					   then Nfi.QuantidadeFiscal" & vbCrLf & _
                       "					   else 0" & vbCrLf & _
                       "				   end) as TotalEstoqueProprio," & vbCrLf & _
                       "				sum(case" & vbCrLf & _
                       "					 when SO.ProdutoDeTerceiro = 1" & vbCrLf & _
                       "					   then Nfi.QuantidadeFiscal" & vbCrLf & _
                       "					   else 0" & vbCrLf & _
                       "				   end) as EstoqueDeTerceiro," & vbCrLf & _
                       "				sum(Nfi.QuantidadeFiscal) as TotalEstoqueGeral" & vbCrLf & _
                       "		  FROM NotasFiscais AS nf" & vbCrLf & _
                       "		 INNER JOIN NotasFiscaisXItens Nfi" & vbCrLf & _
                       "			ON nf.Empresa_Id      = Nfi.Empresa_Id" & vbCrLf & _
                       "		   AND nf.EndEmpresa_Id   = Nfi.EndEmpresa_Id" & vbCrLf & _
                       "		   AND nf.Cliente_Id      = Nfi.Cliente_Id" & vbCrLf & _
                       "		   AND nf.EndCliente_Id   = Nfi.EndCliente_Id" & vbCrLf & _
                       "		   AND nf.EntradaSaida_Id = Nfi.EntradaSaida_Id" & vbCrLf & _
                       "		   AND nf.Serie_Id        = Nfi.Serie_Id" & vbCrLf & _
                       "		   AND nf.Nota_Id         = Nfi.Nota_Id" & vbCrLf & _
                       "		 Inner Join Clientes C" & vbCrLf & _
                       "			on Nf.Cliente_Id    = C.Cliente_Id" & vbCrLf & _
                       "		   and Nf.EndCliente_Id = C.Endereco_Id" & vbCrLf & _
                       "		 Inner join SubOperacoes SO" & vbCrLf & _
                       "			on SO.Operacao_Id     = Nfi.Operacao" & vbCrLf & _
                       "		   and SO.SubOperacoes_Id = Nfi.SubOperacao" & vbCrLf & _
                       "         Inner join Produtos P" & vbCrLf & _
                       "            on Nfi.produto_id = P.Produto_Id" & vbCrLf & _
                       "		WHERE NF.TipoDeDocumento = 1" & vbCrLf & _
                       "		  and NF.Situacao        = 1" & vbCrLf & _
                       "          and Nfi.CFOP_Id        <> 5923" & vbCrLf & _
                       "		  and year(nf.movimento) = " & ddlAno.SelectedValue & vbCrLf & _
                       "		  and nf.movimento between '" & ddlAno.SelectedValue & "-" & ddlMesInicial.SelectedValue & "-01'" & " and '" & CDate(txtDataFinal.Text).ToString("yyyy-MM-dd") & "'" & vbCrLf & _
                       "          and " & selprd(0) & vbCrLf

                If chkConsolidarEmpresa.Checked Then
                    sql &= "		  and left(NF.Empresa_id,8) = '" & Emp(0).Substring(0, 8) & "'" & vbCrLf
                Else
                    sql &= "		  and NF.Empresa_id    ='" & Emp(0) & "'" & vbCrLf & _
                           "          and NF.EndEmpresa_Id = " & Emp(1) & vbCrLf
                End If


                sql &= "		  and SO.Classe not in ('" & eClassesOperacoes.CONTAEORDEM.ToString & "','" & eClassesOperacoes.GLOBAL.ToString & "')" & vbCrLf & _
                       "		group by year(nf.movimento)," & vbCrLf & _
                       "				 month(nf.movimento)," & vbCrLf & _
                       "				 SO.EntradaSaida," & vbCrLf & _
                       "				 Nfi.CFOP_Id " & vbCrLf & _
                       "	  ) sb" & vbCrLf & _
                       " order by sb.Ano, sb.Mes, sb.CFOP" & vbCrLf

                Dim ds As DataSet
                ds = Banco.ConsultaDataSet(sql, "EstoqueFiscalCFOP")

                Dim MenorMesConsulta As Integer = 12
                If ds.Tables(0).Rows.Count > 0 Then
                    For Each row In ds.Tables(0).Rows
                        If row("Mes") > 0 AndAlso row("Mes") < MenorMesConsulta Then
                            MenorMesConsulta = row("Mes")
                        End If
                    Next
                End If

                Dim parameters = New Dictionary(Of String, Object)()
                parameters.Add("PrimeiroMes", MenorMesConsulta)
                parameters.Add("OcultarCFOP", chkOcultarCFOP.Checked)
                parameters.Add("NomeEmpresa", Empresa.Nome & vbCrLf & Empresa.Cidade & "-" & Empresa.CodigoEstado)
                parameters.Add("Parametros", Parametros)

                Funcoes.BindReport(Me.Page, ds, "Cr_EstoqueFiscalPorCFOP", IIf(Pdf, eExportType.PDF, eExportType.ExcelCrystal), parameters)
            Else
                MsgBox(Me.Page, "Usuário sem permissão para emitir relatório.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkConsultar_Click(sender As Object, e As EventArgs) Handles lnkConsultar.Click
        Try
            Dim sql As String

            sql = "select E.Empresa_id + '-' + CAST(E.EndEmpresa_id as varchar) + '-' + Emp.fantasia + ' / ' + Emp.cidade + '-' + Emp.estado as Empresa," & vbCrLf & _
                  "       year(E.data_Id) as Ano," & vbCrLf & _
                  "       P.Produto_Id + '-' + p.nome as Produto,               " & vbCrLf & _
                  "	      isnull(sum(case                                       " & vbCrLf & _
                  "	                   when E.Tipo_id = 'NP'                    " & vbCrLf & _
                  "			             then E.Quantidade                      " & vbCrLf & _
                  "					     else 0                                 " & vbCrLf & _
                  "				     end),0) as EmNossoPoder,                   " & vbCrLf & _
                  "	      isnull(sum(case                                       " & vbCrLf & _
                  "	                   when E.Tipo_id = 'PT'                    " & vbCrLf & _
                  "					     then E.Quantidade                      " & vbCrLf & _
                  "					     else 0                                 " & vbCrLf & _
                  "				     end),0) as EmPoderDeTerceiros,             " & vbCrLf & _
                  "	      isnull(sum(case                                       " & vbCrLf & _
                  "	                   when E.Tipo_id = 'FE'                    " & vbCrLf & _
                  "					     then E.Quantidade                      " & vbCrLf & _
                  "					     else 0                                 " & vbCrLf & _
                  "				     end),0) as ParaFinsDeExportacao,           " & vbCrLf & _
                  "	      isnull(sum(case                                       " & vbCrLf & _
                  "	                   when E.Tipo_id = 'EP'                    " & vbCrLf & _
                  "			            then E.Quantidade                       " & vbCrLf & _
                  "					    else 0                                  " & vbCrLf & _
                  "				     end),0) as TotalEstoqueProprio,            " & vbCrLf & _
                  "	      isnull(sum(case                                       " & vbCrLf & _
                  "	                   when E.Tipo_id = 'ET'                    " & vbCrLf & _
                  "					     then E.Quantidade                      " & vbCrLf & _
                  "					     else 0                                 " & vbCrLf & _
                  "				     end),0) as EstoqueDeTerceiro,              " & vbCrLf & _
                  "       isnull(sum(case                                       " & vbCrLf & _
                  "	                when E.Tipo_id = 'EG'                       " & vbCrLf & _
                  "					  then E.Quantidade                         " & vbCrLf & _
                  "					  else 0                                    " & vbCrLf & _
                  "				  end),0) as TotalEstoqueGeral,                 " & vbCrLf & _
                  "        isnull(sum(case                                      " & vbCrLf & _
                  "	                   when E.Tipo_id = 'NP'                    " & vbCrLf & _
                  "			             then E.valor                           " & vbCrLf & _
                  "					     else 0                                 " & vbCrLf & _
                  "				     end),0) as EmNossoPoderValor,              " & vbCrLf & _
                  "	      isnull(sum(case                                       " & vbCrLf & _
                  "	                   when E.Tipo_id = 'PT'                    " & vbCrLf & _
                  "					     then E.valor                           " & vbCrLf & _
                  "					     else 0                                 " & vbCrLf & _
                  "				     end),0) as EmPoderDeTerceirosValor,        " & vbCrLf & _
                  "	      isnull(sum(case                                       " & vbCrLf & _
                  "	                   when E.Tipo_id = 'FE'                    " & vbCrLf & _
                  "					     then E.valor                           " & vbCrLf & _
                  "					     else 0                                 " & vbCrLf & _
                  "				     end),0) as ParaFinsDeExportacaoValor,      " & vbCrLf & _
                  "	      isnull(sum(case                                       " & vbCrLf & _
                  "	                   when E.Tipo_id = 'EP'                    " & vbCrLf & _
                  "			            then E.valor                            " & vbCrLf & _
                  "					    else 0                                  " & vbCrLf & _
                  "				     end),0) as TotalEstoqueProprioValor,       " & vbCrLf & _
                  "	      isnull(sum(case                                       " & vbCrLf & _
                  "	                   when E.Tipo_id = 'ET'                    " & vbCrLf & _
                  "					     then E.valor                           " & vbCrLf & _
                  "					     else 0                                 " & vbCrLf & _
                  "				     end),0) as EstoqueDeTerceiroValor,         " & vbCrLf & _
                  "       isnull(sum(case                                       " & vbCrLf & _
                  "	                when E.Tipo_id = 'EG'                       " & vbCrLf & _
                  "					  then E.valor                              " & vbCrLf & _
                  "					  else 0                                    " & vbCrLf & _
                  "				  end),0) as TotalEstoqueGeralValor             " & vbCrLf & _
                  "  from EstoqueInicialCFOP E                                  " & vbCrLf & _
                  " inner join Produtos P                                       " & vbCrLf & _
                  "    on E.Produto_id = P.Produto_Id                           " & vbCrLf & _
                  " inner join Clientes Emp                                     " & vbCrLf & _
                  "    on E.Empresa_id    = Emp.cliente_id                      " & vbCrLf & _
                  "   and E.EndEmpresa_id = Emp.Endereco_Id                     " & vbCrLf & _
                  "   AND year(E.data_Id) =                                     " & ddlAnoSI.SelectedValue & vbCrLf

            If Not String.IsNullOrWhiteSpace(ddlEmpresaSI.SelectedValue) Then
                sql &= "AND E.EMPRESA_ID = '" & ddlEmpresaSI.SelectedValue.Split("-")(0) & "'" & vbCrLf & _
                    "AND E.ENDEMPRESA_ID = " & ddlEmpresaSI.SelectedValue.Split("-")(1) & vbCrLf
            End If
            If Not String.IsNullOrWhiteSpace(txtNomeProduto.Text) Then
                sql &= "AND P.Produto_Id = '" & txtNomeProduto.Text.Split("-")(0) & "'" & vbCrLf
            End If

            sql &= " Group by E.Empresa_id + '-' + CAST(E.EndEmpresa_id as varchar) + '-' + Emp.fantasia + ' / ' + Emp.cidade + '-' + Emp.estado," & vbCrLf & _
                "          year(E.data_Id)," & vbCrLf & _
                "          P.Produto_Id + '-' + p.nome" & vbCrLf

            Dim ds As DataSet = Banco.ConsultaDataSet(sql, "SI")
            Session("dsConsulta") = ds

            If ds IsNot Nothing AndAlso ds.Tables IsNot Nothing AndAlso ds.Tables(0).Rows.Count > 0 Then
                gridSI.DataSource = ds
                gridSI.DataBind()
            Else
                MsgBox(Me.Page, "Nenhum resultado encontrado referente a pesquisa.")
                gridSI.DataSource = New List(Of Object)
                gridSI.DataBind()
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub BtnProdutoSI_Click(sender As Object, e As EventArgs) Handles BtnProdutoSI.Click
        Try
            ucConsultaProduto.Limpar()
            ucConsultaProduto.BuscarProduto()
            Popup.ConsultaDeProduto(Me.Page, "objProdutoCFOP" & HID.Value, "txtNome", True)
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

    Protected Sub gridSI_SelectedIndexChanged(sender As Object, e As EventArgs) Handles gridSI.SelectedIndexChanged
        Try
            Dim ds As DataSet = CType(Session("dsConsulta"), DataSet)

            ddlEmpresaSI.SelectedValue = gridSI.SelectedRow.Cells(1).Text.Split("-")(0) & "-" & gridSI.SelectedRow.Cells(1).Text.Split("-")(1)
            txtNomeProduto.Text = gridSI.SelectedRow.Cells(3).Text.Split("-")(0) & "-" & gridSI.SelectedRow.Cells(3).Text.Split("-")(1)
            ddlAnoSI.SelectedValue = gridSI.SelectedRow.Cells(2).Text
            txtNP.Text = gridSI.SelectedRow.Cells(4).Text
            txtPT.Text = gridSI.SelectedRow.Cells(5).Text
            txtFE.Text = gridSI.SelectedRow.Cells(6).Text
            txtEP.Text = gridSI.SelectedRow.Cells(7).Text
            txtET.Text = gridSI.SelectedRow.Cells(8).Text
            txtEG.Text = gridSI.SelectedRow.Cells(9).Text

            txtNPValor.Text = ds.Tables(0).Rows(gridSI.SelectedIndex)(9)
            txtPTValor.Text = ds.Tables(0).Rows(gridSI.SelectedIndex)(10)
            txtFEValor.Text = ds.Tables(0).Rows(gridSI.SelectedIndex)(11)
            txtEPValor.Text = ds.Tables(0).Rows(gridSI.SelectedIndex)(12)
            txtETValor.Text = ds.Tables(0).Rows(gridSI.SelectedIndex)(13)
            txtEGValor.Text = ds.Tables(0).Rows(gridSI.SelectedIndex)(14)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkNovo_Click(sender As Object, e As EventArgs) Handles lnkNovo.Click
        Try
            If ValidaCampos() Then
                AtualizaRegistros()
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

#End Region

#Region "Métodos"

    Public Function valida() As Boolean
        If Not ddlEmpresa.SelectedIndex > 0 Then
            MsgBox(Me.Page, "Informe a empresa.")
            Return False
        End If
        If Not ucSelecaoProduto.TemSelecionado Then
            MsgBox(Me.Page, "Informe o produto.")
            Return False
        End If

        Return True
    End Function

    Public Overrides Sub Carregar(ByVal obj As IBaseEntity)
        If Session("objProdutoCFOP" & HID.Value) IsNot Nothing Then
            Dim objProduto As [Lib].Negocio.Produto = Session("objProdutoCFOP" & HID.Value)
            txtNomeProduto.Text = objProduto.Codigo & "-" & objProduto.Descricao
            Session.Remove("objProdutoCFOP" & HID.Value)
        End If
    End Sub

    Private Sub Limpar()
        txtNomeProduto.Text = String.Empty
        ddlAnoSI.SelectedIndex = 0
        txtNP.Text = String.Empty
        txtPT.Text = String.Empty
        txtFE.Text = String.Empty
        txtEP.Text = String.Empty
        txtET.Text = String.Empty
        txtEG.Text = String.Empty
        txtNPValor.Text = String.Empty
        txtPTValor.Text = String.Empty
        txtFEValor.Text = String.Empty
        txtEPValor.Text = String.Empty
        txtETValor.Text = String.Empty
        txtEGValor.Text = String.Empty

        gridSI.DataSource = New List(Of Object)
        gridSI.DataBind()
        LiberaEmpresa()
    End Sub

    Private Sub LiberaEmpresa()
        If Not UsuarioServidor.LiberaEmpresa Then
            ddlEmpresa.Enabled = False
            ddlEmpresaSI.Enabled = False
        End If
    End Sub

    Private Sub AtualizaRegistros()
        Try
            Dim sql As String = ""
            Dim sqls As New ArrayList()

            Dim tiposValores As String() = {"NP", "PT", "FE", "EP", "ET", "EG"}
            Dim qtdesTipos As String() = {txtNP.Text, txtPT.Text, txtFE.Text, txtEP.Text, txtET.Text, txtEG.Text}
            Dim valoresTipos As String() = {txtNPValor.Text, txtPTValor.Text, txtFEValor.Text, txtEPValor.Text, txtETValor.Text, txtEGValor.Text}

            sql = "DELETE EstoqueInicialCFOP" & vbCrLf & _
                 "	WHERE	Empresa_Id		= '" & ddlEmpresaSI.SelectedValue.Split("-")(0) & "'" & vbCrLf & _
                 "		AND EndEmpresa_Id	= " & ddlEmpresaSI.SelectedValue.Split("-")(1) & vbCrLf & _
                 "		AND Produto_Id		= '" & txtNomeProduto.Text.Split("-")(0) & "'" & vbCrLf & _
                 "		AND YEAR(Data_Id)	= " & ddlAnoSI.SelectedValue & "" & vbCrLf
            sqls.Add(sql)

            For i As Integer = 0 To tiposValores.Length - 1
                sql = "INSERT INTO EstoqueInicialCFOP " & vbCrLf & _
                      "         (EMPRESA_ID, EndEmpresa_Id, Produto_Id, FisicoFiscal_Id, Data_Id, Tipo_Id, Quantidade, Valor)" & vbCrLf & _
                      "     VALUES" & vbCrLf & _
                      "         ('" & ddlEmpresaSI.SelectedValue.Split("-")(0) & "', " & ddlEmpresaSI.SelectedValue.Split("-")(1) & ", '" & txtNomeProduto.Text.Split("-")(0) & "', 1," & _
                      " '" & ddlAnoSI.SelectedValue & "-01-01', '" & tiposValores(i) & "', " & Str(IIf(String.IsNullOrWhiteSpace(qtdesTipos(i)), "0", qtdesTipos(i))) & ", " & Str(IIf(String.IsNullOrWhiteSpace(valoresTipos(i)), "0", valoresTipos(i))) & ")" & vbCrLf
                sqls.Add(sql)
            Next

            If Banco.GravaBanco(sqls) Then
                MsgBox(Me.Page, "Dados atualizados com Sucesso.", eTitulo.Sucess)
                Limpar()
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    'Private Sub AtualizaRegistros()
    '    Dim sql As String = ""

    '    sql = "UPDATE EstoqueInicialCFOP                   " & vbCrLf & _
    '          "	SET Quantidade	=	case                            " & vbCrLf & _
    '          "							WHEN  Tipo_id = 'NP' THEN " & Str(IIf(txtNP.Text = "", 0, txtNP.Text)) & vbCrLf & _
    '          "							WHEN  Tipo_id = 'PT' THEN " & Str(IIf(txtPT.Text = "", 0, txtPT.Text)) & vbCrLf & _
    '          "							WHEN  Tipo_id = 'FE' THEN " & Str(IIf(txtFE.Text = "", 0, txtFE.Text)) & vbCrLf & _
    '          "							WHEN  Tipo_id = 'EP' THEN " & Str(IIf(txtEP.Text = "", 0, txtEP.Text)) & vbCrLf & _
    '          "							WHEN  Tipo_id = 'ET' THEN " & Str(IIf(txtET.Text = "", 0, txtET.Text)) & vbCrLf & _
    '          "							WHEN  Tipo_id = 'EG' THEN " & Str(IIf(txtEG.Text = "", 0, txtEG.Text)) & vbCrLf & _
    '          "						end,                            " & vbCrLf & _
    '          "		Valor		=	case                            " & vbCrLf & _
    '          "							WHEN  Tipo_id = 'NP' THEN " & Str(IIf(txtNPValor.Text = "", 0, txtNPValor.Text)) & vbCrLf & _
    '          "							WHEN  Tipo_id = 'PT' THEN " & Str(IIf(txtPTValor.Text = "", 0, txtPTValor.Text)) & vbCrLf & _
    '          "							WHEN  Tipo_id = 'FE' THEN " & Str(IIf(txtFEValor.Text = "", 0, txtFEValor.Text)) & vbCrLf & _
    '          "							WHEN  Tipo_id = 'EP' THEN " & Str(IIf(txtEPValor.Text = "", 0, txtEPValor.Text)) & vbCrLf & _
    '          "							WHEN  Tipo_id = 'ET' THEN " & Str(IIf(txtETValor.Text = "", 0, txtETValor.Text)) & vbCrLf & _
    '          "							WHEN  Tipo_id = 'EG' THEN " & Str(IIf(txtEGValor.Text = "", 0, txtEGValor.Text)) & vbCrLf & _
    '          "        End                                          " & vbCrLf & _
    '          "	WHERE	Empresa_Id		= '" & ddlEmpresaSI.SelectedValue.Split("-")(0) & "'" & vbCrLf & _
    '          "		AND EndEmpresa_Id	= " & ddlEmpresaSI.SelectedValue.Split("-")(1) & vbCrLf & _
    '          "		AND Produto_Id		= '" & txtNomeProduto.Text.Split("-")(0) & "'" & vbCrLf & _
    '          "		AND YEAR(Data_Id)	= " & ddlAnoSI.SelectedValue & "" & vbCrLf

    '    If Banco.GravaBanco(sql) Then
    '        MsgBox(Me.Page, "Dados atualizados com Sucesso.", eTitulo.Sucess)
    '        Limpar()
    '    End If
    'End Sub

    Private Function ValidaCampos() As Boolean
        If String.IsNullOrWhiteSpace(ddlEmpresaSI.SelectedValue) Then
            MsgBox(Me.Page, "Informe a Empresa.")
            Return False
        ElseIf String.IsNullOrWhiteSpace(txtNomeProduto.Text) Then
            MsgBox(Me.Page, "Informe o produto.")
            Return False
        End If
        Return True
    End Function

#End Region

    'Protected Sub txtEP_TextChanged(sender As Object, e As EventArgs) Handles txtEP.TextChanged
    '    Try
    '        txtEG.Text = CInt(IIf(String.IsNullOrWhiteSpace(txtEP.Text), 0, txtEP.Text)) + CInt(IIf(String.IsNullOrWhiteSpace(txtET.Text), 0, txtET.Text))
    '    Catch ex As Exception
    '        MsgBox(Me.Page, ex.Message, eTitulo.Erro)
    '    End Try
    'End Sub

    'Protected Sub txtET_TextChanged(sender As Object, e As EventArgs) Handles txtET.TextChanged
    '    Try
    '        txtEG.Text = CInt(IIf(String.IsNullOrWhiteSpace(txtEP.Text), 0, txtEP.Text)) + CInt(IIf(String.IsNullOrWhiteSpace(txtET.Text), 0, txtET.Text))
    '    Catch ex As Exception
    '        MsgBox(Me.Page, ex.Message, eTitulo.Erro)
    '    End Try
    'End Sub

    'Protected Sub txtETValor_TextChanged(sender As Object, e As EventArgs) Handles txtETValor.TextChanged
    '    Try
    '        txtEGValor.Text = CInt(IIf(String.IsNullOrWhiteSpace(txtEPValor.Text), 0, txtEPValor.Text)) + CInt(IIf(String.IsNullOrWhiteSpace(txtETValor.Text), 0, txtETValor.Text))
    '    Catch ex As Exception
    '        MsgBox(Me.Page, ex.Message, eTitulo.Erro)
    '    End Try

    'End Sub

    'Protected Sub txtEPValor_TextChanged(sender As Object, e As EventArgs) Handles txtEPValor.TextChanged
    '    Try
    '        txtEGValor.Text = CInt(IIf(String.IsNullOrWhiteSpace(txtEPValor.Text), 0, txtEPValor.Text)) + CInt(IIf(String.IsNullOrWhiteSpace(txtETValor.Text), 0, txtETValor.Text))
    '    Catch ex As Exception
    '        MsgBox(Me.Page, ex.Message, eTitulo.Erro)
    '    End Try
    'End Sub

End Class