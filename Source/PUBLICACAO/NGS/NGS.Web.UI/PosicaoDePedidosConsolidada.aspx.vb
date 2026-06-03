Imports System.Data
Imports System.IO
Imports Microsoft.VisualBasic
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Partial Class PosicaoDePedidosConsolidada
    Inherits BasePage

#Region "Events"

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not IsPostBack And IsConnect Then
            Me.setMenu(eModulo.Comercial)
            If Funcoes.VerificaPermissao("PosicaoDePedidos", "ACESSAR") Then
                HID.Value = Guid.NewGuid.ToString
                txtData.Text = "" 'No_w.Date.ToString("dd/MM/yyyy")
                CarregarEmpresa1()
                Funcoes.VerificaEmpresa(lstEmpresa1)
                ddl.Carregar(ddlSafra, CarregarDDL.Tabela.Safra)
                ucSelecaoProduto.WhereProduto = "Agrupar = 'N'"
            Else
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Comercial.aspx")
                Exit Sub
            End If
        End If
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

    Protected Sub lnkLimpar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkLimpar.Click
        Try
            lstEmpresa1.ClearSelection()
            txtCodigoCliente.Value = 0
            txtCliente.Text = String.Empty
            chkClienteConsolidado.Checked = False
            chkPedidosComSaldo.Checked = False
            ddlFretes.SelectedIndex = 0
            ucSelecaoProduto.CarregarNivel(5)
            ucSelecaoProduto.AbrirAccordion(ucSelecaoProduto.ID, False, True, 100)
            ucSelecaoProduto.Limpar()
            ddlSafra.SelectedIndex = 0
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnCliente_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCliente.Click
        Try
            ucConsultaClientes.Limpar()
            Popup.ConsultaDeClientes(Me.Page, "objCliente" & HID.Value, "txtNome")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

#End Region

#Region "Methods"

    Private Function ValidarCampos() As Boolean
        If String.IsNullOrWhiteSpace(txtData.Text) Or Not IsDate(txtData.Text) Then
            MsgBox(Me.Page, "Informe uma data válida.")
            txtData.Focus()
            Return False
        ElseIf Not ucSelecaoProduto.TemSelecionado Then
            MsgBox(Me.Page, "Selecione ao menos 1 produto.")
            Return False
        ElseIf String.IsNullOrWhiteSpace(ddlSafra.SelectedValue) Then
            MsgBox(Me.Page, "Selecione uma Safra.")
            Return False
        End If
        Return True
    End Function

    Private Function ComporRelatorio(Saldo As Boolean) As String
        Dim Sql As String = String.Empty
        Dim temselecionados As Boolean = lstEmpresa1.GetSelectedValues.Count > 0
        Dim empresasselecionadas As String = String.Join("','", lstEmpresa1.GetSelecteds())

        If ddlFretes.SelectedItem.Text <> "FOB" Then
            Sql &= " /*-------------------------------------------------*/" & vbCrLf & _
              " DECLARE @InicioDeSafra AS Date, @Vencimento AS Date" & vbCrLf & _
              "  SELECT @InicioDeSafra = S.InicioDeSafra, @Vencimento = S.Vencimento " & vbCrLf & _
              "    FROM Safras S" & vbCrLf & _
              "    JOIN #Safra S2" & vbCrLf & _
              "      ON S.Safra_Id = S2.Safra_Id" & vbCrLf & _
              " /*-------------------------------------------------*/" & vbCrLf & _
              " select sb.Cliente AS CodigoDoCliente," & vbCrLf & _
              "        sb.Nome AS NomeDoCliente," & vbCrLf & _
              "        SUM(sb.Contratado)      AS Contratado," & vbCrLf & _
              "        SUM(sb.ContratadoTroca) AS ContratadoPorTroca," & vbCrLf & _
              "        SUM(sb.Entregue)        AS Entregue," & vbCrLf & _
              "        SUM(sb.Cedente)         AS Cedente," & vbCrLf & _
              "        SUM(sb.Cessionario)     AS Cessionario," & vbCrLf & _
              "        SUM(convert(numeric(18,4), isnull(ls.LaudosSemNota,0))) as LaudosSemNota," & vbCrLf & _
              "        SUM(sb.Contratado + sb.ContratadoTroca - (sb.Entregue) + sb.cedente - sb.cessionario) as Saldo, --+ convert(numeric(18,4),isnull(ls.LaudosSemNota,0))" & vbCrLf & _
              "        SUM(sb.AdiantamentoReais) AS AdiantamentoEmReais," & vbCrLf & _
              "        SUM(sb.AdiantamentoDolar) AS AdiantamentoEmDolares," & vbCrLf & _
              "        (isnull(fin.reais,0))  AS Reais," & vbCrLf & _
              "        (isnull(fin.dolar,0))  AS Dolares" & vbCrLf & _
              "   INTO #SaldoFinal " & vbCrLf & _
              "   from (   " & vbCrLf & _
              "	   select b.empresa," & vbCrLf & _
              "	          b.cliente, " & vbCrLf & _
              IIf(Not chkClienteConsolidado.Checked, "	          b.EndCliente,", "") & vbCrLf & _
              "			  (select top 1 nome from clientes where cliente_id = b.cliente " & IIf(Not chkClienteConsolidado.Checked, " AND Endereco_Id = B.EndCliente ", " ") & ") as Nome," & vbCrLf & _
              "			  b.Produto," & vbCrLf & _
              "			  b.NomeProduto," & vbCrLf & _
              "			  sum(case " & vbCrLf & _
              "					 when isnull(p.troca,0) = 0 and so.precofixo = 'S' and so.classe <> '" & eClassesOperacoes.DEPOSITOS.ToString & "'" & vbCrLf & _
              "					   then Contratado" & vbCrLf & _
              "					   else 0" & vbCrLf & _
              "				   end)" & vbCrLf & _
              "				+   " & vbCrLf & _
              "			  sum(case " & vbCrLf & _
              "					 when isnull(p.troca,0) = 0 and so.classe = '" & eClassesOperacoes.AFIXAR.ToString & "'" & vbCrLf & _
              "					   then fixada" & vbCrLf & _
              "					   else 0" & vbCrLf & _
              "				   end) as contratado," & vbCrLf & _
              "			   sum(case " & vbCrLf & _
              "					 when isnull(p.troca,0) = 1 and so.precofixo = 'S' and so.classe <> '" & eClassesOperacoes.DEPOSITOS.ToString & "'" & vbCrLf & _
              "					   then Contratado" & vbCrLf & _
              "					   else 0" & vbCrLf & _
              "				   end) " & vbCrLf & _
              "			   +" & vbCrLf & _
              "			   sum(case " & vbCrLf & _
              "					 when isnull(p.troca,0) = 1 and so.classe = '" & eClassesOperacoes.AFIXAR.ToString & "'" & vbCrLf & _
              "					   then fixada" & vbCrLf & _
              "					   else 0" & vbCrLf & _
              "				   end) 	   				   " & vbCrLf & _
              "				   as ContratadoTroca,		           " & vbCrLf & _
              "			   sum(b.laudo) as Entregue,                 " & vbCrLf & _
              "			   sum(case " & vbCrLf & _
              "					 when so.precofixo = 'N' and b.moeda = 1" & vbCrLf & _
              "					   then b.adiantamento" & vbCrLf & _
              "					   else 0" & vbCrLf & _
              "				   end) as AdiantamentoReais,		           " & vbCrLf & _
              "				sum(case " & vbCrLf & _
              "					 when so.precofixo = 'N' and b.moeda = 3" & vbCrLf & _
              "					   then b.adiantamento" & vbCrLf & _
              "					   else 0" & vbCrLf & _
              "				   end) as AdiantamentoDolar," & vbCrLf & _
              "                sum(Cedente) as cedente," & vbCrLf & _
              "                sum(Cessionario) as Cessionario	" & vbCrLf & _
              "		   from #base b" & vbCrLf & _
              "		  inner join Pedidos P" & vbCrLf & _
              "		     on b.empresa    = P.Empresa_id" & vbCrLf & _
              "		    and b.endEmpresa = P.EndEmpresa_id" & vbCrLf & _
              "		    and b.Pedido     = P.Pedido_id" & vbCrLf & _
              "		   Left Join(SELECT PxD.Empresa_Id,  PxD.EndEmpresa_Id, PxD.Pedido_Id," & vbCrLf & _
              "		                    PxD.Deposito_Id, PxD.EndDeposito_Id" & vbCrLf & _
              "                       FROM PedidosXDepositos PxD" & vbCrLf & _
              "                      where PxD.Tipo = 'OD'" & vbCrLf & _
              "                        and PxD.Principal = 1" & vbCrLf & _
              "		             ) DE" & vbCrLf & _
              "		     on P.Empresa_id    = DE.Empresa_id" & vbCrLf & _
              "		    and P.EndEmpresa_id = DE.EndEmpresa_id" & vbCrLf & _
              "		    and P.Pedido_id     = DE.Pedido_id" & vbCrLf & _
              "		   Left Join(SELECT PxD.Empresa_Id,  PxD.EndEmpresa_Id, PxD.Pedido_Id," & vbCrLf & _
              "		                    PxD.Deposito_Id, PxD.EndDeposito_Id" & vbCrLf & _
              "                       FROM PedidosXDepositos PxD" & vbCrLf & _
              "                      where PxD.Tipo = 'LE'" & vbCrLf & _
              "                        and PxD.Principal = 1" & vbCrLf & _
              "		             ) LE" & vbCrLf & _
              "		     on P.Empresa_id    = LE.Empresa_id" & vbCrLf & _
              "		    and P.EndEmpresa_id = LE.EndEmpresa_id" & vbCrLf & _
              "		    and P.Pedido_id     = LE.Pedido_id" & vbCrLf & _
              "		  inner join Clientes Embarque" & vbCrLf & _
              "		     on Embarque.Cliente_Id  = isnull(LE.Deposito_Id   ,DE.Deposito_Id)" & vbCrLf & _
              "		    and Embarque.Endereco_Id = isnull(LE.EndDeposito_Id,DE.EndDeposito_Id)" & vbCrLf & _
              "		  inner join Suboperacoes SO" & vbCrLf & _
              "		     on b.Operacao    = SO.Operacao_id" & vbCrLf & _
              "		    and b.Suboperacao = SO.SubOperacoes_Id " & vbCrLf & _
              "		  inner join Operacoes OP" & vbCrLf & _
              "		     on OP.operacao_id = SO.Operacao_Id  " & vbCrLf & _
              "		  WHERE 1=1" & vbCrLf

            If temselecionados Then
                Sql &= "  AND b.Empresa + cast(b.EndEmpresa as varchar) in ('" & empresasselecionadas & "')" & vbCrLf
            End If

            If Not String.IsNullOrWhiteSpace(txtCliente.Text) Then
                If chkClienteConsolidado.Checked Then
                    Sql &= "And left(b.Cliente, 8) = '" & Left(txtCodigoCliente.Value.Split("-")(0), 8) & "'"
                Else
                    Sql &= "And b.Cliente    = '" & txtCodigoCliente.Value.Split("-")(0) & "'" & vbCrLf & _
                           "And b.EndCliente = " & txtCodigoCliente.Value.Split("-")(1) & vbCrLf
                End If
            End If

            Sql &= "  And OP.classe in ('" & eClassesOperacoes.COMPRAS.ToString & "','" & eClassesOperacoes.DEPOSITOS.ToString & "') " & vbCrLf & _
            "		  group by b.empresa,b.cliente, b.Produto,b.Nomeproduto" & IIf(Not chkClienteConsolidado.Checked, ", b.EndCliente", "") & vbCrLf & _
            "		) as sb  " & vbCrLf & _
            "  left join " & vbCrLf

            If FinanceiroNovo Then
                Sql &= "               (select clifor as cliente," & vbCrLf & _
                       "                       sum(case" & vbCrLf & _
                       "               		         when T.Moeda = 1" & vbCrLf & _
                       "               			       then ValorOficial" & vbCrLf & _
                       "               			       else 0" & vbCrLf & _
                       "               		       end) reais,    " & vbCrLf & _
                       "                       sum(case" & vbCrLf & _
                       "               			     when T.Moeda = 3" & vbCrLf & _
                       "               			       then ValorMoeda" & vbCrLf & _
                       "               			       else 0" & vbCrLf & _
                       "               		       end) dolar    " & vbCrLf & _
                       "                  from Titulos T" & vbCrLf & _
                       "                  JOIN TitulosxContaContabil Tc" & vbCrLf & _
                       "                    ON T.Titulo_Id = Tc.Titulo_Id" & vbCrLf & _
                       "                   AND T.ContaContabilCliFor = Tc.Conta_Id " & vbCrLf & _
                       "                  left join Pedidos p " & vbCrLf & _
                       "                    on t.EmpresaPedido    = p.Empresa_id" & vbCrLf & _
                       "                   and t.EndEmpresaPedido = p.EndEmpresa_id" & vbCrLf & _
                       "                   and t.Pedido           = p.Pedido_id" & vbCrLf & _
                       "                 where T.situacao = 1" & vbCrLf & _
                       "                   and provisao <> 1 " & vbCrLf & _
                       "                   and (T.reprogramacao < @Vencimento or cr.prorrogacao < '" & CDate(txtData.Text).ToString("yyyy-MM-dd") & "' or isnull(p.safra,'') = '" & ddlSafra.SelectedValue & "')" & vbCrLf

                If temselecionados Then
                    Sql &= "                   and Empresa + cast(EndEmpresa as varchar) in ('" & empresasselecionadas & "') " & vbCrLf
                End If

                Sql &= "                   AND T.RecPag = 'R'" & vbCrLf & _
                       "                 group by clifor" & vbCrLf & _
                       "                ) Fin" & vbCrLf
            Else
                Sql &= "              (select cr.cliente," & vbCrLf & _
                       "	                  sum(case" & vbCrLf & _
                       "							 when cr.moeda = 1" & vbCrLf & _
                       "							   then cr.valordodocumento" & vbCrLf & _
                       "							   else 0" & vbCrLf & _
                       "						   end) reais,    " & vbCrLf & _
                       "					   sum(case" & vbCrLf & _
                       "							 when cr.moeda = 3" & vbCrLf & _
                       "							   then cr.Moedavalordodocumento" & vbCrLf & _
                       "							   else 0" & vbCrLf & _
                       "						   end) dolar    " & vbCrLf & _
                       "				 from contasareceber cr" & vbCrLf & _
                       "				 left join Pedidos p" & vbCrLf & _
                       "                   on cr.EmpresaPedido    = p.Empresa_id" & vbCrLf & _
                       "                  and cr.EndEmpresaPedido = p.EndEmpresa_id" & vbCrLf & _
                       "                  and cr.Pedido           = p.Pedido_id" & vbCrLf & _
                       "				where cr.situacao = 1 " & vbCrLf & _
                       "				  and cr.provisao <> 1 " & vbCrLf & _
                       "				  and (cr.prorrogacao < @Vencimento or cr.prorrogacao < '" & CDate(txtData.Text).ToString("yyyy-MM-dd") & "' or isnull(p.safra,'') = '" & ddlSafra.SelectedValue & "')" & vbCrLf

                If temselecionados Then
                    Sql &= "				  And cr.Empresa + cast(cr.EndEmpresa as varchar) in ('" & empresasselecionadas & "')" & vbCrLf
                End If

                Sql &= "				group by cr.cliente  " & vbCrLf & _
                       "			   ) Fin" & vbCrLf
            End If

            Sql &= "	  on sb.cliente = fin.cliente" & vbCrLf & _
                   "  left join (" & vbCrLf & _
                   "				SELECT R.Empresa_Id, R.EndEmpresa_Id, Ped.cliente, R.Produto, sum(R.PesoLiquido) as LaudosSemNota" & IIf(Not chkClienteConsolidado.Checked, ", Ped.Endcliente", "") & vbCrLf & _
                   "				  FROM Romaneios R" & vbCrLf & _
                   "				 INNER JOIN RomaneiosXPesagens RxP " & vbCrLf & _
                   "					ON R.Empresa_Id       = RxP.Empresa_Id " & vbCrLf & _
                   "				   AND R.EndEmpresa_Id    = RxP.EndEmpresa_Id" & vbCrLf & _
                   "				   AND R.Romaneio_Id      = RxP.Romaneio_Id" & vbCrLf & _
                   "				  LEFT JOIN Pedidos Ped" & vbCrLf & _
                   "					ON R.Empresa_Id    = Ped.Empresa_Id" & vbCrLf & _
                   "				   AND R.EndEmpresa_Id = Ped.EndEmpresa_Id" & vbCrLf & _
                   "				   AND R.Pedido        = Ped.Pedido_Id" & vbCrLf & _
                   "				  LEFT OUTER JOIN Clientes" & vbCrLf & _
                   "					ON Ped.Cliente    = Clientes.Cliente_Id" & vbCrLf & _
                   "				   AND Ped.EndCliente = Clientes.Endereco_Id" & vbCrLf & _
                   "				  LEFT OUTER JOIN NotasFiscaisXRomaneios nfxr" & vbCrLf & _
                   "					ON R.Empresa_Id    = nfxr.Empresa_Id" & vbCrLf & _
                   "				   AND R.EndEmpresa_Id = nfxr.EndEmpresa_Id" & vbCrLf & _
                   "				   AND R.Romaneio_Id   = nfxr.Romaneio_Id " & vbCrLf & _
                   "				   --AND Ped.Cliente         = nfxr.Cliente_id" & vbCrLf & _
                   "				   --AND Ped.EndCliente      = nfxr.EndCliente_id" & vbCrLf & _
                   "                 INNER JOIN Pesagem P" & vbCrLf & _
                   "                    ON RxP.Empresa_Id    = P.Empresa_Id" & vbCrLf & _
                   "                   AND RxP.EndEmpresa_Id = P.EndEmpresa_Id" & vbCrLf & _
                   "                   AND RxP.Pesagem_Id    = P.Pesagem_Id" & vbCrLf & _
                   "                   AND P.Sequencia_Id             = 0" & vbCrLf & _
                   "                   AND P.Situacao                 = 1" & vbCrLf & _
                   "				  LEFT JOIN NotasFiscais NF" & vbCrLf & _
                   "					ON NF.Empresa_Id      = nfxr.Empresa_Id" & vbCrLf & _
                   "				   AND NF.EndEmpresa_Id   = nfxr.EndEmpresa_Id" & vbCrLf & _
                   "				   AND NF.Cliente_Id      = nfxr.Cliente_Id " & vbCrLf & _
                   "				   AND NF.EndCliente_Id   = nfxr.EndCliente_Id" & vbCrLf & _
                   "				   AND NF.EntradaSaida_Id = nfxr.EntradaSaida_Id" & vbCrLf & _
                   "				   AND NF.Serie_Id        = nfxr.Serie_Id " & vbCrLf & _
                   "				   AND NF.Nota_Id         = nfxr.Nota_Id  " & vbCrLf & _
                   "				   --AND NF.Movimento       <=  @Vencimento" & vbCrLf & _
                   "				 Where NF.Nota_Id   IS NULL" & vbCrLf & _
                   "				   --AND R.Movimento       >=  @InicioDeSafra" & vbCrLf & _
                   "				   --AND R.Movimento       <=  @Vencimento" & vbCrLf & _
                   "				   and R.EntradaSaida = 'E'" & vbCrLf & _
                   "                   and Ped.safra = '" & ddlSafra.SelectedValue & "'" & vbCrLf & _
                   "				 group by R.Empresa_Id, R.EndEmpresa_Id, Ped.cliente,R.Produto" & IIf(Not chkClienteConsolidado.Checked, ", Ped.Endcliente", "") & vbCrLf & _
                   "	          ) LS" & vbCrLf & _
                   "	    on sb.empresa = LS.Empresa_Id" & vbCrLf & _
                   "	   and sb.cliente = LS.cliente" & vbCrLf & _
                 IIf(Not chkClienteConsolidado.Checked, "	   and sb.EndCliente = LS.EndCliente", "") & vbCrLf & _
                   "	   and sb.produto = LS.Produto" & vbCrLf & _
                   "     Group by sb.Cliente," & vbCrLf & _
                   "              sb.Nome," & vbCrLf & _
                   "              isnull(fin.reais,0)," & vbCrLf & _
                   "              isnull(fin.dolar,0)" & vbCrLf & _
                   "      order by sb.nome  " & vbCrLf & _
                   "   /*---------------------------------*/  " & vbCrLf

            If Saldo Then
                Sql &= " SELECT * FROM #SaldoFinal " & vbCrLf & _
                       "  WHERE CodigoDoCliente in (" & vbCrLf & _
                       "                            SELECT CodigoDoCliente" & vbCrLf & _
                       "                              FROM #SaldoFinal" & vbCrLf & _
                       "                             GROUP BY CodigoDoCliente" & vbCrLf & _
                       "                            HAVING SUM(Saldo) + sum(LaudosSemNota) > 0)" & vbCrLf
            Else
                Sql &= " SELECT * FROM #SaldoFinal Order by NomeDoCliente " & vbCrLf
            End If
        Else
            '***************************************************************************************************************************************************
            '**********************************************  FOB  **********************************************************************************************
            '***************************************************************************************************************************************************

            Sql &= " /*-------------------------------------------------*/" & vbCrLf & _
              " DECLARE @InicioDeSafra AS Date, @Vencimento AS Date" & vbCrLf & _
              "  SELECT @InicioDeSafra = S.InicioDeSafra, @Vencimento = S.Vencimento " & vbCrLf & _
              "    FROM Safras S" & vbCrLf & _
              "    JOIN #Safra S2" & vbCrLf & _
              "      ON S.Safra_Id = S2.Safra_Id" & vbCrLf & _
              " /*-------------------------------------------------*/" & vbCrLf & _
              " select sb.Cliente AS CodigoDoCliente," & vbCrLf & _
              "        sb.Nome AS NomeDoCliente," & vbCrLf & _
              "        sb.Embarque," & vbCrLf & _
              "        sb.DataEntrega," & vbCrLf & _
              "        SUM(sb.Contratado)      as Contratado," & vbCrLf & _
              "        SUM(sb.ContratadoTroca) as ContratadoPorTroca," & vbCrLf & _
              "        SUM(sb.Entregue)        as Entregue," & vbCrLf & _
              "        SUM(sb.cedente)         as Cedente, " & vbCrLf & _
              "        SUM(sb.cessionario)     as Cessionario, " & vbCrLf & _
              "        SUM(convert(numeric(18,4), isnull(ls.LaudosSemNota,0)))  as LaudosSemNota," & vbCrLf & _
              "        SUM(sb.Contratado + sb.ContratadoTroca - sb.Entregue  + sb.cedente - sb.cessionario) as Saldo," & vbCrLf & _
              "        0.00 AS AdiantamentoEmReais," & vbCrLf & _
              "        0.00 AS AdiantamentoEmDolares," & vbCrLf & _
              "        0.00 as Reais," & vbCrLf & _
              "        0.00 as Dolares" & vbCrLf & _
              "   INTO #SaldoFinal " & vbCrLf & _
              "   from (   " & vbCrLf & _
              "	        select b.empresa," & vbCrLf & _
              "	               b.cliente, " & vbCrLf & _
              IIf(Not chkClienteConsolidado.Checked, " b.Endcliente,", "") & vbCrLf & _
              "			  (select top 1 nome from clientes where cliente_id = b.cliente " & IIf(Not chkClienteConsolidado.Checked, " AND Endereco_Id = B.EndCliente ", " ") & ") as Nome," & vbCrLf & _
              "                b.DataEntrega," & vbCrLf & _
              "                case" & vbCrLf & _
              "                  when isnull(cxt.Tipo_id,0) = 20" & vbCrLf & _
              "                    then embarque.complemento + ' - ' + embarque.cidade + ' - ' + embarque.estado" & vbCrLf & _
              "                    else embarque.nome + ' - ' + embarque.cidade + ' - ' + embarque.estado" & vbCrLf & _
              "                end as Embarque," & vbCrLf & _
              "	               b.Produto," & vbCrLf & _
              "	               b.Nomeproduto," & vbCrLf & _
              "	               sum(case " & vbCrLf & _
              "	                     when isnull(p.troca,0) = 0 and so.precofixo = 'S' and so.classe <> '" & eClassesOperacoes.DEPOSITOS.ToString & "'" & vbCrLf & _
              "	               	       then Contratado" & vbCrLf & _
              "	               	       else 0" & vbCrLf & _
              "	                   end)" & vbCrLf & _
              "	               + sum(case " & vbCrLf & _
              "	                       when isnull(p.troca,0) = 0 and so.classe = '" & eClassesOperacoes.AFIXAR.ToString & "'" & vbCrLf & _
              "	               	         then fixada" & vbCrLf & _
              "	               	         else 0" & vbCrLf & _
              "	                     end) as contratado," & vbCrLf & _
              "	               sum(case " & vbCrLf & _
              "	               	     when isnull(p.troca,0) = 1 and so.precofixo = 'S' and so.classe <> '" & eClassesOperacoes.DEPOSITOS.ToString & "'" & vbCrLf & _
              "	               	       then Contratado" & vbCrLf & _
              "	               	       else 0" & vbCrLf & _
              "	                   end) " & vbCrLf & _
              "	               + sum(case " & vbCrLf & _
              "	               	       when isnull(p.troca,0) = 1 and so.classe = '" & eClassesOperacoes.AFIXAR.ToString & "'" & vbCrLf & _
              "	               	         then fixada" & vbCrLf & _
              "	               	         else 0" & vbCrLf & _
              "	                     end) as ContratadoTroca,		           " & vbCrLf & _
              "	               sum(b.laudo) as Entregue,                 " & vbCrLf & _
              "	               sum(case " & vbCrLf & _
              "	                     when so.precofixo = 'N' and b.moeda = 1" & vbCrLf & _
              "	               	       then b.adiantamento" & vbCrLf & _
              "	               	       else 0" & vbCrLf & _
              "	                   end) as AdiantamentoReais,		           " & vbCrLf & _
              "	               sum(case " & vbCrLf & _
              "	               	     when so.precofixo = 'N' and b.moeda = 3" & vbCrLf & _
              "	               	       then b.adiantamento" & vbCrLf & _
              "	               	       else 0" & vbCrLf & _
              "	                   end) as AdiantamentoDolar," & vbCrLf & _
              "                sum(Cedente) as cedente," & vbCrLf & _
              "                sum(Cessionario) as Cessionario	" & vbCrLf & _
              "	          from #base b" & vbCrLf & _
              "	         inner join Pedidos P" & vbCrLf & _
              "	           on b.empresa    = P.Empresa_id" & vbCrLf & _
              "	          and b.endEmpresa = P.EndEmpresa_id" & vbCrLf & _
              "	          and b.Pedido     = P.Pedido_id" & vbCrLf & _
              "	         Left Join(SELECT PxD.Empresa_Id,  PxD.EndEmpresa_Id, PxD.Pedido_Id," & vbCrLf & _
              "	                          PxD.Deposito_Id, PxD.EndDeposito_Id" & vbCrLf & _
              "                      FROM PedidosXDepositos PxD" & vbCrLf & _
              "                     where PxD.Tipo = 'OD'" & vbCrLf & _
              "                       and PxD.Principal = 1" & vbCrLf & _
              "	                   ) DE" & vbCrLf & _
              "	           on P.Empresa_id    = DE.Empresa_id" & vbCrLf & _
              "	          and P.EndEmpresa_id = DE.EndEmpresa_id" & vbCrLf & _
              "	          and P.Pedido_id     = DE.Pedido_id" & vbCrLf & _
              "	         Left Join(SELECT PxD.Empresa_Id,  PxD.EndEmpresa_Id, PxD.Pedido_Id, " & vbCrLf & _
              "	                          PxD.Deposito_Id, PxD.EndDeposito_Id" & vbCrLf & _
              "                           FROM PedidosXDepositos PxD" & vbCrLf & _
              "                          where PxD.Tipo = 'LE'" & vbCrLf & _
              "                            and PxD.Principal = 1" & vbCrLf & _
              "	                   ) LE" & vbCrLf & _
              "	           on P.Empresa_id    = LE.Empresa_id" & vbCrLf & _
              "	          and P.EndEmpresa_id = LE.EndEmpresa_id" & vbCrLf & _
              "	          and P.Pedido_id     = LE.Pedido_id" & vbCrLf & _
              "	        inner join Clientes Embarque" & vbCrLf & _
              "	           on Embarque.Cliente_Id  = isnull(LE.Deposito_Id   ,DE.Deposito_Id)" & vbCrLf & _
              "	          and Embarque.Endereco_Id = isnull(LE.EndDeposito_Id,DE.EndDeposito_Id)" & vbCrLf & _
              "          left Join clientesxtipos cxt" & vbCrLf & _
              "            on Embarque.Cliente_Id  = cxt.Cliente_Id" & vbCrLf & _
              "           and Embarque.Endereco_Id = cxt.Endereco_Id" & vbCrLf & _
              "           and cxt.Tipo_id = 20" & vbCrLf & _
              "	        inner join Suboperacoes SO" & vbCrLf & _
              "	           on b.Operacao    = SO.Operacao_id" & vbCrLf & _
              "	          and b.Suboperacao = SO.SubOperacoes_Id " & vbCrLf & _
              "	        inner join Operacoes OP" & vbCrLf & _
              "	           on OP.operacao_id = SO.Operacao_Id  " & vbCrLf & _
              "	        WHERE 1=1" & vbCrLf

            If temselecionados Then
                Sql &= "  AND b.Empresa + cast(b.EndEmpresa as varchar) in ('" & empresasselecionadas & "')" & vbCrLf
            End If

            If Not String.IsNullOrWhiteSpace(txtCliente.Text) Then
                If chkClienteConsolidado.Checked Then
                    Sql &= "And left(b.Cliente, 8) = '" & Left(txtCodigoCliente.Value.Split("-")(0), 8) & "'"
                Else
                    Sql &= "And b.Cliente    = '" & txtCodigoCliente.Value.Split("-")(0) & "'" & vbCrLf & _
                           "And b.EndCliente = " & txtCodigoCliente.Value.Split("-")(1) & vbCrLf
                End If
            End If

            Sql &= "    And OP.classe in ('" & eClassesOperacoes.COMPRAS.ToString & "','" & eClassesOperacoes.DEPOSITOS.ToString & "') " & vbCrLf & _
                   "  group by b.empresa,b.cliente, b.Produto, b.Nomeproduto, " & IIf(Not chkClienteConsolidado.Checked, " b.EndCliente, ", "") & vbCrLf & _
                   "           b.DataEntrega, " & vbCrLf & _
                   "           case" & vbCrLf & _
                   "             when isnull(cxt.Tipo_id,0) = 20" & vbCrLf & _
                   "               then embarque.complemento + ' - ' + embarque.cidade + ' - ' + embarque.estado" & vbCrLf & _
                   "               else embarque.nome + ' - ' + embarque.cidade + ' - ' + embarque.estado" & vbCrLf & _
                   "           end" & vbCrLf & _
                   "  ) as sb  " & vbCrLf & _
                   " left join (" & vbCrLf & _
                   "				SELECT R.Empresa_Id, R.EndEmpresa_Id, Ped.cliente, R.Produto, sum(R.PesoLiquido) as LaudosSemNota" & IIf(Not chkClienteConsolidado.Checked, ", Ped.Endcliente", "") & vbCrLf & _
                   "				  FROM Romaneios R " & vbCrLf & _
                   "				 INNER JOIN RomaneiosXPesagens RxP " & vbCrLf & _
                   "					ON R.Empresa_Id       = RxP.Empresa_Id " & vbCrLf & _
                   "				   AND R.EndEmpresa_Id    = RxP.EndEmpresa_Id" & vbCrLf & _
                   "				   AND R.Romaneio_Id      = RxP.Romaneio_Id" & vbCrLf & _
                   "                 INNER JOIN Pesagem P" & vbCrLf & _
                   "                    ON RxP.Empresa_Id    = P.Empresa_Id" & vbCrLf & _
                   "                   AND RxP.EndEmpresa_Id = P.EndEmpresa_Id" & vbCrLf & _
                   "                   AND RxP.Pesagem_Id    = P.Pesagem_Id" & vbCrLf & _
                   "                   AND P.Sequencia_Id             = 0" & vbCrLf & _
                   "                   AND P.Situacao                 = 1" & vbCrLf & _
                   "				  LEFT JOIN Pedidos Ped" & vbCrLf & _
                   "					ON R.Empresa_Id    = Ped.Empresa_Id" & vbCrLf & _
                   "				   AND R.EndEmpresa_Id = Ped.EndEmpresa_Id" & vbCrLf & _
                   "				   AND R.Pedido        = Ped.Pedido_Id" & vbCrLf & _
                   "				  LEFT OUTER JOIN Clientes" & vbCrLf & _
                   "					ON Ped.Cliente    = Clientes.Cliente_Id" & vbCrLf & _
                   "				   AND Ped.EndCliente = Clientes.Endereco_Id" & vbCrLf & _
                   "				  LEFT OUTER JOIN NotasFiscaisXRomaneios nfxr" & vbCrLf & _
                   "					ON R.Empresa_Id    = nfxr.Empresa_Id" & vbCrLf & _
                   "				   AND R.EndEmpresa_Id = nfxr.EndEmpresa_Id" & vbCrLf & _
                   "				   AND Ped.Cliente         = nfxr.Cliente_id" & vbCrLf & _
                   "				   AND Ped.EndCliente      = nfxr.EndCliente_id" & vbCrLf & _
                   "				   AND R.Romaneio_Id   = nfxr.Romaneio_Id " & vbCrLf & _
                   "				  LEFT JOIN NotasFiscais NF" & vbCrLf & _
                   "					ON NF.Empresa_Id      = nfxr.Empresa_Id" & vbCrLf & _
                   "				   AND NF.EndEmpresa_Id   = nfxr.EndEmpresa_Id" & vbCrLf & _
                   "				   AND NF.Cliente_Id      = nfxr.Cliente_Id " & vbCrLf & _
                   "				   AND NF.EndCliente_Id   = nfxr.EndCliente_Id" & vbCrLf & _
                   "				   AND NF.EntradaSaida_Id = nfxr.EntradaSaida_Id" & vbCrLf & _
                   "				   AND NF.Serie_Id        = nfxr.Serie_Id " & vbCrLf & _
                   "				   AND NF.Nota_Id         = nfxr.Nota_Id  " & vbCrLf & _
                   "				   --AND NF.Movimento       <=  @Vencimento				  " & vbCrLf & _
                   "				 Where NF.Nota_Id   IS NULL" & vbCrLf & _
                   "				   --AND R.Movimento       >=  @InicioDeSafra" & vbCrLf & _
                   "				   --AND R.Movimento       <=  @Vencimento	" & vbCrLf & _
                   "				   and R.EntradaSaida = 'E'" & vbCrLf & _
                   "                   and Ped.safra = '" & ddlSafra.SelectedValue & "'" & vbCrLf & _
                   "				 group by R.Empresa_Id, R.EndEmpresa_Id, Ped.cliente,R.Produto" & IIf(Not chkClienteConsolidado.Checked, ", Ped.Endcliente", "") & vbCrLf & _
                   "	          ) LS" & vbCrLf & _
                   "	    on sb.empresa = LS.Empresa_Id" & vbCrLf & _
                   "	   and sb.cliente = LS.cliente  " & vbCrLf & _
                    IIf(Not chkClienteConsolidado.Checked, "	   and sb.Endcliente = LS.Endcliente", "") & vbCrLf & _
                   "	   and sb.produto = LS.Produto     " & vbCrLf & _
                   " GROUP BY sb.Cliente, sb.nome, sb.Embarque, sb.DataEntrega " & vbCrLf

            Sql &= " SELECT * FROM #SaldoFinal where saldo + LaudosSemNota > 0 order by NomeDoCliente, dataentrega" & vbCrLf
        End If

        Return Sql
    End Function

    Private Sub CarregarEmpresa1()
        Dim sql As String = " SELECT DISTINCT Clientes.Cliente_Id as Codigo, Clientes.Endereco_Id, Clientes.Reduzido, Clientes.Nome, Clientes.Cidade, Clientes.Estado " & vbCrLf & _
                            "   FROM GruposXEmpresas " & vbCrLf & _
                            "  INNER JOIN Clientes" & vbCrLf & _
                            "     ON GruposXEmpresas.Cliente_Id    = Clientes.Cliente_Id " & vbCrLf & _
                            "    AND GruposXEmpresas.EndCliente_Id = Clientes.Endereco_Id" & vbCrLf & _
                            "  Inner Join ClientesXEmpresas cxe" & vbCrLf & _
                            "     on cxe.Empresa_Id    = Clientes.Cliente_Id" & vbCrLf & _
                            "    and cxe.EndEmpresa_Id = Clientes.Endereco_Id" & vbCrLf

        Dim ds As DataSet = Banco.ConsultaDataSet(sql, "Empresas")

        lstEmpresa1.Items.Clear()

        If ds IsNot Nothing AndAlso ds.Tables IsNot Nothing AndAlso ds.Tables.Count > 0 AndAlso ds.Tables(0).Rows.Count > 0 Then
            lstEmpresa1.Items.Clear()
            For Each row As DataRow In ds.Tables(0).Rows
                lstEmpresa1.Items.Add(New ListItem(Funcoes.AlinharEsquerda(row("Nome"), 28, ".") & " - " & Funcoes.AlinharEsquerda(row("Cidade"), 20, ".") & " " & row("Estado") & " " & Funcoes.AlinharEsquerda(Funcoes.FormatarCpfCnpj(row("Codigo")), 18, ".") & "-" & CStr(row("Endereco_Id")) & "-" & row("Reduzido"), row("Codigo") & row("Endereco_Id")))
            Next
        End If
    End Sub

    Private Sub EmitirRelatorio(ByVal Pdf As Boolean)
        If ValidarCampos() Then
            Dim dsPosicoes As New DataSet
            Dim strSQL As String = " SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED" & vbCrLf
            strSQL &= SqlPosicao()
            strSQL &= ComporRelatorio(chkPedidosComSaldo.Checked)

            dsPosicoes.Merge(Banco.ConsultaDataSet(strSQL, "PosicaoDePedidosConsolidada"))

            Dim rpt As String = ""

            If ddlFretes.SelectedItem.Text <> "FOB" Then
                If Pdf Then
                    rpt = "Cr_PosicaoDePedidosConsolidada"
                Else
                    rpt = "Cr_PosicaoDePedidosConsolidadaExcell2"
                End If
            Else
                rpt = "Cr_PosicaoDePedidosConsolidadaFOB"
            End If

            Dim parameters = New Dictionary(Of String, Object)()

            Dim Parametros As String = String.Empty
            If lstEmpresa1.GetSelectedValues().Count > 0 Then Parametros &= "Empresa: " & String.Join("-", lstEmpresa1.GetSelecteds()) & vbCrLf
            Dim par As ArrayList = ucSelecaoProduto.GetSqlEParametrosRelatorio("s", "s", "", True)
            Parametros &= par(1)
            If ddlSafra.SelectedIndex > 0 Then Parametros &= "Safra: " & ddlSafra.SelectedItem.Text & vbCrLf
            Parametros &= "Data Limite do Financeiro:" + txtData.Text

            parameters.Add("Parametros", Parametros)
            parameters.Add("NomeEmpresa", HttpContext.Current.Session("ssNomeEmpresa"))
            parameters.Add("Cidade", HttpContext.Current.Session("ssCidadeEmpresa") & " - " & HttpContext.Current.Session("ssEstadoEmpresa"))

            Funcoes.BindReport(Me.Page, dsPosicoes, rpt, IIf(Pdf, eExportType.PDF, eExportType.ExcelCrystal), parameters)
        End If
    End Sub

    Public Function SqlPosicao() As String
        Dim Cliente As Cliente = Nothing
        Dim par As ArrayList = ucSelecaoProduto.GetSqlEParametrosRelatorio("Prd.Grupo", "Prd.Produto_Id")

        If Not String.IsNullOrWhiteSpace(txtCliente.Text) Then
            Cliente = New Cliente(txtCodigoCliente.Value.Split("-")(0), txtCodigoCliente.Value.Split("-")(1))
        End If

        Dim pos As New ListPosicaoDePedido
        Return pos.SqlPosicao(Date.Now, lstEmpresa1.GetSelecteds(), False, Cliente, chkClienteConsolidado.Checked, "#Base", "", ddlSafra.SelectedValue, ddlSafra.SelectedValue, par(0), "", "", "", "", "", ddlFretes.SelectedIndex, "", "", 0, 0, 0, 0, False)
    End Function

    Public Overrides Sub Carregar(ByVal obj As [Lib].Negocio.IBaseEntity)
        Try
            If Not Session("objCliente" & HID.Value) Is Nothing Then
                Dim itemCliente As ListItem = Funcoes.FormatarListItemCliente(CType(Session("objCliente" & HID.Value), [Lib].Negocio.Cliente))
                txtCliente.Text = itemCliente.Text
                txtCodigoCliente.Value = itemCliente.Value
                Session.Remove("objCliente" & HID.Value)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

#End Region

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "PosicaoDePedido")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub

    Protected Sub ddlSafra_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddlSafra.SelectedIndexChanged
        Dim sf As New Safra(ddlSafra.SelectedValue)
        txtData.Text = IIf(sf.Vencimento.Ticks = 0, String.Empty, sf.Vencimento.ToString("dd/MM/yyyy"))
    End Sub
End Class