Imports System.Data
Imports System.IO
Imports Microsoft.VisualBasic
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis
Imports System.Drawing
Imports OfficeOpenXml.Style
Imports OfficeOpenXml

Partial Class PosicaoFinanceiraDeVendas
    Inherits BasePage

    Private Sql As String
    Private ds As DataSet

#Region "Events"

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not IsPostBack And IsConnect Then
            Me.setMenu(eModulo.Comercial)
            If Funcoes.VerificaPermissao("PosicaoFinanceiraDeVendas", "ACESSAR") Then
                Limpar()
            Else
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Comercial.aspx")
                Exit Sub
            End If
        End If
    End Sub

    Protected Sub btnCliente_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCliente.Click
        Try
            ucConsultaClientes.Limpar()
            Popup.ConsultaDeClientes(Me.Page, "objCliente" & HID.Value, "txtNome")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub


    Protected Sub lnkExcel_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkExcel.Click
        Try
            If Funcoes.VerificaPermissao("AlterarChaveNFE", "RELATORIO") Then
                EmitirExcel()
            Else
                MsgBox(Me.Page, "Usuário sem permissão para alterar registro(s).")
                Exit Sub
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkLimpar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkLimpar.Click
        Try
            Limpar()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "PosicaoDePedido")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub

#End Region

#Region "Methods"

    Private Sub Limpar()
        ddl.Carregar(ddlEmpresa, CarregarDDL.Tabela.Empresas, "", True)
        Funcoes.VerificaEmpresa(ddlEmpresa)

        HID.Value = Guid.NewGuid.ToString

        txtData.Text = Format(Today, "dd/MM/yyyy")
        txtCodigoCliente.Value = 0
        txtCliente.Text = String.Empty
        ddlEmpresa.Enabled = False

        chkConsolidarEmpresa.Checked = True
        chkClienteConsolidado.Checked = True
        chkPedidosComSaldo.Checked = True
    End Sub

    Private Function ValidarCampos() As Boolean
        If String.IsNullOrWhiteSpace(txtData.Text) Or Not IsDate(txtData.Text) Then
            MsgBox(Me.Page, "Informe uma data válida.")
            txtData.Focus()
            Return False
        End If
        Return True
    End Function

    Private Function ComporRelatorio(Saldo As Boolean) As String
        'Dim Sql As String = String.Empty
        'Dim temselecionados As Boolean = lstEmpresa1.GetSelectedValues.Count > 0
        'Dim empresasselecionadas As String = String.Join("','", lstEmpresa1.GetSelecteds())

        'If ddlFretes.SelectedItem.Text <> "FOB" Then
        '    Sql &= " /*-------------------------------------------------*/" & vbCrLf &
        '      " DECLARE @InicioDeSafra AS Date, @Vencimento AS Date" & vbCrLf &
        '      "  SELECT @InicioDeSafra = S.InicioDeSafra, @Vencimento = S.Vencimento " & vbCrLf &
        '      "    FROM Safras S" & vbCrLf &
        '      "    JOIN #Safra S2" & vbCrLf &
        '      "      ON S.Safra_Id = S2.Safra_Id" & vbCrLf &
        '      " /*-------------------------------------------------*/" & vbCrLf &
        '      " select sb.Cliente AS CodigoDoCliente," & vbCrLf &
        '      "        sb.Nome AS NomeDoCliente," & vbCrLf &
        '      "        SUM(sb.Contratado)      AS Contratado," & vbCrLf &
        '      "        SUM(sb.ContratadoTroca) AS ContratadoPorTroca," & vbCrLf &
        '      "        SUM(sb.Entregue)        AS Entregue," & vbCrLf &
        '      "        SUM(sb.Cedente)         AS Cedente," & vbCrLf &
        '      "        SUM(sb.Cessionario)     AS Cessionario," & vbCrLf &
        '      "        SUM(convert(numeric(18,4), isnull(ls.LaudosSemNota,0))) as LaudosSemNota," & vbCrLf &
        '      "        SUM(sb.Contratado + sb.ContratadoTroca - (sb.Entregue) + sb.cedente - sb.cessionario) as Saldo, --+ convert(numeric(18,4),isnull(ls.LaudosSemNota,0))" & vbCrLf &
        '      "        SUM(sb.AdiantamentoReais) AS AdiantamentoEmReais," & vbCrLf &
        '      "        SUM(sb.AdiantamentoDolar) AS AdiantamentoEmDolares," & vbCrLf &
        '      "        (isnull(fin.reais,0))  AS Reais," & vbCrLf &
        '      "        (isnull(fin.dolar,0))  AS Dolares" & vbCrLf &
        '      "   INTO #SaldoFinal " & vbCrLf &
        '      "   from (   " & vbCrLf &
        '      "	   select b.empresa," & vbCrLf &
        '      "	          b.cliente, " & vbCrLf &
        '      IIf(Not chkClienteConsolidado.Checked, "	          b.EndCliente,", "") & vbCrLf &
        '      "			  (select top 1 nome from clientes where cliente_id = b.cliente " & IIf(Not chkClienteConsolidado.Checked, " AND Endereco_Id = B.EndCliente ", " ") & ") as Nome," & vbCrLf &
        '      "			  b.Produto," & vbCrLf &
        '      "			  b.NomeProduto," & vbCrLf &
        '      "			  sum(case " & vbCrLf &
        '      "					 when isnull(p.troca,0) = 0 and so.precofixo = 'S' and so.classe <> '" & eClassesOperacoes.DEPOSITOS.ToString & "'" & vbCrLf &
        '      "					   then Contratado" & vbCrLf &
        '      "					   else 0" & vbCrLf &
        '      "				   end)" & vbCrLf &
        '      "				+   " & vbCrLf &
        '      "			  sum(case " & vbCrLf &
        '      "					 when isnull(p.troca,0) = 0 and so.classe = '" & eClassesOperacoes.AFIXAR.ToString & "'" & vbCrLf &
        '      "					   then fixada" & vbCrLf &
        '      "					   else 0" & vbCrLf &
        '      "				   end) as contratado," & vbCrLf &
        '      "			   sum(case " & vbCrLf &
        '      "					 when isnull(p.troca,0) = 1 and so.precofixo = 'S' and so.classe <> '" & eClassesOperacoes.DEPOSITOS.ToString & "'" & vbCrLf &
        '      "					   then Contratado" & vbCrLf &
        '      "					   else 0" & vbCrLf &
        '      "				   end) " & vbCrLf &
        '      "			   +" & vbCrLf &
        '      "			   sum(case " & vbCrLf &
        '      "					 when isnull(p.troca,0) = 1 and so.classe = '" & eClassesOperacoes.AFIXAR.ToString & "'" & vbCrLf &
        '      "					   then fixada" & vbCrLf &
        '      "					   else 0" & vbCrLf &
        '      "				   end) 	   				   " & vbCrLf &
        '      "				   as ContratadoTroca,		           " & vbCrLf &
        '      "			   sum(b.laudo) as Entregue,                 " & vbCrLf &
        '      "			   sum(case " & vbCrLf &
        '      "					 when so.precofixo = 'N' and b.moeda = 1" & vbCrLf &
        '      "					   then b.adiantamento" & vbCrLf &
        '      "					   else 0" & vbCrLf &
        '      "				   end) as AdiantamentoReais,		           " & vbCrLf &
        '      "				sum(case " & vbCrLf &
        '      "					 when so.precofixo = 'N' and b.moeda = 3" & vbCrLf &
        '      "					   then b.adiantamento" & vbCrLf &
        '      "					   else 0" & vbCrLf &
        '      "				   end) as AdiantamentoDolar," & vbCrLf &
        '      "                sum(Cedente) as cedente," & vbCrLf &
        '      "                sum(Cessionario) as Cessionario	" & vbCrLf &
        '      "		   from #base b" & vbCrLf &
        '      "		  inner join Pedidos P" & vbCrLf &
        '      "		     on b.empresa    = P.Empresa_id" & vbCrLf &
        '      "		    and b.endEmpresa = P.EndEmpresa_id" & vbCrLf &
        '      "		    and b.Pedido     = P.Pedido_id" & vbCrLf &
        '      "		   Left Join(SELECT PxD.Empresa_Id,  PxD.EndEmpresa_Id, PxD.Pedido_Id," & vbCrLf &
        '      "		                    PxD.Deposito_Id, PxD.EndDeposito_Id" & vbCrLf &
        '      "                       FROM PedidosXDepositos PxD" & vbCrLf &
        '      "                      where PxD.Tipo = 'OD'" & vbCrLf &
        '      "                        and PxD.Principal = 1" & vbCrLf &
        '      "		             ) DE" & vbCrLf &
        '      "		     on P.Empresa_id    = DE.Empresa_id" & vbCrLf &
        '      "		    and P.EndEmpresa_id = DE.EndEmpresa_id" & vbCrLf &
        '      "		    and P.Pedido_id     = DE.Pedido_id" & vbCrLf &
        '      "		   Left Join(SELECT PxD.Empresa_Id,  PxD.EndEmpresa_Id, PxD.Pedido_Id," & vbCrLf &
        '      "		                    PxD.Deposito_Id, PxD.EndDeposito_Id" & vbCrLf &
        '      "                       FROM PedidosXDepositos PxD" & vbCrLf &
        '      "                      where PxD.Tipo = 'LE'" & vbCrLf &
        '      "                        and PxD.Principal = 1" & vbCrLf &
        '      "		             ) LE" & vbCrLf &
        '      "		     on P.Empresa_id    = LE.Empresa_id" & vbCrLf &
        '      "		    and P.EndEmpresa_id = LE.EndEmpresa_id" & vbCrLf &
        '      "		    and P.Pedido_id     = LE.Pedido_id" & vbCrLf &
        '      "		  inner join Clientes Embarque" & vbCrLf &
        '      "		     on Embarque.Cliente_Id  = isnull(LE.Deposito_Id   ,DE.Deposito_Id)" & vbCrLf &
        '      "		    and Embarque.Endereco_Id = isnull(LE.EndDeposito_Id,DE.EndDeposito_Id)" & vbCrLf &
        '      "		  inner join Suboperacoes SO" & vbCrLf &
        '      "		     on b.Operacao    = SO.Operacao_id" & vbCrLf &
        '      "		    and b.Suboperacao = SO.SubOperacoes_Id " & vbCrLf &
        '      "		  inner join Operacoes OP" & vbCrLf &
        '      "		     on OP.operacao_id = SO.Operacao_Id  " & vbCrLf &
        '      "		  WHERE 1=1" & vbCrLf

        '    If temselecionados Then
        '        Sql &= "  AND b.Empresa + cast(b.EndEmpresa as varchar) in ('" & empresasselecionadas & "')" & vbCrLf
        '    End If

        '    If Not String.IsNullOrWhiteSpace(txtCliente.Text) Then
        '        If chkClienteConsolidado.Checked Then
        '            Sql &= "And left(b.Cliente, 8) = '" & Left(txtCodigoCliente.Value.Split("-")(0), 8) & "'"
        '        Else
        '            Sql &= "And b.Cliente    = '" & txtCodigoCliente.Value.Split("-")(0) & "'" & vbCrLf &
        '                   "And b.EndCliente = " & txtCodigoCliente.Value.Split("-")(1) & vbCrLf
        '        End If
        '    End If

        '    Sql &= "  And OP.classe in ('" & eClassesOperacoes.COMPRAS.ToString & "','" & eClassesOperacoes.DEPOSITOS.ToString & "') " & vbCrLf &
        '    "		  group by b.empresa,b.cliente, b.Produto,b.Nomeproduto" & IIf(Not chkClienteConsolidado.Checked, ", b.EndCliente", "") & vbCrLf &
        '    "		) as sb  " & vbCrLf &
        '    "  left join " & vbCrLf

        '    If FinanceiroNovo Then
        '        Sql &= "               (select clifor as cliente," & vbCrLf &
        '               "                       sum(case" & vbCrLf &
        '               "               		         when T.Moeda = 1" & vbCrLf &
        '               "               			       then ValorOficial" & vbCrLf &
        '               "               			       else 0" & vbCrLf &
        '               "               		       end) reais,    " & vbCrLf &
        '               "                       sum(case" & vbCrLf &
        '               "               			     when T.Moeda = 3" & vbCrLf &
        '               "               			       then ValorMoeda" & vbCrLf &
        '               "               			       else 0" & vbCrLf &
        '               "               		       end) dolar    " & vbCrLf &
        '               "                  from Titulos T" & vbCrLf &
        '               "                  JOIN TitulosxContaContabil Tc" & vbCrLf &
        '               "                    ON T.Titulo_Id = Tc.Titulo_Id" & vbCrLf &
        '               "                   AND T.ContaContabilCliFor = Tc.Conta_Id " & vbCrLf &
        '               "                  left join Pedidos p " & vbCrLf &
        '               "                    on t.EmpresaPedido    = p.Empresa_id" & vbCrLf &
        '               "                   and t.EndEmpresaPedido = p.EndEmpresa_id" & vbCrLf &
        '               "                   and t.Pedido           = p.Pedido_id" & vbCrLf &
        '               "                 where T.situacao = 1" & vbCrLf &
        '               "                   and provisao <> 1 " & vbCrLf &
        '               "                   and (T.reprogramacao < @Vencimento or cr.prorrogacao < '" & CDate(txtData.Text).ToString("yyyy-MM-dd") & "' or isnull(p.safra,'') = '" & ddlSafra.SelectedValue & "')" & vbCrLf

        '        If temselecionados Then
        '            Sql &= "                   and Empresa + cast(EndEmpresa as varchar) in ('" & empresasselecionadas & "') " & vbCrLf
        '        End If

        '        Sql &= "                   AND T.RecPag = 'R'" & vbCrLf &
        '               "                 group by clifor" & vbCrLf &
        '               "                ) Fin" & vbCrLf
        '    Else
        '        Sql &= "              (select cr.cliente," & vbCrLf &
        '               "	                  sum(case" & vbCrLf &
        '               "							 when cr.moeda = 1" & vbCrLf &
        '               "							   then cr.valordodocumento" & vbCrLf &
        '               "							   else 0" & vbCrLf &
        '               "						   end) reais,    " & vbCrLf &
        '               "					   sum(case" & vbCrLf &
        '               "							 when cr.moeda = 3" & vbCrLf &
        '               "							   then cr.Moedavalordodocumento" & vbCrLf &
        '               "							   else 0" & vbCrLf &
        '               "						   end) dolar    " & vbCrLf &
        '               "				 from contasareceber cr" & vbCrLf &
        '               "				 left join Pedidos p" & vbCrLf &
        '               "                   on cr.EmpresaPedido    = p.Empresa_id" & vbCrLf &
        '               "                  and cr.EndEmpresaPedido = p.EndEmpresa_id" & vbCrLf &
        '               "                  and cr.Pedido           = p.Pedido_id" & vbCrLf &
        '               "				where cr.situacao = 1 " & vbCrLf &
        '               "				  and cr.provisao <> 1 " & vbCrLf &
        '               "				  and (cr.prorrogacao < @Vencimento or cr.prorrogacao < '" & CDate(txtData.Text).ToString("yyyy-MM-dd") & "' or isnull(p.safra,'') = '" & ddlSafra.SelectedValue & "')" & vbCrLf

        '        If temselecionados Then
        '            Sql &= "				  And cr.Empresa + cast(cr.EndEmpresa as varchar) in ('" & empresasselecionadas & "')" & vbCrLf
        '        End If

        '        Sql &= "				group by cr.cliente  " & vbCrLf &
        '               "			   ) Fin" & vbCrLf
        '    End If

        '    Sql &= "	  on sb.cliente = fin.cliente" & vbCrLf &
        '           "  left join (" & vbCrLf &
        '           "				SELECT R.Empresa_Id, R.EndEmpresa_Id, Ped.cliente, R.Produto, sum(R.PesoLiquido) as LaudosSemNota" & IIf(Not chkClienteConsolidado.Checked, ", Ped.Endcliente", "") & vbCrLf &
        '           "				  FROM Romaneios R" & vbCrLf &
        '           "				 INNER JOIN RomaneiosXPesagens RxP " & vbCrLf &
        '           "					ON R.Empresa_Id       = RxP.Empresa_Id " & vbCrLf &
        '           "				   AND R.EndEmpresa_Id    = RxP.EndEmpresa_Id" & vbCrLf &
        '           "				   AND R.Romaneio_Id      = RxP.Romaneio_Id" & vbCrLf &
        '           "				  LEFT JOIN Pedidos Ped" & vbCrLf &
        '           "					ON R.Empresa_Id    = Ped.Empresa_Id" & vbCrLf &
        '           "				   AND R.EndEmpresa_Id = Ped.EndEmpresa_Id" & vbCrLf &
        '           "				   AND R.Pedido        = Ped.Pedido_Id" & vbCrLf &
        '           "				  LEFT OUTER JOIN Clientes" & vbCrLf &
        '           "					ON Ped.Cliente    = Clientes.Cliente_Id" & vbCrLf &
        '           "				   AND Ped.EndCliente = Clientes.Endereco_Id" & vbCrLf &
        '           "				  LEFT OUTER JOIN NotasFiscaisXRomaneios nfxr" & vbCrLf &
        '           "					ON R.Empresa_Id    = nfxr.Empresa_Id" & vbCrLf &
        '           "				   AND R.EndEmpresa_Id = nfxr.EndEmpresa_Id" & vbCrLf &
        '           "				   AND R.Romaneio_Id   = nfxr.Romaneio_Id " & vbCrLf &
        '           "				   --AND Ped.Cliente         = nfxr.Cliente_id" & vbCrLf &
        '           "				   --AND Ped.EndCliente      = nfxr.EndCliente_id" & vbCrLf &
        '           "                 INNER JOIN Pesagem P" & vbCrLf &
        '           "                    ON RxP.Empresa_Id    = P.Empresa_Id" & vbCrLf &
        '           "                   AND RxP.EndEmpresa_Id = P.EndEmpresa_Id" & vbCrLf &
        '           "                   AND RxP.Pesagem_Id    = P.Pesagem_Id" & vbCrLf &
        '           "                   AND P.Sequencia_Id             = 0" & vbCrLf &
        '           "                   AND P.Situacao                 = 1" & vbCrLf &
        '           "				  LEFT JOIN NotasFiscais NF" & vbCrLf &
        '           "					ON NF.Empresa_Id      = nfxr.Empresa_Id" & vbCrLf &
        '           "				   AND NF.EndEmpresa_Id   = nfxr.EndEmpresa_Id" & vbCrLf &
        '           "				   AND NF.Cliente_Id      = nfxr.Cliente_Id " & vbCrLf &
        '           "				   AND NF.EndCliente_Id   = nfxr.EndCliente_Id" & vbCrLf &
        '           "				   AND NF.EntradaSaida_Id = nfxr.EntradaSaida_Id" & vbCrLf &
        '           "				   AND NF.Serie_Id        = nfxr.Serie_Id " & vbCrLf &
        '           "				   AND NF.Nota_Id         = nfxr.Nota_Id  " & vbCrLf &
        '           "				   --AND NF.Movimento       <=  @Vencimento" & vbCrLf &
        '           "				 Where NF.Nota_Id   IS NULL" & vbCrLf &
        '           "				   --AND R.Movimento       >=  @InicioDeSafra" & vbCrLf &
        '           "				   --AND R.Movimento       <=  @Vencimento" & vbCrLf &
        '           "				   and R.EntradaSaida = 'E'" & vbCrLf &
        '           "                   and Ped.safra = '" & ddlSafra.SelectedValue & "'" & vbCrLf &
        '           "				 group by R.Empresa_Id, R.EndEmpresa_Id, Ped.cliente,R.Produto" & IIf(Not chkClienteConsolidado.Checked, ", Ped.Endcliente", "") & vbCrLf &
        '           "	          ) LS" & vbCrLf &
        '           "	    on sb.empresa = LS.Empresa_Id" & vbCrLf &
        '           "	   and sb.cliente = LS.cliente" & vbCrLf &
        '         IIf(Not chkClienteConsolidado.Checked, "	   and sb.EndCliente = LS.EndCliente", "") & vbCrLf &
        '           "	   and sb.produto = LS.Produto" & vbCrLf &
        '           "     Group by sb.Cliente," & vbCrLf &
        '           "              sb.Nome," & vbCrLf &
        '           "              isnull(fin.reais,0)," & vbCrLf &
        '           "              isnull(fin.dolar,0)" & vbCrLf &
        '           "      order by sb.nome  " & vbCrLf &
        '           "   /*---------------------------------*/  " & vbCrLf

        '    If Saldo Then
        '        Sql &= " SELECT * FROM #SaldoFinal " & vbCrLf &
        '               "  WHERE CodigoDoCliente in (" & vbCrLf &
        '               "                            SELECT CodigoDoCliente" & vbCrLf &
        '               "                              FROM #SaldoFinal" & vbCrLf &
        '               "                             GROUP BY CodigoDoCliente" & vbCrLf &
        '               "                            HAVING SUM(Saldo) + sum(LaudosSemNota) > 0)" & vbCrLf
        '    Else
        '        Sql &= " SELECT * FROM #SaldoFinal Order by NomeDoCliente " & vbCrLf
        '    End If
        'Else
        '    '***************************************************************************************************************************************************
        '    '**********************************************  FOB  **********************************************************************************************
        '    '***************************************************************************************************************************************************

        '    Sql &= " /*-------------------------------------------------*/" & vbCrLf &
        '      " DECLARE @InicioDeSafra AS Date, @Vencimento AS Date" & vbCrLf &
        '      "  SELECT @InicioDeSafra = S.InicioDeSafra, @Vencimento = S.Vencimento " & vbCrLf &
        '      "    FROM Safras S" & vbCrLf &
        '      "    JOIN #Safra S2" & vbCrLf &
        '      "      ON S.Safra_Id = S2.Safra_Id" & vbCrLf &
        '      " /*-------------------------------------------------*/" & vbCrLf &
        '      " select sb.Cliente AS CodigoDoCliente," & vbCrLf &
        '      "        sb.Nome AS NomeDoCliente," & vbCrLf &
        '      "        sb.Embarque," & vbCrLf &
        '      "        sb.DataEntrega," & vbCrLf &
        '      "        SUM(sb.Contratado)      as Contratado," & vbCrLf &
        '      "        SUM(sb.ContratadoTroca) as ContratadoPorTroca," & vbCrLf &
        '      "        SUM(sb.Entregue)        as Entregue," & vbCrLf &
        '      "        SUM(sb.cedente)         as Cedente, " & vbCrLf &
        '      "        SUM(sb.cessionario)     as Cessionario, " & vbCrLf &
        '      "        SUM(convert(numeric(18,4), isnull(ls.LaudosSemNota,0)))  as LaudosSemNota," & vbCrLf &
        '      "        SUM(sb.Contratado + sb.ContratadoTroca - sb.Entregue  + sb.cedente - sb.cessionario) as Saldo," & vbCrLf &
        '      "        0.00 AS AdiantamentoEmReais," & vbCrLf &
        '      "        0.00 AS AdiantamentoEmDolares," & vbCrLf &
        '      "        0.00 as Reais," & vbCrLf &
        '      "        0.00 as Dolares" & vbCrLf &
        '      "   INTO #SaldoFinal " & vbCrLf &
        '      "   from (   " & vbCrLf &
        '      "	        select b.empresa," & vbCrLf &
        '      "	               b.cliente, " & vbCrLf &
        '      IIf(Not chkClienteConsolidado.Checked, " b.Endcliente,", "") & vbCrLf &
        '      "			  (select top 1 nome from clientes where cliente_id = b.cliente " & IIf(Not chkClienteConsolidado.Checked, " AND Endereco_Id = B.EndCliente ", " ") & ") as Nome," & vbCrLf &
        '      "                b.DataEntrega," & vbCrLf &
        '      "                case" & vbCrLf &
        '      "                  when isnull(cxt.Tipo_id,0) = 20" & vbCrLf &
        '      "                    then embarque.complemento + ' - ' + embarque.cidade + ' - ' + embarque.estado" & vbCrLf &
        '      "                    else embarque.nome + ' - ' + embarque.cidade + ' - ' + embarque.estado" & vbCrLf &
        '      "                end as Embarque," & vbCrLf &
        '      "	               b.Produto," & vbCrLf &
        '      "	               b.Nomeproduto," & vbCrLf &
        '      "	               sum(case " & vbCrLf &
        '      "	                     when isnull(p.troca,0) = 0 and so.precofixo = 'S' and so.classe <> '" & eClassesOperacoes.DEPOSITOS.ToString & "'" & vbCrLf &
        '      "	               	       then Contratado" & vbCrLf &
        '      "	               	       else 0" & vbCrLf &
        '      "	                   end)" & vbCrLf &
        '      "	               + sum(case " & vbCrLf &
        '      "	                       when isnull(p.troca,0) = 0 and so.classe = '" & eClassesOperacoes.AFIXAR.ToString & "'" & vbCrLf &
        '      "	               	         then fixada" & vbCrLf &
        '      "	               	         else 0" & vbCrLf &
        '      "	                     end) as contratado," & vbCrLf &
        '      "	               sum(case " & vbCrLf &
        '      "	               	     when isnull(p.troca,0) = 1 and so.precofixo = 'S' and so.classe <> '" & eClassesOperacoes.DEPOSITOS.ToString & "'" & vbCrLf &
        '      "	               	       then Contratado" & vbCrLf &
        '      "	               	       else 0" & vbCrLf &
        '      "	                   end) " & vbCrLf &
        '      "	               + sum(case " & vbCrLf &
        '      "	               	       when isnull(p.troca,0) = 1 and so.classe = '" & eClassesOperacoes.AFIXAR.ToString & "'" & vbCrLf &
        '      "	               	         then fixada" & vbCrLf &
        '      "	               	         else 0" & vbCrLf &
        '      "	                     end) as ContratadoTroca,		           " & vbCrLf &
        '      "	               sum(b.laudo) as Entregue,                 " & vbCrLf &
        '      "	               sum(case " & vbCrLf &
        '      "	                     when so.precofixo = 'N' and b.moeda = 1" & vbCrLf &
        '      "	               	       then b.adiantamento" & vbCrLf &
        '      "	               	       else 0" & vbCrLf &
        '      "	                   end) as AdiantamentoReais,		           " & vbCrLf &
        '      "	               sum(case " & vbCrLf &
        '      "	               	     when so.precofixo = 'N' and b.moeda = 3" & vbCrLf &
        '      "	               	       then b.adiantamento" & vbCrLf &
        '      "	               	       else 0" & vbCrLf &
        '      "	                   end) as AdiantamentoDolar," & vbCrLf &
        '      "                sum(Cedente) as cedente," & vbCrLf &
        '      "                sum(Cessionario) as Cessionario	" & vbCrLf &
        '      "	          from #base b" & vbCrLf &
        '      "	         inner join Pedidos P" & vbCrLf &
        '      "	           on b.empresa    = P.Empresa_id" & vbCrLf &
        '      "	          and b.endEmpresa = P.EndEmpresa_id" & vbCrLf &
        '      "	          and b.Pedido     = P.Pedido_id" & vbCrLf &
        '      "	         Left Join(SELECT PxD.Empresa_Id,  PxD.EndEmpresa_Id, PxD.Pedido_Id," & vbCrLf &
        '      "	                          PxD.Deposito_Id, PxD.EndDeposito_Id" & vbCrLf &
        '      "                      FROM PedidosXDepositos PxD" & vbCrLf &
        '      "                     where PxD.Tipo = 'OD'" & vbCrLf &
        '      "                       and PxD.Principal = 1" & vbCrLf &
        '      "	                   ) DE" & vbCrLf &
        '      "	           on P.Empresa_id    = DE.Empresa_id" & vbCrLf &
        '      "	          and P.EndEmpresa_id = DE.EndEmpresa_id" & vbCrLf &
        '      "	          and P.Pedido_id     = DE.Pedido_id" & vbCrLf &
        '      "	         Left Join(SELECT PxD.Empresa_Id,  PxD.EndEmpresa_Id, PxD.Pedido_Id, " & vbCrLf &
        '      "	                          PxD.Deposito_Id, PxD.EndDeposito_Id" & vbCrLf &
        '      "                           FROM PedidosXDepositos PxD" & vbCrLf &
        '      "                          where PxD.Tipo = 'LE'" & vbCrLf &
        '      "                            and PxD.Principal = 1" & vbCrLf &
        '      "	                   ) LE" & vbCrLf &
        '      "	           on P.Empresa_id    = LE.Empresa_id" & vbCrLf &
        '      "	          and P.EndEmpresa_id = LE.EndEmpresa_id" & vbCrLf &
        '      "	          and P.Pedido_id     = LE.Pedido_id" & vbCrLf &
        '      "	        inner join Clientes Embarque" & vbCrLf &
        '      "	           on Embarque.Cliente_Id  = isnull(LE.Deposito_Id   ,DE.Deposito_Id)" & vbCrLf &
        '      "	          and Embarque.Endereco_Id = isnull(LE.EndDeposito_Id,DE.EndDeposito_Id)" & vbCrLf &
        '      "          left Join clientesxtipos cxt" & vbCrLf &
        '      "            on Embarque.Cliente_Id  = cxt.Cliente_Id" & vbCrLf &
        '      "           and Embarque.Endereco_Id = cxt.Endereco_Id" & vbCrLf &
        '      "           and cxt.Tipo_id = 20" & vbCrLf &
        '      "	        inner join Suboperacoes SO" & vbCrLf &
        '      "	           on b.Operacao    = SO.Operacao_id" & vbCrLf &
        '      "	          and b.Suboperacao = SO.SubOperacoes_Id " & vbCrLf &
        '      "	        inner join Operacoes OP" & vbCrLf &
        '      "	           on OP.operacao_id = SO.Operacao_Id  " & vbCrLf &
        '      "	        WHERE 1=1" & vbCrLf

        '    If temselecionados Then
        '        Sql &= "  AND b.Empresa + cast(b.EndEmpresa as varchar) in ('" & empresasselecionadas & "')" & vbCrLf
        '    End If

        '    If Not String.IsNullOrWhiteSpace(txtCliente.Text) Then
        '        If chkClienteConsolidado.Checked Then
        '            Sql &= "And left(b.Cliente, 8) = '" & Left(txtCodigoCliente.Value.Split("-")(0), 8) & "'"
        '        Else
        '            Sql &= "And b.Cliente    = '" & txtCodigoCliente.Value.Split("-")(0) & "'" & vbCrLf &
        '                   "And b.EndCliente = " & txtCodigoCliente.Value.Split("-")(1) & vbCrLf
        '        End If
        '    End If

        '    Sql &= "    And OP.classe in ('" & eClassesOperacoes.COMPRAS.ToString & "','" & eClassesOperacoes.DEPOSITOS.ToString & "') " & vbCrLf &
        '           "  group by b.empresa,b.cliente, b.Produto, b.Nomeproduto, " & IIf(Not chkClienteConsolidado.Checked, " b.EndCliente, ", "") & vbCrLf &
        '           "           b.DataEntrega, " & vbCrLf &
        '           "           case" & vbCrLf &
        '           "             when isnull(cxt.Tipo_id,0) = 20" & vbCrLf &
        '           "               then embarque.complemento + ' - ' + embarque.cidade + ' - ' + embarque.estado" & vbCrLf &
        '           "               else embarque.nome + ' - ' + embarque.cidade + ' - ' + embarque.estado" & vbCrLf &
        '           "           end" & vbCrLf &
        '           "  ) as sb  " & vbCrLf &
        '           " left join (" & vbCrLf &
        '           "				SELECT R.Empresa_Id, R.EndEmpresa_Id, Ped.cliente, R.Produto, sum(R.PesoLiquido) as LaudosSemNota" & IIf(Not chkClienteConsolidado.Checked, ", Ped.Endcliente", "") & vbCrLf &
        '           "				  FROM Romaneios R " & vbCrLf &
        '           "				 INNER JOIN RomaneiosXPesagens RxP " & vbCrLf &
        '           "					ON R.Empresa_Id       = RxP.Empresa_Id " & vbCrLf &
        '           "				   AND R.EndEmpresa_Id    = RxP.EndEmpresa_Id" & vbCrLf &
        '           "				   AND R.Romaneio_Id      = RxP.Romaneio_Id" & vbCrLf &
        '           "                 INNER JOIN Pesagem P" & vbCrLf &
        '           "                    ON RxP.Empresa_Id    = P.Empresa_Id" & vbCrLf &
        '           "                   AND RxP.EndEmpresa_Id = P.EndEmpresa_Id" & vbCrLf &
        '           "                   AND RxP.Pesagem_Id    = P.Pesagem_Id" & vbCrLf &
        '           "                   AND P.Sequencia_Id             = 0" & vbCrLf &
        '           "                   AND P.Situacao                 = 1" & vbCrLf &
        '           "				  LEFT JOIN Pedidos Ped" & vbCrLf &
        '           "					ON R.Empresa_Id    = Ped.Empresa_Id" & vbCrLf &
        '           "				   AND R.EndEmpresa_Id = Ped.EndEmpresa_Id" & vbCrLf &
        '           "				   AND R.Pedido        = Ped.Pedido_Id" & vbCrLf &
        '           "				  LEFT OUTER JOIN Clientes" & vbCrLf &
        '           "					ON Ped.Cliente    = Clientes.Cliente_Id" & vbCrLf &
        '           "				   AND Ped.EndCliente = Clientes.Endereco_Id" & vbCrLf &
        '           "				  LEFT OUTER JOIN NotasFiscaisXRomaneios nfxr" & vbCrLf &
        '           "					ON R.Empresa_Id    = nfxr.Empresa_Id" & vbCrLf &
        '           "				   AND R.EndEmpresa_Id = nfxr.EndEmpresa_Id" & vbCrLf &
        '           "				   AND Ped.Cliente         = nfxr.Cliente_id" & vbCrLf &
        '           "				   AND Ped.EndCliente      = nfxr.EndCliente_id" & vbCrLf &
        '           "				   AND R.Romaneio_Id   = nfxr.Romaneio_Id " & vbCrLf &
        '           "				  LEFT JOIN NotasFiscais NF" & vbCrLf &
        '           "					ON NF.Empresa_Id      = nfxr.Empresa_Id" & vbCrLf &
        '           "				   AND NF.EndEmpresa_Id   = nfxr.EndEmpresa_Id" & vbCrLf &
        '           "				   AND NF.Cliente_Id      = nfxr.Cliente_Id " & vbCrLf &
        '           "				   AND NF.EndCliente_Id   = nfxr.EndCliente_Id" & vbCrLf &
        '           "				   AND NF.EntradaSaida_Id = nfxr.EntradaSaida_Id" & vbCrLf &
        '           "				   AND NF.Serie_Id        = nfxr.Serie_Id " & vbCrLf &
        '           "				   AND NF.Nota_Id         = nfxr.Nota_Id  " & vbCrLf &
        '           "				   --AND NF.Movimento       <=  @Vencimento				  " & vbCrLf &
        '           "				 Where NF.Nota_Id   IS NULL" & vbCrLf &
        '           "				   --AND R.Movimento       >=  @InicioDeSafra" & vbCrLf &
        '           "				   --AND R.Movimento       <=  @Vencimento	" & vbCrLf &
        '           "				   and R.EntradaSaida = 'E'" & vbCrLf &
        '           "                   and Ped.safra = '" & ddlSafra.SelectedValue & "'" & vbCrLf &
        '           "				 group by R.Empresa_Id, R.EndEmpresa_Id, Ped.cliente,R.Produto" & IIf(Not chkClienteConsolidado.Checked, ", Ped.Endcliente", "") & vbCrLf &
        '           "	          ) LS" & vbCrLf &
        '           "	    on sb.empresa = LS.Empresa_Id" & vbCrLf &
        '           "	   and sb.cliente = LS.cliente  " & vbCrLf &
        '            IIf(Not chkClienteConsolidado.Checked, "	   and sb.Endcliente = LS.Endcliente", "") & vbCrLf &
        '           "	   and sb.produto = LS.Produto     " & vbCrLf &
        '           " GROUP BY sb.Cliente, sb.nome, sb.Embarque, sb.DataEntrega " & vbCrLf

        '    Sql &= " SELECT * FROM #SaldoFinal where saldo + LaudosSemNota > 0 order by NomeDoCliente, dataentrega" & vbCrLf
        'End If

        'Return Sql
    End Function

    Public Function SqlPosicao() As String
        '    Dim Cliente As Cliente = Nothing
        '    Dim par As ArrayList = ucSelecaoProduto.GetSqlEParametrosRelatorio("Prd.Grupo", "Prd.Produto_Id")

        '    If Not String.IsNullOrWhiteSpace(txtCliente.Text) Then
        '        Cliente = New Cliente(txtCodigoCliente.Value.Split("-")(0), txtCodigoCliente.Value.Split("-")(1))
        '    End If

        '    Dim pos As New ListPosicaoDePedido
        '    Return pos.SqlPosicao(Date.Now, lstEmpresa1.GetSelecteds(), False, Cliente, chkClienteConsolidado.Checked, "#Base", "", ddlSafra.SelectedValue, ddlSafra.SelectedValue, par(0), "", "", "", "", "", ddlFretes.SelectedIndex, "", "", 0, 0, 0, 0, False)
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

    Private Sub EmitirExcel()
        Try
            Dim Empresa = ddlEmpresa.SelectedValue.ToString.Split("-")
            Dim objEmpresa As New [Lib].Negocio.Cliente(Empresa(0), Empresa(1))
            Dim data As String = CStr(txtData.Text)

            ds = getDataset()

            Dim dt As DataTable = New DataTable()
            dt = ds.Tables(0)

            If dt.Rows.Count = 0 Then
                MsgBox(Me.Page, "Nenhum registro encontrado para essa seleção.")
                Exit Sub
            End If

            Dim fileName As String = Server.MapPath("~/PlanilhasExcel/" & Funcoes.GeraNomeArquivo & ".xlsx")

            If File.Exists(fileName) Then
                File.Delete(fileName)
            End If

            Using arquivo As New FileStream(fileName, FileMode.CreateNew)
                Using package As New ExcelPackage(arquivo)

                    'criando planilha e título
                    Dim worksheet As ExcelWorksheet = package.Workbook.Worksheets.Add("Relatório De Posição Financeira De Vendas.")

                    'criando linha com o cabeçalho da planilha
                    Dim rowIndex As Integer = 1
                    Dim columnIndex As Integer = 1

                    'criando linha que informa o nome da empresa e o cnpj
                    worksheet.Cells(rowIndex, columnIndex).Style.Font.Bold = True
                    worksheet.Cells(rowIndex, columnIndex).Style.Font.Color.SetColor(Color.Black)
                    worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                    worksheet.Cells(rowIndex, columnIndex).Value = String.Format("{0} - {1}", objEmpresa.Nome, Funcoes.FormatarCpfCnpj(objEmpresa.Codigo))
                    worksheet.Cells(String.Format("A{0}:D{0}", rowIndex)).Merge = True
                    worksheet.Cells(String.Format("A{0}:D{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                    rowIndex += 1

                    'criando linha que informa o título do relatório
                    worksheet.Cells(rowIndex, columnIndex).Style.Font.Bold = True
                    worksheet.Cells(rowIndex, columnIndex).Style.Font.Color.SetColor(Color.Red)
                    worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                    worksheet.Cells(rowIndex, columnIndex).Value = String.Format("{0}", "POSIÇÃO FINANCEIRA DE VENDAS")
                    worksheet.Cells(String.Format("A{0}:D{0}", rowIndex)).Merge = True
                    worksheet.Cells(String.Format("A{0}:D{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                    rowIndex += 1

                    'criando linha que informa o período selecionado na página
                    worksheet.Cells(rowIndex, columnIndex).Style.Font.Bold = True
                    worksheet.Cells(rowIndex, columnIndex).Style.Font.Color.SetColor(Color.Black)
                    worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                    worksheet.Cells(rowIndex, columnIndex).Value = String.Format("{0}", "PERÍODO LIMITE : " & data)
                    worksheet.Cells(String.Format("A{0}:D{0}", rowIndex)).Merge = True
                    worksheet.Cells(String.Format("A{0}:D{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                    rowIndex += 1

                    ' criando cabeçalho da planilha com os dados do dataset
                    'criando linha com o cabeçalho da planilha
                    For Each col As DataColumn In dt.Columns
                        worksheet.Cells(rowIndex, columnIndex).Value = col.ColumnName
                        columnIndex += 1
                    Next

                    'criando auto filtro na planilha
                    worksheet.Cells("A4:" & rowIndex).AutoFilter = True

                    'aplicando formatação nas células do cabeçalho
                    Using range = worksheet.Cells(rowIndex, 1, rowIndex, dt.Columns.Count)
                        range.Style.Font.Bold = True
                        range.Style.Fill.PatternType = ExcelFillStyle.Solid
                        range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
                        range.Style.Font.Color.SetColor(Color.White)
                    End Using
                    rowIndex += 1

                    ' criando conteúdo da planilha com os dados do dataset
                    For Each row As DataRow In dt.Rows
                        columnIndex = 1

                        For Each col As DataColumn In dt.Columns
                            worksheet.Cells(rowIndex, columnIndex).Value = row(col.ColumnName)
                            columnIndex += 1
                        Next

                        'formatando células datas
                        worksheet.Cells(String.Format("B{0}", rowIndex)).Style.Numberformat.Format = "dd/MM/yyyy"
                        worksheet.Cells(String.Format("M{0}", rowIndex)).Style.Numberformat.Format = "dd/MM/yyyy"


                        'formatando células numéricas
                        worksheet.Cells(String.Format("J{0}", rowIndex)).Style.Numberformat.Format = "#,##0"
                        worksheet.Cells(String.Format("K{0}", rowIndex)).Style.Numberformat.Format = "#,##0.0000000000"
                        worksheet.Cells(String.Format("L{0}", rowIndex)).Style.Numberformat.Format = "#,##0.00"

                        worksheet.Cells(String.Format("N{0}", rowIndex)).Style.Numberformat.Format = "#,##0.00"
                        worksheet.Cells(String.Format("O{0}", rowIndex)).Style.Numberformat.Format = "#,##0.00"
                        worksheet.Cells(String.Format("P{0}", rowIndex)).Style.Numberformat.Format = "#,##0.00"
                        worksheet.Cells(String.Format("Q{0}", rowIndex)).Style.Numberformat.Format = "#,##0.00"
                        worksheet.Cells(String.Format("R{0}", rowIndex)).Style.Numberformat.Format = "#,##0.00"

                        worksheet.Cells(String.Format("S{0}", rowIndex)).Style.Numberformat.Format = "#,##0"
                        worksheet.Cells(String.Format("T{0}", rowIndex)).Style.Numberformat.Format = "#,##0"
                        worksheet.Cells(String.Format("U{0}", rowIndex)).Style.Numberformat.Format = "#,##0"

                        'aplicando formatação nas células do conteúdo
                        If rowIndex Mod 2 = 0 Then
                            Using range = worksheet.Cells(rowIndex, 1, rowIndex, dt.Columns.Count)
                                range.Style.Fill.PatternType = ExcelFillStyle.Solid
                                range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(153, 204, 255))
                            End Using
                        End If
                        rowIndex += 1
                    Next

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

    Function getDataset() As DataSet
        Sql = "Select * from ( SELECT	Emp.Reduzido + '-' + Emp.Estado as Unidade,                                                                                 " & vbCrLf &
              "						Convert(varchar, PED.DataPedido, 103) As DataPedido,                                                                            " & vbCrLf &
              "						Cli.Nome as Cliente,                                                                                                            " & vbCrLf &
              "                        Ped.Pedido_Id As Pedido,                                                                                                     " & vbCrLf &
              "                        Ped.PedidoEfetivo as PedidoCN,                                                                                               " & vbCrLf &
              "                        Cli.Estado As UF,                                                                                                            " & vbCrLf &
              "                        PIT.Produto,                                                                                                                 " & vbCrLf &
              "                        Ped.FreteCIFFOB as Frete,                                                                                                    " & vbCrLf &
              "                        EMB.Descricao As Embalagem,                                                                                                   " & vbCrLf &
              "                        --Format(PIL.ValorPedido, 'N2', 'pt-BR') As ValorPedido,	                                                                    " & vbCrLf &
              "                        --Mantendo ValorPedido como DECIMAL para cálculos e formatando apenas na exibição final                                      " & vbCrLf &
              "						--CAST(PIL.ValorPedido AS DECIMAL(18,2)) AS ValorPedido,                                                                        " & vbCrLf &
              "						--FORMAT(CAST(PIL.ValorPedido AS DECIMAL(18,2)), 'N2', 'pt-BR') AS ValorPedidoFormatado,                                        " & vbCrLf &
              "						   SUM(PXIXL.Quantidade) as Quantidade,                                                                                         " & vbCrLf &
              "						   SUM(PXIXL.UnitarioOficial)  as ValorUnitario,                                                                                " & vbCrLf &
              "						   SUM(PXIXL.TotalOficial) as ValorTotal,                                                                                       " & vbCrLf &
              "						   Ped.VencimentoPedido as Vencimento,                                                                                          " & vbCrLf &
              "						   PIL.ValorPedido,                                                                                                             " & vbCrLf &
              "                        NFXI.Faturado,                                                                                                               " & vbCrLf &
              "                        (PIL.ValorPedido - NFXI.Faturado) As  AFaturar,                                                                               " & vbCrLf &
              "                        NFXT.ValorPago,                                                                                                              " & vbCrLf &
              "                        (NFXI.Faturado - NFXT.ValorPago) As APagar,                                                                                  " & vbCrLf &
              "                        SUM(PXIXL.Quantidade) as QuantidadeEntregar,                                                                                 " & vbCrLf &
              "                        NFXI.Entregue,                                                                                                               " & vbCrLf &
              "                        (SUM(PXIXL.Quantidade) - NFXI.Entregue) As  SaldoEntregar, Emp.Reduzido                                                      " & vbCrLf &
              "                   From Pedidos Ped                                                                                                                  " & vbCrLf &
              "				  INNER Join PedidoXItem PXI ON                                                                                                         " & vbCrLf &
              "                             Ped.Empresa_Id = PXI.Empresa_Id And                                                                                     " & vbCrLf &
              "                              Ped.EndEmpresa_Id = PXI.EndEmpresa_Id And                                                                              " & vbCrLf &
              "                                Ped.Pedido_Id = PXI.Pedido_Id                                                                                        " & vbCrLf &
              "				  INNER Join PedidoXItemXLancamento PXIXL ON                                                                                            " & vbCrLf &
              "                             PXI.Empresa_Id = PXIXL.Empresa_Id And                                                                                     " & vbCrLf &
              "                              PXI.EndEmpresa_Id = PXIXL.EndEmpresa_Id And                                                                              " & vbCrLf &
              "                                PXI.Pedido_Id = PXIXL.Pedido_Id                                                                                        " & vbCrLf &
              "        INNER Join SubOperacoes SO ON                                                                                                                " & vbCrLf &
              "                              SO.Operacao_Id = Ped.Operacao And                                                                                      " & vbCrLf &
              "                              SO.SubOperacoes_Id = Ped.SubOperacao                                                                                   " & vbCrLf &
              "        OUTER APPLY(SELECT Top 1 PRD.Nome as Produto                                                                                                 " & vbCrLf &
              "									   From PedidoXItem PIT                                                                                             " & vbCrLf &
              "									INNER Join	PedidoXItem PXI ON                                                                                      " & vbCrLf &
              "                                                Ped.Empresa_Id = PIT.Empresa_Id And                                                                  " & vbCrLf &
              "                                                 Ped.EndEmpresa_Id = PIT.EndEmpresa_Id And                                                           " & vbCrLf &
              "                                                   Ped.Pedido_Id = PIT.Pedido_Id                                                                     " & vbCrLf &
              "        INNER Join  Produtos PRD ON                                                                                                                  " & vbCrLf &
              "                                                PIT.produto_Id = Prd.Produto_Id                                                                      " & vbCrLf &
              "        WHERE   PXI.Empresa_Id = PIT.Empresa_Id And                                                                                                  " & vbCrLf &
              "                                    PXI.EndEmpresa_Id = PIT.EndEmpresa_Id And                                                                        " & vbCrLf &
              "                                    PXI.Pedido_Id = PIT.Pedido_Id) AS PIT                                                                            " & vbCrLf &
              "				OUTER APPLY(SELECT Sum(Case when PIL.TipoDeLancamento = 'E' then TotalOficial  * -1 else TotalOficial End)   as ValorPedido             " & vbCrLf &
              "							From PedidoXItemXLancamento PIL                                                                                             " & vbCrLf &
              "							Where PXI.Empresa_Id = PIL.Empresa_Id And                                                                                   " & vbCrLf &
              "                                    PXI.EndEmpresa_Id = PIL.EndEmpresa_Id And                                                                        " & vbCrLf &
              "                                    PXI.Pedido_Id = PIL.Pedido_Id) As PIL                                                                            " & vbCrLf &
              "				OUTER APPLY(SELECT                                                                                                                      " & vbCrLf &
              "							isnull(Sum(Case when NFI.EntradaSaida_Id = 'S' then NFI.ValorLiquido else NFI.ValorLiquido * -1 End),0) as Faturado,    	" & vbCrLf &
              "							SUM(NFI.QuantidadeFisica) as Entregue			                                                                            " & vbCrLf &
              "							From NotasFiscais NF                                                                                                        " & vbCrLf &
              "									INNER Join	NotasFiscaisXItens NFI ON                                                                               " & vbCrLf &
              "                                                NF.Empresa_Id = NFI.Empresa_Id And                                                                   " & vbCrLf &
              "                                                NF.EndEmpresa_Id = NFI.EndEmpresa_Id And                                                             " & vbCrLf &
              "                                                NF.Cliente_Id = NFI.Cliente_Id And                                                                   " & vbCrLf &
              "                                                NF.EndCliente_Id = NFI.EndCliente_Id And                                                             " & vbCrLf &
              "                                                NF.EntradaSaida_Id = NFI.EntradaSaida_Id And                                                         " & vbCrLf &
              "                                                NF.Serie_Id = NFI.Serie_Id And                                                                       " & vbCrLf &
              "                                                NF.Nota_Id = NFI.Nota_Id                                                                             " & vbCrLf &
              "        INNER Join SubOperacoes SONF ON                                                                                                              " & vbCrLf &
              "                                                 SONF.Operacao_Id = NF.Operacao And                                                                  " & vbCrLf &
              "                                                 SONF.SubOperacoes_Id = NF.SubOperacao                                                               " & vbCrLf &
              "                                                                                                                                                     " & vbCrLf &
              "        WHERE   Ped.Empresa_Id = NF.Empresa_Id And                                                                                                   " & vbCrLf &
              "                                    Ped.EndEmpresa_Id = NF.EndEmpresa_Id And                                                                         " & vbCrLf &
              "                                    Ped.Cliente = NF.Cliente_Id And                                                                                  " & vbCrLf &
              "                                    Ped.EndCliente = NF.EndCliente_Id And                                                                            " & vbCrLf &
              "                                    SONF.Financeiro = 'S' And                                                                                        " & vbCrLf &
              "                                    NF.Situacao = 1 And                                                                                              " & vbCrLf &
              "                                    Ped.Pedido_Id = NF.Pedido) AS NFXI                                                                               " & vbCrLf &
              "				OUTER APPLY(SELECT                                                                                                                      " & vbCrLf &
              "							isnull(Sum(Case when CAR.Provisao in (1, 2)  Then Car.ValorDoDocumento Else 0 End),0) As ValorTotal,                        " & vbCrLf &
              "							isnull(Sum(Case when CAR.Provisao in (1)  Then Car.ValorLiquido Else 0 End),0) As ValorPago                                 " & vbCrLf &
              "							From NotasFiscais NF									                                                                    " & vbCrLf &
              "									INNER Join	NotaFiscalXTitulo NFXT ON                                                                               " & vbCrLf &
              "                                                NF.Empresa_Id = NFXT.Empresa_Id And                                                                  " & vbCrLf &
              "                                                NF.EndEmpresa_Id = NFXT.EndEmpresa_Id And                                                            " & vbCrLf &
              "                                                NF.Cliente_Id = NFXT.Cliente_Id And                                                                  " & vbCrLf &
              "                                                NF.EndCliente_Id = NFXT.EndCliente_Id And                                                            " & vbCrLf &
              "                                                NF.EntradaSaida_Id = NFXT.EntradaSaida_Id And                                                        " & vbCrLf &
              "                                                NF.Serie_Id = NFXT.Serie_Id And                                                                      " & vbCrLf &
              "                                                NF.Nota_Id = NFXT.Nota_Id                                                                            " & vbCrLf &
              "        INNER Join	ContasAReceber  CAR ON                                                                                                          " & vbCrLf &
              "                                                NFXT.Titulo_Id = CAR.Registro_Id And                                                                 " & vbCrLf &
              "                                                CAR.Situacao = 1                                                                                     " & vbCrLf &
              "        WHERE   PED.Empresa_Id = NF.Empresa_Id And                                                                                                   " & vbCrLf &
              "                                    PED.EndEmpresa_Id = NF.EndEmpresa_Id And                                                                         " & vbCrLf &
              "                                    PED.Cliente = NF.Cliente_Id And                                                                                  " & vbCrLf &
              "                                    PED.EndCliente = NF.EndCliente_Id And                                                                            " & vbCrLf &
              "                                    PED.Pedido_Id = NF.Pedido ) AS NFXT			                                                                    " & vbCrLf &
              "				INNER Join	Clientes Emp ON                                                                                                             " & vbCrLf &
              "                            Ped.Empresa_Id = Emp.Cliente_Id And                                                                                      " & vbCrLf &
              "                            Ped.EndEmpresa_Id = Emp.Endereco_Id                                                                                      " & vbCrLf &
              "        INNER Join	Clientes Cli ON                                                                                                                 " & vbCrLf &
              "                            Ped.Cliente = Cli.Cliente_Id And                                                                                         " & vbCrLf &
              "                            Ped.EndCliente = Cli.Endereco_Id                                                                                         " & vbCrLf &
              "                                                                                                                                                     " & vbCrLf &
              "        Left Join	Embalagens  EMB ON                                                                                                              " & vbCrLf &
              "                            Ped.Embalagem = EMB.Embalagem_Id                                                                                         " & vbCrLf &
              "        WHERE   PIL.ValorPedido Is Not NULL                                                                                                          " & vbCrLf &
              "					And   Ped.Situacao = 1                                                                                                              " & vbCrLf &
              "					And  SO.EntradaSaida = 'S'                                                                                                          " & vbCrLf &
              "					And SO.Financeiro  = 'S'                                                                                                            " & vbCrLf &
              "             		And   Left(PXI.Produto_Id, 5) in ('10101','10102','10103','10105','19101','20101','20102','20103','20104','20105','20106')       " & vbCrLf &
              "					And SO.Classe IN ('Vendas','VENDASAORDEM','GLOBAL')                                                                                 " & vbCrLf

        If Not chkConsolidarEmpresa.Checked Then
            Sql &= " And Ped.Empresa_id = '" & ddlEmpresa.SelectedValue.ToString.Split("-")(0) & "' " & vbCrLf
        End If

        If Not chkClienteConsolidado.Checked Then
            Sql &= " And Ped.Cliente = '" & txtCliente.ToString() & "' " & vbCrLf
        End If

        Sql &= " Group BY                                                                                                                                             " & vbCrLf &
               "        Emp.Reduzido,                                                                                                                                " & vbCrLf &
               "        Emp.Estado,                                                                                                                                  " & vbCrLf &
               "        Ped.DataPedido,                                                                                                                              " & vbCrLf &
               "        Cli.Nome,                                                                                                                                    " & vbCrLf &
               "        Cli.Estado,                                                                                                                                  " & vbCrLf &
               "        Ped.Pedido_Id,                                                                                                                               " & vbCrLf &
               "        Ped.PedidoEfetivo,                                                                                                                           " & vbCrLf &
               "        Ped.VencimentoPedido,                                                                                                                             " & vbCrLf &
               "        PIL.ValorPedido,                                                                                                                             " & vbCrLf &
               "        PIT.Produto,                                                                                                                                 " & vbCrLf &
               "        NFXI.Faturado,                                                                                                                               " & vbCrLf &
               "        NFXT.ValorPago,                                                                                                                              " & vbCrLf &
               "        Ped.FreteCIFFOB,                                                                                                                             " & vbCrLf &
               "        EMB.Descricao,                                                                                                                                " & vbCrLf &
               "        NFXI.Entregue                                                                                                                                " & vbCrLf &
               ") as Consulta                                                                                                                                        " & vbCrLf

        If chkPedidosComSaldo.Checked Then
            Sql &= "Where AFaturar <> 0                                                                                                                                  " & vbCrLf &
               "	And APagar <> 0                                                                                                                                     " & vbCrLf &
               "	And ValorPedido <> 0                                                                                                                                " & vbCrLf &
               "	Order By  Reduzido, Produto, Cliente, Pedido                                                                                                      " & vbCrLf
        End If

        ds = Banco.ConsultaDataSet(Sql, "PosicaoFinanceiraDeVendas")
        Return ds
    End Function

#End Region


End Class