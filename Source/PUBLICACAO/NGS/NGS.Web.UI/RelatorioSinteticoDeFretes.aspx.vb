Imports System.IO
Imports System.Drawing
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports OfficeOpenXml.Drawing
Imports OfficeOpenXml.Style
Imports OfficeOpenXml
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class RelatorioSinteticoDeFretes
    Inherits BasePage

#Region "Eventos"

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Expedicao)
        If Not IsPostBack And IsConnect Then
            If Funcoes.VerificaPermissao("RelatorioSinteticoDeFretes", "ACESSAR") Then
                ddl.Carregar(ddlEmpresa, CarregarDDL.Tabela.Empresas, "", True)
                ddl.Carregar(ddlSafra, CarregarDDL.Tabela.Safra, "", True)
                ucSelecaoProduto.WhereProduto = "Agrupar = 'N'"
                ddl.Carregar(lstClasseOp, CarregarDDL.Tabela.ClasseDeOperacao, "isnull(operacao,0) = 1", False)

                Limpar()
                LiberaEmpresa()
            Else
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Expedicao.aspx")
                Exit Sub
            End If
        End If
    End Sub

    Protected Sub cmdConsultaCliente_Click(sender As Object, e As EventArgs) Handles cmdConsultaCliente.Click
        ucConsultaClientes.Limpar()
        Popup.ConsultaDeClientes(Me.Page, "objClienteNP" & HID.Value)
    End Sub

    Protected Sub BtnTransportador_Click(sender As Object, e As EventArgs) Handles BtnTransportador.Click
        ucConsultaClientes.Limpar()
        ucConsultaClientes.SetarTipoCliente(eTipoCliente.Transportadores)
        Popup.ConsultaDeClientes(Me.Page, "objClienteTrans" & HID.Value)
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
        Limpar()
    End Sub

    Protected Sub rdoPorTransportador_CheckedChanged(sender As Object, e As EventArgs)
        Try
            txtTransportador.Text = String.Empty
            txtCodigoTransportador.Value = String.Empty
            BtnTransportador.Enabled = Not rbPorCliente.Checked
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub rdoPorPedido_CheckedChanged(sender As Object, e As EventArgs)
        Try
            txtTransportador.Text = String.Empty
            txtCodigoTransportador.Value = String.Empty
            BtnTransportador.Enabled = Not rbPorCliente.Checked
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub rdoFinanceiroSim_CheckedChanged(sender As Object, e As EventArgs) Handles rdoFinanceiroSim.CheckedChanged
        chkProvisao.Enabled = rdoFinanceiroSim.Checked
        chkPrevisao.Enabled = rdoFinanceiroSim.Checked
        chkBaixa.Enabled = rdoFinanceiroSim.Checked
    End Sub

    Protected Sub rdoFinanceiroNao_CheckedChanged(sender As Object, e As EventArgs) Handles rdoFinanceiroNao.CheckedChanged
        chkBaixa.Checked = False
        chkPrevisao.Checked = False
        chkProvisao.Checked = False
        chkProvisao.Enabled = rdoFinanceiroSim.Checked
        chkPrevisao.Enabled = rdoFinanceiroSim.Checked
        chkBaixa.Enabled = rdoFinanceiroSim.Checked
    End Sub

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "RelatorioSinteticoDeFretes")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

#End Region

#Region "Métodos"

    Public Overrides Sub Carregar(ByVal obj As IBaseEntity)
        If Session("objClienteNP" & HID.Value) IsNot Nothing Then
            Dim objCliente As [Lib].Negocio.Cliente = CType(Session("objClienteNP" & HID.Value), [Lib].Negocio.Cliente)
            Dim itemCliente As ListItem = Funcoes.FormatarListItemCliente(objCliente)
            txtClientes.Text = itemCliente.Text
            txtCodigoCliente.Value = itemCliente.Value
            Session.Remove("objClienteNP" & HID.Value)
        ElseIf Not Session("objClienteTrans" & HID.Value) Is Nothing Then
            Dim objCliente = CType(Session("objClienteTrans" & HID.Value), [Lib].Negocio.Cliente)
            Dim itemCliente As ListItem = Funcoes.FormatarListItemCliente(objCliente)
            txtTransportador.Text = itemCliente.Text
            txtCodigoTransportador.Value = itemCliente.Value
            Session.Remove("objClienteTrans" & HID.Value)
        End If
    End Sub

    Private Function IsValidFields() As Boolean
        If ddlEmpresa.SelectedIndex = 0 Then
            MsgBox(Me.Page, "Informe a empresa!")
            ddlEmpresa.Focus()
            Return False
        End If
        Return True
    End Function

    Private Sub EmitirRelatorio(ByVal Pdf As Boolean)
        Try
            If Funcoes.VerificaPermissao("RelatorioSinteticoDeFretes", "RELATORIO") Then

                If IsValidFields() Then
                    Dim parametros As String = String.Empty
                    Dim ds As New DataSet
                    Dim sql As String = ""

                    If ddlEmpresa.SelectedIndex > 0 Then parametros &= "Empresa: " & ddlEmpresa.SelectedItem.Text & vbCrLf
                    If txtTransportador.Text.Length > 0 Then parametros &= "Transportador: " & txtTransportador.Text & vbCrLf
                    If txtClientes.Text.Length > 0 Then parametros &= "Cliente: " & txtClientes.Text & vbCrLf
                    Dim op As String = String.Join("','", lstClasseOp.GetSelecteds)

                    If op.Length > 0 Then
                        parametros &= IIf(rdComClasse.Checked, rdComClasse.Text, rdSemClasse.Text) & ": (" & op & ")" & vbCrLf
                    End If

                    parametros &= "Frete pela Empresa " & IIf(rdFreteNaNota.Checked, "na Nota", "no Pedido") & vbCrLf

                    If ddlSafra.SelectedIndex > 0 Then parametros &= ddlSafra.SelectedItem.Text & vbCrLf
                    If chkPeriodo.Checked AndAlso IsDate(txtDataDe.Text) AndAlso IsDate(txtDataAte.Text) Then
                        parametros &= "Período de: " & txtDataDe.Text & " até " & txtDataAte.Text & vbCrLf
                    End If

                    If chkFinalizados.Checked Then
                        parametros &= "Somente Pedidos Finalizados" & vbCrLf
                    End If

                    Dim par As New ArrayList
                    par.Add("")
                    If ucSelecaoProduto.TemSelecionado Then
                        par = ucSelecaoProduto.GetSqlEParametrosRelatorio("F.NotaGrupoProduto", "F.NotaProduto", "", True)
                        parametros &= par(1)
                    End If

                    If rdoSim.Checked Then parametros &= "Apenas com Peso de Chegada" & vbCrLf
                    If rdoSim.Checked Then parametros &= "Apenas com financeiro: "

                    Dim provisao As String = getProvisao(parametros)

                    Dim AgrupadoPor As String = IIf(rbPorCliente.Checked, "Cliente", "Transportador")
                    Dim NomeDoArquivo As String = IIf(rbPorCliente.Checked, "Cr_RelatorioSinteticoDeFretes", "Cr_RelatorioSinteticoDeFretesPorTransportador")

                    sql = getSql(provisao, par(0), AgrupadoPor)

                    ds = Banco.ConsultaDataSet(sql, "Fretes")

                    Dim parameters = New Dictionary(Of String, Object)()
                    parameters.Add("Parametros", parametros)

                    Funcoes.BindReport(Me.Page, ds, NomeDoArquivo, IIf(Pdf, eExportType.PDF, eExportType.ExcelCrystal), parameters)
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para emitir relatório.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub
    Private Function getProvisao(Optional ByRef Parametros As String = "") As String
        Dim Provisao As String = String.Empty
        If rdoFinanceiroSim.Checked Then

            If chkProvisao.Checked Then
                If Not String.IsNullOrWhiteSpace(Provisao) Then
                    Provisao &= ", "
                End If
                Provisao &= "3"
                Parametros &= "Provisão(3) "
            End If
            If chkPrevisao.Checked Then
                If Not String.IsNullOrWhiteSpace(Provisao) Then
                    Provisao &= ", "
                End If
                Provisao &= "2"
                Parametros &= "Previsão(2) "
            End If
            If chkBaixa.Checked Then
                If Not String.IsNullOrWhiteSpace(Provisao) Then
                    Provisao &= ", "
                End If
                Provisao &= "1"
                Parametros &= "Baixado(1) "
            End If
        End If
        Return Provisao
    End Function

    Private Function getSql(ByRef provisao As String, ByVal pSqlProduto As String, ByVal AgruparPor As String)
        Dim Sql As String = String.Empty
        Sql = " /*Relatório Sintético de Fretes */  " & vbCrLf & _
              "  SELECT F.Pedido, F.ClienteCodigo, F.ClienteNome, " & vbCrLf & _
              "         F.ClienteEndereco, F.ClienteFazenda, F.ClienteMunicipio," & vbCrLf & _
              "         F.Nota," & vbCrLf & _
              "		    F.NotaQuantidadeFiscal," & vbCrLf & _
              "		    F.NotaQuantidade," & vbCrLf & _
              "		    F.Notavalor," & vbCrLf & _
              "		    F.LaudoPesoChegada," & vbCrLf & _
              "         F.LaudoPesoBalanca," & vbCrLf & _
              "		    F.PedidoFreteOrcadoUnitario UnitFreteOrcado, " & vbCrLf & _
              "	        F.NotaDeTrocaCompraValor, " & vbCrLf & _
              "		    F.CtePedagioValor," & vbCrLf & _
              "		    F.CteValor," & vbCrLf & _
              "		    F.CteUnitario," & vbCrLf & _
              "		    F.CteQuantidadeFiscal," & vbCrLf & _
              "		    F.RpaUnitario," & vbCrLf & _
              "		    F.RpaValor," & vbCrLf & _
              " 	    F.PedidoFreteOrcadoUnitario," & vbCrLf & _
              "		    F.LaudoQuebraSobra," & vbCrLf & _
              "		    F.LaudoTolerancia," & vbCrLf & _
              "		    F.LaudoQuebraSobraValor," & vbCrLf & _
              "		    CASE WHEN  Pedido.Quantidade >0  THEN Pedido.Quantidade - NFxQS.QuantidadeNFQuebraSobra ELSE 0 END AS PedidoQuantidadeLiquida," & vbCrLf & _
              "         F.NotaTransportador, " & vbCrLf & _
              "         CASE " & vbCrLf & _
              "             WHEN (CteValor + CtePedagioValor + RPAValor) > 0" & vbCrLf & _
              "                 THEN PedidoFreteOrcadoUnitario  + NotaDeTrocaCompraValor - (CteValor + CtePedagioValor + RPAValor) * (CteUnitario + RPAUnitario)  /  (CteValor + CtePedagioValor + RPAValor)" & vbCrLf & _
              "                 ELSE 0" & vbCrLf & _
              "         END AS DiferencaUnitFrete, --DIF. UNIT." & vbCrLf & _
              "		    --CASE" & vbCrLf & _
              "             --WHEN CteNota > 0" & vbCrLf & _
              "                 --THEN (F.PedidoFreteOrcadoUnitario  " & vbCrLf & _
              "                    --   + F.NotaDeTrocaCompraValor " & vbCrLf & _
              "                   --    - (CteValor + CtePedagioValor) * (CteUnitario)  /  CASE WHEN CteValor =0 then 1 else CteValor END)  " & vbCrLf & _
              "                    -- * ISNULL((CASE WHEN NotaQuantidadeFisica <= 0 THEN NotaQuantidadeFiscal ELSE NotaQuantidadeFisica END),0)  /1000" & vbCrLf & _
              "                -- ELSE 0" & vbCrLf & _
              "         --END VlrDiferenca" & vbCrLf & _
              "       CASE" & vbCrLf & _
              "             WHEN (CteValor + CtePedagioValor + RPAValor)> 0" & vbCrLf & _
              "                THEN (F.PedidoFreteOrcadoUnitario  " & vbCrLf & _
              "                       + F.NotaDeTrocaCompraValor " & vbCrLf & _
              "                       - ((CteValor + CtePedagioValor + RPAValor)) * (CteUnitario+RPAUnitario)  /  (CteValor + CtePedagioValor + RPAValor))" & vbCrLf & _
              "                     * LaudoPesoChegada /1000" & vbCrLf & _
              "                ELSE 0" & vbCrLf & _
              "        END VlrDiferenca," & vbCrLf & _
              "        f.CteEstadiaValorLiquido" & vbCrLf
        '"	 CASE" & vbCrLf & _
        ' "WHEN (CteValor + CtePedagioValor + RPAValor) > 0" & vbCrLf & _
        '"   THEN (F.PedidoFreteOrcadoUnitario  " & vbCrLf & _
        ' "         + F.NotaDeTrocaCompraValor " & vbCrLf & _
        ' "         - (CteValor + CtePedagioValor + RPAValor) * (CteUnitario+ RPAUnitario)  /  (CteValor + CtePedagioValor + RPAValor))  " & vbCrLf & _
        ' "       * ISNULL((CASE WHEN NotaQuantidadeFisica <= 0 THEN NotaQuantidadeFiscal ELSE NotaQuantidadeFisica END),0)  /1000" & vbCrLf & _
        ' "  ELSE 0" & vbCrLf & _
        ' "END VlrDiferenca" & vbCrLf

        Sql &= getFrom(provisao)
        Sql &= getWhere(pSqlProduto, provisao)

        Sql &= "  SELECT F.Pedido, F.ClienteCodigo+'-'+CONVERT(VARCHAR,F.ClienteEndereco) AS ClienteCodigo, F.ClienteNome," & vbCrLf & _
               "         F.ClienteEndereco, F.ClienteFazenda AS ClienteComplemento, F.ClienteMunicipio ClienteCidadeUF," & vbCrLf & _
               "		 COUNT(F.Nota) AS NotasEmitidas," & vbCrLf & _
               "		 SUM(case " & vbCrLf & _
               "               When isnull(F.LaudoPesoChegada,0) > 0" & vbCrLf & _
               "                 then 1" & vbCrLf & _
               "                 else 0" & vbCrLf & _
               "             end) as Recebidas," & vbCrLf & _
               "		 COUNT(*) - SUM(case " & vbCrLf & _
               "                          When isnull(F.LaudoPesoChegada,0) > 0" & vbCrLf & _
               "	         				then 1" & vbCrLf & _
               "	         				else 0" & vbCrLf & _
               "	         			end) as Pendentes,	" & vbCrLf & _
               "		 CAST(SUM(CASE WHEN (CteValor + CtePedagioValor + RpaValor) > 0 THEN F.LaudoPesoBalanca ELSE 0 END) AS INT) as PesoBalanca," & vbCrLf & _
               "         CAST(SUM(CASE WHEN (CteValor + CtePedagioValor + RpaValor) > 0 THEN F.LaudoPesoChegada ELSE 0 END) AS INT) as PesoChegada," & vbCrLf & _
               "         CAST(SUM(F.LaudoQuebraSobra) AS INT) AS Quebra," & vbCrLf & _
               "		 --SUM(F.LaudoQuebraSobra) * CAST(SUM(CASE WHEN F.CteValor>0 THEN F.NotaValor ELSE 0 END) / SUM(CASE WHEN F.CteValor>0 THEN F.NotaQuantidadeFiscal ELSE 1 END) AS DECIMAL(18,4)) as ValorQuebra," & vbCrLf & _
               "         SUM(F.LaudoQuebraSobraValor) AS ValorQuebra, " & vbCrLf & _
               "         CAST(SUM(CASE WHEN (CteValor + CtePedagioValor + RpaValor) > 0 THEN F.NotaQuantidadeFiscal ELSE 0 END) AS INT) AS PesoNotaFiscal," & vbCrLf & _
               "         CAST(SUM(CASE WHEN (CteValor + CtePedagioValor + RpaValor) > 0 THEN F.NotaValor ELSE 0 END) / SUM(CASE WHEN (CteValor + CtePedagioValor + RpaValor) > 0 THEN F.NotaQuantidadeFiscal ELSE 1 END) AS DECIMAL(18,4)) as UnitarioNota," & vbCrLf & _
               "         SUM(CASE WHEN (CteValor + CtePedagioValor + RpaValor) > 0 THEN F.NotaValor ELSE 0 END)  as ValorNotaFiscal," & vbCrLf & _
               "		 F.PedidoFreteOrcadoUnitario UnitFreteOrcado, " & vbCrLf & _
               "         CASE WHEN SUM(CASE WHEN NotaDeTrocaCompraValor>0 THEN F.LaudoPesoBalanca ELSE 0 end) >0 then " & vbCrLf & _
               "                SUM(CASE WHEN NotaDeTrocaCompraValor>0 THEN F.LaudoPesoBalanca  * NotaDeTrocaCompraValor  ELSE 0 END)  " & vbCrLf & _
               "                / SUM(CASE WHEN NotaDeTrocaCompraValor>0 THEN F.LaudoPesoBalanca ELSE 0 end)ELSE 0 " & vbCrLf & _
               "         END  UnitFreteCompra," & vbCrLf & _
               "	     --SUM((CteValor + CtePedagioValor) * CteUnitario) / CASE WHEN SUM(CteValor) = 0 THEN 1 ELSE SUM(CteValor) END AS UnitFreteRealizado, --Somando o Pedagio " & vbCrLf & _
               "         CASE " & vbCrLf & _
               "             WHEN SUM(CteValor + CtePedagioValor + RpaValor) > 0   " & vbCrLf & _
               "                 THEN SUM(ISNULL(CteValor + CtePedagioValor + RpaValor,0) * ISNULL(cteUnitario + RPAunitario,0)) / SUM(ISNULL(CteValor + CtePedagioValor + RpaValor,0)) " & vbCrLf & _
               "                 ELSE 0 " & vbCrLf & _
               "         END UnitFreteRealizado, " & vbCrLf & _
               "         SUM(F.VlrDiferenca) / SUM(CASE WHEN (CteValor + CtePedagioValor + RpaValor)>0 THEN F.LaudoPesoBalanca ELSE 1 end) *1000 UnitFreteDiferenca," & vbCrLf & _
               "         CASE WHEN SUM(CteValor + CtePedagioValor + RPAValor)> 0" & vbCrLf & _
               "             THEN SUM(F.VlrDiferenca) / SUM(CASE WHEN (CteValor + CtePedagioValor + RpaValor)>0 THEN F.LaudoPesoBalanca ELSE 1 end) *1000 " & vbCrLf & _
               "             ELSE 0" & vbCrLf & _
               "         END AS UnitFreteDiferenca, " & vbCrLf & _
               "         SUM(F.VlrDiferenca) - sum(f.CteEstadiaValorLiquido) AS VlrFreteDiferenca," & vbCrLf & _
               "		 SUM(CtePedagioValor) AS TarifaPedagio," & vbCrLf & _
               "		 SUM(CteValor + CtePedagioValor + RPAValor) AS ValorFrete," & vbCrLf & _
               "         sum(f.CteEstadiaValorLiquido) * -1 as Estadias," & vbCrLf & _
               "		 CAST(SUM(CASE WHEN PedidoQuantidadeLiquida > 0 THEN  F.CteQuantidadefiscal / PedidoQuantidadeLiquida * 100 ELSE PedidoQuantidadeLiquida END) AS INT)  AS PercentualDeConclusao," & vbCrLf
        Sql &= IIf(AgruparPor = "Transportador", "       F.NotaTransportador AS Transportador", "       '' Transportador ") & vbCrLf


        Sql &= "    FROM #Temp F" & vbCrLf & _
               "GROUP BY F.Pedido, F.ClienteCodigo, F.ClienteEndereco, F.ClienteNome, F.ClienteFazenda, F.ClienteMunicipio," & vbCrLf & _
               "         F.PedidoFreteOrcadoUnitario  " & vbCrLf

        If AgruparPor = "Cliente" Then
            Sql &= "ORDER BY ClienteNome" & vbCrLf
        Else
            Sql &= "		 ,F.NotaTransportador" & vbCrLf & _
                   "ORDER BY CASE WHEN NotaTransportador = 'NENHUM' THEN 'ZZ' ELSE NotaTransportador END" & vbCrLf
        End If
        Sql &= getHaving()

        Return Sql
    End Function

    Private Function getFrom(ByVal provisao As String) As String
        Dim Sql As String = String.Empty
        Sql = "   INTO #Temp" & vbCrLf & _
              "   FROM vw_relatoriodefrete AS F" & vbCrLf & _
              "   LEFT JOIN (select PR.Empresa_Id," & vbCrLf & _
              "		        	     PR.EndEmpresa_Id," & vbCrLf & _
              "		        		 PR.Pedido_Id," & vbCrLf & _
              "                      sum(pr.QuantidadeAtual * valor / case when isnull(pr.QuantidadeAtual,0) = 0 then 1 else pr.QuantidadeAtual end)/count(*) *1000 as ValorPonderado" & vbCrLf & _
              "		        	from pedidos P" & vbCrLf & _
              "		           inner join PedidoXRoteiro PR" & vbCrLf & _
              "		              on P.Empresa_Id    = PR.Empresa_Id" & vbCrLf & _
              "		             and P.EndEmpresa_Id = PR.EndEmpresa_id" & vbCrLf & _
              "		             and P.Pedido_id     = PR.Pedido_Id" & vbCrLf & _
              "		           group by PR.Empresa_Id," & vbCrLf & _
              "				         PR.EndEmpresa_Id," & vbCrLf & _
              "		    	         PR.Pedido_Id) OrcPedido" & vbCrLf & _
              "      ON OrcPedido.Empresa_Id    = F.NotaEmpresa" & vbCrLf & _
              "     AND OrcPedido.EndEmpresa_Id = F.NotaEndEmpresa" & vbCrLf & _
              "     AND OrcPedido.Pedido_id     = F.Pedido" & vbCrLf & _
              "	INNER JOIN (SELECT Pxi.Empresa_Id, PxI.EndEmpresa_Id, PxI.Pedido_Id, " & vbCrLf & _
              "                     SUM(CASE " & vbCrLf & _
              "                        WHEN Pxi.TipoDeLancamento = 'E' " & vbCrLf & _
              "                            THEN Quantidade * -1" & vbCrLf & _
              "                            ELSE Quantidade" & vbCrLf & _
              "                         END) AS Quantidade" & vbCrLf & _
              "                FROM PedidoXItemXLancamento Pxi" & vbCrLf & _
              "	               GROUP BY PxI.Empresa_Id, PxI.EndEmpresa_Id, PxI.Pedido_Id" & vbCrLf & _
              "             ) AS Pedido" & vbCrLf & _
              "     ON F.NotaEmpresa    = Pedido.Empresa_id" & vbCrLf & _
              "    AND F.NotaEndEmpresa = Pedido.EndEmpresa_id" & vbCrLf & _
              "    AND F.Pedido        = Pedido.Pedido_Id" & vbCrLf & _
              "	     LEFT JOIN (" & vbCrLf & _
              "                SELECT NFSQ.Empresa_Id, NFSQ.EndEmpresa_Id, NFSQ.Pedido, SUM(ISNULL(NFSQxI.QuantidadeFiscal, 0)) QuantidadeNFQuebraSobra " & vbCrLf & _
              "                  FROM NotasFiscais AS NFSQ " & vbCrLf & _
              "                  LEFT JOIN  NotasFiscaisXItens NFSQxI " & vbCrLf & _
              "                    ON NFSQ.Empresa_Id = NFSQxI.Empresa_Id " & vbCrLf & _
              "                   AND NFSQ.EndEmpresa_Id = NFSQxI.EndEmpresa_Id " & vbCrLf & _
              "                   AND NFSQ.Cliente_Id = NFSQxI.Cliente_Id " & vbCrLf & _
              "                   AND NFSQ.EndCliente_Id = NFSQxI.EndCliente_Id " & vbCrLf & _
              "                   AND NFSQ.EntradaSaida_Id = NFSQxI.EntradaSaida_Id " & vbCrLf & _
              "                   AND NFSQ.Serie_Id = NFSQxI.Serie_Id " & vbCrLf & _
              "                   AND NFSQ.Nota_Id = NFSQxI.Nota_Id" & vbCrLf & _
              "                   AND NOT ((NFSQ.entradasaida_id = 'S' and NFSQ.ciffob = 'CIF') or (NFSQ.entradasaida_id = 'E' and NFSQ.ciffob = 'FOB'))" & vbCrLf & _
              "                 GROUP BY NFSQ.Empresa_Id, NFSQ.EndEmpresa_Id, NFSQ.Pedido" & vbCrLf & _
              "             ) AS NFxQS" & vbCrLf & _
              "     ON F.NotaEmpresa = NFxQS.Empresa_Id" & vbCrLf & _
              "    AND F.NotaEndEmpresa = NFxQS.EndEmpresa_Id" & vbCrLf & _
              "    AND F.Pedido = NFxQS.Pedido" & vbCrLf

        If rdoFinanceiroSim.Checked AndAlso Not String.IsNullOrWhiteSpace(provisao) Then
            Sql &= "	 LEFT JOIN (      " & vbCrLf & _
                   "                SELECT COUNT(*) as Qtd, ffxi.Empresa_Id, ffxi.EndEmpresa_Id, ffxi.Cliente_Id, ffxi.EndCliente_Id, " & vbCrLf & _
                   "                       ffxi.EntradaSaida_Id, ffxi.Serie_Id, ffxi.Nota_Id  " & vbCrLf & _
                   "                  FROM FaturasDeFretesXItens ffxi      " & vbCrLf & _
                   "                 INNER JOIN FaturasDeFretes ff      " & vbCrLf & _
                   "                    ON  ff.Empresa_Id = ffxi.EmpresaPagadora_Id " & vbCrLf & _
                   "                   AND ff.EndEmpresa_Id = ffxi.EndEmpresaPagadora_Id       " & vbCrLf & _
                   "                   AND ff.Conveniado_Id = ffxi.Conveniado_Id " & vbCrLf & _
                   "                   AND ff.EndConveniado_Id = ffxi.EndConveniado_Id       " & vbCrLf & _
                   "                   AND ff.Fatura_Id = ffxi.Fatura_Id " & vbCrLf & _
                   "                 INNER JOIN ContasAPagar Cp " & vbCrLf & _
                   "                    ON  cp.Registro_Id = ff.Titulo" & vbCrLf & _
                   "                 WHERE Cp.Provisao in (" & provisao & ") " & vbCrLf & _
                   "                 GROUP BY ffxi.Empresa_Id, ffxi.EndEmpresa_Id, ffxi.Cliente_Id, ffxi.EndCliente_Id, " & vbCrLf & _
                   "             		        ffxi.EntradaSaida_Id, ffxi.Serie_Id, ffxi.Nota_Id " & vbCrLf & _
                   "                ) as ff  " & vbCrLf & _
                   "            ON ff.Empresa_Id         = F.CteEmpresa" & vbCrLf & _
                   "           AND ff.EndEmpresa_Id	    = F.CteEndEmpresa  " & vbCrLf & _
                   "           AND ff.Cliente_Id		= F.CteCliente  " & vbCrLf & _
                   "           AND ff.EndCliente_Id	    = F.CteEndCliente   " & vbCrLf & _
                   "           AND ff.EntradaSaida_Id	= F.CteEntradaSaida " & vbCrLf & _
                   "           AND ff.Serie_Id			= F.CteSerie" & vbCrLf & _
                   "           AND ff.Nota_Id			= F.CteNota " & vbCrLf
        End If

        Return Sql
    End Function

    Private Function getWhere(ByVal pSqlProduto As String, ByVal provisao As String) As String
        Dim sql As String = String.Empty

        sql = " WHERE 1=1 " & vbCrLf & _
               "   AND F.NotaEmpresa = '" & ddlEmpresa.SelectedValue.Split("-")(0) & "'" & vbCrLf & _
               "   AND F.NotaEndEmpresa = " & ddlEmpresa.SelectedValue.Split("-")(1) & vbCrLf

        If Not String.IsNullOrWhiteSpace(txtCodigoCliente.Value) Then
            sql &= "   AND F.ClienteCodigo = '" & txtCodigoCliente.Value.Split("-")(0) & "'" & vbCrLf & _
                   "   AND F.ClienteEndereco = " & txtCodigoCliente.Value.Split("-")(1) & vbCrLf
        End If

        Dim op As String = String.Join("','", lstClasseOp.GetSelecteds)
        If op.Length > 0 Then
            sql &= "   AND F.OperacaoClasse " & IIf(rdComClasse.Checked, "", "not") & " in ('" & op & "')" & vbCrLf
        End If

        If rdFreteNaNota.Checked Then
            sql &= "   AND ((F.NotaEntradasaida = 'S' and F.NotaCIFFOB  = 'CIF') OR (F.NotaEntradasaida = 'E' and F.NotaCIFFOB = 'FOB'))" & vbCrLf
        Else
            sql &= "   AND ((F.PedidoEntradasaida = 'S' and F.PedidoCIFFOB  = 'CIF') OR (F.PedidoEntradasaida = 'E' and F.PedidoCIFFOB = 'FOB'))" & vbCrLf
        End If

        If chkPeriodo.Checked Then
            If IsDate(txtDataDe.Text) And IsDate(txtDataAte.Text) Then
                sql &= "   AND " & IIf(rdMovimentoNota.Checked, " F.NotaMovimento ", "F.PedidoData") & " BETWEEN '" & CDate(txtDataDe.Text).ToString("yyyy-MM-dd") & "' and '" & CDate(txtDataAte.Text).ToString("yyyy-MM-dd") & "' " & vbCrLf
            Else
                MsgBox(Me.Page, "Informe a data correta")
                Return ""
            End If
        End If

        If ddlSafra.SelectedIndex > 0 Then
            sql &= "AND F.PedidoSafra = '" & ddlSafra.SelectedValue & "'" & vbCrLf
        End If

        If Not String.IsNullOrEmpty(pSqlProduto) Then
            sql &= "   AND " & pSqlProduto & vbCrLf
        End If

        If rdoFinanceiroSim.Checked AndAlso Not String.IsNullOrWhiteSpace(provisao) Then
            sql &= "    AND ff.Qtd > 0 " & vbCrLf
        End If

        If rdoSim.Checked Then
            sql &= "   	AND F.LaudoPesoChegada > 0   " & vbCrLf
        End If

        If Not String.IsNullOrWhiteSpace(txtCodigoTransportador.Value) Then
            sql &= "    AND CASE WHEN F.CteCliente <> '' THEN F.CteCliente ELSE NotaTransportadorProprietario END = '" & txtCodigoTransportador.Value.Split("-")(0) & "'      " & vbCrLf & _
                   "    AND CASE WHEN F.CteCliente <> '' THEN F.CteEndCliente ELSE NotaTransportadorEndProprietario END = " & txtCodigoTransportador.Value.Split("-")(1) & "  " & vbCrLf
        End If

        Return sql
    End Function

    Private Function getHaving() As String
        Dim sql As String = String.Empty
        If chkFinalizados.Checked Then sql &= "HAVING CAST(SUM(F.CteQuantidadefiscal / PedidoQuantidadeLiquida * 100) AS INT)>=100" & vbCrLf
        Return sql
    End Function

    Private Sub Limpar()
        ddlEmpresa.SelectedIndex = ddlEmpresa.Items.IndexOf(ddlEmpresa.Items.FindByValue(UsuarioServidor.CodigoEmpresa & "-" & UsuarioServidor.EnderecoEmpresa))
        txtClientes.Text = String.Empty
        txtCodigoCliente.Value = String.Empty
        txtTransportador.Text = String.Empty
        txtCodigoTransportador.Value = String.Empty
        rbPorCliente.Checked = True
        rbPorTransportador.Checked = False
        chkPeriodo.Checked = False
        chkPeriodo_CheckedChanged(chkPeriodo, New EventArgs)
        rdoFinanceiroSim.Checked = False
        rdoFinanceiroNao.Checked = True
        rdoFinanceiroNao_CheckedChanged(New Object, New EventArgs)
        rdoSim.Checked = False
        rdoNao.Checked = True
        ddlSafra.SelectedIndex = 0
        chkProvisao.Enabled = rdoFinanceiroSim.Checked
        chkPrevisao.Enabled = rdoFinanceiroSim.Checked
        chkBaixa.Enabled = rdoFinanceiroSim.Checked
        HID.Value = Guid.NewGuid().ToString()
        ucConsultaClientes.SetarHID(HID.Value)
        ucConsultaPedidos.SetarHID(HID.Value)
        txtDataDe.Text = Format(CType(Now.AddDays(-180), Date), "dd/MM/yyyy")
        txtDataAte.Text = Format(CType(Now, Date), "dd/MM/yyyy")
        ucSelecaoProduto.Limpar()
    End Sub

    Private Sub LiberaEmpresa()
        If Not UsuarioServidor.LiberaEmpresa Then
            ddlEmpresa.Enabled = False
        End If
    End Sub

#End Region

    Protected Sub chkPeriodo_CheckedChanged(sender As Object, e As EventArgs) Handles chkPeriodo.CheckedChanged
        If CType(sender, CheckBox).Checked Then
            divTipoDePeriodo.Style.Clear()
        Else
            divTipoDePeriodo.Style.Add("Display", "none")
        End If
    End Sub
End Class