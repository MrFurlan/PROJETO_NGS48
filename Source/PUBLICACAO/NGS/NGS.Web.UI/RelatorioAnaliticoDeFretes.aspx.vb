Imports System.IO
Imports System.Drawing
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports OfficeOpenXml.Drawing
Imports OfficeOpenXml.Style
Imports OfficeOpenXml
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class RelatorioAnaliticoDeFretes
    Inherits BasePage

#Region "Eventos"

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.setMenu(eModulo.Expedicao)
            If Not IsPostBack And IsConnect Then
                If Funcoes.VerificaPermissao("RelatorioAnaliticoDeFretes", "ACESSAR") Then
                    ddl.Carregar(ddlEmpresa, CarregarDDL.Tabela.Empresas, "", True)
                    ddl.Carregar(ddlSafra, CarregarDDL.Tabela.Safra, "", True)
                    ddl.Carregar(lstClasseOp, CarregarDDL.Tabela.ClasseDeOperacao, "isnull(operacao,0) = 1", False)

                    Limpar()
                    LiberaEmpresa()
                Else
                    MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Expedicao.aspx")
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub cmdConsultaCliente_Click(sender As Object, e As EventArgs) Handles cmdConsultaCliente.Click
        Try
            ucConsultaClientes.Limpar()
            txtPedido.Text = String.Empty
            Popup.ConsultaDeClientes(Me.Page, "objClienteNP" & HID.Value)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub BtnTransportador_Click(sender As Object, e As EventArgs) Handles BtnTransportador.Click
        Try
            ucConsultaClientes.Limpar()
            ucConsultaClientes.SetarTipoCliente(eTipoCliente.Transportadores)
            Popup.ConsultaDeClientes(Me.Page, "objClienteTrans" & HID.Value)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnPedido_Click(sender As Object, e As EventArgs) Handles btnPedido.Click
        Try
            If ddlEmpresa.SelectedIndex = 0 Then
                MsgBox(Me.Page, "Informe uma empresa.")
                ddlEmpresa.Focus()
            ElseIf String.IsNullOrWhiteSpace(txtCodigoCliente.Value) Then
                MsgBox(Me.Page, "Informe um cliente.")
                cmdConsultaCliente.Focus()
            Else
                HttpContext.Current.Session("ssCampo" & HID.Value) = "Pedidos"
                Dim strEmpresa() As String = ddlEmpresa.SelectedValue.Split("-")
                Dim strCliente() As String = txtCodigoCliente.Value.Split("-")

                Dim parameters As New Dictionary(Of String, Object)
                parameters("empresa") = strEmpresa(0)
                parameters("enderecoEmpresa") = strEmpresa(1)
                parameters("cliente") = IIf(strCliente(0) <> "", strCliente(0), "")
                parameters("enderecoCliente") = IIf(strCliente.Length > 1, strCliente(1), "")
                parameters("situacao") = eSituacao.Normal
                Popup.ConsultaDePedidos(Me.Page, "objRelatorioAnaliticoDeFretes" & HID.Value, "txtNome")
                ucConsultaPedidos.BindGridView(parameters)
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
            Limpar()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub rdoFinanceiroSim_CheckedChanged(sender As Object, e As EventArgs) Handles rdoFinanceiroSim.CheckedChanged
        Try
            chkProvisao.Enabled = rdoFinanceiroSim.Checked
            chkPrevisao.Enabled = rdoFinanceiroSim.Checked
            chkBaixa.Enabled = rdoFinanceiroSim.Checked
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub rdoFinanceiroNao_CheckedChanged(sender As Object, e As EventArgs) Handles rdoFinanceiroNao.CheckedChanged
        Try
            chkProvisao.Checked = False
            chkPrevisao.Checked = False
            chkBaixa.Checked = False
            chkProvisao.Enabled = rdoFinanceiroSim.Checked
            chkPrevisao.Enabled = rdoFinanceiroSim.Checked
            chkBaixa.Enabled = rdoFinanceiroSim.Checked
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "RelatorioAnaliticoDeFretes")
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
            txtCliente.Text = itemCliente.Text
            txtCodigoCliente.Value = itemCliente.Value
            Session.Remove("objClienteNP" & HID.Value)
        ElseIf Not Session("objClienteTrans" & HID.Value) Is Nothing Then
            Dim objCliente = CType(Session("objClienteTrans" & HID.Value), [Lib].Negocio.Cliente)
            Dim itemCliente As ListItem = Funcoes.FormatarListItemCliente(objCliente)
            txtTransportador.Text = itemCliente.Text
            txtCodigoTransportador.Value = itemCliente.Value
            Session.Remove("objClienteTrans" & HID.Value)
        ElseIf Session("objRelatorioAnaliticoDeFretes" & HID.Value) IsNot Nothing Then
            Dim objPedido As [Lib].Negocio.Pedido = CType(Session("objRelatorioAnaliticoDeFretes" & HID.Value), [Lib].Negocio.Pedido)
            txtPedido.Text = objPedido.Codigo
            Session.Remove("objRelatorioAnaliticoDeFretes" & HID.Value)
        End If
    End Sub

    Private Function IsValidFields() As Boolean
        If ddlEmpresa.SelectedIndex = 0 Then
            MsgBox(Me.Page, "Informe a empresa.")
            ddlEmpresa.Focus()
            Return False
            'ElseIf String.IsNullOrWhiteSpace(txtCodigoCliente.Value) AndAlso (String.IsNullOrWhiteSpace(txtPedido.Text) OrElse CInt(txtPedido.Text) <= 0) Then
            '    MsgBox(Me.Page, "Informe o Cliente ou Pedido.")
            '    cmdConsultaCliente.Focus()
            '    Return False
        ElseIf ucSelecaoProduto.TemSelecionado AndAlso (String.IsNullOrWhiteSpace(txtData1.Text) OrElse String.IsNullOrWhiteSpace(txtData2.Text)) Then
            MsgBox(Me.Page, "Quando algum Grupo ou Produto estiver selecionado é obrigatório o Período.")
            txtData1.Focus()
            Return False
        End If
        Return True
    End Function

    Private Function getSql(ByRef Parametros As String) As String
        Dim ds As New DataSet
        Dim sql As String = ""

        If ddlEmpresa.SelectedIndex > 0 Then Parametros &= "Empresa: " & ddlEmpresa.SelectedItem.Text & vbCrLf
        If txtTransportador.Text.Length > 0 Then Parametros &= "Transportador: " & txtTransportador.Text & vbCrLf
        If txtCliente.Text.Length > 0 Then Parametros &= "Cliente: " & txtCliente.Text & vbCrLf
        If Not String.IsNullOrWhiteSpace(txtPedido.Text) Then
            Parametros &= " Pedido: " & txtPedido.Text.Trim
        End If

        Dim op As String = String.Join("','", lstClasseOp.GetSelecteds)
        If op.Length > 0 Then
            Parametros &= IIf(rdComClasse.Checked, rdComClasse.Text, rdSemClasse.Text) & ": (" & op & ")" & vbCrLf
        End If

        Parametros &= "Frete pela Empresa " & IIf(rdFreteNaNota.Checked, "na Nota", "no Pedido") & vbCrLf
        If ddlSafra.SelectedIndex > 0 Then Parametros &= ddlSafra.SelectedItem.Text & vbCrLf
        If chkPeriodo.Checked AndAlso IsDate(txtData1.Text) AndAlso IsDate(txtData2.Text) Then
            Parametros &= "Período de: " & txtData1.Text & " até " & txtData2.Text & vbCrLf
        End If

        Dim par As New ArrayList
        par.Add("")
        If ucSelecaoProduto.TemSelecionado Then
            par = ucSelecaoProduto.GetSqlEParametrosRelatorio("P.Grupo_Id", "P.Produto_Id", "", True)
            Parametros &= par(1)
        End If

        If rdoSim.Checked Then Parametros &= "Apenas com Peso de Chegada" & vbCrLf
        If rdoSim.Checked Then Parametros &= "Apenas com financeiro: "

        Dim provisao As String = getProvisao(Parametros)

        sql &= " /*RELATORIO ANALITICO DE FRETES*/" & vbCrLf & _
               " SELECT F.ClienteCodigo, F.ClienteNome, F.ClienteEndereco, F.ClienteFazenda, F.ClienteMunicipio," & vbCrLf & _
               "        Pedido, F.NotaDeTrocaCompraCliente," & vbCrLf & _
               "		F.NotaMovimento AS Movimento, F.NotaProduto Produto_Id, F.NotaDescProduto AS DescProduto, F.Nota AS Nota_Id, F.NotaSerie AS Serie_Id, " & vbCrLf & _
               "        F.NotaPlaca Placa, " & vbCrLf & _
               "		F.Laudo NumeroTicket, " & vbCrLf & _
               "		F.LaudoPesoBalanca AS PesoBalanca," & vbCrLf & _
               "        F.LaudoPesoChegada AS  PesoChegada, " & vbCrLf & _
               "        F.LaudoQuebraSobra AS Quebra, " & vbCrLf & _
               "        F.LaudoTolerancia AS Tolerancia, " & vbCrLf & _
               "        F.LaudoQuebraSobraValor AS QuebraSobraValor," & vbCrLf & _
               "        NotaQuantidadeFiscal AS PesoNotaFiscal," & vbCrLf & _
               "        F.NotaUnitario AS Unitario, " & vbCrLf & _
               "		F.NotaValor ValorNotaFiscal, " & vbCrLf & _
               "        CASE " & vbCrLf & _
               "            WHEN (CteValor + CtePedagioValor) > 0 " & vbCrLf & _
               "                THEN CAST(CteNota as varchar) + '-' + CAST(CteSerie as varchar)" & vbCrLf & _
               "            WHEN RPAValor > 0 " & vbCrLf & _
               "                THEN CAST(RPANota AS VARCHAR) + '-' + RPASerie " & vbCrLf & _
               "            WHEN NotaDeTrocaVendaPedido > 0 " & vbCrLf & _
               "                THEN 'Troca: ' + CAST(NotaDeTrocaVendaPedido AS VARCHAR)" & vbCrLf & _
               "                ELSE ' '" & vbCrLf & _
               "        END AS Ctrc_Id," & vbCrLf & _
               "        F.PedidoFreteOrcadoUnitario   UnitFreteOrcado, " & vbCrLf & _
               "        NotaDeTrocaCompraValor AS UnitFreteDif, --FRETE COMPRA" & vbCrLf & _
               "        CASE " & vbCrLf & _
               "            WHEN (CteValor + CtePedagioValor+ RpaValor) > 0 " & vbCrLf & _
               "                THEN(CteValor + CtePedagioValor + RpaValor) * (CteUnitario + RPAUnitario) / (CteValor + CtePedagioValor+ RpaValor) " & vbCrLf & _
               "                ELSE 0" & vbCrLf & _
               "        END AS UnitarioFrete, --Unitario do Frete Realizado " & vbCrLf & _
               "        CASE " & vbCrLf & _
               "            WHEN (CteValor + CtePedagioValor + RPAValor)> 0" & vbCrLf & _
               "               THEN PedidoFreteOrcadoUnitario  + NotaDeTrocaCompraValor - ((CteValor + CtePedagioValor + RPAValor) * (CteUnitario+RPAUnitario))  /  (CteValor + CtePedagioValor + RPAValor)" & vbCrLf & _
               "               ELSE 0" & vbCrLf & _
               "        END AS DiferencaUnitFrete, --DIF. UNIT." & vbCrLf & _
               "       CASE" & vbCrLf & _
               "             WHEN (CteValor + CtePedagioValor + RPAValor)> 0" & vbCrLf & _
               "                THEN (F.PedidoFreteOrcadoUnitario  " & vbCrLf & _
               "                       + F.NotaDeTrocaCompraValor " & vbCrLf & _
               "                       - ((CteValor + CtePedagioValor + RPAValor)) * (CteUnitario+RPAUnitario)  /  (CteValor + CtePedagioValor + RPAValor))" & vbCrLf & _
               "                     * LaudoPesoChegada /1000" & vbCrLf & _
               "                ELSE 0" & vbCrLf & _
               "        END VlrDiferenca," & vbCrLf & _
               "        CtePedagioValor AS TarifaPedagio," & vbCrLf & _
               "        CteValor + CtePedagioValor + RPAValor + CteComplementoValor as ValorFrete," & vbCrLf & _
               "        CteValorLiquido + CtePedagioValor + RPAValor + CteComplementoValorLiquido as ValorLiquidoFrete," & vbCrLf & _
               "        NotaLocalEmbarque LocalEmbarque," & vbCrLf & _
               "        NotaTransportador Transportador," & vbCrLf & _
               "        F.CteEstadiaValorLiquido as Estadia" & vbCrLf & _
               "   INTO #Temp" & vbCrLf & _
               "   FROM vw_RelatorioDeFrete AS F" & vbCrLf

        If rdoFinanceiroSim.Checked AndAlso Not String.IsNullOrWhiteSpace(provisao) Then
            SqlFinanceiro(sql, provisao)
        End If

        SqlWhere(sql, provisao, Parametros)

        sql &= " ORDER BY CASE WHEN F.NotaTransportador = 'NENHUM' THEN 'ZZ' ELSE F.NotaLocalEmbarque END, F.NotaMovimento" & vbCrLf & _
               "SELECT * FROM #Temp" & vbCrLf & _
               "SELECT Pedido, " & vbCrLf & _
               "       CASE " & vbCrLf & _
               "           WHEN SUM(ValorFrete)> 0   " & vbCrLf & _
               "                THEN SUM(ISNULL(ValorFrete,0) * ISNULL(UnitarioFrete,0)) / SUM(ISNULL(ValorFrete,0)) " & vbCrLf & _
               "                ELSE 0 " & vbCrLf & _
               "       END UnitFreteRealizadoMedio, " & vbCrLf & _
               "       SUM(PesoBalanca) QtdeCarregada, SUM(PesoChegada) QtdeDescarregada, " & vbCrLf & _
               "       COUNT(Nota_Id) AS NotasEmitidas, SUM(CASE WHEN isnull(PesoChegada,0) > 0 THEN 1 ELSE 0 END) as NotasRecebidas, " & vbCrLf & _
               "       COUNT(Nota_Id) - SUM(CASE WHEN isnull(PesoChegada,0) > 0 THEN 1 ELSE 0 END) as NotasPendentes," & vbCrLf & _
               "       SUM(VlrDiferenca) / SUM(CASE WHEN ValorFrete>0 THEN PesoBalanca ELSE 1 end) * 1000 DiferencaUnitarioMedio," & vbCrLf & _
               "       Sum(Estadia) as Estadia" & vbCrLf & _
               "  FROM #Temp" & vbCrLf & _
               " GROUP BY Pedido" & vbCrLf & _
               "  DROP TABLE #Temp" & vbCrLf

        Return sql
    End Function

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

    Private Sub SqlFinanceiro(ByRef Sql As String, ByRef Provisao As String)
        Sql &= "	 LEFT JOIN ( " & vbCrLf & _
               "                SELECT COUNT(*) as Qtd, ffxi.Empresa_Id, ffxi.EndEmpresa_Id, ffxi.Cliente_Id, ffxi.EndCliente_Id, " & vbCrLf & _
               "                       ffxi.EntradaSaida_Id, ffxi.Serie_Id, ffxi.Nota_Id  " & vbCrLf & _
               "                  FROM FaturasDeFretesXItens FFxI      " & vbCrLf & _
               "                 INNER JOIN FaturaDeFretexTitulo FFxT " & vbCrLf & _
               "                    ON FFxI.EmpresaPagadora_Id = FFxT.Empresa_Id" & vbCrLf & _
               "                   AND FFxI.EndEmpresaPagadora_Id = FFxT.EndEmpresa_Id" & vbCrLf & _
               "                   AND FFxI.Conveniado_Id = FFxT.Conveniado_Id" & vbCrLf & _
               "                   AND FFxI.EndConveniado_Id = FFxT.EndConveniado_Id" & vbCrLf & _
               "                   AND FFxI.Fatura_Id = FFxT.Fatura_Id " & vbCrLf & _
               "                 INNER JOIN ContasAPagar CP " & vbCrLf & _
               "                    ON CP.Registro_Id = FFxT.Titulo_Id" & vbCrLf & _
               "                 WHERE Cp.Provisao in (" & Provisao & ") " & vbCrLf & _
               "                   AND Cp.Situacao = " & eSituacao.Normal & vbCrLf & _
               "                 GROUP BY FFxI.Empresa_Id, FFxI.EndEmpresa_Id, FFxI.Cliente_Id, FFxI.EndCliente_Id, " & vbCrLf & _
               "             	          FFxI.EntradaSaida_Id, FFxI.Serie_Id, FFxI.Nota_Id " & vbCrLf & _
               "                ) as ff  " & vbCrLf & _
               "            ON ff.Empresa_Id        = F.CteEmpresa" & vbCrLf & _
               "           AND ff.EndEmpresa_Id	    = F.CteEndEmpresa  " & vbCrLf & _
               "           AND ff.Cliente_Id		= F.CteCliente  " & vbCrLf & _
               "           AND ff.EndCliente_Id	    = F.CteEndCliente   " & vbCrLf & _
               "           AND ff.EntradaSaida_Id	= F.CteEntradaSaida " & vbCrLf & _
               "           AND ff.Serie_Id			= F.CteSerie" & vbCrLf & _
               "           AND ff.Nota_Id			= F.CteNota " & vbCrLf
    End Sub

    Private Sub SqlWhere(ByRef Sql As String, ByRef Provisao As String, Optional ByRef Parametros As String = "")
        Sql &= " WHERE F.NotaEmpresa = '" & ddlEmpresa.SelectedValue.Split("-")(0) & "'" & vbCrLf & _
               "   AND F.NotaEndEmpresa = " & ddlEmpresa.SelectedValue.Split("-")(1) & vbCrLf

        If Not String.IsNullOrWhiteSpace(txtPedido.Text) Then
            Sql &= "   AND F.Pedido = " & txtPedido.Text.Trim & vbCrLf

        End If

        If ucSelecaoProduto.TemSelecionado Then
            Sql &= "   AND  " & ucSelecaoProduto.GetSqlEParametrosRelatorio("F.NotaGrupoProduto", "F.NotaProduto", "", True)(0) & vbCrLf
        End If

        If rdoFinanceiroSim.Checked AndAlso Not String.IsNullOrWhiteSpace(Provisao) Then
            Sql &= "   AND ff.Qtd > 0 " & vbCrLf
        End If

        If rdoSim.Checked Then
            Sql &= "   AND F.LaudoPesoChegada > 0  " & vbCrLf
        End If
        If ddlSafra.SelectedIndex > 0 Then
            Sql &= "   AND F.PedidoSafra = '" & ddlSafra.SelectedValue & "'" & vbCrLf
        End If

        If chkPeriodo.Checked Then
            If IsDate(txtData1.Text) And IsDate(txtData2.Text) Then
                Sql &= "   AND " & IIf(rdMovimentoNota.Checked, " F.NotaMovimento ", "F.PedidoData") & " BETWEEN '" & CDate(txtData1.Text).ToString("yyyy-MM-dd") & "' and '" & CDate(txtData2.Text).ToString("yyyy-MM-dd") & "' " & vbCrLf
            End If
        End If

        If Not String.IsNullOrWhiteSpace(txtCodigoCliente.Value) Then
            Sql &= "   AND F.ClienteCodigo = '" & txtCodigoCliente.Value.Split("-")(0) & "' " & vbCrLf & _
                   "   AND F.ClienteEndereco = " & txtCodigoCliente.Value.Split("-")(1) & " " & vbCrLf
        End If

        If Not String.IsNullOrWhiteSpace(txtCodigoTransportador.Value) Then
            Sql &= "   AND CASE WHEN F.CteCliente <> '' THEN F.CteCliente ELSE NotaTransportadorProprietario END = '" & txtCodigoTransportador.Value.Split("-")(0) & "'      " & vbCrLf & _
                   "   AND CASE WHEN F.CteCliente <> '' THEN F.CteEndCliente ELSE NotaTransportadorEndProprietario END = " & txtCodigoTransportador.Value.Split("-")(1) & "  " & vbCrLf
        End If

        Dim op As String = String.Join("','", lstClasseOp.GetSelecteds)
        If op.Length > 0 Then
            Sql &= "   AND F.OperacaoClasse " & IIf(rdComClasse.Checked, "", "not") & " in ('" & op & "')" & vbCrLf
        End If

        If rdFreteNaNota.Checked Then
            Sql &= "   AND ((F.NotaEntradasaida = 'S' and F.NotaCIFFOB  = 'CIF') OR (F.NotaEntradasaida = 'E' and F.NotaCIFFOB = 'FOB'))" & vbCrLf
        Else
            Sql &= "   AND ((F.PedidoEntradasaida = 'S' and F.PedidoCIFFOB  = 'CIF') OR (F.PedidoEntradasaida = 'E' and F.PedidoCIFFOB = 'FOB'))" & vbCrLf
        End If

        If Not String.IsNullOrWhiteSpace(txtNota.Text) Then
            Sql &= "   AND F.Nota = " & txtNota.Text.Trim() & vbCrLf
            Parametros &= "  Nota: " & txtNota.Text.Trim
        End If

        If Not String.IsNullOrWhiteSpace(txtDACTE.Text) Then
            Sql &= "   AND F.CteNota = " & txtDACTE.Text.Trim() & vbCrLf
            Parametros &= "  CTRC: " & txtDACTE.Text.Trim()
        End If
    End Sub

    Private Function getSqlPlanilhaDePagamento() As String
        Dim Sql As String = String.Empty
        Dim Provisao = getProvisao()

        Sql = "SELECT SUBSTRING(NotaLocalEmbarque,0,15) AS LocalDeEmbarque, " & vbCrLf & _
              "       '' AS DataPagamento, " & vbCrLf & _
              "	      NotaMovimento AS Movimento, " & vbCrLf & _
              "	      CAST(CteNota AS VARCHAR) + '-' + CAST(CteSerie AS VARCHAR) CTRC, " & vbCrLf & _
              "       Pedido," & vbCrLf & _
              "	      CAST(Nota AS VARCHAR) + '-' + CAST(NotaSerie as VARCHAR) as NF, " & vbCrLf & _
              "	      CAST(CteComplementoNota as VARCHAR) + '-' + CAST(CteComplementoSerie as varchar) AS Comp," & vbCrLf & _
              "	      CAST (NotaQuantidade AS INT) as PesoNota," & vbCrLf & _
              "	      LaudoPesoChegada AS PesoChegada," & vbCrLf & _
              "       CAST((F.CteValor + F.CteComplementoValor) / F.NotaQuantidade * 1000 AS DECIMAL(18,2)) as Tarifa," & vbCrLf & _
              "       (F.LaudoPesoChegada * (F.CteValor + F.CteComplementoValor) / F.NotaQuantidade) AS VlrFrete," & vbCrLf & _
              "       F.CteComplementoValor AS VlrComp," & vbCrLf & _
              "	      CASE WHEN LaudoQuebraSobra < 0 THEN  LaudoQuebraSobra ELSE 0 END AS Quebra," & vbCrLf & _
              "	      F.LaudoTolerancia AS Tolerancia," & vbCrLf & _
              "	      NotaValor AS ValorNF," & vbCrLf & _
              "	      NotaUnitario AS UnitarioNF," & vbCrLf & _
              "       CASE WHEN F.LaudoQuebraSobraValor<0 THEN F.LaudoQuebraSobraValor ELSE 0 END AS VlrQuebra," & vbCrLf & _
              "       SUBSTRING(F.NotaDestino,0,15) AS Destino" & vbCrLf & _
              "  INTO #Temp" & vbCrLf & _
              "  FROM vw_RelatorioDeFrete F" & vbCrLf & _
              " INNER JOIN ClientesXEmpresas Emp" & vbCrLf & _
              "    ON F.NotaEmpresa = Emp.Empresa_Id" & vbCrLf & _
              "   AND F.NotaEndEmpresa = Emp.EndEmpresa_Id" & vbCrLf
        If rdoFinanceiroSim.Checked AndAlso Not String.IsNullOrWhiteSpace(Provisao) Then
            SqlFinanceiro(Sql, Provisao)
        End If

        SqlWhere(Sql, Provisao)

        'Utilizado apenas quando no critério estiver Financeiro Sim e Provisão marcado e que não estejam marcados Previsão nem Baixa
        'Feito dessa maneira para que não venham na consulta ctes vinculados a faturas de frete que tenham
        ' mais de um título e que pelo menos um deles esteja baixado
        'Quando um título da fatura já foi baixado o(s) outro(s) devem ser tratados diretamente no financeiro
        If rdoFinanceiroSim.Checked AndAlso Not chkBaixa.Checked AndAlso Not chkPrevisao.Checked AndAlso chkProvisao.Checked Then
            Sql &= "AND NOT EXISTS(" & vbCrLf & _
                   "               SELECT 1" & vbCrLf & _
                   "                 FROM FaturasDeFretesXItens FFxI  " & vbCrLf & _
                   "                 JOIN FaturaDeFretexTitulo FFxT " & vbCrLf & _
                   "                   ON FFxI.EmpresaPagadora_Id = FFxT.Empresa_Id" & vbCrLf & _
                   "                  AND FFxI.EndEmpresaPagadora_Id = FFxT.EndEmpresa_Id" & vbCrLf & _
                   "                  AND FFxI.Conveniado_Id = FFxT.Conveniado_Id" & vbCrLf & _
                   "                  AND FFxI.EndConveniado_Id = FFxT.EndConveniado_Id" & vbCrLf & _
                   "                  AND FFxI.Fatura_Id = FFxT.Fatura_Id " & vbCrLf & _
                   "                 JOIN ContasAPagar CP " & vbCrLf & _
                   "                   ON CP.Registro_Id = FFxT.Titulo_Id" & vbCrLf & _
                   "                WHERE FFxI.Empresa_Id      = F.CteEmpresa" & vbCrLf & _
                   "                  AND FFxI.EndEmpresa_Id   = F.CteEndEmpresa  " & vbCrLf & _
                   "                  AND FFxI.Cliente_Id	   = F.CteCliente  " & vbCrLf & _
                   "                  AND FFxI.EndCliente_Id   = F.CteEndCliente   " & vbCrLf & _
                   "                  AND FFxI.EntradaSaida_Id = F.CteEntradaSaida " & vbCrLf & _
                   "                  AND FFxI.Serie_Id		   = F.CteSerie" & vbCrLf & _
                   "                  AND FFxI.Nota_Id		   = F.CteNota " & vbCrLf & _
                   "                  AND CP.PROVISAO IN (1)" & vbCrLf & _
                   "                )" & vbCrLf
        End If


        Sql &= "SELECT LocalDeEmbarque, DataPagamento, Movimento, CTRC, Comp AS CTRCComp, NF, Pedido,  SUM(PesoNota) PesoNota, SUM(PesoChegada) PesoChegada, " & vbCrLf & _
               "       Tarifa,  SUM(VlrFrete - VlrComp) AS VlrFrete, SUM(VlrComp) VlrComp, SUM(Tolerancia) Tolerancia, Quebra,  SUM(ValorNF) ValorNF, SUM(ValorNF /PesoNota) AS Unitario, " & vbCrLf & _
               "       CAST(SUBSTRING(CONVERT(VARCHAR, VlrQuebra), 1,5) AS DECIMAL(18,2)) AS VlrQuebra, " & vbCrLf & _
               "       SUM((VlrFrete) - CAST(SUBSTRING(CONVERT(VARCHAR, VlrQuebra *-1), 1,5) AS DECIMAL(18,2))) AS Saldo, Destino " & vbCrLf & _
               "  FROM #Temp " & vbCrLf & _
               " GROUP BY LocalDeEmbarque, DataPagamento, Movimento, CTRC, Pedido, NF, Comp, Tarifa, Quebra," & vbCrLf & _
               "       CAST(SUBSTRING(CONVERT(VARCHAR, VlrQuebra), 1,5) AS DECIMAL(18,2)), Destino" & vbCrLf & _
               " ORDER BY CTRC, Movimento " & vbCrLf

        Return Sql
    End Function

    Private Sub EmitirRelatorio(ByVal Pdf As Boolean)
        Try
            If Funcoes.VerificaPermissao("RelatorioAnaliticoDeFretes", "RELATORIO") Then

                If rbPlanilhaDePagamento.Checked Then
                    If String.IsNullOrWhiteSpace(txtTransportador.Text) AndAlso String.IsNullOrWhiteSpace(txtPedido.Text) Then
                        MsgBox(Me.Page, "Informe Transportador ou Cliente e Pedido!")
                    ElseIf Not chkProvisao.Checked AndAlso Not chkPrevisao.Checked AndAlso Not chkBaixa.Checked Then
                        MsgBox(Me.Page, "É necessário selecionar ao menos uma situação do financeira!")
                    Else
                        Dim ds As DataSet = Banco.ConsultaDataSet(getSqlPlanilhaDePagamento, "PlanilhaDePagamento")
                        Dim TituloDaPlanilha() As String = {ddlEmpresa.SelectedItem.ToString, IIf(String.IsNullOrWhiteSpace(txtTransportador.Text.Trim()), txtCliente.Text & " Pedido: " & txtPedido.Text, txtTransportador.Text.Trim())}
                        Dim parameters As New MDTwoValues(Of String, String, eTipoCampo)
                        parameters.Add("DataPagamento", "", eTipoCampo.Data)
                        parameters.Add("Movimento", "", eTipoCampo.Data)
                        parameters.Add("Pedido", "", eTipoCampo.Numerico)
                        parameters.Add("PesoNota", "", eTipoCampo.Numerico)
                        parameters.Add("PesoChegada", "", eTipoCampo.Numerico)
                        parameters.Add("Tarifa", "", eTipoCampo.ValorSemTotalizadorCom3CasasDecimais)
                        parameters.Add("VlrFrete", "", eTipoCampo.ValorSemTotalizador)
                        parameters.Add("VlrComp", "", eTipoCampo.ValorSemTotalizador)
                        parameters.Add("Tolerancia", "", eTipoCampo.ValorSemTotalizador)
                        parameters.Add("Quebra", "", eTipoCampo.Numerico)
                        parameters.Add("ValorNF", "", eTipoCampo.ValorSemTotalizador)
                        parameters.Add("Unitario", "", eTipoCampo.ValorSemTotalizadorCom4CasasDecimais)
                        parameters.Add("VlrQuebra", "", eTipoCampo.ValorComTotalizador)
                        parameters.Add("Saldo", "", eTipoCampo.ValorComTotalizador)

                        Funcoes.BindExcelOffice(Me.Page, ds, "CONTROLE DE FRETES", parameters, TituloDaPlanilha)
                    End If

                Else
                    If IsValidFields() Then
                        Dim ds As New DataSet
                        Dim Parametros As String = String.Empty

                        ds = Banco.ConsultaDataSet(getSql(Parametros), "NotasDoPedido")

                        If ds.Tables(0) IsNot Nothing AndAlso ds.Tables(0).Rows.Count > 0 Then
                            ds.Tables(0).TableName = "NotasDoPedido"
                            ds.Tables(1).TableName = "Resumo"

                            Dim nomeDoReport As String = IIf(rdoPorPedido.Checked, "Cr_RelatorioAnaliticoDeFretes", "Cr_RelatorioAnaliticoDeFretesPorTransportador")
                            Dim Par As New Dictionary(Of String, Object)()

                            Dim objEmpresa As Cliente = New Cliente(ddlEmpresa.SelectedValue.Split("-")(0), ddlEmpresa.SelectedValue.Split("-")(1))

                            Par.Add("Parametros", Parametros)

                            Funcoes.BindReport(Me.Page, ds, nomeDoReport, IIf(Pdf, eExportType.PDF, eExportType.ExcelCrystal), Par)
                        Else
                            MsgBox(Me.Page, "Sem Registros!")
                        End If
                    End If
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para emitir relatório.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub Limpar()
        ddlEmpresa.SelectedIndex = ddlEmpresa.Items.IndexOf(ddlEmpresa.Items.FindByValue(UsuarioServidor.CodigoEmpresa & "-" & UsuarioServidor.EnderecoEmpresa))
        txtCliente.Text = String.Empty
        txtCodigoCliente.Value = String.Empty
        txtTransportador.Text = String.Empty
        txtCodigoTransportador.Value = String.Empty
        txtPedido.Text = String.Empty
        lstClasseOp.SelectedIndex = -1
        rdFreteNaNota.Checked = True
        ddlSafra.SelectedIndex = 0
        chkPeriodo.Checked = False
        chkPeriodo_CheckedChanged(chkPeriodo, New EventArgs)
        txtData1.Text = Format(CType(Now.AddDays(-180), Date), "dd/MM/yyyy")
        txtData2.Text = Format(CType(Now, Date), "dd/MM/yyyy")
        rdoPorPedido.Checked = True
        rdoPorTransportador.Checked = False
        rdoFinanceiroSim.Checked = False
        rdoFinanceiroNao.Checked = True
        rdoFinanceiroNao_CheckedChanged(New Object, New EventArgs)
        rdoSim.Checked = False
        rdoNao.Checked = True
        chkProvisao.Enabled = rdoFinanceiroSim.Checked
        chkPrevisao.Enabled = rdoFinanceiroSim.Checked
        chkBaixa.Enabled = rdoFinanceiroSim.Checked
        HID.Value = Guid.NewGuid().ToString()
        ucConsultaClientes.SetarHID(HID.Value)
        ucConsultaPedidos.SetarHID(HID.Value)
        ucSelecaoProduto.Limpar()
    End Sub

    Private Sub LiberaEmpresa()
        If Not UsuarioServidor.LiberaEmpresa Then
            ddlEmpresa.Enabled = False
        End If
    End Sub

#End Region

    Protected Sub rbPlanilhaDePagamento_CheckedChanged(sender As Object, e As EventArgs) Handles rbPlanilhaDePagamento.CheckedChanged
        divPdf.Style.Add("display", "none")
    End Sub

    Protected Sub rdoPorPedido_CheckedChanged(sender As Object, e As EventArgs) Handles rdoPorPedido.CheckedChanged
        divPdf.Style.Clear()
    End Sub

    Protected Sub rdoPorTransportador_CheckedChanged(sender As Object, e As EventArgs) Handles rdoPorTransportador.CheckedChanged
        divPdf.Style.Clear()
    End Sub

    Protected Sub chkPeriodo_CheckedChanged(sender As Object, e As EventArgs) Handles chkPeriodo.CheckedChanged
        If CType(sender, CheckBox).Checked Then
            divTipoDePeriodo.Style.Clear()
        Else
            divTipoDePeriodo.Style.Add("Display", "none")
        End If
    End Sub
End Class