Imports System.Data
Imports System.Collections.Generic
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class ucPedidoFixacao
    Inherits BaseUserControl

#Region "Variaveis Locais"
    Private objPedido As [Lib].Negocio.Pedido
    Private objFixacao As [Lib].Negocio.Fixacao
#End Region

#Region "Session"
    Private Sub SessaoRecuperaPedido()
        objPedido = CType(Session("objPedido" & HID.Value), [Lib].Negocio.Pedido)
        If objPedido Is Nothing Then objPedido = New [Lib].Negocio.Pedido
    End Sub

    Private Sub SessaoSalvaPedido()
        Session("objPedido" & HID.Value) = objPedido
    End Sub

    Private Sub SessaoRecuperaFixacao()
        objFixacao = CType(Session("objFixacao" & HID.Value), [Lib].Negocio.Fixacao)
    End Sub

    Private Sub SessaoSalvaFixacao()
        Session("objFixacao" & HID.Value) = objFixacao
    End Sub
#End Region

#Region "Events"

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            If Not IsPostBack And IsConnect Then
                ddl.Carregar(ddlCondicoesFixacao, CarregarDDL.Tabela.CondicaoDePagamento, "")
                TabVencimentosOld.Visible = Not FinanceiroNovo
                TabVencimentos.Visible = FinanceiroNovo
                If FinanceiroNovo Then ngsFinanceiro.ConfigurarFinanceiro("FX")



            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub imgConsultarFixacao_Click(sender As Object, e As System.Web.UI.ImageClickEventArgs)
        Try
            divVencimento.Visible = False
            lnkCalcularValores.Parent.Visible = False

            SessaoRecuperaPedido()
            Dim Linha As Integer = CType(CType(sender, ImageButton).NamingContainer, GridViewRow).RowIndex
            objFixacao = objPedido.Itens(HIDLinhaProduto.Value).Fixacoes(Linha)
            objFixacao.IUD = "U"

            lblFixacao.Text = objFixacao.Codigo
            txtCessaoDeCreditoAfixar.Text = objFixacao.Procuracao
            txtDataAfixar.Text = objFixacao.Movimento.ToString("dd-MM-yyyy")
            txtIndiceFixadoItem.Text = objFixacao.IndiceFixado.ToString("N8")
            txtQuantidadeAFixar.Text = objFixacao.Quantidade.ToString("N4")
            txtUnitarioAfixar.Text = objFixacao.UnitarioOficial.ToString("N10")
            txtTotalAfixar.Text = objFixacao.TotalOficial.ToString("N2")
            txtTotalAfixarLiquido.Text = objFixacao.Encargos.Where(Function(s) s.CodigoEncargo = "LIQUIDO").First.ValorOficial.ToString("N2")
            ddlSubAFixar.SelectedValue = objFixacao.CodigoOperacao & "-" & objFixacao.CodigoSubOperacao

            gridFixacaoXNotaFiscal.DataSource = objPedido.Itens(HIDLinhaProduto.Value).Fixacoes(Linha).NotasFixacao.ToArray
            gridFixacaoXNotaFiscal.DataBind()

            grdEncargosFixacao.DataSource = objFixacao.Encargos.ToArray
            grdEncargosFixacao.DataBind()

            SessaoSalvaFixacao()
            SessaoSalvaPedido()

            If FinanceiroNovo Then
                If objFixacao.SubOperacao.Financeiro Then
                    ngsFinanceiro.SetarHID(HID.Value)
                    ngsFinanceiro.SetarOrigem = "FX"
                    ngsFinanceiro.ConfigurarFinanceiro("FX")
                    ngsFinanceiro.CarregarControle()
                    lnkNovo.Parent.Visible = True
                End If
                TabContainer1.ActiveTabIndex = 3
            Else
                If objPedido.Vencimentos.Count > 0 Then
                    Dim Where As String
                    Where = "     Empresa       ='" & objPedido.CodigoEmpresa & "'" & vbCrLf & _
                            " And EndEmpresa    = " & objPedido.EnderecoEmpresa & vbCrLf & _
                            " And Pedido        = " & objPedido.Codigo & vbCrLf & _
                            " And PedidoFixacao = " & objFixacao.Codigo
                    Dim Vencimentos As New [Lib].Negocio.ListTitulo(Where)

                    Dim intParcela As Integer = 1
                    Dim TotalPaceladoFixacao As Decimal

                    Dim dtVencimentos As New DataTable()
                    dtVencimentos.Columns.Add("Codigo", Type.GetType("System.Int32"))
                    dtVencimentos.Columns.Add("TipoPagamento", Type.GetType("System.Int32"))
                    dtVencimentos.Columns.Add("Sequencia", Type.GetType("System.Int32"))
                    dtVencimentos.Columns.Add("Parcela", Type.GetType("System.Int32"))
                    dtVencimentos.Columns.Add("Vencimento", Type.GetType("System.DateTime"))
                    dtVencimentos.Columns.Add("Valor", Type.GetType("System.Double"))
                    dtVencimentos.Columns.Add("ValorMoeda", Type.GetType("System.Double"))
                    dtVencimentos.Columns.Add("DataPagamento", Type.GetType("System.DateTime"))
                    dtVencimentos.Columns.Add("ValorPagamento", Type.GetType("System.Double"))
                    dtVencimentos.Columns.Add("ValorPagamentoMoeda", Type.GetType("System.Double"))
                    dtVencimentos.Columns.Add("Saldo", Type.GetType("System.Double"))
                    dtVencimentos.Columns.Add("SaldoMoeda", Type.GetType("System.Double"))
                    dtVencimentos.Columns.Add("Atraso", Type.GetType("System.Double"))

                    For Each tit As [Lib].Negocio.Titulo In Vencimentos
                        Dim drVencimento As DataRow = dtVencimentos.NewRow()
                        drVencimento("Codigo") = tit.Codigo
                        drVencimento("TipoPagamento") = tit.CodigoTipoPgto
                        drVencimento("Sequencia") = 0
                        drVencimento("Parcela") = intParcela
                        drVencimento("Vencimento") = tit.Prorrogacao
                        drVencimento("Valor") = tit.ValorDoDocumento
                        drVencimento("ValorMoeda") = tit.MoedaValorDoDocumento
                        If tit.CodigoProvisao = 1 Then
                            drVencimento("DataPagamento") = tit.Baixa
                            drVencimento("ValorPagamento") = tit.ValorLiquido
                            drVencimento("ValorPagamentoMoeda") = tit.MoedaValorLiquido
                            drVencimento("Saldo") = 0
                            drVencimento("SaldoMoeda") = 0
                            drVencimento("Atraso") = 0
                        Else
                            drVencimento("DataPagamento") = DBNull.Value
                            drVencimento("ValorPagamento") = 0
                            drVencimento("Saldo") = 0
                            drVencimento("SaldoMoeda") = 0
                            drVencimento("Atraso") = 0
                        End If


                        dtVencimentos.Rows.Add(drVencimento)

                        TotalPaceladoFixacao += tit.ValorDoDocumento

                        intParcela += 1
                    Next

                    gridCondicoesFixacao.DataSource = dtVencimentos
                    gridCondicoesFixacao.DataBind()

                    txtTotalPaceladoFixacao.Text = TotalPaceladoFixacao.ToString("N2")

                    TabContainer1.ActiveTabIndex = 3
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub imgExcluirFixacao_Click(sender As Object, e As System.Web.UI.ImageClickEventArgs)
        Try
            If Funcoes.VerificaPermissao("PedidosXItens", "EXCLUIR") Then
                Dim SqlArray As New ArrayList
                Dim linha As Integer = CType(CType(sender, ImageButton).NamingContainer, GridViewRow).RowIndex

                SessaoRecuperaPedido()
                objFixacao = objPedido.Itens(HIDLinhaProduto.Value).Fixacoes(linha)

                If objFixacao.VerNotaFiscal Then
                    MsgBox(Me.Page, "Fixacão com Nota Fiscal não pode ser excluída.")
                    Exit Sub
                End If

                If FinanceiroNovo Then
                    If objFixacao.Titulos.Where(Function(s) s.CodigoSituacao = 1 And s.CodigoProvisao = 1).ToList.Count > 0 Then
                        MsgBox(Me.Page, "Fixacão com Financeiro Pago não pode ser excluída.")
                        Exit Sub
                    End If

                    objFixacao.IUD = "D"

                    For Each tit In objFixacao.Titulos
                        tit.IUD = "D"
                    Next
                    objFixacao.SalvarSql(SqlArray)

                Else
                    If objFixacao.VerVencimentos Then
                        MsgBox(Me.Page, "Fixacão com Financeiro Pago não pode ser excluída.")
                        Exit Sub
                    Else
                        objFixacao.IUD = "D"
                        objFixacao.SalvarSql(SqlArray)

                        Dim Where As String
                        Where = "     Empresa       ='" & objPedido.CodigoEmpresa & "'" & vbCrLf & _
                                " And EndEmpresa    = " & objPedido.EnderecoEmpresa & vbCrLf & _
                                " And Pedido        = " & objPedido.Codigo & vbCrLf & _
                                " And PedidoFixacao = " & objFixacao.Codigo

                        Dim Vencimentos As New [Lib].Negocio.ListTitulo(Where)

                        For Each Titulo As [Lib].Negocio.Titulo In Vencimentos
                            Titulo.IUD = "D"
                            Titulo.SalvarSql(SqlArray)
                        Next
                    End If
                End If

                If Banco.GravaBanco(SqlArray) Then
                    SetarParametros(HIDLinhaProduto.Value, HIDPedido.Value)
                    MsgBox(Me.Page, "Fixação Excluída com Sucesso.", eTitulo.Sucess)
                    CType(Me.Page, PedidosXItens).DesabilitaAlteracao()
                Else
                    MsgBox(Me.Page, "Erro ao Excluir Fixação: " & Funcoes.EliminarCaracteresEspeciais(RTrim(HttpContext.Current.Session("ssMessage").ToString)) & ". Entre em contato com o Suporte.")
                End If

            Else : MsgBox(Me.Page, "Usuário sem permissão para excluir Fixação.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub ddlCondicoesFixacao_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddlCondicoesFixacao.SelectedIndexChanged
        Try
            SessaoRecuperaPedido()
            SessaoRecuperaFixacao()

            Dim Vencimentos As New [Lib].Negocio.ListTitulo()
            objFixacao.CondicaoPagamento = ddlCondicoesFixacao.SelectedValue
            objPedido.IndiceFixado = objFixacao.IndiceFixado
            Vencimentos.ParcelarFixacao(objFixacao, objPedido)
            Session("VencimentosFixacao" & HID.Value) = Vencimentos

            SessaoSalvaPedido()
            SessaoSalvaFixacao()

            AtualizarGridVencimentos()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub gridCondicoesFixacao_SelectedIndexChanged(sender As Object, e As EventArgs) Handles gridCondicoesFixacao.SelectedIndexChanged
        Try
            If Not Session("VencimentosFixacao" & HID.Value) Is Nothing Then
                SessaoRecuperaPedido()

                txtDataVencimentoFixacao.Text = CType(Session("VencimentosFixacao" & HID.Value), [Lib].Negocio.ListTitulo)(gridCondicoesFixacao.SelectedIndex).Prorrogacao

                If objPedido.Moeda.Classificacao = [Lib].Negocio.eTiposMoeda.Oficial Then
                    txtValorVencimentoFixacao.Text = CType(Session("VencimentosFixacao" & HID.Value), [Lib].Negocio.ListTitulo)(gridCondicoesFixacao.SelectedIndex).ValorDoDocumento.ToString("N2")
                    ValorVencimentoFixacao.Value = CType(Session("VencimentosFixacao" & HID.Value), [Lib].Negocio.ListTitulo)(gridCondicoesFixacao.SelectedIndex).ValorDoDocumento.ToString("N2")
                Else
                    txtValorVencimentoFixacao.Text = CType(Session("VencimentosFixacao" & HID.Value), [Lib].Negocio.ListTitulo)(gridCondicoesFixacao.SelectedIndex).MoedaValorDoDocumento.ToString("N2")
                    ValorVencimentoFixacao.Value = CType(Session("VencimentosFixacao" & HID.Value), [Lib].Negocio.ListTitulo)(gridCondicoesFixacao.SelectedIndex).MoedaValorDoDocumento.ToString("N2")
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkCalcularValores_Click(sender As Object, e As EventArgs) Handles lnkCalcularValores.Click
        Try
            SessaoRecuperaPedido()
            SessaoRecuperaFixacao()

            If objPedido.FiscalAberto = False AndAlso Not Funcoes.VerificaAcesso(objPedido.Empresa.Codigo, objPedido.Empresa.CodigoEndereco, IIf(objPedido.DataVencimentoPedido > Now.Date, Now.Date, objPedido.DataVencimentoPedido), "PEDIDOS") Then
                MsgBox(Me.Page, "Movimento já Fechado para esta data.")
            ElseIf txtQuantidadeAFixar.Text.Length = 0 Then
                MsgBox(Me.Page, "Quantidade à Fixar não foi informada")
            ElseIf CDec(txtQuantidadeAFixar.Text) = 0 Then
                MsgBox(Me.Page, "Quantidade à Fixar não pode ser Zero")
            ElseIf txtUnitarioAfixar.Text.Length = 0 Then
                MsgBox(Me.Page, "Unitário à Fixar não foi informado")
            ElseIf CDec(txtUnitarioAfixar.Text) = 0 Then
                MsgBox(Me.Page, "Unitário à Fixar não pode ser Zero")
            ElseIf CDec(txtQuantidadeAFixar.Text) > CDec(txtSaldoAfixar.Text) Then
                MsgBox(Me.Page, "Quantidade à Fixar não pode ser maior que Saldo à Fixar")
            ElseIf objPedido.Itens(HIDLinhaProduto.Value).Fixacoes.SaldoProcuracao > 0 AndAlso _
                CInt(txtCessaoDeCreditoAfixar.Text) = 0 AndAlso _
                CInt(txtQuantidadeAFixar.Text) > (objFixacao.NotasFixacao.Sum(Function(s) s.QtdeAFixar) - objPedido.Itens(HIDLinhaProduto.Value).Fixacoes.SaldoProcuracao) Then
                MsgBox(Me.Page, "Pedido com Cessão de Crédito não pode ter Quantidade à Fixar maior que Saldo à Fixar sem Cessão de Crédito.")
            ElseIf CInt(txtCessaoDeCreditoAfixar.Text) > 0 AndAlso CInt(txtQuantidadeAFixar.Text) > CInt(txtSaldoProcuracao.Text) Then
                MsgBox(Me.Page, "Quantidade à Fixar maior que Saldo à Fixar da Cessão de Crédito.")
                'ElseIf ValorTotalFixacao <> ValorTotalNotaFiscal Then
                '    MsgBox(Me.Page, "O Valor total da(s) fixação(ões) está diferente do valor total de notas.")
            Else

                objFixacao.IUD = "I"
                objFixacao.CodigoOperacao = objPedido.CodigoOperacao
                objFixacao.CodigoSubOperacao = ddlSubAFixar.SelectedValue.Split("-")(1)
                objFixacao.Movimento = CDate(txtDataAfixar.Text)
                objFixacao.IndiceFixado = txtIndiceFixadoItem.Text
                objFixacao.Quantidade = txtQuantidadeAFixar.Text
                objFixacao.Procuracao = CInt(txtCessaoDeCreditoAfixar.Text)

                If objPedido.Moeda.Classificacao = [Lib].Negocio.eTiposMoeda.Oficial Then
                    objFixacao.UnitarioOficial = Convert.ToDecimal(txtUnitarioAfixar.Text.Replace(".", ""))
                    objFixacao.UnitarioMoeda = 0
                    '#FimBaseDeCalculo
                    'objFixacao.TotalOficial = Math.Round(objFixacao.UnitarioOficial * objFixacao.Quantidade / objPedido.Itens(HIDLinhaProduto.Value).Produto.BaseCalculo, 2, MidpointRounding.AwayFromZero)
                    objFixacao.TotalOficial = Math.Round(objFixacao.UnitarioOficial * objFixacao.Quantidade, 2, MidpointRounding.AwayFromZero)
                    objFixacao.TotalMoeda = 0
                Else
                    objFixacao.UnitarioMoeda = Convert.ToDecimal(txtUnitarioAfixar.Text.Replace(".", ""))

                    If objFixacao.IndiceFixado > 0 Then
                        objFixacao.UnitarioOficial = Math.Round(objFixacao.IndiceFixado * objFixacao.UnitarioMoeda, 10, MidpointRounding.AwayFromZero)
                    Else
                        objFixacao.UnitarioOficial = Math.Round(Funcoes.ConverteParaMoedaOficial(objFixacao.UnitarioMoeda, objPedido.Indexador.Codigo, objFixacao.Movimento), 10, MidpointRounding.AwayFromZero)
                    End If
                    '#FimBaseDeCalculo
                    'objFixacao.TotalOficial = Math.Round(objFixacao.UnitarioOficial * objFixacao.Quantidade / objPedido.Itens(HIDLinhaProduto.Value).Produto.BaseCalculo, 2, MidpointRounding.AwayFromZero)
                    'objFixacao.TotalMoeda = Math.Round(objFixacao.UnitarioMoeda * objFixacao.Quantidade / objPedido.Itens(HIDLinhaProduto.Value).Produto.BaseCalculo, 2, MidpointRounding.AwayFromZero)
                    objFixacao.TotalOficial = Math.Round(objFixacao.UnitarioOficial * objFixacao.Quantidade, 2, MidpointRounding.AwayFromZero)
                    objFixacao.TotalMoeda = Math.Round(objFixacao.UnitarioMoeda * objFixacao.Quantidade, 2, MidpointRounding.AwayFromZero)
                End If

                If objFixacao.Encargos Is Nothing OrElse objFixacao.Encargos.Count = 0 Then
                    MsgBox(Me.Page, "Encargos não foram encontrados, verifique!")
                    Exit Sub
                End If

                objFixacao.NotasFixacao.CarregarNotasParaSelecao()

                '********************************************************************************************************************************
                '**************************************************   ADICIONAR FIXACAO   *******************************************************
                '********************************************************************************************************************************
                Dim QtdFixada As Decimal = Convert.ToDecimal(txtQuantidadeAFixar.Text.Replace(".", ""))
                Dim EncQtdAFixar As Decimal = 0
                Dim EncVlrAFixar As Decimal = 0
                Dim EncFacs As Decimal = 0
                Dim EncFethab As Decimal = 0
                Dim EncFundersul As Decimal = 0
                Dim EncFundems As Decimal = 0
                Dim EncFacsAntes As Decimal = 0
                Dim EncFethabAntes As Decimal = 0
                Dim EncFundersulAntes As Decimal = 0
                Dim EncFundemsAntes As Decimal = 0

                If objFixacao.Encargos.Exists(Function(s) s.CodigoEncargo = "FACS") Then EncFacsAntes = objFixacao.Encargos.Find(Function(s) s.CodigoEncargo = "FACS").ValorOficial
                If objFixacao.Encargos.Exists(Function(s) s.CodigoEncargo = "FETHAB") Then EncFethabAntes = objFixacao.Encargos.Find(Function(s) s.CodigoEncargo = "FETHAB").ValorOficial
                If objFixacao.Encargos.Exists(Function(s) s.CodigoEncargo = "FUNDERSUL") Then EncFundersulAntes = objFixacao.Encargos.Find(Function(s) s.CodigoEncargo = "FUNDERSUL").ValorOficial
                If objFixacao.Encargos.Exists(Function(s) s.CodigoEncargo = "FUNDEMS") Then EncFundemsAntes = objFixacao.Encargos.Find(Function(s) s.CodigoEncargo = "FUNDEMS").ValorOficial

                For Each Nota In objFixacao.NotasFixacao
                    If QtdFixada = 0 Then
                        Exit For
                    ElseIf QtdFixada >= Nota.QtdeAFixar Then
                        Nota.Liberada = True
                        Nota.QtdeAFixarLiberado = Nota.QtdeAFixar
                        Nota.ValorAFixarLiberado = Nota.ValorAFixar
                        EncQtdAFixar += Nota.QtdeAFixarLiberado
                        EncVlrAFixar += Nota.ValorAFixarLiberado
                        QtdFixada -= Nota.QtdeAFixar

                        EncFacs += Math.Round((Nota.QtdeAFixar / 1000) * Nota.PercentualFacs, 2, MidpointRounding.AwayFromZero)
                        EncFethab += Math.Round((Nota.QtdeAFixar / 1000) * Nota.PercentualFethab, 2, MidpointRounding.AwayFromZero)
                        EncFundersul += Math.Round((Nota.QtdeAFixar / 1000) * Nota.PercentualFundersul, 2, MidpointRounding.AwayFromZero)
                        EncFundems += Math.Round((Nota.QtdeAFixar / 1000) * Nota.PercentualFudems, 2, MidpointRounding.AwayFromZero)

                        '#ValorFixacaoReal
                        'If Nota.UnitarioNotaFiscal > Nota.Fixacao.UnitarioOficial Then
                        '    Nota.ValorFixacaoReal = Nota.QtdeAFixar * Nota.Fixacao.UnitarioOficial
                        'Else
                        '    Nota.ValorFixacaoReal = Nota.QtdeAFixar * Nota.Fixacao.UnitarioOficial
                        'End If

                    Else
                        Nota.Liberada = True
                        Nota.QtdeAFixarLiberado = QtdFixada
                        '#FimBaseDeCalculo
                        Nota.ValorAFixarLiberado = Math.Round(QtdFixada * Nota.UnitarioNotaFiscal, 2, MidpointRounding.AwayFromZero)
                        EncQtdAFixar += Nota.QtdeAFixarLiberado
                        EncVlrAFixar += Nota.ValorAFixarLiberado

                        EncFacs += Math.Round((QtdFixada / 1000) * Nota.PercentualFacs, 2, MidpointRounding.AwayFromZero)
                        EncFethab += Math.Round((QtdFixada / 1000) * Nota.PercentualFethab, 2, MidpointRounding.AwayFromZero)
                        EncFundersul += Math.Round((QtdFixada / 1000) * Nota.PercentualFundersul, 2, MidpointRounding.AwayFromZero)
                        EncFundems += Math.Round((QtdFixada / 1000) * Nota.PercentualFudems, 2, MidpointRounding.AwayFromZero)

                        '#ValorFixacaoReal
                        'If Nota.UnitarioNotaFiscal > Nota.Fixacao.UnitarioOficial Then
                        '    Nota.ValorFixacaoReal = QtdFixada * Nota.Fixacao.UnitarioOficial
                        'Else
                        '    Nota.ValorFixacaoReal = QtdFixada * Nota.Fixacao.UnitarioOficial
                        'End If

                        Exit For
                    End If
                Next

                If objFixacao.Encargos.All(Function(s) s.CodigoEncargo <> "AFIXAR") Then
                    Dim AFixar As New FixacaoXEncargo(objFixacao)
                    AFixar.CodigoEncargo = "AFIXAR"
                    AFixar.BaseMoeda = objFixacao.Quantidade
                    AFixar.BaseOficial = objFixacao.Quantidade
                    AFixar.Percentual = 100

                    'Passa o valor para trandferir Saldo a Fixar para Fixo no Lote 15 - Furlan - 01/10/2014
                    If objFixacao.TotalOficial > EncVlrAFixar Then
                        AFixar.ValorOficial = EncVlrAFixar
                        AFixar.ValorMoeda = Funcoes.ConverteMoeda(EncVlrAFixar, objFixacao.IndiceFixado, eTiposMoeda.MoedaEstrangeira, True, False, 2)
                    Else
                        AFixar.ValorOficial = objFixacao.TotalOficial
                        AFixar.ValorMoeda = objFixacao.TotalMoeda
                    End If

                    objFixacao.Encargos.Add(AFixar)
                Else
                    Dim AFixar As FixacaoXEncargo = objFixacao.Encargos.Where(Function(s) s.CodigoEncargo = "AFIXAR").FirstOrDefault()
                    AFixar.BaseMoeda = objFixacao.Quantidade
                    AFixar.BaseOficial = objFixacao.Quantidade
                    AFixar.Percentual = 100

                    'Passa o valor para trandferir Saldo a Fixar para Fixo no Lote 15 - Furlan - 01/10/2014
                    If objFixacao.TotalOficial > EncVlrAFixar Then
                        AFixar.ValorOficial = EncVlrAFixar
                        AFixar.ValorMoeda = Funcoes.ConverteMoeda(EncVlrAFixar, objFixacao.IndiceFixado, eTiposMoeda.MoedaEstrangeira, True, False, 2)
                    Else
                        AFixar.ValorOficial = objFixacao.TotalOficial
                        AFixar.ValorMoeda = objFixacao.TotalMoeda
                    End If
                End If

                '********************************************************************************************************************************
                '********************************************************************************************************************************
                '********************************************************************************************************************************

                gridFixacaoXNotaFiscal.DataSource = objFixacao.NotasFixacao.ToArray()
                gridFixacaoXNotaFiscal.DataBind()

                For Each enc In objFixacao.Encargos
                    If enc.CodigoEncargo = "FACS" Then
                        enc.Percentual = 0
                        enc.ValorOficial = EncFacs
                    ElseIf enc.CodigoEncargo = "FETHAB" Then
                        enc.Percentual = 0
                        enc.ValorOficial = EncFethab
                    ElseIf enc.CodigoEncargo = "FUNDERSUL" Then
                        enc.Percentual = 0
                        enc.ValorOficial = EncFundersul
                    ElseIf enc.CodigoEncargo = "FUNDEMS" Then
                        enc.Percentual = 0
                        enc.ValorOficial = EncFundems
                    ElseIf enc.CodigoEncargo = "LIQUIDO" Then
                        enc.ValorOficial += EncFacsAntes
                        enc.ValorOficial += EncFethabAntes
                        enc.ValorOficial += EncFundersulAntes
                        enc.ValorOficial += EncFundemsAntes
                        enc.ValorOficial -= EncFacs
                        enc.ValorOficial -= EncFethab
                    End If
                Next

                grdEncargosFixacao.DataSource = objFixacao.Encargos.ToArray
                grdEncargosFixacao.DataBind()

                If FinanceiroNovo Then
                    lnkCalcularValores.Parent.Visible = False
                    lnkNovo.Parent.Visible = True

                    If objFixacao.SubOperacao.Financeiro Then
                        ngsFinanceiro.SetarHID(HID.Value)
                        ngsFinanceiro.SetarOrigem = "FX"
                        ngsFinanceiro.CarregarControle()
                    End If
                    TabContainer1.ActiveTabIndex = 3
                Else
                    Habilitar_VencimentoFixacao()
                    TabContainer1.ActiveTabIndex = 3
                End If

                If objPedido.Moeda.Classificacao = [Lib].Negocio.eTiposMoeda.Oficial Then
                    txtTotalAfixar.Text = objFixacao.TotalOficial.ToString("N2")
                    txtTotalAfixarLiquido.Text = objFixacao.Encargos.Find(Function(s) s.CodigoEncargo = "LIQUIDO").ValorOficial.ToString("N2")
                Else
                    txtTotalAfixar.Text = objFixacao.TotalMoeda.ToString("N2")
                    txtTotalAfixarLiquido.Text = objFixacao.Encargos.Find(Function(s) s.CodigoEncargo = "LIQUIDO").ValorMoeda.ToString("N2")
                End If

                SessaoSalvaFixacao()
                SessaoSalvaPedido()
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkLimpar_Click(sender As Object, e As EventArgs) Handles lnkLimpar.Click
        Try
            SetarParametros(HIDLinhaProduto.Value, HIDPedido.Value)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkNovo_Click(sender As Object, e As EventArgs) Handles lnkNovo.Click
        Try
            If Funcoes.VerificaPermissao("PedidosXItens", "GRAVAR") Then
                SessaoRecuperaFixacao()
                Dim Sqls As New ArrayList
                Dim sql As String

                SessaoRecuperaPedido()

                sql = "UPDATE Pedidos SET " & _
                      "    Observacoes = '" & objPedido.Observacoes & "'" & vbCrLf & _
                      "   ,UsuarioAlteracao     ='" & Session("ssNomeUsuario") & "'" & vbCrLf & _
                      "   ,UsuarioAlteracaoData = CURRENT_TIMESTAMP" & vbCrLf

                If Left(HttpContext.Current.Session("ssEmpresa").ToString, 8) = "03961253" Then
                    If objPedido.Empresa.Empresa.PedidoBloqueado AndAlso _
                        objPedido.SubOperacao.EntradaSaida = eEntradaSaida.Saida AndAlso _
                        (objPedido.Itens.Exists(Function(s) s.Produto.Grupo.Codigo = "10101") OrElse _
                         objPedido.Itens.Exists(Function(s) s.Produto.Grupo.Codigo = "10102") OrElse _
                         objPedido.Itens.Exists(Function(s) s.Produto.Grupo.Codigo = "10103") OrElse _
                         objPedido.Itens.Exists(Function(s) s.Produto.Grupo.Codigo = "10104") OrElse _
                         objPedido.Itens.Exists(Function(s) s.Produto.Grupo.Codigo = "10105") OrElse _
                         objPedido.Itens.Exists(Function(s) s.Produto.Grupo.Codigo = "10220") OrElse _
                         objPedido.Itens.Exists(Function(s) s.Produto.Grupo.Codigo = "10301") OrElse _
                         objPedido.Itens.Exists(Function(s) s.Produto.Grupo.Codigo = "10401") OrElse _
                         objPedido.Itens.Exists(Function(s) s.Produto.Grupo.Codigo = "20101") OrElse _
                         objPedido.Itens.Exists(Function(s) s.Produto.Grupo.Codigo = "70101")) Then
                        sql &= "   ,PedidoBloqueado  = 1 " & vbCrLf
                    End If
                ElseIf objPedido.Itens.Exists(Function(s) s.Produto.Nome.Contains("SOJA") Or s.Produto.Nome.Contains("LECITINA") And Not s.Produto.Nome.Contains("TAMBOR")) Then
                    If objPedido.Empresa.Empresa.PedidoBloqueado Then
                        sql &= "   ,PedidoBloqueado  = 1 " & vbCrLf
                    End If
                End If

                sql &= " WHERE Empresa_Id    ='" & objPedido.CodigoEmpresa & "'" & vbCrLf & _
                      "   AND EndEmpresa_Id = " & objPedido.EnderecoEmpresa & vbCrLf & _
                      "   AND Pedido_Id     = " & objPedido.Codigo
                Sqls.Add(sql)

                If FinanceiroNovo Then
                    If objFixacao.Titulos.Count = 0 Then
                        MsgBox(Me.Page, "Financeiro não programado!")
                        Exit Sub
                    End If

                    If FinanceiroNovo AndAlso objFixacao.ItemPedido.Pedido IsNot Nothing AndAlso Not objFixacao.ItemPedido.Pedido.Bloquear() Then
                        MsgBox(Me.Page, "O pedido " & objFixacao.ItemPedido.Pedido.Codigo & " foi bloqueado por outro usuário, por favor recarregue o registro!")
                        Exit Sub
                    End If
                    objFixacao.ItemPedido.Pedido.FinanceiroNovo = FinanceiroNovo
                    objFixacao.SalvarSql(Sqls)
                Else
                    objFixacao.SalvarSql(Sqls)
                    For Each tit As [Lib].Negocio.Titulo In CType(Session("VencimentosFixacao" & HID.Value), [Lib].Negocio.ListTitulo)
                        sql = "exec sp_Numerador '" & HttpContext.Current.Session("ssNomeServidor").ToUpper() & "',0,1"
                        For Each Dr As DataRow In Banco.ConsultaDataSet(sql, "Numerador").Tables(0).Rows
                            tit.Codigo = Dr("Sequencia")
                            If Not objFixacao.NotasFixacao.Sum(Function(s) s.ValorAFixarLiberado) = objFixacao.TotalOficial Then tit.CodigoProvisao = 3
                        Next
                        tit.IUD = "I"
                        tit.SalvarSql(Sqls, False)
                    Next
                End If

                If FinanceiroNovo AndAlso Not String.IsNullOrWhiteSpace(objFixacao.ItemPedido.Pedido.Empresa.Empresa.Servidor) AndAlso UsuarioServidor.NomeServidor <> objFixacao.ItemPedido.Pedido.Empresa.Empresa.Servidor Then
                    If FinanceiroNovo AndAlso objFixacao.ItemPedido.Pedido IsNot Nothing Then getSqlException(Sqls, objFixacao.ItemPedido.Pedido)
                    Dim db As New AcessaBanco(2, objFixacao.ItemPedido.Pedido.Empresa.Empresa.Servidor)
                    If db.GravaBanco(Sqls) Then
                        SetarParametros(HIDLinhaProduto.Value, HIDPedido.Value)
                        TabContainer1.ActiveTabIndex = 0
                        MsgBox(Me.Page, "Fixação gravada com Sucesso.", eTitulo.Sucess)
                        CType(Me.Page, PedidosXItens).DesabilitaAlteracao()
                    Else
                        MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                    End If
                Else
                    If Banco.GravaBanco(Sqls) Then
                        SetarParametros(HIDLinhaProduto.Value, HIDPedido.Value)
                        TabContainer1.ActiveTabIndex = 0
                        MsgBox(Me.Page, "Fixação gravada com Sucesso.", eTitulo.Sucess)
                        CType(Me.Page, PedidosXItens).DesabilitaAlteracao()
                    Else
                        MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                    End If
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para gravar fixação!")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkParcelar_Click(sender As Object, e As EventArgs) Handles lnkParcelar.Click
        Try
            If gridCondicoesFixacao.SelectedIndex = -1 Then
                MsgBox(Me.Page, "Selecione um Registro para alteração.")
            Else
                SessaoRecuperaPedido()

                Dim msg As String = ""

                txtDataVencimentoFixacao.Text = Funcoes.ValidaDataUtil(CType(Session("VencimentosFixacao" & HID.Value), [Lib].Negocio.ListTitulo)(gridCondicoesFixacao.SelectedIndex).CodigoEmpresa, CType(Session("VencimentosFixacao" & HID.Value), [Lib].Negocio.ListTitulo)(gridCondicoesFixacao.SelectedIndex).EnderecoEmpresa, CDate(txtDataVencimentoFixacao.Text)).ToShortDateString

                CType(Session("VencimentosFixacao" & HID.Value), [Lib].Negocio.ListTitulo)(gridCondicoesFixacao.SelectedIndex).Vencimento = txtDataVencimentoFixacao.Text
                CType(Session("VencimentosFixacao" & HID.Value), [Lib].Negocio.ListTitulo)(gridCondicoesFixacao.SelectedIndex).Prorrogacao = txtDataVencimentoFixacao.Text

                If objPedido.Moeda.Classificacao = [Lib].Negocio.eTiposMoeda.Oficial Then
                    CType(Session("VencimentosFixacao" & HID.Value), [Lib].Negocio.ListTitulo)(gridCondicoesFixacao.SelectedIndex).ValorDoDocumento = txtValorVencimentoFixacao.Text

                    If CType(Session("VencimentosFixacao" & HID.Value), [Lib].Negocio.ListTitulo)(gridCondicoesFixacao.SelectedIndex).ValorDoDocumento <> CDec(ValorVencimentoFixacao.Value) Then
                        msg = CType(Session("VencimentosFixacao" & HID.Value), [Lib].Negocio.ListTitulo).AjustaParcelas(gridCondicoesFixacao.SelectedIndex, ValorVencimentoFixacao.Value, CType(Session("VencimentosFixacao" & HID.Value), [Lib].Negocio.ListTitulo)(gridCondicoesFixacao.SelectedIndex).ValorDoDocumento)
                        If msg.Length > 0 Then
                            MsgBox(Me.Page, msg)
                        Else
                            AtualizarGridVencimentos()
                        End If
                    Else
                        AtualizarGridVencimentos()
                    End If
                Else
                    CType(Session("VencimentosFixacao" & HID.Value), [Lib].Negocio.ListTitulo)(gridCondicoesFixacao.SelectedIndex).MoedaValorDoDocumento = txtValorVencimentoFixacao.Text

                    If CType(Session("VencimentosFixacao" & HID.Value), [Lib].Negocio.ListTitulo)(gridCondicoesFixacao.SelectedIndex).MoedaValorDoDocumento <> CDec(ValorVencimentoFixacao.Value) Then
                        msg = CType(Session("VencimentosFixacao" & HID.Value), [Lib].Negocio.ListTitulo).AjustaParcelas(gridCondicoesFixacao.SelectedIndex, ValorVencimentoFixacao.Value, CType(Session("VencimentosFixacao" & HID.Value), [Lib].Negocio.ListTitulo)(gridCondicoesFixacao.SelectedIndex).MoedaValorDoDocumento)
                        If msg.Length > 0 Then
                            MsgBox(Me.Page, msg)
                        Else
                            AtualizarGridVencimentos()
                        End If
                    Else
                        AtualizarGridVencimentos()
                    End If
                End If
            End If

            txtDataVencimentoFixacao.Text = ""
            txtValorVencimentoFixacao.Text = ""
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub LnkSair_Click(sender As Object, e As EventArgs) Handles LnkSair.Click
        Try
            If TypeOf Me.Page Is PedidosXItens Then
                CType(Me.Page, PedidosXItens).AtualizarObservacao()
            End If
            Popup.CloseDialog(Me.Page, "divPedidoFixacao")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub txtDataAfixar_TextChanged(sender As Object, e As EventArgs) Handles txtDataAfixar.TextChanged
        Try
            SessaoRecuperaPedido()
            If IsDate(txtDataAfixar.Text) Then
                If Not objPedido.CodigoIndexador = 99 Then
                    txtIndiceFixadoItem.Text = Funcoes.PegarValorConversao(objPedido.CodigoIndexador, CDate(txtDataAfixar.Text)).ToString("N8")

                Else
                    txtIndiceFixadoItem.Text = "0,00000000"
                End If
            Else
                txtDataAfixar.Text = Today.ToString("dd/MM/yyyy")
                If Not objPedido.CodigoIndexador = 99 Then
                    txtIndiceFixadoItem.Text = Funcoes.PegarValorConversao(objPedido.CodigoIndexador, Today).ToString("N8")
                Else
                    txtIndiceFixadoItem.Text = "0,00000000"
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAdicionarObservacaoFixacao_Click(sender As Object, e As EventArgs) Handles lnkAdicionarObservacaoFixacao.Click
        Try
            SessaoRecuperaPedido()
            objPedido.Observacoes = Date.Now.ToString("dd/MM/yyyy HH:mm:ss") & " - Usuario: " & HttpContext.Current.Session("ssNomeUsuario") & vbCrLf & "       " & Funcoes.EliminarCaracteresEspeciais(txtAddObservacaoFixacao.Text) & IIf(objPedido.Observacoes.Length > 0, vbCrLf & objPedido.Observacoes, "")
            txtObservacaoFixacao.Text = objPedido.Observacoes
            SessaoSalvaPedido()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub imgCessaoDeCredito_Click(sender As Object, e As System.Web.UI.ImageClickEventArgs) Handles imgCessaoDeCredito.Click
        Try
            Dim ucFixacaoProcuracao = CType(Me.Page.FindControlRecursive("ucFixacaoProcuracao"), ucFixacaoProcuracao)
            If ucFixacaoProcuracao IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(HIDPedido.Value) Then
                ucFixacaoProcuracao.SetarHID(HID.Value)
                ucFixacaoProcuracao.CarregarProcuracoes(HIDLinhaProduto.Value)
                ucFixacaoProcuracao.MainUserControl = Me
                Popup.ConsultaFixacaoProcuracao(Me.Page, "objFixacaoProcuracao" & HID.Value)
            Else
                MsgBox(Me.Page, "Não encontrado Controle ucFixacaoProcuracao. Entre em contato com o suporte.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

#End Region

#Region "Methods"

    Public Overrides Sub SetarHID(ByVal guid As String)
        HID.Value = guid.ToString
    End Sub

    Public Sub SetarParametros(LinhaProdutoPedido As Integer, ByVal pedido As Integer)
        Limpar()
        HIDLinhaProduto.Value = LinhaProdutoPedido
        HIDPedido.Value = pedido
        SessaoRecuperaPedido()
        'Forca a classe a consultar novamente a lista no BD
        objPedido.Itens(HIDLinhaProduto.Value).Fixacoes = Nothing

        lblProdutoAfixar.Text = objPedido.Itens(LinhaProdutoPedido).CodigoProduto & " - " & objPedido.Itens(LinhaProdutoPedido).Descricao

        ddlSubAFixar.Items.Clear()
        ddl.Carregar(ddlSubAFixar, CarregarDDL.Tabela.OperacaoSubOperacao, "So.Operacao_Id = " & objPedido.CodigoOperacao & " and So.Classe = '" & eClassesOperacoes.COMPLEMENTACOES.ToString & "' and So.Liminar =" & IIf(objPedido.SubOperacao.Liminar, 1, 0), False)

        If Funcoes.VerificaPermissao("LIBERARDATAFIXACAO", "ACESSAR") Then txtDataAfixar.Enabled = True

        'Habilita campos
        CamposFixacao(True)

        gridFixacao.DataSource = objPedido.Itens(HIDLinhaProduto.Value).Fixacoes.ToArray
        gridFixacao.DataBind()

        objFixacao = New Fixacao(objPedido.Itens(LinhaProdutoPedido))
        objFixacao.NotasFixacao.CarregarNotasParaSelecao()

        If objPedido.Itens(HIDLinhaProduto.Value).Fixacoes.Count = 0 Then
            objFixacao.Codigo = 1
            lblFixacao.Text = objFixacao.Codigo
        Else
            objFixacao.Codigo = objPedido.Itens(HIDLinhaProduto.Value).Fixacoes.Max(Function(s) s.Codigo) + 1
            lblFixacao.Text = objFixacao.Codigo
        End If

        txtSaldoAfixar.Text = objFixacao.NotasFixacao.Sum(Function(s) s.QtdeAFixar).ToString("N4")
        lnkCalcularValores.Parent.Visible = objFixacao.NotasFixacao.Sum(Function(s) s.QtdeAFixar) > 0

        Dim sdoProcuracao As Decimal = 0

        If objPedido.Itens(HIDLinhaProduto.Value).Fixacoes.SaldoProcuracao > 0 Then
            sdoProcuracao = objPedido.Itens(HIDLinhaProduto.Value).Fixacoes.SaldoProcuracao
            imgCessaoDeCredito.Visible = True
        End If

        txtSaldoProcuracao.Text = sdoProcuracao

        txtObservacaoFixacao.Text = objPedido.Observacoes
        SessaoSalvaFixacao()
    End Sub

    Public Overrides Sub Carregar(ByVal procuracao As String, ByVal saldo As Decimal)
        If Not Session("objFixacaoProcuracao" & HID.Value) Is Nothing Then
            txtCessaoDeCreditoAfixar.Text = procuracao
            txtSaldoProcuracao.Text = saldo
            Session.Remove("objFixacaoProcuracao" & HID.Value)
        End If
    End Sub

    Public Overrides Sub Limpar()
        SessaoRecuperaPedido()
        lnkNovo.Parent.Visible = False

        'Limpar Fixacao
        If ddlSubAFixar.Items.Count > 0 Then
            ddlSubAFixar.SelectedIndex = 0
        End If

        txtDataAfixar.Text = Today.ToString("dd/MM/yyyy")

        If Not objPedido.CodigoIndexador = 99 Then
            txtIndiceFixadoItem.Text = Funcoes.PegarValorConversao(objPedido.CodigoIndexador, Today).ToString("N8")
        Else
            txtIndiceFixadoItem.Text = New Decimal().ToString("N8")
        End If

        lblFixacao.Text = ""
        txtCessaoDeCreditoAfixar.Text = 0
        txtSaldoProcuracao.Text = 0
        txtQuantidadeAFixar.Text = New Decimal().ToString("N4")
        txtUnitarioAfixar.Text = New Decimal().ToString("N10")
        txtTotalAfixar.Text = ""
        txtTotalAfixarLiquido.Text = ""
        Session.Remove("Fixacoes" & HID.Value)
        Session.Remove("FixacaoXNotaFiscal" & HID.Value)

        gridFixacaoXNotaFiscal.DataSource = Nothing
        gridFixacaoXNotaFiscal.DataBind()

        grdEncargosFixacao.DataSource = Nothing
        grdEncargosFixacao.DataBind()

        'Limpar Vencimentos
        divVencimento.Visible = True

        Session.Remove("VencimentosFixacao" & HID.Value)

        If ddlCondicoesFixacao.Items.Count > 0 Then
            ddlCondicoesFixacao.SelectedIndex = 0
        End If

        ddlCondicoesFixacao.Enabled = False
        txtDataVencimentoFixacao.Text = ""
        txtValorVencimentoFixacao.Text = ""
        txtTotalPaceladoFixacao.Text = ""
        txtTotalPagoFixacao.Text = ""
        txtSaldoVencimentosFixacao.Text = ""
        ValorVencimentoFixacao.Value = 0

        txtDataVencimentoFixacao.Enabled = False
        txtValorVencimentoFixacao.Enabled = False
        txtTotalPaceladoFixacao.Enabled = False
        txtTotalPagoFixacao.Enabled = False
        txtSaldoVencimentosFixacao.Enabled = False

        gridCondicoesFixacao.DataSource = Nothing
        gridCondicoesFixacao.DataBind()

        If FinanceiroNovo Then
            ngsFinanceiro.Limpar()
        End If

        'Setar Aba
        TabContainer1.ActiveTabIndex = 0

        SessaoRecuperaPedido()
        txtObservacaoFixacao.Text = objPedido.Observacoes
    End Sub

    Private Sub AtualizarGridVencimentos()
        Dim intParcela As Integer = 1
        Dim TotalPaceladoFixacao As Decimal
        Dim dtVencimentos As New DataTable()
        dtVencimentos.Columns.Add("Codigo", Type.GetType("System.Int32"))
        dtVencimentos.Columns.Add("TipoPagamento", Type.GetType("System.Int32"))
        dtVencimentos.Columns.Add("Sequencia", Type.GetType("System.Int32"))
        dtVencimentos.Columns.Add("Parcela", Type.GetType("System.Int32"))
        dtVencimentos.Columns.Add("Vencimento", Type.GetType("System.DateTime"))
        dtVencimentos.Columns.Add("Valor", Type.GetType("System.Double"))
        dtVencimentos.Columns.Add("ValorMoeda", Type.GetType("System.Double"))
        dtVencimentos.Columns.Add("DataPagamento", Type.GetType("System.DateTime"))
        dtVencimentos.Columns.Add("ValorPagamento", Type.GetType("System.Double"))
        dtVencimentos.Columns.Add("ValorPagamentoMoeda", Type.GetType("System.Double"))
        dtVencimentos.Columns.Add("Saldo", Type.GetType("System.Double"))
        dtVencimentos.Columns.Add("SaldoMoeda", Type.GetType("System.Double"))
        dtVencimentos.Columns.Add("Atraso", Type.GetType("System.Double"))

        For Each tit As [Lib].Negocio.Titulo In CType(Session("VencimentosFixacao" & HID.Value), [Lib].Negocio.ListTitulo)
            Dim drVencimento As DataRow = dtVencimentos.NewRow()

            drVencimento("Codigo") = tit.Codigo
            drVencimento("TipoPagamento") = tit.CodigoTipoPgto
            drVencimento("Sequencia") = 0
            drVencimento("Parcela") = intParcela
            drVencimento("Vencimento") = tit.Prorrogacao
            drVencimento("Valor") = tit.ValorDoDocumento
            drVencimento("ValorMoeda") = tit.MoedaValorDoDocumento
            drVencimento("DataPagamento") = DBNull.Value
            drVencimento("ValorPagamento") = 0
            drVencimento("ValorPagamentoMoeda") = 0
            drVencimento("Saldo") = 0
            drVencimento("SaldoMoeda") = 0
            drVencimento("Atraso") = 0

            dtVencimentos.Rows.Add(drVencimento)

            TotalPaceladoFixacao += tit.ValorDoDocumento

            intParcela += 1
        Next

        gridCondicoesFixacao.DataSource = dtVencimentos
        gridCondicoesFixacao.DataBind()

        txtTotalPaceladoFixacao.Text = TotalPaceladoFixacao.ToString("N2")

        If gridCondicoesFixacao.Rows.Count > 0 Then lnkNovo.Parent.Visible = True
    End Sub

    Public Sub getSqlException(ByRef Sqls As ArrayList, ByVal objPedido As [Lib].Negocio.Pedido)
        Dim sql As String = "BEGIN TRY " & vbCrLf & _
               "DECLARE @HORA_BLOQUEIO AS DATETIME = DATEADD(MINUTE, 3, (SELECT VERSAOHORARIOBLOQUEIO FROM PEDIDOS WHERE PEDIDO_ID = '" & objPedido.Codigo & "' AND EMPRESA_ID = '" & objPedido.CodigoEmpresa & "' AND ENDEMPRESA_ID = '" & objPedido.EnderecoEmpresa & "')) " & vbCrLf & _
               "PRINT 'HORA_ATUAL: ' + CAST(GETDATE() AS VARCHAR); " & vbCrLf & _
               "PRINT 'HORA_BLOQUEIO: ' + CAST(@HORA_BLOQUEIO AS VARCHAR); " & vbCrLf & _
               "IF (GETDATE() > @HORA_BLOQUEIO) " & vbCrLf & _
               "BEGIN " & vbCrLf & _
               "RAISERROR ('POR FAVOR, ATUALIZE O SALDO FINANCEIRO PARA REALIZAR ESTA AÇÃO!', " & vbCrLf & _
               "16, " & vbCrLf & _
               "1); " & vbCrLf & _
               "END " & vbCrLf & _
               "UPDATE PEDIDOS SET VERSAOHORARIOBLOQUEIO = NULL WHERE PEDIDO_ID = '" & objPedido.Codigo & "' AND EMPRESA_ID = '" & objPedido.CodigoEmpresa & "' AND ENDEMPRESA_ID = '" & objPedido.EnderecoEmpresa & "'; " & vbCrLf & _
               "END TRY " & vbCrLf & _
               "BEGIN CATCH " & vbCrLf & _
               "DECLARE @ErrorMessage NVARCHAR(4000); " & vbCrLf & _
               "DECLARE @ErrorSeverity INT; " & vbCrLf & _
               "DECLARE @ErrorState INT; " & vbCrLf & _
               "SELECT " & vbCrLf & _
               "@ErrorMessage = ERROR_MESSAGE(), " & vbCrLf & _
               "@ErrorSeverity = ERROR_SEVERITY(), " & vbCrLf & _
               "@ErrorState = ERROR_STATE(); " & vbCrLf & _
               "RAISERROR (@ErrorMessage, " & vbCrLf & _
               "@ErrorSeverity, " & vbCrLf & _
               "@ErrorState); " & vbCrLf & _
               "END CATCH;"
        Sqls.Add(sql)
    End Sub

    Public Sub CamposFixacao(Habilitado As Boolean)
        ddlSubAFixar.Enabled = Habilitado
        txtIndiceFixadoItem.Enabled = Habilitado
        txtQuantidadeAFixar.Enabled = Habilitado
        txtUnitarioAfixar.Enabled = Habilitado
        txtDataAfixar.Enabled = Habilitado
    End Sub

    Private Sub Habilitar_VencimentoFixacao()
        lnkCalcularValores.Parent.Visible = False
        lnkNovo.Parent.Visible = True
        divVencimento.Visible = True

        ddlCondicoesFixacao.Enabled = True
        txtDataVencimentoFixacao.Text = ""
        txtValorVencimentoFixacao.Text = ""
        txtTotalPaceladoFixacao.Text = ""
        txtTotalPagoFixacao.Text = ""
        txtSaldoVencimentosFixacao.Text = ""

        txtDataVencimentoFixacao.Enabled = True
        txtValorVencimentoFixacao.Enabled = True

        txtTotalPaceladoFixacao.Enabled = True
        txtTotalPagoFixacao.Enabled = True
        txtSaldoVencimentosFixacao.Enabled = True

        gridCondicoesFixacao.DataSource = Nothing
        gridCondicoesFixacao.DataBind()
    End Sub

#End Region

End Class