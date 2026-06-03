Imports System.Data
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class AlterarNotaFiscal
    Inherits BasePage

    Dim Sql As String
    Private objNotaFiscal As [Lib].Negocio.NotaFiscal

#Region "Load"
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.setMenu(eModulo.Expedicao)
            If Not IsPostBack And IsConnect Then
                If Funcoes.VerificaPermissao("AlterarNotaFiscal", "ACESSAR") Then
                    ddl.Carregar(ddlUnidadeDeNegocio, CarregarDDL.Tabela.UnidadeDeNegocio, "", True)
                    Funcoes.VerificaUnidadeDeNegocio(ddlUnidadeDeNegocio, ddlEmpresa)
                    BuscarFinalidades()
                    txtDataInicial.Text = Format(Today, "dd/MM/yyyy")
                    txtDataFinal.Text = Format(Today, "dd/MM/yyyy")
                    LimparCampos()
                Else
                    MsgBox(Me.Page, "Usuário sem permissão para acessar essa página!", "~/Expedicao.aspx")
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub
#End Region

#Region "Session"
    Private Sub SalvaNotaFiscal()
        Session("objNotaFiscal" & HID.Value) = objNotaFiscal
    End Sub

    Private Sub RecuperaNotaFiscal()
        objNotaFiscal = CType(Session("objNotaFiscal" & HID.Value), [Lib].Negocio.NotaFiscal)
    End Sub
#End Region

#Region "Carregar"
    Public Overrides Sub Carregar(obj As IBaseEntity)
        If Not Session("objClienteOPxNOTA" & HID.Value) Is Nothing Then
            Dim itemCliente As ListItem = Funcoes.FormatarListItemCliente(CType(Session("objClienteOPxNOTA" & HID.Value), [Lib].Negocio.Cliente))
            txtCliente.Text = itemCliente.Text
            txtCodigoCliente.Value = itemCliente.Value
            Session.Remove("objClienteOPxNOTA" & HID.Value)
        End If
    End Sub

    Public Sub CarregarNotaFiscal()
        Try
            If Not Session("objNFConsultaNXI" & HID.Value) Is Nothing Then

                objNotaFiscal = New [Lib].Negocio.NotaFiscal(CType(Session("objNFConsultaNXI" & HID.Value), NotaFiscal))
                objNotaFiscal.NotaFiscalOriginal = New NotaFiscal(objNotaFiscal)

                If objNotaFiscal.CodigoTipoDeDocumento = CInt(eTipoDeDocumento.CTRC) Or
                    objNotaFiscal.CodigoTipoDeDocumento = CInt(eTipoDeDocumento.ReciboDeFrete) Or
                    objNotaFiscal.CodigoTipoDeDocumento = CInt(eTipoDeDocumento.CTRC_SEM_NF) Or
                    objNotaFiscal.CodigoTipoDeDocumento = CInt(eTipoDeDocumento.CT_E_TOM) Or
                    objNotaFiscal.CodigoTipoDeDocumento = CInt(eTipoDeDocumento.CT_E) Or
                    objNotaFiscal.CodigoTipoDeDocumento = CInt(eTipoDeDocumento.CT_E_OUT) Or
                    objNotaFiscal.CodigoTipoDeDocumento = CInt(eTipoDeDocumento.Estadia) Or
                    objNotaFiscal.CodigoTipoDeDocumento = CInt(eTipoDeDocumento.Anulacao) Then

                    MsgBox(Me.Page, "Documento de Frete não pode ser alterado por aqui.")
                    Exit Sub
                End If

                txtNota.Text = objNotaFiscal.Codigo
                txtES.Text = IIf(objNotaFiscal.EntradaSaida = 0, "E", "S")
                txtSerie.Text = objNotaFiscal.Serie
                txtNota.Enabled = False
                txtES.Enabled = False
                txtSerie.Enabled = False

                objNotaFiscal.IUD = "U"
                objNotaFiscal.CarregandoNota = True

                If (objNotaFiscal.CFOP.Codigo <> 5353 And objNotaFiscal.CFOP.Codigo <> 6353) Then
                    Dim Parametros As New Hashtable
                    Parametros.Add("TipoApuracao", 2)
                    Parametros.Add("Empresa", objNotaFiscal.CodigoEmpresa)
                    Parametros.Add("EndEmpresa", objNotaFiscal.EnderecoEmpresa)
                    Parametros.Add("Cliente", objNotaFiscal.CodigoCliente)
                    Parametros.Add("EndCliente", objNotaFiscal.EnderecoCliente)
                    Parametros.Add("Pedido", objNotaFiscal.CodigoPedido)
                    Parametros.Add("Devolucao", IIf(objNotaFiscal.SubOperacao.Devolucao, 1, 0))
                    Parametros.Add("SoValor", IIf(Not objNotaFiscal.SubOperacao.QuantidadeFiscal And Not objNotaFiscal.SubOperacao.QuantidadeFisico, 1, 0))
                    Parametros.Add("Classe", objNotaFiscal.Operacao.CodigoClasse)

                    Dim ListaItensPedido As New ListSaldoPedido2015(Parametros)

                    Dim ObjItemNF As [Lib].Negocio.NotaFiscalXItem
                    For Each row As SaldoPedido2015 In ListaItensPedido
                        Dim codigoproduto As Integer = row.CodigoProduto
                        ObjItemNF = objNotaFiscal.Itens.Where(Function(s) s.CodigoProduto = codigoproduto).FirstOrDefault
                        If ObjItemNF IsNot Nothing Then
                            'GLOBAL É DEVOLUCAO
                            If objNotaFiscal.SubOperacao.Classe = eClassesOperacoes.GLOBAL And objNotaFiscal.SubOperacao.Devolucao Then
                                ObjItemNF.SaldoPedidoFiscal = row.QtdeEntregueFiscalGlobal - row.QtdeEntregueFiscalRemessa + ObjItemNF.QuantidadeFiscal
                                ObjItemNF.SaldoPedidoFisico = row.QtdeEntregueFiscalGlobal - row.QtdeEntregueFiscalRemessa + ObjItemNF.QuantidadeFiscal
                                ObjItemNF.SaldoValorOficial = row.VlrNotaOficialGlobalBruto - row.VlrNotaOficialRemessaBruto + ObjItemNF.ValorTotal
                                ObjItemNF.SaldoValorMoeda = row.VlrNotaMoedaGlobalBruto - row.VlrNotaMoedaRemessaBruto + ObjItemNF.ValorTotalMoeda
                                'GLOBAL NAO É DEVOLUCAO
                            ElseIf objNotaFiscal.SubOperacao.Classe = eClassesOperacoes.GLOBAL And Not objNotaFiscal.SubOperacao.Devolucao Then
                                ObjItemNF.SaldoPedidoFiscal = row.SaldoQtdeGlobalFiscal + ObjItemNF.QuantidadeFiscal
                                ObjItemNF.SaldoPedidoFisico = row.SaldoQtdeGlobalFiscal + ObjItemNF.QuantidadeFiscal
                                ObjItemNF.SaldoValorOficial = row.SaldoValorOficialGlobalDireto + ObjItemNF.ValorTotal
                                ObjItemNF.SaldoValorMoeda = row.SaldoValorMoedaGlobalDireto + ObjItemNF.ValorTotalMoeda
                                'REMESSA É DEVOLUCAO
                            ElseIf objNotaFiscal.SubOperacao.Classe = eClassesOperacoes.REMESSAS And objNotaFiscal.SubOperacao.Devolucao Then
                                ObjItemNF.SaldoPedidoFiscal = row.QtdeEntregueFiscalRemessa + ObjItemNF.QuantidadeFiscal
                                ObjItemNF.SaldoPedidoFisico = row.QtdeEntregueFisicoRemessa + ObjItemNF.QuantidadeFisica
                                ObjItemNF.SaldoValorOficial = row.VlrNotaOficialRemessaBruto + ObjItemNF.ValorTotal
                                ObjItemNF.SaldoValorMoeda = row.VlrNotaMoedaRemessaBruto + ObjItemNF.ValorTotalMoeda
                                'REMESSA NAO É DEVOLUCAO
                            ElseIf objNotaFiscal.SubOperacao.Classe = eClassesOperacoes.REMESSAS And Not objNotaFiscal.SubOperacao.Devolucao Then
                                ObjItemNF.SaldoPedidoFiscal = row.SaldoQtdeRemessaFiscal + ObjItemNF.QuantidadeFiscal
                                ObjItemNF.SaldoPedidoFisico = row.SaldoQtdeRemessaFisica + ObjItemNF.QuantidadeFisica
                                ObjItemNF.SaldoValorOficial = row.SaldoValorOficialRemessa + ObjItemNF.ValorTotal
                                ObjItemNF.SaldoValorMoeda = row.SaldoValorMoedaRemessa + ObjItemNF.ValorTotalMoeda
                                'AFIXAR É DEVOLUCAO
                            ElseIf objNotaFiscal.SubOperacao.Classe = eClassesOperacoes.AFIXAR And objNotaFiscal.SubOperacao.Devolucao Then
                                ObjItemNF.SaldoPedidoFiscal = row.QtdeEntregueFiscalAFixar - row.QtdeFixacao + ObjItemNF.QuantidadeFiscal
                                ObjItemNF.SaldoPedidoFisico = row.QtdeEntregueFisicoAFixar - row.QtdeFixacao + ObjItemNF.QuantidadeFisica
                                ObjItemNF.SaldoValorOficial = row.VlrNotaOficialAFixarBruto - row.VlrFixacaoOficial + ObjItemNF.ValorTotal
                                ObjItemNF.SaldoValorMoeda = row.VlrNotaMoedaAFixarBruto - row.VlrFixacaoMoeda + ObjItemNF.ValorTotalMoeda
                            ElseIf objNotaFiscal.SubOperacao.Devolucao And row.Tipo = 1 Then
                                ObjItemNF.SaldoPedidoFiscal = row.QtdeEntregueFiscalDireta + ObjItemNF.QuantidadeFiscal
                                ObjItemNF.SaldoPedidoFisico = row.QtdeEntregueFisicoDireta + ObjItemNF.QuantidadeFisica
                                ObjItemNF.SaldoValorOficial = row.VlrNotaOficialDiretaBruto + ObjItemNF.ValorTotal
                                ObjItemNF.SaldoValorMoeda = row.VlrNotaMoedaDiretaBruto + ObjItemNF.ValorTotalMoeda
                            ElseIf objNotaFiscal.SubOperacao.Devolucao And row.Tipo = 2 Then
                                ObjItemNF.SaldoPedidoFiscal = row.QtdeEntregueFiscalDeposito + ObjItemNF.QuantidadeFiscal
                                ObjItemNF.SaldoPedidoFisico = row.QtdeEntregueFisicoDeposito + ObjItemNF.QuantidadeFisica
                                ObjItemNF.SaldoValorOficial = row.VlrNotaOficialDepositoBruto + ObjItemNF.ValorTotal
                                ObjItemNF.SaldoValorMoeda = row.VlrNotaMoedaDepositoBruto + ObjItemNF.ValorTotalMoeda
                            Else
                                'Verificar ainda pode nao estar contemplando tudo
                                If row.QtdeProgramada = 0 And row.QtdeProgramadaComercializacao > 0 Then
                                    ObjItemNF.SaldoPedidoFiscal = row.SaldoQtdeComercializacao + ObjItemNF.QuantidadeFiscal
                                    ObjItemNF.SaldoPedidoFisico = row.SaldoQtdeComercializacao + ObjItemNF.QuantidadeFisica
                                    ObjItemNF.SaldoValorOficial = row.SaldoValorOficialGlobalDireto + ObjItemNF.ValorTotal
                                    ObjItemNF.SaldoValorMoeda = row.SaldoValorMoedaGlobalDireto + ObjItemNF.ValorTotalMoeda
                                Else
                                    ObjItemNF.SaldoPedidoFiscal = row.SaldoQtdeDiretoFiscal + ObjItemNF.QuantidadeFiscal
                                    ObjItemNF.SaldoPedidoFisico = row.SaldoQtdeDiretoFisica + ObjItemNF.QuantidadeFisica
                                    ObjItemNF.SaldoValorOficial = row.SaldoValorOficialGlobalDireto + ObjItemNF.ValorTotal
                                    ObjItemNF.SaldoValorMoeda = row.SaldoValorMoedaGlobalDireto + ObjItemNF.ValorTotalMoeda
                                End If
                            End If

                            If ObjItemNF.NotasDevolucao Is Nothing OrElse ObjItemNF.NotasDevolucao.Count = 0 Then
                                ObjItemNF.NotasDevolucao.CarregarNotasUsadasNaDevolucao()
                            End If
                        End If
                    Next
                End If

                ddl.Carregar(cmbSubOperacao, CarregarDDL.Tabela.OperacaoSubOperacao)
                cmbSubOperacao.SelectedValue = objNotaFiscal.CodigoOperacao & "-" & objNotaFiscal.CodigoSubOperacao

                SelecionarIndiceCombo(cmbFinalidade, objNotaFiscal.CodigoFinalidade)

                SalvaNotaFiscal()

                txtPedido.Text = objNotaFiscal.CodigoPedido
                txtRomaneio.Text = objNotaFiscal.CodigoRomaneio

                If objNotaFiscal.CodigoRomaneio > 0 Then
                    lblFisico.Text = objNotaFiscal.Romaneio.PesoLiquido
                    If Not objNotaFiscal.Romaneio.Processo = "NOTA FISCAL" Then
                        btnRomaneio.Enabled = False
                        If objNotaFiscal.Romaneio.Pesagens.Count > 0 Then
                            txtPesagem.Text = objNotaFiscal.Romaneio.Pesagens(0).CodigoPesagem
                        End If
                    End If
                Else
                    lblFisico.Text = "0"
                End If

                lblTotalProduto.Text = objNotaFiscal.TotalProduto.ToString("N2")
                lblTotalNota.Text = objNotaFiscal.TotalNota.ToString("N2")

                pnlDados.Visible = True
                pnlItem.Visible = True

                If Not Funcoes.VerificaAcesso(objNotaFiscal.Empresa.Codigo, objNotaFiscal.Empresa.CodigoEndereco, objNotaFiscal.Movimento, "NOTAS FISCAIS") Then
                    MsgBox(Me.Page, "Movimento " & objNotaFiscal.Movimento.ToString("dd-MM-yyyy") & " da Nota Fiscal já Fechado para esta data.")
                ElseIf Not Funcoes.VerificaAcesso(objNotaFiscal.Empresa.Codigo, objNotaFiscal.Empresa.CodigoEndereco, objNotaFiscal.Movimento, "CONTABIL") Then
                    MsgBox(Me.Page, "Movimento " & objNotaFiscal.Movimento.ToString("dd-MM-yyyy") & " Contábil já Fechado para esta data.")

                Else
                    lnkAtualizar.Parent.Visible = True
                    imgExtratoPedido.Enabled = True
                    lnkConsultar.Parent.Visible = False

                    Dim temDevolucao As Boolean = False

                    For Each ni In objNotaFiscal.Itens
                        If Not ni.NotasDevolucao Is Nothing AndAlso ni.NotasDevolucao.Count > 0 Then temDevolucao = True
                    Next

                    If temDevolucao Then
                        If Funcoes.VerificaPermissao("AlterarNotaFiscal", "LIBERAR") Then
                            ''LIBERA ALTERAÇÃO - FURLAN - 01-11-2024
                        Else
                            MsgBox(Me.Page, "Não é possível alterar nota fiscal com NOTAS DE DEVOLUÇÃO.")
                            lnkAtualizar.Parent.Visible = False
                            Exit Sub
                        End If
                    End If

                    BtnVencimentos.Visible = False

                    If Not IsNothing(objNotaFiscal.VencimentosNota) AndAlso objNotaFiscal.VencimentosNota.Count > 0 Then

                        Dim titulos As String = ""
                        For Each vencimento As Titulo In objNotaFiscal.VencimentosNota
                            titulos &= vencimento.Codigo & ", "
                        Next
                        txtTitulo.Text = titulos.Substring(0, titulos.Length - 2)

                        If objNotaFiscal.VencimentosNota.Where(Function(s) s.CodigoProvisao = eProvisao.Baixa).ToList.Count > 0 Then
                            lnkAtualizar.Parent.Visible = False
                            MsgBox(Me.Page, "Nota fiscal de " & objNotaFiscal.EntradaSaida.ToString() & " com vencimento baixado não pode ser alterada nem excluída!")
                        ElseIf objNotaFiscal.VencimentosNota.Where(Function(s) s.CodigoSituacao = eSituacao.RemessaBancaria).ToList.Count > 0 Then
                            lnkAtualizar.Parent.Visible = False
                            MsgBox(Me.Page, "Não é possível alterar a nota fiscal " & objNotaFiscal.Codigo & ", o financeiro está em Cobrança Bancária.")
                        ElseIf objNotaFiscal.VencimentosNota.Where(Function(s) s.CodigoSituacao = eSituacao.EndossoTitulo).ToList.Count > 0 Then
                            lnkAtualizar.Parent.Visible = False
                            MsgBox(Me.Page, "Não é possível alterar a nota fiscal " & objNotaFiscal.Codigo & ", o financeiro está com Endosso.")
                        ElseIf objNotaFiscal.VencimentosNota.Where(Function(s) s.CodigoSituacao = eSituacao.Normal And s.RegistroMestre > 0).ToList.Count > 0 Then
                            lnkAtualizar.Parent.Visible = False
                            MsgBox(Me.Page, "Nota fiscal de " & objNotaFiscal.EntradaSaida.ToString() & " com vencimento agrupado não pode ser alterada nem excluída!")
                        End If

                    ElseIf objNotaFiscal.SubOperacao.Financeiro Then
                        BtnVencimentos.Visible = True
                        lnkAtualizar.Parent.Visible = False
                    End If

                    Session.Remove("objNFConsultaNXI" & HID.Value)

                    carregarItem()
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Public Sub SalvarVencimentos()
        RecuperaNotaFiscal()

        Dim Sqls As New ArrayList

        If objNotaFiscal.VencimentosNota IsNot Nothing AndAlso objNotaFiscal.VencimentosNota.Count > 0 Then

            objNotaFiscal.IUD = "I"

            Dim parcela As Integer = 0
            Dim parcelas As Integer = objNotaFiscal.VencimentosNota.Count
            For Each tit In objNotaFiscal.VencimentosNota
                parcela += 1
                tit.Historico = "REF. NF " & objNotaFiscal.Codigo & "-" & objNotaFiscal.Serie & "-" & IIf(objNotaFiscal.EntradaSaida = eEntradaSaida.Saida, "S", "E") & ", Parcela " & parcela & "/" & parcelas & ", Pedido: " & objNotaFiscal.CodigoPedido & " / " & objNotaFiscal.Cliente.Nome
            Next

            objNotaFiscal.VencimentosNota.SalvarSQL(Sqls, False)
        End If

        If objNotaFiscal.VencimentosPedido IsNot Nothing AndAlso objNotaFiscal.VencimentosNota.Count > 0 Then
            objNotaFiscal.VencimentosPedido.SalvarSQL(Sqls)
        End If

        If Banco.GravaBanco(Sqls) Then
            Dim espelho As New NotaFiscalEspelho
            espelho.ExibirEspelho(Me.Page, objNotaFiscal)
            MsgBox(Me.Page, "Registro alterado com suscesso")
            LimparCampos()
        Else
            MsgBox(Me.Page, "Erro ao Salvar a Nota: " & Funcoes.EliminarCaracteresEspeciais(RTrim(HttpContext.Current.Session("ssMessage").ToString)) & ". Entre em contato com o Suporte de TI")
        End If

    End Sub

    Public Sub CarregarClassificacao()
        Try
            If Not Session("objClassificacaoALT" & HID.Value.ToString) Is Nothing Then
                RecuperaNotaFiscal()

                If objNotaFiscal.Romaneio.PesoLiquido > objNotaFiscal.Itens(0).SaldoPedidoFisico And objNotaFiscal.SubOperacao.Classe <> eClassesOperacoes.DEPOSITOS And objNotaFiscal.SubOperacao.Classe <> eClassesOperacoes.AFIXAR Then
                    MsgBox(Me.Page, "O Pedido nao tem saldo suficiente para o Romaneio Selecionado. " & objNotaFiscal.Itens(0).SaldoPedidoFisico)
                    LimparCampos()
                ElseIf objNotaFiscal.Romaneio.PesoLiquido > objNotaFiscal.Itens(0).SaldoPedidoFisico And objNotaFiscal.SubOperacao.EntradaSaida = eEntradaSaida.Saida And objNotaFiscal.SubOperacao.Devolucao = True Then
                    MsgBox(Me.Page, "O Pedido nao tem saldo suficiente para o Romaneio Selecionado. " & objNotaFiscal.Itens(0).SaldoPedidoFisico)
                    LimparCampos()
                Else
                    lblNovoFisico.Text = objNotaFiscal.Romaneio.PesoLiquido
                    objNotaFiscal.Itens(0).QuantidadeFisica = objNotaFiscal.Romaneio.PesoLiquido
                    objNotaFiscal.Itens(0).PesoBruto = objNotaFiscal.Romaneio.PesoBruto
                    objNotaFiscal.Itens(0).PesoLiquido = objNotaFiscal.Romaneio.PesoLiquido
                    objNotaFiscal.Itens(0).PesoFiscal = objNotaFiscal.Romaneio.PesoBruto
                    objNotaFiscal.Itens(0).SaldoValorOficial = objNotaFiscal.Itens(0).SaldoPedidoFiscal
                    '#FimBaseDeCalculo
                    'objNotaFiscal.Itens(0).ValorTotal = (objNotaFiscal.Itens(0).QuantidadeFiscal * objNotaFiscal.Itens(0).Unitario) / objNotaFiscal.Itens(0).Produto.BaseCalculo
                    If objNotaFiscal.Itens(0).ValorTotal <> Math.Round(objNotaFiscal.Itens(0).QuantidadeFiscal * objNotaFiscal.Itens(0).Unitario, 2, MidpointRounding.AwayFromZero) Then
                        objNotaFiscal.Itens(0).ValorTotal = (objNotaFiscal.Itens(0).QuantidadeFiscal * objNotaFiscal.Itens(0).Unitario)
                    End If

                    If objNotaFiscal.CriarRomaneio Then
                        objNotaFiscal.CodigoRomaneio = 0
                    Else
                        objNotaFiscal.CodigoRomaneio = objNotaFiscal.Romaneio.Codigo
                    End If

                    SalvaNotaFiscal()

                    Session.Remove("objClassificacaoALT" & HID.Value.ToString)
                    Session.Remove("ProcurandoClassificacao" & HID.Value.ToString)

                    carregarItem()
                End If

            ElseIf Not Session("ProcurandoClassificacao" & HID.Value.ToString) Is Nothing Then
                RecuperaNotaFiscal()
                Session.Remove("ProcurandoClassificacao" & HID.Value.ToString)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Public Sub CarregarRomaneio()
        Try
            If Not Session("objRomaneioALT" & HID.Value.ToString) Is Nothing Then
                RecuperaNotaFiscal()
                Dim ROM As [Lib].Negocio.Romaneio = Session("objRomaneioALT" & HID.Value)
                Session.Remove("objRomaneioALT" & HID.Value)
                Session.Remove("ProcurandoRomaneio" & HID.Value)

                If ROM.CodigoPedido = objNotaFiscal.CodigoPedido AndAlso
                   ROM.EntradaSaida = objNotaFiscal.EntradaSaida.ToString.Substring(0, 1) AndAlso
                   ROM.CodigoOperacao = objNotaFiscal.CodigoOperacao AndAlso
                   ROM.CodigoSubOperacao = objNotaFiscal.CodigoSubOperacao Then
                    If objNotaFiscal.Itens(0).Produto.ControlarRomaneio And ROM.PesoLiquido > objNotaFiscal.Itens(0).SaldoPedidoFisico And objNotaFiscal.SubOperacao.Classe <> eClassesOperacoes.DEPOSITOS And objNotaFiscal.SubOperacao.Classe <> eClassesOperacoes.AFIXAR Then
                        MsgBox(Me.Page, "O Pedido nao tem saldo suficiente para o Romaneio Selecionado. " & objNotaFiscal.Itens(0).SaldoPedidoFisico)
                        LimparCampos()
                    ElseIf objNotaFiscal.Itens(0).Produto.ControlarRomaneio And ROM.PesoLiquido > objNotaFiscal.Itens(0).SaldoPedidoFisico And objNotaFiscal.SubOperacao.EntradaSaida = eEntradaSaida.Saida And objNotaFiscal.SubOperacao.Devolucao = True Then
                        MsgBox(Me.Page, "O Pedido nao tem saldo suficiente para o Romaneio Selecionado. " & objNotaFiscal.Itens(0).SaldoPedidoFisico)
                        LimparCampos()
                    ElseIf objNotaFiscal.Itens(0).Produto.ControlarRomaneio And ROM.PesoLiquido > objNotaFiscal.Itens(0).SaldoPedidoFisico And objNotaFiscal.SubOperacao.EntradaSaida = eEntradaSaida.Entrada And objNotaFiscal.SubOperacao.Devolucao = True Then
                        MsgBox(Me.Page, "O Pedido nao tem saldo suficiente para o Romaneio Selecionado. " & objNotaFiscal.Itens(0).SaldoPedidoFisico)
                        LimparCampos()
                    Else
                        objNotaFiscal.CriarRomaneio = False
                        objNotaFiscal.Romaneio = ROM
                        objNotaFiscal.Deposito = objNotaFiscal.Romaneio.Deposito
                        Dim itemDeposito As ListItem = Funcoes.FormatarListItemCliente(objNotaFiscal.Deposito)

                        With objNotaFiscal.Deposito
                            objNotaFiscal.CodigoDeposito = .Codigo
                            objNotaFiscal.EnderecoDeposito = .CodigoEndereco
                        End With

                        If objNotaFiscal.EntradaSaida = eEntradaSaida.Entrada And Not ROM.NF Is Nothing Then
                            If ROM.NF.Numero > 0 And (objNotaFiscal.NotaProdutor <> ROM.NF.Numero Or objNotaFiscal.SerieNotaProdutor <> ROM.NF.Serie) Then
                                MsgBox(Me.Page, "O Numero da Nota Informada no Romaneio é diferente do Numero Informado Pelo Usuario para Emissao da Nota, Romaneio " & ROM.NF.Numero & "-" & ROM.NF.Serie & " Nota " & objNotaFiscal.NotaProdutor & "-" & objNotaFiscal.SerieNotaProdutor)
                            End If
                        End If

                        If objNotaFiscal.Itens(0).Produto.ControlarRomaneio Then
                            objNotaFiscal.Itens(0).QuantidadeFiscal = ROM.PesoLiquido
                            objNotaFiscal.Itens(0).QuantidadeFisica = ROM.PesoLiquido
                            objNotaFiscal.Itens(0).PesoBruto = ROM.PesoBruto
                            objNotaFiscal.Itens(0).PesoLiquido = ROM.PesoLiquido
                            objNotaFiscal.Itens(0).PesoFiscal = ROM.PesoLiquido
                        Else
                            objNotaFiscal.Itens(0).QuantidadeFiscal = objNotaFiscal.Itens(0).SaldoPedidoFiscal
                            objNotaFiscal.Itens(0).QuantidadeFisica = objNotaFiscal.Itens(0).SaldoPedidoFisico
                            objNotaFiscal.Itens(0).PesoBruto = ROM.PesoBruto
                            objNotaFiscal.Itens(0).PesoLiquido = ROM.PesoLiquido
                            objNotaFiscal.Itens(0).PesoFiscal = objNotaFiscal.Itens(0).SaldoPedidoFiscal
                        End If

                        objNotaFiscal.Itens(0).SaldoValorOficial = Math.Max(objNotaFiscal.Itens(0).SaldoPedidoFisico, objNotaFiscal.Itens(0).SaldoPedidoFiscal)
                        '#FimBaseDeCalculo
                        'objNotaFiscal.Itens(0).ValorTotal = (objNotaFiscal.Itens(0).QuantidadeFiscal * objNotaFiscal.Itens(0).Unitario) / objNotaFiscal.Itens(0).Produto.BaseCalculo
                        objNotaFiscal.Itens(0).ValorTotal = (objNotaFiscal.Itens(0).QuantidadeFiscal * objNotaFiscal.Itens(0).Unitario)
                        objNotaFiscal.CodigoRomaneio = ROM.Codigo

                        objNotaFiscal.CodigoTransportador = ROM.CodigoTransportador
                        objNotaFiscal.EnderecoTransportador = ROM.EnderecoTransportador
                        objNotaFiscal.CodigoMotorista = ROM.CodigoMotorista
                        objNotaFiscal.EnderecoMotorista = ROM.EnderecoMotorista
                        objNotaFiscal.PlacaTransportador = ROM.Placa
                        objNotaFiscal.CodigoAutorizacao = ROM.CodigoAutorizacao
                        If ROM.CodigoAutorizacao > 0 Then objNotaFiscal.ObservacoesDeEmbarque = objNotaFiscal.Autorizacao.Observacao

                        SalvaNotaFiscal()
                    End If
                Else
                    MsgBox(Me.Page, "Pedido do Romaneio " & ROM.CodigoPedido.ToString & " não pode ser diferente do Pedido da Nota Fiscal " & objNotaFiscal.CodigoPedido.ToString)
                    LimparCampos()
                End If
            ElseIf Not Session("ProcurandoRomaneio" & HID.Value) Is Nothing Then
                Session.Remove("ProcurandoRomaneio" & HID.Value)

                RecuperaNotaFiscal()
                If objNotaFiscal.SubOperacao.QuantidadeFisico Then
                    BuscaClassificacao()
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub
#End Region

    Private Sub SelecionarIndiceCombo(ByVal Combo As DropDownList, ByVal Valor As String)
        If Valor = "0" Then
            Combo.SelectedIndex = 0
        Else
            Combo.SelectedIndex = Combo.Items.IndexOf(Combo.Items.FindByValue(Valor))
        End If
    End Sub

    Private Sub BuscarFinalidades()
        ddl.Carregar(cmbFinalidade, CarregarDDL.Tabela.Finalidade, "")
    End Sub

    Private Sub LimparCampos()
        txtES.Text = ""
        txtNota.Text = ""
        txtSerie.Text = ""
        txtCliente.Text = ""
        txtCodigoCliente.Value = ""
        txtNota.Enabled = True
        txtES.Enabled = True
        txtSerie.Enabled = True
        lnkAtualizar.Parent.Visible = False
        lnkConsultar.Parent.Visible = True
        imgExtratoPedido.Enabled = False
        btnRomaneio.Enabled = True
        pnlDados.Visible = False
        pnlItem.Visible = False
        pnlGeral.Visible = False
        txtTitulo.Text = ""
        txtNaturezaDaOperacao.Text = ""
        txtNovaNaturezaDaOperacao.Text = ""
        txtSituacaoTributaria.Text = ""
        txtSituacaoTributariaIPI.Text = ""
        txtSituacaoTributariaPISCOFINS.Text = ""
        txtOperacaoXEstado.Text = ""
        txtNovaSituacaoTributaria.Text = ""
        txtNovaSituacaoTributariaIPI.Text = ""
        txtNovaSituacaoTributariaPISCOFINS.Text = ""
        txtNovaOperacaoXEstado.Text = ""
        txtNovaSituacaoTributaria.ForeColor = Drawing.Color.Blue
        txtNovaSituacaoTributariaIPI.ForeColor = Drawing.Color.Blue
        txtNovaSituacaoTributariaPISCOFINS.ForeColor = Drawing.Color.Blue
        lblTotalProduto.Text = "0,00"
        lblTotalNota.Text = "0,00"
        lblFisico.Text = "0"
        lblNovoTotalProduto.Text = "0,00"
        lblNovoTotalNota.Text = "0,00"
        lblNovoFisico.Text = "0"
        cmbSubOperacao.Items.Clear()
        cmbNovaSubOperacao.Items.Clear()
        cmbNovaSubOperacao.Enabled = False
        cmbFinalidade.SelectedIndex = 0
        btnFisico2.Visible = False
        btnFiscal2.Visible = False
        btnFisico2.Visible = False
        btnFiscal2.Visible = False
        Session.Remove("objClienteOPxNOTA" & HID.Value)
        Session.Remove("objNFConsultaNXI" & HID.Value)
        'Session.Remove("objNotaOriginalxOP" & HID.Value)
        Session.Remove("objClassificacaoALT" & HID.Value)
        Session.Remove("ProcurandoClassificacao" & HID.Value)
        Session.Remove("objRomaneioALT" & HID.Value)
        Session.Remove("ProcurandoRomaneio" & HID.Value)
        Session.Remove("Op" & HID.Value.ToString)
        Session.Remove("SOp" & HID.Value.ToString)
        Session.Remove("Unitario" & HID.Value.ToString)
        Session.Remove("ssCampo" & HID.Value)

        gridItem.DataSource = Nothing
        gridItem.DataBind()
        gridEncargos.DataSource = Nothing
        gridEncargos.DataBind()
        gridEncargosNovos.DataSource = Nothing
        gridEncargosNovos.DataBind()

        HID.Value = Guid.NewGuid().ToString
        ucConsultaPedidosXNotas.SetarHID(HID.Value)
        ucConsultaClientes.SetarHID(HID.Value)
        ucNotaFiscalXClassificacao.SetarHID(HID.Value)
        ucConsultaRomaneios.SetarHID(HID.Value)
        Try
            objNotaFiscal = New [Lib].Negocio.NotaFiscal()
            SalvaNotaFiscal()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub carregarItem()
        Try
            RecuperaNotaFiscal()

            Dim dtItens As New DataTable("Itens")
            dtItens.Columns.Add("Produto", Type.GetType("System.String"))
            dtItens.Columns.Add("NomeProduto", Type.GetType("System.String"))
            dtItens.Columns.Add("BaseCalculo", Type.GetType("System.String"))
            dtItens.Columns.Add("Lote", Type.GetType("System.String"))
            dtItens.Columns.Add("Classificacao", Type.GetType("System.String"))
            dtItens.Columns.Add("Embalagem", Type.GetType("System.String"))
            dtItens.Columns.Add("Saldo", Type.GetType("System.Decimal"))
            dtItens.Columns.Add("Quantidade", Type.GetType("System.Decimal"))
            dtItens.Columns.Add("QuantidadeF", Type.GetType("System.Decimal"))
            dtItens.Columns.Add("Unitario", Type.GetType("System.Decimal"))
            dtItens.Columns.Add("Total", Type.GetType("System.Decimal"))
            dtItens.Columns.Add("NumeroPecas", Type.GetType("System.Int32"))

            For Each row As [Lib].Negocio.NotaFiscalXItem In objNotaFiscal.Itens
                Dim drItem As DataRow = dtItens.NewRow()
                objNotaFiscal.Especie = row.Produto.Embalagem.Descricao
                row.CodigoDeposito = objNotaFiscal.CodigoDeposito
                row.EnderecoDeposito = objNotaFiscal.EnderecoDeposito
                drItem("Produto") = row.CodigoProduto
                drItem("Lote") = row.Lote
                drItem("Classificacao") = row.Classificacao
                drItem("Embalagem") = row.CodigoEmbalagemIndea + "-" + row.CodigoTipoDeEmbalagem + "-" + row.CapacidadeEmbalagem.ToString
                '#FimBaseDeCalculo
                'drItem("BaseCalculo") = row.Produto.BaseCalculo
                drItem("BaseCalculo") = 1
                drItem("Saldo") = IIf(row.QuantidadeFiscal = 0, row.SaldoValorOficial.ToString("N4"), row.QuantidadeFiscal.ToString("N4"))
                drItem("Total") = row.ValorTotal.ToString("N2")
                drItem("Unitario") = row.Unitario.ToString("N10")
                drItem("NumeroPecas") = row.NumeroPecas

                If (row.Produto.ControlarLote Or row.Produto.ControlarEmbalagem) And objNotaFiscal.SubOperacao.Classe <> eClassesOperacoes.GLOBAL Then
                    drItem("NomeProduto") = "$" & row.Produto.Nome
                    If objNotaFiscal.IUD <> "U" Then row.QuantidadeFiscal = IIf(row.Lote.Length > 0 Or row.CodigoEmbalagemIndea.Length > 0, row.QuantidadeFiscal, 0)
                Else
                    drItem("NomeProduto") = row.Produto.Nome
                End If

                drItem("Quantidade") = row.QuantidadeFiscal.ToString("N4")
                drItem("QuantidadeF") = row.QuantidadeFisica.ToString("N4")
                dtItens.Rows.Add(drItem)
            Next

            gridItem.DataSource = dtItens
            gridItem.DataBind()

            Dim i As Integer = 0
            While i < gridItem.Rows.Count
                If objNotaFiscal.NossaEmissao Then
                    CType(gridItem.Rows(i).FindControl("txtQuantidadeItem"), TextBox).Enabled = False
                    CType(gridItem.Rows(i).FindControl("txtUnitarioItem"), TextBox).Enabled = False
                    CType(gridItem.Rows(i).FindControl("txtTotalItem"), TextBox).Enabled = False
                Else
                    CType(gridItem.Rows(i).FindControl("txtQuantidadeItem"), TextBox).Enabled = True
                    CType(gridItem.Rows(i).FindControl("txtUnitarioItem"), TextBox).Enabled = True
                    CType(gridItem.Rows(i).FindControl("txtTotalItem"), TextBox).Enabled = True
                End If
                i += 1
            End While
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub ddlUnidadeDeNegocio_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            ddl.Carregar(ddlEmpresa, CarregarDDL.Tabela.Empresas, ddlUnidadeDeNegocio.SelectedValue, True)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnCliente_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            ucConsultaClientes.Limpar()
            Popup.ConsultaDeClientes(Me.Page, "objClienteOPxNOTA" & HID.Value.ToString, "txtCliente")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub ListarRomaneios(ByVal Pedido As String)
        Try
            RecuperaNotaFiscal()
            Dim operacao() As String = cmbSubOperacao.SelectedValue.Split("-")
            Session("ProcurandoRomaneio" & HID.Value.ToString) = True
            Session("Op" & HID.Value.ToString) = operacao(0)
            Session("SOp" & HID.Value.ToString) = operacao(1)
            Session("Unitario" & HID.Value.ToString) = Str(objNotaFiscal.Itens(0).Unitario)
            Popup.ConsultaDeRomaneios(Me, "objRomaneioALT" & HID.Value.ToString)
            ucConsultaRomaneios.LimparCampos()
            ucConsultaRomaneios.CargaRomaneios()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnRomaneio_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            RecuperaNotaFiscal()
            If objNotaFiscal.CodigoPedido = 0 Then
                MsgBox(Me.Page, "Para consultar o(s) Romaneio(s) é necessário informar o Pedido")
            Else
                Dim temDevolucao As Boolean = False

                For Each ni In objNotaFiscal.Itens
                    If Not ni.NotasDevolucao Is Nothing AndAlso ni.NotasDevolucao.Count > 0 Then temDevolucao = True
                Next

                If temDevolucao Then
                    If Funcoes.VerificaPermissao("AlterarNotaFiscal", "LIBERAR") Then
                        ''LIBERA ALTERAÇÃO - FURLAN - 01-11-2024
                    Else
                        MsgBox(Me.Page, "Não é possível alterar nota fiscal com NOTAS DE DEVOLUÇÃO.")
                        lnkAtualizar.Parent.Visible = False
                        Exit Sub
                    End If
                End If

                If objNotaFiscal.Itens(0).Produto.ControlarPesagem And objNotaFiscal.SubOperacao.Classe <> eClassesOperacoes.GLOBAL Then
                        If TemRomaneio(objNotaFiscal.CodigoPedido) = True AndAlso objNotaFiscal.SubOperacao.QuantidadeFisico Then
                            ListarRomaneios(objNotaFiscal.CodigoPedido)
                        Else
                            If objNotaFiscal.SubOperacao.QuantidadeFisico Then
                                BuscaClassificacao()
                            End If
                        End If
                    Else
                        MsgBox(Me.Page, "Não é informado a Classificação para este Produto")
                    End If
                End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Function TemRomaneio(ByVal Pedido As String) As Boolean
        RecuperaNotaFiscal()
        If objNotaFiscal.SubOperacao.Classe = eClassesOperacoes.GLOBAL Then
            MsgBox(Me.Page, "Não existe romaneios para Notas com Operação Global")
            Return False
        End If

        Dim ds As New DataSet
        Dim operacao() As String = cmbSubOperacao.SelectedValue.Split("-")
        Sql = "SELECT 1 " & vbCrLf &
              "  FROM Romaneios R" & vbCrLf &
              " INNER JOIN Produtos" & vbCrLf &
              "    ON R.Produto = Produtos.Produto_Id" & vbCrLf &
              " INNER JOIN RomaneiosXPesagens RxP" & vbCrLf &
              "    ON R.Empresa_Id    = RxP.Empresa_Id" & vbCrLf &
              "   AND R.EndEmpresa_Id = RxP.EndEmpresa_Id" & vbCrLf &
              "   AND R.Romaneio_Id   = RxP.Romaneio_Id" & vbCrLf &
              " INNER JOIN Pesagem Pe" & vbCrLf &
              "    ON RxP.Empresa_Id    = Pe.Empresa_Id" & vbCrLf &
              "   AND RxP.EndEmpresa_Id = Pe.EndEmpresa_Id" & vbCrLf &
              "   AND RxP.Pesagem_Id    = Pe.Pesagem_Id" & vbCrLf &
              "   AND RxP.Sequencia_Id  = Pe.Sequencia_Id" & vbCrLf &
              "  FULL JOIN NotasFiscaisXRomaneios NFR" & vbCrLf &
              "    ON R.Empresa_Id    = NFR.Empresa_Id" & vbCrLf &
              "   AND R.EndEmpresa_Id = NFR.EndEmpresa_Id" & vbCrLf &
              "   AND R.Romaneio_Id   = NFR.Romaneio_Id " & vbCrLf &
              " WHERE R.Empresa_id     ='" & objNotaFiscal.CodigoEmpresa & "'" & vbCrLf &
              "   AND R.EndEmpresa_id  = " & objNotaFiscal.EnderecoEmpresa & vbCrLf &
              "   AND R.Pedido         = " & Pedido & vbCrLf &
              "   AND NFR.Nota_Id IS NULL" & vbCrLf &
              "   AND Pe.Situacao = 1" & vbCrLf &
              "   AND R.Operacao       = " & operacao(0) & vbCrLf &
              "   AND R.SubOperacao    = " & operacao(1)

        ds = Banco.ConsultaDataSet(Sql, "Romaneios")

        If ds Is Nothing Then
            Return False
        ElseIf ds.Tables(0).Rows.Count = 0 Then
            Return False
        Else
            Return True
        End If
    End Function

    Private Sub BuscaClassificacao(Optional ByVal delay As Boolean = False)
        RecuperaNotaFiscal()
        If objNotaFiscal.SubOperacao.Classe = eClassesOperacoes.GLOBAL Then
            MsgBox(Me.Page, "Não é Possivel haver classificação de Produto para Notas Globais")
            Exit Sub
        End If
        If objNotaFiscal.CodigoRomaneio > 0 AndAlso Not objNotaFiscal.Romaneio.Processo = "NOTA FISCAL" Then
            MsgBox(Me.Page, "A Nota Já Possui um Romaneio / Pesagem")
            Exit Sub
        End If

        Session("ProcurandoClassificacao" & HID.Value.ToString) = True
        Popup.ConsultaDeNotaFiscalXClassificacao(Me, "objClassificacaoALT" & HID.Value.ToString, "txtPrimeiraPesagem", delay, 1000)
        ucNotaFiscalXClassificacao.LimparCampos()
        ucNotaFiscalXClassificacao.BindGridView()
    End Sub

    Private Function validarCampos() As Boolean
        RecuperaNotaFiscal()

        If cmbFinalidade.SelectedIndex = 0 Then
            MsgBox(Me.Page, "Finalidade não foi selecionada")
            Return False
        End If

        If objNotaFiscal.SubOperacao.QuantidadeFisico Then
            For Each item As [Lib].Negocio.NotaFiscalXItem In objNotaFiscal.Itens
                If item.QuantidadeFisica = 0 Then
                    MsgBox(Me.Page, "Quantidade Física não foi informada")
                    Return False
                End If
            Next
        End If

        Return True
    End Function

    Protected Sub cmbFinalidade_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            If cmbFinalidade.SelectedIndex > 0 Then
                RecuperaNotaFiscal()
                objNotaFiscal.CodigoFinalidade = cmbFinalidade.SelectedValue
                SalvaNotaFiscal()
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub gridItem_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles gridItem.SelectedIndexChanged
        Try
            RecuperaNotaFiscal()

            Dim temDevolucao As Boolean = False

            For Each ni In objNotaFiscal.Itens
                If Not ni.NotasDevolucao Is Nothing AndAlso ni.NotasDevolucao.Count > 0 Then temDevolucao = True
            Next

            If temDevolucao Then
                If Funcoes.VerificaPermissao("AlterarNotaFiscal", "LIBERAR") Then
                    ''LIBERA ALTERAÇÃO - FURLAN - 01-11-2024
                Else
                    MsgBox(Me.Page, "Não é possível alterar nota fiscal com NOTAS DE DEVOLUÇÃO.")
                    lnkAtualizar.Parent.Visible = False
                    Exit Sub
                End If
            End If


            gridEncargos.DataSource = objNotaFiscal.Itens(gridItem.SelectedIndex).Encargos.ToArray
            gridEncargos.DataBind()

            btnFisico.Visible = True
            btnFiscal.Visible = True

            If objNotaFiscal.SubOperacao.QuantidadeFisico Then
                btnFisico.BackColor = Drawing.Color.Green
            Else
                btnFisico.BackColor = Drawing.Color.Red
            End If

            If objNotaFiscal.SubOperacao.QuantidadeFiscal Then
                btnFiscal.BackColor = Drawing.Color.Green
            Else
                btnFiscal.BackColor = Drawing.Color.Red
            End If

            For Each enc As [Lib].Negocio.NotaFiscalXItemXEncargo In objNotaFiscal.Itens(gridItem.SelectedIndex).Encargos
                If enc.Codigo = "PRODUTO" Then
                    Dim SitTrib As New [Lib].Negocio.SituacaoTributaria(enc.SituacaoTributaria)
                    txtSituacaoTributaria.Text = SitTrib.Codigo & " - " & SitTrib.Descricao

                    Dim SitTribIPI As New [Lib].Negocio.SituacaoTributariaIPI(enc.SituacaoTributariaIPI)
                    txtSituacaoTributariaIPI.Text = SitTribIPI.Codigo & " - " & SitTribIPI.Descricao

                    Dim SitTribPIS As New [Lib].Negocio.SituacaoTributariaPISCOFINS(enc.SituacaoTributariaPISCOFINS)
                    txtSituacaoTributariaPISCOFINS.Text = SitTribPIS.Codigo & " - " & SitTribPIS.Descricao
                End If
            Next

            txtOperacaoXEstado.Text = objNotaFiscal.Itens(gridItem.SelectedIndex).CodigoOperacaoEstado

            If objNotaFiscal.Itens(gridItem.SelectedIndex).Produto.ControlarPecas AndAlso objNotaFiscal.Itens(gridItem.SelectedIndex).SubOperacao.ControlarPecas Then
                CType(gridItem.Rows(gridItem.SelectedIndex).FindControl("txtNumeroPecas"), TextBox).Enabled = True
            Else
                CType(gridItem.Rows(gridItem.SelectedIndex).FindControl("txtNumeroPecas"), TextBox).Enabled = False
            End If

            txtNaturezaDaOperacao.Text = objNotaFiscal.NaturezaDaOperacao

            Dim Par As New Hashtable()
            Par.Add("Empresa", objNotaFiscal.CodigoEmpresa)
            Par.Add("EndEmpresa", objNotaFiscal.EnderecoEmpresa)
            Par.Add("Operacao", objNotaFiscal.CodigoOperacao)
            Par.Add("SubOperacao", objNotaFiscal.CodigoSubOperacao)
            Par.Add("Pedido", objNotaFiscal.CodigoPedido)

            Par.Add("EstadoOrigem", objNotaFiscal.Empresa.Municipio.CodigoEstado)
            Par.Add("EstadoDestino", objNotaFiscal.Cliente.Municipio.CodigoEstado)
            Par.Add("RegiaoDestino", objNotaFiscal.Cliente.Estado.Regiao)
            Par.Add("GrupoDeProduto", objNotaFiscal.Itens(0).Produto.CodigoGrupo)
            Par.Add("Produto", objNotaFiscal.Itens(0).CodigoProduto)

            'Usado para listar apenas suboperações com situação 1 - Normal.
            ddl.Carregar(cmbNovaSubOperacao, CarregarDDL.Tabela.OperacaoSubOperacaoPermitidasNaNota, "Situacao = 1", True, Par)

            cmbNovaSubOperacao.Enabled = True
            pnlGeral.Visible = True
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub gridItem_RowCreated(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs)
        Try
            If e.Row.RowType = DataControlRowType.DataRow Then
                Dim btnItem As Button = CType(e.Row.FindControl("btnItem"), Button)
                btnItem.CommandArgument = e.Row.RowIndex.ToString()
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub gridItem_RowCommand(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewCommandEventArgs)
        Try
            If e.CommandName = "OK" Then
                RecuperaNotaFiscal()

                Dim temDevolucao As Boolean = False

                For Each ni In objNotaFiscal.Itens
                    If Not ni.NotasDevolucao Is Nothing AndAlso ni.NotasDevolucao.Count > 0 Then temDevolucao = True
                Next

                If temDevolucao Then
                    If Funcoes.VerificaPermissao("AlterarNotaFiscal", "LIBERAR") Then
                        ''LIBERA ALTERAÇÃO - FURLAN - 10-03-2024
                    Else
                        MsgBox(Me.Page, "Não é possível alterar nota fiscal com NOTAS DE DEVOLUÇÃO.")
                        lnkAtualizar.Parent.Visible = False
                        Exit Sub
                    End If
                End If

                If Not objNotaFiscal.VencimentosNota Is Nothing AndAlso
                    objNotaFiscal.VencimentosNota.Count > 0 Then

                    If objNotaFiscal.VencimentosNota.Where(Function(s) s.IUD <> "D").Any(Function(s) s.CodigoProvisao = eProvisao.Baixa) Then
                        MsgBox(Me.Page, "Não pode ser alterado o Valor da Nota Fiscal com Financeiro baixado, volte o título para PREVISÃO e refaça o processo.")
                        lnkAtualizar.Parent.Visible = False
                        Exit Sub
                    ElseIf objNotaFiscal.VencimentosNota.Where(Function(s) s.IUD <> "D").Any(Function(s) s.CodigoSituacao = eSituacao.RemessaBancaria) Then
                        MsgBox(Me.Page, "Não é possível alterar a nota fiscal com o financeiro em Cobrança Bancária.")
                        lnkAtualizar.Parent.Visible = False
                        Exit Sub
                    ElseIf objNotaFiscal.VencimentosNota.Where(Function(s) s.IUD <> "D").Any(Function(s) s.CodigoSituacao = eSituacao.EndossoTitulo) Then
                        MsgBox(Me.Page, "Não é possível alterar a nota fiscal com o financeiro em Endosso.")
                        lnkAtualizar.Parent.Visible = False
                        Exit Sub
                    End If
                End If

                Dim intIndice As Integer = Convert.ToInt32(e.CommandArgument)
                Dim intIndiceNota As Integer = intIndice

                objNotaFiscal.Itens(intIndice).IUD = "U"
                objNotaFiscal.Itens(intIndice).Encargos = Nothing
                objNotaFiscal.Itens(intIndice).PesoFiscal = CType(gridItem.Rows(intIndice).FindControl("txtQuantidadeItem"), TextBox).Text
                objNotaFiscal.Itens(intIndice).QuantidadeFisica = CType(gridItem.Rows(intIndice).FindControl("txtQuantidadeFItem"), TextBox).Text
                objNotaFiscal.Itens(intIndice).QuantidadeFiscal = CType(gridItem.Rows(intIndice).FindControl("txtQuantidadeItem"), TextBox).Text
                objNotaFiscal.Itens(intIndice).Unitario = CType(gridItem.Rows(intIndice).FindControl("txtUnitarioItem"), TextBox).Text

                If objNotaFiscal.Itens(intIndice).NumeroPecas > 0 Then
                    objNotaFiscal.Itens(intIndice).NumeroPecas = CType(gridItem.Rows(intIndice).FindControl("txtNumeroPecas"), TextBox).Text
                    objNotaFiscal.Itens(intIndice).QuantidadeFisica = CType(gridItem.Rows(intIndice).FindControl("txtQuantidadeItem"), TextBox).Text
                    CType(gridItem.Rows(intIndice).FindControl("txtQuantidadeFItem"), TextBox).Text = objNotaFiscal.Itens(intIndice).QuantidadeFisica
                Else
                    '#FimBaseDeCalculo
                    'Dim Valor As Decimal = IIf(objNotaFiscal.Itens(intIndice).QuantidadeFiscal = 0, CType(gridItem.Rows(intIndice).FindControl("txtTotalItem"), TextBox).Text, Math.Round((objNotaFiscal.Itens(intIndice).QuantidadeFiscal * objNotaFiscal.Itens(intIndice).Unitario) / objNotaFiscal.Itens(intIndiceNota).Produto.BaseCalculo, 2, MidpointRounding.AwayFromZero))
                    Dim Valor As Decimal = IIf(objNotaFiscal.Itens(intIndice).QuantidadeFiscal = 0, CType(gridItem.Rows(intIndice).FindControl("txtTotalItem"), TextBox).Text, Math.Round((objNotaFiscal.Itens(intIndice).QuantidadeFiscal * objNotaFiscal.Itens(intIndice).Unitario), 2, MidpointRounding.AwayFromZero))
                    objNotaFiscal.Itens(intIndiceNota).ValorTotal = Valor
                End If

                CType(gridItem.Rows(intIndice).FindControl("txtTotalItem"), TextBox).Text = objNotaFiscal.Itens(intIndiceNota).ValorTotal

                gridEncargos.DataSource = objNotaFiscal.Itens(intIndice).Encargos.ToArray
                gridEncargos.DataBind()

                lblTotalProduto.Text = objNotaFiscal.TotalProduto.ToString("N2")
                lblTotalNota.Text = objNotaFiscal.TotalNota.ToString("N2")

                'Fianceiro NF de Entrada
                If objNotaFiscal.VencimentosNota IsNot Nothing AndAlso objNotaFiscal.VencimentosNota.Count > 0 Then
                    If Not objNotaFiscal.NossaEmissao AndAlso objNotaFiscal.EntradaSaida = eEntradaSaida.Entrada AndAlso objNotaFiscal.TotalNota <> objNotaFiscal.VencimentosNota.Sum(Function(x) x.ValorDoDocumento) Then
                        objNotaFiscal.VencimentosNota.ReajFinanceiro = New ReajusteFinanceiro(objNotaFiscal, False)
                        objNotaFiscal.VencimentosNota.ReajFinanceiro.ReajustaNotaDeEntrada()
                    End If
                End If

                SalvaNotaFiscal()
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub cmbNovaSubOperacao_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbNovaSubOperacao.SelectedIndexChanged
        Try
            RecuperaNotaFiscal()

            If cmbNovaSubOperacao.SelectedIndex > 0 Then
                Dim so() As String = cmbNovaSubOperacao.SelectedValue.ToString.Split("-")
                Dim sOpe As New [Lib].Negocio.SubOperacao(so(0), so(1))

                If Not sOpe.EntradaSaida = objNotaFiscal.EntradaSaida Then
                    MsgBox(Me.Page, "Operação selecionada de " & IIf(sOpe.EntradaSaida = eEntradaSaida.Entrada, "ENTRADA", "SAÍDA") & " não pode ser usada na Nota Fiscal de " & IIf(objNotaFiscal.EntradaSaida = eEntradaSaida.Entrada, "ENTRADA", "SAÍDA"))
                    cmbNovaSubOperacao.SelectedIndex = cmbSubOperacao.SelectedIndex
                    Exit Sub
                End If



                objNotaFiscal.CodigoOperacao = sOpe.CodigoOperacao
                objNotaFiscal.CodigoSubOperacao = sOpe.Codigo

                Dim listaDeEncargosOld As New ListNotaFiscalXItemXEncargo

                For x As Integer = 0 To objNotaFiscal.Itens.Count - 1

                    objNotaFiscal.Itens(x).IUD = objNotaFiscal.IUD

                    objNotaFiscal.Itens(x).CodigoOperacao = sOpe.CodigoOperacao
                    objNotaFiscal.Itens(x).CodigoSubOperacao = sOpe.Codigo
                    objNotaFiscal.Itens(x).CodigoOperacaoEstado = 0
                    objNotaFiscal.Itens(x).Encargos.Clear()

                    objNotaFiscal.Itens(x).CodigoOperacaoEstado = ConsultarIDFiscal(x)

                    objNotaFiscal.Itens(x).CarregandoEncargos = True
                    objNotaFiscal.Itens(x).Encargos = New [Lib].Negocio.ListNotaFiscalXItemXEncargo(objNotaFiscal.Itens(x), New List(Of eEtapaEncago) From {eEtapaEncago.Normal})
                    objNotaFiscal.Itens(x).CarregandoEncargos = False

                    If objNotaFiscal.Itens(x).Encargos.Count = 0 Then
                        MsgBox(Me.Page, "O Produto " & objNotaFiscal.Itens(x).Produto.Nome & ", não tem encargos cadastrados na Operação:" & objNotaFiscal.Itens(x).CodigoOperacao & "-" & objNotaFiscal.Itens(x).CodigoSubOperacao)
                        Exit Sub
                    End If

                    objNotaFiscal.Itens(x).Encargos.AtualizaLiquido()

                    Dim codigocfop As Integer = objNotaFiscal.Itens.OrderByDescending(Function(s) s.ValorTotal).First.OperacaoEstado.CodigoFiscal
                    objNotaFiscal.CFOP = New [Lib].Negocio.CFOP(codigocfop)

                    Dim saldo As Decimal = objNotaFiscal.Itens(x).SaldoValorOficial + objNotaFiscal.NotaFiscalOriginal.Itens(x).QuantidadeFiscal
                    'Dim saldo As Decimal = objNotaFiscal.Itens(x).Saldo + objNotaFiscal.NotaFiscalOriginal.Itens(x).QuantidadeFiscal

                    If objNotaFiscal.Itens(x).CFOP = 0 Then
                        MsgBox(Me.Page, "Operação:" & objNotaFiscal.Itens(x).CodigoOperacao & "-" & objNotaFiscal.Itens(x).CodigoSubOperacao & " está sem CFOP informada")
                        Exit Sub
                    ElseIf objNotaFiscal.Itens(x).CFOP < 5000 And objNotaFiscal.EntradaSaida = eEntradaSaida.Saida Then
                        MsgBox(Me.Page, "Operação " & objNotaFiscal.Itens(x).CodigoOperacao & "-" & objNotaFiscal.Itens(x).CodigoSubOperacao & " com CFOP " & objNotaFiscal.Itens(x).CFOP & " não pode ser usada na saída.")
                        Exit Sub
                    ElseIf objNotaFiscal.Itens(x).CFOP > 5000 And objNotaFiscal.EntradaSaida = eEntradaSaida.Entrada Then
                        MsgBox(Me.Page, "Operação " & objNotaFiscal.Itens(x).CodigoOperacao & "-" & objNotaFiscal.Itens(x).CodigoSubOperacao & " com CFOP " & objNotaFiscal.Itens(x).CFOP & " não pode ser usada na entrada.")
                        Exit Sub
                    ElseIf objNotaFiscal.EntradaSaida = eEntradaSaida.Saida And objNotaFiscal.Itens(x).SaldoValorOficial <= 0 And objNotaFiscal.SubOperacao.Classe <> eClassesOperacoes.DEPOSITOS And objNotaFiscal.SubOperacao.Classe <> eClassesOperacoes.AFIXAR And objNotaFiscal.SubOperacao.Classe <> eClassesOperacoes.GLOBAL Then
                        MsgBox(Me.Page, "Pedido com saldo 0 ou negativo não pode ser usado para emissão de Nota Fiscal. Saldo: " & objNotaFiscal.Itens(x).SaldoValorOficial)
                        Exit Sub
                    End If
                Next

                For Each encOri As [Lib].Negocio.NotaFiscalXItemXEncargo In objNotaFiscal.NotaFiscalOriginal.Itens(gridItem.SelectedIndex).Encargos
                    If encOri.Codigo.Contains("DESP.ADUANEIRAS") _
                        OrElse encOri.Codigo.Contains("IPI") _
                        OrElse encOri.Codigo.Contains("SEGURO") _
                        OrElse encOri.Codigo.Contains("FRETES") _
                        OrElse encOri.Codigo.Contains("DESCONTOS") _
                        OrElse encOri.Codigo.Contains("INSS") _
                        OrElse encOri.Codigo.Contains("IRRF PF") _
                        OrElse encOri.Codigo.Contains("CSRF") _
                        OrElse encOri.Codigo.Contains("IRRF PJ") _
                        OrElse encOri.Codigo.Contains("ISS") _
                        OrElse encOri.Codigo.Contains("ISS PJ") _
                        OrElse encOri.Codigo.Contains("ICMS") Then
                        For Each encDes As [Lib].Negocio.NotaFiscalXItemXEncargo In objNotaFiscal.Itens(gridItem.SelectedIndex).Encargos
                            If encDes.Codigo.Contains("ICMS") AndAlso encOri.Codigo.Contains("ICMS") Then
                                encDes.Base = encOri.Base
                                encDes.Percentual = encOri.Percentual
                                encDes.PercentualExibicao = encOri.PercentualExibicao
                                encDes.Valor = encOri.Valor
                            ElseIf encDes.Codigo = encOri.Codigo Then
                                encDes.Base = encOri.Base
                                encDes.Valor = encOri.Valor
                            End If
                        Next
                    End If
                Next

                btnFisico2.Visible = True
                btnFiscal2.Visible = True
                lnkAtualizar.Parent.Visible = True
                panelNovaOperacao.Visible = True

                If sOpe.QuantidadeFisico Then
                    btnFisico2.BackColor = Drawing.Color.Green
                Else
                    btnFisico2.BackColor = Drawing.Color.Red
                    objNotaFiscal.Itens(gridItem.SelectedIndex).QuantidadeFisica = 0
                End If

                If sOpe.QuantidadeFiscal Then
                    btnFiscal2.BackColor = Drawing.Color.Green
                Else
                    btnFiscal2.BackColor = Drawing.Color.Red
                End If

                gridEncargosNovos.DataSource = objNotaFiscal.Itens(gridItem.SelectedIndex).Encargos.ToArray
                gridEncargosNovos.DataBind()

                objNotaFiscal.AtualizaTotais()

                Dim totalEncNF As Decimal = 0

                For Each enc As [Lib].Negocio.NotaFiscalXItemXEncargo In objNotaFiscal.Itens(gridItem.SelectedIndex).Encargos
                    If enc.Codigo = "PRODUTO" Then
                        Dim SitTrib As New [Lib].Negocio.SituacaoTributaria(enc.SituacaoTributaria)
                        txtNovaSituacaoTributaria.Text = SitTrib.Codigo & " - " & SitTrib.Descricao

                        Dim SitTribIPI As New [Lib].Negocio.SituacaoTributariaIPI(enc.SituacaoTributariaIPI)
                        txtNovaSituacaoTributariaIPI.Text = SitTribIPI.Codigo & " - " & SitTribIPI.Descricao

                        Dim SitTribPIS As New [Lib].Negocio.SituacaoTributariaPISCOFINS(enc.SituacaoTributariaPISCOFINS)
                        txtNovaSituacaoTributariaPISCOFINS.Text = SitTribPIS.Codigo & " - " & SitTribPIS.Descricao
                    ElseIf enc.Codigo.Contains("FACS") AndAlso enc.Sinal = "-" Then
                        totalEncNF += enc.Valor
                    ElseIf enc.Codigo.Contains("FETHAB") AndAlso enc.Sinal = "-" Then
                        totalEncNF += enc.Valor
                    ElseIf enc.Codigo.Contains("SENAR") AndAlso enc.Sinal = "-" Then
                        totalEncNF += enc.Valor
                    ElseIf enc.Codigo.Contains("FUNRURAL") AndAlso enc.Sinal = "-" Then
                        totalEncNF += enc.Valor
                    End If
                Next

                txtNovaNaturezaDaOperacao.Text = objNotaFiscal.NaturezaDaOperacao
                lblNovoTotalProduto.Text = objNotaFiscal.TotalProduto.ToString("N2")
                lblNovoTotalNota.Text = objNotaFiscal.TotalNota.ToString("N2")

                txtNovaOperacaoXEstado.Text = objNotaFiscal.Itens(gridItem.SelectedIndex).CodigoOperacaoEstado

                Dim liberaEcargo As Boolean = True

                If objNotaFiscal.NossaEmissao Then

                    liberaEcargo = False

                    Dim i As Integer = 0
                    Dim j As Integer = 0
                    Dim temEnc As Boolean
                    Dim vlrOk As Boolean

                    While j < gridEncargos.Rows.Count
                        i = 0
                        temEnc = False
                        vlrOk = False
                        While i < gridEncargosNovos.Rows.Count
                            If gridEncargosNovos.Rows(i).Cells(0).Text() = gridEncargos.Rows(j).Cells(0).Text() Then
                                temEnc = True
                                If CDec(CType(gridEncargosNovos.Rows(i).FindControl("txtValorEncargoItem"), TextBox).Text) = CDec(CType(gridEncargos.Rows(j).FindControl("txtValorEncargoItem"), TextBox).Text) Then
                                    vlrOk = True
                                End If
                            End If
                            i += 1
                        End While

                        If temEnc = False AndAlso totalEncNF = 0 Then
                            If Not gridEncargos.Rows(j).Cells(0).Text().Contains("PIS") AndAlso
                                Not gridEncargos.Rows(j).Cells(0).Text().Contains("COFINS") AndAlso
                                Not gridEncargos.Rows(j).Cells(0).Text().Contains("FETHAB") AndAlso
                                Not gridEncargos.Rows(j).Cells(0).Text().Contains("FACS") AndAlso
                                Not gridEncargos.Rows(j).Cells(0).Text().Contains("SENAR") AndAlso
                                Not gridEncargos.Rows(j).Cells(0).Text().Contains("FUNRURAL") Then
                                MsgBox(Me.Page, "Encargo " & gridEncargos.Rows(j).Cells(0).Text() & " não foi encontrado na nova lista da Operação")
                                lnkAtualizar.Parent.Visible = False
                                Exit Sub
                            End If
                        ElseIf vlrOk = False AndAlso totalEncNF = 0 Then
                            If Not gridEncargos.Rows(j).Cells(0).Text().Contains("PIS") AndAlso Not gridEncargos.Rows(j).Cells(0).Text().Contains("COFINS") AndAlso Not gridEncargos.Rows(j).Cells(0).Text().Contains("FETHAB VENDA") Then
                                Exit Sub
                            End If

                            If Not gridEncargos.Rows(j).Cells(0).Text().Contains("PIS") AndAlso
                                Not gridEncargos.Rows(j).Cells(0).Text().Contains("COFINS") AndAlso
                                Not gridEncargos.Rows(j).Cells(0).Text().Contains("FETHAB") AndAlso
                                Not gridEncargos.Rows(j).Cells(0).Text().Contains("FACS") AndAlso
                                Not gridEncargos.Rows(j).Cells(0).Text().Contains("SENAR") AndAlso
                                Not gridEncargos.Rows(j).Cells(0).Text().Contains("FUNRURAL") Then
                                MsgBox(Me.Page, "Valor do Encargo " & gridEncargos.Rows(j).Cells(0).Text() & " não é o mesmo na nova lista da Operação")
                                lnkAtualizar.Parent.Visible = False
                                Exit Sub
                            End If
                        End If

                        If (gridEncargos.Rows(j).Cells(0).Text().Contains("PIS") And gridEncargos.Rows(j).Cells(14).Text().Contains("=")) OrElse
                                (gridEncargos.Rows(j).Cells(0).Text().Contains("COFINS") And gridEncargos.Rows(j).Cells(14).Text().Contains("=")) OrElse
                                (gridEncargos.Rows(j).Cells(0).Text().Contains("FETHAB") And gridEncargos.Rows(j).Cells(14).Text().Contains("-")) OrElse
                                (gridEncargos.Rows(j).Cells(0).Text().Contains("FACS") And gridEncargos.Rows(j).Cells(14).Text().Contains("-")) OrElse
                                (gridEncargos.Rows(j).Cells(0).Text().Contains("SENAR") And gridEncargos.Rows(j).Cells(14).Text().Contains("-")) OrElse
                                (gridEncargos.Rows(j).Cells(0).Text().Contains("FUNRURAL") And gridEncargos.Rows(j).Cells(14).Text().Contains("-")) Then
                            liberaEcargo = True
                        End If

                        If gridEncargos.Rows(j).Cells(0).Text().Contains("ICMS") Then
                            liberaEcargo = False
                            Exit While
                        End If

                        j += 1
                    End While

                    If txtSituacaoTributaria.Text <> txtNovaSituacaoTributaria.Text Then
                        MsgBox(Me.Page, "Situação Tributária ICMS da Operação não pode ser diferente da informada na Nota Eletrônica")
                        txtNovaSituacaoTributaria.ForeColor = Drawing.Color.Red
                        lnkAtualizar.Parent.Visible = False
                        Exit Sub
                    ElseIf txtSituacaoTributariaIPI.Text <> txtNovaSituacaoTributariaIPI.Text Then
                        MsgBox(Me.Page, "Situação Tributária IPI da Operação não pode ser diferente da informada na Nota Eletrônica")
                        txtNovaSituacaoTributariaIPI.ForeColor = Drawing.Color.Red
                        lnkAtualizar.Parent.Visible = False
                        Exit Sub
                    ElseIf txtSituacaoTributariaPISCOFINS.Text <> txtNovaSituacaoTributariaPISCOFINS.Text Then
                        MsgBox(Me.Page, "Situação Tributária PIS/COFINS da Operação não pode ser diferente da informada na Nota Eletrônica")
                        txtNovaSituacaoTributariaPISCOFINS.ForeColor = Drawing.Color.Red
                    ElseIf txtNaturezaDaOperacao.Text <> txtNovaNaturezaDaOperacao.Text Then
                        MsgBox(Me.Page, "CFOP da Operação não pode ser diferente do informado na Nota Eletrônica")
                        lnkAtualizar.Parent.Visible = False
                        Exit Sub
                    ElseIf CDec(lblTotalProduto.Text) <> CDec(lblNovoTotalProduto.Text) Then
                        MsgBox(Me.Page, "Valor do Produto não pode ser diferente do informado na Nota Eletrônica")
                        lnkAtualizar.Parent.Visible = False
                        Exit Sub
                    ElseIf CDec(lblTotalNota.Text) <> CDec(lblNovoTotalNota.Text) Then
                        If CDec(lblNovoTotalNota.Text) > CDec(lblTotalNota.Text) AndAlso (CDec(lblNovoTotalNota.Text) - CDec(lblTotalNota.Text) = totalEncNF) Then
                            'Deixa passar pela diferença ser facs, fethab, senar ou funrural
                        ElseIf CDec(lblNovoTotalNota.Text) < CDec(lblTotalNota.Text) AndAlso (CDec(lblTotalNota.Text) - CDec(lblNovoTotalNota.Text) = totalEncNF) Then
                            'Deixa passar pela diferença ser facs, fethab, senar ou funrural
                        Else
                            If Not liberaEcargo Then
                                MsgBox(Me.Page, "Total da Nota Fiscal não pode ser diferente do informado na Nota Eletrônica")
                                lnkAtualizar.Parent.Visible = False
                                Exit Sub

                            End If
                        End If
                    End If
                Else
                    For Each enc As [Lib].Negocio.NotaFiscalXItemXEncargo In objNotaFiscal.Itens(gridItem.SelectedIndex).Encargos
                        If enc.Codigo = "ICMS" AndAlso enc.Valor > 0 AndAlso enc.SituacaoTributaria <> 0 AndAlso enc.SituacaoTributaria <> 20 AndAlso enc.SituacaoTributaria <> 51 Then
                            MsgBox(Me.Page, "Nota Fiscal com ICMS difere da Situação tributária informada")
                            txtNovaSituacaoTributaria.ForeColor = Drawing.Color.Red
                            lnkAtualizar.Parent.Visible = False
                            Exit Sub
                        ElseIf enc.Codigo.Contains("IPI") Then
                            If objNotaFiscal.EntradaSaida = eEntradaSaida.Entrada AndAlso enc.Valor > 0 AndAlso enc.SituacaoTributariaIPI <> 0 Then
                                MsgBox(Me.Page, "Nota Fiscal com IPI difere da Situação tributária informada")
                                txtNovaSituacaoTributariaIPI.ForeColor = Drawing.Color.Red
                                lnkAtualizar.Parent.Visible = False
                                Exit Sub
                            ElseIf objNotaFiscal.EntradaSaida = eEntradaSaida.Saida AndAlso enc.Valor > 0 AndAlso enc.SituacaoTributariaIPI <> 50 Then
                                MsgBox(Me.Page, "Nota Fiscal com IPI difere da Situação tributária informada")
                                txtNovaSituacaoTributariaIPI.ForeColor = Drawing.Color.Red
                                lnkAtualizar.Parent.Visible = False
                                Exit Sub
                            ElseIf objNotaFiscal.EntradaSaida = eEntradaSaida.Entrada AndAlso enc.Valor = 0 AndAlso enc.SituacaoTributariaIPI <> 0 Then
                                MsgBox(Me.Page, "Nota Fiscal sem IPI difere da Situação tributária informada")
                                txtNovaSituacaoTributariaIPI.ForeColor = Drawing.Color.Red
                                lnkAtualizar.Parent.Visible = False
                                Exit Sub
                            ElseIf objNotaFiscal.EntradaSaida = eEntradaSaida.Saida AndAlso enc.Valor = 0 AndAlso enc.SituacaoTributariaIPI = 50 Then
                                MsgBox(Me.Page, "Nota Fiscal sem IPI difere da Situação tributária informada")
                                txtNovaSituacaoTributariaIPI.ForeColor = Drawing.Color.Red
                                lnkAtualizar.Parent.Visible = False
                                Exit Sub
                            End If
                        ElseIf enc.Codigo.Contains("PIS") Then
                            If txtSituacaoTributariaPISCOFINS.Text <> txtNovaSituacaoTributariaPISCOFINS.Text Then
                                txtNovaSituacaoTributariaPISCOFINS.ForeColor = Drawing.Color.Red
                                MsgBox(Me.Page, "Situação Tributária PIS/COFINS da Operação é diferente da informada na Operação anterior, tenha certeza antes de ATUALIZAR para as novas configurações")
                            End If
                        End If
                    Next
                End If

                If objNotaFiscal.SubOperacao.QuantidadeFisico AndAlso objNotaFiscal.Itens(0).Produto.ControlarPesagem AndAlso (objNotaFiscal.CodigoRomaneio = 0 OrElse (objNotaFiscal.CodigoRomaneio > 0 AndAlso objNotaFiscal.Romaneio.Processo = "NOTA FISCAL")) Then
                    If objNotaFiscal.CodigoRomaneio = 0 Then
                        objNotaFiscal.CriarRomaneio = True
                    End If
                    SalvaNotaFiscal()

                    BuscaClassificacao()
                ElseIf objNotaFiscal.SubOperacao.QuantidadeFisico Then
                    SalvaNotaFiscal()
                    carregarItem()
                    If objNotaFiscal.Romaneio IsNot Nothing AndAlso objNotaFiscal.Romaneio.Codigo > 0 Then lblNovoFisico.Text = objNotaFiscal.Romaneio.PesoLiquido
                Else
                    SalvaNotaFiscal()
                    carregarItem()
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Function ConsultarIDFiscal(ByVal item As Integer) As Integer

        Dim bConsultaPorProduto As Boolean = False

ConsultaOperacaoXEstado:

        Dim Parametros As New OperacaoXEstado
        Parametros.Empresa = Left(objNotaFiscal.CodigoEmpresa, 8)
        Parametros.CodigoGrupoProduto = objNotaFiscal.Itens(item).Produto.CodigoGrupo


        'Consulta operação por produto
        If bConsultaPorProduto Then
            Parametros.CodigoProduto = objNotaFiscal.Itens(item).CodigoProduto
        End If

        Parametros.CodigoOperacao = objNotaFiscal.Itens(item).CodigoOperacao
        Parametros.CodigoSubOperacao = objNotaFiscal.Itens(item).CodigoSubOperacao
        Parametros.EstadoOrigem = objNotaFiscal.Empresa.CodigoEstado
        Parametros.EstadoDestino = IIf(objNotaFiscal.SubOperacao.Operacao.UFDepositoDestino, objNotaFiscal.Destino.CodigoEstado, objNotaFiscal.Cliente.CodigoEstado)
        Parametros.InicioVigencia = objNotaFiscal.Movimento

        Dim OXE As New OperacaoXEstado(Parametros)

        If Parametros.Codigo = 0 And bConsultaPorProduto = False Then
            bConsultaPorProduto = True
            GoTo ConsultaOperacaoXEstado
        End If

        Return OXE.Codigo

    End Function

    Protected Sub imgExtratoPedido_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)
        Try
            RecuperaNotaFiscal()
            Extrato.Emitir(Me.Page, FinanceiroNovo, objNotaFiscal.CodigoEmpresa, objNotaFiscal.EnderecoEmpresa, "T", objNotaFiscal.CodigoPedido)
            'Dim strQueryString As String = "?fim=" & DateTime.Now.ToString("dd/MM/yyyy")
            'strQueryString &= "&empresa=" & objNotaFiscal.CodigoEmpresa & "-" & objNotaFiscal.EnderecoEmpresa
            'strQueryString &= "&cliente=" & objNotaFiscal.CodigoCliente & "-" & objNotaFiscal.EnderecoCliente
            'strQueryString &= "&pedido=" & objNotaFiscal.CodigoPedido
            'strQueryString &= "&es=ES"
            'ScriptManager.RegisterClientScriptBlock(Me, Me.GetType, "CarregarHTML", "window.open(""ExtratoPedidoHTML.aspx" & strQueryString & """, ""relatorio"", ""status=no, toolbar=yes, location=no, menubar=yes, resizable=yes, height=600, width=800, scrollbars=yes"");", True)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnAtualizar_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnAtualizar.Click
        Try
            RecuperaNotaFiscal()

            Dim so() As String = {"", ""}
            If Not String.IsNullOrWhiteSpace(cmbNovaSubOperacao.SelectedValue) Then
                so(0) = cmbNovaSubOperacao.SelectedValue.Split("-")(0)
                so(1) = cmbNovaSubOperacao.SelectedValue.Split("-")(1)
            Else
                so(0) = cmbSubOperacao.SelectedValue.Split("-")(0)
                so(1) = cmbSubOperacao.SelectedValue.Split("-")(1)
            End If

            'Dim so() As String = cmbSubOperacao.SelectedValue.ToString.Split("-")
            'Dim so() As String = cmbNovaSubOperacao.SelectedValue.ToString.Split("-")

            Dim sOpe As New [Lib].Negocio.SubOperacao(so(0), so(1))

            If Not sOpe.EntradaSaida = objNotaFiscal.EntradaSaida Then
                MsgBox(Me.Page, "Operação selecionada de " & IIf(sOpe.EntradaSaida = eEntradaSaida.Entrada, "ENTRADA", "SAÍDA") & " não pode ser usada na Nota Fiscal de " & IIf(objNotaFiscal.EntradaSaida = eEntradaSaida.Entrada, "ENTRADA", "SAÍDA"))
                cmbNovaSubOperacao.SelectedIndex = 0
                Exit Sub
            End If

            objNotaFiscal.CodigoOperacao = sOpe.CodigoOperacao
            objNotaFiscal.CodigoSubOperacao = sOpe.Codigo

            objNotaFiscal.Itens(gridItem.SelectedIndex).IUD = "U"
            objNotaFiscal.Itens(gridItem.SelectedIndex).CodigoOperacao = sOpe.CodigoOperacao
            objNotaFiscal.Itens(gridItem.SelectedIndex).CodigoSubOperacao = sOpe.Codigo

            If objNotaFiscal.Itens(gridItem.SelectedIndex).Encargos.Count = 0 Then
                MsgBox(Me.Page, "O Produto " & objNotaFiscal.Itens(gridItem.SelectedIndex).Produto.Nome & ", não tem encargos cadastrados na Operação:" & objNotaFiscal.Itens(gridItem.SelectedIndex).CodigoOperacao & "-" & objNotaFiscal.Itens(gridItem.SelectedIndex).CodigoSubOperacao)
                Exit Sub
            ElseIf objNotaFiscal.Itens(gridItem.SelectedIndex).CFOP = 0 Then
                MsgBox(Me.Page, "Operação:" & objNotaFiscal.Itens(gridItem.SelectedIndex).CodigoOperacao & "-" & objNotaFiscal.Itens(gridItem.SelectedIndex).CodigoSubOperacao & " está sem CFOP informada")
                Exit Sub
            ElseIf objNotaFiscal.Itens(gridItem.SelectedIndex).CFOP < 5000 And objNotaFiscal.EntradaSaida = eEntradaSaida.Saida Then
                MsgBox(Me.Page, "Operação " & objNotaFiscal.Itens(gridItem.SelectedIndex).CodigoOperacao & "-" & objNotaFiscal.Itens(gridItem.SelectedIndex).CodigoSubOperacao & " com CFOP " & objNotaFiscal.Itens(gridItem.SelectedIndex).CFOP & " não pode ser usada na saída.")
                Exit Sub
            ElseIf objNotaFiscal.Itens(gridItem.SelectedIndex).CFOP > 5000 And objNotaFiscal.EntradaSaida = eEntradaSaida.Entrada Then
                MsgBox(Me.Page, "Operação " & objNotaFiscal.Itens(gridItem.SelectedIndex).CodigoOperacao & "-" & objNotaFiscal.Itens(gridItem.SelectedIndex).CodigoSubOperacao & " com CFOP " & objNotaFiscal.Itens(gridItem.SelectedIndex).CFOP & " não pode ser usada na entrada.")
                Exit Sub
            ElseIf objNotaFiscal.EntradaSaida = eEntradaSaida.Saida And objNotaFiscal.Itens(gridItem.SelectedIndex).SaldoValorOficial <= 0 And objNotaFiscal.SubOperacao.Classe <> eClassesOperacoes.DEPOSITOS And objNotaFiscal.SubOperacao.Classe <> eClassesOperacoes.AFIXAR And objNotaFiscal.SubOperacao.Classe <> eClassesOperacoes.GLOBAL Then
                MsgBox(Me.Page, "Pedido com saldo 0 ou negativo não pode ser usado para emissão de Nota Fiscal. Saldo: " & objNotaFiscal.Itens(gridItem.SelectedIndex).SaldoValorOficial)
                Exit Sub
            End If

            btnFisico2.Visible = True
            btnFiscal2.Visible = True
            lnkAtualizar.Parent.Visible = True

            If sOpe.QuantidadeFisico Then
                btnFisico2.BackColor = Drawing.Color.Green
            Else
                btnFisico2.BackColor = Drawing.Color.Red
                objNotaFiscal.Itens(gridItem.SelectedIndex).QuantidadeFisica = 0
            End If

            If sOpe.QuantidadeFiscal Then
                btnFiscal2.BackColor = Drawing.Color.Green
            Else
                btnFiscal2.BackColor = Drawing.Color.Red
            End If

            objNotaFiscal.Itens(gridItem.SelectedIndex).Encargos.Clear()

            objNotaFiscal.Itens(gridItem.SelectedIndex).CarregandoEncargos = True
            objNotaFiscal.Itens(gridItem.SelectedIndex).Encargos = New [Lib].Negocio.ListNotaFiscalXItemXEncargo(objNotaFiscal.Itens(gridItem.SelectedIndex), New List(Of eEtapaEncago) From {eEtapaEncago.Normal})
            objNotaFiscal.Itens(gridItem.SelectedIndex).CarregandoEncargos = False

            objNotaFiscal.Itens(gridItem.SelectedIndex).Encargos.AtualizaLiquido()

            gridEncargosNovos.DataSource = objNotaFiscal.Itens(gridItem.SelectedIndex).Encargos.ToArray
            gridEncargosNovos.DataBind()

            objNotaFiscal.AtualizaTotais()

            Dim totalEncNF As Decimal = 0

            For Each enc As [Lib].Negocio.NotaFiscalXItemXEncargo In objNotaFiscal.Itens(gridItem.SelectedIndex).Encargos
                If enc.Codigo = "PRODUTO" Then
                    Dim SitTrib As New [Lib].Negocio.SituacaoTributaria(enc.SituacaoTributaria)
                    txtNovaSituacaoTributaria.Text = SitTrib.Codigo & " - " & SitTrib.Descricao

                    Dim SitTribIPI As New [Lib].Negocio.SituacaoTributariaIPI(enc.SituacaoTributariaIPI)
                    txtNovaSituacaoTributariaIPI.Text = SitTribIPI.Codigo & " - " & SitTribIPI.Descricao

                    Dim SitTribPIS As New [Lib].Negocio.SituacaoTributariaPISCOFINS(enc.SituacaoTributariaPISCOFINS)
                    txtNovaSituacaoTributariaPISCOFINS.Text = SitTribPIS.Codigo & " - " & SitTribPIS.Descricao
                ElseIf enc.Codigo.Contains("FACS") AndAlso enc.Sinal = "-" Then
                    totalEncNF += enc.Valor
                ElseIf enc.Codigo.Contains("FETHAB") AndAlso enc.Sinal = "-" Then
                    totalEncNF += enc.Valor
                ElseIf enc.Codigo.Contains("SENAR") AndAlso enc.Sinal = "-" Then
                    totalEncNF += enc.Valor
                ElseIf enc.Codigo.Contains("FUNRURAL") AndAlso enc.Sinal = "-" Then
                    totalEncNF += enc.Valor
                End If
            Next

            txtNovaNaturezaDaOperacao.Text = objNotaFiscal.NaturezaDaOperacao
            lblNovoTotalProduto.Text = objNotaFiscal.TotalProduto.ToString("N2")
            lblNovoTotalNota.Text = objNotaFiscal.TotalNota.ToString("N2")
            txtNovaOperacaoXEstado.Text = objNotaFiscal.Itens(gridItem.SelectedIndex).CodigoOperacaoEstado

            If objNotaFiscal.NossaEmissao Then
                Dim i As Integer = 0
                Dim j As Integer = 0
                Dim temEnc As Boolean
                Dim vlrOk As Boolean

                While j < gridEncargos.Rows.Count
                    i = 0
                    temEnc = False
                    vlrOk = False
                    While i < gridEncargosNovos.Rows.Count
                        If gridEncargosNovos.Rows(i).Cells(0).Text() = gridEncargos.Rows(j).Cells(0).Text() Then
                            temEnc = True
                            If CDec(CType(gridEncargosNovos.Rows(i).FindControl("txtValorEncargoItem"), TextBox).Text) = CDec(CType(gridEncargos.Rows(j).FindControl("txtValorEncargoItem"), TextBox).Text) Then
                                vlrOk = True
                            End If
                        End If
                        i += 1
                    End While

                    If temEnc = False AndAlso totalEncNF = 0 Then
                        If Not gridEncargos.Rows(j).Cells(0).Text().Contains("PIS") AndAlso
                            Not gridEncargos.Rows(j).Cells(0).Text().Contains("COFINS") AndAlso
                            Not gridEncargos.Rows(j).Cells(0).Text().Contains("FETHAB") AndAlso
                            Not gridEncargos.Rows(j).Cells(0).Text().Contains("FACS") AndAlso
                            Not gridEncargos.Rows(j).Cells(0).Text().Contains("SENAR") AndAlso
                            Not gridEncargos.Rows(j).Cells(0).Text().Contains("FUNRURAL") Then

                            MsgBox(Me.Page, "Encargo " & gridEncargos.Rows(j).Cells(0).Text() & " não foi encontrado na nova lista da Operação")
                            lnkAtualizar.Parent.Visible = False
                            Exit Sub
                        End If
                    ElseIf vlrOk = False AndAlso totalEncNF = 0 Then
                        If Not gridEncargos.Rows(j).Cells(0).Text().Contains("PIS") AndAlso
                            Not gridEncargos.Rows(j).Cells(0).Text().Contains("COFINS") AndAlso
                            Not gridEncargos.Rows(j).Cells(0).Text().Contains("FETHAB") AndAlso
                            Not gridEncargos.Rows(j).Cells(0).Text().Contains("FACS") AndAlso
                            Not gridEncargos.Rows(j).Cells(0).Text().Contains("SENAR") AndAlso
                            Not gridEncargos.Rows(j).Cells(0).Text().Contains("FUNRURAL") Then

                            MsgBox(Me.Page, "Valor do Encargo " & gridEncargos.Rows(j).Cells(0).Text() & " não é o mesmo na nova lista da Operação")
                            lnkAtualizar.Parent.Visible = False
                            Exit Sub
                        End If
                    End If

                    j += 1
                End While

                If txtSituacaoTributaria.Text <> txtNovaSituacaoTributaria.Text Then
                    MsgBox(Me.Page, "Situação Tributária ICMS da Operação não pode ser diferente da informada na Nota Eletrônica")
                    txtNovaSituacaoTributariaPISCOFINS.ForeColor = Drawing.Color.Red
                    lnkAtualizar.Parent.Visible = False
                    Exit Sub
                ElseIf txtSituacaoTributariaIPI.Text <> txtNovaSituacaoTributariaIPI.Text Then
                    MsgBox(Me.Page, "Situação Tributária IPI da Operação não pode ser diferente da informada na Nota Eletrônica")
                    txtNovaSituacaoTributariaIPI.ForeColor = Drawing.Color.Red
                    lnkAtualizar.Parent.Visible = False
                    Exit Sub
                ElseIf txtSituacaoTributariaPISCOFINS.Text <> txtNovaSituacaoTributariaPISCOFINS.Text Then
                    MsgBox(Me.Page, "Situação Tributária PIS/COFINS da Operação não pode ser diferente da informada na Nota Eletrônica")
                    txtNovaSituacaoTributariaPISCOFINS.ForeColor = Drawing.Color.Red
                ElseIf txtNaturezaDaOperacao.Text <> txtNovaNaturezaDaOperacao.Text Then
                    MsgBox(Me.Page, "CFOP da Operação não pode ser diferente do informado na Nota Eletrônica")
                    lnkAtualizar.Parent.Visible = False
                    Exit Sub
                ElseIf CDec(lblTotalProduto.Text) <> CDec(lblNovoTotalProduto.Text) Then
                    MsgBox(Me.Page, "Valor do Produto não pode ser diferente do informado na Nota Eletrônica")
                    lnkAtualizar.Parent.Visible = False
                    Exit Sub
                ElseIf CDec(lblTotalNota.Text) <> CDec(lblNovoTotalNota.Text) Then
                    If CDec(lblNovoTotalNota.Text) > CDec(lblTotalNota.Text) AndAlso (CDec(lblNovoTotalNota.Text) - CDec(lblTotalNota.Text) = totalEncNF) Then
                        'Deixa passar pela diferença ser facs, fethab, senar ou funrural
                    ElseIf CDec(lblNovoTotalNota.Text) < CDec(lblTotalNota.Text) AndAlso (CDec(lblTotalNota.Text) - CDec(lblNovoTotalNota.Text) = totalEncNF) Then
                        'Deixa passar pela diferença ser facs, fethab, senar ou funrural
                    Else
                        MsgBox(Me.Page, "Total da Nota Fiscal não pode ser diferente do informado na Nota Eletrônica")
                        lnkAtualizar.Parent.Visible = False
                        Exit Sub
                    End If
                End If
            Else
                For Each enc As [Lib].Negocio.NotaFiscalXItemXEncargo In objNotaFiscal.Itens(gridItem.SelectedIndex).Encargos
                    If enc.Codigo = "ICMS" AndAlso enc.Valor > 0 AndAlso enc.SituacaoTributaria <> 0 AndAlso enc.SituacaoTributaria <> 20 AndAlso enc.SituacaoTributaria <> 51 Then
                        MsgBox(Me.Page, "Nota Fiscal com ICMS difere da Situação tributária informada")
                        txtNovaSituacaoTributaria.ForeColor = Drawing.Color.Red
                        lnkAtualizar.Parent.Visible = False
                        Exit Sub
                    ElseIf enc.Codigo.Contains("IPI") Then
                        If objNotaFiscal.EntradaSaida = eEntradaSaida.Entrada AndAlso enc.Valor > 0 AndAlso enc.SituacaoTributariaIPI <> 0 Then
                            MsgBox(Me.Page, "Nota Fiscal com IPI difere da Situação tributária informada")
                            txtNovaSituacaoTributariaIPI.ForeColor = Drawing.Color.Red
                            lnkAtualizar.Parent.Visible = False
                            Exit Sub
                        ElseIf objNotaFiscal.EntradaSaida = eEntradaSaida.Saida AndAlso enc.Valor > 0 AndAlso enc.SituacaoTributariaIPI <> 50 Then
                            MsgBox(Me.Page, "Nota Fiscal com IPI difere da Situação tributária informada")
                            txtNovaSituacaoTributariaIPI.ForeColor = Drawing.Color.Red
                            lnkAtualizar.Parent.Visible = False
                            Exit Sub
                        ElseIf objNotaFiscal.EntradaSaida = eEntradaSaida.Entrada AndAlso enc.Valor = 0 AndAlso enc.SituacaoTributariaIPI = 0 Then
                            MsgBox(Me.Page, "Nota Fiscal sem IPI difere da Situação tributária informada")
                            txtNovaSituacaoTributariaIPI.ForeColor = Drawing.Color.Red
                            lnkAtualizar.Parent.Visible = False
                            Exit Sub
                        ElseIf objNotaFiscal.EntradaSaida = eEntradaSaida.Saida AndAlso enc.Valor = 0 AndAlso enc.SituacaoTributariaIPI = 50 Then
                            MsgBox(Me.Page, "Nota Fiscal sem IPI difere da Situação tributária informada")
                            txtNovaSituacaoTributariaIPI.ForeColor = Drawing.Color.Red
                            lnkAtualizar.Parent.Visible = False
                            Exit Sub
                        End If
                    ElseIf enc.Codigo.Contains("PIS") Then
                        If txtSituacaoTributariaPISCOFINS.Text <> txtNovaSituacaoTributariaPISCOFINS.Text Then
                            txtNovaSituacaoTributariaPISCOFINS.ForeColor = Drawing.Color.Red
                            MsgBox(Me.Page, "Situação Tributária PIS/COFINS da Operação é diferente da informada na Operação anterior, tenha certeza antes de ATUALIZAR para as novas configurações")
                        End If
                    End If
                Next
            End If

            If objNotaFiscal.SubOperacao.QuantidadeFisico AndAlso objNotaFiscal.Itens(0).Produto.ControlarPesagem AndAlso (objNotaFiscal.CodigoRomaneio = 0 OrElse (objNotaFiscal.CodigoRomaneio > 0 AndAlso objNotaFiscal.Romaneio.Processo = "NOTA FISCAL")) Then
                If objNotaFiscal.CodigoRomaneio = 0 Then
                    objNotaFiscal.CriarRomaneio = True
                End If
                SalvaNotaFiscal()

                BuscaClassificacao()

            ElseIf objNotaFiscal.SubOperacao.QuantidadeFisico Then
                SalvaNotaFiscal()
                carregarItem()
            Else
                SalvaNotaFiscal()
                carregarItem()
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkConsultar_Click(sender As Object, e As EventArgs) Handles lnkConsultar.Click
        Try
            RecuperaNotaFiscal()
            Dim Empresa() As String = ddlEmpresa.SelectedValue.ToString.Split("-")
            If String.IsNullOrWhiteSpace(ddlEmpresa.SelectedValue) Then
                MsgBox(Me.Page, "É necessário selecionar a empresa!")
                Exit Sub
            End If

            objNotaFiscal.CodigoEmpresa = Empresa(0)
            objNotaFiscal.EnderecoEmpresa = Empresa(1)
            objNotaFiscal.DataNota = CDate(txtDataInicial.Text)
            objNotaFiscal.Movimento = CDate(txtDataFinal.Text)

            If txtCodigoCliente.Value.ToString.Length > 0 Then
                Dim Cliente() As String = txtCodigoCliente.Value.ToString.Split("-")
                objNotaFiscal.CodigoCliente = Cliente(0)
                objNotaFiscal.EnderecoCliente = Cliente(1)
            End If

            If txtES.Text.Length > 0 Then objNotaFiscal.EntradaSaida = IIf(txtES.Text = "E", eEntradaSaida.Entrada, eEntradaSaida.Saida)
            If txtSerie.Text.Length > 0 Then objNotaFiscal.Serie = txtSerie.Text
            If txtNota.Text.Length > 0 Then objNotaFiscal.Codigo = txtNota.Text

            Session("ssCampo" & HID.Value) = "NotaXItens"
            Popup.ConsultaDePedidosXNotas(Me.Page, "objNFConsultaNXI" & HID.Value)
            ucConsultaPedidosXNotas.BindGridView()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAtualizar_Click(sender As Object, e As EventArgs) Handles lnkAtualizar.Click
        Try
            If validarCampos() Then
                RecuperaNotaFiscal()

                If objNotaFiscal.SubOperacao.QuantidadeFisico AndAlso objNotaFiscal.Itens(0).Produto.ControlarPesagem AndAlso (objNotaFiscal.CodigoRomaneio = -1 OrElse objNotaFiscal.CodigoRomaneio = 0 OrElse (objNotaFiscal.CodigoRomaneio > 0 AndAlso objNotaFiscal.Romaneio.Processo = "NOTA FISCAL")) Then
                    If objNotaFiscal.CodigoRomaneio = -1 Or objNotaFiscal.CodigoRomaneio = 0 Then
                        objNotaFiscal.CriarRomaneio = True
                    End If
                End If

                objNotaFiscal.NotaFiscalOriginal.TotalNotaValorModificado = objNotaFiscal.TotalNota
                objNotaFiscal.TotalNotaValorModificado = objNotaFiscal.TotalNota

                Dim Sqls As New ArrayList

                'ALTERAÇÃO CÓDIGO GRAVAÇÃO - FURLAN - 23/02-2015
                'EXCLUINDO
                Dim original As [Lib].Negocio.NotaFiscal = objNotaFiscal.NotaFiscalOriginal
                original.IUD = "D"
                For Each item In original.Itens
                    item.IUD = "D"
                Next
                original.CarregandoNota = True
                original.SalvaAlteracaoDaNotaFiscal(Sqls)

                'INCLUINDO NOVAMENTE

                objNotaFiscal.PesoBruto = original.PesoBruto
                objNotaFiscal.PesoLiquido = original.PesoLiquido
                objNotaFiscal.Quantidade = original.Quantidade
                objNotaFiscal.Marca = original.Marca
                objNotaFiscal.Numero = original.Numero
                objNotaFiscal.Especie = original.Especie

                If String.IsNullOrWhiteSpace(objNotaFiscal.ObservacoesControleInterno) Then
                    objNotaFiscal.ObservacoesControleInterno = "ALTERADO NO PROCESSO ALTERARNOTAFISCAL EM " & Now().ToString("yyyy-MM-dd HH:mm:ss")
                Else
                    objNotaFiscal.ObservacoesControleInterno = objNotaFiscal.ObservacoesControleInterno & ". ALTERADO NO PROCESSO ALTERARNOTAFISCAL EM " & Now().ToString("yyyy-MM-dd HH:mm:ss")
                End If

                'Atualiza a Operação e Finalide da Nota Fiscal
                Sql = " Update NotasFiscais set" & vbCrLf &
                      "     Finalidade                 = " & objNotaFiscal.CodigoFinalidade & vbCrLf &
                      "	   ,Operacao                   = " & objNotaFiscal.CodigoOperacao & vbCrLf &
                      "	   ,SubOperacao                = " & objNotaFiscal.CodigoSubOperacao & vbCrLf &
                      "	   ,UsuarioAlteracao           ='" & HttpContext.Current.Session("ssNomeUsuario") & "'" & vbCrLf &
                      "	   ,UsuarioAlteracaoData       ='" & Now().ToString("yyyy-MM-dd HH:mm:ss") & "'" & vbCrLf &
                      "	   ,ObservacoesControleInterno ='" & objNotaFiscal.ObservacoesControleInterno & "'" & vbCrLf &
                      "  Where Empresa_Id      ='" & objNotaFiscal.CodigoEmpresa & "'" & vbCrLf &
                      "    and EndEmpresa_Id   = " & objNotaFiscal.EnderecoEmpresa & vbCrLf &
                      "	   and Cliente_Id      ='" & objNotaFiscal.CodigoCliente & "'" & vbCrLf &
                      "    and EndCliente_Id   = " & objNotaFiscal.EnderecoCliente & vbCrLf &
                      "    and EntradaSaida_Id ='" & objNotaFiscal.EntradaSaida.ToString.Substring(0, 1) & "'" & vbCrLf &
                      "    and Nota_Id         = " & objNotaFiscal.Codigo & vbCrLf &
                      "    and Serie_Id        ='" & objNotaFiscal.Serie & "'"
                Sqls.Add(Sql)

                'Inclui Novamente as Relacionadas
                objNotaFiscal.IUD = "I"

                objNotaFiscal.CarregandoNota = True
                objNotaFiscal.SalvaAlteracaoDaNotaFiscal(Sqls)

                If Not objNotaFiscal.Romaneio Is Nothing Then
                    If Not objNotaFiscal.CriarRomaneio AndAlso objNotaFiscal.CodigoRomaneio > 0 Then
                        Sql = "Update Romaneios Set Operacao = " & objNotaFiscal.CodigoOperacao & ", SubOperacao = " & objNotaFiscal.CodigoSubOperacao &
                              " Where Empresa_id = '" & objNotaFiscal.CodigoEmpresa & "'" &
                              "   and EndEmpresa_Id = " & objNotaFiscal.EnderecoEmpresa &
                              "   and Romaneio_Id = " & objNotaFiscal.Romaneio.Codigo
                        Sqls.Add(Sql)
                    End If

                    If Not objNotaFiscal.Romaneio.Pesagens Is Nothing AndAlso objNotaFiscal.Romaneio.Pesagens.Count > 0 Then
                        Sql = "Update Pesagem Set Operacao = " & objNotaFiscal.CodigoOperacao & ", SubOperacao = " & objNotaFiscal.CodigoSubOperacao &
                              " Where Empresa_id = '" & objNotaFiscal.CodigoEmpresa & "'" &
                              "   and EndEmpresa_Id = " & objNotaFiscal.EnderecoEmpresa &
                              "   and Pesagem_Id = " & objNotaFiscal.Romaneio.Pesagens(0).CodigoPesagem &
                              "   and Sequencia_Id = 0 "
                        Sqls.Add(Sql)
                    End If
                End If

                'Fianceiro NF de Entrada
                If objNotaFiscal.VencimentosNota IsNot Nothing AndAlso objNotaFiscal.VencimentosNota.Count > 0 Then
                    If objNotaFiscal.EntradaSaida = eEntradaSaida.Entrada AndAlso objNotaFiscal.TotalNota <> objNotaFiscal.VencimentosNota.Sum(Function(x) x.ValorDoDocumento) Then
                        objNotaFiscal.VencimentosNota.ReajFinanceiro = New ReajusteFinanceiro(objNotaFiscal, False)
                        objNotaFiscal.VencimentosNota.ReajFinanceiro.ReajustaNotaDeEntrada()
                    ElseIf objNotaFiscal.EntradaSaida = eEntradaSaida.Saida AndAlso objNotaFiscal.TotalNota <> objNotaFiscal.VencimentosNota.Sum(Function(x) x.ValorDoDocumento) Then
                        objNotaFiscal.VencimentosNota.ReajFinanceiro = New ReajusteFinanceiro(objNotaFiscal, False)
                        objNotaFiscal.VencimentosNota.ReajFinanceiro.ReajustaNotaDeSaida()
                    End If
                End If

                'Atualiza o financeiro do pedido
                For Each tit In objNotaFiscal.VencimentosPedido.Where(Function(x) x.CodigoProvisao = eProvisao.Previsao Or x.CodigoProvisao = eProvisao.Provisao)
                    If tit.ValorDoDocumento = 0 Then
                        tit.IUD = "U"
                        tit.CodigoSituacao = eSituacao.Excluido
                    End If
                    tit.SalvarSql(Sqls, False)
                Next

                'Atualiza o financeiro da NF
                If Not objNotaFiscal.VencimentosNota Is Nothing AndAlso objNotaFiscal.VencimentosNota.Count > 0 Then
                    objNotaFiscal.VencimentosNota.NF = objNotaFiscal
                    For Each tit In objNotaFiscal.VencimentosNota.Where(Function(x) x.CodigoProvisao = eProvisao.Previsao Or x.CodigoProvisao = eProvisao.Provisao)
                        If tit.ValorDoDocumento = 0 Then
                            tit.CodigoSituacao = eSituacao.Excluido
                        End If
                        If tit.IUD Is Nothing Then
                            tit.IUD = "I"
                        End If
                        tit.SalvarSql(Sqls, False)
                        If tit.IUD = "I" Then Sqls.Add(tit.AddNotaxTituloSql(tit.Codigo, tit.IUD, objNotaFiscal))
                    Next
                End If

                If Banco.GravaBanco(Sqls) Then
                    Dim espelho As New NotaFiscalEspelho
                    espelho.ExibirEspelho(Me.Page, objNotaFiscal)
                    MsgBox(Me.Page, "Registro alterado com suscesso")
                    LimparCampos()
                Else
                    MsgBox(Me.Page, "Erro ao Salvar a Nota: " & Funcoes.EliminarCaracteresEspeciais(RTrim(HttpContext.Current.Session("ssMessage").ToString)) & ". Entre em contato com o Suporte de TI")
                End If
            End If
        Catch ex As Exception
            MsgBox(Me, ex.Message)
        End Try
    End Sub

    Protected Sub lnkLimpar_Click(sender As Object, e As EventArgs) Handles lnkLimpar.Click
        Try
            LimparCampos()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "AlterarNotaFiscal")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub

    Protected Sub BtnVencimentos_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnVencimentos.Click
        Try
            RecuperaNotaFiscal()

            If objNotaFiscal.Codigo = 0 Then
                MsgBox(Me.Page, "Número da Nota não foi informado")
            ElseIf objNotaFiscal.Serie.Length = 0 Then
                MsgBox(Me.Page, "Série da Nota não foi informada")
            ElseIf Not objNotaFiscal.SubOperacao.Devolucao Then
                If objNotaFiscal.Pedido.MomentoFinanceiro <> 3 AndAlso objNotaFiscal.IUD = "U" AndAlso Not objNotaFiscal.SubOperacao.Financeiro Then
                    MsgBox(Me.Page, "Momento Financeiro pelo Pedido não pode ser alterado na Nota Fiscal")
                    Exit Sub
                End If

                If objNotaFiscal.Cliente.ApenasAVista AndAlso Not objNotaFiscal.Pedido.CondicaoPagamento.AVista Then
                    MsgBox(Me.Page, "Para Cliente com Condição de Pagamento apenas a Vista, o Pedido deve sestar na mesma condição.")
                    Exit Sub
                End If

                If objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D").Sum(Function(s) s.ValorLiquido) = 0 Then
                    MsgBox(Me.Page, "Valor Liquido da Nota Fiscal não pode ser Zero.")
                    Exit Sub
                End If

                Dim parameters = New Dictionary(Of String, Object)
                parameters.Add("tipo", "FRT")
                parameters.Add("Origem", "NF")
                parameters.Add("Indice", objNotaFiscal.IndiceNota)
                ucVencimentos.Limpar()
                ucVencimentos.SetarHID(HID.Value)
                ucVencimentos.BindGridView(parameters)
                Popup.ConsultaDeVencimentos(Me.Page, "objVencimentosNxI" & HID.Value)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, "Erro ao gerar Finaceiro da Nota Fiscal")
            'lnkNovo.Parent.Visible = False
        End Try
    End Sub
End Class