Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis
Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.Sql
Imports System.Data.SqlTypes
Imports System.Drawing
Imports System.Xml
Imports System.IO


Public Class NotasFiscaisGerais
    Inherits BasePage

#Region "Retorno UC"
    Public Overrides Sub Carregar(ByVal obj As [Lib].Negocio.IBaseEntity)

        If Not Session("objCliente" & HID.Value) Is Nothing Then
            SessaoRecuperaNotaFiscal()
            Dim objCliente As [Lib].Negocio.Cliente = Session("objCliente" & HID.Value)
            Session.Remove("objCliente" & HID.Value)

            'Não estava atualizando o Cliente dos Titulos na Inclusão, usuário preenchia todas informações da nota e antes de gravar trocava o Cliente. Caso tenha baixa também não pode trocar - Furlan - 06/10/2016
            If objNotaFiscal.CodigoOperacao > 0 AndAlso objNotaFiscal.VencimentosNota IsNot Nothing AndAlso objNotaFiscal.VencimentosNota.Count > 0 Then
                For Each row In objNotaFiscal.VencimentosNota
                    If row.CodigoProvisao = 1 Then
                        MsgBox(Me.Page, "Não pode ser alterado o Cliente com Financeiro Baixado - Título " & row.Codigo)
                        LimparCampos()
                        Exit Sub
                    Else
                        If Not objNotaFiscal.SubOperacao.Financeiro Then row.IUD = "D"

                        row.CodigoUnidadeDeNegocio = objNotaFiscal.CodigoUnidadeDeNegocio
                        row.CodigoEmpresa = objNotaFiscal.CodigoEmpresa
                        row.EnderecoEmpresa = objNotaFiscal.EnderecoEmpresa

                        row.CodigoEmpresaPedido = objNotaFiscal.CodigoEmpresa
                        row.EndEmpresaPedido = objNotaFiscal.EnderecoEmpresa
                        row.CodigoPedido = objNotaFiscal.CodigoPedido

                        row.CodigoCliente = objCliente.Codigo
                        row.EndCliente = objCliente.CodigoEndereco

                        row.CodigoDestinatario = objCliente.Codigo
                        row.EndDestinatario = objCliente.CodigoEndereco
                        row.NomeDoDestinatario = ""

                        row.Movimento = objNotaFiscal.Movimento

                        row.Historico = "PAGTO " & Trim(objNotaFiscal.TipoDeDocumento.Historico) & " " & Trim(objNotaFiscal.Codigo) & "-" & Trim(objNotaFiscal.Serie) & " - " & Trim(objCliente.Nome) & " - " & Trim(row.Carteira.Descricao)
                        row.Observacoes = objNotaFiscal.Observacoes
                    End If
                Next
            End If

            Dim itemCliente As ListItem = Funcoes.FormatarListItemCliente(objCliente)
            hdfCodigoCliente.Value = itemCliente.Value
            txtNomeCliente.Text = itemCliente.Text

            If Not objNotaFiscal.Itens Is Nothing AndAlso objNotaFiscal.Itens.Count > 0 AndAlso objNotaFiscal.Cliente.CodigoEstado <> objCliente.CodigoEstado Then

                Dim fCliente As String = objCliente.Codigo
                Dim fEndCliente As Integer = objCliente.CodigoEndereco

                If objNotaFiscal.CodigoTipoDeDocumento = 2 OrElse
                    objNotaFiscal.CodigoTipoDeDocumento = 10 OrElse
                    objNotaFiscal.CodigoTipoDeDocumento = 14 OrElse
                    objNotaFiscal.CodigoTipoDeDocumento = 57 OrElse
                    objNotaFiscal.CodigoTipoDeDocumento = 67 Then

                    objNotaFiscal.TipoDeDocumentoFrete = eTipoDeDocumentoFrete.Comprovacao

                    If Not objNotaFiscal.NotasTrocaOrigem Is Nothing AndAlso
                        objNotaFiscal.NotasTrocaOrigem.Count > 0 AndAlso
                        objCliente.CodigoEstado <> objNotaFiscal.NotasTrocaOrigem(0).Cliente.CodigoEstado Then
                        objNotaFiscal.CodigoCliente = objNotaFiscal.NotasTrocaOrigem(0).CodigoCliente
                        objNotaFiscal.EnderecoCliente = objNotaFiscal.NotasTrocaOrigem(0).EnderecoCliente
                    End If
                Else
                    objNotaFiscal.CodigoCliente = objCliente.Codigo
                    objNotaFiscal.EnderecoCliente = objCliente.CodigoEndereco
                End If

                For Each row In objNotaFiscal.Itens
                    If row.OperacaoEstado.Encargos.Count = 0 Then
                        MsgBox(Me.Page, "Não foram encontrados encargos do produto " & row.CodigoProduto & "-" & row.Produto.Nome & " para o cliente selecionado!")
                        LimparCampos()
                        Exit Sub
                    End If
                    row.CarregandoEncargos = True
                    If row.Encargos IsNot Nothing AndAlso row.Encargos.Count > 0 AndAlso row.Encargos.EncProduto.CentroDeCusto IsNot Nothing AndAlso row.Encargos.EncProduto.CentroDeCusto.Count > 0 Then
                        Dim cc As String = row.Encargos.EncProduto.CentroDeCusto
                        If Not String.IsNullOrWhiteSpace(cc) AndAlso cc <> "0" Then
                            row.Encargos = Nothing
                            row.Encargos.EncProduto.CentroDeCusto = cc
                            row.CarregandoEncargos = False
                            row.Encargos.AtualizaLiquido()
                        End If
                    End If
                Next
                SessaoSalvaNotaFiscal()
                AtualizarCFOP(False)
                SessaoRecuperaNotaFiscal()
                If objNotaFiscal.CodigoTipoDeDocumento = 2 OrElse
                    objNotaFiscal.CodigoTipoDeDocumento = 10 OrElse
                    objNotaFiscal.CodigoTipoDeDocumento = 14 OrElse
                    objNotaFiscal.CodigoTipoDeDocumento = 57 OrElse
                    objNotaFiscal.CodigoTipoDeDocumento = 67 Then
                    objNotaFiscal.TipoDeDocumentoFrete = Nothing


                    objNotaFiscal.CodigoCliente = fCliente
                    objNotaFiscal.EnderecoCliente = fEndCliente
                End If
            Else
                objNotaFiscal.CodigoCliente = objCliente.Codigo
                objNotaFiscal.EnderecoCliente = objCliente.CodigoEndereco
            End If

            SessaoSalvaNotaFiscal()
        ElseIf Session("objSubOperacaoNFG" & HID.Value) IsNot Nothing Then
            SessaoRecuperaNotaFiscal()
            Dim objSubOperacao As [Lib].Negocio.SubOperacao = CType(obj, [Lib].Negocio.SubOperacao)

            objNotaFiscal.CarregandoNota = True
            objNotaFiscal.CodigoOperacao = objSubOperacao.CodigoOperacao
            objNotaFiscal.CodigoSubOperacao = objSubOperacao.Codigo
            For Each item As [Lib].Negocio.NotaFiscalXItem In objNotaFiscal.Itens
                item.CodigoOperacao = objSubOperacao.CodigoOperacao
                item.CodigoSubOperacao = objSubOperacao.Codigo
            Next
            objNotaFiscal.CarregandoNota = False

            For Each nf As [Lib].Negocio.NotaFiscal In objNotaFiscal.NotasTrocaOrigem
                nf.CarregandoNota = True
                nf.CodigoOperacao = objSubOperacao.CodigoOperacao
                nf.CodigoSubOperacao = objSubOperacao.Codigo
                For Each item As [Lib].Negocio.NotaFiscalXItem In nf.Itens
                    item.CodigoOperacao = objSubOperacao.CodigoOperacao
                    item.CodigoSubOperacao = objSubOperacao.Codigo
                Next
                nf.CarregandoNota = False
            Next

            grdProdutos.DataSource = objNotaFiscal.Itens
            grdProdutos.DataBind()

            grdNotasFretes.DataSource = objNotaFiscal.NotasTrocaOrigem.Where(Function(s) s.IUD <> "D")
            grdNotasFretes.DataBind()

            SessaoSalvaNotaFiscal()
            Session.Remove("objSubOperacaoNFG" & HID.Value)
        ElseIf Session("objDadosBancariosNFG" & HID.Value) IsNot Nothing Then
            SessaoRecuperaNotaFiscal()
            Dim objConta As [Lib].Negocio.ClienteXContaBancaria = CType(Session("objDadosBancariosNFG" & HID.Value), [Lib].Negocio.ClienteXContaBancaria)
            lblBanco.Text = objConta.CodigoBanco & " | " & objConta.NomeBanco
            lblAgencia.Text = objConta.CodigoAgencia & "-" & objConta.DigitoAgencia & " | " & objConta.Praca
            lblContaCorrente.Text = objConta.ContaCorrente & "-" & objConta.DigitoConta
            Session.Remove("objDadosBancariosNFG" & HID.Value)

            If objNotaFiscal.VencimentosNota IsNot Nothing AndAlso objNotaFiscal.VencimentosNota.Count > 0 Then
                For Each row As [Lib].Negocio.Titulo In objNotaFiscal.VencimentosNota
                    row.CodigoBancoCliente = objConta.CodigoBanco
                    row.CodigoAgenciaCliente = objConta.CodigoAgencia
                    row.DigitoAgenciaCliente = objConta.DigitoAgencia
                    row.ContaCliente = objConta.ContaCorrente
                    row.DigitoContaCliente = objConta.DigitoConta
                Next
            End If
            SessaoSalvaNotaFiscal()

        ElseIf Session("objClienteNFReferencial" & HID.Value) IsNot Nothing Then 'Cliente para o pedido das notas fiscais referenciais de origem do RPA
            Dim objCliente As [Lib].Negocio.Cliente = CType(Session("objClienteNFReferencial" & HID.Value), [Lib].Negocio.Cliente)
            Dim itemCliente As ListItem = Funcoes.FormatarListItemCliente(objCliente)
            txtCodigoClienteNFReferencial.Value = itemCliente.Value
            txtPedidoNFReferencial.Text = String.Empty
            Session.Remove("objClienteNFReferencial" & HID.Value)
            ConsultarPedido()

        ElseIf Session("objPedidoNFReferencial" & HID.Value) IsNot Nothing Then 'Pedido para as Notas fiscais referenciais de origem do RPA
            Dim objPedido As [Lib].Negocio.Pedido = CType(Session("objPedidoNFReferencial" & HID.Value), [Lib].Negocio.Pedido)
            txtPedidoNFReferencial.Text = objPedido.Codigo
            If Not String.IsNullOrWhiteSpace(objPedido.Codigo) AndAlso Int32.TryParse(objPedido.Codigo, 0) Then
                lnkAddNota_Click(lnkAddNota, New EventArgs())
            End If
            Session.Remove("objPedidoNFReferencial" & HID.Value)
        ElseIf Session("objNotaFiscalReferencial" & HID.Value) IsNot Nothing Then
            SessaoRecuperaNotaFiscal()
            For Each objNotaFiscalReferencial In CType(obj, [Lib].Negocio.ListNotaFiscalReferencial)
                objNotaFiscalReferencial.Parent = objNotaFiscal.Itens(0)
                If Not objNotaFiscal.NotasReferenciais Is Nothing Then
                    objNotaFiscal.NotasReferenciais.Add(objNotaFiscalReferencial)
                Else
                    objNotaFiscal.NotasReferenciais = New ListNotaFiscalReferencial()
                    objNotaFiscal.NotasReferenciais.Add(objNotaFiscalReferencial)
                End If
            Next

            If Not objNotaFiscal.NotasReferenciais Is Nothing AndAlso objNotaFiscal.NotasReferenciais.Count > 0 Then
                txtPedidoNFReferencial.Text = objNotaFiscal.NotasReferenciais(0).ParentOrigem.NotaFiscal.CodigoPedido
                txtPedidoNFReferencial.ReadOnly = True
            End If

            grdNotasReferenciais.DataSource = objNotaFiscal.NotasReferenciais.Where(Function(s) s.IUD <> "D")
            grdNotasReferenciais.DataBind()

            Session.Remove("objNotaFiscalReferencial" & HID.Value)
            SessaoSalvaNotaFiscal()
        ElseIf Session("objNavioXInvoice" & HID.Value) IsNot Nothing Then
            Dim objNavioXInvoice = CType(Session("objNavioXInvoice" & HID.Value), [Lib].Negocio.NavioXInvoice)
            txtNaviosXInvoice.Text = objNavioXInvoice.Codigo & " - " & objNavioXInvoice.Descricao

        ElseIf Session("objRepresentante" & HID.Value) IsNot Nothing Then

            SessaoRecuperaNotaFiscal()

            objNotaFiscal.CodigoRepresentante = CType(Session("objRepresentante" & HID.Value), [Lib].Negocio.Cliente).Codigo
            objNotaFiscal.EndRepresentante = CType(Session("objRepresentante" & HID.Value), [Lib].Negocio.Cliente).CodigoEndereco

            If Not objNotaFiscal.CodigoRepresentante Is Nothing AndAlso objNotaFiscal.CodigoRepresentante.Length > 0 Then
                Funcoes.FormatarClienteTXT(txtRepresentante, objNotaFiscal.Representante)
            End If

            Session.Remove("objRepresentante" & HID.Value)
            SessaoSalvaNotaFiscal()

        Else
            lblBanco.Text = "Banco"
            lblAgencia.Text = "Agência"
            lblContaCorrente.Text = "Conta"
        End If
    End Sub

    Protected Sub lnkLimparRepresentante_Click(sender As Object, e As EventArgs) Handles lnkLimparRepresentante.Click
        Try

            txtRepresentante.Text = String.Empty
            SessaoRecuperaNotaFiscal()

            objNotaFiscal.CodigoRepresentante = String.Empty
            objNotaFiscal.EndRepresentante = 0

            SessaoSalvaNotaFiscal()

        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Public Sub AdicionarArquivoBD()
        SessaoRecuperaNotaFiscal()

        If Not Funcoes.VerificaAcesso(objNotaFiscal.CodigoEmpresa, objNotaFiscal.EnderecoEmpresa, objNotaFiscal.Movimento, "NOTAS FISCAIS") Then
            MsgBox(Me.Page, "Movimento da Nota Fiscal já fechado para esta data", eTitulo.Info)
        Else
            If objNotaFiscal.IUD = "U" Then
                ucFile.Salvar(objNotaFiscal.Arquivos)

                Dim Sqls As New ArrayList

                For Each arq In objNotaFiscal.Arquivos
                    If arq.IUD = "I" Then
                        arq.CodigoEmpresa = objNotaFiscal.CodigoEmpresa
                        arq.EnderecoEmpresa = objNotaFiscal.EnderecoEmpresa
                        arq.CodigoCliente = objNotaFiscal.CodigoCliente
                        arq.EnderecoCliente = objNotaFiscal.EnderecoCliente
                        arq.CodigoNota = objNotaFiscal.Codigo
                        arq.Serie = objNotaFiscal.Serie
                        arq.CodigoPedido = objNotaFiscal.CodigoPedido
                        arq.SalvarSql(Sqls)
                        arq.IUD = "U"
                    End If
                Next

                If Sqls.Count > 0 Then
                    If Banco.GravaBanco(Sqls) Then
                        SessaoSalvaNotaFiscal()
                        MsgBox(Me.Page, "Arquivo adicionado com sucesso. Caso tenha consultado a nota apenas para adicionar o arquivo, não precisa clicar em alterar pois o mesmo já foi adicionado.", eTitulo.Sucess)
                    Else
                        MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                    End If
                End If
            End If
        End If

    End Sub

    Public Sub RemoverArquivoBD()
        SessaoRecuperaNotaFiscal()

        If Not Funcoes.VerificaAcesso(objNotaFiscal.CodigoEmpresa, objNotaFiscal.EnderecoEmpresa, objNotaFiscal.Movimento, "NOTAS FISCAIS") Then
            MsgBox(Me.Page, "Movimento da Nota Fiscal já fechado para esta data", eTitulo.Info)
        Else
            If objNotaFiscal.IUD = "U" Then
                ucFile.Salvar(objNotaFiscal.Arquivos)

                Dim Sqls As New ArrayList

                For Each arq In objNotaFiscal.Arquivos
                    If arq.IUD = "D" Then
                        arq.CodigoEmpresa = objNotaFiscal.CodigoEmpresa
                        arq.EnderecoEmpresa = objNotaFiscal.EnderecoEmpresa
                        arq.CodigoCliente = objNotaFiscal.CodigoCliente
                        arq.EnderecoCliente = objNotaFiscal.EnderecoCliente
                        arq.CodigoNota = objNotaFiscal.Codigo
                        arq.Serie = objNotaFiscal.Serie
                        arq.CodigoPedido = objNotaFiscal.CodigoPedido
                        arq.SalvarSql(Sqls)
                    End If
                Next

                If Sqls.Count > 0 Then
                    If Banco.GravaBanco(Sqls) Then
                        SessaoSalvaNotaFiscal()
                        MsgBox(Me.Page, "Arquivo removido com sucesso. Caso tenha consultado a nota apenas para remover o arquivo, não precisa clicar em alterar pois o mesmo já foi removido.", eTitulo.Sucess)
                    Else
                        MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                    End If
                End If
            End If
        End If
    End Sub

    Public Sub AtualizarProdutoXML()

        If Not SessaoDsXML Is Nothing And Not objProdutoNota Is Nothing And 1 = 2 Then

            If Not objProdutoNota.Produto Is Nothing Then

                For i As Integer = 0 To objNotaFiscal.Itens.Count - 1

                    If objNotaFiscal.Itens(0).Produto.NCM = objProdutoNota.Produto.NCM Then

                        If objNotaFiscal.Itens(i).CodigoProduto = "0" Then
                            objNotaFiscal.Itens(i).CodigoProduto = objProdutoNota.CodigoProduto
                            objNotaFiscal.Itens(i).EmbalagemProduto = objProdutoNota.EmbalagemProduto
                            objNotaFiscal.Itens(i).IndiceProdutoNota = objProdutoNota.IndiceProdutoNota
                            objNotaFiscal.Itens(i).ObservacoesDoProduto = objProdutoNota.ObservacoesDoProduto
                            objNotaFiscal.Itens(i).Produto = objProdutoNota.Produto
                            objNotaFiscal.Itens(i).PesoQuantidade = objProdutoNota.PesoQuantidade
                        End If

                        If objNotaFiscal.Itens(i).CodigoOperacao = "0" Then
                            objNotaFiscal.Itens(i).CodigoOperacao = objProdutoNota.CodigoOperacao
                            objNotaFiscal.Itens(i).CodigoOperacaoEstado = objProdutoNota.CodigoOperacaoEstado
                            objNotaFiscal.Itens(i).CodigoSubOperacao = objProdutoNota.CodigoSubOperacao
                            objNotaFiscal.Itens(i).Operacao = objProdutoNota.Operacao
                            objNotaFiscal.Itens(i).OperacaoEstado = objProdutoNota.OperacaoEstado
                            objNotaFiscal.Itens(i).SubOperacao = objProdutoNota.SubOperacao
                        End If

                        objNotaFiscal.Itens(i).ValorTotal = objNotaFiscal.Itens(i).Unitario * objNotaFiscal.Itens(i).QuantidadeFiscal

                    End If

                Next

            End If

        End If

    End Sub

    Public Function VerificarProdutos() As Boolean

        Return grdProdutos.Rows.Count > 0

    End Function

    Public Sub AtualizarItensNoGrid()

        grdProdutos.Columns(6).Visible = True

        If objNotaFiscal Is Nothing Then SessaoRecuperaNotaFiscal()
        SessaoRecuperaProdutoNota()

        Dim objEmpresaXCliente As New [Lib].Negocio.ClienteXEmpresa(objNotaFiscal.CodigoEmpresa, objNotaFiscal.EnderecoEmpresa)
        If objEmpresaXCliente IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(objEmpresaXCliente.CodigoContaGrupoComissoes) Then
            lnkComissoes.Parent.Visible = objNotaFiscal.Itens.Any(Function(s) s.SubOperacao IsNot Nothing AndAlso s.SubOperacao.CodigoGrupoContas = objEmpresaXCliente.CodigoContaGrupoComissoes)
        End If

        grdProdutos.DataSource = objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D").ToList.OrderBy(Function(s) s.Sequencia)
        grdProdutos.DataBind()

        For ni As Integer = 0 To objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D").Count - 1
            If objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(ni).Encargos.Count > 0 Then
                If objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(ni).Encargos.EncProduto.CentroDeCusto = 0 Then
                    CType(grdProdutos.Rows(ni).FindControl("btnRateio"), Button).Visible = True

                Else
                    CType(grdProdutos.Rows(ni).FindControl("btnRateio"), Button).Visible = False

                End If
            End If
        Next

        gridEncargosGerais.DataSource = From nfEnc In objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D").SelectMany(Function(s) s.Encargos)
                                        Group By nfEnc.Codigo Into ValorOficial = Sum(nfEnc.Valor)
                                        Order By IIf(Codigo = "PRODUTO", 1, IIf(Codigo = "LIQUIDO", 3, 2))
                                        Select Codigo, ValorOficial
        gridEncargosGerais.DataBind()

        If objNotaFiscal.Itens IsNot Nothing AndAlso objNotaFiscal.Itens.Count > 0 Then

            objNotaFiscal.AtualizaTotais()

            If objNotaFiscal.CodigoOperacao > 0 AndAlso objNotaFiscal.TotalNota > 0 AndAlso objNotaFiscal.SubOperacao.Financeiro Then
                TabVencimentosold.Visible = True
            Else
                TabVencimentosold.Visible = False
            End If

            If objNotaFiscal.SubOperacao IsNot Nothing AndAlso objNotaFiscal.SubOperacao.Financeiro Then

                Dim TotalPago As Decimal = objNotaFiscal.VencimentosNota.Where(Function(s) s.CodigoProvisao = 1).Sum(Function(s) s.ValorDoDocumento)
                If Not objNotaFiscal.VencimentosNota Is Nothing AndAlso
                    objNotaFiscal.VencimentosNota.Count > 0 Then

                    grdCondicoes.DataSource = objNotaFiscal.VencimentosNota.Where(Function(s) s.IUD <> "D")
                    grdCondicoes.DataBind()

                    For x As Integer = 0 To objNotaFiscal.VencimentosNota.Count - 1
                        If objNotaFiscal.VencimentosNota(x).CodigoProvisao = 1 Then
                            grdCondicoes.Rows(x).Cells(1).ForeColor = Drawing.Color.Red
                            grdCondicoes.Rows(x).Cells(2).ForeColor = Drawing.Color.Red
                            grdCondicoes.Rows(x).Cells(3).ForeColor = Drawing.Color.Red
                            grdCondicoes.Rows(x).Cells(4).ForeColor = Drawing.Color.Red
                            grdCondicoes.Rows(x).Cells(5).ForeColor = Drawing.Color.Red
                            grdCondicoes.Rows(x).Cells(6).ForeColor = Drawing.Color.Red
                            grdCondicoes.Rows(x).Cells(7).ForeColor = Drawing.Color.Red
                            grdCondicoes.Rows(x).Cells(8).ForeColor = Drawing.Color.Red
                            grdCondicoes.Rows(x).Cells(9).ForeColor = Drawing.Color.Red
                        End If
                    Next
                End If

                txtTotalNota.Text = objNotaFiscal.TotalNota.ToString("N2")
                txtTotalParcelado.Text = objNotaFiscal.VencimentosNota.Where(Function(s) s.IUD <> "D").Sum(Function(s) s.ValorDoDocumento).ToString("N2")
                txtTotalPago.Text = TotalPago.ToString("N2")
                txtSaldoVencimentos.Text = (CDec(txtTotalParcelado.Text) - TotalPago).ToString("N2")
            End If

            For Each row As GridViewRow In grdProdutos.Rows
                If objNotaFiscal.Itens(row.RowIndex).Rateios IsNot Nothing AndAlso objNotaFiscal.Itens(row.RowIndex).Rateios.Count > 0 Then
                    CType(row.FindControl("btnRateio"), Button).BackColor = Drawing.Color.Green
                    CType(row.FindControl("btnRateio"), Button).ForeColor = Drawing.Color.White
                End If
            Next
        End If

        ddlNaturezaDeRendimento.Items.Clear()

        If objNotaFiscal.CodigoNaturezaDeRendimento > 0 Then
            chkNaturezaDeRendimento.Checked = True
            divNaturezaDeRendimento.Visible = True

            BuscarNaturezaDeRendimentos(objNotaFiscal)

            ddlNaturezaDeRendimento.SelectedValue = objNotaFiscal.CodigoNaturezaDeRendimento
        Else
            chkNaturezaDeRendimento.Checked = False
            divNaturezaDeRendimento.Visible = False
        End If

        If lnkNovo.Parent.Visible = False AndAlso grdProdutos.Rows.Count > 0 Then lnkNovo.Parent.Visible = objNotaFiscal.IUD <> "U" AndAlso objNotaFiscal.IUD <> "D"

        SessaoSalvaNotaFiscal()

        If objNotaFiscal.SubOperacao IsNot Nothing AndAlso objNotaFiscal.SubOperacao.Financeiro AndAlso grdCondicoes.Rows.Count = 0 AndAlso objNotaFiscal.Itens.Count() > 0 Then

            If objNotaFiscal.SubOperacao.EntradaSaida = eEntradaSaida.Entrada Then
                Dim codigoCarteira As String = objNotaFiscal.Itens(0).Produto.CodigoCarteiraCompra
                If Not String.IsNullOrWhiteSpace(objNotaFiscal.Itens(0).Produto.CodigoCarteiraCompra) Then
                    'Feito dessa forma para evitar o erro, caso o ddl não contenha a carteira
                    For Each item As ListItem In cmbCarteira.Items
                        If item.Value = objNotaFiscal.Itens(0).Produto.CodigoCarteiraCompra Then
                            cmbCarteira.SelectedValue = item.Value
                            Exit For
                        End If
                    Next
                End If
            Else
                If Not String.IsNullOrWhiteSpace(objNotaFiscal.Itens(0).Produto.CodigoCarteiraVenda) Then
                    'Feito dessa forma para evitar o erro, caso o ddl não contenha a carteira
                    For Each item As ListItem In cmbCarteira.Items
                        If item.Value = objNotaFiscal.Itens(0).Produto.CodigoCarteiraVenda Then
                            cmbCarteira.SelectedValue = item.Value
                            Exit For
                        End If
                    Next
                End If
            End If

            If Not String.IsNullOrWhiteSpace(cmbCarteira.SelectedValue) Then
                cmbCarteira_SelectedIndexChanged(cmbCarteira, Nothing)
            End If

        End If

        AtualizarCFOP(False)

    End Sub

    Public Sub AtualizarValorParcelamento()
        If objNotaFiscal.SubOperacao IsNot Nothing Then

            If objNotaFiscal.SubOperacao.Financeiro Then
                SessaoRecuperaNotaFiscal()

                If objNotaFiscal.TotalNota > 0 AndAlso TabVencimentosold.Visible = False Then
                    TabVencimentosold.Visible = True
                End If

                If Not objNotaFiscal.TotalNota = objNotaFiscal.VencimentosNota.Sum(Function(s) s.ValorDoDocumento) AndAlso objNotaFiscal.VencimentosNota.Count > 0 Then
                    If objNotaFiscal.VencimentosNota.All(Function(s) s.CodigoProvisao = 1) Then
                        MsgBox(Me.Page, "Atenção, valor da Nota Fiscal não pode ser alteração sem Financeiro estar em aberto")
                    Else
                        If FinanceiroVirtual Then
                            objNotaFiscal.VencimentosNota.ParcelarNotasFiscaisGerais(hdCondicaoDePagamento.Value, CDec(txtTotalNota.Text), CDec(txtTotalParcelado.Text), CDec(txtTotalPago.Text))
                        Else
                            objNotaFiscal.VencimentosNota.ParcelarNotasFiscaisGerais(hdCondicaoDePagamento.Value, CDec(txtTotalNota.Text), CDec(txtTotalParcelado.Text), CDec(txtTotalPago.Text), IIf(chkProvisao.Checked, 3, 2))
                        End If

                        grdCondicoes.DataSource = objNotaFiscal.VencimentosNota.Where(Function(s) s.IUD <> "D")
                        grdCondicoes.DataBind()

                        For x As Integer = 0 To objNotaFiscal.VencimentosNota.Count - 1
                            If objNotaFiscal.VencimentosNota(x).CodigoProvisao = 1 Then
                                grdCondicoes.Rows(x).Cells(1).ForeColor = Drawing.Color.Red
                                grdCondicoes.Rows(x).Cells(2).ForeColor = Drawing.Color.Red
                                grdCondicoes.Rows(x).Cells(3).ForeColor = Drawing.Color.Red
                                grdCondicoes.Rows(x).Cells(4).ForeColor = Drawing.Color.Red
                                grdCondicoes.Rows(x).Cells(5).ForeColor = Drawing.Color.Red
                                grdCondicoes.Rows(x).Cells(6).ForeColor = Drawing.Color.Red
                                grdCondicoes.Rows(x).Cells(7).ForeColor = Drawing.Color.Red
                                grdCondicoes.Rows(x).Cells(8).ForeColor = Drawing.Color.Red
                                grdCondicoes.Rows(x).Cells(9).ForeColor = Drawing.Color.Red
                            End If
                        Next

                        Dim TotalPago As Decimal = objNotaFiscal.VencimentosNota.Where(Function(s) s.CodigoProvisao = 1).Sum(Function(s) s.ValorDoDocumento)
                        txtTotalParcelado.Text = objNotaFiscal.VencimentosNota.Where(Function(s) s.IUD <> "D").Sum(Function(s) s.ValorDoDocumento).ToString("N2")
                        txtTotalPago.Text = TotalPago.ToString("N2")
                        txtSaldoVencimentos.Text = (CDec(txtTotalParcelado.Text) - TotalPago).ToString("N2")

                        SessaoSalvaNotaFiscal()
                        If objNotaFiscal.VencimentosNota.Count > 0 Then MsgBox(Me.Page, "Pela alteração de valores, verifique o Financeiro")
                    End If
                End If
            End If
        End If
    End Sub

    Public Sub AtualizarDadosInformados()

        SessaoRecuperaNotaFiscal()
        ImportarProdutoXML(objNotaFiscal, SessaoXmlDoc, True)
        CarregarFormComAClasse(objNotaFiscal, True)
        AdicionarProdutos(False, True)

    End Sub

    Public Sub CarregarNotaFiscal()
        If Session("objNFConsultaNFG" & HID.Value) IsNot Nothing Then
            Dim objNota As NotaFiscal = CType(Session("objNFConsultaNFG" & HID.Value), [Lib].Negocio.NotaFiscal)
            objNotaFiscalOriginal = New NotaFiscal(objNota)
            If Not objNotaFiscalOriginal.NotasTrocaOrigem Is Nothing AndAlso objNotaFiscalOriginal.NotasTrocaOrigem.Count > 0 Then divPedidoReferenciaRPA.Style.Add("display", "none")

            If objNotaFiscalOriginal.Empresa.Empresa.ControlaDataMovimentoNFG Then txtMovimento.Enabled = False

            SessaoSalvaNotaFiscalOriginal()

            objNota.IUD = "U"

            CarregarFormComAClasse(objNota)
        End If
    End Sub

    Public Sub CarregarNotasDeOrigem()
        SessaoRecuperaNotaFiscal()
        If objNotaFiscal.NotasTrocaOrigem IsNot Nothing AndAlso objNotaFiscal.NotasTrocaOrigem.Count > 0 Then
            grdNotasFretes.DataSource = objNotaFiscal.NotasTrocaOrigem
            grdNotasFretes.DataBind()
            TabContainer1.ActiveTabIndex = 0

            Dim obs As String = String.Empty

            If objNotaFiscal.CodigoTipoDeDocumento = CInt(eTipoDeDocumento.CTRC) OrElse
                objNotaFiscal.CodigoTipoDeDocumento = CInt(eTipoDeDocumento.CT_E) Then

                obs = "FRETE REF. NF(S). "
            End If

            If objNotaFiscal.CodigoTipoDeDocumento = CInt(eTipoDeDocumento.Estadia) Then
                obs = "ESTADIA REF. NF(S). "
            End If

            If objNotaFiscal.CodigoTipoDeDocumento = CInt(eTipoDeDocumento.ComplementoDeFrete) Then
                obs = "COMPLEMENTO DE FRETE REF. NF(S). "
            End If

            Dim separador As String = ""
            For Each item In objNotaFiscal.NotasTrocaOrigem
                obs += item.Codigo & "-" & item.Cliente.Nome & separador
                separador = ","
            Next

            txtObservacao.Text = obs
            TabNFOrigem.Visible = True

        End If
    End Sub

    'Carrega a NFe apartir do user control de arquivo.
    Public Overrides Sub Carregar(ByVal pString As String)

        SessaoRecuperaNotaFiscal()
        PreencherNFeXML(objNotaFiscal, pString, True)

        If objNotaFiscal.CodigoTipoDeDocumento = 57 Or objNotaFiscal.CodigoTipoDeDocumento = 58 Then
            AdicionarProdutos(False, False)
        Else
            AdicionarProdutos(True, False)
        End If

    End Sub
#End Region

#Region "Variáveis Locais"
    Private objNotaFiscalOriginal As NotaFiscal
    Private objNotaFiscal As NotaFiscal
    Private objProdutoNota As NotaFiscalXItem
    Private dblEncargos As Double
    Private dblLiquido As Double
    Private dblTotal As Double
    Private chaveXMLautomatico As String
    Private mensagemErro As String

    Dim logs As FuncoesLogs
#End Region

#Region "Inicialização"
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Expedicao)
        If IsConnect AndAlso Not IsPostBack Then

            If Not UsuarioServidor.KeyCodeActive Then
                MsgBox(Me.Page, "Sistema com chave de licença expirada. Entre em contato com o suporte.", "~/Expedicao.aspx", eTitulo.Info)
                Exit Sub
            End If

            If Funcoes.VerificaPermissao("NotasFiscaisGerais", "ACESSAR") Then
                ddl.Carregar(ddlUnidadeNegocio, CarregarDDL.Tabela.UnidadeDeNegocio, "", True)
                ddl.Carregar(ddlTipoDeDocumento, CarregarDDL.Tabela.TipoDeDocumento, " Codigo_Id <> 4", True)
                ddl.Carregar(ddlSituacao, CarregarDDL.Tabela.Situacao, "")
                ddl.Carregar(ddlCondicaoDePagamento, CarregarDDL.Tabela.CondicaoDePagamento, "", True)
                'ddl.Carregar(cmbCarteira, CarregarDDL.Tabela.CarteiraFinanceira, "Adiantamento = 'N' AND Produto_Id NOT IN(SELECT Carteira_Id from CarteirasXTributos where Carteira_Id = Produto_Id) ", True)
                ddl.Carregar(cmbCarteira, CarregarDDL.Tabela.CarteiraFinanceiraConta, "Cart.Adiantamento = 'N' AND Produto_Id NOT IN(SELECT Carteira_Id from CarteirasXTributos where Carteira_Id = Produto_Id) ", True)

                Dim Parametros As New Hashtable
                Parametros.Clear()
                Parametros.Add("listarTudo", "N")

                ddl.Carregar(cmbFormas, CarregarDDL.Tabela.TipoDePagamento, "", True, Parametros)

                ddl.Carregar(ddlSafra, CarregarDDL.Tabela.Safra, "", True)
                VerificaUnidade()

                If Not Session("chaveXMLautomacao") Is Nothing Then
                    chaveXMLautomatico = Session("chaveXMLautomacao" & HID.Value).ToString()
                End If

                LimparCampos()
                chkEspelho.Checked = True
                With ddlSafra
                    .SelectedIndex = .Items.IndexOf(.Items.FindByValue("NENHUMA"))
                End With
                'Para conferência de Documento Fiscal
                CarregarDocumento()
                carregarNotaXMLautomatico(sender, e)

                If Not Directory.Exists("C:\NGS\Log") Then Directory.CreateDirectory("C:\NGS\Log")

            Else
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Expedicao.aspx")
                Exit Sub
            End If
        End If
    End Sub

    Private Sub VerificaUnidade()
        Dim Sql As String = ""
        Sql = "SELECT Top 1 isnull(AcessoUnidade, '') as AcessoUnidade," & vbCrLf &
              "             isnull(AcessoEmpresa, '') as AcessoEmpresa," & vbCrLf &
              "             isnull(AcessoEndEmpresa, 0) as AcessoEndEmpresa" & vbCrLf &
              "  from Usuarios" & vbCrLf &
              " where Usuario_Id = '" & UsuarioServidor.NomeUsuario & "'"

        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Consulta").Tables(0).Rows
            ddlUnidadeNegocio.SelectedValue = Dr("AcessoUnidade")
            ddl.Carregar(ddlEmpresa, CarregarDDL.Tabela.Empresas, ddlUnidadeNegocio.SelectedValue, True)
            ddlEmpresa.SelectedValue = Dr("AcessoEmpresa") & "-" & Dr("AcessoEndEmpresa")
        Next
    End Sub
#End Region

#Region "Sessões"
    Private Sub SessaoSalvaNotaFiscalOriginal()
        Session("objNotaFiscalOriginal" & HID.Value) = objNotaFiscalOriginal
    End Sub

    Private Sub SessaoRecuperaNotaFiscalOriginal()
        objNotaFiscalOriginal = CType(Session("objNotaFiscalOriginal" & HID.Value), [Lib].Negocio.NotaFiscal)
    End Sub

    Private Sub SessaoSalvaNotaFiscal()
        Session("objNotaFiscal" & HID.Value) = objNotaFiscal
    End Sub

    Private Sub SessaoRecuperaNotaFiscal()
        objNotaFiscal = CType(Session("objNotaFiscal" & HID.Value), [Lib].Negocio.NotaFiscal)
    End Sub

    Private Sub SessaoSalvaProdutoNota()
        Session("objProdutoNota" & HID.Value) = objProdutoNota
    End Sub

    Private Sub SessaoRecuperaProdutoNota()
        objProdutoNota = CType(Session("objProdutoNota" & HID.Value), [Lib].Negocio.NotaFiscalXItem)
    End Sub

    Public Property SessaoDsXML() As DataSet
        Get
            Return CType(Session("dsXML" & HID.Value), DataSet)
        End Get
        Set(ByVal value As DataSet)
            If value Is Nothing Then
                Session.Remove("dsXml" & HID.Value)
            Else
                Session("dsXml" & HID.Value) = value
            End If
        End Set
    End Property

    Public Property SessaoXmlDoc() As XmlDocument
        Get
            Return CType(Session("XmlDoc" & HID.Value), XmlDocument)
        End Get
        Set(ByVal value As XmlDocument)
            If value Is Nothing Then
                Session.Remove("XmlDoc" & HID.Value)
            Else
                Session("XmlDoc" & HID.Value) = value
            End If
        End Set
    End Property

#End Region

#Region "Cabeçalho"
    Protected Sub ddlUnidadeNegocio_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ddlUnidadeNegocio.SelectedIndexChanged
        SessaoRecuperaNotaFiscal()
        If ddlUnidadeNegocio.SelectedIndex > 0 Then
            ddl.Carregar(ddlEmpresa, CarregarDDL.Tabela.Empresas, ddlUnidadeNegocio.SelectedValue, True)
            objNotaFiscal.CodigoUnidadeDeNegocio = ddlUnidadeNegocio.SelectedValue
        Else
            ddlEmpresa.Items.Clear()
            objNotaFiscal.CodigoUnidadeDeNegocio = String.Empty
            objNotaFiscal.CodigoEmpresa = String.Empty
            objNotaFiscal.EnderecoEmpresa = 0
        End If
        SessaoSalvaNotaFiscal()
    End Sub

    Protected Sub ddlEmpresa_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ddlEmpresa.SelectedIndexChanged
        SessaoRecuperaNotaFiscal()

        If ddlEmpresa.SelectedIndex > 0 Then
            Dim strEmpresa As String() = ddlEmpresa.SelectedValue.Split("-")

            'Caso mude a Empresa tem que buscar numerador do pedido. Caso tenha baixa não pode trocar - Furlan - 06/10/2016
            If objNotaFiscal.CodigoOperacao > 0 AndAlso objNotaFiscal.VencimentosNota IsNot Nothing AndAlso objNotaFiscal.VencimentosNota.Count > 0 AndAlso Not objNotaFiscal.CodigoEmpresa = strEmpresa(0) Then
                objNotaFiscal.CodigoEmpresa = strEmpresa(0)
                objNotaFiscal.EnderecoEmpresa = strEmpresa(1)

                Dim objPedido As New Pedido(objNotaFiscal, hdCondicaoDePagamento.Value)

                'NUMERADOR DOS PEDIDOS
                Dim SqlN As String = "exec sp_Numerador '" & objPedido.CodigoEmpresa & "'," & objPedido.EnderecoEmpresa & "," & [Lib].Negocio.eTiposNumerador.Pedido & ""
                Dim dsN As New DataSet
                dsN = Banco.ConsultaDataSet(SqlN, "Numerador")

                Dim CodigoNumerador As Integer = 0
                If dsN IsNot Nothing AndAlso dsN.Tables(0).Rows.Count > 0 Then
                    CodigoNumerador = dsN.Tables(0).Rows(0).Item(0)
                End If

                If Not CodigoNumerador > 0 Then
                    MsgBox(Me.Page, "Numerador de Pedidos não cadastrado!")
                    Exit Sub
                End If

                objPedido.Codigo = CodigoNumerador
                objNotaFiscal.CodigoPedido = CodigoNumerador
                objNotaFiscal.Pedido = objPedido
                objNotaFiscal.CIFFOB = objPedido.FreteCIFFOB

                For Each row In objNotaFiscal.VencimentosNota
                    If row.CodigoProvisao = 1 Then
                        MsgBox(Me.Page, "Não pode ser alterado a Empresa com Financeiro Baixado - Título " & row.Codigo)
                        LimparCampos()
                        Exit Sub
                    Else
                        If Not objNotaFiscal.SubOperacao.Financeiro Then row.IUD = "D"

                        row.CodigoUnidadeDeNegocio = objNotaFiscal.CodigoUnidadeDeNegocio
                        row.CodigoEmpresa = objNotaFiscal.CodigoEmpresa
                        row.EnderecoEmpresa = objNotaFiscal.EnderecoEmpresa

                        row.CodigoEmpresaPedido = objNotaFiscal.CodigoEmpresa
                        row.EndEmpresaPedido = objNotaFiscal.EnderecoEmpresa
                        row.CodigoPedido = objNotaFiscal.CodigoPedido

                        row.CodigoCliente = objNotaFiscal.CodigoCliente
                        row.EndCliente = objNotaFiscal.EnderecoCliente

                        row.CodigoDestinatario = objNotaFiscal.CodigoCliente
                        row.EndDestinatario = objNotaFiscal.EnderecoCliente
                        row.NomeDoDestinatario = ""

                        row.Movimento = objNotaFiscal.Movimento

                        row.Historico = "PAGTO " & Trim(objNotaFiscal.TipoDeDocumento.Historico) & " " & Trim(objNotaFiscal.Codigo) & "-" & Trim(objNotaFiscal.Serie) & " - " & Trim(objNotaFiscal.Cliente.Nome) & " - " & Trim(row.Carteira.Descricao)
                        row.Observacoes = objNotaFiscal.Observacoes
                    End If
                Next
            Else
                objNotaFiscal.CodigoEmpresa = strEmpresa(0)
                objNotaFiscal.EnderecoEmpresa = strEmpresa(1)
            End If

            'Verifica se a empresa está habilitada para gravar arquivo
            Dim Empresa As New [Lib].Negocio.Cliente(UsuarioServidor.CodigoEmpresa, UsuarioServidor.EnderecoEmpresa)
            ucFile.Parent.Visible = Not String.IsNullOrWhiteSpace(Empresa.Empresa.PathDownloadNFe)
        Else
            objNotaFiscal.CodigoEmpresa = String.Empty
            objNotaFiscal.EnderecoEmpresa = 0
        End If

        SessaoSalvaNotaFiscal()
    End Sub

    Protected Sub btnConsultaClientes_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnConsultaClientes.Click
        ucConsultaClientes.Limpar()
        Popup.ConsultaDeClientes(Me.Page, "objCliente" & HID.Value, "txtNome")
    End Sub

    Protected Sub ddlTipoDeDocumento_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddlTipoDeDocumento.SelectedIndexChanged
        SessaoRecuperaNotaFiscal()

        If Left(objNotaFiscal.CodigoEmpresa, 8) = "04854422" Then
            If ddlTipoDeDocumento.SelectedValue = 2 Or
                ddlTipoDeDocumento.SelectedValue = 3 Or
                ddlTipoDeDocumento.SelectedValue = 4 Or
                ddlTipoDeDocumento.SelectedValue = 10 Or
                ddlTipoDeDocumento.SelectedValue = 11 Or
                ddlTipoDeDocumento.SelectedValue = 12 Then

                'MsgBox(Me.Page, "Tipo de documento só pode ser lançado pelo novo sistema de fretes!")
                'ddlTipoDeDocumento.SelectedIndex = 0
                'Exit Sub

                'APENAS PARA LANÇAMENTO DOS CONHECIMENTOS ABAIXO QUE SÃO REF. NOTAS DE 2014. - TEMPORÁRIO - REMOVO ATÉ O FINAL DE SEMANA
                If ddlTipoDeDocumento.SelectedValue = 8 AndAlso
                    Left(objNotaFiscal.CodigoEmpresa, 8) = "03961253" AndAlso
                    (txtNumeroNota.Text = "76989" OrElse
                     txtNumeroNota.Text = "76990" OrElse
                     txtNumeroNota.Text = "76991" OrElse
                     txtNumeroNota.Text = "76992" OrElse
                     txtNumeroNota.Text = "76993" OrElse
                     txtNumeroNota.Text = "76995" OrElse
                     txtNumeroNota.Text = "76959" OrElse
                     txtNumeroNota.Text = "76966" OrElse
                     txtNumeroNota.Text = "76976" OrElse
                     txtNumeroNota.Text = "76977" OrElse
                     txtNumeroNota.Text = "76984" OrElse
                     txtNumeroNota.Text = "76987" OrElse
                     txtNumeroNota.Text = "76988") Then
                Else
                    MsgBox(Me.Page, "Tipo de documento só pode ser lançado pelo novo sistema de fretes!")
                    ddlTipoDeDocumento.SelectedIndex = 0
                    Exit Sub
                End If
            End If
        End If

        'NUBA NÃO PODE MAIS LANÇAR FRETE POR AQUI - FURLAN - 16/06/2025
        If Left(objNotaFiscal.CodigoEmpresa, 8) = "53267147" AndAlso
            (ddlTipoDeDocumento.SelectedValue = 2 Or
            ddlTipoDeDocumento.SelectedValue = 3 Or
            ddlTipoDeDocumento.SelectedValue = 4 Or
            ddlTipoDeDocumento.SelectedValue = 8 Or
            ddlTipoDeDocumento.SelectedValue = 10 Or
            ddlTipoDeDocumento.SelectedValue = 11 Or
            ddlTipoDeDocumento.SelectedValue = 14 Or
            ddlTipoDeDocumento.SelectedValue = 57 Or
            ddlTipoDeDocumento.SelectedValue = 58 Or
            ddlTipoDeDocumento.SelectedValue = 67) Then

            MsgBox(Me.Page, "Tipo de documento só pode ser lançado pelo novo sistema de fretes!")
            ddlTipoDeDocumento.SelectedIndex = 0
            Exit Sub

        End If

        'ORIX NÃO PODE MAIS LANÇAR FRETE POR AQUI - FURLAN - 18/09/2025
        If Left(objNotaFiscal.CodigoEmpresa, 8) = "62747840" AndAlso
            (ddlTipoDeDocumento.SelectedValue = 2 Or
            ddlTipoDeDocumento.SelectedValue = 3 Or
            ddlTipoDeDocumento.SelectedValue = 4 Or
            ddlTipoDeDocumento.SelectedValue = 8 Or
            ddlTipoDeDocumento.SelectedValue = 10 Or
            ddlTipoDeDocumento.SelectedValue = 11 Or
            ddlTipoDeDocumento.SelectedValue = 14 Or
            ddlTipoDeDocumento.SelectedValue = 57 Or
            ddlTipoDeDocumento.SelectedValue = 58 Or
            ddlTipoDeDocumento.SelectedValue = 67) Then

            MsgBox(Me.Page, "Tipo de documento só pode ser lançado pelo novo sistema de fretes!")
            ddlTipoDeDocumento.SelectedIndex = 0
            Exit Sub

        End If

        'HORUS NÃO LANÇA FRETE POR AQUI - FURLAN - 22/01/2026
        If (Left(objNotaFiscal.CodigoEmpresa, 8) = "62780383" OrElse Left(objNotaFiscal.CodigoEmpresa, 8) = "63358210") AndAlso
            (ddlTipoDeDocumento.SelectedValue = 2 Or
            ddlTipoDeDocumento.SelectedValue = 3 Or
            ddlTipoDeDocumento.SelectedValue = 4 Or
            ddlTipoDeDocumento.SelectedValue = 8 Or
            ddlTipoDeDocumento.SelectedValue = 10 Or
            ddlTipoDeDocumento.SelectedValue = 11 Or
            ddlTipoDeDocumento.SelectedValue = 14 Or
            ddlTipoDeDocumento.SelectedValue = 57 Or
            ddlTipoDeDocumento.SelectedValue = 58 Or
            ddlTipoDeDocumento.SelectedValue = 67) Then

            MsgBox(Me.Page, "Tipo de documento só pode ser lançado pelo novo sistema de fretes!")
            ddlTipoDeDocumento.SelectedIndex = 0
            Exit Sub

        End If


        If String.IsNullOrWhiteSpace(objNotaFiscal.CodigoCliente) Then
            ddlTipoDeDocumento.Enabled = True
            ddlTipoDeDocumento.SelectedIndex = 0
            MsgBox(Me.Page, "É necessário selecionar um cliente!")
            Exit Sub
        End If

        If (ddlTipoDeDocumento.SelectedValue = 2 Or
            ddlTipoDeDocumento.SelectedValue = 3 Or
            ddlTipoDeDocumento.SelectedValue = 4 Or
            ddlTipoDeDocumento.SelectedValue = 8 Or
            ddlTipoDeDocumento.SelectedValue = 10 Or
            ddlTipoDeDocumento.SelectedValue = 11 Or
            ddlTipoDeDocumento.SelectedValue = 14 Or
            ddlTipoDeDocumento.SelectedValue = 57 Or
            ddlTipoDeDocumento.SelectedValue = 58 Or
            ddlTipoDeDocumento.SelectedValue = 67) AndAlso
            Not objNotaFiscal.Cliente.Tipos.Any(Function(s) s.CodigoTipo = eTipoCliente.Transportadores) Then
            MsgBox(Me.Page, "É necessário selecionar um Tipo de Cliente Transportador!")
            ddlTipoDeDocumento.SelectedIndex = 0
            Exit Sub
        End If

        objNotaFiscal.CodigoTipoDeDocumento = ddlTipoDeDocumento.SelectedValue

        If ddlTipoDeDocumento.SelectedValue = 2 OrElse
            ddlTipoDeDocumento.SelectedValue = 9 OrElse
            ddlTipoDeDocumento.SelectedValue = 10 OrElse
            ddlTipoDeDocumento.SelectedValue = 14 OrElse
            ddlTipoDeDocumento.SelectedValue = 57 Then
            TabNFOrigem.Visible = True
            grdNotasFretes.Visible = True
            grdNotasReferenciais.Visible = False
            divPedidoReferenciaRPA.Style.Add("display", "none")
        ElseIf ddlTipoDeDocumento.SelectedValue = CInt(eTipoDeDocumento.RPA) Then 'RPA
            TabNFOrigem.Visible = True
            grdNotasFretes.Visible = False
            grdNotasReferenciais.Visible = True
            divPedidoReferenciaRPA.Style.Remove("display")
        Else
            TabNFOrigem.Visible = False
        End If

        If objNotaFiscal.CodigoTipoDeDocumento = CInt(eTipoDeDocumento.CTRC) OrElse
            objNotaFiscal.CodigoTipoDeDocumento = CInt(eTipoDeDocumento.Estadia) OrElse
            objNotaFiscal.CodigoTipoDeDocumento = CInt(eTipoDeDocumento.ComplementoDeFrete) OrElse
            objNotaFiscal.CodigoTipoDeDocumento = CInt(eTipoDeDocumento.CT_E) Then
            TabContainer1.ActiveTabIndex = 2
            ucNFOrigem.Limpar()
            Popup.ConsultaNotaOrigem(Me.Page, "objOrigemNFG" & HID.Value)
        End If

        If ddlTipoDeDocumento.SelectedValue = 1 OrElse ddlTipoDeDocumento.SelectedValue = 2 OrElse ddlTipoDeDocumento.SelectedValue = 15 OrElse ddlTipoDeDocumento.SelectedValue = 57 OrElse ddlTipoDeDocumento.SelectedValue = 58 Then
            lnkVerificarChaveNFE.Visible = True
        Else
            lnkVerificarChaveNFE.Visible = False
        End If

        If ddlTipoDeDocumento.SelectedValue = 8 OrElse ddlTipoDeDocumento.SelectedValue = 9 OrElse ddlTipoDeDocumento.SelectedValue = 67 Then
            chkProvisao.Parent.Visible = True
        Else
            chkProvisao.Parent.Visible = False
            chkProvisao.Checked = False
        End If

        SessaoSalvaNotaFiscal()
    End Sub

    Protected Sub imgExtratoPedido_Click(sender As Object, e As System.Web.UI.ImageClickEventArgs) Handles imgExtratoPedido.Click
        SessaoRecuperaNotaFiscal()
        If String.IsNullOrWhiteSpace(objNotaFiscal.CodigoCliente) Then
            MsgBox(Me.Page, "Cliente não foi selecionado")
            Exit Sub
        End If
        'Carrega a nova ficha do pedido caso seja financeiro novo
        'If Not FinanceiroNovo Then
        '    Dim strQueryString As String = "?fim=" & DateTime.Now.ToString("dd/MM/yyyy")
        '    strQueryString &= "&empresa=" & objNotaFiscal.CodigoEmpresa & "-" & objNotaFiscal.EnderecoEmpresa
        '    strQueryString &= "&cliente=" & objNotaFiscal.Pedido.CodigoCliente & "-" & objNotaFiscal.Pedido.EnderecoCliente
        '    strQueryString &= "&pedido=" & objNotaFiscal.CodigoPedido
        '    strQueryString &= "&es=ES"

        '    ScriptManager.RegisterClientScriptBlock(Me, Me.GetType, "CarregarHTML", "window.open(""ExtratoPedidoHTML.aspx" & strQueryString & """, ""relatorio"", ""status=no, toolbar=yes, location=no, menubar=yes, resizable=yes, height=600, width=800, scrollbars=yes"");", True)
        'Else

        Extrato.Emitir(Me.Page, FinanceiroNovo, objNotaFiscal.CodigoEmpresa, objNotaFiscal.EnderecoEmpresa, "T", objNotaFiscal.CodigoPedido)
        'End If

    End Sub

    Protected Sub lnkVerificarChaveNFE_Click(sender As Object, e As EventArgs) Handles lnkVerificarChaveNFE.Click
        If String.IsNullOrWhiteSpace(txtChaveNFe.Text.Replace(".", "")) Then
            MsgBox(Me.Page, "Chave da Nota Fiscal não foi informada.")
        ElseIf txtChaveNFe.Text.Replace(".", "").Length <> 44 Then
            txtChaveNFe.Text = ""
            MsgBox(Me.Page, "Chave da nota eletrônica deve ter 44 dígitos.")
        ElseIf String.IsNullOrWhiteSpace(txtNumeroNota.Text) AndAlso Not ucFile.Parent.Visible Then
            MsgBox(Me.Page, "Número da Nota Fiscal deve ser Informado.")
        ElseIf String.IsNullOrWhiteSpace(txtSerie.Text) AndAlso Not ucFile.Parent.Visible Then
            MsgBox(Me.Page, "Série da Nota Fiscal deve ser Informada.")
        ElseIf String.IsNullOrWhiteSpace(hdfCodigoCliente.Value) AndAlso Not ucFile.Parent.Visible Then
            MsgBox(Me.Page, "Cliente da Nota Fiscal deve ser Informado.")
        ElseIf Not ucFile.Parent.Visible AndAlso (Not ddlTipoDeDocumento.SelectedValue = 1 And Not ddlTipoDeDocumento.SelectedValue = 15) Then
            MsgBox(Me.Page, "Validação só pode ser feita para Documento do Tipo Nota fiscal.")
        ElseIf String.IsNullOrWhiteSpace(txtDataNota.Text) AndAlso Not ucFile.Parent.Visible Then
            MsgBox(Me.Page, "Data da Nota Fiscal deve ser Informada.")
        Else
            SessaoRecuperaNotaFiscal()
            objNotaFiscal.ChaveNFE = txtChaveNFe.Text.Replace(".", "")

            objNotaFiscal.Codigo = CInt(Mid(objNotaFiscal.ChaveNFE, 26, 9))
            txtNumeroNota.Text = CInt(Mid(objNotaFiscal.ChaveNFE, 26, 9))

            objNotaFiscal.Serie = CInt(Mid(objNotaFiscal.ChaveNFE, 23, 3))
            txtSerie.Text = objNotaFiscal.Serie

            If Not String.IsNullOrWhiteSpace(txtObservacao.Text) Then objNotaFiscal.Observacoes = Funcoes.EliminarCaracteresEspeciaisNF(txtObservacao.Text)

            If Not String.IsNullOrWhiteSpace(txtMovimento.Text) Then objNotaFiscal.Movimento = CDate(txtMovimento.Text)

            Dim ModeloNFe As String = Mid(objNotaFiscal.ChaveNFE, 21, 2)

            Dim valida As Boolean = True

            If (Mid(objNotaFiscal.ChaveNFE, 7, 14) = "02935843000105" Or
                Mid(objNotaFiscal.ChaveNFE, 7, 14) = "01409655000180" Or
                Mid(objNotaFiscal.ChaveNFE, 7, 14) = "78393592000146" Or
                Mid(objNotaFiscal.ChaveNFE, 7, 14) = "82951310000156" Or
                Mid(objNotaFiscal.ChaveNFE, 7, 14) = "87958674000181" Or
                Mid(objNotaFiscal.ChaveNFE, 7, 14) = "58290502000184" Or
                Mid(objNotaFiscal.ChaveNFE, 7, 14) = "76416890000189" Or
                Mid(objNotaFiscal.ChaveNFE, 7, 14) = "03507415000578") Then
                valida = False
            End If

            'Ver com o Furlan
            'If String.IsNullOrWhiteSpace(objNotaFiscal.ChaveNFE) Then
            '    MsgBox(Me.Page, "Chave Eletrônica não do informada na Nota Fiscal.")
            'ElseIf valida AndAlso objNotaFiscal.CodigoCliente.Length = 11 AndAlso Not Mid(objNotaFiscal.ChaveNFE, 10, 11) = objNotaFiscal.CodigoCliente Then
            '    MsgBox(Me.Page, "Cliente na Chave Eletrônica diferente do informado na Nota Fiscal.")
            'ElseIf valida AndAlso ModeloNFe.Equals("55") AndAlso objNotaFiscal.CodigoCliente.Length = 14 AndAlso Not Mid(objNotaFiscal.ChaveNFE, 7, 14) = objNotaFiscal.CodigoCliente Then
            '    MsgBox(Me.Page, "Cliente na Chave Eletrônica diferente do informado na Nota Fiscal.")
            'ElseIf ModeloNFe.Equals("55") AndAlso (Not ddlTipoDeDocumento.SelectedValue = 1 And Not ddlTipoDeDocumento.SelectedValue = 15) Then
            '    MsgBox(Me.Page, "Validação só pode ser feita para Documento do Tipo Nota fiscal.")
            'ElseIf ModeloNFe.Equals("57") AndAlso Not (objNotaFiscal.CodigoTipoDeDocumento = 57 Or objNotaFiscal.CodigoTipoDeDocumento = 58) Then
            '    MsgBox(Me.Page, "Validação só pode ser feita para Documento do Tipo Conhecimento de Transporte.")
            'ElseIf Not ucFile.Parent.Visible AndAlso Not CInt(Left(objNotaFiscal.ChaveNFE, 2)) = objNotaFiscal.Cliente.Municipio.EstadoIbge Then
            '    MsgBox(Me.Page, "Estado do Cliente na Chave Eletrônica diferente do informado na Nota Fiscal.")
            'ElseIf Not ucFile.Parent.Visible AndAlso Not CInt(Mid(objNotaFiscal.ChaveNFE, 26, 9)) = CInt(txtNumeroNota.Text) Then
            '    MsgBox(Me.Page, "Número da Nota na Chave Eletrônica diferente do informado na Nota Fiscal.")
            'ElseIf Not ucFile.Parent.Visible AndAlso Not CInt(Mid(objNotaFiscal.ChaveNFE, 23, 3)) = CInt(txtSerie.Text) Then
            '    MsgBox(Me.Page, "Série da Nota na Chave Eletrônica diferente do informado na Nota Fiscal.")
            'ElseIf Not ucFile.Parent.Visible AndAlso Not Mid(objNotaFiscal.ChaveNFE, 3, 2) = String.Format("{0:yy}", CDate(txtDataNota.Text)) Then
            '    MsgBox(Me.Page, "Ano da Nota na Chave Eletrônica diferente do informado na Nota Fiscal.")
            'ElseIf Not ucFile.Parent.Visible AndAlso Not Mid(objNotaFiscal.ChaveNFE, 5, 2) = String.Format("{0:MM}", CDate(txtDataNota.Text)) Then
            '    MsgBox(Me.Page, "Mês da Nota na Chave Eletrônica diferente do informado na Nota Fiscal.")
            'Else
            Dim msgResultado As String = String.Empty
            '****************************** (INICIO) colocar em produção quando estiver configurado o BD de Arquivo ******************************

            'Realiza o manifesto da NFe
            If ucFile.Parent.Visible AndAlso (ModeloNFe.Equals("55") Or ModeloNFe.Equals("57")) Then '(Modelo: 55 NFe, 57 CTe)

                Dim tpExt As String = String.Empty

                If ModeloNFe.Equals("55") Then tpExt = "-nfe.xml"
                If ModeloNFe.Equals("57") Then tpExt = "-CTe.xml"

                'Download do Arquivo.
                Dim bytes As Byte() = New FilesManager().getFileXml(String.Format("{0}{1}", objNotaFiscal.ChaveNFE, tpExt))
                If bytes Is Nothing Then
                    If ModeloNFe.Equals("55") Then
                        MsgBox(Me.Page, "XML da NFe não foi encontrado, favor inserir manualmente.")
                    Else
                        MsgBox(Me.Page, "XML do CTe não foi encontrado, favor inserir manualmente.")
                    End If

                    Exit Sub
                Else
                    Dim caminhoArquivoFile As String = Server.MapPath(String.Format("~/Files/{0}{1}", objNotaFiscal.ChaveNFE, tpExt))
                    If Not File.Exists(caminhoArquivoFile) Then
                        System.IO.File.WriteAllBytes(caminhoArquivoFile, bytes)
                    End If
                End If

                If objNotaFiscal.Empresa.Empresa.NotaFiscalEletronica AndAlso ModeloNFe.Equals("55") AndAlso Not objNotaFiscal.TemConfirmacaoDaOperacao Then

                    Dim verStatusNFe As String = DocumentoEletronico.ConsultaNFEFornecedor(objNotaFiscal)
                    Dim statusNFE As String() = verStatusNFe.Split(";")

                    If statusNFE(0) = "100" Then
                        If Not DocumentoEletronico.ManifestoNFe(objNotaFiscal, eTipoManifesto.ConfirmacaoDaOperacao, msgResultado) Then
                            MsgBox(Me.Page, msgResultado)
                            Exit Sub
                        End If
                    ElseIf statusNFE(0) = "101" Then
                        MsgBox(Me.Page, "Nota Fiscal Cancelada pelo Fornecedor não pode ser utilizada.")
                        Exit Sub
                    Else
                        MsgBox(Me.Page, statusNFE(1))
                        Exit Sub
                    End If
                End If

                If bytes IsNot Nothing Then
                    Try
                        If bytes IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(objNotaFiscal.ChaveNFE) Then
                            Dim caminhoArquivo As String = Server.MapPath(String.Format("~/Files/{0}{1}", objNotaFiscal.ChaveNFE, tpExt))
                            If System.IO.File.Exists(caminhoArquivo) Then
                                System.IO.File.Delete(caminhoArquivo)
                            End If
                            System.IO.File.WriteAllBytes(caminhoArquivo, bytes)

                            Dim arquivo As String = objNotaFiscal.ChaveNFE & tpExt

                            If File.Exists(caminhoArquivo) Then
                                Dim sourceDir As String = Server.MapPath("~/Files/")
                                Dim backupDir As String = "C:/Alfasig/LeituraNFe/-download/"

                                If Not File.Exists(backupDir & arquivo) Then
                                    File.Copy(Path.Combine(sourceDir, arquivo), Path.Combine(backupDir, arquivo))
                                End If
                            End If
                        End If

                        Dim temArquivo As Boolean = False

                        If objNotaFiscal.Arquivos IsNot Nothing AndAlso objNotaFiscal.Arquivos.Count > 0 Then
                            For Each arq In objNotaFiscal.Arquivos
                                If arq.Descricao = String.Format("{0}{1}", objNotaFiscal.ChaveNFE, tpExt) Then
                                    temArquivo = True
                                End If
                            Next
                        End If

                        If Not temArquivo Then
                            objNotaFiscal.Arquivos.Add(New [Lib].Negocio.Arquivo() With {
                             .IUD = "I",
                             .Codigo = String.Empty,
                             .Descricao = String.Format("{0}{1}", objNotaFiscal.ChaveNFE, tpExt),
                             .Arquivo = bytes})
                        End If


                        'Incluir a DANFE automaticamente
                        Dim msgPDF As String = ""
                        Dim bytesPDF As Byte()


                        If ModeloNFe.Equals("55") Then

                            If DocumentoEletronico.ImprimirNFeDanfe(objNotaFiscal, 1, msgPDF) Then
                                bytesPDF = New FilesManager().getFile(String.Format("{0}.pdf", objNotaFiscal.ChaveNFE), eTipoDeDocumento.Nota)
                            End If

                        ElseIf ModeloNFe.Equals("57") Then

                            If DocumentoEletronico.ImprimirCTeDacte(objNotaFiscal, 1, msgPDF) Then
                                bytesPDF = New FilesManager().getFile(String.Format("{0}.pdf", objNotaFiscal.ChaveNFE), eTipoDeDocumento.CT_E)
                            End If

                        End If

                        If bytesPDF IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(objNotaFiscal.ChaveNFE) Then

                            temArquivo = False

                            Dim caminhoArquivo As String = Server.MapPath(String.Format("~/Files/{0}.pdf", objNotaFiscal.ChaveNFE))
                            If System.IO.File.Exists(caminhoArquivo) Then
                                System.IO.File.Delete(caminhoArquivo)
                            End If
                            System.IO.File.WriteAllBytes(caminhoArquivo, bytesPDF)

                            If objNotaFiscal.Arquivos IsNot Nothing AndAlso objNotaFiscal.Arquivos.Count > 0 Then
                                For Each arq In objNotaFiscal.Arquivos
                                    If arq.Descricao = String.Format("{0}.pdf", objNotaFiscal.ChaveNFE) Then
                                        temArquivo = True
                                    End If
                                Next
                            End If

                            If Not temArquivo Then
                                objNotaFiscal.Arquivos.Add(New [Lib].Negocio.Arquivo() With {
                                                             .IUD = "I",
                                                             .Codigo = String.Empty,
                                                             .Descricao = String.Format("{0}.pdf", objNotaFiscal.ChaveNFE),
                                                             .Arquivo = bytesPDF})
                            End If

                        End If

                        ucFile.Bind(objNotaFiscal.Arquivos)

                        PreencherNFeXML(objNotaFiscal, String.Format("{0}{1}", objNotaFiscal.ChaveNFE, tpExt), False)
                        '****************************** (FIM) colocar em produção quando estiver configurado o BD de Arquivo ******************************

                        SessaoSalvaNotaFiscal()

                        If msgResultado.Length > 0 Then MsgBox(Me.Page, msgResultado, eTitulo.Info, False)
                    Catch ex As Exception
                        MsgBox(Me.Page, ex.Message.ToString())
                        If ex.Message = "Nota de origem não encontrada!" Then
                            ucNFOrigem.Limpar()
                            Popup.ConsultaNotaOrigem(Me.Page, "objOrigemNFG" & HID.Value)
                        End If
                        Exit Sub
                    End Try
                Else
                    MsgBox(Me.Page, "XML não foi encontrado, favor inserir o manualmente.")
                    Exit Sub
                End If
            Else
                If Not ModeloNFe.Equals("55") Then
                    MsgBox(Me.Page, "Manifesto permitido somente para Nota Fiscal")
                End If
                txtChaveNFe.Enabled = True
                txtNumeroNota.Enabled = True
                txtSerie.Enabled = True
                txtDataNota.Enabled = True
            End If
        End If
        'End If
    End Sub

    Protected Sub cmbCarteira_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbCarteira.SelectedIndexChanged
        ddlTributos.Items.Clear()

        If cmbCarteira.SelectedIndex > 0 Then
            ddl.Carregar(ddlTributos, CarregarDDL.Tabela.CarteirasXTributos, "Carteira_Id = '" & cmbCarteira.SelectedValue & "'", True)

            SessaoRecuperaNotaFiscal()

            If Not objNotaFiscal.VencimentosNota Is Nothing AndAlso objNotaFiscal.VencimentosNota.Count > 0 Then
                Dim objCarteira As New [Lib].Negocio.CarteiraFinanceira(cmbCarteira.SelectedValue)

                For Each rowTit As [Lib].Negocio.Titulo In objNotaFiscal.VencimentosNota
                    rowTit.CodigoCarteira = objCarteira.CodigoCarteira
                    rowTit.ContaContabilCliente = objCarteira.CodigoContaCliente
                Next

                SessaoSalvaNotaFiscal()

                grdCondicoes.DataSource = objNotaFiscal.VencimentosNota.Where(Function(s) s.IUD <> "D")
                grdCondicoes.DataBind()
            End If
        End If
    End Sub

    Protected Sub ddlTributos_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddlTributos.SelectedIndexChanged
        If ddlTributos.SelectedIndex > 0 Then
            SessaoRecuperaNotaFiscal()

            If Not objNotaFiscal.VencimentosNota Is Nothing AndAlso objNotaFiscal.VencimentosNota.Count > 0 Then
                Dim objTributo As New [Lib].Negocio.Encargo(ddlTributos.SelectedValue)

                For Each rowTit As [Lib].Negocio.Titulo In objNotaFiscal.VencimentosNota
                    rowTit.Tributo = objTributo.Codigo
                    rowTit.ContaContabilCliente = objTributo.ContaCredito
                Next

                SessaoSalvaNotaFiscal()

                grdCondicoes.DataSource = objNotaFiscal.VencimentosNota.Where(Function(s) s.IUD <> "D")
                grdCondicoes.DataBind()
            End If
        End If
    End Sub

    Protected Sub chkNaturezaDeRendimento_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            SessaoRecuperaNotaFiscal()

            ddlNaturezaDeRendimento.Items.Clear()

            If chkNaturezaDeRendimento.Checked Then

                If objNotaFiscal.CodigoCliente.Length = 0 Then
                    MsgBox(Me.Page, "Cliente não foi selecionado.")
                    chkNaturezaDeRendimento.Checked = False
                    Exit Sub
                End If

                divNaturezaDeRendimento.Visible = True

                BuscarNaturezaDeRendimentos(objNotaFiscal)
            Else
                divNaturezaDeRendimento.Visible = False
            End If

        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub chkImportarXML_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)

        Dim chk As CheckBox = CType(sender, CheckBox)

        If chk.ID = "chkImportarProdutoUnico" And chk.Checked Then
            chkAGruparNCM.Checked = False
        ElseIf chk.ID = "chkAGruparNCM" And chk.Checked Then
            chkImportarProdutoUnico.Checked = False
        ElseIf chk.ID = "chkInformarDados" And chk.Checked Then
            chkImportarProdutoUnico.Checked = False
            chkAGruparNCM.Checked = False
        End If

    End Sub

    Private Sub BuscarNaturezaDeRendimentos(ByVal objNota As NotaFiscal)
        Dim lstNaturezaDeRendimentos As New [Lib].Negocio.NaturezaDeRendimentos("Where TipoDePessoa in (3," & IIf(objNota.CodigoCliente.Length = 11, 1, 2) & ")")

        ddlNaturezaDeRendimento.Items.Add(New ListItem("", 0))

        For Each item In lstNaturezaDeRendimentos
            ddlNaturezaDeRendimento.Items.Add(New ListItem(item.Codigo & " - " & item.Descricao, item.Codigo))
        Next
    End Sub

    Protected Sub ddlNaturezaDeRendimento_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddlNaturezaDeRendimento.SelectedIndexChanged
        SessaoRecuperaNotaFiscal()

        objNotaFiscal.CodigoNaturezaDeRendimento = ddlNaturezaDeRendimento.SelectedValue

        SessaoSalvaNotaFiscal()
    End Sub

#End Region

#Region "TAB Produtos"
    Protected Sub lnkAdicionarItem_Click(sender As Object, e As EventArgs) Handles lnkAdicionarItem.Click

        AdicionarProdutos(True, Not SessaoDsXML Is Nothing)

    End Sub

    Protected Sub lnkComissoes_Click(sender As Object, e As EventArgs) Handles lnkComissoes.Click
        SessaoRecuperaNotaFiscal()
        ucComissoesXBaixas.Limpar()
        Dim parameters = New Dictionary(Of String, Object)
        parameters.Add("representanteId", objNotaFiscal.CodigoCliente)
        parameters.Add("enderecoRepId", objNotaFiscal.EnderecoCliente)
        parameters.Add("valorTotal", objNotaFiscal.Itens.Sum(Function(s) s.ValorTotal))
        ucComissoesXBaixas.BindGridView(parameters)
        Popup.ConsultaDeComissoesXBaixas(Me.Page, "objComissoesXBaixas" & HID.Value)
    End Sub

    Protected Sub btnEncargos_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim btn As Button = CType(sender, Button)
        Dim row As GridViewRow = CType(btn.NamingContainer, GridViewRow)
        ucNFEncargo.Limpar()
        ucNFEncargo.Inicializar(row.RowIndex)
        Popup.NFEncargo(Me.Page, "objNFEncargo" & HID.Value)
    End Sub

    Protected Sub btnRateio_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        SessaoRecuperaNotaFiscal()
        Dim btn As Button = CType(sender, Button)
        Dim row As GridViewRow = CType(btn.NamingContainer, GridViewRow)

        Dim i As Integer = objNotaFiscal.Itens.FindIndex(Function(s) s.CodigoProduto = row.Cells(1).Text)

        ucRateio.Limpar()
        ucRateio.SetarEmpresa(ddlUnidadeNegocio.SelectedValue, ddlEmpresa.SelectedValue)
        ucRateio.Inicializar(i)
        Popup.ConsultaRateio(Me.Page, "objRateioNFG" & HID.Value)
    End Sub

    Protected Sub lnkSelecionarProduto_Click(sender As Object, e As EventArgs)
        If (ddlTipoDeDocumento.SelectedValue = 2 OrElse ddlTipoDeDocumento.SelectedValue = 10 OrElse ddlTipoDeDocumento.SelectedValue = 14 OrElse ddlTipoDeDocumento.SelectedValue = 57) AndAlso grdNotasFretes.Rows.Count = 0 Then
            MsgBox(Me.Page, "Para lançamento de Conhecimento o Vinculo da Nota Fiscal deve ser selecionado antes.")
            Exit Sub
        ElseIf String.IsNullOrWhiteSpace(ddlTipoDeDocumento.SelectedValue) Then
            MsgBox(Me.Page, "É necessário selecionar o tipo de documento!")
            Exit Sub
        ElseIf String.IsNullOrWhiteSpace(txtNumeroNota.Text) Then
            MsgBox(Me.Page, "É necessário informar o número da nota!")
            Exit Sub
        ElseIf String.IsNullOrWhiteSpace(txtSerie.Text) Then
            MsgBox(Me.Page, "É necessário informar a série da nota!")
            Exit Sub
        ElseIf String.IsNullOrWhiteSpace(txtMovimento.Text) Then
            MsgBox(Me.Page, "É necessário informar a Data do Movimento!")
            Exit Sub
        ElseIf String.IsNullOrWhiteSpace(txtDataNota.Text) Then
            MsgBox(Me.Page, "É necessário informar a Data da Nota!")
            Exit Sub
        Else

            SessaoRecuperaNotaFiscal()

            Dim columnIndexProduto As Integer
            Dim columnIndexSequencia As Integer

            ' Itera pelas colunas da grid
            For iColumn As Integer = 0 To grdProdutos.Columns.Count - 1
                If grdProdutos.Columns(iColumn).HeaderText = "Produto" Then
                    columnIndexProduto = iColumn
                ElseIf grdProdutos.Columns(iColumn).HeaderText = "Sequência" Then
                    columnIndexSequencia = iColumn
                End If
            Next

            If Not SessaoDsXML Is Nothing Then

                Dim lnkProduto As LinkButton = CType(sender, LinkButton)
                Dim row As GridViewRow = CType(lnkProduto.NamingContainer, GridViewRow)
                Dim i As Integer
                i = objNotaFiscal.Itens.FindIndex(Function(s) s.CodigoProduto = row.Cells(columnIndexProduto).Text.Replace("&nbsp;", "") And s.Sequencia = row.Cells(columnIndexSequencia).Text)

                If i = -1 And objNotaFiscal.Itens.Count > 0 Then
                    i = row.RowIndex
                End If

                ucProdutoNFG.InicializarUC(i, SessaoDsXML)
                ucProdutoNFG.BloqueioEdicaoProdutoXML()
                Popup.ConsultaProdutoNFG(Me.Page, "objProdutoNotaNF" & HID.Value)

            Else

                Dim lnkProduto As LinkButton = CType(sender, LinkButton)
                Dim row As GridViewRow = CType(lnkProduto.NamingContainer, GridViewRow)
                Dim i As Integer = objNotaFiscal.Itens.FindIndex(Function(s) s.CodigoProduto = row.Cells(columnIndexProduto).Text And s.Sequencia = row.Cells(columnIndexSequencia).Text)
                ucProdutoNFG.InicializarUC(i, False, chkReaproveitar.Checked)

                If Not SessaoDsXML Is Nothing Then
                    ucProdutoNFG.BloqueioEdicaoProduto()
                End If

                Popup.ConsultaProdutoNFG(Me.Page, "objProdutoNotaNF" & HID.Value)

            End If


        End If
    End Sub

    Protected Sub imgObsProduto_Click(sender As Object, e As ImageClickEventArgs)
        Try
            SessaoRecuperaNotaFiscal()
            If Not objNotaFiscal Is Nothing Then
                Dim Imgproduto As ImageButton = CType(sender, ImageButton)
                Dim row As GridViewRow = CType(Imgproduto.NamingContainer, GridViewRow)
                ucNFObsProduto.Limpar()
                ucNFObsProduto.CarregarObs(row.RowIndex)
                Popup.NFObsProduto(Me.Page, "ObjNFObsProduto" & HID.Value)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub imgExcluir_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)
        SessaoRecuperaNotaFiscal()
        Dim img As ImageButton = CType(sender, ImageButton)
        Dim row As GridViewRow = CType(img.NamingContainer, GridViewRow)

        Dim i As Integer = objNotaFiscal.Itens.FindIndex(Function(s) s.CodigoProduto = row.Cells(1).Text.Replace("&nbsp;", "") And s.Sequencia = row.Cells(2).Text)

        If objNotaFiscal.IUD = "I" Or objNotaFiscal.Itens(i).IUD = "I" Then
            objNotaFiscal.Itens.Remove(objNotaFiscal.Itens(i))
        Else
            objNotaFiscal.Itens(i).IUD = "D"
        End If
        SessaoSalvaNotaFiscal()
        AtualizarItensNoGrid()
        If FinanceiroNovo Then AtualizarValorParcelamento()
    End Sub
#End Region

#Region "TAB NF Origem"
    Protected Sub lnkAddNota_Click(sender As Object, e As EventArgs) Handles lnkAddNota.Click
        If ddlTipoDeDocumento.SelectedValue = 13 Then
            SessaoRecuperaNotaFiscal()

            If objNotaFiscal.NotasReferenciais IsNot Nothing AndAlso objNotaFiscal.NotasReferenciais.Where(Function(s) s.IUD <> "D").Count > 0 Then
                MsgBox(Me.Page, "Já existem notas referenciais. Para prosseguir, estas devem ser excluídas!")
                'ElseIf String.IsNullOrWhiteSpace(txtPedidoNFReferencial.Text) Then
                '    MsgBox(Me.Page, "Indique o pedido que será utilizado como filtro para buscar as notas que serão vinculadas ao RPA.")
            Else
                Dim Empresa() As String = {objNotaFiscal.CodigoEmpresa, objNotaFiscal.EnderecoEmpresa}
                Dim Cliente() As String = txtCodigoClienteNFReferencial.Value.Split("-")

                Dim Parametros As New Dictionary(Of String, Object)()
                Parametros.Add("TipoDeDocumento", eTipoDeDocumento.Nota)
                Parametros.Add("Empresa", Empresa)
                Parametros.Add("Cliente", Cliente)
                Parametros.Add("TipoReferencial", eTipoReferencial.RPA)

                If Not String.IsNullOrWhiteSpace(txtPedidoNFReferencial.Text) Then Parametros.Add("Pedido", txtPedidoNFReferencial.Text)

                Parametros.Add("ValorNFOrigem", objNotaFiscal.TotalNotaBruto)

                Parametros.Add("DataInicio", CDate(txtDataInicio.Text).ToString("yyyy-MM-dd"))
                Parametros.Add("DataFim", CDate(txtDataFim.Text).ToString("yyyy-MM-dd"))

                ucNotaFiscalReferencial.Limpar()
                ucNotaFiscalReferencial.ConsultarNotas(Parametros)
                ucNotaFiscalReferencial.MainUserControl = Nothing

                Popup.ConsultaNotaFiscalReferencial(Me.Page, "objNotaFiscalReferencial" & HID.Value)
            End If
        Else
            ucNFOrigem.Limpar()
            Popup.ConsultaNotaOrigem(Me.Page, "objOrigemNFG" & HID.Value)
        End If
    End Sub

    Protected Sub imgExcluirNF_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)
        SessaoRecuperaNotaFiscal()
        Dim img As ImageButton = CType(sender, ImageButton)
        Dim row As GridViewRow = CType(img.NamingContainer, GridViewRow)
        objNotaFiscal.NotasTrocaOrigem(row.RowIndex).IUD = "D"
        grdNotasFretes.DataSource = objNotaFiscal.NotasTrocaOrigem.Where(Function(s) s.IUD <> "D")
        grdNotasFretes.DataBind()
        objNotaFiscal.Itens.ForEach(Function(nxi)
                                        nxi.QuantidadeFiscal = objNotaFiscal.NotasTrocaOrigem.Where(Function(s) s.IUD <> "D").SelectMany(Function(s) s.Itens).Sum(Function(s) s.QuantidadeFiscal)
                                        nxi.QuantidadeFisica = objNotaFiscal.NotasTrocaOrigem.Where(Function(s) s.IUD <> "D").SelectMany(Function(s) s.Itens).Sum(Function(s) s.QuantidadeFisica)
                                        Return True
                                    End Function)
        SessaoSalvaNotaFiscal()
    End Sub
#End Region

#Region "TAB Contabilização"
    Protected Sub ddlProdutoContabilizacao_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddlProdutoContabilizacao.SelectedIndexChanged
        SessaoRecuperaNotaFiscal()
        objNotaFiscal.LancamentosContabeis.CalcularSaldo()
        Select Case ddlProdutoContabilizacao.SelectedValue
            Case 0
                gridRazao.DataSource = objNotaFiscal.LancamentosContabeis.OrderBy(Function(s) s.Sequencia)
                gridRazao.DataBind()

                lblDebito.Text = objNotaFiscal.LancamentosContabeis.Sum(Function(s) s.DebitoOficial).ToString("N2")
                lblCredito.Text = objNotaFiscal.LancamentosContabeis.Sum(Function(s) s.CreditoOficial).ToString("N2")
            Case Else
                gridRazao.DataSource = objNotaFiscal.LancamentosContabeis.Where(Function(S) S.CodigoProdutoNF = ddlProdutoContabilizacao.SelectedValue).OrderBy(Function(s) s.Sequencia)
                gridRazao.DataBind()

                lblDebito.Text = objNotaFiscal.LancamentosContabeis.Where(Function(S) S.CodigoProdutoNF = ddlProdutoContabilizacao.SelectedValue).Sum(Function(s) s.DebitoOficial).ToString("N2")
                lblCredito.Text = objNotaFiscal.LancamentosContabeis.Where(Function(S) S.CodigoProdutoNF = ddlProdutoContabilizacao.SelectedValue).Sum(Function(s) s.CreditoOficial).ToString("N2")
        End Select
    End Sub
#End Region

#Region "TAB Vencimentos"
    Protected Sub LnkParcelamento_Click(sender As Object, e As EventArgs) Handles LnkParcelamento.Click

        If cmbCarteira.SelectedIndex = 0 Then
            MsgBox(Me.Page, "Carteira não foi selecionada.")
            Exit Sub
        ElseIf ddlCondicaoDePagamento.SelectedIndex = 0 Then
            MsgBox(Me.Page, "Condições de Pagamento não foi selecionado.")
            Exit Sub
        ElseIf cmbFormas.SelectedIndex = 0 Then
            MsgBox(Me.Page, "Forma de Pagamento não foi selecionada.")
            Exit Sub
        ElseIf CDec(txtTotalNota.Text) < CDec(txtTotalPago.Text) Then
            MsgBox(Me.Page, "Total da nota nao pode ser menor do que o total ja pago da mesma")
            Exit Sub
        End If

        Dim objCarteira As New [Lib].Negocio.CarteiraFinanceira(cmbCarteira.SelectedValue)
        Dim objTributo As New [Lib].Negocio.Encargo()

        If Not String.IsNullOrWhiteSpace(ddlTributos.SelectedValue) Then objTributo = New [Lib].Negocio.Encargo(ddlTributos.SelectedValue)

        If String.IsNullOrWhiteSpace(objCarteira.CodigoContaCliente) AndAlso String.IsNullOrWhiteSpace(objTributo.Codigo) Then
            MsgBox(Me.Page, "Encargo não foi selecionado.")
            Exit Sub
        End If

        SessaoRecuperaNotaFiscal()

        If objNotaFiscal.IUD = "I" Then
            cmbFormas.Enabled = True
        Else
            cmbFormas.Enabled = False
        End If

        objNotaFiscal.VencimentosNota.NF = objNotaFiscal
        hdCondicaoDePagamento.Value = ddlCondicaoDePagamento.SelectedValue

        Dim sReplicarHistorico As String = ""
        If objNotaFiscal.VencimentosNota.Count() = 1 Then
            sReplicarHistorico = objNotaFiscal.VencimentosNota.FirstOrDefault().Historico
        End If

        If FinanceiroVirtual Then
            objNotaFiscal.VencimentosNota.ParcelarNotasFiscaisGerais(hdCondicaoDePagamento.Value, CDec(txtTotalNota.Text), CDec(txtTotalParcelado.Text), CDec(txtTotalPago.Text))
        Else
            objNotaFiscal.VencimentosNota.ParcelarNotasFiscaisGerais(hdCondicaoDePagamento.Value, CDec(txtTotalNota.Text), CDec(txtTotalParcelado.Text), CDec(txtTotalPago.Text), IIf(chkProvisao.Checked, 3, 2))
        End If

        txtTotalParcelado.Text = objNotaFiscal.VencimentosNota.Where(Function(s) s.IUD <> "D").Sum(Function(s) s.ValorDoDocumento).ToString("N2")
        txtSaldoVencimentos.Text = CDec(txtTotalParcelado.Text) - CDec(txtTotalPago.Text)

        For Each rowTit As [Lib].Negocio.Titulo In objNotaFiscal.VencimentosNota
            rowTit.CodigoCarteira = objCarteira.CodigoCarteira
            rowTit.CodigoTipoPgto = cmbFormas.SelectedValue
            rowTit.Tributo = IIf(Not String.IsNullOrWhiteSpace(objTributo.Codigo) AndAlso objTributo.Codigo.Length > 0, objTributo.Codigo, "")
            If Not String.IsNullOrWhiteSpace(objTributo.Codigo) AndAlso objTributo.Codigo.Length > 0 Then
                rowTit.ContaContabilCliente = objTributo.ContaCredito
            Else
                rowTit.ContaContabilCliente = objCarteira.CodigoContaCliente
            End If
            If rowTit.Historico.Length = 0 Then
                rowTit.Historico = sReplicarHistorico
            End If
        Next

        SessaoSalvaNotaFiscal()

        grdCondicoes.DataSource = objNotaFiscal.VencimentosNota.Where(Function(s) s.IUD <> "D")
        grdCondicoes.DataBind()

        If cmbFormas.SelectedValue = 4 Then
            txtCodigoDeBarras.Enabled = True
            BtValidarCodBarras.Enabled = True
            CkbCodigoDeBarras.Enabled = True
            ckPreImpresso.Enabled = True
        Else
            txtCodigoDeBarras.Text = ""
            CkbCodigoDeBarras.Checked = False
            ckPreImpresso.Checked = False
            txtCodigoDeBarras.Enabled = False
            BtValidarCodBarras.Enabled = False
            CkbCodigoDeBarras.Enabled = False
            ckPreImpresso.Enabled = False
        End If

        If (cmbFormas.SelectedValue.ToString.Contains("3") Or
                cmbFormas.SelectedValue.ToString.Contains("6") Or
                cmbFormas.SelectedValue.ToString.Contains("7") Or
                cmbFormas.SelectedValue.ToString.Contains("11")) AndAlso Not divBanco.Visible Then
            divBanco.Visible = True
            ucConsultaDadosBancarios.Limpar()
            ucConsultaDadosBancarios.CarregaGrid(objNotaFiscal.CodigoCliente, objNotaFiscal.EnderecoCliente)
            Popup.ConsultaDeDadosBancarios(Me.Page, "objDadosBancariosNFG" & HID.Value)
        Else
            divBanco.Visible = False
        End If
    End Sub

    Protected Sub BtValidarCodBarras_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtValidarCodBarras.Click
        If ckPreImpresso.Checked = False Then
            If ddlFormasDePagamento.SelectedValue = 4 Then
                If Trim(txtCodigoDeBarras.Text) <> "" Then
                    'Deve passar a Empresa do Título para procurar em Datas não programáveis, não o Fornecedor - Furlan - 18/11/2013
                    'Dim strCliente As String() = hdfCodigoCliente.Value.Split(";")
                    Dim strEmpresa As String() = ddlEmpresa.SelectedValue.Split("-")
                    If CkbCodigoDeBarras.Checked Then txtCodigoDeBarras.Text = Funcoes.FormataLinhaDigitavelOriginal(txtCodigoDeBarras.Text)
                    If Funcoes.ValidaCodigoBarras(txtCodigoDeBarras.Text, CkbCodigoDeBarras.Checked, txtDataVencimento.Text, txtValorVencimento.Text, strEmpresa(0), strEmpresa(1), Banco) Then
                        MsgBox(Me.Page, "Codigo Barras Valido!")
                    Else
                        MsgBox(Me.Page, Funcoes.EliminarCaracteresEspeciais(RTrim(HttpContext.Current.Session("ssMessage").ToString)))
                    End If
                Else
                    MsgBox(Me.Page, "Código de Barras não foi informado")
                End If
            Else
                MsgBox(Me.Page, "Preenchimento do Codigo de Barras Somente Aceito para Boletos Bancarios")
            End If
        Else
            MsgBox(Me.Page, "Sistema não Valida Codigo De Barras de Boletos Pré Impressos...")
        End If
    End Sub

    Protected Sub ddlFormasDePagamento_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ddlFormasDePagamento.SelectedIndexChanged
        If ddlFormasDePagamento.SelectedValue = 4 Then
            txtCodigoDeBarras.Enabled = True
            BtValidarCodBarras.Enabled = True
            CkbCodigoDeBarras.Enabled = True
            ckPreImpresso.Enabled = True
        Else
            If (ddlFormasDePagamento.SelectedValue.ToString.Contains("3") Or
                ddlFormasDePagamento.SelectedValue.ToString.Contains("6") Or
                ddlFormasDePagamento.SelectedValue.ToString.Contains("7") Or
                ddlFormasDePagamento.SelectedValue.ToString.Contains("11")) AndAlso Not divBanco.Visible Then
                divBanco.Visible = True
                ucConsultaDadosBancarios.Limpar()
                SessaoRecuperaNotaFiscal()
                ucConsultaDadosBancarios.CarregaGrid(objNotaFiscal.CodigoCliente, objNotaFiscal.EnderecoCliente)
                Popup.ConsultaDeDadosBancarios(Me.Page, "objDadosBancariosNFG" & HID.Value)
            Else
                divBanco.Visible = False
            End If

            txtCodigoDeBarras.Text = ""
            CkbCodigoDeBarras.Checked = False
            ckPreImpresso.Checked = False
            txtCodigoDeBarras.Enabled = False
            BtValidarCodBarras.Enabled = False
            CkbCodigoDeBarras.Enabled = False
            ckPreImpresso.Enabled = False
        End If
    End Sub

    Protected Sub grdCondicoes_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        SessaoRecuperaNotaFiscal()
        divTitulo.Visible = True

        If objNotaFiscal.CodigoPedido > 0 Then ddlCondicaoDePagamento.SelectedValue = objNotaFiscal.Pedido.CodigoCondicaoPagamento

        cmbCarteira.SelectedValue = objNotaFiscal.VencimentosNota(grdCondicoes.SelectedIndex).CodigoCarteira

        If Not String.IsNullOrWhiteSpace(objNotaFiscal.VencimentosNota(grdCondicoes.SelectedIndex).Tributo) Then
            ddl.Carregar(ddlTributos, CarregarDDL.Tabela.CarteirasXTributos, "Carteira_Id = '" & cmbCarteira.SelectedValue & "'", True)
            ddlTributos.SelectedValue = objNotaFiscal.VencimentosNota(grdCondicoes.SelectedIndex).Tributo
        End If

        Dim Parametros As New Hashtable
        Parametros.Clear()
        Parametros.Add("listarTudo", "N")

        ddl.Carregar(ddlFormasDePagamento, CarregarDDL.Tabela.TipoDePagamento, "", True, Parametros)

        lblTitulo.Text = objNotaFiscal.VencimentosNota(grdCondicoes.SelectedIndex).Codigo
        txtDataVencimento.Text = objNotaFiscal.VencimentosNota(grdCondicoes.SelectedIndex).Prorrogacao
        txtValorVencimento.Text = String.Format("{0:N2}", objNotaFiscal.VencimentosNota(grdCondicoes.SelectedIndex).ValorDoDocumento)
        txtCodigoDeBarras.Text = objNotaFiscal.VencimentosNota(grdCondicoes.SelectedIndex).CodigoDeBarras
        CkbCodigoDeBarras.Checked = objNotaFiscal.VencimentosNota(grdCondicoes.SelectedIndex).CodigoDigitado
        ckPreImpresso.Checked = objNotaFiscal.VencimentosNota(grdCondicoes.SelectedIndex).CodigoDeBarrasPreImpresso

        If objNotaFiscal.VencimentosNota(grdCondicoes.SelectedIndex).CodigoTipoPgto = 0 Then
            objNotaFiscal.VencimentosNota(grdCondicoes.SelectedIndex).CodigoTipoPgto = 1
        End If

        ddlFormasDePagamento.SelectedValue = objNotaFiscal.VencimentosNota(grdCondicoes.SelectedIndex).CodigoTipoPgto
        cmbFormas.SelectedValue = objNotaFiscal.VencimentosNota(grdCondicoes.SelectedIndex).TipoPagto.TipoDePagamento

        ddlFormasDePagamento.Enabled = True
        ddlFormasDePagamento_SelectedIndexChanged(Nothing, Nothing)


        SessaoSalvaNotaFiscal()

        If objNotaFiscal.VencimentosNota(grdCondicoes.SelectedIndex).CodigoProvisao = 2 Then
            ddlFormasDePagamento.Enabled = True
            txtValorVencimento.Enabled = True
            cmdOkVencimento.Enabled = True
            txtDataVencimento.Enabled = True

            If ddlFormasDePagamento.SelectedValue = 4 Then
                txtCodigoDeBarras.Enabled = True
                BtValidarCodBarras.Enabled = True
                CkbCodigoDeBarras.Enabled = True
                ckPreImpresso.Enabled = True
            Else
                If (ddlFormasDePagamento.SelectedValue.ToString.Contains("3") Or
                    ddlFormasDePagamento.SelectedValue.ToString.Contains("6") Or
                    ddlFormasDePagamento.SelectedValue.ToString.Contains("7") Or
                    ddlFormasDePagamento.SelectedValue.ToString.Contains("11")) Then
                    divBanco.Visible = True
                Else
                    divBanco.Visible = False
                End If

                txtCodigoDeBarras.Text = ""
                CkbCodigoDeBarras.Checked = False
                ckPreImpresso.Checked = False
                CkbCodigoDeBarras.Checked = False
                ckPreImpresso.Checked = False
                txtCodigoDeBarras.Enabled = False
                BtValidarCodBarras.Enabled = False
                CkbCodigoDeBarras.Enabled = False
                ckPreImpresso.Enabled = False
            End If
        Else
            txtCodigoDeBarras.Enabled = False
            BtValidarCodBarras.Enabled = False
            CkbCodigoDeBarras.Enabled = False
            ckPreImpresso.Enabled = False
            ddlFormasDePagamento.Enabled = False
            txtValorVencimento.Enabled = False
            cmdOkVencimento.Enabled = False
            txtDataVencimento.Enabled = False
        End If
    End Sub

    Protected Sub cmdOkVencimento_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdOkVencimento.Click
        SessaoRecuperaNotaFiscal()
        Dim strEmpresa As String() = ddlEmpresa.SelectedValue.Split("-")

        If Not IsDate(txtDataVencimento.Text) OrElse Convert.ToDateTime(txtDataVencimento.Text) < Convert.ToDateTime(txtDataNota.Text) Then
            MsgBox(Me.Page, "Data de vencimento não pode ser menor que " & txtDataNota.Text)
            Exit Sub
        End If

        If ddlFormasDePagamento.SelectedValue = 4 AndAlso Not ckPreImpresso.Checked AndAlso Not Funcoes.ValidaCodigoBarras(txtCodigoDeBarras.Text, CkbCodigoDeBarras.Checked, txtDataVencimento.Text, txtValorVencimento.Text, strEmpresa(0), strEmpresa(1), Banco) Then
            MsgBox(Me.Page, "Código de Barras Inválido!")
            Exit Sub
        End If

        Dim DataUtil = Funcoes.ValidaDataUtil(objNotaFiscal.CodigoEmpresa, objNotaFiscal.EnderecoEmpresa, CDate(txtDataVencimento.Text))

        If objNotaFiscal.VencimentosNota.Count = 1 Then
            If objNotaFiscal.IUD = "I" Then objNotaFiscal.VencimentosNota(grdCondicoes.SelectedIndex).Vencimento = Funcoes.ValidaDataUtil(objNotaFiscal.CodigoEmpresa, objNotaFiscal.EnderecoEmpresa, CDate(txtDataVencimento.Text))
            objNotaFiscal.VencimentosNota(grdCondicoes.SelectedIndex).Prorrogacao = DataUtil
            objNotaFiscal.VencimentosNota(grdCondicoes.SelectedIndex).Baixa = CDate(txtMovimento.Text)
            objNotaFiscal.VencimentosNota(grdCondicoes.SelectedIndex).CodigoTipoPgto = ddlFormasDePagamento.SelectedValue
            objNotaFiscal.VencimentosNota(grdCondicoes.SelectedIndex).ValorDoDocumento = txtValorVencimento.Text
            objNotaFiscal.VencimentosNota(grdCondicoes.SelectedIndex).CodigoDigitado = CkbCodigoDeBarras.Checked
            objNotaFiscal.VencimentosNota(grdCondicoes.SelectedIndex).CodigoDeBarrasPreImpresso = ckPreImpresso.Checked
            objNotaFiscal.VencimentosNota(grdCondicoes.SelectedIndex).CodigoDeBarras = txtCodigoDeBarras.Text.Replace(" ", "").Replace(".", "")

            txtTotalParcelado.Text = txtValorVencimento.Text
            SessaoSalvaNotaFiscal()
            grdCondicoes.DataSource = objNotaFiscal.VencimentosNota
            grdCondicoes.DataBind()
            ddlFormasDePagamento.Enabled = True
            txtValorVencimento.Enabled = True
            LimparVencimentos()
            Exit Sub
        End If

        Dim msg As String = objNotaFiscal.VencimentosNota.AjustaParcelas(grdCondicoes.SelectedIndex, objNotaFiscal.VencimentosNota(grdCondicoes.SelectedIndex).ValorDoDocumento, txtValorVencimento.Text.Trim())
        If String.IsNullOrWhiteSpace(msg) Then
            objNotaFiscal.VencimentosNota(grdCondicoes.SelectedIndex).Codigo = lblTitulo.Text.Trim()
            If objNotaFiscal.IUD = "I" Then objNotaFiscal.VencimentosNota(grdCondicoes.SelectedIndex).Vencimento = Funcoes.ValidaDataUtil(objNotaFiscal.CodigoEmpresa, objNotaFiscal.EnderecoEmpresa, CDate(txtDataVencimento.Text))
            objNotaFiscal.VencimentosNota(grdCondicoes.SelectedIndex).Prorrogacao = DataUtil
            objNotaFiscal.VencimentosNota(grdCondicoes.SelectedIndex).Baixa = CDate(txtMovimento.Text)
            objNotaFiscal.VencimentosNota(grdCondicoes.SelectedIndex).ValorDoDocumento = txtValorVencimento.Text.Trim()
            objNotaFiscal.VencimentosNota(grdCondicoes.SelectedIndex).CodigoDigitado = CkbCodigoDeBarras.Checked
            objNotaFiscal.VencimentosNota(grdCondicoes.SelectedIndex).CodigoDeBarrasPreImpresso = ckPreImpresso.Checked
            objNotaFiscal.VencimentosNota(grdCondicoes.SelectedIndex).CodigoDeBarras = txtCodigoDeBarras.Text.Replace(" ", "").Replace(".", "")
            objNotaFiscal.VencimentosNota(grdCondicoes.SelectedIndex).CodigoTipoPgto = ddlFormasDePagamento.SelectedValue
            grdCondicoes.DataSource = objNotaFiscal.VencimentosNota
            grdCondicoes.DataBind()
            SessaoSalvaNotaFiscal()
        Else
            MsgBox(Me.Page, msg)
        End If
        LimparVencimentos()
    End Sub

    Private Sub LimparVencimentos()
        lblTitulo.Text = ""
        txtDataVencimento.Text = ""
        CkbCodigoDeBarras.Checked = False
        ckPreImpresso.Checked = False
        txtCodigoDeBarras.Enabled = False
        BtValidarCodBarras.Enabled = False
        CkbCodigoDeBarras.Enabled = False
        ckPreImpresso.Enabled = False
        txtCodigoDeBarras.Text = ""
        ddlFormasDePagamento.SelectedValue = ""
        grdCondicoes.SelectedIndex = -1
        divTitulo.Visible = False
    End Sub

    Protected Sub btnDadosBancarios_Click(sender As Object, e As EventArgs) Handles btnDadosBancarios.Click
        SessaoRecuperaNotaFiscal()
        ucConsultaDadosBancarios.Limpar()
        ucConsultaDadosBancarios.CarregaGrid(objNotaFiscal.CodigoCliente, objNotaFiscal.EnderecoCliente)
        Popup.ConsultaDeDadosBancarios(Me.Page, "objDadosBancariosNFG" & HID.Value)
    End Sub

    Protected Sub lnkConsultarNaviosXInvoice_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lnkConsultarNaviosXInvoice.Click
        Try
            Popup.ConsultarNaviosXInvoice(Me.Page, "objConsultarNaviosXInvoice")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub
#End Region

#Region "Methods & Functions"

    Private Sub InformarDadosProdutos(ByVal xmlDoc As XmlDocument)

        SessaoSalvaNotaFiscal()
        Popup.InformarDadosProdutoNFG(Me.Page, "objProdutoNotaNF" & HID.Value)

    End Sub

    Private Sub AdicionarProdutos(ByVal bAbrirPopUp As Boolean, ByVal bXMLAutomatico As Boolean)

        If String.IsNullOrWhiteSpace(hdfCodigoCliente.Value) Then
            MsgBox(Me.Page, "É necessário selecionar um cliente!")
            Exit Sub
        ElseIf String.IsNullOrWhiteSpace(ddlTipoDeDocumento.SelectedValue) Then
            MsgBox(Me.Page, "É necessário selecionar o tipo de documento!")
            Exit Sub
        ElseIf (ddlTipoDeDocumento.SelectedValue = 2 OrElse ddlTipoDeDocumento.SelectedValue = 10 OrElse ddlTipoDeDocumento.SelectedValue = 14 OrElse ddlTipoDeDocumento.SelectedValue = 57) AndAlso grdNotasFretes.Rows.Count = 0 Then
            MsgBox(Me.Page, "Para lançamento de Conhecimento o Vinculo da Nota Fiscal deve ser selecionado antes.")
            ucNFOrigem.Limpar()
            Popup.ConsultaNotaOrigem(Me.Page, "objOrigemNFG" & HID.Value)
        ElseIf String.IsNullOrWhiteSpace(txtNumeroNota.Text) Then
            MsgBox(Me.Page, "É necessário informar o número da nota!")
            Exit Sub
        ElseIf String.IsNullOrWhiteSpace(txtSerie.Text) Then
            MsgBox(Me.Page, "É necessário informar a série da nota!")
            Exit Sub
        ElseIf String.IsNullOrWhiteSpace(txtMovimento.Text) Then
            MsgBox(Me.Page, "É necessário informar a Data do Movimento!")
            Exit Sub
        ElseIf String.IsNullOrWhiteSpace(txtDataNota.Text) Then
            MsgBox(Me.Page, "É necessário informar a Data da Nota!")
            Exit Sub
        Else

            'SessaoRecuperaNotaFiscal()
            If bAbrirPopUp Then
                Popup.ConsultaProdutoNFG(Me.Page, "objProdutoNotaNF" & HID.Value)
            End If

            ucProdutoNFG.InicializarUC(-1, bXMLAutomatico, chkReaproveitar.Checked)

            If bAbrirPopUp = False Then
                ucProdutoNFG.AdicionarProduto(bXMLAutomatico, chkReaproveitar.Checked)
            End If

        End If

    End Sub

    Private Sub carregarNotaXMLautomatico(ByVal sender As Object, ByVal e As System.EventArgs)
        If Not chaveXMLautomatico Is Nothing Then
            txtChaveNFe.Text = chaveXMLautomatico
            lnkVerificarChaveNFE_Click(sender, e)
        End If
    End Sub

    Private Function ValidarCampos() As Boolean

        'Verificar se foi feito a recusa da Nota Fiscal
        If objNotaFiscal.TemRecusa Then
            mensagemErro = "Nota Fiscal não pode ser lançada pois a mesma foi lançada como recusada."
            Return False
        End If

        'Dim i As Integer = objNotaFiscal.Titulos.AdiantamentosAbertos.Where(Function(s) s.VlrBaixa > 0 AndAlso s.VlrBaixa > s.Valor).Count

        Dim valida As Boolean = True

        If (Mid(objNotaFiscal.ChaveNFE, 7, 14) = "02935843000105" Or
            Mid(objNotaFiscal.ChaveNFE, 7, 14) = "01409655000180" Or
            Mid(objNotaFiscal.ChaveNFE, 7, 14) = "78393592000146" Or
            Mid(objNotaFiscal.ChaveNFE, 7, 14) = "82951310000156" Or
            Mid(objNotaFiscal.ChaveNFE, 7, 14) = "87958674000181" Or
            Mid(objNotaFiscal.ChaveNFE, 7, 14) = "58290502000184" Or
            Mid(objNotaFiscal.ChaveNFE, 7, 14) = "76416890000189" Or
            Mid(objNotaFiscal.ChaveNFE, 7, 14) = "03507415000578") Then
            valida = False
        End If

        If chkNaturezaDeRendimento.Checked AndAlso ddlNaturezaDeRendimento.SelectedIndex = 0 Then
            mensagemErro = "Natureza de Rendimento não foi selecionada."
            Return False
        End If

        'Repete a Validacao da chave da nota caso seja eletronica S e nao seja nossa emissao
        If objNotaFiscal.Eletronica Then
            If valida AndAlso objNotaFiscal.CodigoCliente.Length = 11 AndAlso Not Mid(objNotaFiscal.ChaveNFE, 10, 11) = objNotaFiscal.CodigoCliente Then
                mensagemErro = "Cliente na Chave Eletrônica diferente do informado na Nota Fiscal."
                Return False
            ElseIf valida AndAlso objNotaFiscal.CodigoCliente.Length = 14 AndAlso Not Mid(objNotaFiscal.ChaveNFE, 7, 14) = objNotaFiscal.CodigoCliente Then
                MsgBox(Me.Page, "Cliente na Chave Eletrônica diferente do informado na Nota Fiscal.")
                Return False
            ElseIf objNotaFiscal.Cliente.Municipio.EstadoIbge = 0 Then
                MsgBox(Me.Page, "o Código IBGE do cliente não foi informado no cadastro! Favor revisar o cadastro do cliente.")
                Return False
            ElseIf valida AndAlso Not CInt(Left(objNotaFiscal.ChaveNFE, 2)) = objNotaFiscal.Cliente.Municipio.EstadoIbge Then
                mensagemErro = "Estado do Cliente na Chave Eletrônica diferente do informado na Nota Fiscal."
                Return False
            ElseIf Not CInt(Mid(objNotaFiscal.ChaveNFE, 26, 9)) = CInt(txtNumeroNota.Text) Then
                mensagemErro = "Número da Nota na Chave Eletrônica diferente do informado na Nota Fiscal."
                Return False
            ElseIf Not CInt(Mid(objNotaFiscal.ChaveNFE, 23, 3)) = CInt(txtSerie.Text) Then
                mensagemErro = "Série da Nota na Chave Eletrônica diferente do informado na Nota Fiscal."
                Return False
            ElseIf Not Mid(objNotaFiscal.ChaveNFE, 3, 2) = String.Format("{0:yy}", CDate(txtDataNota.Text)) Then
                mensagemErro = "Ano da Nota na Chave Eletrônica diferente do informado na Nota Fiscal."
                Return False
            ElseIf Not Mid(objNotaFiscal.ChaveNFE, 5, 2) = String.Format("{0:MM}", CDate(txtDataNota.Text)) Then
                mensagemErro = "Mês da Nota na Chave Eletrônica diferente do informado na Nota Fiscal."
                Return False
            End If
        End If

        If ddlTipoDeDocumento.SelectedValue = 1 AndAlso objNotaFiscal.Empresa.Empresa.ObrigaChaveNf AndAlso Not objNotaFiscal.ChaveNFE.Length = 44 Then
            mensagemErro = "Para Tipo do Documento Nota Fiscal Eletrônica é Obrigatório a Chave da NFe."
            Return False
        End If

        If (ddlTipoDeDocumento.SelectedValue = 2 Or ddlTipoDeDocumento.SelectedValue = 57 Or ddlTipoDeDocumento.SelectedValue = 58) AndAlso objNotaFiscal.Empresa.Empresa.ObrigaChaveNfg AndAlso Not objNotaFiscal.ChaveNFE.Length = 44 Then
            mensagemErro = "Para Tipo do Documento Conhecimento de Transporte Eletrônico é Obrigatório a Chave do CTe."
            Return False
        End If

        'Verificar Valor Documento ZERADO
        If objNotaFiscal.TotalNota = 0 Then
            Dim liberaTotalZero As Boolean = False

            For Each item In objNotaFiscal.Itens.Where(Function(s) Not s.IUD = "D")
                For Each enc In item.Encargos
                    If enc.Codigo = "ICMS" AndAlso enc.Valor > 0 Then
                        'Apenas complemento de ICMS
                        liberaTotalZero = True
                    End If
                Next
            Next

            If Not liberaTotalZero Then
                mensagemErro = "Total da Nota Fiscal não pode ser Zero"
                Return False
            End If
        End If

        If String.IsNullOrWhiteSpace(objNotaFiscal.CodigoEmpresa) Then
            mensagemErro = "Selecione uma Empresa para Continuar"
            Return False
        ElseIf String.IsNullOrWhiteSpace(objNotaFiscal.CodigoCliente) Then
            mensagemErro = "Selecione um Cliente para Continuar"
            Return False
        ElseIf Not Funcoes.VerificaAcesso(objNotaFiscal.CodigoEmpresa, objNotaFiscal.EnderecoEmpresa, objNotaFiscal.Movimento, "NOTAS FISCAIS") Then
            mensagemErro = "Movimento da Nota Fiscal já fechado para esta data"
            Return False
        ElseIf Not Funcoes.VerificaAcesso(objNotaFiscal.CodigoEmpresa, objNotaFiscal.EnderecoEmpresa, objNotaFiscal.Movimento, "CONTABIL") Then
            mensagemErro = "Movimento Contábil já fechado para esta data"
            Return False
        ElseIf objNotaFiscal.CodigoSituacao = 0 Then
            mensagemErro = "Situação não foi selecionada"
            Return False
        ElseIf objNotaFiscal.IUD = "I" AndAlso objNotaFiscal.CodigoSituacao <> 1 Then
            mensagemErro = "Situação na gravação deve ser Normal na Inclusão"
            Return False
        ElseIf objNotaFiscal.CodigoTipoDeDocumento = 0 Then
            mensagemErro = "Tipo de Documento não foi selecionado"
            Return False
        ElseIf objNotaFiscal.Codigo = 0 Then
            mensagemErro = "Número da Nota Fiscal não foi informado"
            Return False
        ElseIf String.IsNullOrWhiteSpace(objNotaFiscal.Serie) Then
            mensagemErro = "Série da Nota Fiscal não foi informada"
            Return False
        ElseIf Trim(objNotaFiscal.Serie) = "00" OrElse Trim(objNotaFiscal.Serie) = "000" Then
            mensagemErro = "Série invalida"
            Return False
        ElseIf objNotaFiscal.Itens.Count = 0 Then
            mensagemErro = "Não foi adicionado nenhum Produto na Nota Fiscal"
            Return False
        ElseIf String.IsNullOrWhiteSpace(objNotaFiscal.CodigoSafra) Then
            mensagemErro = "Safra não foi selecionada"
            Return False
        ElseIf objNotaFiscal.Movimento < objNotaFiscal.DataNota Then
            mensagemErro = "A Data de Emissão não pode ser Posterior a Data do Movimento."
            Return False
        ElseIf (objNotaFiscal.CodigoTipoDeDocumento = 2 OrElse objNotaFiscal.CodigoTipoDeDocumento = 10 OrElse objNotaFiscal.CodigoTipoDeDocumento = 14 OrElse objNotaFiscal.CodigoTipoDeDocumento = 57) AndAlso (objNotaFiscal.NotasTrocaOrigem Is Nothing OrElse objNotaFiscal.NotasTrocaOrigem.Count = 0) Then
            mensagemErro = "Vinculo do Conhecimento não foi selecionado."
            Return False
        ElseIf objNotaFiscal.CodigoTipoDeDocumento = eTipoDeDocumento.RPA AndAlso (objNotaFiscal.NotasReferenciais Is Nothing OrElse objNotaFiscal.NotasReferenciais.Where(Function(s) s.IUD <> "D").Count = 0) Then
            mensagemErro = "Vincule as Notas Fiscais ao RPA."
            Return False
        ElseIf Not String.IsNullOrWhiteSpace(objNotaFiscal.ChaveNFE) AndAlso objNotaFiscal.ChaveNFE.Replace(".", "").Length < 44 Then
            mensagemErro = "Verifique chave Nota Eletrônica, deve ter 44 dígitos"
            Return False
        ElseIf objNotaFiscal.SubOperacao Is Nothing Then
            mensagemErro = "Operação da Nota Fiscal não foi selecionada"
            Return False
        ElseIf objNotaFiscal.ChaveNFE.Length > 0 AndAlso objNotaFiscal.ChaveNFE.Length < 44 Then
            mensagemErro = "Verifique chave Nota Eletrônica, deve ter 44 dígitos."
            Return False
        ElseIf objNotaFiscal.Itens.Any(Function(s) s.IUD <> "D" AndAlso s.CFOP = 0) Then
            mensagemErro = "Verifique o lançamento, existe cfop sem informação!"
            Return False
        ElseIf ddlTipoDeDocumento.SelectedValue = 1 AndAlso (objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(0).CFOP = 1352 OrElse
                                                             objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(0).CFOP = 2352 OrElse
                                                             objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(0).CFOP = 5352 OrElse
                                                             objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(0).CFOP = 6352 OrElse
                                                             objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(0).CFOP = 1353 OrElse
                                                             objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(0).CFOP = 2353 OrElse
                                                             objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(0).CFOP = 5353 OrElse
                                                             objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(0).CFOP = 6353) Then
            mensagemErro = "Verifique esse lançamento, cfop de transporte não pode ser usado para esse tipo de documento!"
            Return False
        ElseIf (ddlTipoDeDocumento.SelectedValue = 2 Or
            ddlTipoDeDocumento.SelectedValue = 3 Or
            ddlTipoDeDocumento.SelectedValue = 4 Or
            ddlTipoDeDocumento.SelectedValue = 8 Or
            ddlTipoDeDocumento.SelectedValue = 10 Or
            ddlTipoDeDocumento.SelectedValue = 11 Or
            ddlTipoDeDocumento.SelectedValue = 14 Or
            ddlTipoDeDocumento.SelectedValue = 57 Or
            ddlTipoDeDocumento.SelectedValue = 58 Or
            ddlTipoDeDocumento.SelectedValue = 67) AndAlso
            Not objNotaFiscal.Cliente.Tipos.Any(Function(s) s.CodigoTipo = eTipoCliente.Transportadores) Then
            mensagemErro = "É necessário selecionar um Tipo de Cliente Transportador!"
            Return False
        ElseIf objNotaFiscal.SubOperacao.Financeiro AndAlso objNotaFiscal.TotalNota > 0 Then
            If Not FinanceiroNovo Then
                If objNotaFiscal.VencimentosNota Is Nothing OrElse objNotaFiscal.VencimentosNota.Count = 0 Then
                    mensagemErro = "É necessário programar o financeiro para esta nota!"
                    Return False
                ElseIf objNotaFiscal.IUD = "I" AndAlso cmbCarteira.SelectedIndex = 0 Then
                    mensagemErro = "Carteira não foi selecionada"
                    Return False
                ElseIf objNotaFiscal.IUD = "I" AndAlso cmbCarteira.SelectedIndex > 0 AndAlso ddlTributos.Items.Count > 1 AndAlso ddlTributos.SelectedIndex = 0 Then
                    mensagemErro = "Encargo não foi selecionado"
                    Return False
                ElseIf objNotaFiscal.TotalNota <> objNotaFiscal.VencimentosNota.Where(Function(s) Not s.IUD = "D").Sum(Function(s) s.ValorDoDocumento) Then
                    mensagemErro = "Valor Liquido da Nota Fiscal não pode ser diferente do Valor Programando no Financeiro."
                    Return False
                ElseIf objNotaFiscal.VencimentosNota.All(Function(s) s.CodigoTipoPgto = 3) OrElse
                       objNotaFiscal.VencimentosNota.All(Function(s) s.CodigoTipoPgto = 6) OrElse
                       objNotaFiscal.VencimentosNota.All(Function(s) s.CodigoTipoPgto = 7) OrElse
                       objNotaFiscal.VencimentosNota.All(Function(s) s.CodigoTipoPgto = 11) Then

                    If Not objNotaFiscal.VencimentosNota.All(Function(s) s.CodigoBancoCliente > 0 AndAlso Len(s.CodigoAgenciaCliente) > 0) Then
                        mensagemErro = "Para forma de pagamento 3,6,7 e 11 deve ser informado os Dados Bancários do Cliente."
                        Return False
                    End If
                End If

                Dim dataParcela = Today()
                Dim primeiro = True

                For Each tit In objNotaFiscal.VencimentosNota

                    If tit.Agrupado = eAgrupamentoFinanceiro.Agrupado Then
                        mensagemErro = "Título em Agrupamento não pode ser alterado."
                        Return False
                    ElseIf tit.CodigoTipoPgto = 4 AndAlso Not tit.CodigoDeBarrasPreImpresso AndAlso Not tit.CodigoProvisao = 1 Then
                        If tit.CodigoDeBarras.Length = 0 Then
                            mensagemErro = "Preenchimento do Codigo de Barras é Obrigatório para Boletos Bancarios."
                            Return False
                        Else
                            If tit.CodigoDigitado Then tit.CodigoDeBarras = Funcoes.FormataLinhaDigitavelOriginal(tit.CodigoDeBarras)
                            If Not Funcoes.ValidaCodigoBarras(tit.CodigoDeBarras, tit.CodigoDigitado, tit.Prorrogacao, tit.ValorDoDocumento, tit.CodigoEmpresa, tit.EnderecoEmpresa, Banco) Then
                                mensagemErro = Funcoes.EliminarCaracteresEspeciais(RTrim(HttpContext.Current.Session("ssMessage").ToString))
                                Return False
                            End If
                        End If
                    End If

                    If primeiro AndAlso objNotaFiscal.VencimentosNota.Count > 1 AndAlso tit.Prorrogacao >= dataParcela Then
                        dataParcela = tit.Prorrogacao
                        primeiro = False
                    End If

                    If objNotaFiscal.Empresa.Empresa.ControlaDataMovimentoNFG AndAlso tit.Prorrogacao < dataParcela Then
                        mensagemErro = "Verifique o(s) vencimento(s) no financeiro programado pois a ordem não está correspondente ou a data da parcela maior que a data do primeiro titulo!"
                        Return False
                    ElseIf tit.ValorDoDocumento = 0 Then
                        mensagemErro = "Verifique o(s) vencimento(s), valor programado no título " & tit.Codigo & " não pode ser 0(ZERO)."
                        Return False
                    End If

                    dataParcela = tit.Prorrogacao

                Next

                'Valida o conta contábil do título com a carteira(Baixa)
                For Each titulo In objNotaFiscal.VencimentosNota.Where(Function(s) s.CodigoProvisao = eProvisao.Baixa)
                    If Not objNotaFiscal.SubOperacao.CodigoGrupoContas = titulo.Carteira.CodigoContaCliente Then
                        mensagemErro = "Conta Contábil do produto diferente da conta contábil do(s) título(s) já baixados!"
                        Return False
                    ElseIf Convert.ToDateTime(txtMovimento.Text) > titulo.Baixa Then
                        mensagemErro = "Data de Movimento superior a data de baixa do titulo: " & titulo.Codigo
                        Return False
                    End If
                Next

                If Not objNotaFiscal.VencimentosNota(0).Carteira.CodigoContaCliente = objNotaFiscal.Itens.Where(Function(s) Not s.IUD = "D")(0).SubOperacao.CodigoGrupoContas Then
                    mensagemErro = "Conta de Cliente " & objNotaFiscal.VencimentosNota(0).Carteira.CodigoContaCliente & " da Carteira selecionada é diferente da Conta de Cliente da Operação da Nota " & objNotaFiscal.Itens.Where(Function(s) Not s.IUD = "D")(0).SubOperacao.CodigoGrupoContas
                    Return False
                End If
            Else
                If (objNotaFiscal.Titulos Is Nothing OrElse objNotaFiscal.Titulos.Count = 0) AndAlso (Not objNotaFiscal.Titulos.AdiantamentosAbertos.Count > 0) Then
                    mensagemErro = "É necessário programar o financeiro para esta nota!"
                    Return False
                ElseIf cmbFormas.SelectedValue.ToString.Contains("3") _
                OrElse cmbFormas.SelectedValue.ToString.Contains("6") _
                OrElse cmbFormas.SelectedValue.ToString.Contains("7") _
                OrElse cmbFormas.SelectedValue.ToString.Contains("11") Then
                    If Not objNotaFiscal.Titulos.All(Function(s) s.CodigoBancoCliFor > 0 AndAlso Len(s.CodigoAgenciaCliFor) > 0) Then
                        mensagemErro = "Para forma de pagamento 3,6,7 e 11 deve ser informado os Dados Bancários do Cliente."
                        Return False
                    End If
                ElseIf objNotaFiscal.TotalNota <> IIf(objNotaFiscal.Titulos.AdiantamentosAbertos.Count > 0, objNotaFiscal.Titulos.Where(Function(s) Not s.IUD = "D").Sum(Function(s) s.Valores.EncargoValorDocumento.ValorOficial) + objNotaFiscal.Titulos.AdiantamentosAbertos.ValorTotalInformadoParaBaixa, objNotaFiscal.Titulos.Where(Function(s) Not s.IUD = "D").Sum(Function(s) s.Valores.EncargoValorDocumento.ValorOficial)) Then
                    mensagemErro = "Valor Liquido da Nota Fiscal não pode ser diferente do Valor Programado no Financeiro."
                    Return False
                ElseIf objNotaFiscal.Titulos.AdiantamentosAbertos.Where(Function(s) s.VlrBaixa > 0 AndAlso s.VlrBaixa > s.Valor).Count > 0 Then
                    mensagemErro = "Valor Informado para baixa de adiantamentos é maior que o valor da Nota Fiscal"
                    Return False
                End If

                For Each titulo In objNotaFiscal.Titulos.Where(Function(s) s.CodigoProvisao = eProvisao.Baixa)
                    If Not objNotaFiscal.SubOperacao.CodigoGrupoContas = titulo.CodigoContaContabilCliFor Then
                        mensagemErro = "Conta Contábil do produto diferente da conta contábil do(s) título(s) já baixados!"
                        Return False
                    ElseIf Convert.ToDateTime(txtMovimento.Text) > titulo.DataBaixa Then
                        mensagemErro = "Data de Movimento superior a data de baixa do titulo: " & titulo.Codigo
                        Return False
                    End If
                Next

                For Each tit In objNotaFiscal.Titulos
                    If tit.CodigoTipoPgto = 4 AndAlso Not tit.CodigoDeBarrasPreImpresso AndAlso Not String.IsNullOrWhiteSpace(tit.CodigoDeBarras) Then
                        If tit.CodigoDeBarrasDigitado Then tit.CodigoDeBarras = Funcoes.FormataLinhaDigitavelOriginal(tit.CodigoDeBarras)
                        If Not Funcoes.ValidaCodigoBarras(tit.CodigoDeBarras, tit.CodigoDeBarrasDigitado, tit.Reprogramacao, tit.Valores.EncargoValorDocumento.Valor, tit.CodigoEmpresa, tit.EnderecoEmpresa, Banco) Then
                            mensagemErro = Funcoes.EliminarCaracteresEspeciais(RTrim(HttpContext.Current.Session("ssMessage").ToString))
                            Return False
                        End If
                    End If
                Next

            End If
        End If

        'Para operações q não tem financeiro e possui vencimentos.(Na alteração de operação na nota)
        If Not FinanceiroNovo Then
            If Not objNotaFiscal.SubOperacao.Financeiro AndAlso objNotaFiscal.VencimentosNota.Count > 0 AndAlso objNotaFiscal.VencimentosNota.All(Function(s) s.CodigoProvisao = 1) Then
                MsgBox(Me.Page, "Atenção! operação:" & objNotaFiscal.CodigoOperacao & "-" & objNotaFiscal.CodigoSubOperacao & " não gera financeiro, porém a NF já possuí títulos baixados")
                Return False
            End If
        Else
            If Not objNotaFiscal.SubOperacao.Financeiro AndAlso objNotaFiscal.VencimentosNota.Count > 0 AndAlso objNotaFiscal.Titulos.All(Function(s) s.CodigoProvisao = 1) Then
                MsgBox(Me.Page, "Atenção! operação:" & objNotaFiscal.CodigoOperacao & "-" & objNotaFiscal.CodigoSubOperacao & " não gera financeiro, porém a NF já possuí títulos baixados")
                Return False
            End If
        End If

        'furlan - Não deixar gravar encargo caso não tenha conta débito ou crédito - 18/07/2014
        '*********************************************************************************************************************************************************
        If objNotaFiscal.SubOperacao.Contabil Then
            For Each item In objNotaFiscal.Itens
                If String.IsNullOrWhiteSpace(item.SubOperacao.CodigoGrupoContas) Then
                    If item.Encargos.EncLiquido.ContaDeCredito.Length = 0 And item.Encargos.EncLiquido.ContaDeDebito.Length = 0 Then
                        mensagemErro = "Operação " & objNotaFiscal.CodigoOperacao & "-" & objNotaFiscal.CodigoSubOperacao & " não possui conta contábil, verifique."
                        Return False
                    End If
                End If

                For Each enc In item.Encargos
                    If enc.Valor > 0 Then

                        'Ignora LIQUIDO
                        If enc.Codigo = "LIQUIDO" Then
                            Continue For
                        End If

                        'IBS/CBS neste momento não valida conta contábil para empresas do grupo Verde
                        If (enc.Codigo = "IBS" OrElse enc.Codigo = "CBS") AndAlso objNotaFiscal.NossaEmissao = False Then
                            Continue For
                        End If

                        If String.IsNullOrWhiteSpace(enc.OperacaoEncargo.CodigoDebitaConta) And String.IsNullOrWhiteSpace(enc.OperacaoEncargo.CodigoCreditaConta) Then
                            mensagemErro = "Produto " & item.CodigoProduto & " ENCARGO " & enc.Codigo & " não possui conta DÉBITO/CRÉDITO, verifique."
                            Return False
                        End If

                    End If
                Next
            Next
        End If

        'ATÉ AQUI
        If Left(objNotaFiscal.CodigoEmpresa, 8) = "04854422" Then
            If ddlTipoDeDocumento.SelectedValue = 2 Or
                ddlTipoDeDocumento.SelectedValue = 3 Or
                ddlTipoDeDocumento.SelectedValue = 4 Or
                ddlTipoDeDocumento.SelectedValue = 10 Or
                ddlTipoDeDocumento.SelectedValue = 11 Or
                ddlTipoDeDocumento.SelectedValue = 12 Then
                mensagemErro = "Tipo de documento só pode ser lançado pelo novo sistema de fretes!"
                Return False
            End If
        End If

        'Arquivo
        If objNotaFiscal.Empresa.Empresa.ArquivoNFE AndAlso Not objNotaFiscal.Arquivos.Count > 0 Then
            MsgBox(Me.Page, "O campo Arquivo é de preenchimento obrigatório!")
            Return False
        End If
        'Conferencia Fiscal de Documento
        If chkConferencia.Visible AndAlso objNotaFiscal.Empresa.Empresa.ConferenciaNFE AndAlso
            Not chkConferencia.Checked AndAlso String.IsNullOrWhiteSpace(objNotaFiscal.UsuarioConferencia) Then
            MsgBox(Me.Page, "O campo Conferência é de preenchimento obrigatório!")
            Return False
        End If

        For ni As Integer = 0 To objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D").Count - 1
            If Not objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(ni).Encargos.EncProduto.GrupoDeContaDebito Is Nothing AndAlso
                objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(ni).Encargos.EncProduto.GrupoDeContaDebito.TemCentroDeCusto AndAlso
                objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(ni).Encargos.EncProduto.CentroDeCusto = "0" AndAlso
                objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(ni).Rateios.Where(Function(x) x.IUD <> "D").Count = 0 Then
                MsgBox(Me.Page, "Centro de Custo deve ser informado no Item ou no Rateio do Produto " & objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(ni).CodigoProduto & "-" & objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(ni).Produto.Nome)
                Return False
            End If

            If Not objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(ni).Encargos.EncProduto.GrupoDeContaDebito Is Nothing AndAlso
                objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(ni).Encargos.EncProduto.GrupoDeContaDebito.TemCentroDeCusto AndAlso
                objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(ni).Encargos.EncProduto.CentroDeCusto = "0" AndAlso
                objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(ni).ValorTotal <> objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(ni).Rateios.Where(Function(x) x.IUD <> "D").Sum(Function(z) z.Valor) Then
                MsgBox(Me.Page, "Valor total do rateio " & objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(ni).Rateios.Where(Function(x) x.IUD <> "D").Sum(Function(z) z.Valor) &
                      " está diferente do valor da NF: " & objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(ni).ValorTotal &
                      " no produto: " & objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(ni).CodigoProduto & "-" & objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(ni).Produto.Nome)
                Return False
            End If
        Next
        Return True
    End Function

    Public Sub CarregarFormComAClasse(ByVal objNota As NotaFiscal, Optional ByVal pXmlNFe As Boolean = False)

        Dim dsXML As DataSet = SessaoDsXML

        ddlUnidadeNegocio.SelectedValue = objNota.CodigoUnidadeDeNegocio
        ddl.Carregar(ddlEmpresa, CarregarDDL.Tabela.Empresas, ddlUnidadeNegocio.SelectedValue, True)
        ddlEmpresa.SelectedValue = objNota.CodigoEmpresa & "-" & objNota.EnderecoEmpresa
        hdfCodigoCliente.Value = objNota.CodigoCliente & "-" & objNota.EnderecoCliente
        If objNota.Cliente IsNot Nothing Then Funcoes.FormatarClienteTXT(txtNomeCliente, objNota.Cliente)

        txtChaveNFe.Text = Funcoes.FormatarChaveNFe(objNota.ChaveNFE)

        ddlSafra.SelectedValue = objNota.CodigoSafra
        txtObservacao.Text = objNota.Observacoes
        ddlSituacao.SelectedValue = objNota.CodigoSituacao
        ddlSituacao.Enabled = False
        ddlTipoDeDocumento.SelectedValue = objNota.CodigoTipoDeDocumento

        lnkAdicionarItem.Parent.Parent.Visible = True
        grdProdutos.Columns(0).Visible = True
        grdProdutos.Columns(13).Visible = True

        If ddlTipoDeDocumento.SelectedValue = 8 Then
            chkProvisao.Parent.Visible = True
        Else
            chkProvisao.Parent.Visible = False
            chkProvisao.Checked = False
        End If

        lblPedido.Text = objNota.CodigoPedido
        txtNumeroNota.Text = objNota.Codigo
        txtSerie.Text = objNota.Serie
        txtMovimento.Text = objNota.Movimento.ToShortDateString()
        txtDataNota.Text = objNota.DataNota.ToShortDateString()

        If Not objNotaFiscal Is Nothing Then
            If objNotaFiscal.Itens.Count > 0 Then
                grdProdutos.DataSource = Nothing
                grdProdutos.DataSource = objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D").ToList.OrderBy(Function(s) s.Sequencia)
                grdProdutos.DataBind()

                If objNotaFiscal.Itens(0).Encargos.Count > 0 Then
                    AtualizarItensNoGrid()

                    If dsXML.Tables.Contains("dup") Then

                        ''
                        Dim objCarteira As New [Lib].Negocio.CarteiraFinanceira(objNotaFiscal.Itens(0).Produto.CodigoCarteiraCompra)
                        cmbCarteira.SelectedValue = objNotaFiscal.Itens(0).Produto.CodigoCarteiraCompra
                        cmbFormas.SelectedValue = 1

                        Dim objTributo As New [Lib].Negocio.Encargo()

                        Dim objPagamentos As CondicaoPagamento = New CondicaoPagamento(0, dsXML.Tables("dup").Rows.Count)

                        If objPagamentos.Codigo <> 0 Then

                            ddlCondicaoDePagamento.SelectedValue = objPagamentos.Codigo

                            If Not String.IsNullOrWhiteSpace(ddlTributos.SelectedValue) Then objTributo = New [Lib].Negocio.Encargo(ddlTributos.SelectedValue)

                            ''VER COM O FURLAN
                            'If String.IsNullOrWhiteSpace(objCarteira.CodigoContaCliente) AndAlso String.IsNullOrWhiteSpace(objTributo.Codigo) Then
                            '    MsgBox(Me.Page, "Encargo não foi selecionado.")
                            '    Exit Sub
                            'End If

                            'SessaoRecuperaNotaFiscal()

                            If objNota.IUD = "I" Then
                                cmbFormas.Enabled = True
                            Else
                                cmbFormas.Enabled = False
                            End If

                            objNota.VencimentosNota = New ListTitulo()

                            objNota.VencimentosNota.NF = objNota
                            hdCondicaoDePagamento.Value = ddlCondicaoDePagamento.SelectedValue

                            If FinanceiroVirtual Then
                                objNota.VencimentosNota.ParcelarNotasFiscaisGerais(hdCondicaoDePagamento.Value, CDec(txtTotalNota.Text), CDec(txtTotalParcelado.Text), CDec(txtTotalPago.Text))
                            Else
                                objNota.VencimentosNota.ParcelarNotasFiscaisGerais(hdCondicaoDePagamento.Value, CDec(txtTotalNota.Text), CDec(txtTotalParcelado.Text), CDec(txtTotalPago.Text), IIf(chkProvisao.Checked, 3, 2))
                            End If

                            For i As Integer = 0 To dsXML.Tables("dup").Rows.Count - 1

                                objNota.VencimentosNota(i).Vencimento = Funcoes.ValidaDataUtil(objNota.Empresa.Codigo, objNota.Empresa.CodigoEndereco, dsXML.Tables("dup").Rows(i).Item("dVenc"))
                                objNota.VencimentosNota(i).Prorrogacao = Funcoes.ValidaDataUtil(objNota.Empresa.Codigo, objNota.Empresa.CodigoEndereco, dsXML.Tables("dup").Rows(i).Item("dVenc"))
                                objNota.VencimentosNota(i).ValorDoDocumento = dsXML.Tables("dup").Rows(i).Item("vDup").ToString.Replace(".", ",")

                            Next

                            txtTotalParcelado.Text = objNota.VencimentosNota.Where(Function(s) s.IUD <> "D").Sum(Function(s) s.ValorDoDocumento).ToString("N2")
                            txtSaldoVencimentos.Text = CDec(txtTotalParcelado.Text) - CDec(txtTotalPago.Text)

                            For Each rowTit As [Lib].Negocio.Titulo In objNota.VencimentosNota
                                rowTit.CodigoCarteira = objCarteira.CodigoCarteira
                                rowTit.CodigoTipoPgto = cmbFormas.SelectedValue
                                rowTit.Tributo = IIf(Not String.IsNullOrWhiteSpace(objTributo.Codigo) AndAlso objTributo.Codigo.Length > 0, objTributo.Codigo, "")
                                If Not String.IsNullOrWhiteSpace(objTributo.Codigo) AndAlso objTributo.Codigo.Length > 0 Then
                                    rowTit.ContaContabilCliente = objTributo.ContaCredito
                                Else
                                    rowTit.ContaContabilCliente = objCarteira.CodigoContaCliente
                                End If
                            Next

                            SessaoSalvaNotaFiscal()

                            grdCondicoes.DataSource = objNota.VencimentosNota.Where(Function(s) s.IUD <> "D")
                            grdCondicoes.DataBind()

                        Else

                            MsgBox(Me.Page, "Não foi possível localizar o número de parcelas desse XML !")

                        End If



                        ''

                    End If

                End If

            End If
        End If







        If pXmlNFe Then
            'Desabilitado temporariamente - Furlan - 02/06/2022
            'lnkAdicionarItem_Click(lnkAdicionarItem, Nothing)
            Exit Sub
        End If

        If (Not objNota.NotasTrocaOrigem Is Nothing AndAlso objNota.NotasTrocaOrigem.Count > 0) OrElse
            objNota.TipoDeDocumento.Codigo = eTipoDeDocumento.CTRC OrElse
            objNota.TipoDeDocumento.Codigo = eTipoDeDocumento.CT_E OrElse
            objNota.TipoDeDocumento.Codigo = eTipoDeDocumento.Estadia OrElse
            objNota.TipoDeDocumento.Codigo = eTipoDeDocumento.ComplementoDeFrete Then
            TabNFOrigem.Visible = True
            grdNotasReferenciais.Visible = False
            grdNotasFretes.Visible = True

            divPedidoReferenciaRPA.Style.Add("display", "none")

            grdNotasFretes.DataSource = objNota.NotasTrocaOrigem
            grdNotasFretes.DataBind()

            Dim i As Integer = 0
            While i < grdNotasFretes.Rows.Count
                CType(grdNotasFretes.Rows(i).FindControl("imgExcluirNF"), ImageButton).Visible = False

                i += 1
            End While
        End If


        If Not objNota.NotaTrocaDestino Is Nothing AndAlso objNota.NotasOrigemDestino.Count > 0 Then
            TabNFOrigem.Visible = True
            grdNotasReferenciais.Visible = False
            grdNotasFretes.Visible = True

            divPedidoReferenciaRPA.Style.Add("display", "none")

            grdNotasFretes.DataSource = objNota.NotasOrigemDestino
            grdNotasFretes.DataBind()

            Dim i As Integer = 0
            While i < grdNotasFretes.Rows.Count
                CType(grdNotasFretes.Rows(i).FindControl("imgExcluirNF"), ImageButton).Visible = False

                i += 1
            End While
        End If


        'Tipo de Documento RPA - Recibo de Pagamento a Autônomo
        If objNota.TipoDeDocumento IsNot Nothing AndAlso objNota.TipoDeDocumento.Codigo = eTipoDeDocumento.RPA Then
            'Dim objListaDeNotasFiscaisReferenciais As [Lib].Negocio.ListNotaFiscalReferencial
            'objListaDeNotasFiscaisReferenciais = New ListNotaFiscalReferencial(objNota, eTipoReferencial.RPA)
            objNota.NotasReferenciais = New ListNotaFiscalReferencial(objNota, eTipoReferencial.RPA)

            If Not objNota.NotasReferenciais Is Nothing AndAlso objNota.NotasReferenciais.Count > 0 Then
                TabNFOrigem.Visible = True
                grdNotasFretes.Visible = False
                grdNotasReferenciais.Visible = True
                divPedidoReferenciaRPA.Style.Remove("display")

                txtPedidoNFReferencial.Text = objNota.NotasReferenciais(0).ParentOrigem.NotaFiscal.CodigoPedido
                txtPedidoNFReferencial.ReadOnly = True
                btnPedidoNFReferencial.Enabled = False

                grdNotasReferenciais.DataSource = objNota.NotasReferenciais
                grdNotasReferenciais.DataBind()

                SessaoRecuperaNotaFiscalOriginal()
                objNotaFiscalOriginal.NotasReferenciais = objNota.NotasReferenciais
                SessaoSalvaNotaFiscalOriginal()
            End If
        End If

        ddlProdutoContabilizacao.Items.Clear()
        If objNota.SubOperacao.Contabil Then
            TabContabil.Visible = True
            ddlProdutoContabilizacao.Items.Add(New ListItem("Todos os Lançamentos", "0"))
            For Each row In objNota.Itens
                ddlProdutoContabilizacao.Items.Add(New ListItem("Produto " & row.CodigoProduto & " - " & row.Produto.Nome, row.CodigoProduto))
            Next

            objNota.LancamentosContabeis.CalcularSaldo()
            gridRazao.DataSource = objNota.LancamentosContabeis.OrderBy(Function(s) s.Sequencia)
            gridRazao.DataBind()

            lblDebito.Text = objNota.LancamentosContabeis.Sum(Function(s) s.DebitoOficial).ToString("N2")
            lblCredito.Text = objNota.LancamentosContabeis.Sum(Function(s) s.CreditoOficial).ToString("N2")
        Else
            TabContabil.Visible = False
        End If

        If objNota.SubOperacao.Financeiro AndAlso objNota.TotalNota > 0 AndAlso objNota.VencimentosNota IsNot Nothing AndAlso objNota.VencimentosNota.Count > 0 Then
            If objNota.CodigoPedido > 0 Then
                hdCondicaoDePagamento.Value = objNota.Pedido.CodigoCondicaoPagamento
            Else
                hdCondicaoDePagamento.Value = objNota.VencimentosNota(0).CodigoTipoPgto
            End If
            txtTotalNota.Text = objNota.TotalNota.ToString("N2")

            TabVencimentosold.Visible = True

            grdCondicoes.DataSource = objNota.VencimentosNota
            grdCondicoes.DataBind()

            For x As Integer = 0 To objNota.VencimentosNota.Count - 1
                If objNota.VencimentosNota(x).CodigoProvisao = 1 Then
                    grdCondicoes.Rows(x).Cells(1).ForeColor = Drawing.Color.Red
                    grdCondicoes.Rows(x).Cells(2).ForeColor = Drawing.Color.Red
                    grdCondicoes.Rows(x).Cells(3).ForeColor = Drawing.Color.Red
                    grdCondicoes.Rows(x).Cells(4).ForeColor = Drawing.Color.Red
                    grdCondicoes.Rows(x).Cells(5).ForeColor = Drawing.Color.Red
                    grdCondicoes.Rows(x).Cells(6).ForeColor = Drawing.Color.Red
                    grdCondicoes.Rows(x).Cells(7).ForeColor = Drawing.Color.Red
                    grdCondicoes.Rows(x).Cells(8).ForeColor = Drawing.Color.Red
                    grdCondicoes.Rows(x).Cells(9).ForeColor = Drawing.Color.Red
                End If
            Next
        Else
            TabVencimentosold.Visible = False
        End If

        objNotaFiscal = objNota
        SessaoSalvaNotaFiscal()

        AtualizarItensNoGrid()

        If Not objNotaFiscal.CodigoRepresentante Is Nothing AndAlso objNotaFiscal.CodigoRepresentante.Length > 0 Then

            Funcoes.FormatarClienteTXT(txtRepresentante, objNotaFiscal.Representante)

        End If

        ddlUsuarios.Items.Clear()
        If Not String.IsNullOrWhiteSpace(objNotaFiscal.UsuarioInclusao) Then ddlUsuarios.Items.Add("Inc.- " & objNotaFiscal.UsuarioInclusao)
        If Not String.IsNullOrWhiteSpace(objNotaFiscal.UsuarioAlteracao) Then ddlUsuarios.Items.Add("Alt.- " & objNotaFiscal.UsuarioAlteracao)

        'ToolBar
        lnkAtualizar.Parent.Visible = objNotaFiscal.IUD = "U"
        lnkNovo.Parent.Visible = objNotaFiscal.IUD <> "U" AndAlso objNotaFiscal.IUD <> "D"
        lnkExcluir.Parent.Visible = objNotaFiscal.IUD <> "I" AndAlso Not String.IsNullOrWhiteSpace(objNotaFiscal.IUD)
        lnkEspelho.Parent.Visible = objNotaFiscal.IUD <> "I" AndAlso objNotaFiscal.CodigoSituacao = 1
        lnkRecontabilizar.Parent.Visible = objNotaFiscal.IUD <> "I" AndAlso objNotaFiscal.CodigoSituacao = 1 AndAlso objNotaFiscal.SubOperacao.Contabil
        lnkConsultar.Parent.Visible = False
        imgExtratoPedido.Visible = objNotaFiscal.IUD = "U"

        If Not String.IsNullOrWhiteSpace(objNotaFiscal.ChaveNFE) AndAlso objNotaFiscal.CodigoTipoDeDocumento = 1 AndAlso objNotaFiscal.Arquivos.Count > 0 Then
            txtChaveNFe.Enabled = False
            txtNumeroNota.Enabled = False
            txtSerie.Enabled = False
            txtDataNota.Enabled = False
            lnkVerificarChaveNFE.Visible = True
        End If
        'Arquivo
        If ucFile.Parent.Visible Then
            ucFile.Bind(objNota.Arquivos)
        End If

        If objNotaFiscal.Empresa.Empresa.ConferenciaNFE Then
            If (objNotaFiscal.Conferencia OrElse Funcoes.VerificaPermissao("Fiscal", "ACESSAR")) Then
                chkConferencia.Visible = True
                divConferencia.Visible = True
                chkConferencia.Checked = objNotaFiscal.Conferencia
                chkConferencia.Enabled = Funcoes.VerificaPermissao("Fiscal", "ACESSAR")
            End If
        End If

        If Not objNotaFiscal.CodigoSituacao = 1 Then
            'ToolBar
            lnkAtualizar.Parent.Visible = False
            lnkNovo.Parent.Visible = False
            lnkExcluir.Parent.Visible = False
            chkEspelho.Checked = False
            chkEspelho.Enabled = False
        ElseIf Not Funcoes.VerificaAcesso(objNotaFiscal.CodigoEmpresa, objNotaFiscal.EnderecoEmpresa, txtMovimento.Text, "NOTAS FISCAIS") Then
            'Toolbar
            lnkAtualizar.Parent.Visible = False
            lnkExcluir.Parent.Visible = False
            lnkAdicionarItem.Parent.Parent.Visible = False
            grdProdutos.Columns(0).Visible = False
            grdProdutos.Columns(13).Visible = False
            'lnkRecontabilizar.Parent.Visible = False
            MsgBox(Me.Page, "Movimento Fiscal já Fechado para esta data")
        ElseIf Not Funcoes.VerificaAcesso(objNotaFiscal.CodigoEmpresa, objNotaFiscal.EnderecoEmpresa, txtMovimento.Text, "CONTABIL") Then
            'ToolBar
            lnkAtualizar.Parent.Visible = False
            lnkExcluir.Parent.Visible = False
            lnkRecontabilizar.Parent.Visible = False
            lnkAdicionarItem.Parent.Parent.Visible = False
            grdProdutos.Columns(0).Visible = False
            grdProdutos.Columns(13).Visible = False
        ElseIf objNotaFiscal.VencimentosNota.Count > 0 AndAlso objNotaFiscal.VencimentosNota.All(Function(s) s.RegistroMestre > 0) Then
            'ToolBar
            lnkExcluir.Parent.Visible = False
            ddlUnidadeNegocio.Enabled = False
            ddlTipoDeDocumento.Enabled = False
            ddlEmpresa.Enabled = False
            btnConsultaClientes.Visible = False

            LnkParcelamento.Enabled = False
            cmbCarteira.Enabled = False
            ddlTributos.Enabled = False
            ddlCondicaoDePagamento.Enabled = False
            cmbFormas.Enabled = False
            MsgBox(Me.Page, "Atenção, só é permitido alterar notas com Titulos agrupados desde que o Valor Líquido de Nota Fiscal não seja alterado")
        ElseIf objNotaFiscal.VencimentosNota.Where(Function(s) s.CodigoProvisao = 1).ToList.Count > 0 Then
            'ToolBar
            lnkExcluir.Parent.Visible = False
            ddlUnidadeNegocio.Enabled = False
            ddlTipoDeDocumento.Enabled = False
            ddlEmpresa.Enabled = False
            btnConsultaClientes.Visible = False

            LnkParcelamento.Enabled = False
            cmbCarteira.Enabled = False
            ddlTributos.Enabled = False
            ddlCondicaoDePagamento.Enabled = False
            cmbFormas.Enabled = False
            If objNotaFiscal.VencimentosNota.All(Function(s) s.CodigoProvisao = 1) Then
                MsgBox(Me.Page, "Atenção, só é permitido alterar notas com Titulos Baixados desde que o Valor Líquido de Nota Fiscal não seja alterado")
            End If

        ElseIf FinanceiroNovo AndAlso objNotaFiscal.Titulos.Where(Function(s) s.CodigoSituacao And s.CodigoProvisao = 1).ToList.Count > 0 Then
            MsgBox(Me.Page, "Atenção, só é permitido alterar notas com Titulos Baixados desde que o Valor Líquido de Nota Fiscal não seja alterado")
            'ToolBar
            lnkExcluir.Parent.Visible = False
            ddlUnidadeNegocio.Enabled = False
            ddlTipoDeDocumento.Enabled = False
            ddlEmpresa.Enabled = False
            btnConsultaClientes.Visible = False
        ElseIf Not objNota.NotaTrocaDestino Is Nothing AndAlso objNota.NotasOrigemDestino.Count > 0 Then
            If objNotaFiscal.CodigoTipoDeDocumento = CInt(eTipoDeDocumento.CTRC) OrElse
                objNotaFiscal.CodigoTipoDeDocumento = CInt(eTipoDeDocumento.CT_E) OrElse
                objNotaFiscal.CodigoTipoDeDocumento = CInt(eTipoDeDocumento.CT_E_TOM) OrElse
                objNotaFiscal.CodigoTipoDeDocumento = CInt(eTipoDeDocumento.CT_E_OUT) OrElse
                objNotaFiscal.CodigoTipoDeDocumento = CInt(eTipoDeDocumento.Estadia) OrElse
                objNotaFiscal.CodigoTipoDeDocumento = CInt(eTipoDeDocumento.ComplementoDeFrete) Then
                'NÃO FAZ NADA
            Else
                lnkAtualizar.Parent.Visible = False
                lnkExcluir.Parent.Visible = False
                lnkRecontabilizar.Parent.Visible = False
                lnkAdicionarItem.Parent.Parent.Visible = False
                grdProdutos.Columns(0).Visible = False
                grdProdutos.Columns(13).Visible = False
                MsgBox(Me.Page, "Nota Fiscal com Vinculo não pode ser alterada nem excluída, apenas consulta.", eTitulo.Info)
            End If
        ElseIf Not Funcoes.VerificaPermissao("NotasFiscaisGerais", "ALTERAR") Then
            lnkAtualizar.Parent.Visible = False
        End If



    End Sub

    Public Sub LimparCampos()
        If Not objNotaFiscal Is Nothing Then
            If Not String.IsNullOrWhiteSpace(objNotaFiscal.IUD) Then
                SessaoRecuperaNotaFiscalOriginal()
                logs = New FuncoesLogs(2, Session("ssCnpjDaEmpresa"))
                logs.RegistrarLog(objNotaFiscal.IUD, objNotaFiscal)
            End If
        End If

        SessaoDsXML = Nothing
        SessaoXmlDoc = Nothing

        lblUsuario.Text = Session("ssNomeUsuario")
        Session.Remove("objConsultarNaviosXInvoice")

        lblPedido.Text = "0"
        txtNumeroNota.Text = String.Empty
        txtSerie.Text = String.Empty
        txtChaveNFe.Text = String.Empty
        txtObservacao.Text = String.Empty
        txtMovimento.Text = String.Empty
        txtDataNota.Text = String.Empty
        txtPedidoNFReferencial.ReadOnly = False
        txtPedidoNFReferencial.Text = String.Empty
        btnPedidoNFReferencial.Enabled = True

        txtMovimento.Enabled = True
        cmbCarteira.Enabled = True

        lnkComissoes.Parent.Visible = False
        grdProdutos.DataSource = Nothing
        grdProdutos.DataBind()

        gridEncargosGerais.DataSource = Nothing
        gridEncargosGerais.DataBind()

        gridRazao.DataSource = Nothing
        gridRazao.DataBind()
        lblDebito.Text = "0,00"
        lblCredito.Text = "0,00"
        TabContabil.Visible = False

        ddlNaturezaDeRendimento.Items.Clear()
        divNaturezaDeRendimento.Visible = False
        chkNaturezaDeRendimento.Checked = False

        txtDataNota.Enabled = True
        txtNumeroNota.Enabled = True
        txtSerie.Enabled = True
        txtChaveNFe.Enabled = True
        ddlSituacao.Enabled = True

        btnRepresentante.Enabled = True
        txtRepresentante.Enabled = False
        txtRepresentante.Text = String.Empty

        'Toolbar
        lnkNovo.Parent.Visible = False
        lnkAtualizar.Parent.Visible = False
        lnkExcluir.Parent.Visible = False
        lnkEspelho.Parent.Visible = False
        lnkRecontabilizar.Parent.Visible = False
        lnkConsultar.Parent.Visible = True

        ddlUnidadeNegocio.Enabled = True
        ddlTipoDeDocumento.Enabled = True
        ddlEmpresa.Enabled = True
        btnConsultaClientes.Visible = True
        cmbCarteira.Enabled = True
        ddlTributos.Enabled = True
        LnkParcelamento.Enabled = True
        ddlCondicaoDePagamento.Enabled = True
        cmbFormas.Enabled = True
        grdCondicoes.Enabled = True
        cmdOkVencimento.Enabled = True
        txtDataVencimento.Enabled = True
        txtCodigoDeBarras.Text = String.Empty
        txtCodigoDeBarras.Enabled = False
        BtValidarCodBarras.Enabled = False
        CkbCodigoDeBarras.Checked = False
        ckPreImpresso.Checked = False
        CkbCodigoDeBarras.Enabled = False
        ckPreImpresso.Enabled = False

        chkProvisao.Checked = False
        If FinanceiroVirtual Then chkProvisao.Visible = False

        lnkVerificarChaveNFE.Visible = True
        If chkEspelho.Enabled = False Then chkEspelho.Enabled = True

        ddlProdutoContabilizacao.Items.Clear()
        imgExtratoPedido.Visible = False
        lnkAdicionarItem.Parent.Parent.Visible = True
        grdProdutos.Columns(0).Visible = True
        grdProdutos.Columns(13).Visible = True

        'Tab Vencimentos
        TabVencimentosold.Visible = False
        divTitulo.Visible = False

        grdNotasFretes.DataSource = Nothing
        grdNotasFretes.DataBind()
        txtDataInicio.Text = Date.Now.ToString("dd/MM/yyyy")
        txtDataFim.Text = Date.Now.ToString("dd/MM/yyyy")
        grdNotasReferenciais.DataSource = Nothing
        grdNotasReferenciais.DataBind()
        TabNFOrigem.Visible = False

        lblBanco.Text = "Banco"
        lblAgencia.Text = "Agência"
        lblContaCorrente.Text = "Conta"
        divBanco.Visible = False

        hdCondicaoDePagamento.Value = 0
        ddlCondicaoDePagamento.SelectedIndex = 0
        cmbFormas.SelectedIndex = 0
        txtDataVencimento.Text = String.Empty
        txtValorVencimento.Text = String.Empty
        grdCondicoes.DataSource = Nothing
        grdCondicoes.DataBind()

        txtTotalNota.Text = "0,00"
        txtTotalParcelado.Text = "0,00"
        txtTotalPago.Text = "0,00"
        txtSaldoVencimentos.Text = "0,00"

        chkConferencia.Visible = False
        divConferencia.Visible = False
        '*****************

        objNotaFiscalOriginal = Nothing
        SessaoSalvaNotaFiscalOriginal()

        If chkReaproveitar.Checked Then
            SessaoRecuperaNotaFiscal()
            Dim nf As New NotaFiscal
            nf.IUD = "I"
            nf.Usuario = Session("ssNomeUsuario")
            nf.UsuarioInclusao = Session("ssNomeUsuario")
            nf.DataInclusao = Date.Now.ToString("yyyy-MM-dd")

            nf.CodigoUnidadeDeNegocio = objNotaFiscal.CodigoUnidadeDeNegocio
            nf.CodigoEmpresa = objNotaFiscal.CodigoEmpresa
            nf.EnderecoEmpresa = objNotaFiscal.EnderecoEmpresa
            nf.CodigoCliente = objNotaFiscal.CodigoCliente
            nf.EnderecoCliente = objNotaFiscal.EnderecoCliente
            hdfCodigoCliente.Value = objNotaFiscal.CodigoCliente & "-" & objNotaFiscal.EnderecoCliente
            nf.Movimento = objNotaFiscal.Movimento
            nf.DataNota = objNotaFiscal.DataNota
            nf.CodigoTipoDeDocumento = objNotaFiscal.CodigoTipoDeDocumento
            ddlTipoDeDocumento.SelectedValue = nf.CodigoTipoDeDocumento
            nf.CodigoSituacao = objNotaFiscal.CodigoSituacao
            nf.CodigoOperacao = objNotaFiscal.CodigoOperacao
            nf.CodigoSubOperacao = objNotaFiscal.CodigoSubOperacao
            nf.CodigoSafra = objNotaFiscal.CodigoSafra
            ddlSafra.SelectedValue = nf.CodigoSafra

            Dim objItemNF As [Lib].Negocio.NotaFiscalXItem

            For Each it In objNotaFiscal.Itens
                objItemNF = New [Lib].Negocio.NotaFiscalXItem(nf)
                With objItemNF
                    .CodigoProduto = it.CodigoProduto
                    .PesoQuantidade = it.Produto.PesoQuantidade
                    .CodigoOperacao = nf.CodigoOperacao
                    .CodigoSubOperacao = nf.CodigoSubOperacao
                    .TemCentroDeCusto = it.TemCentroDeCusto
                    If Not it.CodigoProdutoCusto Is Nothing AndAlso it.CodigoProdutoCusto.Length > 0 Then
                        .CodigoProdutoCusto = it.CodigoProdutoCusto
                    Else
                        .CodigoProdutoCusto = it.CentroDeCustoInformado
                    End If

                    objItemNF.CarregandoEncargos = True
                    .Encargos = New [Lib].Negocio.ListNotaFiscalXItemXEncargo(objItemNF, New List(Of eEtapaEncago) From {eEtapaEncago.Normal})
                    objItemNF.CarregandoEncargos = False
                End With
                nf.Itens.Add(objItemNF)
            Next

            objNotaFiscal = nf

            If objNotaFiscal.CodigoTipoDeDocumento = eTipoDeDocumento.CTRC OrElse
                objNotaFiscal.CodigoTipoDeDocumento = eTipoDeDocumento.Estadia OrElse
                objNotaFiscal.CodigoTipoDeDocumento = eTipoDeDocumento.ComplementoDeFrete OrElse
                objNotaFiscal.CodigoTipoDeDocumento = eTipoDeDocumento.CT_E Then
                TabNFOrigem.Visible = True
                grdNotasFretes.Visible = True
                grdNotasReferenciais.Visible = False
                divPedidoReferenciaRPA.Style.Add("display", "none")
            ElseIf ddlTipoDeDocumento.SelectedValue = CInt(eTipoDeDocumento.RPA) Then 'RPA
                TabNFOrigem.Visible = True
                grdNotasFretes.Visible = False
                grdNotasReferenciais.Visible = True
                divPedidoReferenciaRPA.Style.Remove("display")
            End If

            lnkNovo.Parent.Visible = True

            HttpContext.Current.Session("ssMessage") = String.Empty

        Else
            objNotaFiscal = New NotaFiscal()
            objNotaFiscal.IUD = "I"
            objNotaFiscal.CodigoUnidadeDeNegocio = ddlUnidadeNegocio.SelectedValue
            Dim strEmpresa As String() = ddlEmpresa.SelectedValue.Split("-")
            objNotaFiscal.CodigoEmpresa = strEmpresa(0)
            objNotaFiscal.EnderecoEmpresa = strEmpresa(1)
            objNotaFiscal.Usuario = Session("ssNomeUsuario")
            objNotaFiscal.UsuarioInclusao = Session("ssNomeUsuario")
            objNotaFiscal.DataInclusao = Date.Now.ToString("yyyy-MM-dd")

            txtNomeCliente.Text = String.Empty
            hdfCodigoCliente.Value = String.Empty
            ddlTipoDeDocumento.SelectedIndex = 0
            ddlSituacao.SelectedValue = 1
            cmbCarteira.SelectedIndex = 0
            ddlTributos.Items.Clear()
            With ddlSafra
                ddlSafra.SelectedIndex = .Items.IndexOf(.Items.FindByValue("NENHUMA"))
            End With
        End If

        ddlUsuarios.Items.Clear()
        ddlUsuarios.Items.Add("Inc.- " & Session("ssNomeUsuario"))

        Session.Remove("objNotaFiscal" & HID.Value)
        Session.Remove("objClienteNFReferencial" & HID.Value)
        Session.Remove("objPedidoNFReferencial" & HID.Value)
        Session.Remove("objNotaFiscalReferencial" & HID.Value)
        Session.Remove("objRepresentante" & HID.Value)
        Session.Remove("chaveXMLautomacao" & HID.Value)

        HID.Value = Guid.NewGuid().ToString
        ucConsultaClientes.SetarHID(HID.Value)
        ucConsultaDadosBancarios.SetarHID(HID.Value)
        ucConsultaPedidosXNotas.SetarHID(HID.Value)
        ucComissoesXBaixas.SetarHID(HID.Value)
        ucNFEncargo.SetarHID(HID.Value)
        ucProdutoNFG.SetarHID(HID.Value)
        ucInformarDadosProdutoNFG.SetarHID(HID.Value)
        ucNFOrigem.SetarHID(HID.Value)
        ucRateio.SetarHID(HID.Value)
        ucNotaFiscalReferencial.SetarHID(HID.Value)
        ucNFObsProduto.SetarHID(HID.Value)
        ucConsultarNaviosXInvoice.SetarHID(HID.Value)
        ucFile.Clear()

        objNotaFiscal.NFG = True
        SessaoSalvaNotaFiscal()

        'Verifica se a empresa está habilitada para gravar arquivo
        Dim Empresa As New [Lib].Negocio.Cliente(UsuarioServidor.CodigoEmpresa, UsuarioServidor.EnderecoEmpresa)
        ucFile.Parent.Visible = Not String.IsNullOrWhiteSpace(Empresa.Empresa.PathDownloadNFe)

        If Empresa.Empresa.ObrigaNavio Then
            divNaviosXInvoice.Visible = True
        Else
            divNaviosXInvoice.Visible = False
        End If

        SessaoRecuperaNotaFiscal()

        If chkReaproveitar.Checked AndAlso objNotaFiscal.Itens IsNot Nothing AndAlso objNotaFiscal.Itens.Count > 0 Then AtualizarItensNoGrid()
        LiberaEmpresa()
    End Sub

    Private Sub LiberaEmpresa()
        If Not UsuarioServidor.LiberaEmpresa Then
            ddlUnidadeNegocio.Enabled = False
            ddlEmpresa.Enabled = False
        End If
    End Sub

    Private Sub MostrarMensagemFinal()
        If objNotaFiscalOriginal IsNot Nothing Then
            If objNotaFiscalOriginal.IUD = "D" Then
                mensagemErro = "A nota " & objNotaFiscalOriginal.Codigo & "-" & objNotaFiscalOriginal.Serie & " foi deletada com sucesso! \n"
                mensagemErro &= "A nota " & objNotaFiscal.Codigo & "-" & objNotaFiscal.Serie & " foi gerada com sucesso! \n"
            Else
                mensagemErro = "A nota " & objNotaFiscalOriginal.Codigo & "-" & objNotaFiscalOriginal.Serie & " foi " & IIf(objNotaFiscalOriginal.IUD = "U", "alterada", "gerada") & " com sucesso! \n"
            End If
        Else
            mensagemErro = "A nota " & objNotaFiscal.Codigo & "-" & objNotaFiscal.Serie & " foi " & IIf(objNotaFiscal.IUD = "U", "alterada", "gerada") & " com sucesso! \n"
        End If

        mensagemErro &= "Foi criado também o pedido número " & objNotaFiscal.CodigoPedido & " \n"

        If objNotaFiscal.SubOperacao.Financeiro Then
            mensagemErro &= "E o" & IIf(objNotaFiscal.VencimentosNota.Count <> 1, "s", "") & " seguinte" & IIf(objNotaFiscal.VencimentosNota.Count <> 1, "s", "") & " título" & IIf(objNotaFiscal.VencimentosNota.Count <> 1, "s", "") & ": \n"

            If FinanceiroNovo Then
                For Each row In objNotaFiscal.Titulos.Where(Function(t) t.IUD <> "D")
                    mensagemErro &= "  *" & row.Codigo & ", com vencimento em " &
                                   Convert.ToDateTime(row.Reprogramacao).ToString("dd/MM/yyyy") & ", no valor de " &
                                   Convert.ToDouble(row.Valores.EncargoValorDocumento.Valor).ToString("N2") & " \n"
                Next
            Else
                For Each row In objNotaFiscal.VencimentosNota.Where(Function(t) t.IUD <> "D")
                    mensagemErro &= "  *" & row.Codigo & ", com vencimento em " &
                                   Convert.ToDateTime(row.Prorrogacao).ToString("dd/MM/yyyy") & ", no valor de " &
                                   Convert.ToDouble(row.ValorDoDocumento).ToString("N2") & " \n"
                Next
            End If
        End If

        MsgBox(Me.Page, mensagemErro, eTitulo.Info, False)
    End Sub

    Private Sub AtualizarCFOP(ByVal Consulta As Boolean)
        SessaoRecuperaNotaFiscal()

        If (objNotaFiscal.CodigoTipoDeDocumento = 2 Or
            objNotaFiscal.CodigoTipoDeDocumento = 10 Or
            objNotaFiscal.CodigoTipoDeDocumento = 14 Or
            objNotaFiscal.CodigoTipoDeDocumento = 57) AndAlso
            Not objNotaFiscal.NotasTrocaOrigem Is Nothing AndAlso
            objNotaFiscal.NotasTrocaOrigem.Count > 0 AndAlso
            objNotaFiscal.Cliente.CodigoEstado <> objNotaFiscal.NotasTrocaOrigem(0).Cliente.CodigoEstado Then

            For Each item In objNotaFiscal.Itens

                If Not item.OperacaoEstado Is Nothing AndAlso item.OperacaoEstado.Encargos.Count > 0 Then

                    If Not item.CFOP = item.OperacaoEstado.CodigoFiscal Then
                        item.CFOP = item.OperacaoEstado.CodigoFiscal
                        item.Encargos.ForEach(Function(s)
                                                  s.EstadoDestino = item.OperacaoEstado.EstadoDestino
                                                  Return True
                                              End Function)
                    End If

                Else
                    LimparCampos()
                    MsgBox(Me.Page, "Não foram encontrados encargos do produto " & item.CodigoProduto & "-" & item.Produto.Descricao & " para o Cliente selecionado.")
                    Exit Sub
                End If
            Next
        End If

        If objNotaFiscal IsNot Nothing AndAlso objNotaFiscal.Itens IsNot Nothing AndAlso objNotaFiscal.Itens.Count > 0 Then
            Dim cfop As Integer = objNotaFiscal.Itens(0).CFOP
            If Not objNotaFiscal.CodigoTipoDeDocumento = CInt(eTipoDeDocumento.CTRC) AndAlso
                Not objNotaFiscal.CodigoTipoDeDocumento = CInt(eTipoDeDocumento.CTRC_SEM_NF) AndAlso
                Not objNotaFiscal.CodigoTipoDeDocumento = CInt(eTipoDeDocumento.CT_E_TOM) AndAlso
                Not objNotaFiscal.CodigoTipoDeDocumento = CInt(eTipoDeDocumento.CT_E) AndAlso
                Not objNotaFiscal.CodigoTipoDeDocumento = CInt(eTipoDeDocumento.CT_E_OUT) AndAlso
                Not objNotaFiscal.CodigoTipoDeDocumento = CInt(eTipoDeDocumento.ComplementoDeFrete) AndAlso
                Not objNotaFiscal.CodigoTipoDeDocumento = CInt(eTipoDeDocumento.Estadia) Then
                If objNotaFiscal.TotalNota > 0 AndAlso (cfop = 1352 OrElse cfop = 2352 OrElse cfop = 5352 OrElse cfop = 6352 OrElse cfop = 1353 OrElse cfop = 2353 OrElse cfop = 5353 OrElse cfop = 6353) Then
                    If Consulta Then
                        MsgBox(Me.Page, "Verifique esse lançamento, cfop de transporte não pode ser usado para esse tipo de documento!")
                    Else
                        MsgBox(Me.Page, "Cfop de transporte não pode ser usado para esse tipo de documento!")
                        LimparCampos()
                        Exit Sub
                    End If
                End If
            End If

            If objNotaFiscal.CodigoTipoDeDocumento = CInt(eTipoDeDocumento.CTRC) OrElse
                objNotaFiscal.CodigoTipoDeDocumento = CInt(eTipoDeDocumento.CTRC_SEM_NF) OrElse
                objNotaFiscal.CodigoTipoDeDocumento = CInt(eTipoDeDocumento.CT_E_TOM) OrElse
                objNotaFiscal.CodigoTipoDeDocumento = CInt(eTipoDeDocumento.CT_E) OrElse
                objNotaFiscal.CodigoTipoDeDocumento = CInt(eTipoDeDocumento.CT_E_OUT) OrElse
                objNotaFiscal.CodigoTipoDeDocumento = CInt(eTipoDeDocumento.ComplementoDeFrete) OrElse
                objNotaFiscal.CodigoTipoDeDocumento = CInt(eTipoDeDocumento.Estadia) Then
                If objNotaFiscal.TotalNota > 0 AndAlso Not cfop = 1352 AndAlso Not cfop = 2352 AndAlso Not cfop = 5352 AndAlso Not cfop = 6352 AndAlso Not cfop = 1353 AndAlso Not cfop = 2353 AndAlso Not cfop = 5353 AndAlso Not cfop = 6353 Then
                    If (Not objNotaFiscal.CodigoTipoDeDocumento = CInt(eTipoDeDocumento.CTRC_SEM_NF) Or Not objNotaFiscal.CodigoTipoDeDocumento = CInt(eTipoDeDocumento.CT_E_TOM)) AndAlso (cfop = 1933 OrElse cfop = 2933 OrElse cfop = 5933 OrElse cfop = 6933) Then
                        If Consulta Then
                            MsgBox(Me.Page, "Verifique esse lançamento, cfop de transporte não pode ser usado para esse tipo de documento!")
                        Else
                            MsgBox(Me.Page, "Para o tipo de documento selecionado deve ser informado uma operação com cfop de transporte!")
                            LimparCampos()
                            Exit Sub
                        End If
                    End If
                End If
            End If

            grdProdutos.DataSource = objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D").ToList.OrderBy(Function(s) s.Sequencia)
            grdProdutos.DataBind()

            For ni As Integer = 0 To objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D").Count - 1
                If objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(ni).Encargos.Count > 0 Then
                    If objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D")(ni).Encargos.EncProduto.CentroDeCusto = 0 Then
                        CType(grdProdutos.Rows(ni).FindControl("btnRateio"), Button).Visible = True
                    Else
                        CType(grdProdutos.Rows(ni).FindControl("btnRateio"), Button).Visible = False
                    End If
                End If
            Next

            SessaoSalvaNotaFiscal()
        End If
    End Sub

    Public Sub getSqlException(ByRef Sqls As ArrayList, ByVal objPedido As [Lib].Negocio.Pedido)
        Dim sql As String = "BEGIN TRY " & vbCrLf &
               "DECLARE @HORA_BLOQUEIO AS DATETIME = DATEADD(MINUTE, 3, (SELECT VERSAOHORARIOBLOQUEIO FROM PEDIDOS WHERE PEDIDO_ID = '" & objPedido.Codigo & "' AND EMPRESA_ID = '" & objPedido.CodigoEmpresa & "' AND ENDEMPRESA_ID = '" & objPedido.EnderecoEmpresa & "')) " & vbCrLf &
               "PRINT 'HORA_ATUAL: ' + CAST(GETDATE() AS VARCHAR); " & vbCrLf &
               "PRINT 'HORA_BLOQUEIO: ' + CAST(@HORA_BLOQUEIO AS VARCHAR); " & vbCrLf &
               "IF (GETDATE() > @HORA_BLOQUEIO) " & vbCrLf &
               "BEGIN " & vbCrLf &
               "RAISERROR ('POR FAVOR, ATUALIZE O SALDO FINANCEIRO PARA REALIZAR ESTA AÇÃO!', " & vbCrLf &
               "16, " & vbCrLf &
               "1); " & vbCrLf &
               "END " & vbCrLf &
               "UPDATE PEDIDOS SET VERSAOHORARIOBLOQUEIO = NULL WHERE PEDIDO_ID = '" & objPedido.Codigo & "' AND EMPRESA_ID = '" & objPedido.CodigoEmpresa & "' AND ENDEMPRESA_ID = '" & objPedido.EnderecoEmpresa & "'; " & vbCrLf &
               "END TRY " & vbCrLf &
               "BEGIN CATCH " & vbCrLf &
               "DECLARE @ErrorMessage NVARCHAR(4000); " & vbCrLf &
               "DECLARE @ErrorSeverity INT; " & vbCrLf &
               "DECLARE @ErrorState INT; " & vbCrLf &
               "SELECT " & vbCrLf &
               "@ErrorMessage = ERROR_MESSAGE(), " & vbCrLf &
               "@ErrorSeverity = ERROR_SEVERITY(), " & vbCrLf &
               "@ErrorState = ERROR_STATE(); " & vbCrLf &
               "RAISERROR (@ErrorMessage, " & vbCrLf &
               "@ErrorSeverity, " & vbCrLf &
               "@ErrorState); " & vbCrLf &
               "END CATCH;"
        Sqls.Add(sql)
    End Sub

    Protected Sub CarregarDocumento()
        If Session("ssConferenciaNFe") IsNot Nothing Then
            Dim param As New Dictionary(Of String, Object)
            param = Session("ssConferenciaNFe")
            If param IsNot Nothing Then
                Session("ssCnpjDaEmpresa" & HID.Value) = param("CodigoEmpresa")
                Session("ssEndDaEmpresa" & HID.Value) = param("EnderecoEmpresa")
                Session("txtCnpjDoCliente" & HID.Value) = param("CodigoCliente")
                Session("txtEndDoCliente" & HID.Value) = param("EnderecoCliente")
                Session("txtSerie" & HID.Value) = param("Serie")
                Session("txtNumero" & HID.Value) = param("Nota")
                Session("ssCampo" & HID.Value) = "NotaFiscalGeral"

                Popup.ConsultaDePedidosXNotas(Me.Page, "objNFConsultaNFG" & HID.Value)
                Dim numberRows As Integer = ucConsultaPedidosXNotas.BindGridView()
                If numberRows = 1 Then
                    Popup.CloseDialog(Me.Page, "objNFConsultaNFG" & HID.Value)
                End If
                'Habilita os campos para conferencia
                divConferencia.Visible = True
                chkConferencia.Visible = True
                Session.Remove("ssConferenciaNFe")
            End If
        End If

    End Sub
#End Region

#Region "ToolBar"

    Protected Sub lnkConsultar_Click(sender As Object, e As EventArgs) Handles lnkConsultar.Click
        Try
            If Funcoes.VerificaPermissao("NotasFiscaisGerais", "LEITURA") Then

                If String.IsNullOrWhiteSpace(txtNumeroNota.Text) And (String.IsNullOrWhiteSpace(txtMovimento.Text) Or String.IsNullOrWhiteSpace(txtDataNota.Text)) And String.IsNullOrWhiteSpace(hdfCodigoCliente.Value) Then
                    MsgBox(Me.Page, "Informe o número da Nota Fiscal ou período: Data Inicial(Movimento) até Data Final(Data da Nota) ou então um cliente!")
                    Exit Sub
                End If

                If Not String.IsNullOrWhiteSpace(ddlEmpresa.SelectedValue) Then
                    Dim empresa As String() = ddlEmpresa.SelectedValue.Split("-")
                    Session("ssCnpjDaEmpresa" & HID.Value) = empresa(0)
                    Session("ssEndDaEmpresa" & HID.Value) = empresa(1)
                Else
                    Session.Remove("ssCnpjDaEmpresa" & HID.Value)
                    Session.Remove("ssEndDaEmpresa" & HID.Value)
                End If

                If Not String.IsNullOrWhiteSpace(hdfCodigoCliente.Value) Then
                    Dim cliente As String() = hdfCodigoCliente.Value.Split("-")
                    If cliente IsNot Nothing AndAlso cliente.Length > 0 Then
                        Session("txtCnpjDoCliente" & HID.Value) = cliente(0)
                        Session("txtEndDoCliente" & HID.Value) = cliente(1)
                    End If
                Else
                    Session.Remove("txtCnpjDoCliente" & HID.Value)
                    Session.Remove("txtEndDoCliente" & HID.Value)
                End If

                If Not String.IsNullOrWhiteSpace(txtMovimento.Text) Then
                    Session("txtDataDeEmissao" & HID.Value) = txtMovimento.Text
                Else
                    Session.Remove("txtDataDeEmissao" & HID.Value)
                End If

                If Not String.IsNullOrWhiteSpace(txtDataNota.Text) Then
                    Session("txtDataDeEntrada" & HID.Value) = txtDataNota.Text
                Else
                    Session.Remove("txtDataDeEntrada" & HID.Value)
                End If

                Session("txtSerie" & HID.Value) = txtSerie.Text
                Session("txtNumero" & HID.Value) = txtNumeroNota.Text
                Session("ssCampo" & HID.Value) = "NotaFiscalGeral"

                Popup.ConsultaDePedidosXNotas(Me.Page, "objNFConsultaNFG" & HID.Value)
                Dim numberRows As Integer = ucConsultaPedidosXNotas.BindGridView()
                If numberRows = 1 Then
                    Popup.CloseDialog(Me.Page, "objNFConsultaNFG" & HID.Value)
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para consultar registro!")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível consultar a nota")
        End Try
    End Sub

    Protected Sub lnkNovo_Click(sender As Object, e As EventArgs) Handles lnkNovo.Click
        Try
            If Funcoes.VerificaPermissao("NotasFiscaisGerais", "GRAVAR") Then

                If String.IsNullOrWhiteSpace(txtDataNota.Text) OrElse String.IsNullOrWhiteSpace(txtMovimento.Text) Then
                    MsgBox(Me.Page, "Informe a data da Nota e a data de Movimento.")
                    Exit Sub
                End If

                SessaoRecuperaNotaFiscal()

                objNotaFiscal.Codigo = IIf(String.IsNullOrWhiteSpace(txtNumeroNota.Text), 0, txtNumeroNota.Text)
                objNotaFiscal.Serie = txtSerie.Text.ToUpper
                objNotaFiscal.NotaProdutor = IIf(String.IsNullOrWhiteSpace(txtNumeroNota.Text), 0, txtNumeroNota.Text)
                objNotaFiscal.SerieNotaProdutor = txtSerie.Text.ToUpper
                objNotaFiscal.CarregandoNota = True
                objNotaFiscal.DataNota = CDate(txtDataNota.Text)
                objNotaFiscal.Movimento = CDate(txtMovimento.Text)
                objNotaFiscal.CodigoTipoDeDocumento = IIf(ddlTipoDeDocumento.SelectedIndex = 0, 0, ddlTipoDeDocumento.SelectedValue)
                objNotaFiscal.CodigoSituacao = IIf(ddlSituacao.SelectedIndex = 0, 0, ddlSituacao.SelectedValue)
                objNotaFiscal.CodigoSafra = ddlSafra.SelectedValue
                objNotaFiscal.ChaveNFE = txtChaveNFe.Text.Replace(".", "")
                objNotaFiscal.CodigoDeposito = objNotaFiscal.CodigoEmpresa
                objNotaFiscal.EnderecoDeposito = objNotaFiscal.EnderecoEmpresa
                objNotaFiscal.CodigoDestino = objNotaFiscal.CodigoCliente
                objNotaFiscal.EnderecoDestino = objNotaFiscal.EnderecoCliente

                If objNotaFiscal.Itens.Count > 0 Then
                    objNotaFiscal.CodigoOperacao = objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D").FirstOrDefault.CodigoOperacao
                    objNotaFiscal.CodigoSubOperacao = objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D").FirstOrDefault.CodigoSubOperacao
                End If
                If String.IsNullOrWhiteSpace(objNotaFiscal.ChaveNFE) Then
                    objNotaFiscal.Eletronica = False
                Else
                    objNotaFiscal.Eletronica = True
                End If
                objNotaFiscal.Observacoes = Funcoes.EliminarCaracteresEspeciaisNF(txtObservacao.Text)
                If (FinanceiroNovo AndAlso objNotaFiscal.Pedido IsNot Nothing) Then hdCondicaoDePagamento.Value = objNotaFiscal.Pedido.CodigoCondicaoPagamento
                'Recupera os arquivos da Nota
                If ucFile.Parent.Visible Then
                    ucFile.Salvar(objNotaFiscal.Arquivos)
                End If

                Dim Sqls As New ArrayList
                Dim dsXmlProduto As DataSet = SessaoDsXML

                'Vamos gravar os dados de xml de para primeiro antes de gravar a nota, pq caso ocorra algum problema, esses
                'dados ficam salvos para uma posterior importação
                If Not SessaoDsXML Is Nothing Then

                    Dim produtosAdicionados As New List(Of Integer)

                    For Each item As NotaFiscalXItem In objNotaFiscal.Itens

                        'Produto de para'
                        Dim ProdutoDePara = New XmlProdutoXDePara(objNotaFiscal.Cliente.Codigo, 0, item.ProdutoXML)

                        Dim codigoProdutoXML = Funcoes.EliminarCaracteresEspeciais(item.ProdutoXML)
                        Dim nomeProdutoXML = Funcoes.EliminarCaracteresEspeciais(item.NomeProdutoXML)
                        'Se n tem cria'
                        If ProdutoDePara.CodigoProduto.Length = 0 Then

                            ProdutoDePara = New XmlProdutoXDePara()

                            ProdutoDePara.IUD = "I"
                            ProdutoDePara.CodigoCliente = objNotaFiscal.Cliente.Codigo
                            ProdutoDePara.EndCliente = 0

                            If codigoProdutoXML.ToString.Length > 30 Then
                                codigoProdutoXML = codigoProdutoXML.ToString.Substring(0, 30)
                            End If
                            If nomeProdutoXML.ToString.Length > 120 Then
                                nomeProdutoXML = nomeProdutoXML.ToString.Substring(0, 120)

                            End If
                            ProdutoDePara.CodigoProdutoXML = codigoProdutoXML
                            ProdutoDePara.NomeProdutoXML = nomeProdutoXML
                            ProdutoDePara.NCMProdutoXML = item.NCMProdutoXML
                            ProdutoDePara.UnidadeProdutoXML = item.UnidadeProdutoXML
                            ProdutoDePara.CodigoProduto = item.CodigoProduto
                            ProdutoDePara.ClienteConsolidado = 0 '?

                            If Not produtosAdicionados.Contains(ProdutoDePara.CodigoProduto) Then
                                ProdutoDePara.SalvarSql(Sqls)
                                produtosAdicionados.Add(ProdutoDePara.CodigoProduto)
                            End If

                        ElseIf ProdutoDePara.NCMProdutoXML.Length = 0 Then

                            ProdutoDePara.IUD = "U"
                            ProdutoDePara.CodigoCliente = objNotaFiscal.Cliente.Codigo
                            ProdutoDePara.EndCliente = 0

                            If codigoProdutoXML.ToString.Length > 30 Then
                                codigoProdutoXML = codigoProdutoXML.ToString.Substring(0, 30)
                            End If
                            If nomeProdutoXML.ToString.Length > 120 Then
                                nomeProdutoXML = nomeProdutoXML.ToString.Substring(0, 120)
                            End If
                            ProdutoDePara.CodigoProdutoXML = codigoProdutoXML
                            ProdutoDePara.NomeProdutoXML = nomeProdutoXML
                            ProdutoDePara.NCMProdutoXML = item.NCMProdutoXML
                            ProdutoDePara.UnidadeProdutoXML = item.UnidadeProdutoXML
                            ProdutoDePara.CodigoProduto = item.CodigoProduto
                            ProdutoDePara.ClienteConsolidado = 0 '?

                            If Not produtosAdicionados.Contains(ProdutoDePara.CodigoProduto) Then
                                ProdutoDePara.SalvarSql(Sqls)
                                produtosAdicionados.Add(ProdutoDePara.CodigoProduto)
                            End If

                        End If

                    Next

                End If

                Dim OPAdicionadas As New List(Of String)

                For Each item As [Lib].Negocio.NotaFiscalXItem In objNotaFiscal.Itens

                    Dim HistoricoOP = New ClientesXHistoricoOperacoes(objNotaFiscal.Cliente.Codigo, objNotaFiscal.Cliente.CodigoEndereco, item.Produto.Codigo, item.CodigoOperacao, item.CodigoSubOperacao, ddlTipoDeDocumento.SelectedValue)

                    'Se n tem cria'
                    If HistoricoOP.Operacao_Id = 0 Then

                        HistoricoOP = New ClientesXHistoricoOperacoes()

                        HistoricoOP.IUD = "I"

                        HistoricoOP.Cliente_Id = objNotaFiscal.CodigoCliente
                        HistoricoOP.EndCliente_Id = objNotaFiscal.EnderecoCliente
                        HistoricoOP.Produto_Id = item.Produto.Codigo
                        HistoricoOP.Operacao_Id = item.CodigoOperacao
                        HistoricoOP.SubOperacao_Id = item.CodigoSubOperacao
                        HistoricoOP.TipoDocumento = ddlTipoDeDocumento.SelectedValue

                        If Not OPAdicionadas.Contains((HistoricoOP.Produto_Id & HistoricoOP.Operacao_Id & HistoricoOP.SubOperacao_Id)) Then
                            HistoricoOP.SalvarSql(Sqls)
                            OPAdicionadas.Add((HistoricoOP.Produto_Id & HistoricoOP.Operacao_Id & HistoricoOP.SubOperacao_Id))
                        End If

                    End If

                Next

                If Banco.GravaBanco(Sqls) = False Then
                    MsgBox(Me.Page, "Erro ao gravar Informações de histórico!")
                    Exit Sub
                End If

                If objNotaFiscal.Empresa.Empresa.ControlaDataMovimentoNFG AndAlso objNotaFiscal.Movimento <> DateTime.Today Then
                    MsgBox(Me.Page, "A Data de Movimento não pode ser Posterior a Data de Hoje.")
                    Exit Sub
                End If

                If ValidarCampos() Then

                    Sqls = New ArrayList

                    'CRIA O PEDIDO COM BASE NA NOTA
                    'RECUPERA O SQL DE GRAVACAO DO PEDIDO
                    'RECUPERA O SQL DE GRAVACAO DA NOTA

                    Dim objPedido As New Pedido(objNotaFiscal, hdCondicaoDePagamento.Value)

                    'NUMERADOR DOS PEDIDOS
                    Dim SqlN As String = "exec sp_Numerador '" & objPedido.CodigoEmpresa & "'," & objPedido.EnderecoEmpresa & "," & [Lib].Negocio.eTiposNumerador.Pedido & ""
                    Dim dsN As New DataSet
                    dsN = Banco.ConsultaDataSet(SqlN, "Numerador")

                    Dim CodigoNumerador As Integer = 0
                    If dsN IsNot Nothing AndAlso dsN.Tables(0).Rows.Count > 0 Then
                        CodigoNumerador = dsN.Tables(0).Rows(0).Item(0)
                    End If

                    If Not CodigoNumerador > 0 Then
                        MsgBox(Me.Page, "Numerador de Pedidos não cadastrado!")
                        Exit Sub
                    End If

                    objPedido.Codigo = CodigoNumerador

                    If (Left(objNotaFiscal.CodigoEmpresa, 8) = "53267147" OrElse Left(objNotaFiscal.CodigoEmpresa, 8) = "62747840" OrElse Left(objNotaFiscal.CodigoEmpresa, 8) = "62780383" OrElse Left(objNotaFiscal.CodigoEmpresa, 8) = "63358210") Then
                        objPedido.FreteCIFFOB = eTiposFrete.CIF
                    End If

                    objNotaFiscal.CodigoPedido = CodigoNumerador
                    objNotaFiscal.Pedido = objPedido
                    objNotaFiscal.CIFFOB = objPedido.FreteCIFFOB

                    'If objPedido.Empresa.Empresa.ObrigaNavio Then
                    '    objPedido.InvoiceNavio = CInt(txtNaviosXInvoice.Text)
                    'End If

                    If FinanceiroNovo Then
                        For Each row In objNotaFiscal.Titulos
                            If Not row.CodigoProvisao = eProvisao.Baixa Then
                                row.CodigoUnidadeDeNegocio = objNotaFiscal.CodigoUnidadeDeNegocio
                                row.CodigoEmpresa = objNotaFiscal.CodigoEmpresa
                                row.EnderecoEmpresa = objNotaFiscal.EnderecoEmpresa
                                row.CodigoPedido = objNotaFiscal.CodigoPedido
                                row.Movimento = objNotaFiscal.Movimento
                                row.Historico = "PAGTO " & Trim(objNotaFiscal.TipoDeDocumento.Historico) & " " & Trim(objNotaFiscal.Codigo) & "-" & Trim(objNotaFiscal.Serie) & " - " & Trim(objNotaFiscal.Cliente.Nome) & " - " & Trim(row.Valores.EncargoValorDocumento.Descricao)
                                row.Observacoes = objNotaFiscal.Observacoes
                            End If
                        Next
                    Else
                        For Each row In objNotaFiscal.VencimentosNota
                            If Not row.CodigoProvisao = 1 Then
                                If Not objNotaFiscal.SubOperacao.Financeiro Then row.IUD = "D"

                                row.CodigoUnidadeDeNegocio = objNotaFiscal.CodigoUnidadeDeNegocio
                                row.CodigoEmpresa = objNotaFiscal.CodigoEmpresa
                                row.EnderecoEmpresa = objNotaFiscal.EnderecoEmpresa

                                row.CodigoEmpresaPedido = objNotaFiscal.CodigoEmpresa
                                row.EndEmpresaPedido = objNotaFiscal.EnderecoEmpresa
                                row.CodigoPedido = objNotaFiscal.CodigoPedido

                                row.CodigoCliente = objNotaFiscal.CodigoCliente
                                row.EndCliente = objNotaFiscal.EnderecoCliente

                                row.CodigoDestinatario = objNotaFiscal.CodigoCliente
                                row.EndDestinatario = objNotaFiscal.EnderecoCliente
                                row.NomeDoDestinatario = ""

                                row.Movimento = objNotaFiscal.Movimento

                                row.Historico = "PAGTO " & Trim(objNotaFiscal.TipoDeDocumento.Historico) & " " & Trim(objNotaFiscal.Codigo) & "-" & Trim(objNotaFiscal.Serie) & " - " & Trim(objNotaFiscal.Cliente.Nome) & " - " & Trim(row.Carteira.Descricao)
                                row.Observacoes = objNotaFiscal.Observacoes
                            End If
                        Next
                    End If
                    'Arquivos associados a NF
                    'If ucFile.Parent.Visible Then
                    '    For Each arq As [Lib].Negocio.Arquivo In objNotaFiscal.Arquivos
                    '        arq.CodigoEmpresa = objNotaFiscal.CodigoEmpresa
                    '        arq.EnderecoEmpresa = objNotaFiscal.EnderecoEmpresa
                    '        arq.CodigoCliente = objNotaFiscal.CodigoCliente
                    '        arq.EnderecoEmpresa = objNotaFiscal.EnderecoCliente
                    '        arq.CodigoNota = objNotaFiscal.Codigo
                    '        arq.Serie = objNotaFiscal.Serie
                    '        arq.CodigoPedido = objNotaFiscal.CodigoPedido
                    '        arq.SalvarSql(Sqls)
                    '    Next
                    'End If

                    'Setar IUD = "X" para as notas referenciais que não tenham o IUD configurado já que seria setado do item da nota que está como IUD= "I" causando erro de duplicidade na inserção.
                    If objNotaFiscal.CodigoTipoDeDocumento = eTipoDeDocumento.RPA AndAlso
                        objNotaFiscal.IUD = "U" AndAlso
                        objNotaFiscal.NotasReferenciais IsNot Nothing AndAlso
                        objNotaFiscal.NotasReferenciais.Count > 0 Then
                        For Each ref In objNotaFiscal.NotasReferenciais
                            If String.IsNullOrEmpty(ref.IUD) Then
                                ref.IUD = "X"
                            End If
                        Next
                    End If

                    SessaoSalvaNotaFiscal()
                    objPedido.SalvarSql(Sqls)
                    objNotaFiscal.SalvarSql(Sqls)

                    If Banco.GravaBanco(Sqls) Then
                        objNotaFiscalOriginal = objNotaFiscal
                        SessaoSalvaNotaFiscalOriginal()
                        MostrarMensagemFinal()
                        'chkReaproveitar.Checked = False
                        If chkEspelho.Checked Then lnkEspelho_Click(Nothing, Nothing)
                        LimparCampos()
                    Else
                        MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                    End If
                Else
                    MsgBox(Me.Page, mensagemErro)
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para incluir registro!")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível salvar a nota." & ex.Message)
        End Try
    End Sub

    Protected Sub lnkAtualizar_Click(sender As Object, e As EventArgs) Handles lnkAtualizar.Click
        Try
            Dim IUD_NFORI As String = ""
            Dim IUD_PDORI As String = ""

            If Not Funcoes.VerificaPermissao("NotasFiscaisGerais", "ALTERAR") Then
                MsgBox(Me.Page, "Usuário sem permissão para atualizar registro!")
                Exit Sub
            End If

            SessaoRecuperaNotaFiscalOriginal()
            SessaoRecuperaNotaFiscal()

            '************************************************************************************************************************************
            '******************************************** Ataliza dados Principais **************************************************************
            '************************************************************************************************************************************
            objNotaFiscal.Codigo = IIf(String.IsNullOrWhiteSpace(txtNumeroNota.Text), 0, txtNumeroNota.Text)
            objNotaFiscal.Serie = txtSerie.Text.ToUpper
            objNotaFiscal.NotaProdutor = IIf(String.IsNullOrWhiteSpace(txtNumeroNota.Text), 0, txtNumeroNota.Text)
            objNotaFiscal.SerieNotaProdutor = txtSerie.Text.ToUpper
            objNotaFiscal.CarregandoNota = True
            objNotaFiscal.DataNota = CDate(txtDataNota.Text)
            objNotaFiscal.Movimento = CDate(txtMovimento.Text)
            objNotaFiscal.CodigoTipoDeDocumento = IIf(ddlTipoDeDocumento.SelectedIndex = 0, 0, ddlTipoDeDocumento.SelectedValue)
            objNotaFiscal.CodigoSituacao = IIf(ddlSituacao.SelectedIndex = 0, 0, ddlSituacao.SelectedValue)
            objNotaFiscal.CodigoSafra = ddlSafra.SelectedValue
            objNotaFiscal.ChaveNFE = txtChaveNFe.Text.Replace(".", "")
            objNotaFiscal.NFG = True

            If objNotaFiscal.Itens.Count > 0 Then
                objNotaFiscal.CodigoOperacao = objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D").FirstOrDefault.CodigoOperacao
                objNotaFiscal.CodigoSubOperacao = objNotaFiscal.Itens.Where(Function(s) s.IUD <> "D").FirstOrDefault.CodigoSubOperacao
            End If

            If String.IsNullOrWhiteSpace(objNotaFiscal.ChaveNFE) Then
                objNotaFiscal.Eletronica = False
            Else
                objNotaFiscal.Eletronica = True
            End If

            objNotaFiscal.Observacoes = Funcoes.EliminarCaracteresEspeciaisNF(txtObservacao.Text)

            'Conferencia de Documento Fiscal
            If chkConferencia.Visible Then
                objNotaFiscal.Conferencia = chkConferencia.Checked
                objNotaFiscal.UsuarioConferencia = UsuarioServidor.NomeUsuario
                objNotaFiscal.UsuarioConferenciaData = DateTime.Now
            End If

            'Arquivo ****************************************************************************************************************************
            '************************************************************************************************************************************
            If ucFile.Parent.Visible Then
                ucFile.Salvar(objNotaFiscal.Arquivos)
            End If

            'Validar Campos *********************************************************************************************************************
            '************************************************************************************************************************************
            If Not ValidarCampos() Then
                MsgBox(Me.Page, mensagemErro)
                Exit Sub
            End If

            If objNotaFiscal.NossaEmissao AndAlso objNotaFiscal.SubOperacao.EntradaSaida = eEntradaSaida.Saida And objNotaFiscal.SubOperacao.Classe <> eClassesOperacoes.SERVICOS Then
                MsgBox(Me.Page, "Notas de saída são somente para visualização!")
                Exit Sub
            End If

            'Atualizacao ************************************************************************************************************************
            '************************************************************************************************************************************
            Dim Sqls As New ArrayList

            'Mudou a chave da Nota **************************************************************************************************************
            '************************************************************************************************************************************
            If objNotaFiscalOriginal.CodigoEmpresa <> objNotaFiscal.CodigoEmpresa _
            OrElse objNotaFiscalOriginal.EnderecoEmpresa <> objNotaFiscal.EnderecoEmpresa _
            OrElse objNotaFiscalOriginal.CodigoCliente <> objNotaFiscal.CodigoCliente _
            OrElse objNotaFiscalOriginal.EnderecoCliente <> objNotaFiscal.EnderecoCliente _
            OrElse objNotaFiscalOriginal.EntradaSaida <> objNotaFiscal.EntradaSaida _
            OrElse objNotaFiscalOriginal.Codigo <> objNotaFiscal.Codigo _
            OrElse objNotaFiscalOriginal.Serie <> objNotaFiscal.Serie Then

                'Deleta a Nota Fiscal Original
                objNotaFiscalOriginal.IUD = "D"

                'Deleta o vinculo da Nota com o Titulo
                For Each tit In objNotaFiscalOriginal.VencimentosNota
                    Sqls.Add(objNotaFiscalOriginal.VencimentosNota.NotaxTituloSql(tit.Codigo, objNotaFiscalOriginal.IUD))
                Next
                'Carrega os sqls de deleção da nota após o script de deleção dos títulos por causa da FK
                objNotaFiscalOriginal.SalvarSql(Sqls)

                'Muda O marcador da nota para Inclusao
                objNotaFiscal.IUD = "I" 'NÃO ESTAVA SALVANDO A NOTA CASO ALTERAÇÃO MUDASSE A CHAVE PORQUE CONSULTA  IUD=U - FURLAN - 24/03/2014
                IUD_NFORI = "D" 'Marcador para identificar que a NF original vai ser Deletada
            End If
            '************************************************************************************************************************************
            '************************************************************************************************************************************
            'Cria um Novo Pedido Baseado na Classe da Nota Fiscal
            Dim PedNovo As New Pedido(objNotaFiscal, hdCondicaoDePagamento.Value)

            'Mudou a Empresa ********************************************************************************************************************
            '************************************************************************************************************************************
            If objNotaFiscalOriginal.CodigoEmpresa <> objNotaFiscal.CodigoEmpresa _
            OrElse objNotaFiscalOriginal.EnderecoEmpresa <> objNotaFiscal.EnderecoEmpresa Then
                'Deleta o Pedido Original
                objNotaFiscalOriginal.Pedido.IUD = "D"
                objNotaFiscalOriginal.Pedido.SalvarSql(Sqls)
                IUD_PDORI = "D" 'Marcador para identificar que o Pedido original vai ser Deletado

                'Adquire um numero novo para o Pedido **********************************************
                '***********************************************************************************
                Dim SqlN As String = "exec sp_Numerador '" & objNotaFiscal.CodigoEmpresa & "'," & objNotaFiscal.EnderecoEmpresa & "," & [Lib].Negocio.eTiposNumerador.Pedido & ""
                Dim dsN As New DataSet
                dsN = Banco.ConsultaDataSet(SqlN, "Numerador")
                Dim CodigoNumerador As Integer = 0
                If dsN IsNot Nothing AndAlso dsN.Tables(0).Rows.Count > 0 Then
                    CodigoNumerador = dsN.Tables(0).Rows(0).Item(0)
                End If

                If Not CodigoNumerador > 0 Then
                    MsgBox(Me.Page, "Numerador não cadastrado!")
                    Exit Sub
                End If
                PedNovo.Codigo = CodigoNumerador

                'If PedNovo.Empresa.Empresa.ObrigaNavio Then
                '    PedNovo.InvoiceNavio = CInt(txtNaviosXInvoice.Text)
                'End If
                '***********************************************************************************
                'Limpar a classe Vencimentos do novo pedido a nota criara os vencimentos
                PedNovo.Vencimentos.Clear()
            Else
                'Se nao mudou a empresa conserva o mesmo numero de Pedido
                PedNovo.Codigo = objNotaFiscalOriginal.CodigoPedido

                If ddlCondicaoDePagamento.SelectedIndex > 0 Then
                    Dim cond As New CondicaoPagamento(ddlCondicaoDePagamento.SelectedValue)
                    If cond.Parcelas = grdCondicoes.Rows.Count Then
                        PedNovo.CodigoCondicaoPagamento = ddlCondicaoDePagamento.SelectedValue
                    Else
                        PedNovo.CodigoCondicaoPagamento = objNotaFiscalOriginal.Pedido.CodigoCondicaoPagamento
                    End If
                Else
                    PedNovo.CodigoCondicaoPagamento = objNotaFiscalOriginal.Pedido.CodigoCondicaoPagamento
                End If
            End If

            'Atualiza o codigo do Pedido na Nota que esta sendo atualizada
            objNotaFiscal.CodigoPedido = PedNovo.Codigo
            objNotaFiscal.Pedido = PedNovo


            SessaoSalvaNotaFiscal()

            'Verifica se o Pedido Original vai ser deletado 
            If IUD_PDORI = "D" Then
                'Insere o novo Pedido com as alteracoes
                PedNovo.SalvarSql(Sqls)
            Else
                'Altera o Pedido 
                PedNovo.IUD = "U"

                For Each row In PedNovo.Itens
                    If objNotaFiscalOriginal.Pedido.Itens.Where(Function(s) s.CodigoProduto = row.CodigoProduto).Count > 0 Then
                        row.IUD = "U"
                    Else
                        row.IUD = "I"
                    End If
                Next

                objNotaFiscalOriginal.Pedido.IUD = ""
                For Each row In objNotaFiscalOriginal.Pedido.Itens
                    If objNotaFiscal.Pedido.Itens.Where(Function(s) s.CodigoProduto = row.CodigoProduto).Count = 0 Then
                        row.IUD = "D"
                    End If
                Next

                objNotaFiscalOriginal.Pedido.Parcelas.ForEach(Function(s)
                                                                  s.IUD = "D"
                                                                  Return True
                                                              End Function)

                'Vai retornar os deletes dos produtos q existiam no original e nao existem mais no atual e vai deletar as parcelas 
                objNotaFiscalOriginal.Pedido.SalvarSql(Sqls)
                'Vai retornar os sqls para Alteracao do Pedido
                PedNovo.SalvarSql(Sqls)
            End If


            If objNotaFiscalOriginal.VencimentosNota IsNot Nothing AndAlso objNotaFiscalOriginal.VencimentosNota.Count > 0 Then
                For Each tit In objNotaFiscalOriginal.VencimentosNota
                    If objNotaFiscal.VencimentosNota.Where(Function(s) s.Codigo = tit.Codigo).Count = 0 Then
                        tit.IUD = "D"
                        Sqls.Add(objNotaFiscal.VencimentosNota.NotaxTituloSql(tit.Codigo, "D"))
                    End If
                Next
                objNotaFiscalOriginal.VencimentosNota.SalvarSQL(Sqls)

                For Each tit In objNotaFiscal.VencimentosNota.Where(Function(s) s.IUD <> "D")
                    'se o IUD_PDORI = "D" é por o pedido foi mudado de empresa entao a numeracao dos antigos titulos serao descartados e adquiridos novos na inclusa

                    tit.CodigoEmpresaPedido = objNotaFiscal.CodigoEmpresa
                    tit.EndEmpresaPedido = objNotaFiscal.EnderecoEmpresa
                    tit.CodigoPedido = objNotaFiscal.CodigoPedido

                    If objNotaFiscalOriginal.VencimentosNota.Where(Function(s) s.Codigo = tit.Codigo).Count = 0 OrElse IUD_PDORI = "D" OrElse IUD_NFORI = "D" Then
                        tit.IUD = "I"
                    Else
                        tit.IUD = "U"
                    End If
                Next
            End If

            objNotaFiscal.CodigoDeposito = objNotaFiscal.CodigoEmpresa
            objNotaFiscal.EnderecoDeposito = objNotaFiscal.EnderecoEmpresa
            objNotaFiscal.CodigoDestino = objNotaFiscal.CodigoCliente
            objNotaFiscal.EnderecoDestino = objNotaFiscal.EnderecoCliente

            'retorno os sqls da nota Fiscal
            objNotaFiscal.SalvarSql(Sqls)

            'Passei essa parte para depois pois estava tentando gravar uma NotaXTitulo sem ter a Nota Fiscal Gravada - Furlan - 12/09/2016
            If FinanceiroNovo Then
                For Each row In objNotaFiscal.Titulos
                    If Not row.CodigoProvisao = eProvisao.Baixa Then
                        If String.IsNullOrEmpty(row.IUD) Then row.IUD = "U"
                        row.CodigoUnidadeDeNegocio = objNotaFiscal.CodigoUnidadeDeNegocio
                        row.CodigoEmpresa = objNotaFiscal.CodigoEmpresa
                        row.EnderecoEmpresa = objNotaFiscal.EnderecoEmpresa

                        row.CodigoPedido = objNotaFiscal.CodigoPedido
                        row.CodigoContaContabilCliFor = objNotaFiscal.SubOperacao.CodigoGrupoContas
                        row.CodigoCliFor = objNotaFiscal.CodigoCliente
                        row.EnderecoCliFor = objNotaFiscal.EnderecoCliente

                        row.Movimento = objNotaFiscal.Movimento
                        row.Historico = "PAGTO " & Trim(objNotaFiscal.TipoDeDocumento.Historico) & " " & Trim(objNotaFiscal.Codigo) & "-" & Trim(objNotaFiscal.Serie) & " - " & Trim(objNotaFiscal.Cliente.Nome) & " - " & Trim(row.Valores.EncargoValorDocumento.Descricao)

                        row.Observacoes = objNotaFiscal.Observacoes

                        'Vinculo de nota fiscal e título
                        row.NotaTitulo.IUD = row.IUD
                        row.NotaTitulo.NotaFiscal = objNotaFiscal
                    End If
                Next
            Else
                For Each row In objNotaFiscal.VencimentosNota
                    If Not row.CodigoProvisao = 1 Then
                        If Not objNotaFiscal.SubOperacao.Financeiro Then
                            row.IUD = "D"
                        Else
                            If row.IUD Is Nothing OrElse row.IUD <> "D" Then row.IUD = "I"
                        End If

                        row.CodigoUnidadeDeNegocio = objNotaFiscal.CodigoUnidadeDeNegocio
                        row.CodigoEmpresa = objNotaFiscal.CodigoEmpresa
                        row.EnderecoEmpresa = objNotaFiscal.EnderecoEmpresa

                        row.CodigoEmpresaPedido = objNotaFiscal.CodigoEmpresa
                        row.EndEmpresaPedido = objNotaFiscal.EnderecoEmpresa
                        row.CodigoPedido = objNotaFiscal.CodigoPedido

                        row.CodigoCliente = objNotaFiscal.CodigoCliente
                        row.EndCliente = objNotaFiscal.EnderecoCliente

                        row.CodigoDestinatario = objNotaFiscal.CodigoCliente
                        row.EndDestinatario = objNotaFiscal.EnderecoCliente
                        row.NomeDoDestinatario = ""

                        row.Movimento = objNotaFiscal.Movimento
                        row.Historico = "PAGTO " & Trim(objNotaFiscal.TipoDeDocumento.Historico) & " " & Trim(objNotaFiscal.Codigo) & "-" & Trim(objNotaFiscal.Serie) & " - " & Trim(objNotaFiscal.Cliente.Nome) & " - " & Trim(row.Carteira.Descricao)
                        row.Observacoes = objNotaFiscal.Observacoes
                    ElseIf IUD_NFORI = "D" Then
                        'Vincula a Nota aos títulos quando alterada a chave e possui titulos baixados.
                        Sqls.Add(objNotaFiscal.VencimentosNota.NotaxTituloSql(row.Codigo, objNotaFiscal.IUD))
                    End If
                Next
            End If

            If FinanceiroNovo AndAlso objNotaFiscal.Pedido IsNot Nothing AndAlso Not objNotaFiscal.Pedido.Bloquear() Then
                MsgBox(Me.Page, "O pedido " & objNotaFiscal.Pedido.Codigo & " foi bloqueado por outro usuário, por favor recarregue o registro!")
                Exit Sub
            End If

            If FinanceiroNovo AndAlso Not String.IsNullOrWhiteSpace(objNotaFiscal.Empresa.Empresa.Servidor) AndAlso UsuarioServidor.NomeServidor <> objNotaFiscal.Empresa.Empresa.Servidor Then
                If FinanceiroNovo AndAlso PedNovo IsNot Nothing Then getSqlException(Sqls, PedNovo)
                Dim db As New AcessaBanco(2, objNotaFiscal.Empresa.Empresa.Servidor)
                If db.GravaBanco(Sqls) Then
                    objNotaFiscalOriginal = objNotaFiscal
                    SessaoSalvaNotaFiscalOriginal()
                    MostrarMensagemFinal()
                    chkReaproveitar.Checked = False
                    If chkEspelho.Checked Then lnkEspelho_Click(Nothing, Nothing)
                    LimparCampos()
                Else
                    MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                End If
            Else
                If Banco.GravaBanco(Sqls) Then
                    MostrarMensagemFinal()
                    If chkEspelho.Checked Then lnkEspelho_Click(Nothing, Nothing)
                    LimparCampos()
                Else
                    MsgBox(Me.Page, "Erro ao Alterar Nota: " & Funcoes.EliminarCaracteresEspeciais(RTrim(HttpContext.Current.Session("ssMessage").ToString)) & ". Entre em contato com o Suporte de TI")
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível atualizar a nota." & ex.Message)
        End Try
    End Sub

    Protected Sub lnkExcluir_Click(sender As Object, e As EventArgs) Handles lnkExcluir.Click
        Try
            If Funcoes.VerificaPermissao("NotasFiscaisGerais", "EXCLUIR") Then
                SessaoRecuperaNotaFiscalOriginal()
                If objNotaFiscalOriginal Is Nothing Then
                    MsgBox(Me.Page, "Antes de excluir, faça uma consulta e selecione uma das notas fiscais listadas")
                    Exit Sub
                End If

                Dim sqls As New ArrayList
                objNotaFiscalOriginal.NFG = True

                'Removido pois não faz sentido em Notas Gerais - Furlan - 17/09/2021
                'objNotaFiscalOriginal.SubOperacao.EntradaSaida = eEntradaSaida.Saida And objNotaFiscalOriginal.SubOperacao.Classe <> eClassesOperacoes.SERVICOS Then
                '                    MsgBox(Me.Page, "Não é permitido Excluir uma nota de Saida")
                '                    Exit Sub
                '                ElseIf 

                If objNotaFiscalOriginal.Empresa.Empresa.ConferenciaNFE AndAlso objNotaFiscalOriginal.Conferencia Then
                    MsgBox(Me.Page, "Não é permitido Excluir Documento conferido pelo setor fiscal")
                    Exit Sub
                ElseIf objNotaFiscalOriginal.SubOperacao.EntradaSaida = eEntradaSaida.Saida And objNotaFiscalOriginal.SubOperacao.Classe = eClassesOperacoes.SERVICOS Then
                    objNotaFiscalOriginal.IUD = "C"
                    objNotaFiscalOriginal.SalvarSql(sqls)

                    objNotaFiscalOriginal.Pedido.IUD = "U"
                    objNotaFiscalOriginal.Pedido.CodigoSituacao = 2
                    objNotaFiscalOriginal.Pedido.SalvarSql(sqls)
                Else
                    objNotaFiscalOriginal.IUD = "D"
                    objNotaFiscalOriginal.SalvarSql(sqls)

                    objNotaFiscalOriginal.Pedido.IUD = "D"
                    objNotaFiscalOriginal.Pedido.SalvarSql(sqls)
                End If

                If objNotaFiscalOriginal.TemNotaTroca AndAlso objNotaFiscalOriginal.CodigoTipoDeDocumento = eTipoDeDocumento.NotaDeProdutor Then
                    MsgBox(Me.Page, "Nota Fiscal não pode ser excluída pois está vinculada a Nota Fiscal " & objNotaFiscalOriginal.NotaTrocaDestino.Codigo & " de " & objNotaFiscalOriginal.NotaTrocaDestino.Movimento.ToString("dd/MM/yyyy") & ".", eTitulo.Info)
                    Exit Sub
                End If

                If objNotaFiscalOriginal.VencimentosNota IsNot Nothing AndAlso objNotaFiscalOriginal.VencimentosNota.Count > 0 Then
                    For Each tit In objNotaFiscalOriginal.VencimentosNota

                        If tit.Agrupado = eAgrupamentoFinanceiro.Agrupado Then
                            MsgBox(Me.Page, "Nota Fiscal com Título em Agrupamento não pode ser excluído.", eTitulo.Info)
                            Exit Sub
                        ElseIf tit.CodigoProvisao = 1 Then
                            MsgBox(Me.Page, "Nota Fiscal com Título baixado não pode ser excluído.", eTitulo.Info)
                            Exit Sub
                        End If

                    Next
                End If

                'Arquivos da Nota
                If ucFile.Parent.Visible Then
                    For Each arq As [Lib].Negocio.Arquivo In objNotaFiscalOriginal.Arquivos
                        arq.IUD = "D"
                        arq.SalvarSql(sqls)
                    Next
                End If
                'Deleta os titulos conforme a Nota.
                If FinanceiroNovo Then
                    For Each row In objNotaFiscalOriginal.Titulos
                        If Not row.CodigoProvisao = 1 Then
                            row.IUD = objNotaFiscalOriginal.IUD
                            row.NotaTitulo.IUD = objNotaFiscalOriginal.IUD
                            row.SalvarSql(sqls)
                        End If
                    Next
                End If

                If Banco.GravaBanco(sqls) Then
                    LimparCampos()
                    MsgBox(Me.Page, "Nota excluida com Sucesso.", eTitulo.Sucess)
                Else
                    MsgBox(Me.Page, "Erro ao Excluir a Nota: " & Funcoes.EliminarCaracteresEspeciais(RTrim(HttpContext.Current.Session("ssMessage").ToString)) & ". Entre em contato com o Suporte de TI")
                End If
            Else
                MsgBox(Me.Page, "Usuario sem permissao para excluir Registro")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível excluir a nota")
        End Try
    End Sub

    Protected Sub lnkLimpar_Click(sender As Object, e As EventArgs) Handles lnkLimpar.Click
        Try
            LimparCampos()
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível limpar a nota")
        End Try
    End Sub

    Private Function ObterEncargosXmlNota(ByVal dsXml As DataSet) As HashSet(Of String)

        Dim encargosXml As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

        Dim detTable As DataTable = Nothing
        If dsXml Is Nothing OrElse Not dsXml.Tables.Contains("det") Then Return encargosXml

        detTable = dsXml.Tables("det")

        For Each detRow As DataRow In detTable.Rows

            Dim impostoRelations = detRow.Table.ChildRelations.Cast(Of DataRelation)().
                               Where(Function(rel) rel.ChildTable.TableName = "imposto").ToList()

            For Each impostoRelation As DataRelation In impostoRelations

                Dim impostoRows As DataRow() = detRow.GetChildRows(impostoRelation)

                For Each impostoRow As DataRow In impostoRows

                    Dim impostoChildRelations = impostoRow.Table.ChildRelations.Cast(Of DataRelation)().
                                            Where(Function(rel) rel.ChildTable.TableName = "ICMS" OrElse
                                                                rel.ChildTable.TableName = "IPI" OrElse
                                                                rel.ChildTable.TableName = "PIS" OrElse
                                                                rel.ChildTable.TableName = "COFINS" OrElse
                                                                rel.ChildTable.TableName = "ISSQN" OrElse
                                                                rel.ChildTable.TableName = "II").ToList()

                    For Each impostoChildRelation As DataRelation In impostoChildRelations

                        Dim impostoChildRows As DataRow() = impostoRow.GetChildRows(impostoChildRelation)

                        For Each impostoChildRow As DataRow In impostoChildRows

                            Select Case impostoChildRow.Table.TableName

                                Case "ICMS"

                                    'ICMS00 -> ICMS
                                    Dim icms00 = impostoChildRow.GetChildRows(impostoChildRow.Table.ChildRelations.Cast(Of DataRelation)().
                                                                         FirstOrDefault(Function(rel) rel.ChildTable.TableName = "ICMS00"))
                                    If icms00 IsNot Nothing AndAlso icms00.Length > 0 Then
                                        If ExisteValorMaiorQueZero(icms00, "vICMS") Then encargosXml.Add("ICMS")
                                    End If

                                    'ICMS60 -> ICMS-ST (no teu código você usa assim)
                                    Dim icms60 = impostoChildRow.GetChildRows(impostoChildRow.Table.ChildRelations.Cast(Of DataRelation)().
                                                                         FirstOrDefault(Function(rel) rel.ChildTable.TableName = "ICMS60"))
                                    If icms60 IsNot Nothing AndAlso icms60.Length > 0 Then
                                        If ExisteValorMaiorQueZero(icms60, "vICMS") Then encargosXml.Add("ICMS-ST")
                                    End If

                                    'ICMSUFDEST -> diferencial
                                    Dim icmsUfDest = impostoChildRow.GetChildRows(impostoChildRow.Table.ChildRelations.Cast(Of DataRelation)().
                                                                             FirstOrDefault(Function(rel) rel.ChildTable.TableName = "ICMSUFDEST"))
                                    If icmsUfDest IsNot Nothing AndAlso icmsUfDest.Length > 0 Then
                                        If ExisteValorMaiorQueZero(icmsUfDest, "vICMSUFDEST") Then encargosXml.Add("ICMS DIFERENCIAL")
                                        If ExisteValorMaiorQueZero(icmsUfDest, "vICMSDeson") Then encargosXml.Add("ICMS DESONERADO")
                                    End If

                                Case "IPI"

                                    Dim ipi = impostoChildRow.GetChildRows(impostoChildRow.Table.ChildRelations.Cast(Of DataRelation)().
                                                                      FirstOrDefault(Function(rel) rel.ChildTable.TableName = "IPINT"))
                                    If ipi IsNot Nothing AndAlso ipi.Length > 0 Then
                                        If ExisteValorMaiorQueZero(ipi, "vIPI") Then encargosXml.Add("IPI")
                                    End If

                                Case "PIS"

                                    Dim pis = impostoChildRow.GetChildRows(impostoChildRow.Table.ChildRelations.Cast(Of DataRelation)().
                                                                      FirstOrDefault(Function(rel) rel.ChildTable.TableName = "PISAliq"))
                                    If pis IsNot Nothing AndAlso pis.Length > 0 Then
                                        If ExisteValorMaiorQueZero(pis, "vPIS") Then encargosXml.Add("PIS")
                                    End If

                                    'PISST
                                    Dim pisst = impostoChildRow.GetChildRows(impostoChildRow.Table.ChildRelations.Cast(Of DataRelation)().
                                                                        FirstOrDefault(Function(rel) rel.ChildTable.TableName = "PISST"))
                                    If pisst IsNot Nothing AndAlso pisst.Length > 0 Then
                                        If ExisteValorMaiorQueZero(pisst, "vPISST") Then encargosXml.Add("PISST")
                                    End If

                                Case "COFINS"

                                    Dim cofins = impostoChildRow.GetChildRows(impostoChildRow.Table.ChildRelations.Cast(Of DataRelation)().
                                                                         FirstOrDefault(Function(rel) rel.ChildTable.TableName = "COFINSAliq"))
                                    If cofins IsNot Nothing AndAlso cofins.Length > 0 Then
                                        If ExisteValorMaiorQueZero(cofins, "vCOFINS") Then encargosXml.Add("COFINS")
                                    End If

                                    'COFINSST
                                    Dim cofinsst = impostoChildRow.GetChildRows(impostoChildRow.Table.ChildRelations.Cast(Of DataRelation)().
                                                                           FirstOrDefault(Function(rel) rel.ChildTable.TableName = "COFINSST"))
                                    If cofinsst IsNot Nothing AndAlso cofinsst.Length > 0 Then
                                        If ExisteValorMaiorQueZero(cofinsst, "vCOFINSST") Then encargosXml.Add("COFINSST")
                                    End If

                                Case "ISSQN"

                                    Dim issqn = impostoChildRow.GetChildRows(impostoChildRow.Table.ChildRelations.Cast(Of DataRelation)().
                                                                         FirstOrDefault(Function(rel) rel.ChildTable.TableName = "ISSQN"))
                                    If issqn IsNot Nothing AndAlso issqn.Length > 0 Then
                                        If ExisteValorMaiorQueZero(issqn, "vISSQN") Then encargosXml.Add("ISSQN")
                                    End If

                                Case "II"

                                    'se o teu dsXml tiver II
                                    If impostoChildRow.Table.Columns.Contains("vII") Then
                                        If Not IsDBNull(impostoChildRow("vII")) Then
                                            If ToDecimal(impostoChildRow("vII")) > 0D Then encargosXml.Add("II")
                                        End If
                                    End If

                            End Select

                        Next
                    Next
                Next
            Next
        Next

        Return encargosXml

    End Function

    Private Function ExisteValorMaiorQueZero(ByVal rows As DataRow(), ByVal col As String) As Boolean
        For Each r As DataRow In rows
            If r.Table.Columns.Contains(col) AndAlso Not IsDBNull(r(col)) Then
                If ToDecimal(r(col)) > 0D Then Return True
            End If
        Next
        Return False
    End Function

    Private Function ToDecimal(ByVal v As Object) As Decimal
        If v Is Nothing OrElse v Is DBNull.Value Then Return 0D
        Dim s As String = v.ToString().Trim().Replace(".", ",")
        Dim d As Decimal = 0D
        Decimal.TryParse(s, Globalization.NumberStyles.Any, Globalization.CultureInfo.GetCultureInfo("pt-BR"), d)
        Return d
    End Function

    Private Function ObterEncargosNota(ByVal ObjNotaFiscal As [Lib].Negocio.NotaFiscal) As HashSet(Of String)

        Dim encargosNF As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

        For Each e In ObjNotaFiscal.Itens.SelectMany(Function(s) s.Encargos)
            If Not String.IsNullOrWhiteSpace(e.Codigo) Then
                encargosNF.Add(e.Codigo.Trim())
            End If
        Next

        Return encargosNF

    End Function

    Private Function ValidarEncargosXmlVsNota(ByVal ObjNotaFiscal As [Lib].Negocio.NotaFiscal, ByVal dsXml As DataSet) As String

        'Não validaremos nesse momento essas questões de impostos no XML - 2026-01-16
        Return String.Empty

        Dim xmlSet As HashSet(Of String) = ObterEncargosXmlNota(dsXml)
        Dim nfSet As HashSet(Of String) = ObterEncargosNota(ObjNotaFiscal)

        Dim faltantes As New List(Of String)

        For Each enc In xmlSet
            If Not nfSet.Contains(enc) Then
                faltantes.Add(enc)
            End If
        Next

        If faltantes.Count > 0 Then
            Return "Não é possível salvar a nota. O XML contém impostos que não existem como encargos na NF: " & String.Join(", ", faltantes)
        End If

        Return String.Empty

    End Function

    Protected Sub lnkEspelho_Click(sender As Object, e As EventArgs) Handles lnkEspelho.Click
        Try
            Dim espelho As New [Lib].Negocio.NotaFiscalEspelho

            SessaoRecuperaNotaFiscal()

            If objNotaFiscal Is Nothing Then
                MsgBox(Me.Page, "Salve a nota para poder visualizar o espelho")
            Else
                espelho.ExibirEspelho(Me, objNotaFiscal)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível emitir o espelho da nota")
        End Try
    End Sub

    Protected Sub lnkRecontabilizar_Click(sender As Object, e As EventArgs) Handles lnkRecontabilizar.Click
        Try
            SessaoRecuperaNotaFiscalOriginal()

            If objNotaFiscalOriginal Is Nothing Then
                MsgBox(Me.Page, "Carregue a nota antes de Recontabilizar")
                Exit Sub
            End If

            If Funcoes.VerificaPermissao("RECONTABILIZAR", "ALTERAR") Then
                If Not Funcoes.VerificaAcesso(objNotaFiscalOriginal.CodigoEmpresa, objNotaFiscalOriginal.EnderecoEmpresa, objNotaFiscalOriginal.Movimento, "CONTABIL") Then
                    MsgBox(Me.Page, "Movimento Contábil já Fechado para esta data")
                    Exit Sub
                End If

                Dim sqls As New ArrayList

                'Fiz um ajuste em locas ref. recontabilização, por hora para evitar duplicação de contabilização por não saber se os registros
                'estão no lote 9 ou 10 ou 21, vamos deletar todos e recontabilizar no correto como está abaixo. - furlan 13-03-2014
                objNotaFiscalOriginal.Razao.ExcluiContabilizacaoNotaSQL(sqls, 9)
                objNotaFiscalOriginal.Razao.ExcluiContabilizacaoNotaSQL(sqls, 10)
                objNotaFiscalOriginal.Razao.ExcluiContabilizacaoNotaSQL(sqls, 11)
                objNotaFiscalOriginal.Razao.ExcluiContabilizacaoNotaSQL(sqls, 21)

                If objNotaFiscalOriginal.CodigoTipoDeDocumento = CInt(eTipoDeDocumento.CTRC) Or
                objNotaFiscalOriginal.CodigoTipoDeDocumento = CInt(eTipoDeDocumento.ReciboDeFrete) Or
                objNotaFiscalOriginal.CodigoTipoDeDocumento = CInt(eTipoDeDocumento.CTRC_SEM_NF) Or
                objNotaFiscalOriginal.CodigoTipoDeDocumento = CInt(eTipoDeDocumento.CT_E_TOM) Or
                objNotaFiscalOriginal.CodigoTipoDeDocumento = CInt(eTipoDeDocumento.CT_E) Or
                objNotaFiscalOriginal.CodigoTipoDeDocumento = CInt(eTipoDeDocumento.CT_E_OUT) Or
                objNotaFiscalOriginal.CodigoTipoDeDocumento = CInt(eTipoDeDocumento.Estadia) Or
                objNotaFiscalOriginal.CodigoTipoDeDocumento = CInt(eTipoDeDocumento.ComplementoDeFrete) Or
                objNotaFiscalOriginal.CodigoTipoDeDocumento = CInt(eTipoDeDocumento.Anulacao) Then
                    objNotaFiscalOriginal.Razao.ContabilizarNotaSql(sqls, 21)
                Else
                    objNotaFiscalOriginal.Razao.ContabilizarNotaSql(sqls, 9)
                End If

                Dim Banco As New AcessaBanco

                If Banco.GravaBanco(sqls) Then
                    Dim objNotaAntes = New NotaFiscal(objNotaFiscalOriginal)
                    objNotaFiscalOriginal = New NotaFiscal(objNotaAntes)
                    SessaoSalvaNotaFiscalOriginal()

                    objNotaFiscalOriginal.LancamentosContabeis.CalcularSaldo()
                    gridRazao.DataSource = objNotaFiscalOriginal.LancamentosContabeis.OrderBy(Function(s) s.Sequencia)
                    gridRazao.DataBind()

                    MsgBox(Me.Page, "Nota Contabilizada com Sucesso.", eTitulo.Sucess)
                Else
                    MsgBox(Me.Page, "Erro durante o Processo de recontabilizacao")
                End If

            Else
                MsgBox(Me.Page, "Usuário sem permissão para recontabilizar nota fiscal")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível recontabilizar a nota")
        End Try
    End Sub

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "NotasFiscaisGerais")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub

#End Region

#Region "Xml NFe"

    Protected Sub PreencherNFeXML(ByVal objNotaFiscal As [Lib].Negocio.NotaFiscal, ByVal pNomeArquivo As String, ByVal pOrigem As Boolean)

        Dim DsXml As New DataSet

        Dim nome As String = pNomeArquivo
        If Not nome.ToLower().EndsWith(".xml") Then nome &= ".xml"

        DsXml.ReadXml(Server.MapPath(String.Format("~/Files/{0}", nome)))
        SessaoDsXML = DsXml
        objNotaFiscal.NossaEmissao = False
        objNotaFiscal.XMLImportado = True
        objNotaFiscal.NFG = True

        Dim xmlDoc As New XmlDocument
        xmlDoc.Load(Server.MapPath(String.Format("~/Files/{0}", nome)))

        SessaoXmlDoc = xmlDoc

        'Chave NFe ou CTe
        If xmlDoc.GetElementsByTagName("protNFe").Count = 1 Then
            objNotaFiscal.ChaveNFE = xmlDoc.GetElementsByTagName("protNFe").GetNode("protNFe").ChildNodes.GetNode("infProt").ChildNodes.GetNode("chNFe").InnerText
        Else
            objNotaFiscal.ChaveNFE = xmlDoc.GetElementsByTagName("protCTe").GetNode("protCTe").ChildNodes.GetNode("infProt").ChildNodes.GetNode("chCTe").InnerText
        End If

        If xmlDoc.GetElementsByTagName("protNFe").Count = 1 Then
            'Codigo NFe
            objNotaFiscal.Codigo = xmlDoc.GetElementsByTagName("nNF").GetNode("nNF").InnerText
            objNotaFiscal.CodigoTipoDeDocumento = eTipoDeDocumento.Nota
        Else

            'Codigo CTe
            objNotaFiscal.Codigo = xmlDoc.GetElementsByTagName("nCT").GetNode("nCT").InnerText

            Dim chaveNFEOrigem As String = ""

            If xmlDoc.GetElementsByTagName("infNFe").Count > 0 Then

                Dim infNFeElement As XmlElement = CType(xmlDoc.GetElementsByTagName("infNFe")(0), XmlElement)

                If infNFeElement.GetElementsByTagName("chave").Count > 0 Then

                    chaveNFEOrigem = infNFeElement.GetElementsByTagName("chave")(0).InnerText

                    objNotaFiscal.CodigoTipoDeDocumento = eTipoDeDocumento.CT_E 'CTRC = 2
                    objNotaFiscal.TipoDeDocumento.Codigo = eTipoDeDocumento.CT_E 'CTRC = 2

                    Dim nf As New [Lib].Negocio.NotaFiscal()

                    nf.CodigoEmpresa = objNotaFiscal.CodigoEmpresa
                    nf.EnderecoEmpresa = objNotaFiscal.EnderecoEmpresa
                    'nf.CodigoCliente = Row.Cells(1).Text.Split("-")(0)
                    'nf.EnderecoCliente = Row.Cells(1).Text.Split("-")(1)
                    'nf.EntradaSaida = IIf(Row.Cells(5).Text.Split("-")(0) = "E", eEntradaSaida.Entrada, eEntradaSaida.Saida)
                    'nf.Codigo = Row.Cells(5).Text.Split("-")(1)
                    'nf.Serie = Row.Cells(5).Text.Split("-")(2)

                    nf.CarregarNotaComChaveXML(chaveNFEOrigem)

                    If nf.Codigo > 0 Then
                        nf = New [Lib].Negocio.NotaFiscal(nf)
                        If Not (objNotaFiscal.NotasTrocaOrigem IsNot Nothing AndAlso objNotaFiscal.NotasTrocaOrigem.Count > 0) Then
                            objNotaFiscal.NotasTrocaOrigem = New List(Of [Lib].Negocio.NotaFiscal)
                        End If

                        If objNotaFiscal.NotasTrocaOrigem.Count = 0 Or objNotaFiscal.NotasTrocaOrigem.Where(Function(x) x.Empresa.Codigo = nf.Empresa.Codigo And x.Cliente.Codigo = nf.Cliente.Codigo And x.Codigo = nf.Codigo).Count = 0 Then
                            objNotaFiscal.NotasTrocaOrigem.Add(nf)
                            CarregarNotasDeOrigem()
                        End If

                    End If

                End If

            End If

            If chaveNFEOrigem.Length > 0 Then
                objNotaFiscal.CodigoTipoDeDocumento = eTipoDeDocumento.CT_E
                objNotaFiscal.TipoDeDocumento.Codigo = eTipoDeDocumento.CT_E
            Else
                objNotaFiscal.CodigoTipoDeDocumento = eTipoDeDocumento.CT_E_TOM 'CTRC_SEM_NF = 8
                objNotaFiscal.TipoDeDocumento.Codigo = eTipoDeDocumento.CT_E_TOM 'CTRC_SEM_NF = 8
            End If

        End If

        'Série
        objNotaFiscal.Serie = xmlDoc.GetElementsByTagName("serie").GetNode("serie").InnerText

        If xmlDoc.GetElementsByTagName("protNFe").Count = 1 Then
            'Emissão
            If xmlDoc.GetElementsByTagName("dEmi").GetNode("dEmi") IsNot Nothing Then
                objNotaFiscal.DataNota = xmlDoc.GetElementsByTagName("dEmi").GetNode("dEmi").InnerText
            Else
                Dim strData As String = xmlDoc.GetElementsByTagName("dhEmi").GetNode("dhEmi").InnerText
                objNotaFiscal.DataNota = strData.Remove(strData.Length - 6)
            End If
        Else
            'Emissão
            If xmlDoc.GetElementsByTagName("dhEmi").GetNode("dhEmi") IsNot Nothing Then
                objNotaFiscal.DataNota = xmlDoc.GetElementsByTagName("dhEmi").GetNode("dhEmi").InnerText
            Else
                Dim strData As String = xmlDoc.GetElementsByTagName("dhEmi").GetNode("dhEmi").InnerText
                objNotaFiscal.DataNota = strData.Remove(strData.Length - 6)
            End If
        End If

        Dim Empresa As New [Lib].Negocio.ListCliente

        If xmlDoc.GetElementsByTagName("protNFe").Count = 1 Then
            'Empresa 
            Dim codigo As String = xmlDoc.GetElementsByTagName("dest").GetNode("dest").ChildNodes.GetNode("CNPJ").InnerText
            'Dim cep As String = xmlDoc.GetElementsByTagName("dest").GetNode("dest").ChildNodes.GetNode("enderDest").ChildNodes.GetNode("CEP").InnerText
            'Dim ie As String = xmlDoc.GetElementsByTagName("dest").GetNode("dest").ChildNodes.GetNode("IE").InnerText

            Empresa = Empresa.FindCliente(codigo)
            If Empresa.Count = 1 Then
                objNotaFiscal.CodigoEmpresa = Empresa(0).Codigo
                objNotaFiscal.EnderecoEmpresa = Empresa(0).CodigoEndereco
                objNotaFiscal.CodigoUnidadeDeNegocio = Empresa(0).CodigoUnidadeDeNegocio
                objNotaFiscal.Empresa = Empresa(0)
            End If


            'Fornecedor
            If objNotaFiscal.CodigoCliente.Length = 11 Then
                codigo = xmlDoc.GetElementsByTagName("emit").GetNode("emit").ChildNodes.GetNode("CPF").InnerText
            Else
                codigo = xmlDoc.GetElementsByTagName("emit").GetNode("emit").ChildNodes.GetNode("CNPJ").InnerText
            End If

            Dim Fornecedor As New [Lib].Negocio.ListCliente
            Fornecedor = Empresa.FindCliente(codigo)

            If Fornecedor.Count = 1 Then
                objNotaFiscal.CodigoCliente = Fornecedor(0).Codigo
                objNotaFiscal.EnderecoCliente = Fornecedor(0).CodigoEndereco
            ElseIf Fornecedor.Count = 0 Then
                MsgBox(Me.Page, "Cliente não encontrado!", eTitulo.Info, False)
                txtNomeCliente.Text = String.Empty
            Else
                ucConsultaClientes.Limpar()
                ucConsultaClientes.CarregarCNPJ(codigo)
                Popup.ConsultaDeClientes(Me.Page, "objCliente" & HID.Value, "txtNome")
            End If

        Else

            'Fornecedor 
            Dim codigo As String = xmlDoc.GetElementsByTagName("emit").GetNode("emit").ChildNodes.GetNode("CNPJ").InnerText

            'Fornecedor
            If objNotaFiscal.CodigoCliente.Length = 11 Then
                codigo = xmlDoc.GetElementsByTagName("emit").GetNode("emit").ChildNodes.GetNode("CPF").InnerText
            Else
                codigo = xmlDoc.GetElementsByTagName("emit").GetNode("emit").ChildNodes.GetNode("CNPJ").InnerText
            End If

            Dim Fornecedor As New [Lib].Negocio.ListCliente
            Fornecedor = Empresa.FindCliente(codigo)

            If Fornecedor.Count = 1 Then
                objNotaFiscal.CodigoCliente = Fornecedor(0).Codigo
                objNotaFiscal.EnderecoCliente = Fornecedor(0).CodigoEndereco
            ElseIf Fornecedor.Count = 0 Then
                MsgBox(Me.Page, "Cliente não encontrado!", eTitulo.Info, False)
                txtNomeCliente.Text = String.Empty
            Else
                ucConsultaClientes.Limpar()
                ucConsultaClientes.CarregarCNPJ(codigo)
                Popup.ConsultaDeClientes(Me.Page, "objCliente" & HID.Value, "txtNome")
            End If

        End If


        'VerificarProdutos se a nota existe

        Dim strSQL As String

        strSQL = "  SELECT 1 " & vbCrLf &
                 "  FROM NotasFiscais NF" & vbCrLf &
                 "  INNER JOIN NotasFiscaisxItens NFxI" & vbCrLf &
                 "      ON NF.Empresa_Id        = NFxI.Empresa_Id" & vbCrLf &
                 "      AND NF.EndEmpresa_Id    = NFxI.EndEmpresa_Id" & vbCrLf &
                 "      AND NF.Cliente_Id       = NFxI.Cliente_Id" & vbCrLf &
                 "      AND NF.EndCliente_Id    = NFxI.EndCliente_Id" & vbCrLf &
                 "      AND NF.EntradaSaida_Id  = NFxI.EntradaSaida_Id" & vbCrLf &
                 "      AND NF.Serie_Id         = NFxI.Serie_Id" & vbCrLf &
                 "      AND NF.Nota_Id          = NFxI.Nota_Id" & vbCrLf &
                 "  WHERE NF.Empresa_Id         = '" & objNotaFiscal.Empresa.Codigo & "'" & vbCrLf &
                 "      AND NF.Cliente_Id       = '" & objNotaFiscal.Cliente.Codigo & "'" & vbCrLf &
                 "      AND NF.Nota_Id          = " & objNotaFiscal.Codigo & vbCrLf &
                 "      AND NF.Serie_Id         = '" & objNotaFiscal.Serie & "'"

        Dim ds As DataSet = Banco.ConsultaDataSet(strSQL, "NotasFiscaisGerais")

        If Not ds Is Nothing AndAlso ds.Tables(0).Rows.Count() > 0 Then
            MsgBox(Me.Page, "XML já possui uma nota fiscal cadastrada, não é possivel importar!", eTitulo.Info, False)
            Exit Sub
        End If

        'Situação
        objNotaFiscal.CodigoSituacao = eSituacao.Normal

        'Safra
        objNotaFiscal.CodigoSafra = "NENHUMA"

        'Xml Obsoleto
        'objNotaFiscal.XML = xmlDoc.InnerXml

        'bloqueia os campos
        txtChaveNFe.Enabled = False
        txtNumeroNota.Enabled = False
        txtSerie.Enabled = False
        txtDataNota.Enabled = False


        'Manifesta a NFe quando a origem for ucFile
        Dim msgResult As String = String.Empty
        If pOrigem Then
            If xmlDoc.GetElementsByTagName("protNFe").Count = 1 Then
                If Not DocumentoEletronico.ManifestoNFe(objNotaFiscal, eTipoManifesto.ConfirmacaoDaOperacao, msgResult) Then
                    MsgBox(Me.Page, msgResult, eTitulo.Info, False)
                    'LimparCampos()
                    'Exit Sub
                Else
                    MsgBox(Me.Page, msgResult, eTitulo.Info, False)
                End If
            End If
        End If

        'A verificação tem que ser antes de adicionar os produtos
        ' Verifica se a tabela "ICMSTot" existe no DataSet
        If DsXml.Tables.Contains("ICMSTot") Then
            Dim icmsTotTable As DataTable = DsXml.Tables("ICMSTot")

            ' Verifica se há pelo menos uma linha na tabela "ICMSTot"
            If icmsTotTable.Rows.Count > 0 Then
                Dim icmsTotRow As DataRow = icmsTotTable.Rows(0)

                ' Extrai os valores das tags vProd e vNF
                Dim vProd As String = If(icmsTotRow.Table.Columns.Contains("vProd") AndAlso Not IsDBNull(icmsTotRow("vProd")), icmsTotRow("vProd").ToString(), "")
                Dim vNF As String = If(icmsTotRow.Table.Columns.Contains("vNF") AndAlso Not IsDBNull(icmsTotRow("vNF")), icmsTotRow("vNF").ToString(), "")

                ' Exemplo: Atribuir os valores a outras variáveis
                Dim valorProduto As Decimal = Decimal.Parse(vProd.Replace(".", ",")) ' Converte para Decimal
                Dim valorNotaFiscal As Decimal = Decimal.Parse(vNF.Replace(".", ",")) ' Converte para Decimal

                objNotaFiscal.DiferencaValorNFXProdutoXML = valorNotaFiscal - valorProduto

            End If
        End If

        If Not DsXml.Tables("transp") Is Nothing Then
            '*** Transporte
            If CInt(DsXml.Tables("transp").Rows(0)("modfrete")) = 0 OrElse CInt(DsXml.Tables("transp").Rows(0)("modfrete")) = 3 Then
                objNotaFiscal.CIFFOB = eTiposFrete.CIF
            ElseIf CInt(DsXml.Tables("transp").Rows(0)("modfrete")) = 1 OrElse CInt(DsXml.Tables("transp").Rows(0)("modfrete")) = 4 Then
                objNotaFiscal.CIFFOB = eTiposFrete.FOB
            ElseIf CInt(DsXml.Tables("transp").Rows(0)("modfrete")) = 2 Then
                objNotaFiscal.CIFFOB = eTiposFrete.TER
            ElseIf CInt(DsXml.Tables("transp").Rows(0)("modfrete")) = 9 Then
                objNotaFiscal.CIFFOB = eTiposFrete.NEN
            End If
        End If

        If Not DsXml.Tables("transporta") Is Nothing Then

            Dim temTransporte As Boolean = False

            If DsXml.Tables("transporta").Columns.Contains("CNPJ") Then
                temTransporte = True
                objNotaFiscal.CodigoTransportador = DsXml.Tables("transporta").Rows(0)("CNPJ").ToString()
            ElseIf DsXml.Tables("transporta").Columns.Contains("CPF") Then
                temTransporte = True
                objNotaFiscal.CodigoTransportador = DsXml.Tables("transporta").Rows(0)("CPF").ToString()
            End If

            If temTransporte Then

                Dim cTransportador As String = String.Empty
                If DsXml.Tables("transporta").Columns.Contains("CNPJ") Then
                    cTransportador = DsXml.Tables("transporta").Rows(0)("CNPJ").ToString()
                    objNotaFiscal.CodigoTransportador = DsXml.Tables("transporta").Rows(0)("CNPJ").ToString()
                ElseIf DsXml.Tables("transporta").Columns.Contains("CPF") Then
                    cTransportador = DsXml.Tables("transporta").Rows(0)("CPF").ToString()
                    objNotaFiscal.CodigoTransportador = DsXml.Tables("transporta").Rows(0)("CPF").ToString()
                End If

                Dim TranspCodigo As New ListCliente("", Nothing, cTransportador)

                Dim TranpCidadeTemp As String = String.Empty
                If DsXml.Tables("transporta").Columns.Contains("xMun") Then
                    TranpCidadeTemp = DsXml.Tables("transporta").Rows(0)("xMun")
                End If

                Dim TranpEstadoTemp As String = String.Empty
                If DsXml.Tables("transporta").Columns.Contains("UF") Then
                    TranpEstadoTemp = DsXml.Tables("transporta").Rows(0)("UF")
                End If

                Dim TranspEnderecoTemp As String = String.Empty
                If DsXml.Tables("transporta").Columns.Contains("xEnder") Then
                    TranspEnderecoTemp = DsXml.Tables("transporta").Rows(0)("xEnder")
                End If

                Dim TranspEnd As Integer = 0
                Dim temTransportador As Boolean = False

                For Each c In TranspCodigo
                    If TranpEstadoTemp.Length > 0 AndAlso TranpCidadeTemp.Length > 0 AndAlso TranspEnderecoTemp.Length > 0 Then
                        'AndAlso c.Endereco.Contains(TranspEnderecoTemp) - Removido - Furlan - 11/01/2024
                        If c.CodigoEstado = TranpEstadoTemp AndAlso c.Cidade.ToUpper() = TranpCidadeTemp.ToUpper() Then
                            TranspEnd = c.CodigoEndereco

                            temTransportador = True
                        End If
                    End If
                Next

                If temTransportador Then
                    objNotaFiscal.EnderecoTransportador = TranspEnd
                End If

            End If
        End If

        If Not DsXml.Tables("veicTransp") Is Nothing AndAlso Not DsXml.Tables("veicTransp").Rows(0)("placa") Is Nothing Then

            objNotaFiscal.PlacaTransportador = DsXml.Tables("veicTransp").Rows(0)("placa")
            If objNotaFiscal.PlacaDetalhes Is Nothing OrElse objNotaFiscal.PlacaDetalhes.Placa01.Length = 0 Then
                objNotaFiscal.PlacaDetalhes = New [Lib].Negocio.Placa
            End If

            objNotaFiscal.PlacaDetalhes.Placa01 = DsXml.Tables("veicTransp").Rows(0)("placa")

            If DsXml.Tables("veicTransp").Columns.Contains("UF") Then
                objNotaFiscal.PlacaDetalhes.EstadoPlaca01 = DsXml.Tables("veicTransp").Rows(0)("UF")
            Else
                objNotaFiscal.PlacaDetalhes.EstadoPlaca01 = ""
            End If

            If Not DsXml.Tables("reboque") Is Nothing Then
                objNotaFiscal.PlacaDetalhes.Placa02 = DsXml.Tables("reboque").Rows(0)("placa")

                If DsXml.Tables("reboque").Columns.Contains("UF") Then
                    objNotaFiscal.PlacaDetalhes.EstadoPlaca02 = DsXml.Tables("reboque").Rows(0)("UF")
                Else
                    objNotaFiscal.PlacaDetalhes.EstadoPlaca02 = ""
                End If

            End If

            objNotaFiscal.PlacaDetalhes.Motorista = New [Lib].Negocio.Cliente
            objNotaFiscal.PlacaDetalhes.Motorista.Codigo = ""
            objNotaFiscal.PlacaDetalhes.Motorista.CodigoEndereco = 0

        End If

        objNotaFiscal.Itens.Clear()

        If chkImportarProdutoUnico.Checked Or chkAGruparNCM.Checked Then
            If chkImportarProdutoUnico.Checked Then
                ImportarProdutoUnico(objNotaFiscal, xmlDoc)
            Else
                AgruparNCM(objNotaFiscal, xmlDoc)
            End If
        Else
            If chkInformarDados.Checked Then
                InformarDadosProdutos(xmlDoc)
                Exit Sub
            End If
            ImportarProdutoXML(objNotaFiscal, xmlDoc, True)
        End If

        CarregarFormComAClasse(objNotaFiscal, True)
        AdicionarProdutos(False, True)

    End Sub

    Private Sub ImportarProdutoXML(ByVal objNotaFiscal As NotaFiscal, ByVal xmlDoc As XmlDocument, ByVal bProcurarProduto As Boolean)

        Dim nodelist As XmlNodeList = xmlDoc.DocumentElement.ChildNodes

        For Each noPai In nodelist 'Le os nós principais da NFe

            If noPai.Name = "NFe" Then

                For Each noFilho In noPai 'Lê os Nós secundários

                    For Each noNeto In noFilho 'Lê as Tags da NFe\CTe

                        If noNeto.Name = "det" Then 'Verifica os detalhes dos produtos

                            Dim pFatorConversao = 1
                            Dim objProdutoNotaItem As New NotaFiscalXItem(objNotaFiscal)
                            Dim operacaoNossa As Integer = 0
                            Dim SubOperacaoNossa As Integer = 0

                            objProdutoNotaItem.DescricaoProdutoXML = Funcoes.EliminarCaracteresEspeciais(If(noNeto.GetElementsByTagName("xProd").Count > 0, noNeto.GetElementsByTagName("xProd").Item(0).InnerText, ""))
                            objProdutoNotaItem.NomeProdutoXML = Funcoes.EliminarCaracteresEspeciais(If(noNeto.GetElementsByTagName("xProd").Count > 0, noNeto.GetElementsByTagName("xProd").Item(0).InnerText, ""))
                            objProdutoNotaItem.NCMProdutoXML = If(noNeto.GetElementsByTagName("NCM").Count > 0, noNeto.GetElementsByTagName("NCM").Item(0).InnerText, "")
                            objProdutoNotaItem.UnidadeProdutoXML = If(noNeto.GetElementsByTagName("uTrib").Count > 0, noNeto.GetElementsByTagName("uTrib").Item(0).InnerText, "")

                            objProdutoNotaItem.ProdutoXML = If(noNeto.GetElementsByTagName("cProd").Count > 0, noNeto.GetElementsByTagName("cProd").Item(0).InnerText, "")

                            Dim ProdutoDeParaXML As New XmlProdutoXDePara
                            Dim produtoNosso As New Produto

                            If bProcurarProduto Then

                                Dim listProdutos As New ListXmlProdutoXDePara(objProdutoNotaItem.NCMProdutoXML, objNotaFiscal.CodigoCliente, objNotaFiscal.EnderecoCliente, objProdutoNotaItem.ProdutoXML)

                                If listProdutos.Count() = 1 Then
                                    ProdutoDeParaXML = listProdutos.FirstOrDefault()
                                ElseIf listProdutos.Count() > 1 And listProdutos.Where(Function(x) x.CodigoCliente = objNotaFiscal.CodigoCliente And x.CodigoProdutoXML = objProdutoNotaItem.ProdutoXML).Count() > 0 Then
                                    ProdutoDeParaXML = listProdutos.Where(Function(x) x.CodigoCliente = objNotaFiscal.CodigoCliente And x.CodigoProdutoXML = objProdutoNotaItem.ProdutoXML).FirstOrDefault()
                                ElseIf listProdutos.Count() > 1 Then
                                    ProdutoDeParaXML = listProdutos.FirstOrDefault()
                                End If

                                'Se tiver'
                                If ProdutoDeParaXML.CodigoProduto.Length > 0 Then

                                    produtoNosso = New Produto(ProdutoDeParaXML.CodigoProduto)

                                    Dim historicoOP As ListClientesXHistoricoOperacoes = New ListClientesXHistoricoOperacoes("", 0, ProdutoDeParaXML.CodigoProduto, 0, 0, objNotaFiscal.TipoDeDocumento.Codigo)

                                    If historicoOP.Count > 0 Then

                                        If historicoOP.Where(Function(x) x.Cliente_Id = objNotaFiscal.CodigoCliente).Count() > 0 Then
                                            operacaoNossa = historicoOP.Where(Function(x) x.Cliente_Id = objNotaFiscal.CodigoCliente).FirstOrDefault().Operacao_Id
                                            SubOperacaoNossa = historicoOP.Where(Function(x) x.Cliente_Id = objNotaFiscal.CodigoCliente).FirstOrDefault().SubOperacao_Id
                                        Else
                                            operacaoNossa = historicoOP(0).Operacao_Id
                                            SubOperacaoNossa = historicoOP(0).SubOperacao_Id
                                        End If

                                    End If

                                End If

                            End If

                            objProdutoNotaItem.Sequencia = IIf(noNeto.Attributes("nItem").Value = "", 0, noNeto.Attributes("nItem").Value)
                            objProdutoNotaItem.CodigoProduto = ProdutoDeParaXML.CodigoProduto

                            If chkInformarDados.Checked Then

                                SessaoRecuperaProdutoNota()
                                operacaoNossa = objProdutoNota.CodigoOperacao
                                SubOperacaoNossa = objProdutoNota.CodigoSubOperacao
                                objProdutoNotaItem.CodigoProduto = objProdutoNota.CodigoProduto
                                objProdutoNotaItem.TemCentroDeCusto = objProdutoNota.TemCentroDeCusto
                                If objProdutoNotaItem.TemCentroDeCusto Then
                                    objProdutoNotaItem.CodigoProdutoCusto = objProdutoNota.CentroDeCustoInformado
                                End If

                            End If

                            objProdutoNotaItem.InfAdicionalProdutoXML = If(noNeto.GetElementsByTagName("infAdProd").Count > 0, noNeto.GetElementsByTagName("infAdProd").Item(0).InnerText, "")

                            objProdutoNotaItem.PesoFiscal = If(noNeto.GetElementsByTagName("qTrib").Count > 0, noNeto.GetElementsByTagName("qTrib").Item(0).InnerText.Replace(".", ","), "0").Replace(",,", ",")
                            objProdutoNotaItem.QuantidadeFiscal = If(noNeto.GetElementsByTagName("qTrib").Count > 0, noNeto.GetElementsByTagName("qTrib").Item(0).InnerText.Replace(".", ","), "0").Replace(",,", ",")

                            Dim sOP As New SubOperacao(operacaoNossa, SubOperacaoNossa)

                            If sOP.QuantidadeFisico Then
                                objProdutoNotaItem.QuantidadeFisica = objProdutoNotaItem.QuantidadeFiscal
                            End If

                            objProdutoNotaItem.PesoLiquido = If(noNeto.GetElementsByTagName("qTrib").Count > 0, noNeto.GetElementsByTagName("qTrib").Item(0).InnerText.Replace(".", ","), "0").Replace(",,", ",")
                            objProdutoNotaItem.PesoBruto = If(noNeto.GetElementsByTagName("qCom").Count > 0, noNeto.GetElementsByTagName("qCom").Item(0).InnerText.Replace(".", ","), "0").ToString.Replace(".", ",").Replace(",,", ",")

                            objProdutoNotaItem.Unitario = If(noNeto.GetElementsByTagName("vUnTrib").Count > 0, noNeto.GetElementsByTagName("vUnTrib").Item(0).InnerText.Replace(".", ","), "0").Replace(",,", ",")

                            Dim converterTons As String() = {"TO", "TON", "TN", "TNF", "BG"}

                            If converterTons.Contains(noNeto.GetElementsByTagName("uCom").Item(0).InnerText) Then
                                objProdutoNotaItem.Unitario = objProdutoNotaItem.Unitario / 1000
                                objProdutoNotaItem.PesoBruto = objProdutoNotaItem.PesoBruto * 1000
                                objProdutoNotaItem.PesoFiscal = objProdutoNotaItem.PesoFiscal * 1000
                                objProdutoNotaItem.PesoLiquido = objProdutoNotaItem.PesoLiquido * 1000
                                objProdutoNotaItem.QuantidadeFiscal = objProdutoNotaItem.QuantidadeFiscal * 1000
                                If sOP.QuantidadeFisico Then
                                    objProdutoNotaItem.QuantidadeFisica = objProdutoNotaItem.QuantidadeFiscal
                                End If
                            End If

                            objProdutoNotaItem.Volumes = 1 '?
                            objProdutoNotaItem.Numeracao = 1  '???
                            objProdutoNotaItem.PesoQuantidade = produtoNosso.PesoQuantidade

                            objProdutoNotaItem.ValorLiquido = If(noNeto.GetElementsByTagName("vProd").Count > 0, noNeto.GetElementsByTagName("vProd").Item(0).InnerText.Replace(".", ","), "0").Replace(",,", ",")
                            objProdutoNotaItem.ValorTotal = If(noNeto.GetElementsByTagName("vProd").Count > 0, noNeto.GetElementsByTagName("vProd").Item(0).InnerText.Replace(".", ","), "0").Replace(",,", ",")

                            objProdutoNotaItem.ValorFreteXML = If(noNeto.GetElementsByTagName("vFrete").Count > 0, noNeto.GetElementsByTagName("vFrete").Item(0).InnerText.Replace(".", ","), "0").Replace(",,", ",")
                            objProdutoNotaItem.ValorDescontoXML = If(noNeto.GetElementsByTagName("vDesc").Count > 0, noNeto.GetElementsByTagName("vDesc").Item(0).InnerText.Replace(".", ","), "0").Replace(",,", ",")

                            If operacaoNossa <> 0 Then

                                Dim operacaoHistorico = New Operacao(operacaoNossa)
                                Dim SubOperacaoHistorico = New SubOperacao(operacaoNossa, SubOperacaoNossa)

                                'VER COM O FURLAN'
                                objNotaFiscal.Operacao = operacaoHistorico
                                objNotaFiscal.SubOperacao = SubOperacaoHistorico

                                objNotaFiscal.CodigoOperacao = operacaoNossa
                                objNotaFiscal.CodigoSubOperacao = SubOperacaoNossa

                                objProdutoNotaItem.Operacao = operacaoHistorico
                                objProdutoNotaItem.SubOperacao = SubOperacaoHistorico

                                objProdutoNotaItem.CodigoOperacao = operacaoNossa
                                objProdutoNotaItem.CodigoSubOperacao = SubOperacaoNossa

                                objProdutoNotaItem.CarregandoEncargos = True
                                objProdutoNotaItem.Encargos = New [Lib].Negocio.ListNotaFiscalXItemXEncargo(objProdutoNotaItem, New List(Of eEtapaEncago) From {eEtapaEncago.Normal})
                                objProdutoNotaItem.CarregandoEncargos = False

                                If chkInformarDados.Checked Then
                                    objProdutoNotaItem.Encargos.EncProduto.CentroDeCusto = objProdutoNotaItem.CodigoProdutoCusto
                                End If

                            Else

                                objProdutoNotaItem.Encargos = New ListNotaFiscalXItemXEncargo()

                            End If

                            objNotaFiscal.Itens.Add(objProdutoNotaItem)

                            objNotaFiscal.CarregandoItens = True

                            For Each noBisneto In noNeto

                                'If noBisneto.Name = "imposto" Then 'Dados dos Impostos

                                '    For Each noTetra In noBisneto
                                '        If noTetra.Name = "ICMS" Then

                                '            Dim encargoICMS As NotaFiscalXItemXEncargo = New NotaFiscalXItemXEncargo

                                '            encargoICMS.Codigo = "ICMS"
                                '            encargoICMS.


                                '            objItem.orig = If(noTetra.GetElementsByTagName("orig").Count > 0, noTetra.GetElementsByTagName("orig").Item(0).InnerText, "")
                                '            objItem.CSTICMS = If(noTetra.GetElementsByTagName("CST").Count > 0, noTetra.GetElementsByTagName("CST").Item(0).InnerText, "")
                                '            objItem.modBCICMS = If(noTetra.GetElementsByTagName("modBC").Count > 0, noTetra.GetElementsByTagName("modBC").Item(0).InnerText, "")
                                '            objItem.pRedBCICMS = If(noTetra.GetElementsByTagName("pRedBC").Count > 0, noTetra.GetElementsByTagName("pRedBC").Item(0).InnerText.Replace(".", ","), "0").Replace(",,", ",")
                                '            objItem.vBCICMS = If(noTetra.GetElementsByTagName("vBC").Count > 0, noTetra.GetElementsByTagName("vBC").Item(0).InnerText.Replace(".", ","), "0").Replace(",,", ",")
                                '            objItem.pICMS = If(noTetra.GetElementsByTagName("pICMS").Count > 0, noTetra.GetElementsByTagName("pICMS").Item(0).InnerText.Replace(".", ","), "0").Replace(",,", ",")
                                '            objItem.vICMS = If(noTetra.GetElementsByTagName("vICMS").Count > 0, noTetra.GetElementsByTagName("vICMS").Item(0).InnerText.Replace(".", ","), "0").Replace(",,", ",")
                                '            objItem.modBCSTICMS = If(noTetra.GetElementsByTagName("modBCST").Count > 0, noTetra.GetElementsByTagName("modBCST").Item(0).InnerText, "")
                                '            objItem.pmVASTICMS = If(noTetra.GetElementsByTagName("pmVAST").Count > 0, noTetra.GetElementsByTagName("pmVAST").Item(0).InnerText.Replace(".", ","), "0").Replace(",,", ",")
                                '            objItem.pRedBCSTICMS = If(noTetra.GetElementsByTagName("pRedBCST").Count > 0, noTetra.GetElementsByTagName("pRedBCST").Item(0).InnerText.Replace(".", ","), "0").Replace(",,", ",")
                                '            objItem.vBCSTICMS = If(noTetra.GetElementsByTagName("vBCST").Count > 0, noTetra.GetElementsByTagName("vBCST").Item(0).InnerText.Replace(".", ","), "0").Replace(",,", ",")
                                '            objItem.pICMSST = If(noTetra.GetElementsByTagName("pICMSST").Count > 0, noTetra.GetElementsByTagName("pICMSST").Item(0).InnerText.Replace(".", ","), "0").Replace(",,", ",")
                                '            objItem.vICMSST = If(noTetra.GetElementsByTagName("vICMSST").Count > 0, noTetra.GetElementsByTagName("vICMSST").Item(0).InnerText.Replace(".", ","), "0").Replace(",,", ",")
                                '            objItem.vBCSTRetICMS = If(noTetra.GetElementsByTagName("vBCSTRet").Count > 0, noTetra.GetElementsByTagName("vBCSTRet").Item(0).InnerText.Replace(".", ","), "0").Replace(",,", ",")
                                '            objItem.vICMSSTRet = If(noTetra.GetElementsByTagName("vICMSSTRet").Count > 0, noTetra.GetElementsByTagName("vICMSSTRet").Item(0).InnerText.Replace(".", ","), "0").Replace(",,", ",")
                                '            objItem.vBCICMSSTDest = If(noTetra.GetElementsByTagName("vBCICMSSTDest").Count > 0, noTetra.GetElementsByTagName("vBCICMSSTDest").Item(0).InnerText.Replace(".", ","), "0").Replace(",,", ",")
                                '            objItem.vICMSSTDest = If(noTetra.GetElementsByTagName("vICMSSTDest").Count > 0, noTetra.GetElementsByTagName("vICMSSTDest").Item(0).InnerText.Replace(".", ","), "0").Replace(",,", ",")
                                '            objItem.motDesICMS = If(noTetra.GetElementsByTagName("motDesICMS").Count > 0, noTetra.GetElementsByTagName("motDesICMS").Item(0).InnerText, "")
                                '            objItem.pBCOpICMS = If(noTetra.GetElementsByTagName("pBCOp").Count > 0, noTetra.GetElementsByTagName("pBCOp").Item(0).InnerText.Replace(".", ","), "0").Replace(",,", ",")
                                '            objItem.UFSTICMS = If(noTetra.GetElementsByTagName("UFST").Count > 0, noTetra.GetElementsByTagName("UFST").Item(0).InnerText, "")
                                '            objItem.pCredSNICMS = If(noTetra.GetElementsByTagName("pCredSN").Count > 0, noTetra.GetElementsByTagName("pCredSN").Item(0).InnerText.Replace(".", ","), "0").Replace(",,", ",")
                                '            objItem.vCredICMSSN = If(noTetra.GetElementsByTagName("vCredICMSSN").Count > 0, noTetra.GetElementsByTagName("vCredICMSSN").Item(0).InnerText.Replace(".", ","), "0").Replace(",,", ",")
                                '            objItem.vDesonICMS = If(noTetra.GetElementsByTagName("vICMSDeson").Count > 0, noTetra.GetElementsByTagName("vICMSDeson").Item(0).InnerText.Replace(".", ","), "0").Replace(",,", ",")
                                '            objItem.vICMSOp = If(noTetra.GetElementsByTagName("vICMSOp").Count > 0, noTetra.GetElementsByTagName("vICMSOp").Item(0).InnerText.Replace(".", ","), "0").Replace(",,", ",")
                                '            objItem.vICMSDif = If(noTetra.GetElementsByTagName("vICMSDif").Count > 0, noTetra.GetElementsByTagName("vICMSDif").Item(0).InnerText.Replace(".", ","), "0").Replace(",,", ",")
                                '            objItem.CSOSN = If(noTetra.GetElementsByTagName("CSOSN").Count > 0, noTetra.GetElementsByTagName("CSOSN").Item(0).InnerText.Replace(".", ","), "0").Replace(",,", ",")
                                '            objItem.pDifICMS = If(noTetra.GetElementsByTagName("pDif").Count > 0, noTetra.GetElementsByTagName("pDif").Item(0).InnerText.Replace(".", ","), "0").Replace(",,", ",")

                                '            objItem.vICMSDeson = If(noTetra.GetElementsByTagName("vICMSDeson").Count > 0, noTetra.GetElementsByTagName("vICMSDeson").Item(0).InnerText.Replace(".", ","), "0").Replace(",,", ",")
                                '            objItem.pICMSDeson = If(noTetra.GetElementsByTagName("pICMSDeson").Count > 0, noTetra.GetElementsByTagName("pICMSDeson").Item(0).InnerText.Replace(".", ","), "0").Replace(",,", ",")
                                '            objItem.vBICMSDeson = If(noTetra.GetElementsByTagName("vBICMSDeson").Count > 0, noTetra.GetElementsByTagName("vBICMSDeson").Item(0).InnerText.Replace(".", ","), "0").Replace(",,", ",")

                                '            objItem.vBCFCPST = If(noTetra.GetElementsByTagName("vBCFCPST").Count > 0, noTetra.GetElementsByTagName("vBCFCPST").Item(0).InnerText.Replace(".", ","), "0").Replace(",,", ",")
                                '            objItem.pFCPST = If(noTetra.GetElementsByTagName("pFCPST").Count > 0, noTetra.GetElementsByTagName("pFCPST").Item(0).InnerText.Replace(".", ","), "0").Replace(",,", ",")
                                '            objItem.vFCPST = If(noTetra.GetElementsByTagName("vFCPST").Count > 0, noTetra.GetElementsByTagName("vFCPST").Item(0).InnerText.Replace(".", ","), "0").Replace(",,", ",")

                                '            If noBisneto.GetElementsByTagName("ICMSSN101").Count > 0 Or noBisneto.GetElementsByTagName("ICMSSN201").Count > 0 Then
                                '                objItem.vICMSSN = noTetra.GetElementsByTagName("vCredICMSSN").Item(0).InnerText.Replace(".", ",").Replace(",,", ",")
                                '                objItem.pICMSSN = noTetra.GetElementsByTagName("pCredSN").Item(0).InnerText.Replace(".", ",").Replace(",,", ",")
                                '                objItem.vBICMSSN = pValorToTalProdutos

                                '            End If



                                '        End If

                                '        If noTetra.Name = "PIS" Then

                                '            objItem.CSTPIS = If(noTetra.GetElementsByTagName("CST").Count > 0, noTetra.GetElementsByTagName("CST").Item(0).InnerText, "")
                                '            objItem.vBCPIS = If(noTetra.GetElementsByTagName("vBC").Count > 0, noTetra.GetElementsByTagName("vBC").Item(0).InnerText.Replace(".", ","), "0").Replace(",,", ",")
                                '            objItem.pPIS = If(noTetra.GetElementsByTagName("pPIS").Count > 0, noTetra.GetElementsByTagName("pPIS").Item(0).InnerText.Replace(".", ","), "0").Replace(",,", ",")
                                '            objItem.vPIS = If(noTetra.GetElementsByTagName("vPIS").Count > 0, noTetra.GetElementsByTagName("vPIS").Item(0).InnerText.Replace(".", ","), "0").Replace(",,", ",")
                                '            objItem.qBCProdPIS = If(noTetra.GetElementsByTagName("qBCProd").Count > 0, noTetra.GetElementsByTagName("qBCProd").Item(0).InnerText.Replace(".", ","), "0").Replace(",,", ",")
                                '            objItem.vAliqProdPIS = If(noTetra.GetElementsByTagName("vAliqProd").Count > 0, noTetra.GetElementsByTagName("vAliqProd").Item(0).InnerText.Replace(".", ","), "0").Replace(",,", ",")

                                '        End If

                                '        If noTetra.Name = "II" Then

                                '            objItem.vBCII = If(noTetra.GetElementsByTagName("vBC").Count > 0, noTetra.GetElementsByTagName("vBC").Item(0).InnerText.Replace(".", ","), "0").Replace(",,", ",")
                                '            objItem.vDespAduII = If(noTetra.GetElementsByTagName("vDespAdu").Count > 0, noTetra.GetElementsByTagName("vDespAdu").Item(0).InnerText.Replace(".", ","), "0").Replace(",,", ",")
                                '            objItem.vII = If(noTetra.GetElementsByTagName("vII").Count > 0, noTetra.GetElementsByTagName("vII").Item(0).InnerText.Replace(".", ","), "0").Replace(",,", ",")
                                '            objItem.vIOFII = If(noTetra.GetElementsByTagName("vIOF").Count > 0, noTetra.GetElementsByTagName("vIOF").Item(0).InnerText.Replace(".", ","), "0").Replace(",,", ",")

                                '        End If

                                '        If noTetra.Name = "COFINS" Then

                                '            objItem.CSTCOFINS = If(noTetra.GetElementsByTagName("CST").Count > 0, noTetra.GetElementsByTagName("CST").Item(0).InnerText, "")
                                '            objItem.vBCCOFINS = If(noTetra.GetElementsByTagName("vBC").Count > 0, noTetra.GetElementsByTagName("vBC").Item(0).InnerText.Replace(".", ","), "0").Replace(",,", ",")
                                '            objItem.pCOFINS = If(noTetra.GetElementsByTagName("pCOFINS").Count > 0, noTetra.GetElementsByTagName("pCOFINS").Item(0).InnerText.Replace(".", ","), "0").Replace(",,", ",")
                                '            objItem.vCOFINS = If(noTetra.GetElementsByTagName("vCOFINS").Count > 0, noTetra.GetElementsByTagName("vCOFINS").Item(0).InnerText.Replace(".", ","), "0").Replace(",,", ",")
                                '            objItem.qBCProdCOFINS = If(noTetra.GetElementsByTagName("qBCProd").Count > 0, noTetra.GetElementsByTagName("qBCProd").Item(0).InnerText.Replace(".", ","), "0").Replace(",,", ",")
                                '            objItem.vAliqProdCOFINS = If(noTetra.GetElementsByTagName("vAliqProd").Count > 0, noTetra.GetElementsByTagName("vAliqProd").Item(0).InnerText.Replace(".", ","), "0").Replace(",,", ",")

                                '        End If


                                '        If noTetra.Name = "PISST" Then

                                '            objItem.vBCPISST = If(noTetra.GetElementsByTagName("vBC").Count > 0, noTetra.GetElementsByTagName("vBC").Item(0).InnerText.Replace(".", ","), "0").Replace(",,", ",")
                                '            objItem.pPISST = If(noTetra.GetElementsByTagName("pPIS").Count > 0, noTetra.GetElementsByTagName("pPIS").Item(0).InnerText.Replace(".", ","), "0").Replace(",,", ",")
                                '            objItem.vPISST = If(noTetra.GetElementsByTagName("vPIS").Count > 0, noTetra.GetElementsByTagName("vPIS").Item(0).InnerText.Replace(".", ","), "0").Replace(",,", ",")
                                '            objItem.qBCProdPISST = If(noTetra.GetElementsByTagName("qBCProd").Count > 0, noTetra.GetElementsByTagName("qBCProd").Item(0).InnerText.Replace(".", ","), "0").Replace(",,", ",")
                                '            objItem.vAliqProdPISST = If(noTetra.GetElementsByTagName("vAliqProd").Count > 0, noTetra.GetElementsByTagName("vAliqProd").Item(0).InnerText.Replace(".", ","), "0").Replace(",,", ",")

                                '        End If

                                '        If noTetra.Name = "COFINSST" Then

                                '            objItem.vBCCOFINSST = If(noTetra.GetElementsByTagName("vBC").Count > 0, noTetra.GetElementsByTagName("vBC").Item(0).InnerText.Replace(".", ","), "0").Replace(",,", ",")
                                '            objItem.pCOFINSST = If(noTetra.GetElementsByTagName("pCOFINS").Count > 0, noTetra.GetElementsByTagName("pCOFINS").Item(0).InnerText.Replace(".", ","), "0").Replace(",,", ",")
                                '            objItem.vCOFINSST = If(noTetra.GetElementsByTagName("vCOFINS").Count > 0, noTetra.GetElementsByTagName("vCOFINS").Item(0).InnerText.Replace(".", ","), "0").Replace(",,", ",")
                                '            objItem.qBCProdCOFINSST = If(noTetra.GetElementsByTagName("qBCProd").Count > 0, noTetra.GetElementsByTagName("qBCProd").Item(0).InnerText.Replace(".", ","), "0").Replace(",,", ",")
                                '            objItem.vAliqProdCOFINSST = If(noTetra.GetElementsByTagName("vAliqProd").Count > 0, noTetra.GetElementsByTagName("vAliqProd").Item(0).InnerText.Replace(".", ","), "0").Replace(",,", ",")

                                '        End If

                                '        If noTetra.Name = "ISSQN" Then

                                '            objItem.vBCISS = If(noTetra.GetElementsByTagName("vBC").Count > 0, noTetra.GetElementsByTagName("vBC").Item(0).InnerText.Replace(".", ","), "0").Replace(",,", ",")
                                '            objItem.vAliqISS = If(noTetra.GetElementsByTagName("vAliq").Count > 0, noTetra.GetElementsByTagName("vAliq").Item(0).InnerText.Replace(".", ","), "0").Replace(",,", ",")
                                '            objItem.vISSQN = If(noTetra.GetElementsByTagName("vISSQN").Count > 0, noTetra.GetElementsByTagName("vISSQN").Item(0).InnerText.Replace(".", ","), "0").Replace(",,", ",")
                                '            objItem.cMunFGISS = If(noTetra.GetElementsByTagName("cMunFG").Count > 0, noTetra.GetElementsByTagName("cMunFG").Item(0).InnerText, "")
                                '            objItem.cListServISS = If(noTetra.GetElementsByTagName("cListServ").Count > 0, noTetra.GetElementsByTagName("cListServ").Item(0).InnerText, "")
                                '            objItem.vDeduccao_OpcISS = If(noTetra.GetElementsByTagName("vDeduccao_Opc").Count > 0, noTetra.GetElementsByTagName("vDeduccao_Opc").Item(0).InnerText.Replace(".", ","), "0").Replace(",,", ",")
                                '            objItem.vOutro_OpcISS = If(noTetra.GetElementsByTagName("vOutro_Opc").Count > 0, noTetra.GetElementsByTagName("vOutro_Opc").Item(0).InnerText.Replace(".", ","), "0").Replace(",,", ",")
                                '            objItem.vDescIncond_OpcISS = If(noTetra.GetElementsByTagName("vDescIncond_Opc").Count > 0, noTetra.GetElementsByTagName("vDescIncond_Opc").Item(0).InnerText.Replace(".", ","), "0").Replace(",,", ",")
                                '            objItem.vDescCondISS = If(noTetra.GetElementsByTagName("vDescCond").Count > 0, noTetra.GetElementsByTagName("vDescCond").Item(0).InnerText.Replace(".", ","), "0").Replace(",,", ",")
                                '            objItem.vISSRet_OpcISS = If(noTetra.GetElementsByTagName("vISSRet_Opc").Count > 0, noTetra.GetElementsByTagName("vISSRet_Opc").Item(0).InnerText.Replace(".", ","), "0").Replace(",,", ",")
                                '            objItem.indISS = If(noTetra.GetElementsByTagName("indISS").Count > 0, noTetra.GetElementsByTagName("indISS").Item(0).InnerText.Replace(".", ","), "0").Replace(",,", ",")
                                '            objItem.cServico_OpcISS = If(noTetra.GetElementsByTagName("cServico_Opc").Count > 0, noTetra.GetElementsByTagName("cServico_Opc").Item(0).InnerText, "")
                                '            objItem.cMun_OpcISS = If(noTetra.GetElementsByTagName("cMun_Opc").Count > 0, noTetra.GetElementsByTagName("cMun_Opc").Item(0).InnerText, "")
                                '            objItem.cPais_OpcISS = If(noTetra.GetElementsByTagName("cPais_Opc").Count > 0, noTetra.GetElementsByTagName("cPais_Opc").Item(0).InnerText, "")
                                '            objItem.nProcesso_OpcISS = If(noTetra.GetElementsByTagName("nProcesso_Opc").Count > 0, noTetra.GetElementsByTagName("nProcesso_Opc").Item(0).InnerText, "")
                                '            objItem.indIncentivoISS = If(noTetra.GetElementsByTagName("indIncentivoISS").Count > 0, noTetra.GetElementsByTagName("indIncentivoISS").Item(0).InnerText.Replace(".", ","), "0").Replace(",,", ",")

                                '        End If

                                '        If noTetra.Name = "ICMSUFDest" Then

                                '            objItem.vBCUFDest = If(noTetra.GetElementsByTagName("vBCUFDest").Count > 0, noTetra.GetElementsByTagName("vBCUFDest").Item(0).InnerText.Replace(".", ","), "0").Replace(",,", ",")
                                '            objItem.pFCPUFDest = If(noTetra.GetElementsByTagName("pFCPUFDest").Count > 0, noTetra.GetElementsByTagName("pFCPUFDest").Item(0).InnerText.Replace(".", ","), "0").Replace(",,", ",")
                                '            objItem.pICMSUFDest = If(noTetra.GetElementsByTagName("pICMSUFDest").Count > 0, noTetra.GetElementsByTagName("pICMSUFDest").Item(0).InnerText.Replace(".", ","), "0").Replace(",,", ",")
                                '            objItem.pICMSInterDest = If(noTetra.GetElementsByTagName("pICMSInter").Count > 0, noTetra.GetElementsByTagName("pICMSInter").Item(0).InnerText.Replace(".", ","), "0").Replace(",,", ",")
                                '            objItem.pICMSInterPartDest = If(noTetra.GetElementsByTagName("pICMSInterPart").Count > 0, noTetra.GetElementsByTagName("pICMSInterPart").Item(0).InnerText.Replace(".", ","), "0").Replace(",,", ",")
                                '            objItem.vFCPUFDest = If(noTetra.GetElementsByTagName("vFCPUFDest").Count > 0, noTetra.GetElementsByTagName("vFCPUFDest").Item(0).InnerText.Replace(".", ","), "0").Replace(",,", ",")
                                '            objItem.vICMSUFDest = If(noTetra.GetElementsByTagName("vICMSUFDest").Count > 0, noTetra.GetElementsByTagName("vICMSUFDest").Item(0).InnerText.Replace(".", ","), "0").Replace(",,", ",")
                                '            objItem.vICMSUFRemet = If(noTetra.GetElementsByTagName("vICMSUFRemet").Count > 0, noTetra.GetElementsByTagName("vICMSUFRemet").Item(0).InnerText.Replace(".", ","), "0").Replace(",,", ",")

                                '        End If

                                '        If noTetra.Name = "IPI" Then

                                '            objItem.clEnqIPI = If(noTetra.GetElementsByTagName("clEnq").Count > 0, noTetra.GetElementsByTagName("clEnq").Item(0).InnerText, "")
                                '            objItem.CNPJProdIPI = If(noTetra.GetElementsByTagName("CNPJProd").Count > 0, noTetra.GetElementsByTagName("CNPJProd").Item(0).InnerText, "")
                                '            objItem.cSeloIPI = If(noTetra.GetElementsByTagName("cSelo").Count > 0, noTetra.GetElementsByTagName("cSelo").Item(0).InnerText, "")
                                '            objItem.qSeloIPI = If(noTetra.GetElementsByTagName("qSelo").Count > 0, noTetra.GetElementsByTagName("qSelo").Item(0).InnerText.Replace(".", ","), "0").Replace(",,", ",")
                                '            objItem.cEnqIPI = If(noTetra.GetElementsByTagName("cEnq").Count > 0, noTetra.GetElementsByTagName("cEnq").Item(0).InnerText, "")
                                '            objItem.CSTIPI = If(noTetra.GetElementsByTagName("CST").Count > 0, noTetra.GetElementsByTagName("CST").Item(0).InnerText, "")
                                '            objItem.vBCIPI = If(noTetra.GetElementsByTagName("vBC").Count > 0, noTetra.GetElementsByTagName("vBC").Item(0).InnerText.Replace(".", ","), "0").Replace(",,", ",")
                                '            objItem.pIPI = If(noTetra.GetElementsByTagName("pIPI").Count > 0, noTetra.GetElementsByTagName("pIPI").Item(0).InnerText.Replace(".", ","), "0").Replace(",,", ",")
                                '            objItem.vIPI = If(noTetra.GetElementsByTagName("vIPI").Count > 0, noTetra.GetElementsByTagName("vIPI").Item(0).InnerText.Replace(".", ","), "0").Replace(",,", ",")
                                '            objItem.qUnidIPI = If(noTetra.GetElementsByTagName("qUnid").Count > 0, noTetra.GetElementsByTagName("qUnid").Item(0).InnerText.Replace(".", ","), "0").Replace(",,", ",")
                                '            objItem.vUnidIPI = If(noTetra.GetElementsByTagName("vUnid").Count > 0, noTetra.GetElementsByTagName("vUnid").Item(0).InnerText.Replace(".", ","), "0").Replace(",,", ",")

                                '        End If

                                '    Next

                                'End If

                            Next

                        End If

                    Next

                Next

            End If

        Next

    End Sub

    Private Sub ImportarProdutoUnico(objNotaFiscal As NotaFiscal, xmlDoc As XmlDocument)

        ImportarProdutoXML(objNotaFiscal, xmlDoc, False)

        Dim produto As New NotaFiscalXItem(objNotaFiscal)

        For Each item As [Lib].Negocio.NotaFiscalXItem In objNotaFiscal.Itens

            produto.QuantidadeFiscal += item.QuantidadeFiscal

            If Not item.SubOperacao Is Nothing AndAlso item.SubOperacao.QuantidadeFisico Then
                produto.QuantidadeFisica += item.QuantidadeFiscal
                item.QuantidadeFisica = item.QuantidadeFiscal
            End If

            produto.PesoFiscal += item.PesoFiscal

            produto.PesoBruto += item.PesoBruto
            produto.PesoLiquido += item.PesoLiquido
            produto.ValorLiquido += item.ValorLiquido
            produto.ValorTotal += item.ValorTotal

        Next

        If produto.ValorTotal > 0 And produto.QuantidadeFiscal > 0 Then
            produto.Unitario = produto.ValorTotal / produto.QuantidadeFiscal
        End If

        objNotaFiscal.Itens.Clear()
        objNotaFiscal.Itens.Add(produto)

    End Sub

    Private Sub AgruparNCM(objNotaFiscal As NotaFiscal, xmlDoc As XmlDocument)

        ImportarProdutoXML(objNotaFiscal, xmlDoc, False)

        Dim itensAgrupadoPorNCMProduto = From item In objNotaFiscal.Itens
                                         Group item By item.NCMProdutoXML Into Grupo = Group
                                         Select New With {
                                          .NCMProdutoXML = NCMProdutoXML,
                                           .Itens = Grupo.ToList()
                                      }

        Dim listaDePordutos As New ListNotaFiscalXItem(objNotaFiscal)

        For Each grupo In itensAgrupadoPorNCMProduto

            ' Acessando o valor de NCMProdutoXML
            Dim ncmProduto As String = grupo.NCMProdutoXML

            Dim produto As New NotaFiscalXItem(objNotaFiscal)

            ' Iterando sobre os itens dentro do grupo
            For Each item In grupo.Itens

                ' Aqui você pode acessar as propriedades do item
                produto.QuantidadeFiscal += item.QuantidadeFiscal

                If Not item.SubOperacao Is Nothing AndAlso item.SubOperacao.QuantidadeFisico Then
                    produto.QuantidadeFisica += item.QuantidadeFiscal
                    item.QuantidadeFisica = item.QuantidadeFiscal
                End If

                produto.PesoFiscal += item.PesoFiscal

                produto.PesoBruto += item.PesoBruto
                produto.PesoLiquido += item.PesoLiquido
                produto.ValorLiquido += item.ValorLiquido
                produto.ValorTotal += item.ValorTotal
            Next

            If produto.ValorTotal > 0 And produto.QuantidadeFiscal > 0 Then
                produto.Unitario = produto.ValorTotal / produto.QuantidadeFiscal
            End If

            listaDePordutos.Add(produto)

        Next

        objNotaFiscal.Itens.Clear()
        objNotaFiscal.Itens = listaDePordutos

    End Sub

#End Region


    Protected Sub btnPedidoNFReferencial_Click(sender As Object, e As EventArgs) Handles btnPedidoNFReferencial.Click
        Try
            ucConsultaClientes.Limpar()
            Popup.ConsultaDeClientes(Me.Page, "objClienteNFReferencial" & HID.Value)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub ConsultarPedido()
        HttpContext.Current.Session("ssCampo" & HID.Value) = "Pedidos"
        Dim strEmpresa() As String = ddlEmpresa.SelectedValue.Split("-")
        Dim strCliente() As String = txtCodigoClienteNFReferencial.Value.Split("-")

        Dim parameters As New Dictionary(Of String, Object)
        parameters("empresa") = strEmpresa(0)
        parameters("enderecoEmpresa") = strEmpresa(1)
        parameters("cliente") = IIf(strCliente(0) <> "", strCliente(0), "")
        parameters("enderecoCliente") = IIf(strCliente.Length > 1, strCliente(1), "")
        If Not String.IsNullOrWhiteSpace(txtPedidoNFReferencial.Text) Then
            parameters("pedido") = txtPedidoNFReferencial.Text.Trim()
        End If
        Popup.ConsultaDePedidos(Me.Page, "objPedidoNFReferencial" & HID.Value, "")
        ucConsultaPedidos.BindGridView(parameters)
    End Sub


    Protected Sub txtPedidoNFReferencial_TextChanged(sender As Object, e As EventArgs) Handles txtPedidoNFReferencial.TextChanged
        txtCodigoClienteNFReferencial.Value = String.Empty
    End Sub

    Protected Sub imgDelNFReferencial_Click(sender As Object, e As ImageClickEventArgs)
        Try
            Dim cancRe As ImageButton = CType(sender, ImageButton)
            Dim row As GridViewRow = CType(cancRe.NamingContainer, GridViewRow)

            Dim lblNota As Label = row.FindControl("lblNota")

            SessaoRecuperaNotaFiscal()

            If objNotaFiscal.IUD = "I" Then
                objNotaFiscal.NotasReferenciais.RemoveAt(row.RowIndex)
            Else
                objNotaFiscal.NotasReferenciais.Where(Function(s) s.IUD <> "D" And s.Nota_Id = lblNota.Text).Single.IUD = "D"

                For i As Integer = 0 To objNotaFiscal.NotasReferenciais.Count - 1
                    If String.IsNullOrWhiteSpace(objNotaFiscal.NotasReferenciais(i).IUD) OrElse objNotaFiscal.NotasReferenciais(i).IUD <> "D" Then
                        objNotaFiscal.NotasReferenciais(i).IUD = "X"
                    End If
                Next
            End If
            If objNotaFiscal.NotasReferenciais IsNot Nothing Then
                If objNotaFiscal.NotasReferenciais.Where(Function(s) s.IUD <> "D").Count = 0 Then
                    txtPedidoNFReferencial.ReadOnly = False
                    txtPedidoNFReferencial.Text = String.Empty
                    btnPedidoNFReferencial.Enabled = True
                End If
            End If

            SessaoSalvaNotaFiscal()

            grdNotasReferenciais.DataSource = objNotaFiscal.NotasReferenciais.Where(Function(s) s.IUD <> "D")
            grdNotasReferenciais.DataBind()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub grdProdutos_RowDataBound(sender As Object, e As GridViewRowEventArgs) Handles grdProdutos.RowDataBound
        If e.Row.RowType = DataControlRowType.DataRow Then
            Dim CentroDecusto = CType(e.Row.FindControl("lblCentroDeCusto"), Label)

            If Not String.IsNullOrWhiteSpace(CentroDecusto.Text) AndAlso CentroDecusto.Text > 0 Then
                CentroDecusto.ToolTip = New CentroDeCusto(CentroDecusto.Text).Descricao
            End If

            If objNotaFiscal.Itens(e.Row.RowIndex).ProdutoXML.Length > 0 Then
                Dim prdXML = CType(e.Row.FindControl("imgProdutoDeTerceiro"), ImageButton)
                prdXML.Visible = True
                prdXML.ToolTip = objNotaFiscal.Itens(e.Row.RowIndex).ProdutoXMLDeTerceiro
            End If

        End If
    End Sub

    Protected Sub btnRepresentante_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            ucConsultaClientes.Limpar()
            Dim txtNome As TextBox = CType(ucConsultaClientes.FindControlRecursive("txtNome"), TextBox)
            Popup.ConsultaDeClientes(Me, "objRepresentante" & HID.Value, txtNome.ClientID)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

End Class
