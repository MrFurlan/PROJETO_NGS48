Imports System.Data
Imports System.Collections.Generic
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis


Public Class ucNFObsProduto
    Inherits BaseUserControl

    'Carrega a Nota Fiscal armazenada na Sessão'
    Private Property objNotaFiscal As [Lib].Negocio.NotaFiscal
        Get
            Return CType(Session("objNotaFiscal" & HID.Value), [Lib].Negocio.NotaFiscal)
        End Get
        Set(value As [Lib].Negocio.NotaFiscal)
            Session("objNotaFiscal" & HID.Value) = CType(value, [Lib].Negocio.NotaFiscal)
        End Set
    End Property
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not IsPostBack Then
        End If
    End Sub
    Public Overrides Sub SetarHID(ByVal guid As String)
        HID.Value = guid.ToString
    End Sub

    Public Overrides Sub Limpar()
        txtObservacoesDoProduto.Text = String.Empty

        If pnlPecas.Visible Then txtPecas.Text = String.Empty

        If objNotaFiscal.Itens(Session("objIndiceProduto" & HID.Value)).Produto.ControlarNumeroDoLote AndAlso objNotaFiscal.Itens(Session("objIndiceProduto" & HID.Value)).SubOperacao.ControlarNumeroDoLote Then
            divInfLote.Visible = True
            LimparInfLote()
        Else
            divInfLote.Visible = False
        End If

    End Sub

    Public Overrides Sub Carregar(obj As IBaseEntity)
        If Session("objLoteFornecedor" & HID.Value) IsNot Nothing Then

            For Each drItemLote As DataRow In CType(Session("objLoteFornecedor" & HID.Value), DataTable).Rows
                If drItemLote("Produto") = objNotaFiscal.Itens(Session("objIndiceProduto" & HID.Value)).CodigoProduto Then

                    Dim Sql As String = "SELECT nXt.Nota_Id, nXt.Fabricado, nXt.Validade" & vbCrLf &
                                          "FROM NotaFiscalXLote nXt " & vbCrLf &
                                          "	inner join NotasFiscais n" & vbCrLf &
                                          "			ON n.Empresa_Id       = nXt.Empresa_Id" & vbCrLf &
                                          "			and n.EndEmpresa_iD   = nXt.EndEmpresa_Id" & vbCrLf &
                                          "			and n.Cliente_Id      = nXt.Cliente_Id" & vbCrLf &
                                          "			and n.EndCliente_Id   = nXt.EndCliente_Id" & vbCrLf &
                                          "			and n.EntradaSaida_Id = nXt.EntradaSaida_Id" & vbCrLf &
                                          "			and n.Serie_Id        = nXt.Serie_Id" & vbCrLf &
                                          "			and n.Nota_Id         = nXt.Nota_Id" & vbCrLf &
                                          "Where nXt.Empresa_Id      = '" & objNotaFiscal.CodigoEmpresa & "'" & vbCrLf &
                                          "  and nXt.EndEmpresa_Id   = " & objNotaFiscal.EnderecoEmpresa & vbCrLf &
                                          "  and nXt.EntradaSaida_Id = 'E'" & vbCrLf &
                                          "  and nXt.Produto_Id      = '" & objNotaFiscal.Itens(Session("objIndiceProduto" & HID.Value)).CodigoProduto & "'" & vbCrLf &
                                          "  and nXt.Lote_Id         = '" & drItemLote("Lote") & "'" & vbCrLf &
                                          "  and n.Situacao          = 1" & vbCrLf &
                                          "UNION ALL" & vbCrLf &
                                          "SELECT 1 as Nota_Id, op.Movimento as Fabricado, op.Validade" & vbCrLf &
                                          "FROM OrdemDeProducao op " & vbCrLf &
                                          "INNER JOIN OrdemDeProducaoXProduto OPXP " & vbCrLf &
                                          "   ON op.Empresa_Id		        = OPXP.Empresa_Id " & vbCrLf &
                                          "   AND op.EndEmpresa_Id	        = OPXP.EndEmpresa_Id " & vbCrLf &
                                          "   AND op.Ordem_Id			    = OPXP.Ordem_Id " & vbCrLf &
                                          "Where op.Empresa_Id              = '" & objNotaFiscal.CodigoEmpresa & "'" & vbCrLf &
                                          "  and op.EndEmpresa_Id           = " & objNotaFiscal.EnderecoEmpresa & vbCrLf &
                                          "  and OPXP.Produto_Id            = '" & objNotaFiscal.Itens(Session("objIndiceProduto" & HID.Value)).CodigoProduto & "'" & vbCrLf &
                                          "  and OPXP.Lote                  = '" & drItemLote("Lote") & "'" & vbCrLf &
                                          "  and op.Situacao                = 1 "

                    Dim ds As DataSet = Banco.ConsultaDataSet(Sql, "NotaFiscalXLote")

                    If ds IsNot Nothing AndAlso ds.Tables(0) IsNot Nothing AndAlso ds.Tables(0).Rows.Count > 0 Then
                        For Each row As DataRow In ds.Tables(0).Rows
                            drItemLote("Fabricado") = CDate(row("Fabricado")).ToString("yyyy-MM-dd")
                            Exit For
                        Next
                    End If

                    If drItemLote("Consumo") > 0 Then
                        Dim lote As New NotaFiscalXLote(objNotaFiscal.Itens(Session("objIndiceProduto" & HID.Value)))
                        lote.IUD = "I"
                        lote.Lote = drItemLote("Lote")
                        lote.Fabricado = drItemLote("Fabricado")
                        lote.Validade = drItemLote("Validade")
                        lote.Quantidade = drItemLote("Consumo")
                        objNotaFiscal.Itens(Session("objIndiceProduto" & HID.Value)).Lotes.Add(lote)
                    End If
                End If
            Next


            If objNotaFiscal.IUD = "I" Then

                If Left(objNotaFiscal.Empresa.Codigo, 8) = "40938762" OrElse Left(objNotaFiscal.Empresa.Codigo, 8) = "49673784" Then
                    'CASO VOLTE A MOSTRAR MEXER NO DOCUMENTO ELETRÔNICO EM RASTRO
                    'SE FOR BAXI(FOODS E DISTRIBUIDORA) NÃO INFORMAR LOTE ATÉ SEGUNDA ORDEM, SOLICITAÇÃO DOUGLAS - FURLAN - 06/09/2024
                Else
                    Dim sSeparador As String = String.Empty
                    objNotaFiscal.Itens(Session("objIndiceProduto" & HID.Value)).ObservacoesDoProduto = Funcoes.EliminarCaracteresEspeciaisNF(txtObservacoesDoProduto.Text)

                    For Each infLote In objNotaFiscal.Itens(Session("objIndiceProduto" & HID.Value)).Lotes
                        objNotaFiscal.Itens(Session("objIndiceProduto" & HID.Value)).ObservacoesDoProduto += sSeparador & "Lote: " & infLote.Lote & " Fabricado: " & infLote.Fabricado.ToString("dd/MM/yyyy") & " Validade: " & infLote.Validade.ToString("dd/MM/yyyy")
                        sSeparador = " - "
                    Next

                    txtObservacoesDoProduto.Text = objNotaFiscal.Itens(Session("objIndiceProduto" & HID.Value)).ObservacoesDoProduto
                End If
            End If

            LimparInfLote()

            atualizarGridLote()

        End If
    End Sub

    Public Sub CarregarObs(ByRef pIndiceDoProduto As Integer)
        If pIndiceDoProduto >= 0 Then
            Session("objIndiceProduto" & HID.Value) = pIndiceDoProduto
            If (objNotaFiscal.IUD = "U" AndAlso objNotaFiscal.NossaEmissao AndAlso objNotaFiscal.Eletronica) Then
                liNovo.Style.Add("display", "none")
                txtObservacoesDoProduto.ReadOnly = True
            Else
                liNovo.Style.Clear()
                txtObservacoesDoProduto.ReadOnly = False
            End If

            'Observação do Produto'
            txtObservacoesDoProduto.Text = objNotaFiscal.Itens(Session("objIndiceProduto" & HID.Value)).ObservacoesDoProduto

            If Not objNotaFiscal.NFG Then
                'Mensagem em Notas de Devolução
                If objNotaFiscal.SubOperacao.Devolucao And objNotaFiscal.IUD = "I" Then
                    txtMsgDevolucaoProduto.Text = objNotaFiscal.Itens(Session("objIndiceProduto" & HID.Value)).NotasDevolucao.MsgDevolucao
                ElseIf objNotaFiscal.SubOperacao.Devolucao And objNotaFiscal.IUD = "U" Then
                    txtMsgDevolucaoProduto.Text = "Após a gravação a mensagem de devolução é armazenada nos dados adicionais da nota."
                End If

                'Impostos
                lblImpostos.Text = objNotaFiscal.Itens(Session("objIndiceProduto" & HID.Value)).Encargos.MensagemImpostos

                pnlDevolucao.Visible = True

                'Controlar Peças'
                If objNotaFiscal.Itens(Session("objIndiceProduto" & HID.Value)).Produto.ControlarPecas AndAlso objNotaFiscal.Itens(Session("objIndiceProduto" & HID.Value)).SubOperacao.ControlarPecas Then
                    If objNotaFiscal.Itens(Session("objIndiceProduto" & HID.Value)).NumeroPecas > 0 Then
                        txtPecas.Text = objNotaFiscal.Itens(Session("objIndiceProduto" & HID.Value)).NumeroPecas
                    End If
                    pnlPecas.Visible = True
                Else
                    pnlPecas.Visible = False
                End If

                'Informações do Lote
                If divInfLote.Visible Then
                    If objNotaFiscal.Itens(Session("objIndiceProduto" & HID.Value)).Lotes.Count > 0 Then
                        gridInfLote.DataSource = objNotaFiscal.Itens(Session("objIndiceProduto" & HID.Value)).Lotes.Where(Function(s) s.IUD <> "D")
                    Else
                        gridInfLote.DataSource = Nothing
                    End If
                    gridInfLote.DataBind()
                End If
            Else
                pnlDevolucao.Visible = False
                pnlPecas.Visible = False
            End If
        End If
    End Sub

    Protected Sub lnkNovo_Click(sender As Object, e As EventArgs) Handles lnkNovo.Click
        Try
            If objNotaFiscal.NFG Then
                If String.IsNullOrWhiteSpace(txtObservacoesDoProduto.Text) Then
                    MsgBox(Me.Page, "Observação não foi informada.")
                Else
                    objNotaFiscal.Itens(Session("objIndiceProduto" & HID.Value)).ObservacoesDoProduto = Funcoes.EliminarCaracteresEspeciaisNF(txtObservacoesDoProduto.Text)

                    Popup.CloseDialog(Me.Page, "divNFObsProduto")
                End If
            Else
                objNotaFiscal.Itens(Session("objIndiceProduto" & HID.Value)).ObservacoesDoProduto = Funcoes.EliminarCaracteresEspeciaisNF(txtObservacoesDoProduto.Text)

                If pnlPecas.Visible Then
                    objNotaFiscal.Itens(Session("objIndiceProduto" & HID.Value)).NumeroPecas = txtPecas.Text
                End If

                If objNotaFiscal.Itens(Session("objIndiceProduto" & HID.Value)).Produto.ControlarNumeroDoLote AndAlso objNotaFiscal.Itens(Session("objIndiceProduto" & HID.Value)).SubOperacao.ControlarNumeroDoLote Then
                    If TypeOf Me.Page Is NotaFiscalXItens Then
                        CType(Me.Page, NotaFiscalXItens).DesabilitaLinhaProduto(Session("objIndiceProduto" & HID.Value))
                    End If
                End If

                Popup.CloseDialog(Me.Page, "divNFObsProduto")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnAdicionarInfLote_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)
        If String.IsNullOrEmpty(txtNumeroDoLote.Text) Then
            MsgBox(Me.Page, "Informe o número do Lote.", eTitulo.Info)
        ElseIf Not IsDate(txtDataFabricadoLote.Text) Then
            MsgBox(Me.Page, "Informe uma data de Fabricação válida.", eTitulo.Info)
        ElseIf Not IsDate(txtDataValidadeLote.Text) Then
            MsgBox(Me.Page, "Informe uma data de Validade válida.", eTitulo.Info)
        ElseIf CDate(txtDataValidadeLote.Text) < Now.Date Then
            MsgBox(Me.Page, "Data de Validade não pode ser menor que a data atual.", eTitulo.Info)
        ElseIf String.IsNullOrEmpty(txtQuantidadeLote.Text) OrElse CDec(txtQuantidadeLote.Text) = 0 Then
            MsgBox(Me.Page, "Informe a quantidade do Lote.", eTitulo.Info)
        Else

            Dim Sql As String = "SELECT nXt.Nota_Id, nXt.Fabricado, nXt.Validade" & vbCrLf &
                                "FROM NotaFiscalXLote nXt " & vbCrLf &
                                          "	inner join NotasFiscais n" & vbCrLf &
                                          "			ON n.Empresa_Id       = nXt.Empresa_Id" & vbCrLf &
                                          "			and n.EndEmpresa_iD   = nXt.EndEmpresa_Id" & vbCrLf &
                                          "			and n.Cliente_Id      = nXt.Cliente_Id" & vbCrLf &
                                          "			and n.EndCliente_Id   = nXt.EndCliente_Id" & vbCrLf &
                                          "			and n.EntradaSaida_Id = nXt.EntradaSaida_Id" & vbCrLf &
                                          "			and n.Serie_Id        = nXt.Serie_Id" & vbCrLf &
                                          "			and n.Nota_Id         = nXt.Nota_Id" & vbCrLf &
                                "Where nXt.Empresa_Id      = '" & objNotaFiscal.CodigoEmpresa & "'" & vbCrLf &
                                "  and nXt.EndEmpresa_Id   = " & objNotaFiscal.EnderecoEmpresa & vbCrLf &
                                "  and nXt.EntradaSaida_Id = '" & objNotaFiscal.EntradaSaida.ToString.Substring(0, 1) & "'" & vbCrLf &
                                "  and nXt.Produto_Id      = '" & objNotaFiscal.Itens(Session("objIndiceProduto" & HID.Value)).CodigoProduto & "'" & vbCrLf &
                                "  and nXt.Lote_Id         = '" & Trim(txtNumeroDoLote.Text.ToUpper()) & "'" & vbCrLf &
                                "  and n.Situacao = 1" & vbCrLf &
                                "UNION ALL" & vbCrLf &
                                "SELECT 1 as Nota_Id, op.Movimento as Fabricado, op.Validade" & vbCrLf &
                                "FROM OrdemDeProducao op" & vbCrLf &
                                "INNER JOIN OrdemDeProducaoXProduto OPXP " & vbCrLf &
                                "   ON op.Empresa_Id		        = OPXP.Empresa_Id " & vbCrLf &
                                "   AND op.EndEmpresa_Id	        = OPXP.EndEmpresa_Id " & vbCrLf &
                                "   AND op.Ordem_Id			        = OPXP.Ordem_Id " & vbCrLf &
                                "Where op.Empresa_Id                = '" & objNotaFiscal.CodigoEmpresa & "'" & vbCrLf &
                                "  and op.EndEmpresa_Id             = " & objNotaFiscal.EnderecoEmpresa & vbCrLf &
                                "  and OPXP.Produto_Id              = '" & objNotaFiscal.Itens(Session("objIndiceProduto" & HID.Value)).CodigoProduto & "'" & vbCrLf &
                                "  and OPXP.Lote                    = '" & Trim(txtNumeroDoLote.Text.ToUpper()) & "'" & vbCrLf &
                                "  and op.Situacao                  = 1"

            Dim ds As DataSet = Banco.ConsultaDataSet(Sql, "NotaFiscalXLote")

            If ds IsNot Nothing AndAlso ds.Tables(0) IsNot Nothing AndAlso ds.Tables(0).Rows.Count > 0 Then
                For Each row As DataRow In ds.Tables(0).Rows
                    If Not CDate(row("Validade")).ToString("yyyy-MM-dd") = CDate(txtDataValidadeLote.Text).ToString("yyyy-MM-dd") Then
                        MsgBox(Me.Page, "Número de Lote já existe com outra data na Nota Fiscal " & row("Nota_Id") & ". O mesmo lote não pode ter vencimento(s) diferente(s).", eTitulo.Info)
                        Exit Sub
                    End If
                Next
            End If

            If gridInfLote.Rows.Count > 0 Then
                Dim quantidade As Decimal = CDec(txtQuantidadeLote.Text)
                For Each pRow As GridViewRow In gridInfLote.Rows
                    If txtNumeroDoLote.Text.ToUpper() = pRow.Cells(1).Text Then
                        MsgBox(Me.Page, "Número de Lote já existe na lista.", eTitulo.Info)
                        Exit Sub
                    Else
                        quantidade += CDec(pRow.Cells(4).Text)
                    End If
                Next

                If objNotaFiscal.SubOperacao.EstoqueFisico Then
                    If CDec(txtQuantidadeLote.Text) > objNotaFiscal.Itens(Session("objIndiceProduto" & HID.Value)).QuantidadeFisica Then
                        MsgBox(Me.Page, "Quantidade(s) de Lote(s) informada não pode ser maior que a Quantidade do produto.", eTitulo.Info)
                    Else
                        Dim lote As New NotaFiscalXLote(objNotaFiscal.Itens(Session("objIndiceProduto" & HID.Value)))
                        lote.IUD = "I"
                        lote.Lote = Trim(txtNumeroDoLote.Text.ToUpper())
                        lote.Fabricado = CDate(txtDataFabricadoLote.Text)
                        lote.Validade = CDate(txtDataValidadeLote.Text)
                        lote.Quantidade = CDec(txtQuantidadeLote.Text)
                        objNotaFiscal.Itens(Session("objIndiceProduto" & HID.Value)).Lotes.Add(lote)
                    End If
                Else
                    If quantidade > objNotaFiscal.Itens(Session("objIndiceProduto" & HID.Value)).QuantidadeFiscal Then
                        MsgBox(Me.Page, "Quantidade(s) de Lote(s) informada não pode ser maior que a Quantidade do produto.", eTitulo.Info)
                    Else
                        Dim lote As New NotaFiscalXLote(objNotaFiscal.Itens(Session("objIndiceProduto" & HID.Value)))
                        lote.IUD = "I"
                        lote.Lote = Trim(txtNumeroDoLote.Text.ToUpper())
                        lote.Fabricado = CDate(txtDataFabricadoLote.Text)
                        lote.Validade = CDate(txtDataValidadeLote.Text)
                        lote.Quantidade = CDec(txtQuantidadeLote.Text)
                        objNotaFiscal.Itens(Session("objIndiceProduto" & HID.Value)).Lotes.Add(lote)
                    End If
                End If
            Else
                If objNotaFiscal.SubOperacao.EstoqueFisico Then
                    If CDec(txtQuantidadeLote.Text) > objNotaFiscal.Itens(Session("objIndiceProduto" & HID.Value)).QuantidadeFisica Then
                        MsgBox(Me.Page, "Quantidade(s) de Lote(s) informada não pode ser maior que a Quantidade do produto.", eTitulo.Info)
                    Else
                        Dim lote As New NotaFiscalXLote(objNotaFiscal.Itens(Session("objIndiceProduto" & HID.Value)))
                        lote.IUD = "I"
                        lote.Lote = Trim(txtNumeroDoLote.Text.ToUpper())
                        lote.Fabricado = CDate(txtDataFabricadoLote.Text)
                        lote.Validade = CDate(txtDataValidadeLote.Text)
                        lote.Quantidade = CDec(txtQuantidadeLote.Text)
                        objNotaFiscal.Itens(Session("objIndiceProduto" & HID.Value)).Lotes.Add(lote)
                    End If
                Else
                    If CDec(txtQuantidadeLote.Text) > objNotaFiscal.Itens(Session("objIndiceProduto" & HID.Value)).QuantidadeFiscal Then
                        MsgBox(Me.Page, "Quantidade(s) de Lote(s) informada não pode ser maior que a Quantidade do produto.", eTitulo.Info)
                    Else
                        Dim lote As New NotaFiscalXLote(objNotaFiscal.Itens(Session("objIndiceProduto" & HID.Value)))
                        lote.IUD = "I"
                        lote.Lote = Trim(txtNumeroDoLote.Text.ToUpper())
                        lote.Fabricado = CDate(txtDataFabricadoLote.Text)
                        lote.Validade = CDate(txtDataValidadeLote.Text)
                        lote.Quantidade = CDec(txtQuantidadeLote.Text)
                        objNotaFiscal.Itens(Session("objIndiceProduto" & HID.Value)).Lotes.Add(lote)
                    End If
                End If
            End If

            LimparInfLote()

            atualizarGridLote()
        End If
    End Sub

    Protected Sub imgExcluirLote_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)
        Dim img As ImageButton = CType(sender, ImageButton)
        Dim row As GridViewRow = CType(img.NamingContainer, GridViewRow)

        For Each Lote In objNotaFiscal.Itens(Session("objIndiceProduto" & HID.Value)).Lotes.Where(Function(s) s.IUD <> "D")
            If Lote.Lote = row.Cells(1).Text.ToUpper() Then
                objNotaFiscal.Itens(Session("objIndiceProduto" & HID.Value)).Lotes.RemoveAt(row.RowIndex)
                Exit For
            End If
        Next

        atualizarGridLote()
    End Sub

    Private Sub LimparInfLote()
        txtNumeroDoLote.Enabled = True
        txtDataFabricadoLote.Enabled = True
        txtDataValidadeLote.Enabled = True
        txtQuantidadeLote.Enabled = True

        btnAdicionarInfLote.Visible = True
        imgSelecionaLote.Visible = False

        txtNumeroDoLote.Text = String.Empty
        txtDataFabricadoLote.Text = String.Empty
        txtDataValidadeLote.Text = String.Empty
        txtQuantidadeLote.Text = "0,0000"

        gridInfLote.DataSource = Nothing
        gridInfLote.DataBind()

        If objNotaFiscal.CodigoEmpresa = "05272759000147" AndAlso objNotaFiscal.SubOperacao.EntradaSaida = eEntradaSaida.Saida AndAlso _
            (objNotaFiscal.Itens(Session("objIndiceProduto" & HID.Value)).Produto.CodigoGrupo = "10101" Or objNotaFiscal.Itens(Session("objIndiceProduto" & HID.Value)).Produto.CodigoGrupo = "10102" Or objNotaFiscal.Itens(Session("objIndiceProduto" & HID.Value)).Produto.CodigoGrupo = "10103" Or objNotaFiscal.Itens(Session("objIndiceProduto" & HID.Value)).Produto.CodigoGrupo = "30101" Or objNotaFiscal.Itens(Session("objIndiceProduto" & HID.Value)).Produto.CodigoGrupo = "30102" Or objNotaFiscal.Itens(Session("objIndiceProduto" & HID.Value)).Produto.CodigoGrupo = "10110" Or objNotaFiscal.Itens(Session("objIndiceProduto" & HID.Value)).Produto.CodigoGrupo = "10111" Or objNotaFiscal.Itens(Session("objIndiceProduto" & HID.Value)).Produto.CodigoGrupo = "10112" Or objNotaFiscal.Itens(Session("objIndiceProduto" & HID.Value)).Produto.CodigoGrupo = "10105" Or objNotaFiscal.Itens(Session("objIndiceProduto" & HID.Value)).Produto.CodigoGrupo = "10106") Then
            txtNumeroDoLote.Enabled = False
            txtDataFabricadoLote.Enabled = False
            txtDataValidadeLote.Enabled = False
            txtQuantidadeLote.Enabled = False
            btnAdicionarInfLote.Visible = False

            If objNotaFiscal.IUD = "I" Then
                imgSelecionaLote.Visible = True
            Else
                imgSelecionaLote.Visible = False
            End If

        ElseIf objNotaFiscal.SubOperacao.EntradaSaida = eEntradaSaida.Saida AndAlso (Left(objNotaFiscal.CodigoEmpresa, 8) = "40938762" Or Left(objNotaFiscal.CodigoEmpresa, 8) = "49673784") Then
            txtNumeroDoLote.Enabled = False
            txtDataFabricadoLote.Enabled = False
            txtDataValidadeLote.Enabled = False
            txtQuantidadeLote.Enabled = False
            btnAdicionarInfLote.Visible = False

            If objNotaFiscal.IUD = "I" Then
                imgSelecionaLote.Visible = True
            Else
                imgSelecionaLote.Visible = False
            End If
        End If
    End Sub
    Private Sub atualizarGridLote()
        gridInfLote.DataSource = objNotaFiscal.Itens(Session("objIndiceProduto" & HID.Value)).Lotes.Where(Function(s) s.IUD <> "D")
        gridInfLote.DataBind()

        For Each row As GridViewRow In gridInfLote.Rows
            If objNotaFiscal.SubOperacao.EntradaSaida = eEntradaSaida.Entrada Then
                CType(gridInfLote.Rows(row.RowIndex).FindControl("imgExcluirLote"), ImageButton).Enabled = True
            Else
                CType(gridInfLote.Rows(row.RowIndex).FindControl("imgExcluirLote"), ImageButton).Enabled = False
            End If
        Next
    End Sub

    Protected Sub imgSelecionaLote_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)

        Dim ucConsultaLote = CType(Me.Page.FindControlRecursive("ucConsultaLote"), ucConsultaLote)
        If ucConsultaLote IsNot Nothing Then
            ucConsultaLote.Limpar()
            ucConsultaLote.SetarHID(HID.Value)
            ucConsultaLote.MainUserControl = Me
            ucConsultaLote.CarregarLote(objNotaFiscal.CodigoEmpresa, objNotaFiscal.EnderecoEmpresa, objNotaFiscal.Itens(Session("objIndiceProduto" & HID.Value)))
        End If

        Popup.ConsultaDeLote(Me.Page, "objConsultaDeLote" & HID.Value)
    End Sub

    Protected Sub lnkLimpar_Click(sender As Object, e As EventArgs) Handles lnkLimpar.Click
        Limpar()
    End Sub
    Protected Sub lnkFechar_Click(sender As Object, e As EventArgs) Handles lnkFechar.Click
        Popup.CloseDialog(Me.Page, "divNFObsProduto")
    End Sub
End Class